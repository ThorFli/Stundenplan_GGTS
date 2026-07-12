# CP-SAT Stundenplan-Solver

Automatische Stundenplangenerierung für eine Grundschule auf Basis von Google OR-Tools CP-SAT und Excel-Eingabedaten.

---

## Inhaltsverzeichnis

1. [Voraussetzungen](#voraussetzungen)
2. [Installation](#installation)
3. [Ausführung](#ausführung)
4. [Eingabe-Arbeitsmappe](#eingabe-arbeitsmappe)
5. [Ausgabe-Arbeitsmappe](#ausgabe-arbeitsmappe)
6. [Regelwerk (13_Regeln)](#regelwerk-13_regeln)
7. [Lehrerwünsche (10_Wuensche)](#lehrerwünsche-10_wuensche)
8. [Modellarchitektur](#modellarchitektur)
9. [Optimierungsstrategie (2-Phasen)](#optimierungsstrategie-2-phasen)
10. [Besondere Fächer](#besondere-fächer)
11. [Bekannte Einschränkungen](#bekannte-einschränkungen)
12. [Änderungshistorie](#änderungshistorie)

---

## Voraussetzungen

- Python 3.13 oder neuer
- Lokale Installation von OR-Tools (wird **nicht** in Cloud-Umgebungen mitgeliefert)
- openpyxl

---

## Installation

```bash
# Virtuelle Umgebung anlegen (empfohlen)
python3 -m venv .venv
source .venv/bin/activate

# Abhängigkeiten installieren
pip install -r requirements_cp_sat.txt
# oder direkt:
pip install ortools openpyxl
```

---

## Ausführung

```bash
python3 stundenplan_cpsat_solver.py \
    --input  Stundenplan_Vorlage_Wuensche_Lehrer.xlsx \
    --output stundenplan_cpsat_ergebnis.xlsx \
    --time-limit 300 \
    --workers 8
```

| Parameter | Standard | Beschreibung |
|---|---|---|
| `--input` | `Stundenplan_Vorlage_Wuensche_Lehrer.xlsx` | Pfad zur Eingabe-Arbeitsmappe |
| `--output` | `stundenplan_cpsat_ergebnis.xlsx` | Pfad zur Ausgabe-Arbeitsmappe |
| `--time-limit` | `300` | Gesamte Rechenzeit in Sekunden (Phase 1 + Phase 2) |
| `--workers` | `8` | Anzahl paralleler Solver-Threads |

Die Rechenzeit teilt sich automatisch auf: 40 % für Phase 1 (Primäroptimierung), 60 % für Phase 2 (Sekundäroptimierung).

---

## Eingabe-Arbeitsmappe

Die Datei `Stundenplan_Vorlage_Wuensche_Lehrer.xlsx` enthält folgende relevante Blätter:

### 02_Zeiten
Definiert alle planbaren Zeitslots der Woche.

| Spalte | Bedeutung |
|---|---|
| `Tag` | Wochentag (Montag–Freitag) |
| `Block` | Stundennummer (1–5, Leseband, UF1, UF2, GF1, GF2) |
| `Beginn` / `Ende` | Uhrzeit |
| `Klassenstufen_Gruppe` | `1+4` oder `2+3` (welche Klassen diesen Slot nutzen) |
| `Typ` | z. B. `Kernzeit`, `GF`, `UF` |
| `Planbar` | `Ja` = wird vom Solver berücksichtigt |

### 03_Klassen
Alle Klassen der Schule.

| Spalte | Bedeutung |
|---|---|
| `Klasse` | z. B. `1.1`, `2.3` |
| `Klassenstufe` | Zahl 1–4 |
| `Klassenlehrer` | Kürzel der Klassenlehrkraft (wird für Mo-1/Fr-5-Regel benötigt) |
| `Max_Std_Tag` | Maximale Unterrichtsstunden pro Tag |

> **Wichtig:** Die Spalte `Klassenlehrer` muss befüllt sein, damit die Regel `Klassenlehrer_Wochen_Start_Ende` greifen kann.

### 04_Faecher
Liste aller Fächer (informativ, wird nicht direkt eingelesen).

### 05_Unterrichtsbedarf
Wochenstundenbedarf je Klasse und Fach.

| Spalte | Bedeutung |
|---|---|
| `Klasse` | z. B. `1.1` |
| `Fach` | z. B. `Deutsch`, `Mathematik` |
| `Wochenstunden` | Anzahl benötigter Stunden pro Woche |
| `Lehrer` | Optional: feste Lehrkraftzuordnung (Kürzel) |
| `Ressource` | z. B. `Sporthalle` |

### 06_Lehrer
Stammdaten aller Lehrkräfte.

| Spalte | Bedeutung |
|---|---|
| `Kuerzel` | Eindeutiges Kürzel |
| `Name` | Vollständiger Name |
| `Rolle` | `Klassenlehrer` oder `Fachlehrer` |
| `Max_Std_Woche` | Wochenstunden-Soll; fehlt oder `0` → Lehrkraft wird nicht eingeplant |
| `Deputatsstunden_Blocke` | Feste Deputat-Blöcke, die Unterrichtszeit blockieren |
| `Kommentar` | Enthält ggf. Marker wie `kein Einsatz im Unterricht` |

Marker, die eine Lehrkraft komplett vom Unterricht ausschließen:
- `kein Einsatz im Unterricht`
- `keine Unterrichtszuteilung`
- `langzeiterkrankt`
- `EZ bis`

### 07_Lehrer_Verfuegbarkeit
Verfügbarkeit je Lehrkraft und Wochentag.

| Spalte | Bedeutung |
|---|---|
| `Kuerzel` | Lehrkraft-Kürzel |
| `Montag` … `Freitag` | Kommagetrennte Liste verfügbarer Blöcke; `frei` = ganztägig nicht verfügbar |

### 09_Feste_Termine
Noch nicht vollständig implementiert; DAZ-Fixslots und Dienstag-Religion/Sport werden aktuell direkt im Code erzeugt (→ `create_tasks`).

### 10_Wuensche
Wünsche der Lehrkräfte (Details → [Abschnitt Lehrerwünsche](#lehrerwünsche-10_wuensche)).

### 11_Freizeit
Informativ; Freizeitblöcke werden aus `02_Zeiten` abgeleitet.

### 13_Regeln
Aktivierung/Deaktivierung einzelner Solver-Regeln (Details → [Abschnitt Regelwerk](#regelwerk-13_regeln)).

---

## Ausgabe-Arbeitsmappe

Die erzeugte Datei enthält vier Blätter:

### Klassenplaene
Zeilenweise Auflistung aller Unterrichtszuweisungen.

| Spalte | Inhalt |
|---|---|
| `Klasse` | z. B. `1.1` |
| `Tag` | Wochentag |
| `Block` | Stundenblock |
| `Zeit` | Uhrzeit `HH:MM-HH:MM` |
| `Fach` | Fachbezeichnung |
| `Lehrer` | Kürzel; bei Förderunterricht mit Co-Lehrer: `Lehrer1 + Lehrer2` |
| `Status` | `OK` oder `UNASSIGNED` |
| `Hinweis` | Zusatzinformation (z. B. Überhangstunde, Co-Lehrer-Vermerk) |

Rot hinterlegte Zeilen = nicht zugewiesene Stunden.

### Klassenmatrix
Tabellarische Wochenübersicht je Klasse: Wochentage als Spalten, Stundenblöcke als Zeilen. Fach und Lehrerkürzel stehen in jeder Zelle.

### Lehrerplaene
Chronologisch sortierte Auflistung aller Stunden je Lehrkraft.

### Validierung
- Lösungsstatistiken (Status, Objective, Rechenzeit, Regelübersicht)
- Alle Warnungen (Unterdeckungen, MUSS-Verstöße, Fallback-Hinweise)

---

## Regelwerk (13_Regeln)

Jede Regel wird mit `Aktiv = Ja` aktiviert. Folgende Regeln sind implementiert:

| Regelname | Art | Beschreibung |
|---|---|---|
| `Aufbau_Klassenmatrix` | Hart | Klassenstufen-Gruppen (`1+4`, `2+3`) teilen dieselben Zeitslots |
| `DAZ_taeglich_1_4` | Hart | Klasse 1.4 (DAZ) belegt Mo–Fr die Blöcke 1–4 mit fixen Slots |
| `Deputatsstunden_blockieren_Unterricht` | Hart | Deputatstunden in `06_Lehrer` sperren die entsprechenden Slots für Unterricht |
| `Deutsch_Mathe` | Hart | Deutsch und Mathematik dürfen nicht in Block 5 stattfinden |
| `Fach_Lehrer` | Hart | Pro Klasse und Fach wird genau eine Lehrkraft für die ganze Woche zugewiesen (Ausnahmen: Förderunterricht, DAZ, Gebundene/Ungebundene Freizeit) |
| `Faecher_verteilt` | Weich | Gleiches Fach soll möglichst nicht mehrfach am selben Tag stattfinden |
| `Gemeinsame_UF1_UF2_pro_Klassenstufe` | Hart | UF-Slots gelten klassenstufen-übergreifend |
| `Hauptfaecher_vormittags` | Weich | Deutsch, Mathematik u. a. Hauptfächer bevorzugt in Blöcken 1–4 |
| `Keine_Doppelbuchung_Klasse` | Hart | Eine Klasse hat pro Slot maximal eine Unterrichtseinheit |
| `Keine_Doppelbuchung_Lehrer` | Hart | Eine Lehrkraft kann pro Slot nur eine Klasse unterrichten |
| `Klassenlehrer_Wochen_Start_Ende` | **Hart** | Der Klassenlehrer muss Montag Block 1 und Freitag Block 5 in seiner Klasse stehen |
| `Klassenlehrer_parallele_Foerderstunde` | Hart | Förderunterricht wird von einer Fachlehrkraft erteilt; der Klassenlehrer ist parallel anwesend (Co-Teaching, zählt für beide Deputate) |
| `Lehrer_Verfuegbarkeit_beachten` | Hart | Lehrkräfte werden nur in verfügbaren Blöcken eingeplant |
| `Nachmittagsmodell_Fachlehrer` | Hart | Fachlehrkräfte haben max. 2 GF2-Tage pro Woche |
| `Nachmittagsmodell_Klassenlehrer` | Hart | Klassenlehrkräfte haben max. 1 GF2-Tag und max. 2 GF1/GF2-Tage pro Woche |
| `Regel_Bauer` | (intern) | Schulspezifische Sonderregel |
| `Religion_Dienstag_4_5` | Hart | Religion (und Sport für 1.1) findet dienstags in Block 4+5 statt |
| `Schwimmen_Halbjahr_Logik` | Hart | Schwimmen wird gemäß Halbjahresplan auf Sporthallen-Slots beschränkt |
| `Sport_moeglichst_Doppelstunde` | Hart | Sport findet in Doppelstunden (zwei aufeinanderfolgende Blöcke) statt; gilt für Klassen 2–4 |
| `Sporthalle_nur_eine_Klasse` | Hart | Die Sporthalle kann pro Slot von maximal einer Klasse genutzt werden |
| `Wochenstunden_erfuellen` | Hart/Weich | Unterrichtsbedarf aus `05_Unterrichtsbedarf` wird vollständig verplant; Lehrkräfte nicht über ihr Wochenstunden-Soll hinaus belastet |

---

## Lehrerwünsche (10_Wuensche)

| Spalte | Bedeutung |
|---|---|
| `Kuerzel` | Lehrkraft-Kürzel |
| `Aktiv` | `ja` = Wunsch wird berücksichtigt |
| `Harte_Wuensche` | Freitext; wird geparst und in Modellbedingungen übersetzt (siehe unten) |
| `Weiche_Wuensche` | Freitext; beeinflusst Präferenz-Penalty, aber kein harter Zwang |
| `Bevorzugte_Faecher` | Bevorzugte Fächer (reduziert Penalty) |
| `Nicht_Bevorzugte_Faecher` | Nicht erwünschte Fächer (erhöht Primärkosten bei Zuweisung) |
| `Bevorzugte_Klassenstufen` | z. B. `2, 3` – Zuweisung außerhalb der Präferenz erhöht Primärkosten |
| `Beginn_1_Stunde_JN` | `nein` = möchte nicht in Block 1 beginnen |

### Geparste Harte-Wunsch-Muster

Der Parser (`parse_hard_wishes`) erkennt folgende Texteingaben und erzeugt daraus **harte Modellbedingungen**:

| Muster (Beispiel) | Wirkung |
|---|---|
| `kein langer Tag am Mittwoch` | Mittwoch wird kein UF1/UF2/GF1/GF2 eingeplant |
| `Mittwochs 1. bis 4. Stunde KOOP` | Mittwoch Blöcke 1–4 werden gesperrt |
| `Dienstag 4. und 5. Stunde evangelische Religion` | Dienstag Block 4+5 muss dieser Lehrer Religion unterrichten |
| `Deutsch Klasse 4.1` | Dieser Lehrer muss Deutsch in 4.1 unterrichten |
| `je eine Förderstunde in 4.1 und 4.3` | Dieser Lehrer muss je eine Förderstunde in 4.1 und 4.3 übernehmen |
| `jeden Tag 1.-3. Stunde in DaZ Gruppe` | Dieser Lehrer belegt Mo–Fr die Blöcke 1–3 mit DAZ |
| `Dienstag, Mittwoch, Donnerstag, Freitag, 4. Std in DaZ Gruppe` | Diese Tage Block 4 mit DAZ belegen |

Kann ein Harter Wunsch wegen fehlender Kandidaten nicht direkt modelliert werden, erscheint ein `Harte_Wuensche Fallback`-Hinweis im Blatt `Validierung`.

---

## Modellarchitektur

### Datenfluss

```
Excel-Eingabe
    └── load_data()         → Slots, Klassen, Bedarfe, Lehrer, Klassenlehrer, Regeln
    └── create_tasks()      → Liste der LessonTask-Objekte (eine Task = eine Unterrichtsstunde)
    └── compute_core_overload_by_class()  → Überhang für Klassen 3+4
    └── select_overflow_task_ids()        → Tasks, die auf GF-Slot gelegt werden dürfen
    └── solve_cp_sat()      → CP-SAT Modell aufbauen + lösen
    └── write_output()      → Excel-Ausgabe
```

### Entscheidungsvariablen

Für jede Kombination aus `(Task, Slot, Lehrer)` gibt es eine Boolesche Variable `x`. Genau eine dieser Variablen pro Task muss `True` sein (`AddExactlyOne`).

Für `Fach_Lehrer`-Fächer (alle außer Förderunterricht, DAZ, GF, UF) gibt es zusätzlich eine Variable `y` pro `(Klasse, Fach, Lehrer)`, die erzwingt, dass wochenweit nur ein Lehrer dieses Fach in dieser Klasse unterrichtet.

### Überhang-Mechanismus (Klassen 3+4)

Wenn der Unterrichtsbedarf einer Klasse der Stufe 3 oder 4 die 25 Kernstunden übersteigt, werden die Überhangstunden als `UNASSIGNED` markiert und im Plan als „Klassenlehrer legt diese Stunde auf GF" ausgegeben. Die Anzahl UNASSIGNED-Stunden pro Klasse wird exakt erzwungen.

### Förderunterricht mit Co-Teaching

- Förderunterricht wird von einer Fachlehrkraft (nicht dem Klassenlehrer) unterrichtet.
- Der Klassenlehrer ist parallel anwesend und wird für diesen Slot in seinen Zeitplan eingetragen.
- Beide Lehrkräfte zählen dafür je eine Stunde auf ihr Wochenstunden-Deputat.
- Im Blatt `Klassenplaene` erscheinen beide Lehrkräfte in einer einzigen Zeile (`Lehrer1 + Lehrer2`).

---

## Optimierungsstrategie (2-Phasen)

### Phase 1 – Primäre Ziele (40 % der Rechenzeit)
Minimiert die gewichtete Summe harter Verstöße:

| Ziel | Gewicht |
|---|---|
| Nicht zugewiesene Pflicht-Stunde (`UNASSIGNED`) | 1 000 |
| Freistunde in Kernzeit (Block 1–5) | 1 200 |
| Unterschreitung 75 % des Stundensolls je Lehrer | 800 |
| Verstoß Bevorzugte Klassenstufen | 300 |
| Verstoß Nicht-Bevorzugte Fächer | 500 |

Wenn Phase 1 keine Lösung findet (`UNKNOWN`), wird automatisch ein Fallback aktiviert: Der Solver sucht zunächst eine erste machbare Lösung ohne Zielfunktion und übergibt diese als Startpunkt für Phase 2.

### Phase 2 – Sekundäre Ziele (60 % der Rechenzeit)
Optimiert unter der Bedingung `primary_expr ≤ Phase1-Ergebnis`:

| Ziel | Gewicht |
|---|---|
| Nicht zugewiesene optionale Stunde | 20 000 |
| Freistunde in Kernzeit | 25 000 |
| Unterschreitung 75 % des Stundensolls | 15 000 |
| Unterdeckung gegenüber Wochenstunden-Soll | 1 000 |
| Fach mehrfach am selben Tag | 120 |
| Klassenlehrer Mo-1 / Fr-5 | (hart, kein Penalty) |
| Präferenz-Penalty (Soft-Wünsche) | variabel |

---

## Besondere Fächer

| Fach | Besonderheit |
|---|---|
| `Deutsch als Zweitsprache` (DAZ) | Klasse 1.4; Mo–Fr Blöcke 1–4 fix; **kein** Fach-Lehrer-Zwang; Verteilung nach Lehrerwünschen aus `10_Wuensche` |
| `Förderunterricht` | Co-Teaching (Fachlehrkraft + Klassenlehrer); **kein** Fach-Lehrer-Zwang; zählt je eine Stunde für beide Lehrkräfte |
| `Gebundene Freizeit` (GF1, GF2) | Nur in GF-Slots; optional UNASSIGNED; **kein** Fach-Lehrer-Zwang; wird frei verteilt |
| `Ungebundene Freizeit` (UF1, UF2) | Nur in UF-Slots; optional UNASSIGNED; **kein** Fach-Lehrer-Zwang; wird frei verteilt |
| `Sport` | Klassen 2–4: Doppelstunden-Pflicht; Sporthallen-Exklusivnutzung; nicht parallel zu Förderunterricht in derselben Klasse |
| `Schwimmen` | Nutzung der Sporthallen-Slots; Halbjahreslogik |
| `Religion` | Dienstag Block 4+5 (außer Klasse 1.4) |
| `Leseband` | Nur in Leseband-Slot |

---

## Bekannte Einschränkungen

- **UNKNOWN-Status**: Bei sehr kurzen Rechenzeiten (< 60 s) oder sehr vielen harten Constraints kann Phase 1 keine Lösung finden. In diesem Fall greift der Fallback (erste machbare Lösung ohne Optimierung). Die Lösung ist dann gültig, aber nicht qualitätsoptimiert.
- **Underload**: Wenn das Gesamtstunden-Angebot der Lehrkräfte den Unterrichtsbedarf übersteigt, ist eine Unterauslastung mathematisch unvermeidlich. Der Solver verteilt diese möglichst gleichmäßig.
- **DAZ-Parserfehler**: Freitexteingaben in `Harte_Wuensche` für DAZ werden nur erkannt, wenn sie dem dokumentierten Muster entsprechen (z. B. „jeden Tag 1.-3. Stunde in DaZ Gruppe").
- **Deputatsstunden**: Die Blockierung durch Deputatstunden wird über einfache Texterkennung in der Spalte `Deputatsstunden_Blocke` vorgenommen. Nicht normierte Einträge werden ggf. ignoriert.

---

## Änderungshistorie

| Datum | Änderung |
|---|---|
| 2026-07-12 | Erstversion README |
| 2026-07-12 | `Fach_Lehrer`-Constraint: UF1/UF2/GF1/GF2 von Einzel-Lehrer-Zwang befreit |
| 2026-07-12 | `Fach_Lehrer`-Constraint: DAZ von Einzel-Lehrer-Zwang befreit |
| 2026-07-12 | DAZ-Parser in `parse_hard_wishes`: „jeden Tag X.-Y. Stunde in DaZ" und Mehrtagsmuster erkannt |
| 2026-07-12 | `Klassenlehrer_Wochen_Start_Ende` von Soft-Penalty auf harte Bedingung umgestellt |
| 2026-07-12 | 2-Phasen-Strategie: Phase 1 optimiert jetzt wirklich (40 % Rechenzeit), Fallback bei UNKNOWN |
| 2026-07-12 | MUSS-Verstoss-Gewichte: `Bevorzugte_Klassenstufen` 1→300, `Nicht_Bevorzugte_Faecher` 1→500 |
| 2026-07-12 | `MIN_LOAD_RATIO` von 0,5 auf 0,75 erhöht (Mindestauslastung 75 % des Solls) |
| 2026-07-12 | `Harte_Wuensche` Slot-Anforderungen sind jetzt echte harte Bedingungen (kein Slack mehr) |
| 2026-07-12 | Förderunterricht: Co-Teaching als eine Klassenstunde modelliert; Ausgabe in einer Zeile (`L1 + L2`) |
| 2026-07-12 | `Klassenlehrer`-Spalte in `03_Klassen` als primäre Quelle für Klassenlehrkraft-Zuordnung |
| 2026-07-12 | `Fach_Lehrer`-Constraint von `<= 1` auf `== 1` verschärft |
| 2026-07-12 | `stop_after_first_solution`-Hint-Strategie für Phase-2-Start eingeführt |

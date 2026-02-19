# Matrix-NX – SCPI Serial Terminal (Python GUI)

Python-GUI zum seriellen Verbinden mit der Maschine, SCPI senden/empfangen, Logs exportieren und Parametervergleich.

## Was jetzt konkret unterstützt wird

- Laden eines **SCPI-Katalogs** aus `.xlsx`, `.xlsm`, `.csv`, `.tsv`, `.txt`.
- Erkennung typischer Spalten wie `Category`, `SCPI Command`, `Instance`, `Parameter`, `Response`, `Unit`, `Description`.
- Auswahl eines Befehls aus dem Katalog und Übernahme ins Terminal.
- Speziell für `SERVice:XALL?` und `ALL?`:
  - integrierte Zuordnung für `TEMPeratures`, `STEPper`, `OTHers` (XALL) sowie die 14er-ALL-Sequenz,
  - optional eigene Feldliste laden (eine Zeile = Feldname).
- Vergleich aktueller Werte gegen Referenzdatei (`.txt`, `.csv`, `.xlsx`) mit Markierung von Abweichungen.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Start

```bash
python scpi_terminal_gui.py
```

## Workflow

1. COM-Port verbinden.
2. `SCPI Katalog laden` (deine große Command-Tabelle).
3. Befehl auswählen und `Befehl übernehmen`.
4. Beispiel: `SERVice:XALL? TEMPeratures` senden.
5. `Aktuelle Parameter aus letztem Kommando`.
6. Referenzdatei laden, dann vergleichen.

## Hinweis XALL

Für `SERVice:XALL? TEMPeratures|STEPper|OTHers` und `ALL?` wird automatisch die feste Reihenfolge verwendet.
Wenn du ein anderes Antwortformat hast, kannst du mit `ALL/XALL Feldliste laden` eine eigene Reihenfolge vorgeben.

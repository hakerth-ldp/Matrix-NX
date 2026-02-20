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

## Troubleshooting: UPort/COM-Port unter Windows

Wenn in der App nur `COM3` auswählbar ist, im Gerätemanager aber zusätzlich ein Gerät wie `UPort 1110` unter **Andere Geräte** auftaucht, fehlt in der Regel der passende USB-Seriell-Treiber.

- `UPort 1110` unter **Andere Geräte** = Windows erkennt den Adapter, aber hat **keinen funktionierenden COM-Treiber** zugeordnet.
- Erst wenn der Treiber korrekt installiert ist, erscheint der Adapter unter **Anschlüsse (COM & LPT)** mit einem eigenen COM-Port (z. B. `COM5`).
- Solange der Adapter dort nicht auftaucht, kann die GUI ihn nicht als seriellen Port anbieten.

Kurz: Ja, du brauchst sehr wahrscheinlich den originalen Treiber für deinen Moxo/MOXA UPort RS232-auf-USB-Adapter.

### Wenn `*IDN?` nur Zeichenmüll liefert

Wenn eine Verbindung aufgebaut ist, aber Antworten wie `*IDN?` als unlesbare Zeichen erscheinen, sind fast immer die seriellen Parameter falsch (vor allem Baudrate).

Typische Ursachen:
- Falsche **Baudrate** (häufigster Grund).
- Falsche **Parität/Stopbits/Datenbits** (z. B. Gerät erwartet `8N1`).
- Falsches **Zeilenende** beim Senden (`LF`, `CR` oder `CRLF`).

Die GUI bietet dafür nun zusätzliche Einstellungen: Datenbits, Parität, Stopbits, Encoding und Zeilenende.
Diese findest du in der separaten Box **"Serielle Einstellungen"** direkt unter der COM-Verbindung.
Beim Start zeigt die App zusätzlich die Build-Anzeige (z. B. `Build: v2.1`) sowie den geladenen Script-Pfad im Log, damit du sofort siehst, dass wirklich die aktuelle Datei läuft.
Zusätzlich loggt die App jetzt `CWD`, `Git Root`, `Git Branch` und `Git Commit`, damit man sofort sieht, aus welchem Repo/Branch die laufende Instanz wirklich kommt.
Wenn dort kein Git-Repo gefunden wird, startest du sehr wahrscheinlich eine Datei außerhalb deines Projektordners.
Teste am besten zuerst diese Kombination:
- `115200`, `8`, `None`, `1` (8N1)
- Zeilenende `LF`, alternativ `CRLF`
- Encoding `ascii`

Wenn weiterhin Müll kommt, stimmt meist weiterhin Baudrate oder Parität nicht mit den Laser-Controller-Defaults überein.

Zusatz bei **Error 600** (SCPI Timeout / end of command indicator):
- Stelle zuerst **Zeilenende = `CR`** ein (bei vielen Lasern korrekt), alternativ `CRLF`.
- Prüfe im Log den Eintrag `TX HEX`:
  - `... 0d` = CR
  - `... 0a` = LF
  - `... 0d 0a` = CRLF
- Nutze den Button **IDN Terminator-Test**, um `*IDN?` automatisch mit CR/LF/CRLF/None zu senden.

## Hinweis XALL

Für `SERVice:XALL? TEMPeratures|STEPper|OTHers` und `ALL?` wird automatisch die feste Reihenfolge verwendet.
Wenn du ein anderes Antwortformat hast, kannst du mit `ALL/XALL Feldliste laden` eine eigene Reihenfolge vorgeben.

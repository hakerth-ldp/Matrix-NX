# Matrix-NX – SCPI Serial Terminal (Python GUI)

Python-GUI zum seriellen Verbinden mit der Maschine, SCPI senden/empfangen, Logs exportieren und Parametervergleich.

## Was jetzt konkret unterstützt wird

- Laden eines **SCPI-Katalogs** aus `.xlsx`, `.xlsm`, `.csv`, `.tsv`, `.txt`.
- Erkennung typischer Spalten wie `Category`, `SCPI Command`, `Instance`, `Parameter`, `Response`, `Unit`, `Description`.
- Auswahl eines Befehls aus dem Katalog und Übernahme ins Terminal.
- Serielle Verbindungsparameter in der GUI einstellbar: Baudrate, Datenbits, Parität, Stopbits, Timeout.
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

## Temperatur-Monitor SHG/THG

Es gibt den Button `monitor Temp Stage SHG/THG` mehrfach sichtbar (oben in der Verbindungszeile und unten in der Aktionszeile), damit er nicht übersehen wird.

Beim Start passiert automatisch:

1. Set-Werte lesen:
   - `SOURce:TEMPerature:LEVel:SET? TEMP_STAGE_SHG`
   - `SOURce:TEMPerature:LEVel:SET? TEMP_STAGE_THG`
2. Danach Polling im 100-ms-Takt:
   - `SOURce:TEMPerature:ACTual? TEMP_STAGE_SHG`
   - `SOURce:TEMPerature:ACTual? TEMP_STAGE_THG`
3. Logging in der Terminal-Ansicht und in CSV (Dateiname wird beim Start abgefragt).
4. Live-Plot zeigt die Abweichung `Δ = Ist - Set` für SHG/THG.

Stoppen:
- Erneut auf den Button klicken (`Stop monitor Temp Stage SHG/THG`) oder Verbindung trennen.

## Hinweis XALL

Für `SERVice:XALL? TEMPeratures|STEPper|OTHers` und `ALL?` wird automatisch die feste Reihenfolge verwendet.
Wenn du ein anderes Antwortformat hast, kannst du mit `ALL/XALL Feldliste laden` eine eigene Reihenfolge vorgeben.

## Git-Merge hängt in VS Code? (Soforthilfe, 60 Sekunden)

Wenn du im Merge „festhängst“, ist Git fast immer **nicht kaputt** – es wartet nur auf einen letzten Schritt.

### 1) Prüfen, ob ein Merge offen ist

```bash
git status
```

Wenn dort `You are currently merging` steht, dann gilt:

### 2) Merge **fertig machen** (empfohlen)

```bash
git commit --no-edit
git push origin HEAD
```

`--no-edit` nimmt die vorhandene Merge-Nachricht und umgeht den Editor komplett (kein Vim/COMMIT_EDITMSG-Stopp).

### 3) Oder Merge sauber **abbrechen**

```bash
git merge --abort
```

Danach bist du wieder auf dem Stand vor dem Merge.

### 4) Falls VS Code „Continue Merge“ geklickt wurde und Userdaten kamen

Das ist normal (Git fragt nach Benutzername/Token, wenn noch nicht gespeichert). Danach einfach nochmal:

```bash
git status
git commit --no-edit
git push origin HEAD
```

### 5) Einmalig gegen zukünftige Editor-Hänger (optional)

```bash
git config --global core.editor "code --wait"
```

Dann öffnet Git Commit-Messages künftig direkt in VS Code statt im Terminal-Editor.


## Update-Workflow ohne Branch-Chaos (nur `main`)

Nutze künftig nur diesen Block:

```bash
git fetch --all --prune
git checkout main
git reset --hard origin/main
git pull origin main
python scpi_terminal_gui.py
```

Das stellt sicher, dass lokal exakt der aktuelle `origin/main`-Stand läuft.


### Einmalbefehl für Windows (PowerShell)

Wenn du nicht jedes Mal mehrere Git-Befehle tippen willst, nutze das Script im Repo:

```powershell
powershell -ExecutionPolicy Bypass -File .\update_and_run_main.ps1
```

Das Script macht automatisch `fetch`, `checkout main`, `reset --hard origin/main`, `pull` und startet danach die GUI.

Schnellcheck auf den richtigen Stand:

```bash
git log --oneline -n 3
```

## 30-Sekunden Diagnoseblock (wenn wieder etwas hängt)

```bash
git status
git branch --show-current
git log --oneline -n 5
git remote -v
```

Mit diesen 4 Befehlen sieht man sofort: Merge offen, falscher Branch, falscher Commit oder fehlender Remote.

# Anleitung: PPTX erzeugen

In diesem Ordner findest du ein Node.js Skript, das die Präsentation erzeugt.

Dateien:
- create_pptx_simple.js — Node-Skript: erzeugt `Imperialismus_Praesentation.pptx`.

Wichtige Hinweise:
- Titelfolie: kein Name (bleibt als Platzhalter). Das Datum auf der Titelfolie wurde gesetzt auf 30.06.2026.
- Branch: export/imperialismus-praesentation

So erzeugst du die Datei lokal (Windows PowerShell):

1. Öffne PowerShell.
2. Wechsle in einen Ordner, z. B. Desktop: `cd $HOME\Desktop`.
3. Erstelle und wechsle in ein Arbeitsverzeichnis: `mkdir imperialismus-pptx; cd imperialismus-pptx`.
4. Kopiere die Datei `create_pptx_simple.js` aus diesem Repo-Ordner in dein Arbeitsverzeichnis oder lade das Repo lokal herunter.
5. Initialisiere Node-Projekt und installiere Abhängigkeit:
   ```
   npm init -y
   npm install pptxgenjs
   ```
6. Führe das Skript aus:
   ```
   node create_pptx_simple.js
   ```
7. Die Datei `Imperialismus_Praesentation.pptx` wird im aktuellen Ordner erstellt. Öffne sie mit PowerPoint.

Wenn du möchtest, kann ich die PPTX auch direkt in dein Repo hochladen — dazu benötige ich die fertige PPTX-Datei oder die Erlaubnis, eine binäre Datei hier hochzuladen. Weil ich die PPTX nicht auf meinem System ausführen kann, habe ich das Skript und die Anweisungen in dein Repository gelegt, damit du es lokal erstellst.

Bei Problemen: kopiere die Fehlermeldung aus PowerShell hierher, ich helfe Schritt für Schritt weiter.

# namitogl: Stufenlisten aus NaMi

Dieses Tool macht aus dem NaMi-Export für jede Stufe eine Excel-Liste für die Gruppenleitung. Die Listen enthalten nur Vor- und Nachname, bei Schnuppermitgliedern zusätzlich das Anmeldedatum. Schnuppermitglieder, die schon länger als sechs Wochen dabei sind, sind rot markiert.

---

## Mit der App

Die App läuft im Browser, aber nur auf deinem Computer. Du brauchst dafür kein Programmierwissen.

### Einmalig: Python installieren

Lade Python von [python.org](https://www.python.org/downloads/) herunter und installiere es mit den vorgeschlagenen Einstellungen.

### Einmalig: Tool herunterladen

Auf der GitHub-Seite des Projekts auf **Code** → **Download ZIP** klicken. Dann die ZIP-Datei **entpacken**: Rechtsklick auf die Datei → **Alle extrahieren …**. Ein Doppelklick auf `starten.bat` direkt in der ZIP-Datei funktioniert nicht.

### Starten

Doppelklick auf `starten.bat`. Beim ersten Start richtet sich das Tool ein, das dauert ein paar Minuten. Danach öffnet sich die App im Browser. Das schwarze Fenster offen lassen, solange du die App benutzt; zum Beenden einfach schließen.

Zeigt Windows eine Sicherheitswarnung, liegt das daran, dass die Datei aus dem Internet stammt. Dann das Ausführen erlauben.

### In der App

1. In NaMi einen Excel-Export der Mitglieder erstellen und die heruntergeladene `.xls`-Datei hochladen.
2. Die Listen in den Reitern durchsehen. Im Reiter **Fehler** stehen Einträge, die in NaMi korrigiert werden sollten, z. B. Mitglieder ohne Stufe.
3. **Alle Listen als ZIP herunterladen** oder einzelne Listen über die Knöpfe darunter.

Die Listen enthalten Namen von Kindern und Jugendlichen. Gib jede Liste nur an die Gruppenleitung der jeweiligen Stufe weiter.

---

## Ohne App (Kommandozeile)

```bash
pip install -r requirements.txt
python namitogl.py nami_export.xls
```

Die Listen landen im Ordner `ausgabe`. Ein anderer Ordner kann als zweites Argument angegeben werden.

Oben in `namitogl.py` lassen sich die Stufen (`STUFEN`) und die Schnupperzeit (`SCHNUPPER_WARNUNG_WOCHEN`) anpassen.

---

## Datenschutz

Das Tool läuft vollständig lokal. Es werden keine Daten ins Internet übertragen, und die App sendet keine Nutzungsstatistik an Streamlit. Die `.gitignore`-Datei sorgt dafür, dass Excel-Dateien nicht versehentlich auf GitHub landen.

Gut Pfad :)

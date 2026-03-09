# Debattenbericht-Template

Statisches Frontend für redaktionelle Debattenberichte mit:

- Eingabemasken für Titel, Untertitel, Lead, zwei Textblöcke, Bilder und Zitat
- Sofortfeedback zu Zeichenlimits, W-Fragen, Satzlängen und Aktivstil
- Browser-Rechtschreibprüfung in allen relevanten Textfeldern
- Live-Vorschau im Layout der gelieferten Debattenbericht-Vorlage
- Export als echte `.docx`-Datei
- GitHub-Pages-Deployment per Workflow

## Projektstruktur

```text
debattenbericht-template/
├── .github/workflows/pages.yml
├── config.js
├── index.html
├── language-tool-service.js
├── styles.css
├── app.js
├── README.md
├── GITHUB_METADATA.md
└── LICENSE
```

## Lokal starten

Da das Projekt rein statisch ist, reicht eines der folgenden Verfahren:

1. Datei direkt im Browser öffnen:

```bash
open index.html
```

2. Oder lokal per Mini-Server:

```bash
python3 -m http.server 4173
```

Dann im Browser `http://localhost:4173` aufrufen.

## Funktionslogik

- `Titel`: maximal 50 Zeichen.
- `Untertitel`: maximal 80 Zeichen; das Feedback prüft, ob Debattenthema und Vornamen der Debattierenden auftauchen.
- `Lead`: maximal 500 Zeichen; das Feedback sucht heuristisch nach Wer, Was, Wann, Wo, Warum und einer Bilanz.
- `Block 1` und `Block 2`: je maximal 3000 Zeichen; zusätzlich Hinweise zu langen Sätzen und Passivmustern.
- `Bildlegende Debattenbild`: maximal 100 Zeichen.
- `DOCX-Export`: erzeugt eine echte `.docx`-Datei mit eingebetteten PNG-/JPG-Bildern.
- `LanguageTool`: prüft Texte gegen einen selbst gehosteten HTTP-Server.
- `GitHub Pages`: Das Repo ist für statisches Hosting vorbereitet.

## Bildformate

Für den DOCX-Export sind aktuell `PNG` und `JPG` vorgesehen. Diese Formate werden direkt in das Word-Dokument eingebettet.

## LanguageTool lokal starten

Die App ist statisch und spricht einen externen LanguageTool-HTTP-Server direkt aus dem Browser an.
Die Basis-URL steht zentral in `config.js`.

Default:

```js
languageToolBaseUrl: "http://localhost:8081/v2/check"
languageToolLanguage: "de-DE"
```

Lege bei Bedarf zuerst eine leere Datei `server.properties` an. Typischer lokaler Start auf Port `8081`:

```bash
languagetool --http --config server.properties --port 8081 --allow-origin "*"
```

Falls du LanguageTool nicht als CLI im PATH hast, starte den HTTP-Server über Java:

```bash
java -cp languagetool-server.jar org.languagetool.server.HTTPServer --config server.properties --port 8081 --allow-origin
```

Danach:

1. Die URL in `config.js` bei Bedarf anpassen.
2. Diese App lokal starten, z. B. mit `python3 -m http.server 4173`.
3. Im Browser auf `http://localhost:4173` gehen.
4. In der Redaktionsmaske den Button `Text prüfen` verwenden.

## CORS-Hinweis

Da dieses Unterprojekt kein eigenes Backend hat, muss der LanguageTool-Server Requests vom Origin der App akzeptieren.
Für lokale Tests ist `--allow-origin "*"` der einfachste Weg. Ohne passende CORS-Freigabe blockiert der Browser die Anfrage, obwohl der Server läuft.

## GitHub-Repo vorbereiten

Wenn du dieses Unterprojekt als eigenes Repository veröffentlichen willst:

```bash
cd debattenbericht-template
git init
git add .
git commit -m "Initial commit for debate report template"
git branch -M main
git remote add origin <DEIN_GITHUB_REPO>
git push -u origin main
```

Danach kannst du in GitHub unter `Settings > Pages` den Eintrag `GitHub Actions` aktivieren. Der Workflow unter `.github/workflows/pages.yml` übernimmt dann das Deployment automatisch.

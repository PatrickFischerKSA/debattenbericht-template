# Debattenbericht-Template

Statisches Frontend fuer redaktionelle Debattenberichte mit:

- Eingabemasken fuer Titel, Untertitel, Lead, zwei Textbloecke, Bilder und Zitat
- Sofortfeedback zu Zeichenlimits, W-Fragen, Satzlaengen und Aktivstil
- Browser-Rechtschreibpruefung in allen relevanten Textfeldern
- Live-Vorschau im Layout der gelieferten Debattenbericht-Vorlage
- Export als echte `.docx`-Datei
- GitHub-Pages-Deployment per Workflow

## Projektstruktur

```text
debattenbericht-template/
├── .github/workflows/pages.yml
├── index.html
├── styles.css
├── app.js
├── README.md
├── GITHUB_METADATA.md
└── LICENSE
```

## Lokal starten

Da das Projekt rein statisch ist, reicht eines der folgenden Verfahren:

1. Datei direkt im Browser oeffnen:

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
- `Untertitel`: maximal 80 Zeichen; das Feedback prueft, ob Debattenthema und Vornamen der Debattierenden auftauchen.
- `Lead`: maximal 500 Zeichen; das Feedback sucht heuristisch nach Wer, Was, Wann, Wo, Warum und einer Bilanz.
- `Block 1` und `Block 2`: je maximal 3000 Zeichen; zusaetzlich Hinweise zu langen Saetzen und Passivmustern.
- `Bildlegende Debattenbild`: maximal 100 Zeichen.
- `DOCX-Export`: erzeugt eine echte `.docx`-Datei mit eingebetteten PNG-/JPG-Bildern.
- `GitHub Pages`: Das Repo ist fuer statisches Hosting vorbereitet.

## Bildformate

Fuer den DOCX-Export sind aktuell `PNG` und `JPG` vorgesehen. Diese Formate werden direkt in das Word-Dokument eingebettet.

## GitHub-Repo vorbereiten

Wenn du dieses Unterprojekt als eigenes Repository veroeffentlichen willst:

```bash
cd debattenbericht-template
git init
git add .
git commit -m "Initial commit for debate report template"
git branch -M main
git remote add origin <DEIN_GITHUB_REPO>
git push -u origin main
```

Danach kannst du in GitHub unter `Settings > Pages` den Eintrag `GitHub Actions` aktivieren. Der Workflow unter `.github/workflows/pages.yml` uebernimmt dann das Deployment automatisch.

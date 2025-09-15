# ppt-to-glypheo

Extrahiert zweisprachige Texte aus PowerPoint-Dateien (`.pptx`) und schreibt sie in eine CSV-Datei. Der deutschsprachige Text wird anhand einer definierten Farbe erkannt (Standard: weiß), andersfarbige Abschnitte werden als Englisch exportiert.

## Voraussetzungen
- Python 3.8+
- Abhängigkeiten: `python-pptx`, `Pillow`

```bash
pip install python-pptx pillow
```

## Verwendung

```bash
python app.py INPUT.pptx OUTPUT.csv [Optionen]
```

### Wichtige Optionen
- `--de-color` HEX[,HEX...] – Farbe(n) für Deutsch (Standard: FFFFFF)
- `--en-color` HEX[,HEX...] – Farbe(n) für Englisch (Standard: alle nicht-deutschen)
- `--tolerance` N – Farbtoleranz je Kanal (Standard: 8)
- `--unknown-policy {german,english,skip}` – Verhalten für Runs ohne explizite Farbe
- `--interactive` – interaktives Zuordnen der Farben

Die CSV besitzt immer den Aufbau `FOLIENNUMMER;DEUTSCH;ENGLISCH` und wird mit UTF‑8‑BOM sowie Semikolon als Trennzeichen geschrieben. Leere Folien erzeugen eine leere Zeile, um die Reihenfolge zu erhalten.

## Entwicklung

```bash
python app.py INPUT.pptx OUTPUT.csv
```

## Lizenz

MIT – siehe `LICENSE`.

# ppt-to-xlsx-glypheo

Extrahiert zweisprachige Texte aus PowerPoint-Dateien (`.pptx`) und schreibt sie wahlweise in eine CSV- oder XLSX-Datei. Der deutschsprachige Text wird anhand einer definierten Farbe erkannt (Standard: weiß), andersfarbige Abschnitte werden als Englisch exportiert.

## Voraussetzungen
- Python 3.8+
- Abhängigkeiten: `python-pptx`, `Pillow`, optional `openpyxl` für XLSX-Ausgabe

```bash
pip install python-pptx pillow
```

## Verwendung

```bash
python app.py INPUT.pptx OUTPUT.csv|OUTPUT.xlsx [Optionen]

# oder gesamter Ordner:
python app.py ORDNER auto [Optionen]
```

### Wichtige Optionen
- `--de-color` HEX[,HEX...] – Farbe(n) für Deutsch (Standard: FFFFFF)
- `--en-color` HEX[,HEX...] – Farbe(n) für Englisch (Standard: alle nicht-deutschen)
- `--tolerance` N – Farbtoleranz je Kanal (Standard: 8)
- `--unknown-policy {german,english,skip}` – Verhalten für Runs ohne explizite Farbe
- `--interactive` – interaktives Zuordnen der Farben
- `--xlsx` – Im `auto`-Modus bzw. bei Ausgabeverzeichnissen XLSX statt CSV schreiben

Die Ausgabedatei besitzt immer den Aufbau `FOLIENNUMMER;DEUTSCH;ENGLISCH`. Bei CSV wird UTF‑8‑BOM sowie Semikolon als Trennzeichen verwendet. Leere Folien erzeugen eine leere Zeile, um die Reihenfolge zu erhalten.

## Entwicklung

```bash
python app.py INPUT.pptx OUTPUT.xlsx
```

## Lizenz

MIT – siehe `LICENSE`.

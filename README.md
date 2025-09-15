# pptx\_to\_csv\_by\_color

> Extract bilingual slide text from PowerPoint (`.pptx`) into a CSV based on **font colors**.

**Default behavior**

* **CSV format (always):** `FOLIENNUMMER;DEUTSCH;ENGLISCH` (with header)
* **Empty slides:** exported as an **empty row** to preserve slide order
* **German = white** text (`#FFFFFF`), **English = every non‑German color** (unless `--en-color` is specified)
* **Mixed-color paragraphs are split** so each colored segment goes to the matching column
* **Placeholders** (slide number, date, header/footer) are **ignored by default**; slide numbers can be included explicitly

Also see the German version: **[README\_de.md](README_de.md)**

---

## Quick Start

```bash
pip install python-pptx pillow

# Default: German = white (#FFFFFF), English = all other colors, placeholders ignored
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv

# Process all .pptx files in a folder with auto-named CSVs
python pptx_to_csv_by_color.py /path/to/folder Auto
```

**Windows / PowerShell example**

```powershell
python .\pptx_to_csv_by_color.py ".\alt\Am Ende_Uebertitel_Premiere.pptx" ".\alt\Am_Ende.csv"
```

---

## Installation

* **Requirements:** Python 3.8+
* **Dependencies:** `python-pptx`, `Pillow`

```bash
pip install python-pptx pillow
```

---

## Usage

```bash
python pptx_to_csv_by_color.py INPUT.pptx|FOLDER OUTPUT.csv|Auto [options]
```

`INPUT` may be a single PowerPoint file or a folder. When `OUTPUT` is `Auto`,
the tool writes each CSV next to its source `.pptx` using the same base name.

### Common options

| Option                                   | Description                                                                                |
| ---------------------------------------- | ------------------------------------------------------------------------------------------ |
| `--interactive`                          | Scan the deck and interactively map detected colors to German/English.                     |
| `--de-color HEX[,HEX...]`                | Explicit color(s) for **German** (default: `FFFFFF`). Multiple hex values comma-separated. |
| `--en-color HEX[,HEX...]`                | Explicit color(s) for **English** (default: **all non‑German** colors).                    |
| `--tolerance N`                          | Per-channel RGB tolerance (0–255). Default: `8`.                                           |
| `--unknown-policy {german,english,skip}` | Assign runs with **no explicit RGB** (theme/auto). Default: `german`.                      |
| `--no-skip-placeholders`                 | Do not skip date/footer/header placeholders (slide numbers remain excluded by default).    |
| `--include-slide-numbers`                | Also include **slide-number** placeholders as text (optional, for special cases).          |

> The tool always writes a header and uses **semicolon** (`;`) as CSV delimiter with **UTF‑8‑BOM** (`utf-8-sig`) for easy Excel import.

---

## CSV format

```
FOLIENNUMMER;DEUTSCH;ENGLISCH
1;"Bitte schalten Sie Ihr Mobiltelefon aus. Bild- und Tonaufnahmen sind nicht gestattet.";"Please switch off your cell phone. Picture and sound recordings are not allowed."
2;;
3;"Nur Deutsch";"Only English on other slides"
```

* Slide numbers are **1-based** and follow PowerPoint order.
* If a slide contains multiple paragraphs per language, they are joined with line breaks (`\n`) inside the cell.

---

## Color mapping workflow (`--interactive`)

1. The tool scans all runs with explicit RGB colors and lists them with usage counts.
2. You assign which colors represent **German** and (optionally) **English**.
3. Any text colored with a mapped German color → **DEUTSCH**; any other colored text (or any explicitly mapped English color) → **ENGLISCH**.
4. Runs with *no* explicit RGB (theme inheritance) follow `--unknown-policy`.

Example:

```bash
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv --interactive
```

---

## Placeholders & slide numbers

* By default, **slide-number placeholders are ignored**, preventing stray digits like `… allowed 2`.
* If you *do* want them, pass:

  ```bash
  python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv --include-slide-numbers --no-skip-placeholders
  ```
* Defensive cleanup additionally removes a trailing, isolated slide number that may have been inserted into a text line.

---

## Examples

**Explicit colors**

```bash
# German = white; English = amber + light amber
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv --de-color FFFFFF --en-color FFC000,FFD966
```

**Higher tolerance for slight color variations**

```bash
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv --tolerance 16
```

**Treat theme/auto colors as English**

```bash
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv --unknown-policy english
```

**Intentionally include placeholders (incl. slide numbers)**

```bash
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv --no-skip-placeholders --include-slide-numbers
```

---

## Troubleshooting

* **“Stray digits” (e.g., `… 2`)**
  Usually a slide number. By default the tool ignores slide numbers. If you used `--no-skip-placeholders`, either remove it or add `--include-slide-numbers` only when you truly need them. A defensive cleanup also strips trailing isolated slide numbers.

* **Everything ends up in German**
  Likely due to theme/auto colors without explicit RGB. Use `--interactive`, specify `--en-color …`, or set `--unknown-policy english`.

* **Excel shows all in one column**
  The file is written with `;` as delimiter and `UTF‑8‑BOM`. Make sure your import settings match.

* **Expected color not recognized**
  Increase `--tolerance` (e.g., `16`) or define the color explicitly via `--de-color` / `--en-color`.

---

## Development

```bash
# Run from source
python pptx_to_csv_by_color.py INPUT.pptx OUTPUT.csv
```

Contributions and bug reports are welcome. Please include:

* a minimal `.pptx` sample,
* your command line,
* expected vs. actual CSV rows.

---

## Changelog (highlights)

* **rev3**: always ignore slide-number placeholders by default; `--include-slide-numbers` to include; defensive removal of trailing slide numbers.
* **rev2**: fix English detection when `--en-color` is unset (English = non‑German); placeholder skip option.
* **rev1**: initial release with interactive color mapping and fixed CSV header.

---

## License

MIT — see `LICENSE`.

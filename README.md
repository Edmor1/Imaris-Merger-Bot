# Imaris Merger Bot

A tool that takes the `.xls` files Imaris produces for each dendrite and merges them into a single overview Excel workbook with two sheets: **IMARIS RAW DENDRITES** and **IMARIS RAW SPINES**.

## What it does

For each Imaris `.xls` file in a folder, the bot:

- Reads the **Average** sheet (per-dendrite measurements) and copies it into the dendrite overview
- Reads the **Algorithm** sheet and pulls out 6 parameters (diameter thresholds, spine settings, etc.)
- Reads the **Spines** sheet (per-spine-type measurements) and copies it into the spine overview
- Auto-fills the blinded name, dendrite number, and ROI from the filename
- Highlights the first row of each new dendrite in yellow so it's easy to scroll through
- Skips any dendrite that's already in the output, so you can re-run as many times as you like

What it leaves blank for you to fill in manually:

- Animal ID
- Slide number

## What you need

- A Mac or PC
- Python 3 installed ([download from python.org](https://www.python.org/downloads/))
- About 10 minutes for first-time setup

You don't need any prior coding knowledge.

## Quick start

The full step-by-step tutorial is in [`imaris_bot_tutorial.md`](./imaris_bot_tutorial.md).

In short:

1. Put `merge_imaris.py` in a folder on your computer
2. Make a sub-folder called `inputs` inside that folder
3. Rename your Imaris `.xls` files to follow this pattern: `<name> <dendrite number> <roi>.xls`
   - For example: `nina 1 oriens.xls`, `kiara 3 radiatum.xls`, `mary poppins 2 oriens.xls`
4. Drop them in the `inputs` folder
5. Open Terminal, go to your folder, and run:
   ```
   python3 merge_imaris.py Overview_updated.xlsx inputs
   ```

A new `Overview_updated.xlsx` file will then appear with all your data merged. If the code has already been run and created the `Overview_updated.xlsx` file it will just update that file.

## Filename rules

The bot reads the name, dendrite number, and ROI **from the filename**, so the format matters. Examples:

| Filename | Becomes |
|---|---|
| `nina 1 oriens.xls` | Nina, 001, Oriens |
| `kiara 3 radiatum.xls` | Kiara, 003, Radiatum |
| `mary poppins 2 oriens.xls` | Mary Poppins, 002, Oriens |
| `snow white 4 oriens.xls` | Snow White, 004, Oriens |

The name can be any word(s), the number is the dendrite number, and the ROI is whatever you write after the number. Underscores or hyphens work in place of spaces if you prefer (`nina_1_oriens.xls`).

## Adding more data later

When you have new files to add, just:

1. Rename them in the same format as the others
2. Drop them into the `inputs` folder alongside the old ones
3. Run the same command again

The bot remembers what's already in the output and only writes new data. Old files staying in the folder won't cause duplicates.

## Output

The bot creates an Excel file with two sheets matching the standard format Dinia and I used in her Omega-3 project:

- **IMARIS RAW DENDRITES** — all per-dendrite data, one row per variable
- **IMARIS RAW SPINES** — all per-spine-type data, one row per variable × spine type (Stubby, Mushroom, Long Thin, Filopodia)

Numeric cells use Imaris's standard light-green fill and 2-decimal-place display, purely for continuity. The Animal ID column is set up so typing a number like `5` displays as `005`.

## Troubleshooting

The bot prints a clear summary at the end of every run. If anything goes wrong, it shows up in a `PROBLEMS` section with the filename and what went wrong. Common fixes are covered in the [tutorial](./imaris_bot_tutorial.md#troubleshooting).

## Built with

- [Python](https://www.python.org/) 3
- [pandas](https://pandas.pydata.org/) — for reading the Imaris XLS files
- [openpyxl](https://openpyxl.readthedocs.io/) — for writing the formatted overview workbook
- [xlrd](https://xlrd.readthedocs.io/) — required to read older `.xls` files

## Background

I built this during my placement year in Oslo, where the lab uses Imaris to quantify dendritic spine density in a mouse dietary intervention study. Each mouse has multiple dendrites imaged, each of which produce its own XLS export, these were being copied into the master spreadsheet by hand, which is slow and error-prone.

The tool is general enough to work with any XLS file produced by an Imaris export following our lab's standard protocol.

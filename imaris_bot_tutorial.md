# Imaris Bot — Tutorial

Python script that takes Imaris `.xls` exports and merges them into a single Excel workbook with two sheets: **IMARIS RAW DENDRITES** and **IMARIS RAW SPINES**.

This tutorial covers everything from a clean Mac with no tools installed through to a finished merged Excel workbook.

---

## Table of contents

1. [What you'll end up with](#what-youll-end-up-with)
2. [Prerequisites](#prerequisites)
3. [One-time setup](#one-time-setup)
4. [Preparing your input files](#preparing-your-input-files)
5. [Running the bot](#running-the-bot)
6. [Reading the summary](#reading-the-summary)
7. [Adding more data later](#adding-more-data-later)
8. [Troubleshooting](#troubleshooting)
9. [What the bot actually does](#what-the-bot-actually-does)

---

## What you'll end up with

A folder on your desktop like this:

```
Desktop/imaris_bot/
├── merge_imaris.py              ← the script
├── Overview_updated.xlsx        ← your merged output (created on first run)
└── inputs/
    ├── nina 1 oriens.xls
    ├── sarabi 1 oriens.xls
    ├── kiara 3 radiatum.xls
    └── ...
```

You drop new `.xls` files into `inputs/`, run the script from Terminal, and the output workbook fills up. Already-processed dendrites are automatically skipped on re-runs, so you can run it as many times as you like.

---

## Prerequisites

You need:

- A Mac (this tutorial is written for macOS; Windows works too but the commands differ slightly)
- About 10 minutes for first-time setup
- Imaris `.xls` export files from your dendrite measurements

---

## One-time setup

### Step 1 — Install Python

Most Macs come with Python pre-installed, but it's old. Install a fresh version from [python.org](https://www.python.org/downloads/) — download the latest macOS installer and run it with default settings.

To check it worked, open **Terminal** (Cmd + Space, type `Terminal`, hit Enter) and type:

```
python3 --version
```

You should see something like `Python 3.14.0`. If you see `command not found`, restart Terminal and try again.

### Step 2 — Install the Python packages the script needs

In Terminal, paste this and hit Enter:

```
pip3 install pandas openpyxl xlrd
```

This installs three small libraries. You'll see a lot of text scroll past — as long as the last line says something like "Successfully installed..." you're good. This only needs to be done once.

### Step 3 — Create the project folder

On your Desktop, make a new folder called `imaris_bot`. (No spaces, underscore is fine.)

Inside that folder, make another folder called `inputs`.

So you now have:

```
Desktop/imaris_bot/
└── inputs/
```

### Step 4 — Save the script

Save `merge_imaris.py` into the `imaris_bot` folder (**not** inside `inputs`).

You're done with setup. You won't need to repeat any of these steps again unless you get a new computer.

---

## Preparing your input files

The bot reads the blinded name, dendrite number, and ROI (region of interest) **from the filename** — so the filenames have to follow a specific pattern.

### The filename pattern

```
<name> <number> <roi>.xls
```

### Examples

| Filename | Parses as |
|---|---|
| `nina 1 oriens.xls` | Nina, dendrite 001, Oriens |
| `kiara 3 radiatum.xls` | Kiara, dendrite 003, Radiatum |
| `mary poppins 2 oriens.xls` | Mary Poppins, dendrite 002, Oriens |
| `snow white 4 oriens.xls` | Snow White, dendrite 004, Oriens |

### Rules

- The **name** can be anything (one or more words). It's the text before the first number.
- The **number** is the dendrite number. It gets zero-padded (`1` → `001`) automatically.
- The **ROI** is whatever comes after the number. It can be `oriens`, `radiatum`, or anything else — whatever you type will appear in the ROI column of the output.
- Separators can be spaces, underscores, or hyphens. `nina 1 oriens.xls`, `nina_1_oriens.xls`, and `nina-1-oriens.xls` all work identically.
- Capitalisation doesn't matter in the filename — the bot capitalises everything in the output.

### Rename your files

Go through your Imaris save files, copy them into the inputs folder and rename each file to follow this pattern.

---

## Running the bot

Open Terminal.

### Step 1 — Navigate to the project folder

Type this and hit Enter:

```
cd ~/Desktop/imaris_bot
```

The prompt should now end in `imaris_bot %`. That means you're in the right place.

### Step 2 — Run the script

Type this and hit Enter:

```
python3 merge_imaris.py Overview_updated.xlsx inputs
```

What this means:
- `python3` — run Python
- `merge_imaris.py` — the script file
- `Overview_updated.xlsx` — the output file (it'll be created if it doesn't exist)
- `inputs` — the folder containing your `.xls` files

### What you'll see

Each file gets processed in turn, with its parsed name/number/ROI printed, then how many rows were written. At the end there's a summary.

Example:

```
Creating fresh overview at Overview_updated.xlsx

nina 1 oriens.xls
  parsed as: Nina 001 Oriens
  dendrites: wrote 103 rows
  spines:    wrote 412 rows

sarabi 2 oriens.xls
  parsed as: Sarabi 002 Oriens
  dendrites: wrote 103 rows
  spines:    wrote 412 rows

Saved to Overview_updated.xlsx

============================================================
SUMMARY
============================================================
Files processed:            2
Dendrites written (files):  2  (206 rows)
Spines written (files):     2  (824 rows)

No problems encountered.
============================================================
```

Open `Overview_updated.xlsx` in Excel to see your data.

---

## Reading the summary

After every run the bot prints a summary. Here's how to read it:

### Counts

```
Files processed:            12
Dendrites written (files):  10  (1030 rows)
Spines written (files):     10  (4120 rows)
```

- **Files processed** — every file in the input folder
- **Dendrites / Spines written** — how many new entries were added to each sheet

### Already-present skips

```
Dendrites skipped as already-present: 2
  - Nina 001 Oriens (nina 1 oriens.xls)
  - Sarabi 002 Oriens (sarabi 2 oriens.xls)
```

This isn't an error, it's just a note that the bot saw these dendrites were already in the output workbook and skipped them. Which is the expected behaviour on re-runs.

### PROBLEMS

```
PROBLEMS (1):
  - weird_file.xls: filename unparseable -- ...
```

**This section is the one to pay attention to.** Each entry tells you which file had a problem and what went wrong. Common issues and how to fix them are in the [Troubleshooting](#troubleshooting) section.

If everything went fine, you'll see `No problems encountered.` instead.

---

## Adding more data later

When you have new `.xls` files to add:

1. **Rename them** to follow the `<name> <number> <roi>.xls` pattern (see [Preparing your input files](#preparing-your-input-files))
2. **Drop them into the `inputs` folder** alongside your existing files
3. **Run the same command** as before:
   ```
   cd ~/Desktop/imaris_bot
   python3 merge_imaris.py Overview_updated.xlsx inputs
   ```

The bot will skip everything that's already in the output workbook, and add only the new dendrites. The old files can stay in the inputs folder — they won't cause duplicates.

### If you want a fresh start

Delete `Overview_updated.xlsx` and re-run the command. A new empty workbook will be created and every file in the `inputs` folder will be processed from scratch.

---

## Troubleshooting

### `command not found: python3`

Python isn't installed or your Terminal hasn't picked it up yet. Try closing and reopening Terminal, or re-run the Python installer from [python.org](https://www.python.org/downloads/).

### `No such file or directory: 'Overview_updated.xlsx'`

You're not in the right folder in Terminal. Run `pwd` to see where you are, then `cd ~/Desktop/imaris_bot` to get to the right place.

### `No .xls/.xlsx files found in inputs`

Your input folder is empty, or you're pointing at the wrong folder. Make sure the `inputs` folder exists inside `imaris_bot` and has your `.xls` files in it.

### In the summary: "filename unparseable"

The filename doesn't follow the `<name> <number> <roi>` pattern. Rename the file and re-run.

### In the summary: "no sheet matching 'Spines'" or similar

One of your Imaris export files is missing a sheet (or it's been renamed to something unrecognisable). Open the file in Excel, look at the tabs at the bottom, and either:
- Rename the tab to match what's expected (`Average`, `Algorithm`, or `Spines`), **or**
- Accept that part of the data is missing and move on. Dendrite data will still load even if spine data fails, and vice versa.

### Two dendrites with the same name and number get merged

The bot treats `(Name, Dendrite number)` as a unique key. If you have, say, two files both parsing as `Cinderella 001`, only the first one gets written and the second is skipped as a duplicate. Check your filenames — usually this means a typo in one of them. Rename the duplicate properly, delete the incorrectly-written rows from the Excel file, and re-run.

### I want to see what went wrong for a specific file

Just scroll up in the Terminal output — every file gets its own section with any errors printed inline, not just in the summary.

---

## What the bot actually does

For each `.xls` file in your inputs folder, the bot:

1. **Parses the filename** into name, dendrite number, and ROI.
2. **Reads the `Average` sheet** from the Imaris export — this has one row per variable (Dendrite Length, Mean Diameter, Spine Density, etc.).
3. **Reads the `Algorithm` sheet** and pulls out 6 parameters (Dendrite Diameter Threshold, Spine Seed Point Diameter, Spine Maximum Length, Spine Seed Point Threshold, Spine Diameter Threshold, Spine Diameter Algorithm).
4. **Reads the `Spines` sheet** — this has one row per variable × spine type (Stubby, Mushroom, Long Thin, Filopodia).
5. **Writes new rows** into the two overview sheets, with the name/number/ROI filled in, the 6 algorithm parameters on the first row for that dendrite only, and the full data copied over.
6. **Applies formatting** so numeric cells get a light-green fill and 2-decimal-place display, matching the Imaris convention.
7. **Tracks which dendrites are already in the output** so re-runs don't duplicate anything.

### What stays blank

Two columns aren't populated by the bot — they're left blank for you to fill in manually:

- **Animal ID** — the bot has no way to know which blinded name maps to which real animal
- **Slide number** — not recorded in the Imaris export

### Sheet matching is forgiving

The bot tolerates:
- Case differences (`Algorithm` / `algorithm`)
- Whitespace (` Algorithm `)
- Small typos (`agorithm` will still match)

So you don't have to clean up the Imaris exports before running.

---

## Summary of the full workflow

Once set up, your routine is:

```
1. Rename new .xls files: "<name> <number> <roi>.xls"
2. Drop them into ~/Desktop/imaris_bot/inputs/
3. Open Terminal, run:
     cd ~/Desktop/imaris_bot
     python3 merge_imaris.py Overview_updated.xlsx inputs
4. Read the summary.
5. Open Overview_updated.xlsx in Excel.
```

That's it.

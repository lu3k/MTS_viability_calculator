# MTS Viability Data Builder

Fills in the calculated absorbance/viability blocks in this lab's MTS-assay
Excel templates, given a fresh raw export straight off the Tecan Spark.
Every calculated cell is written as a **live Excel formula**, so opening the
result in Excel and editing a raw value recalculates everything downstream.

Two ways to use it — pick whichever fits:

## Web app (recommended)

No install, runs entirely in your browser. Nothing is ever uploaded.

```sh
cd webapp
npm install
npm run dev
```

Open the printed local URL, upload a raw export, and download the
processed file. See [`webapp/README.md`](webapp/README.md) for details
(single file vs. batch, dose axis options, medium-control references).

## CLI (Python)

For scripting or automating a batch of files.

```sh
pip install -r requirements.txt
python build_viability_data.py RAW.xlsx
```

```sh
# override drug / cell line labels (rows B-D vs. rows E-G)
python build_viability_data.py RAW.xlsx --drug1 Brigatinib --drug2 Crizotinib --cell-line1 "MOLM13 resistant"
python build_viability_data.py RAW.xlsx --drug1 Brigatinib --cell-line1 "MOLM13 parental" --cell-line2 "MOLM13 resistant"

# batch mode: process every *.xlsx in a directory
python build_viability_data.py path/to/directory
```

Run `python build_viability_data.py --help` for the full flag list
(medium-control overrides, output path, etc).

## What it expects

A fresh raw export with:

- an anchor cell `"<>"` marking the top-left of the 96-well grid
- rows B–D / E–G: two drugs, 10-point dilution series, triplicate
- row F/G, plate column 12: two medium-control (blank) wells

From that, it fills in 4 calculated blocks to the right of the raw grid —
raw absorbance reshaped, corrected absorbance, % viability, and %
viability corrected — matching an existing template's cell layout exactly.

See the docstring at the top of [`build_viability_data.py`](build_viability_data.py)
for the full layout writeup.

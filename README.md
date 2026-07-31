# MTS Viability Data Builder

Fills in the calculated absorbance/viability blocks in this lab's MTS-assay
Excel templates, given a raw Tecan Spark export to be pasted in Prism.
Every calculated cell is written as a **live Excel formula**.

## Web app
```sh
cd webapp & npm install
npm run dev
```
or
```sh
npm run build   # static output in dist/, deploy anywhere
```


## CLI (Python3)

Install 
```sh
pip install -r requirements.txt

# Single file : 
python build_viability_data.py RAW.xlsx
# Batch directory : 
python build_viability_data.py path/to/directory
```

override drug / cell line labels (rows B-D (=1) vs. rows E-G (=2))
```sh
# 2 different drugs
python build_viability_data.py RAW.xlsx --drug1 DRUG_NAME1 --drug2 DRUG_NAME2 --cell-line1 CELL_LINE_NAME
# 2 different cellines 
python build_viability_data.py RAW.xlsx --drug1 DRUG_NAME1 --cell-line1 CELL_LINE_NAME1 --cell-line2 CELL_LINE_NAME2
´´´

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

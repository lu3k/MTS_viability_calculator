# MTS Viability Data Builder (web app)

Browser port of `build_viability_data.py`. Fills in the same 4 calculated
data blocks (raw absorbance reshaped, corrected absorbance, % viability, %
viability corrected) as live Excel formulas, given a fresh raw Tecan Spark
export — no Python needed. All processing happens client-side in the
browser (via `exceljs`); no file is ever uploaded to a server.

See `src/lib/buildViabilityData.ts` for the ported calculation logic, and
`../build_viability_data.py` for the original layout writeup it's based on.

## Run locally

```sh
npm install
npm run dev
```

## Build for production

```sh
npm run build
```

Outputs a static site in `dist/` — can be hosted anywhere that serves
static files (no backend required).

## Usage

- **Single file**: pick one raw `.xlsx` export, optionally set drug 1/drug
  2/cell line labels and override the two medium control values, then
  download the processed file.
- **Batch**: pick multiple files (or a whole folder via the second file
  picker), optionally override the medium control values for the whole
  batch, and download a `.zip` of all processed files. Cell line/drug
  labels are left as placeholders in batch mode, same as the CLI.

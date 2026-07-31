# MTS Viability Data Builder

Fills in the calculated absorbance/viability blocks in this lab's MTS-assay
Excel templates from a raw Tecan Spark export — no Python needed. Runs
entirely in your browser; nothing is uploaded anywhere.

## Quick start

```sh
npm install
npm run dev
```

Open the printed local URL, upload a raw export, and download the
processed file.

```sh
npm run build   # static output in dist/, deploy anywhere
```

## How to use

1. **Assay settings** — pick a dose axis (dilution series, or type your own
   concentrations) and, optionally, the medium-control wells (Excel cell
   refs or Tecan well numbers; defaults to the plate's own wells).
2. **Files** — choose **Single file** or **Batch**:
   - *Single file*: set drug 1/2 and cell line 1/2 (same or different per
     row block), pick a `.xlsx`, then **Build & download**.
   - *Batch*: pick multiple files or a folder, then **Build & download
     all** for a zip. Drug/cell line labels stay as placeholders — edit by
     hand afterward.

## More detail

- `src/lib/buildViabilityData.ts` — the calculation logic (ported from
  `../build_viability_data.py`, the original CLI).
- Every calculated cell is a live Excel formula, so opening the result in
  Excel and editing a raw value recalculates everything downstream.

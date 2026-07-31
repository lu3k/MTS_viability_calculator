import ExcelJS from "exceljs";

/**
 * Port of build_viability_data.py's build() logic to the browser.
 * Fills in the 4 calculated data blocks that sit to the right of the raw
 * plate grid in this lab's MTS-assay Result sheets, given a fresh raw
 * export straight off the Tecan Spark. See build_viability_data.py for the
 * full layout writeup this was reverse-engineered from.
 */

export const PLACEHOLDER_DRUG1 = "DRUG1";
export const PLACEHOLDER_DRUG2 = "DRUG2";
export const PLACEHOLDER_CELL_LINE = "CELL_LINE";

export const TOP_DOSE_NM = 10000.0;
export const DILUTION_FACTOR = 2;
// The raw plate only ever has 10 physical dilution columns per drug (plate
// columns 2-11 of the 96-well grid), so this isn't user-configurable the
// way top dose / dilution factor / manual concentrations are - the well
// reshaping in build() below assumes exactly this many raw columns exist.
export const N_DOSES = 10;
export const DOSE_UNIT = "nM";

/** A geometric dilution series computed from a top dose, written as live formulas. */
export interface DoseAxisSeries {
  mode: "series";
  topDose: number;
  dilutionFactor: number;
}

/** An explicit, arbitrary list of N_DOSES concentrations (low -> high), written as literals. */
export interface DoseAxisManual {
  mode: "manual";
  concentrations: number[];
}

export type DoseAxis = DoseAxisSeries | DoseAxisManual;

export interface BuildOptions {
  /** Drug 1 label, rows B-D. Pass the same value as drug2 if both blocks are the same drug. */
  drug1?: string;
  /** Drug 2 label, rows E-G. */
  drug2?: string;
  /** Cell line label, rows B-D. Pass the same value as cellLine2 if both blocks are the same cell line. */
  cellLine1?: string;
  /** Cell line label, rows E-G. */
  cellLine2?: string;
  /** Comma-separated list of medium-control wells: each item is either a
   * literal number (e.g. "0.223") or a reference resolved according to
   * mediumControlRefMode (e.g. "F47" as an Excel cell, or "F12" as a Tecan
   * well). The medium control used for background subtraction is the
   * AVERAGE of every item in the list. Falls back to the plate's own two
   * medium-control wells (row F/G, plate col 12) if omitted or blank. */
  mediumControls?: string;
  /** How to interpret non-numeric items in mediumControls: "excel" reads
   * them as spreadsheet cell addresses (column letter + row number, e.g.
   * "F47"); "tecan" reads them as Tecan plate wells (row letter A-H + plate
   * column 1-12, e.g. "F12") and resolves them relative to the raw plate's
   * anchor, same as the physical well grid. Defaults to "excel". */
  mediumControlRefMode?: MediumControlRefMode;
  doseAxis?: DoseAxis;
  doseUnit?: string;
}

/** "excel": raw spreadsheet cell address (e.g. "F47"). "tecan": physical plate well (row A-H + column 1-12, e.g. "F12"). */
export type MediumControlRefMode = "excel" | "tecan";

export interface BuildInfo {
  cellLine1: string;
  cellLine2: string;
  drug1: string;
  drug2: string;
  /** Resolved medium-control items actually used, e.g. ["M47", "M48"] or ["0.223", "H12"]. */
  mediumControlRefs: string[];
  /** AVERAGE of mediumControlRefs' values - the background value subtracted in block 2/4. */
  mediumControlAverage: number;
}

export interface BuildResult {
  workbook: ExcelJS.Workbook;
  info: BuildInfo;
}

/** openpyxl.utils.get_column_letter equivalent (1-based column number -> letters). */
export function colLetter(n: number): string {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cellValue(ws: ExcelJS.Worksheet, row: number, col: number): unknown {
  const cell = ws.getCell(row, col);
  const v = cell.value;
  if (v && typeof v === "object" && !Array.isArray(v) && "result" in (v as object)) {
    return (v as unknown as { result: unknown }).result;
  }
  return v;
}

function setFormula(ws: ExcelJS.Worksheet, row: number, col: number, formula: string): void {
  ws.getCell(row, col).value = { formula } as unknown as ExcelJS.CellValue;
}

/** "F" -> 6, "AA" -> 27, ... (inverse of colLetter). */
function colNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

const NUMERIC_RE = /^-?\d+(\.\d+)?$/;
const CELL_REF_RE = /^([A-Za-z]{1,3})(\d+)$/;
const TECAN_WELL_RE = /^([A-Ha-h])(1[0-2]|[1-9])$/;

interface MediumControlItem {
  /** Human-readable form for BuildInfo, in whatever notation was used to specify it (Excel cell or Tecan well). */
  display: string;
  /** Always a real spreadsheet cell address (or a literal number) - what actually goes into the AVERAGE formula. */
  formulaRef: string;
  value: number;
}

/** Row letter (A-H) + 1-based plate column (1-12) -> the spreadsheet cell that well lives in, relative to the raw-plate anchor. */
function tecanWellCell(anchor: { row: number; col: number }, letter: string, plateCol: number): { row: number; col: number; address: string } {
  const letterIndex = letter.toUpperCase().charCodeAt(0) - 65; // A=0..H=7
  const row = anchor.row + 1 + letterIndex;
  const col = anchor.col + plateCol;
  return { row, col, address: `${colLetter(col)}${row}` };
}

/**
 * Resolve one medium-control list item: either a literal number ("0.223"),
 * an Excel cell address ("F47"), or a Tecan well ("F12", row A-H + plate
 * column 1-12) depending on `mode` - so a lab member can point at whatever
 * they're looking at (the raw spreadsheet or the plate itself) instead of
 * retyping its value.
 */
function resolveMediumControlItem(
  ws: ExcelJS.Worksheet,
  token: string,
  mode: MediumControlRefMode,
  anchor: { row: number; col: number }
): MediumControlItem {
  const trimmed = token.trim();

  if (NUMERIC_RE.test(trimmed)) {
    return { display: trimmed, formulaRef: trimmed, value: parseFloat(trimmed) };
  }

  if (mode === "tecan") {
    const m = TECAN_WELL_RE.exec(trimmed);
    if (!m) {
      throw new Error(`Medium control: "${token}" is not a number or a Tecan well like "F12" (row A-H, column 1-12)`);
    }
    const letter = m[1].toUpperCase();
    const plateCol = parseInt(m[2], 10);
    const display = `${letter}${plateCol}`;
    const { row, col, address } = tecanWellCell(anchor, letter, plateCol);
    const v = cellValue(ws, row, col);
    if (v === null || v === undefined) {
      throw new Error(`Medium control: well ${display} is empty`);
    }
    const value = Number(v);
    if (Number.isNaN(value)) {
      throw new Error(`Medium control: well ${display} does not contain a number (got ${JSON.stringify(v)})`);
    }
    return { display, formulaRef: address, value };
  }

  const m = CELL_REF_RE.exec(trimmed);
  if (m) {
    const row = parseInt(m[2], 10);
    const col = colNumber(m[1]);
    const ref = trimmed.toUpperCase();
    const v = cellValue(ws, row, col);
    if (v === null || v === undefined) {
      throw new Error(`Medium control: cell ${ref} is empty`);
    }
    const value = Number(v);
    if (Number.isNaN(value)) {
      throw new Error(`Medium control: cell ${ref} does not contain a number (got ${JSON.stringify(v)})`);
    }
    return { display: ref, formulaRef: ref, value };
  }

  throw new Error(`Medium control: "${token}" is not a number or a cell reference like "F47"`);
}

/**
 * Resolve the full medium-control list (comma-separated numbers and/or
 * cell/well references). Falls back to `defaultItems()` (the plate's own
 * two medium-control wells) if the input is missing/blank.
 */
function resolveMediumControls(
  ws: ExcelJS.Worksheet,
  input: string | undefined,
  mode: MediumControlRefMode,
  anchor: { row: number; col: number },
  defaultItems: () => MediumControlItem[]
): { items: MediumControlItem[]; average: number } {
  const trimmed = input?.trim();
  let items: MediumControlItem[];
  if (!trimmed) {
    items = defaultItems();
  } else {
    const tokens = trimmed.split(",").map((t) => t.trim()).filter((t) => t.length > 0);
    if (tokens.length === 0) {
      throw new Error("Medium control: provide at least one value or reference");
    }
    items = tokens.map((t) => resolveMediumControlItem(ws, t, mode, anchor));
  }
  const average = items.reduce((sum, i) => sum + i.value, 0) / items.length;
  return { items, average };
}

/** Locate the '<>' cell marking the top-left corner of the raw plate grid. */
export function findAnchor(ws: ExcelJS.Worksheet): { row: number; col: number } {
  let found: { row: number; col: number } | null = null;
  ws.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (found) return;
      if (typeof cell.value === "string" && cell.value.trim() === "<>") {
        found = { row: rowNumber, col: colNumber };
      }
    });
  });
  if (!found) {
    throw new Error("Could not find plate-grid anchor cell ('<>') in this file");
  }
  return found;
}

type RawPlateGrid = Record<string, Record<number, unknown>>;

/** 8 well rows (A-H) x 12 plate columns, keyed by row letter and 1-based plate col. */
export function readRawPlate(ws: ExcelJS.Worksheet, anchorRow: number, anchorCol: number): RawPlateGrid {
  const grid: RawPlateGrid = {};
  const letters = "ABCDEFGH";
  for (let i = 0; i < letters.length; i++) {
    const letter = letters[i];
    const r = anchorRow + 1 + i;
    const rowValues: Record<number, unknown> = {};
    for (let plateCol = 1; plateCol <= 12; plateCol++) {
      rowValues[plateCol] = cellValue(ws, r, anchorCol + plateCol);
    }
    grid[letter] = rowValues;
  }
  return grid;
}

/**
 * Sanity-check a fresh export has the dose wells this template needs before
 * formulas get written that would otherwise silently reference blanks.
 * Medium control wells (row F/G, plate col 12) are checked separately in
 * build(), since they can also be supplied by the caller.
 */
export function validateRawPlate(grid: RawPlateGrid): void {
  for (const letter of "BCDEFG") {
    for (let c = 2; c <= 11; c++) {
      if (grid[letter][c] === null || grid[letter][c] === undefined) {
        throw new Error(`Row ${letter} is missing well values in plate cols 2-11`);
      }
    }
  }
}

/**
 * col0..col0+9 = dose 0..9 ascending.
 *
 * Series mode: highest dose is a literal constant (topDose); each column to
 * its left is that constant divided by dilutionFactor one more time; dose 0
 * is a literal 0. Written as live formulas, same as the CLI.
 *
 * Manual mode: each of the N_DOSES concentrations is written as-is, low to
 * high, left to right - no formula, since there's no fixed ratio between
 * them.
 */
function writeDoseHeader(ws: ExcelJS.Worksheet, row: number, col0: number, doseAxis: DoseAxis): void {
  if (doseAxis.mode === "manual") {
    if (doseAxis.concentrations.length !== N_DOSES) {
      throw new Error(
        `Manual concentrations must have exactly ${N_DOSES} values (one per plate dilution column, low to high); got ${doseAxis.concentrations.length}`
      );
    }
    doseAxis.concentrations.forEach((v, i) => {
      ws.getCell(row, col0 + i).value = v;
    });
    return;
  }

  const topCol = col0 + N_DOSES - 1;
  ws.getCell(row, topCol).value = doseAxis.topDose;
  for (let c = topCol - 1; c > col0; c--) {
    setFormula(ws, row, c, `${colLetter(c + 1)}${row}/${doseAxis.dilutionFactor}`);
  }
  ws.getCell(row, col0).value = 0;
}

export async function build(arrayBuffer: ArrayBuffer, options: BuildOptions = {}): Promise<BuildResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  const ws = workbook.worksheets[0];
  if (!ws) {
    throw new Error("Workbook has no worksheets");
  }

  const drug1 = options.drug1 || PLACEHOLDER_DRUG1;
  const drug2 = options.drug2 || PLACEHOLDER_DRUG2;
  const cellLine1 = options.cellLine1 || PLACEHOLDER_CELL_LINE;
  const cellLine2 = options.cellLine2 || PLACEHOLDER_CELL_LINE;
  // Rows B-D and E-G can be the same cell line + different drugs, the same
  // drug + different cell lines, or both different. The title cell only has
  // room for one label, so collapse it to a single name when both blocks
  // share a cell line, and show "cellLine1 / cellLine2" (rows B-D / rows
  // E-G, same left-to-right order as drug1/drug2) when they don't.
  const cellLineLabel = cellLine1 === cellLine2 ? cellLine1 : `${cellLine1} / ${cellLine2}`;
  const doseAxis: DoseAxis = options.doseAxis ?? { mode: "series", topDose: TOP_DOSE_NM, dilutionFactor: DILUTION_FACTOR };
  const doseUnit = options.doseUnit || DOSE_UNIT;

  const { row: anchorRow, col: anchorCol } = findAnchor(ws);
  const grid = readRawPlate(ws, anchorRow, anchorCol);
  validateRawPlate(grid);

  // column offsets, relative to the anchor column (normally A -> col O etc.)
  const col_O = anchorCol + 14;
  const col_P = anchorCol + 15;
  const col_Z = anchorCol + 25;
  const col_AA = anchorCol + 26;
  const col_AB = anchorCol + 27;
  const col_AC = anchorCol + 28;
  const col_AM = anchorCol + 38;
  const col_M = anchorCol + 12; // medium control (blank) well column
  const col_N = anchorCol + 13; // averaged medium control cell

  const row_title1 = anchorRow - 1; // 40
  const row_dose1 = anchorRow + 1; // 42
  const row_data1 = anchorRow + 2; // 43: B,C,D (drug1) then E,F,G (drug2)
  const row_title2 = anchorRow + 10; // 51
  const row_dose2 = anchorRow + 12; // 53
  const row_data2 = anchorRow + 13; // 54

  const blank_row1 = anchorRow + 6; // row F, absolute (M47)
  const blank_row2 = anchorRow + 7; // row G, absolute (M48)
  const medium_avg_row = row_title2; // N51
  const ctrl_row1_raw = row_data1 + 1; // 44 (middle of B,C,D)
  const ctrl_row2_raw = row_data1 + 4; // 47 (middle of E,F,G)
  const ctrl_row1_corr = row_data2 + 1; // 55
  const ctrl_row2_corr = row_data2 + 4; // 58

  // Medium control (background) wells: average of a comma-separated list of
  // values and/or references (Excel cell addresses or Tecan wells,
  // depending on mediumControlRefMode), defaulting to the plate's own two
  // medium-control wells (row F/G, plate col 12 - M47/M48 in the default
  // anchor position) if omitted or blank. The AVERAGE formula is built
  // directly from whatever was resolved, so it stays live no matter how
  // many wells (or literal overrides) are involved.
  const mediumControlRefMode: MediumControlRefMode = options.mediumControlRefMode ?? "excel";
  const defaultMediumControlItems = (): MediumControlItem[] =>
    ([
      ["F", blank_row1],
      ["G", blank_row2],
    ] as const).map(([letter, row]) => {
      const address = `${colLetter(col_M)}${row}`;
      const v = cellValue(ws, row, col_M);
      if (v === null || v === undefined) {
        throw new Error(
          `Medium control well ${letter}12 (${address}) is empty; supply medium control values`
        );
      }
      return {
        display: mediumControlRefMode === "tecan" ? `${letter}12` : address,
        formulaRef: address,
        value: Number(v),
      };
    });
  const { items: mediumControlItems, average: mediumControlAverage } = resolveMediumControls(
    ws,
    options.mediumControls,
    mediumControlRefMode,
    { row: anchorRow, col: anchorCol },
    defaultMediumControlItems
  );
  const mediumControlRefs = mediumControlItems.map((i) => i.display);
  setFormula(
    ws,
    medium_avg_row,
    col_N,
    `AVERAGE(${mediumControlItems.map((i) => i.formulaRef).join(",")})`
  );

  // --- Block 1: raw absorbance, reshaped from the raw well grid ---
  ws.getCell(row_title1, col_P).value = "MTS Absorbance 490 nm";
  ws.getCell(row_dose1, col_O).value = cellLineLabel;
  writeDoseHeader(ws, row_dose1, col_P, doseAxis);
  ws.getCell(row_dose1, col_Z).value = doseUnit;
  for (let i = 0; i < 6; i++) {
    // rows 43..48
    const r = row_data1 + i;
    for (let j = 0; j < N_DOSES; j++) {
      const rawCol = anchorCol + (11 - j); // plate col 11-j
      setFormula(ws, r, col_P + j, `${colLetter(rawCol)}${r}`);
    }
  }
  ws.getCell(ctrl_row1_raw, col_O).value = drug1;
  setFormula(
    ws,
    ctrl_row1_raw,
    col_AA,
    `AVERAGE(${colLetter(col_P)}${row_data1}:${colLetter(col_P)}${row_data1 + 2})`
  );
  ws.getCell(ctrl_row1_raw, col_AB).value = drug1;
  ws.getCell(ctrl_row2_raw, col_O).value = drug2;
  setFormula(
    ws,
    ctrl_row2_raw,
    col_AA,
    `AVERAGE(${colLetter(col_P)}${row_data1 + 3}:${colLetter(col_P)}${row_data1 + 5})`
  );
  ws.getCell(ctrl_row2_raw, col_AB).value = drug2;

  // --- Block 2: corrected absorbance = block 1 minus the averaged medium control ---
  ws.getCell(row_title2, col_P).value = "MTS Absorbance 490 nm Corrected";
  writeDoseHeader(ws, row_dose2, col_P, doseAxis);
  ws.getCell(row_dose2, col_Z).value = doseUnit;
  for (let i = 0; i < 6; i++) {
    // rows 54..59
    const r1 = row_data1 + i;
    const r2 = row_data2 + i;
    for (let j = 0; j < N_DOSES; j++) {
      const c = col_P + j;
      setFormula(ws, r2, c, `${colLetter(c)}${r1}-$${colLetter(col_N)}$${medium_avg_row}`);
    }
  }
  ws.getCell(ctrl_row1_corr, col_O).value = drug1;
  setFormula(
    ws,
    ctrl_row1_corr,
    col_AA,
    `AVERAGE(${colLetter(col_P)}${row_data2}:${colLetter(col_P)}${row_data2 + 2})`
  );
  ws.getCell(ctrl_row1_corr, col_AB).value = drug1;
  ws.getCell(ctrl_row2_corr, col_O).value = drug2;
  setFormula(
    ws,
    ctrl_row2_corr,
    col_AA,
    `AVERAGE(${colLetter(col_P)}${row_data2 + 3}:${colLetter(col_P)}${row_data2 + 5})`
  );
  ws.getCell(ctrl_row2_corr, col_AB).value = drug2;

  // --- Block 3: % viability (DMSO adjusted) = block 1 / its own control mean ---
  ws.getCell(row_title1, col_AC).value = "% Viability Relative to Control (DMSO adjusted)";
  writeDoseHeader(ws, row_dose1, col_AC, doseAxis);
  ws.getCell(row_dose1, col_AM).value = doseUnit;
  for (let i = 0; i < 6; i++) {
    const r = row_data1 + i;
    const ctrlRow = i < 3 ? ctrl_row1_raw : ctrl_row2_raw;
    for (let j = 0; j < N_DOSES; j++) {
      const srcCol = col_P + j;
      const dstCol = col_AC + j;
      setFormula(ws, r, dstCol, `${colLetter(srcCol)}${r}/$${colLetter(col_AA)}$${ctrlRow}`);
    }
  }

  // --- Block 4: % viability (DMSO adjusted) corrected = block 2 / its own control mean ---
  ws.getCell(row_title2, col_AC).value = "% Viability Relative to Control (DMSO adjusted) Corrected";
  writeDoseHeader(ws, row_dose2, col_AC, doseAxis);
  ws.getCell(row_dose2, col_AM).value = doseUnit;
  for (let i = 0; i < 6; i++) {
    const r = row_data2 + i;
    const ctrlRow = i < 3 ? ctrl_row1_corr : ctrl_row2_corr;
    for (let j = 0; j < N_DOSES; j++) {
      const srcCol = col_P + j;
      const dstCol = col_AC + j;
      setFormula(ws, r, dstCol, `${colLetter(srcCol)}${r}/$${colLetter(col_AA)}$${ctrlRow}`);
    }
  }

  return {
    workbook,
    info: { cellLine1, cellLine2, drug1, drug2, mediumControlRefs, mediumControlAverage },
  };
}

export function outputFileName(inputName: string): string {
  const dot = inputName.lastIndexOf(".");
  const stem = dot === -1 ? inputName : inputName.slice(0, dot);
  const ext = dot === -1 ? ".xlsx" : inputName.slice(dot);
  return `${stem}_processed${ext}`;
}

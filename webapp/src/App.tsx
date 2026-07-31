import { useRef, useState } from "react";
import JSZip from "jszip";
import {
  build,
  outputFileName,
  N_DOSES,
  TOP_DOSE_NM,
  DILUTION_FACTOR,
  DOSE_UNIT,
  type BuildInfo,
  type DoseAxis,
  type MediumControlRefMode,
} from "./lib/buildViabilityData";
import "./App.css";

type Mode = "single" | "batch";
type DoseMode = "series" | "manual";

interface BatchResult {
  name: string;
  outName?: string;
  info?: BuildInfo;
  error?: string;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function FlaskIcon() {
  return (
    <svg viewBox="0 0 48 48" fill="none" aria-hidden="true">
      <path
        d="M19 6h10v10.2l8 15.6c1.7 3.3-.7 7.2-4.4 7.2H15.4c-3.7 0-6.1-3.9-4.4-7.2l8-15.6V6Z"
        stroke="#fff"
        strokeWidth="2.6"
        strokeLinejoin="round"
      />
      <path d="M16.5 27h15" stroke="#fff" strokeWidth="2.6" strokeLinecap="round" />
      <circle cx="20.5" cy="32.5" r="1.6" fill="#fff" />
      <circle cx="26" cy="35.5" r="1.2" fill="#fff" />
      <circle cx="29.5" cy="30.5" r="1" fill="#fff" />
    </svg>
  );
}

function LockIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <rect x="5" y="10.5" width="14" height="9.5" rx="2" stroke="currentColor" strokeWidth="1.8" />
      <path d="M8 10.5V7.5a4 4 0 0 1 8 0v3" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    </svg>
  );
}

function UploadIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 15V4m0 0 4 4m-4-4-4 4M5 16v2a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-2"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function FolderIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M4 6.5A1.5 1.5 0 0 1 5.5 5h4.4a1.5 1.5 0 0 1 1.2.6l1.4 1.9h6a1.5 1.5 0 0 1 1.5 1.5v9a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 18V6.5Z"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function CheckCircleIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
      <path d="M8.5 12.5 11 15l4.5-5.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function AlertIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 3.5 21.5 20h-19L12 3.5Z"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinejoin="round"
      />
      <path d="M12 10v4.2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
      <circle cx="12" cy="17" r="0.9" fill="currentColor" />
    </svg>
  );
}

function isSkippable(name: string): boolean {
  if (name.startsWith("~$")) return true;
  const dot = name.lastIndexOf(".");
  const stem = dot === -1 ? name : name.slice(0, dot);
  return stem.endsWith("_processed");
}

/** Parses the manual-concentrations textarea: comma/whitespace/newline separated numbers. */
function parseConcentrations(text: string): number[] {
  return text
    .split(/[\s,]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map(Number);
}

/** Trims float noise for display (the live Excel formula still keeps full precision). */
function formatNumber(n: number): string {
  return Number(n.toFixed(6)).toString();
}

function App() {
  const [mode, setMode] = useState<Mode>("single");

  // --- shared assay settings ---
  const [doseMode, setDoseMode] = useState<DoseMode>("series");
  const [topDose, setTopDose] = useState(String(TOP_DOSE_NM));
  const [dilutionFactor, setDilutionFactor] = useState(String(DILUTION_FACTOR));
  const [doseUnit, setDoseUnit] = useState(DOSE_UNIT);
  const [manualConcentrations, setManualConcentrations] = useState("");
  const [mediumControls, setMediumControls] = useState("");
  const [mediumControlRefMode, setMediumControlRefMode] = useState<MediumControlRefMode>("excel");

  // --- single-file mode ---
  const [singleFile, setSingleFile] = useState<File | null>(null);
  // Rows B-D and E-G can be the same cell line + different drugs (the
  // common case), the same drug + different cell lines, or both different -
  // these two toggles pick which fields are shown/shared between the two
  // row blocks.
  const [sameCellLine, setSameCellLine] = useState(true);
  const [sameDrug, setSameDrug] = useState(false);
  const [drug1, setDrug1] = useState("");
  const [drug2, setDrug2] = useState("");
  const [cellLine1, setCellLine1] = useState("");
  const [cellLine2, setCellLine2] = useState("");

  // --- batch mode ---
  const [batchFiles, setBatchFiles] = useState<File[]>([]);

  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [singleResult, setSingleResult] = useState<{
    info: BuildInfo;
    outName: string;
  } | null>(null);
  const [batchResults, setBatchResults] = useState<BatchResult[] | null>(
    null
  );

  const folderInputRef = useRef<HTMLInputElement>(null);

  function attachFolderAttrs(el: HTMLInputElement | null) {
    if (el) {
      el.setAttribute("webkitdirectory", "");
      el.setAttribute("directory", "");
    }
    folderInputRef.current = el;
  }

  /** Builds the shared DoseAxis from current state, or throws with a user-facing message. */
  function resolveDoseAxis(): DoseAxis {
    if (doseMode === "manual") {
      const concentrations = parseConcentrations(manualConcentrations);
      if (concentrations.some((n) => Number.isNaN(n))) {
        throw new Error("Manual concentrations must all be numbers.");
      }
      if (concentrations.length !== N_DOSES) {
        throw new Error(
          `Manual concentrations must have exactly ${N_DOSES} values (one per plate dilution column, low to high) — got ${concentrations.length}.`
        );
      }
      return { mode: "manual", concentrations };
    }
    const top = Number(topDose);
    const factor = Number(dilutionFactor);
    if (Number.isNaN(top) || top <= 0) {
      throw new Error("Top dose must be a positive number.");
    }
    if (Number.isNaN(factor) || factor <= 0) {
      throw new Error("Dilution factor must be a positive number.");
    }
    return { mode: "series", topDose: top, dilutionFactor: factor };
  }

  async function handleSingleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!singleFile) {
      setError("Choose a raw Tecan .xlsx export first.");
      return;
    }
    setProcessing(true);
    setError(null);
    setSingleResult(null);
    try {
      const doseAxis = resolveDoseAxis();
      const buffer = await singleFile.arrayBuffer();
      const d1 = drug1.trim() || undefined;
      const d2 = (sameDrug ? drug1 : drug2).trim() || undefined;
      const c1 = cellLine1.trim() || undefined;
      const c2 = (sameCellLine ? cellLine1 : cellLine2).trim() || undefined;
      const { workbook, info } = await build(buffer, {
        drug1: d1,
        drug2: d2,
        cellLine1: c1,
        cellLine2: c2,
        mediumControls: mediumControls.trim() || undefined,
        mediumControlRefMode,
        doseAxis,
        doseUnit: doseUnit.trim() || undefined,
      });
      const outBuffer = await workbook.xlsx.writeBuffer();
      const outName = outputFileName(singleFile.name);
      downloadBlob(new Blob([outBuffer], { type: XLSX_MIME }), outName);
      setSingleResult({ info, outName });
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setProcessing(false);
    }
  }

  async function handleBatchSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (batchFiles.length === 0) {
      setError("Choose one or more raw Tecan .xlsx exports first.");
      return;
    }
    setProcessing(true);
    setError(null);
    setBatchResults(null);

    let doseAxis: DoseAxis;
    try {
      doseAxis = resolveDoseAxis();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setProcessing(false);
      return;
    }

    const mc = mediumControls.trim() || undefined;
    const unit = doseUnit.trim() || undefined;

    const results: BatchResult[] = [];
    const zip = new JSZip();
    let okCount = 0;

    for (const file of batchFiles) {
      if (isSkippable(file.name)) continue;
      try {
        const buffer = await file.arrayBuffer();
        const { workbook, info } = await build(buffer, {
          mediumControls: mc,
          mediumControlRefMode,
          doseAxis,
          doseUnit: unit,
        });
        const outBuffer = await workbook.xlsx.writeBuffer();
        const outName = outputFileName(file.name);
        zip.file(outName, outBuffer);
        results.push({ name: file.name, outName, info });
        okCount++;
      } catch (err) {
        results.push({
          name: file.name,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    setBatchResults(results);

    if (okCount > 0) {
      const zipBlob = await zip.generateAsync({ type: "blob" });
      downloadBlob(zipBlob, "viability_data_processed.zip");
    }

    setProcessing(false);
  }

  const assaySettings = (
    <div className="card">
      <h2>Assay settings</h2>

      <div className="field">
        <span>Dose axis</span>
        <div className="mode-switch inline" role="tablist">
          <button
            type="button"
            role="tab"
            aria-selected={doseMode === "series"}
            className={doseMode === "series" ? "active" : ""}
            onClick={() => setDoseMode("series")}
          >
            Dilution series
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={doseMode === "manual"}
            className={doseMode === "manual" ? "active" : ""}
            onClick={() => setDoseMode("manual")}
          >
            Manual concentrations
          </button>
        </div>
      </div>

      {doseMode === "series" ? (
        <div className="field-row">
          <label className="field">
            <span>Top dose</span>
            <input
              type="number"
              step="any"
              min="0"
              value={topDose}
              onChange={(e) => setTopDose(e.target.value)}
            />
          </label>
          <label className="field field-narrow">
            <span>Dilution factor</span>
            <input
              type="number"
              step="any"
              min="0"
              value={dilutionFactor}
              onChange={(e) => setDilutionFactor(e.target.value)}
            />
          </label>
          <label className="field field-narrow">
            <span>Unit</span>
            <input
              type="text"
              value={doseUnit}
              onChange={(e) => setDoseUnit(e.target.value)}
            />
          </label>
        </div>
      ) : (
        <>
          <label className="field">
            <span>
              Concentrations, low&nbsp;&rarr;&nbsp;high ({N_DOSES} values,
              comma or newline separated)
            </span>
            <textarea
              rows={2}
              placeholder={`0, 4.88, 9.77, 19.5, ... (${N_DOSES} values)`}
              value={manualConcentrations}
              onChange={(e) => setManualConcentrations(e.target.value)}
            />
          </label>
          <label className="field">
            <span>Unit</span>
            <input
              type="text"
              value={doseUnit}
              onChange={(e) => setDoseUnit(e.target.value)}
            />
          </label>
        </>
      )}

      <div className="field">
        <span>Medium control wells</span>
        <div className="mode-switch inline" role="tablist">
          <button
            type="button"
            role="tab"
            aria-selected={mediumControlRefMode === "excel"}
            className={mediumControlRefMode === "excel" ? "active" : ""}
            onClick={() => setMediumControlRefMode("excel")}
          >
            Excel cells
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mediumControlRefMode === "tecan"}
            className={mediumControlRefMode === "tecan" ? "active" : ""}
            onClick={() => setMediumControlRefMode("tecan")}
          >
            Tecan wells
          </button>
        </div>
      </div>
      <label className="field">
        <input
          type="text"
          placeholder={
            mediumControlRefMode === "excel"
              ? "defaults to the plate's own M47, M48"
              : "defaults to the plate's own F12, G12"
          }
          value={mediumControls}
          onChange={(e) => setMediumControls(e.target.value)}
        />
      </label>
      {mediumControlRefMode === "excel" ? (
        <p className="hint">
          Comma-separated list of values and/or Excel cell references (e.g.{" "}
          <code>F47, F48</code>) &mdash; the medium control used is the average of all of them. 
        </p>
      ) : (
        <p className="hint">
          Comma-separated list of values and/or Tecan wells, row A&ndash;H +
          column 1&ndash;12 (e.g. <code>F12, G12</code>) &mdash; the medium control used is the average of all of them. 
        </p>
      )}
    </div>
  );

  const filesCard = (
    <form
      className="card"
      onSubmit={(e) => (mode === "single" ? handleSingleSubmit(e) : handleBatchSubmit(e))}
    >
      <div className="card-header-row">
        <h2>Files</h2>
        <div className="mode-switch inline" role="tablist">
          <button
            type="button"
            role="tab"
            aria-selected={mode === "single"}
            className={mode === "single" ? "active" : ""}
            onClick={() => {
              setMode("single");
              setError(null);
              setBatchResults(null);
            }}
          >
            Single file
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mode === "batch"}
            className={mode === "batch" ? "active" : ""}
            onClick={() => {
              setMode("batch");
              setError(null);
              setSingleResult(null);
            }}
          >
            Batch (multiple files)
          </button>
        </div>
      </div>

      {mode === "single" ? (
        <>
          <label className="field">
            <span>Raw export (.xlsx)</span>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setSingleFile(e.target.files?.[0] ?? null)}
            />
          </label>

          <div className="field">
            <span>What differs between rows B&ndash;D and E&ndash;G?</span>
            <div className="checkbox-row">
              <label className="checkbox">
                <input
                  type="checkbox"
                  checked={!sameDrug}
                  onChange={(e) => setSameDrug(!e.target.checked)}
                />
                Different drugs
              </label>
              <label className="checkbox">
                <input
                  type="checkbox"
                  checked={!sameCellLine}
                  onChange={(e) => setSameCellLine(!e.target.checked)}
                />
                Different cell lines
              </label>
            </div>
          </div>

          <div className="field-grid">
            {sameDrug ? (
              <label className="field">
                <span>Drug</span>
                <input
                  type="text"
                  placeholder="DRUG1"
                  value={drug1}
                  onChange={(e) => setDrug1(e.target.value)}
                />
              </label>
            ) : (
              <>
                <label className="field">
                  <span>Drug 1 (rows B&ndash;D)</span>
                  <input
                    type="text"
                    placeholder="DRUG1"
                    value={drug1}
                    onChange={(e) => setDrug1(e.target.value)}
                  />
                </label>
                <label className="field">
                  <span>Drug 2 (rows E&ndash;G)</span>
                  <input
                    type="text"
                    placeholder="DRUG2"
                    value={drug2}
                    onChange={(e) => setDrug2(e.target.value)}
                  />
                </label>
              </>
            )}

            {sameCellLine ? (
              <label className="field">
                <span>Cell line</span>
                <input
                  type="text"
                  placeholder="CELL_LINE"
                  value={cellLine1}
                  onChange={(e) => setCellLine1(e.target.value)}
                />
              </label>
            ) : (
              <>
                <label className="field">
                  <span>Cell line (rows B&ndash;D)</span>
                  <input
                    type="text"
                    placeholder="CELL_LINE"
                    value={cellLine1}
                    onChange={(e) => setCellLine1(e.target.value)}
                  />
                </label>
                <label className="field">
                  <span>Cell line (rows E&ndash;G)</span>
                  <input
                    type="text"
                    placeholder="CELL_LINE"
                    value={cellLine2}
                    onChange={(e) => setCellLine2(e.target.value)}
                  />
                </label>
              </>
            )}
          </div>
        </>
      ) : (
        <>
          <div className="field">
            <span>Raw exports (.xlsx)</span>
            <div className="file-choice-grid">
              <label className="file-choice">
                <span className="file-choice-label">
                  <UploadIcon />
                  Select multiple files
                </span>
                <input
                  type="file"
                  accept=".xlsx"
                  multiple
                  onChange={(e) =>
                    setBatchFiles(Array.from(e.target.files ?? []))
                  }
                />
              </label>
              <label className="file-choice">
                <span className="file-choice-label">
                  <FolderIcon />
                  Select a folder
                </span>
                <input type="file" multiple ref={attachFolderAttrs} onChange={(e) =>
                    setBatchFiles(Array.from(e.target.files ?? []))
                  }
                />
              </label>
            </div>
            {batchFiles.length > 0 && (
              <p className="hint">{batchFiles.length} file(s) selected</p>
            )}
          </div>

          <p className="hint">
            Files already ending in <code>_processed.xlsx</code> are skipped.
          </p>
        </>
      )}

      <button type="submit" className="submit" disabled={processing}>
        {processing && <span className="spinner" aria-hidden="true" />}
        {processing
          ? "Processing…"
          : mode === "single"
          ? "Build & download"
          : "Build & download all (.zip)"}
      </button>
    </form>
  );

  return (
    <div className="page">
      <header className="page-header">
        <span className="brand-mark">
          <FlaskIcon />
        </span>
        <h1>MTS Viability Data Builder</h1>
        <p className="subtitle">
          Fill in the raw-absorbance and %-viability calculation blocks for
          this lab's MTS-assay Result sheets from a fresh Tecan Spark export
          &mdash; no Python or Excel formulas to hand-type.
        </p>
      </header>

      <div className="settings-grid">
        {assaySettings}
        {filesCard}
      </div>

      {error && (
        <div className="alert error">
          <span className="alert-icon">
            <AlertIcon />
          </span>
          <div className="alert-body">{error}</div>
        </div>
      )}

      {singleResult && (
        <div className="alert success">
          <span className="alert-icon">
            <CheckCircleIcon />
          </span>
          <div className="alert-body">
            <strong>Wrote {singleResult.outName}</strong>
            <dl>
              <dt>Drug 1 (rows B&ndash;D)</dt>
              <dd>{singleResult.info.drug1}</dd>
              <dt>Drug 2 (rows E&ndash;G)</dt>
              <dd>{singleResult.info.drug2}</dd>
              <dt>Cell line (rows B&ndash;D)</dt>
              <dd>{singleResult.info.cellLine1}</dd>
              <dt>Cell line (rows E&ndash;G)</dt>
              <dd>{singleResult.info.cellLine2}</dd>
              <dt>Medium control wells</dt>
              <dd>{singleResult.info.mediumControlRefs.join(", ")}</dd>
              <dt>Medium control average</dt>
              <dd>{formatNumber(singleResult.info.mediumControlAverage)}</dd>
            </dl>
          </div>
        </div>
      )}

      {batchResults && (
        <div className="card">
          <h2>
            {batchResults.filter((r) => !r.error).length}/{batchResults.length}{" "}
            processed
          </h2>
          <div className="results-table-wrap">
          <table className="results-table">
            <thead>
              <tr>
                <th>File</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {batchResults.map((r) => (
                <tr key={r.name}>
                  <td>{r.name}</td>
                  <td>
                    {r.error ? (
                      <span className="status-error">FAILED: {r.error}</span>
                    ) : (
                      <span className="status-ok">&rarr; {r.outName}</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
        </div>
      )}

      <footer className="page-footer">
        <p>
          <LockIcon />
          All processing happens locally in your browser; no file is
          uploaded anywhere.
        </p>
      </footer>
    </div>
  );
}

export default App;

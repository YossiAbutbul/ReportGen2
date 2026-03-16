import { useMemo, useState } from "react";

import {
  TEMPLATE_SCOPE,
  TEMPLATE_TESTED_POWER,
  TEMPLATE_TITLE,
} from "./constants";
import { parseWorkbook } from "./services/excelParser";
import { buildDocx } from "./services/wordReportBuilder";
import type { SummaryData } from "./types";
import { formatFrequency, formatNumber, todayAsDDMMYYYY } from "./utils/format";
import "./reportGenerator.css";

function downloadBlob(blob: Blob, filename: string) {
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = objectUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}

export default function ReportGeneratorPage() {
  const [excelFileName, setExcelFileName] = useState("");
  const [author, setAuthor] = useState("Yossi Abutbul");
  const [reportTitle, setReportTitle] = useState(TEMPLATE_TITLE);
  const [reportDate, setReportDate] = useState(todayAsDDMMYYYY());
  const [scopeOfTesting, setScopeOfTesting] = useState(TEMPLATE_SCOPE);
  const [fwVersion, setFwVersion] = useState("");
  const [hwVersion, setHwVersion] = useState("");
  const [testedPower, setTestedPower] = useState(TEMPLATE_TESTED_POWER);
  const [parsed, setParsed] = useState<SummaryData | null>(null);
  const [error, setError] = useState("");
  const [isBuilding, setIsBuilding] = useState(false);

  const totals = useMemo(() => {
    if (!parsed) return { rows: 0, units: 0, frequencies: 0 };

    return {
      rows: parsed.rows.length,
      units: parsed.uniqueUnitIds.length,
      frequencies: parsed.uniqueFrequencies.length,
    };
  }, [parsed]);

  async function handleUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    setError("");
    setParsed(null);
    setExcelFileName(file.name);

    try {
      const result = await parseWorkbook(file);

      if (!result.rows.length) {
        throw new Error(
          "No valid data rows were found. Make sure the Excel contains headers like Unit ID, Frequency, TRP, and Max Peak."
        );
      }

      setParsed(result);
    } catch (err) {
      setParsed(null);
      setError(err instanceof Error ? err.message : "Failed to parse the file.");
    }
  }

  async function handleGenerate() {
    if (!parsed) {
      setError("Please upload a valid Excel file first.");
      return;
    }

    setIsBuilding(true);
    setError("");

    try {
      const dateForFileName = (reportDate.trim() || todayAsDDMMYYYY()).replace(
        /[^0-9.\\-]+/g,
        "_"
      );

      const blob = await buildDocx({
        title: reportTitle.trim() || TEMPLATE_TITLE,
        author: author.trim() || "Author",
        dateText: reportDate.trim() || todayAsDDMMYYYY(),
        scopeOfTesting: scopeOfTesting.trim() || TEMPLATE_SCOPE,
        fwVersion: fwVersion.trim(),
        hwVersion: hwVersion.trim(),
        testedPower: testedPower.trim() || "-",
        frequencies: parsed.uniqueFrequencies,
        unitIds: parsed.uniqueUnitIds,
        rows: parsed.rows,
      });

      const safeName = (reportTitle.trim() || "test-summary").replace(
        /[^a-z0-9-_]+/gi,
        "_"
      );

      downloadBlob(blob, `${safeName}-${dateForFileName}.docx`);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Failed to generate the Word file."
      );
    } finally {
      setIsBuilding(false);
    }
  }

  function clearAll() {
    setExcelFileName("");
    setParsed(null);
    setError("");
    setReportDate(todayAsDDMMYYYY());
    setFwVersion("");
    setHwVersion("");
    setScopeOfTesting(TEMPLATE_SCOPE);
  }

  return (
    <div className="page">
      <div className="layout">
        <div className="panel">
          <h1>Test Report Generator</h1>
          <p className="muted">
            Upload the Excel, insert FW and HW versions, and export a Word summary.
          </p>

          <div className="section">
            <label className="uploadBox">
              <div className="uploadTitle">Upload Test Result Excel file</div>
              {/* <div className="uploadSub">
                Unnamed columns are ignored, and the Cell Factor table is skipped.
              </div> */}
              <input type="file" accept=".xlsx,.xls" onChange={handleUpload} />
            </label>

            {excelFileName && (
              <div className="fileBadge">
                <strong>Loaded:</strong> {excelFileName}
              </div>
            )}
          </div>

          <div className="formGrid">
            <div>
              <label>Report title</label>
              <input
                value={reportTitle}
                onChange={(event) => setReportTitle(event.target.value)}
              />
            </div>

            <div>
              <label>Author</label>
              <input
                value={author}
                onChange={(event) => setAuthor(event.target.value)}
              />
            </div>

            <div>
              <label>Date</label>
              <input
                value={reportDate}
                onChange={(event) => setReportDate(event.target.value)}
              />
            </div>

            <div>
              <label>Tested power</label>
              <input
                value={testedPower}
                onChange={(event) => setTestedPower(event.target.value)}
              />
            </div>

            <div>
              <label>F.W. Version</label>
              <input
                placeholder="e.g. 2E.51"
                value={fwVersion}
                onChange={(event) => setFwVersion(event.target.value)}
              />
            </div>

            <div>
              <label>H.W. Version</label>
              <input
                placeholder="e.g. 08.06"
                value={hwVersion}
                onChange={(event) => setHwVersion(event.target.value)}
              />
            </div>
          </div>

          <div className="section">
            <label>Scope of Testing</label>
            <textarea
              rows={4}
              value={scopeOfTesting}
              onChange={(event) => setScopeOfTesting(event.target.value)}
              placeholder="Describe the test scope for the report"
            />
          </div>

          <div className="actions">
            <button onClick={handleGenerate} disabled={!parsed || isBuilding}>
              {isBuilding ? "Generating..." : "Generate Word File"}
            </button>
            <button className="secondary" onClick={clearAll}>
              Clear
            </button>
          </div>

          {error && <div className="errorBox">{error}</div>}
        </div>

        <div className="panel">
          <h2>Parsed Summary</h2>

          <div className="stats">
            <div className="statCard">
              <div className="muted">Units</div>
              <div className="statValue">{totals.units}</div>
            </div>
            <div className="statCard">
              <div className="muted">Frequencies</div>
              <div className="statValue">{totals.frequencies}</div>
            </div>
            <div className="statCard">
              <div className="muted">Result Rows</div>
              <div className="statValue">{totals.rows}</div>
            </div>
          </div>

          <div className="section">
            <h3>Unit IDs</h3>
            <div className="box">
              {parsed?.uniqueUnitIds.length ? (
                <ul>
                  {parsed.uniqueUnitIds.map((id) => (
                    <li key={id}>{id}</li>
                  ))}
                </ul>
              ) : (
                <div className="muted">Upload a file to see unit IDs.</div>
              )}
            </div>
          </div>

          <div className="section">
            <h3>Detected Frequencies</h3>
            <div className="box">
              {parsed?.uniqueFrequencies.length
                ? parsed.uniqueFrequencies
                    .map((value) => formatFrequency(value))
                    .join(", ")
                : "Upload a file to detect frequencies."}
            </div>
          </div>

          <div className="section">
            <h3>Preview of radiated results</h3>
            <div className="tableWrap">
              <table>
                <thead>
                  <tr>
                    <th>Unit ID</th>
                    <th>Frequency</th>
                    <th>TRP</th>
                    <th>Max Peak</th>
                  </tr>
                </thead>
                <tbody>
                  {parsed?.rows.length ? (
                    parsed.rows.map((row, index) => (
                      <tr key={`${row.unitId}-${row.frequencyMHz ?? "na"}-${index}`}>
                        <td>{row.unitId}</td>
                        <td>{formatFrequency(row.frequencyMHz)}</td>
                        <td>{formatNumber(row.trp)}</td>
                        <td>{formatNumber(row.maxPeak)}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={4} className="emptyCell">
                        No data loaded yet.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

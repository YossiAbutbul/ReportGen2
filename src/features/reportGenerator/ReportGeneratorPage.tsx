import { useMemo, useRef, useState } from "react";

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
  const fileInputRef = useRef<HTMLInputElement | null>(null);
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
  const [selectedPhoto, setSelectedPhoto] = useState<string | null>(null);

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

  function openFilePicker() {
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }

    fileInputRef.current?.click();
  }

  function normalizePhotoSource(value?: string) {
    if (!value) return "";

    const trimmed = value.trim();
    if (!trimmed) return "";

    if (/^[a-z]:\\/i.test(trimmed)) {
      return `file:///${trimmed.replace(/\\/g, "/")}`;
    }

    if (/^\\\\/.test(trimmed)) {
      return `file:${trimmed.replace(/\\/g, "/")}`;
    }

    return trimmed;
  }

  function hasPhoto(value?: string) {
    return Boolean(value?.trim());
  }

  return (
    <div className="page">
      <div className="shell">
        <section className="heroCard">
          <div className="heroCopy">
            <h1>Test Report Generator</h1>
            <p className="heroText">
              Upload the Excel results, complete the report details, and export a
              polished Word report without touching the raw data manually.
            </p>
          </div>

          <div className="heroStats">
            <div className="statCard">
              <div className="muted">Units</div>
              <div className="statValue">{totals.units}</div>
            </div>
            <div className="statCard">
              <div className="muted">Frequencies</div>
              <div className="statValue">{totals.frequencies}</div>
            </div>
            <div className="statCard accent">
              <div className="muted">Result Rows</div>
              <div className="statValue">{totals.rows}</div>
            </div>
          </div>
        </section>

        <div className="layout">
          <div className="panel setupPanel">
            <div className="setupPanelBody">
              <div className="panelHeader">
                <div>
                  <p className="panelEyebrow">Setup</p>
                  <h2>Report Details</h2>
                </div>
              </div>

              <div className="section">
                <div className="uploadRow">
                  <div className="uploadField">
                    {excelFileName || "Select .xlsx or .xls workbook"}
                  </div>
                  <button
                    type="button"
                    className="uploadButton"
                    onClick={openFilePicker}
                  >
                    Choose file
                  </button>
                  <input
                    ref={fileInputRef}
                    className="uploadInput"
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleUpload}
                  />
                </div>
              </div>

              <div className="formGrid">
                <div>
                  <label htmlFor="report-title">Report title</label>
                  <input
                    id="report-title"
                    value={reportTitle}
                    onChange={(event) => setReportTitle(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="report-author">Author</label>
                  <input
                    id="report-author"
                    value={author}
                    onChange={(event) => setAuthor(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="report-date">Date</label>
                  <input
                    id="report-date"
                    value={reportDate}
                    onChange={(event) => setReportDate(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="tested-power">Tested power</label>
                  <input
                    id="tested-power"
                    value={testedPower}
                    onChange={(event) => setTestedPower(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="fw-version">F.W. Version</label>
                  <input
                    id="fw-version"
                    placeholder="e.g. 2E.51"
                    value={fwVersion}
                    onChange={(event) => setFwVersion(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="hw-version">H.W. Version</label>
                  <input
                    id="hw-version"
                    placeholder="e.g. 08.06"
                    value={hwVersion}
                    onChange={(event) => setHwVersion(event.target.value)}
                  />
                </div>
              </div>

              <div className="section scopeSection">
                <label htmlFor="scope-of-testing">Scope of Testing</label>
                <textarea
                  id="scope-of-testing"
                  className="scopeTextarea"
                  rows={4}
                  value={scopeOfTesting}
                  onChange={(event) => setScopeOfTesting(event.target.value)}
                  placeholder="Describe the test scope for the report"
                />
              </div>

              {error && <div className="errorBox">{error}</div>}
            </div>

            <div className="actions stickyActions">
              <button type="button" className="secondary" onClick={clearAll}>
                Clear
              </button>
              <button
                type="button"
                className="primaryAction"
                onClick={handleGenerate}
                disabled={!parsed || isBuilding}
              >
                {isBuilding ? "Generating..." : "Generate Report"}
              </button>
            </div>
          </div>

          <div className="panel overviewPanel">
            <div className="panelHeader">
              <div>
                <p className="panelEyebrow">Overview</p>
                <h2>Parsed Summary</h2>
              </div>
            </div>

            <div className="infoGrid">
              <div className="summaryCard compactSection">
                <div className="summaryCardHeader">
                  <h3>Detected Frequencies</h3>
                </div>
                <div className="listArea">
                  {parsed?.uniqueFrequencies.length ? (
                    <div className="badgeGrid">
                      {parsed.uniqueFrequencies.map((value) => (
                        <span
                          key={value}
                          className="dataBadge dataBadgeAccent"
                        >
                          {formatFrequency(value)}
                        </span>
                      ))}
                    </div>
                  ) : (
                    <div className="muted">Upload a file to detect frequencies.</div>
                  )}
                </div>
              </div>

              <div className="summaryCard compactSection unitIdsCard">
                <div className="summaryCardHeader">
                  <h3>Unit IDs</h3>
                </div>
                <div className="listArea listAreaScroll unitIdsListArea">
                  {parsed?.uniqueUnitIds.length ? (
                    <div className="badgeGrid badgeGridAuto">
                      {parsed.uniqueUnitIds.map((id) => (
                        <span key={id} className="dataBadge">
                          {id}
                        </span>
                      ))}
                    </div>
                  ) : (
                    <div className="muted">Upload a file to see unit IDs.</div>
                  )}
                </div>
              </div>
            </div>

            <div className="section tableSection">
              <div className="tableHeader">
                <div>
                  <h3>Preview of radiated results</h3>
                </div>
              </div>
              <div className="tableWrap">
                <table className="resultsTable resultsTableHead">
                  <colgroup>
                    <col className="colUnitId" />
                    <col className="colFrequency" />
                    <col className="colTrp" />
                    <col className="colPeak" />
                    <col className="colPhoto" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Unit ID</th>
                      <th>Frequency</th>
                      <th>TRP</th>
                      <th>Max Peak</th>
                      <th>Photo</th>
                    </tr>
                  </thead>
                </table>
                <div className="tableBody">
                  <table className="resultsTable">
                    <colgroup>
                      <col className="colUnitId" />
                      <col className="colFrequency" />
                      <col className="colTrp" />
                      <col className="colPeak" />
                      <col className="colPhoto" />
                    </colgroup>
                    <tbody>
                      {parsed?.rows.length ? (
                        parsed.rows.map((row, index) => (
                          <tr key={`${row.unitId}-${row.frequencyMHz ?? "na"}-${index}`}>
                            <td>{row.unitId}</td>
                            <td>{formatFrequency(row.frequencyMHz)}</td>
                            <td>{formatNumber(row.trp)}</td>
                            <td>{formatNumber(row.maxPeak)}</td>
                            <td className="photoCell">
                              {hasPhoto(row.photoValue) ? (
                                <button
                                  type="button"
                                  className="photoIconButton"
                                  onClick={() =>
                                    setSelectedPhoto(
                                      normalizePhotoSource(row.photoValue) || null
                                    )
                                  }
                                  aria-label={`View 3D photo for ${row.unitId}`}
                                  title="View 3D photo"
                                >
                                  <svg
                                    viewBox="0 0 24 24"
                                    aria-hidden="true"
                                    className="photoIcon"
                                  >
                                    <path
                                      d="M8 7.5 9.7 5h4.6L16 7.5H19a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2zm4 9a4 4 0 1 0 0-8 4 4 0 0 0 0 8m0-1.8a2.2 2.2 0 1 1 0-4.4 2.2 2.2 0 0 1 0 4.4"
                                      fill="currentColor"
                                    />
                                  </svg>
                                </button>
                              ) : (
                                <span className="photoPlaceholder">-</span>
                              )}
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={5} className="emptyCell">
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
      </div>

      {selectedPhoto && (
        <div
          className="photoModalBackdrop"
          onClick={() => setSelectedPhoto(null)}
          role="presentation"
        >
          <div
            className="photoModal"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
            aria-label="3D photo preview"
          >
            <button
              type="button"
              className="photoModalClose"
              onClick={() => setSelectedPhoto(null)}
              aria-label="Close photo preview"
            >
              Close
            </button>
            <img
              className="photoModalImage"
              src={selectedPhoto}
              alt="3D preview"
            />
          </div>
        </div>
      )}
    </div>
  );
}

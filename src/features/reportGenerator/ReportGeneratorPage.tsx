import { useMemo, useRef, useState, useEffect } from "react";
import { EllipsisVertical, Download, Filter } from "lucide-react";

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

type FilterState = {
  unitTypes: string[];
  unitIds: string[];
  frequencies: string[];
};

type FilterSearchState = Record<keyof FilterState, string>;

const EMPTY_FILTERS: FilterState = {
  unitTypes: [],
  unitIds: [],
  frequencies: [],
};

const EMPTY_FILTER_SEARCH: FilterSearchState = {
  unitTypes: "",
  unitIds: "",
  frequencies: "",
};

function formatFilterSummary(
  key: keyof FilterState,
  selectedValues: string[]
) {
  if (!selectedValues.length) {
    return "Select values";
  }

  const labels = selectedValues.map((value) =>
    key === "frequencies" ? formatFrequency(Number(value)) : value
  );

  const firstLabel = labels[0] ?? "";
  if (labels.length === 1) {
    return firstLabel;
  }

  return `${firstLabel} +${labels.length - 1}`;
}

function filterRowsBySelections(
  rows: SummaryData["rows"],
  filters: FilterState,
  ignoredKey?: keyof FilterState
) {
  const selectedUnitTypes =
    ignoredKey === "unitTypes" ? new Set<string>() : new Set(filters.unitTypes);
  const selectedUnitIds =
    ignoredKey === "unitIds" ? new Set<string>() : new Set(filters.unitIds);
  const selectedFrequencies =
    ignoredKey === "frequencies" ? new Set<string>() : new Set(filters.frequencies);

  return rows.filter((row) => {
    if (selectedUnitTypes.size && !selectedUnitTypes.has(row.unitType)) {
      return false;
    }

    if (selectedUnitIds.size && !selectedUnitIds.has(row.unitId)) {
      return false;
    }

    if (
      selectedFrequencies.size &&
      !selectedFrequencies.has(String(row.frequencyMHz ?? ""))
    ) {
      return false;
    }

    return true;
  });
}

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
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS);
  const [error, setError] = useState("");
  const [isBuilding, setIsBuilding] = useState(false);
  const [selectedPhoto, setSelectedPhoto] = useState<string | null>(null);
  const [isUploadMenuOpen, setIsUploadMenuOpen] = useState(false);
  const [isFilterMenuOpen, setIsFilterMenuOpen] = useState(false);
  const [openFilterField, setOpenFilterField] = useState<keyof FilterState | null>(
    null
  );
  const [filterSearch, setFilterSearch] =
    useState<FilterSearchState>(EMPTY_FILTER_SEARCH);
  const uploadMenuRef = useRef<HTMLDivElement | null>(null);
  const filterMenuRef = useRef<HTMLDivElement | null>(null);

  const filterOptions = useMemo(() => {
    if (!parsed) {
      return {
        unitTypes: [] as string[],
        unitIds: [] as string[],
        frequencies: [] as string[],
      };
    }

    const rowsForUnitTypes = filterRowsBySelections(parsed.rows, filters, "unitTypes");
    const rowsForUnitIds = filterRowsBySelections(parsed.rows, filters, "unitIds");
    const rowsForFrequencies = filterRowsBySelections(
      parsed.rows,
      filters,
      "frequencies"
    );

    return {
      unitTypes: Array.from(
        new Set(rowsForUnitTypes.map((row) => row.unitType).filter(Boolean))
      ),
      unitIds: Array.from(
        new Set(rowsForUnitIds.map((row) => row.unitId).filter(Boolean))
      ),
      frequencies: Array.from(
        new Set(
          rowsForFrequencies
            .map((row) => row.frequencyMHz)
            .filter((value): value is number => value != null)
            .map((value) => String(value))
        )
      ).sort((left, right) => Number(left) - Number(right)),
    };
  }, [filters, parsed]);

  const filteredRows = useMemo(() => {
    if (!parsed) return [];

    return filterRowsBySelections(parsed.rows, filters);
  }, [filters, parsed]);

  const filteredUnitIds = useMemo(
    () => Array.from(new Set(filteredRows.map((row) => row.unitId).filter(Boolean))),
    [filteredRows]
  );

  const filteredFrequencies = useMemo(
    () =>
      Array.from(
        new Set(
          filteredRows
            .map((row) => row.frequencyMHz)
            .filter((value): value is number => value != null)
        )
      ).sort((a, b) => a - b),
    [filteredRows]
  );

  const totals = useMemo(() => {
    if (!parsed) return { rows: 0, units: 0, frequencies: 0 };

    return {
      rows: filteredRows.length,
      units: filteredUnitIds.length,
      frequencies: filteredFrequencies.length,
    };
  }, [filteredFrequencies.length, filteredRows.length, filteredUnitIds.length, parsed]);

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

      setFilters(EMPTY_FILTERS);
      setFilterSearch(EMPTY_FILTER_SEARCH);
      setOpenFilterField(null);
      setParsed(result);
    } catch (err) {
      setParsed(null);
      setFilters(EMPTY_FILTERS);
      setFilterSearch(EMPTY_FILTER_SEARCH);
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
    setFilters(EMPTY_FILTERS);
    setFilterSearch(EMPTY_FILTER_SEARCH);
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

  function downloadTemplate() {
    const anchor = document.createElement("a");
    anchor.href = "/report-generator-template.xlsx";
    anchor.download = "report-generator-template.xlsx";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setIsUploadMenuOpen(false);
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

  function toggleFilterValue(key: keyof FilterState, value: string) {
    setFilters((current) => {
      const nextValues = current[key].includes(value)
        ? current[key].filter((item) => item !== value)
        : [...current[key], value];

      return { ...current, [key]: nextValues };
    });
  }

  function clearFilters() {
    setFilters(EMPTY_FILTERS);
    setFilterSearch(EMPTY_FILTER_SEARCH);
    setOpenFilterField(null);
  }

  const activeFilterCount = useMemo(
    () => Object.values(filters).reduce((total, values) => total + values.length, 0),
    [filters]
  );

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      const target = event.target as Node;

      if (uploadMenuRef.current && !uploadMenuRef.current.contains(target)) {
        setIsUploadMenuOpen(false);
      }

      if (filterMenuRef.current && !filterMenuRef.current.contains(target)) {
        setIsFilterMenuOpen(false);
        setOpenFilterField(null);
      }
    }

    if (isUploadMenuOpen || isFilterMenuOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isFilterMenuOpen, isUploadMenuOpen]);

  useEffect(() => {
    setFilters((current) => {
      const nextFilters: FilterState = {
        unitTypes: current.unitTypes.filter((value) =>
          filterOptions.unitTypes.includes(value)
        ),
        unitIds: current.unitIds.filter((value) =>
          filterOptions.unitIds.includes(value)
        ),
        frequencies: current.frequencies.filter((value) =>
          filterOptions.frequencies.includes(value)
        ),
      };

      const changed =
        nextFilters.unitTypes.length !== current.unitTypes.length ||
        nextFilters.unitIds.length !== current.unitIds.length ||
        nextFilters.frequencies.length !== current.frequencies.length;

      return changed ? nextFilters : current;
    });
  }, [filterOptions]);

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
                  <div className="uploadActions" ref={uploadMenuRef}>
                    <button
                      type="button"
                      className="uploadButton"
                      onClick={openFilePicker}
                    >
                      Choose file
                    </button>

                    <button
                      type="button"
                      className="uploadMenuButton"
                      aria-label="Open upload actions"
                      aria-haspopup="menu"
                      aria-expanded={isUploadMenuOpen}
                      onClick={() => setIsUploadMenuOpen((current) => !current)}
                    >
                      <EllipsisVertical size={16} />
                    </button>

                    {isUploadMenuOpen && (
                      <div className="uploadMenu" role="menu">
                        <button
                          type="button"
                          className="uploadMenuItem"
                          role="menuitem"
                          onClick={downloadTemplate}
                        >
                          <Download size={16} />
                          <span>Download template</span>
                        </button>
                      </div>
                    )}
                  </div>
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
                    autoComplete="off"
                    value={reportTitle}
                    onChange={(event) => setReportTitle(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="report-author">Author</label>
                  <input
                    id="report-author"
                    autoComplete="off"
                    value={author}
                    onChange={(event) => setAuthor(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="report-date">Date</label>
                  <input
                    id="report-date"
                    autoComplete="off"
                    value={reportDate}
                    onChange={(event) => setReportDate(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="tested-power">Tested power</label>
                  <input
                    id="tested-power"
                    autoComplete="off"
                    value={testedPower}
                    onChange={(event) => setTestedPower(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="fw-version">F.W. Version</label>
                  <input
                    id="fw-version"
                    autoComplete="off"
                    placeholder="e.g. 2E.51"
                    value={fwVersion}
                    onChange={(event) => setFwVersion(event.target.value)}
                  />
                </div>

                <div>
                  <label htmlFor="hw-version">H.W. Version</label>
                  <input
                    id="hw-version"
                    autoComplete="off"
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
                  autoComplete="off"
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
                  {filteredFrequencies.length ? (
                    <div className="badgeGrid">
                      {filteredFrequencies.map((value) => (
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
                  {filteredUnitIds.length ? (
                    <div className="badgeGrid badgeGridAuto">
                      {filteredUnitIds.map((id) => (
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
                <div className="filterMenuAnchor" ref={filterMenuRef}>
                  <button
                    type="button"
                    className={`filterToggleButton${activeFilterCount ? " isActive" : ""}`}
                    onClick={() => setIsFilterMenuOpen((current) => !current)}
                    aria-haspopup="dialog"
                    aria-expanded={isFilterMenuOpen}
                    disabled={!parsed || !parsed.rows.length}
                  >
                    <Filter size={16} />
                    <span>Filter</span>
                    {activeFilterCount ? (
                      <span className="filterCountBadge">{activeFilterCount}</span>
                    ) : null}
                  </button>

                  {isFilterMenuOpen ? (
                    <div
                      className="filterMenuPanel"
                      role="dialog"
                      aria-label="Filter rows"
                      onMouseDown={(event) => {
                        const target = event.target as HTMLElement;
                        if (!target.closest(".filterField")) {
                          setOpenFilterField(null);
                        }
                      }}
                    >
                      <div className="filterMenuGrid">
                        {(["unitTypes", "unitIds", "frequencies"] as const).map((key) => {
                          const labelMap = {
                            unitTypes: "Unit type",
                            unitIds: "Unit ID",
                            frequencies: "Frequency",
                          };
                          const values = filterOptions[key];
                          const selectedValues = filters[key];
                          const isOpen = openFilterField === key;
                          const searchValue = filterSearch[key];
                          const normalizedSearch = searchValue.trim().toLowerCase();
                          const visibleValues = values.filter((value) => {
                            const label =
                              key === "frequencies"
                                ? formatFrequency(Number(value))
                                : value;

                            return (
                              !normalizedSearch ||
                              label.toLowerCase().includes(normalizedSearch)
                            );
                          });

                          return (
                            <div key={key} className="filterField">
                              <label>{labelMap[key]}</label>
                              <button
                                type="button"
                                className="filterSelectButton"
                                onClick={() =>
                                  setOpenFilterField((current) =>
                                    current === key ? null : key
                                  )
                                }
                              >
                                <span>
                                  {formatFilterSummary(key, selectedValues)}
                                </span>
                              </button>

                              {isOpen ? (
                                <div className="filterSelectMenu">
                                  <input
                                    className="filterSearchInput"
                                    autoComplete="off"
                                    value={searchValue}
                                    onChange={(event) =>
                                      setFilterSearch((current) => ({
                                        ...current,
                                        [key]: event.target.value,
                                      }))
                                    }
                                    placeholder={`Search ${labelMap[key].toLowerCase()}`}
                                  />
                                  <div className="filterOptionList">
                                    {visibleValues.length ? (
                                      visibleValues.map((value) => (
                                      <label key={value} className="filterOptionItem">
                                        <input
                                          type="checkbox"
                                          checked={selectedValues.includes(value)}
                                          onChange={() => toggleFilterValue(key, value)}
                                        />
                                        <span>
                                          {key === "frequencies"
                                            ? formatFrequency(Number(value))
                                            : value}
                                        </span>
                                      </label>
                                      ))
                                    ) : (
                                      <div className="filterEmptyState">
                                        No matching options
                                      </div>
                                    )}
                                  </div>
                                </div>
                              ) : null}
                            </div>
                          );
                        })}
                      </div>

                      <div className="filterMenuActions">
                        <button
                          type="button"
                          className="secondary"
                          onClick={clearFilters}
                        >
                          Clear filters
                        </button>
                      </div>
                    </div>
                  ) : null}
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
                  {filteredRows.length ? (
                    <table className="resultsTable">
                      <colgroup>
                        <col className="colUnitId" />
                        <col className="colFrequency" />
                        <col className="colTrp" />
                        <col className="colPeak" />
                        <col className="colPhoto" />
                      </colgroup>
                      <tbody>
                        {filteredRows.map((row, index) => (
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
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <div className="emptyStateTable">No data loaded yet.</div>
                  )}
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

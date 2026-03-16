export type ResultRow = {
  unitId: string;
  unitType: string;
  frequencyMHz?: number;
  trp?: number;
  maxPeak?: number;
  graphValue?: string;
  photoValue?: string;
};

export type SummaryData = {
  rows: ResultRow[];
  uniqueUnitIds: string[];
  uniqueFrequencies: number[];
  uniqueUnitTypes: string[];
};

export type MatrixCell = string | number | boolean | null | undefined;
export type SheetMatrix = MatrixCell[][];

export type ReportBuildParams = {
  title: string;
  author: string;
  dateText: string;
  scopeOfTesting: string;
  fwVersion: string;
  hwVersion: string;
  testedPower: string;
  frequencies: number[];
  unitIds: string[];
  rows: ResultRow[];
};

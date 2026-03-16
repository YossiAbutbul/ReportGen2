export type ResultRow = {
  unitId: string;
  frequencyMHz?: number;
  trp?: number;
  maxPeak?: number;
  graphValue?: string;
};

export type SummaryData = {
  rows: ResultRow[];
  uniqueUnitIds: string[];
  uniqueFrequencies: number[];
};

export type MatrixCell = string | number | null;
export type SheetMatrix = MatrixCell[][];

export type ReportBuildParams = {
  title: string;
  author: string;
  dateText: string;
  fwVersion: string;
  hwVersion: string;
  testedPower: string;
  frequencies: number[];
  unitIds: string[];
  rows: ResultRow[];
};

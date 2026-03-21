# ReportGen2

ReportGen2 is a browser-based test report generator built with React, TypeScript, and Vite.

It lets you:

- upload Excel test result files
- parse units from multiple sheets
- filter results by unit type, unit ID, and frequency
- preview parsed data before export
- generate a Word report with grouped tables by unit type

## Live App

https://yossiabutbul.github.io/ReportGen2/

## Main Features

- Excel parsing across multiple worksheets
- Unit type detection from sheet names and section headers
- Searchable cascading filters
- Word report generation with separate sections per unit type
- Local browser-based workflow with no backend required

## Tech Stack

- React
- TypeScript
- Vite
- `exceljs`
- `docx`

## Local Development

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

## Deploy To GitHub Pages

```bash
npm run deploy
```

## Repository

https://github.com/YossiAbutbul/ReportGen2

// Evaluates an ExcelJS workbook's formulas with HyperFormula (an Excel-compatible engine).
// This is the independent reference the web app is tested against.
import { HyperFormula } from "hyperformula";

function toSheets(wb) {
  const sheets = {};
  wb.eachSheet((ws) => {
    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const cells = [];
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        cells[c - 1] = cell.formula ? `=${cell.formula}` : (cell.value ?? null);
      });
      rows[r - 1] = cells;
    });
    sheets[ws.name] = Array.from(rows, (r) => r ?? []);
  });
  return sheets;
}

// overrides: { Inputs: { B2: 100000, ... } }. Caller must hf.destroy().
export function openWorkbook(wb, overrides = {}) {
  const hf = HyperFormula.buildFromSheets(toSheets(wb), { licenseKey: "gpl-v3" });
  for (const [sheetName, cells] of Object.entries(overrides)) {
    const sheet = hf.getSheetId(sheetName);
    for (const [ref, value] of Object.entries(cells)) {
      const { row, col } = hf.simpleCellAddressFromString(ref, sheet);
      hf.setCellContents({ sheet, row, col }, value);
    }
  }
  return hf;
}

export function evaluateWorkbook(wb, overrides = {}) {
  const hf = openWorkbook(wb, overrides);
  const out = {};
  for (const name of hf.getSheetNames()) out[name] = hf.getSheetValues(hf.getSheetId(name));
  hf.destroy();
  return out;
}

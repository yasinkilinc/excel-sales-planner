// Builds sales-planner.xlsx: the "client spreadsheet" the web app reproduces.
// Cached results are computed by HyperFormula so the file opens with values in any viewer.
import ExcelJS from "exceljs";
import { evaluateWorkbook } from "./test/workbook-engine.mjs";

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const WEIGHTS = [0.7, 0.8, 1, 1.1, 1.1, 1, 0.8, 0.8, 1.1, 1.2, 1.2, 0.9];

const inputs = [
  ["Annual income goal ($)", 120000],
  ["Average commission per sale ($)", 850],
  ["Lead → appointment rate", 0.3],
  ["Appointment show rate", 0.8],
  ["Close rate", 0.25],
  ["Working weeks per year", 48],
  ["Working days per week", 5],
];

const summary = [
  ["Sales needed", "ROUNDUP(Inputs!B2/Inputs!B3,0)"],
  ["Appointments held", "ROUNDUP(B2/Inputs!B6,0)"],
  ["Appointments booked", "ROUNDUP(B3/Inputs!B5,0)"],
  ["Leads needed", "ROUNDUP(B4/Inputs!B4,0)"],
  ["Sales per week", "ROUND(B2/Inputs!B7,1)"],
  ["Appointments per week", "ROUND(B3/Inputs!B7,1)"],
  ["Leads per week", "ROUND(B5/Inputs!B7,1)"],
  ["Leads per working day", "ROUND(B5/(Inputs!B7*Inputs!B8),1)"],
  ["Projected income ($)", "B2*Inputs!B3"],
  ["Surplus over goal ($)", "B10-Inputs!B2"],
];

const wb = new ExcelJS.Workbook();
wb.calcProperties.fullCalcOnLoad = true;

const inSheet = wb.addWorksheet("Inputs");
inSheet.addRow(["Input", "Value", "Month", "Seasonality weight"]);
for (let i = 0; i < 12; i++) {
  const [label, value] = inputs[i] ?? [null, null];
  inSheet.addRow([label, value, MONTHS[i], WEIGHTS[i]]);
}

const plan = wb.addWorksheet("Plan");
plan.addRow(["Metric", "Value"]);
for (const [label, formula] of summary) plan.addRow([label, { formula }]);
plan.addRow([]);
plan.addRow(["Month", "Share", "Income target ($)", "Sales", "Appointments", "Leads"]);
for (let i = 0; i < 12; i++) {
  const r = 14 + i;
  plan.addRow([
    { formula: `Inputs!C${i + 2}` },
    { formula: `Inputs!D${i + 2}/SUM(Inputs!$D$2:$D$13)` },
    { formula: `ROUND(Inputs!$B$2*B${r},2)` },
    { formula: `ROUND($B$2*B${r},1)` },
    { formula: `ROUND($B$3*B${r},1)` },
    { formula: `ROUND($B$5*B${r},1)` },
  ]);
}
plan.addRow(["Total", ..."BCDEF".split("").map((c) => ({ formula: `SUM(${c}14:${c}25)` }))]);

// Fill cached results from the formula engine
const values = evaluateWorkbook(wb);
plan.eachRow((row, r) =>
  row.eachCell((cell, c) => {
    if (cell.formula) cell.value = { formula: cell.formula, result: values.Plan[r - 1][c - 1] };
  }),
);

for (const ws of [inSheet, plan]) {
  ws.getRow(1).font = { bold: true };
  ws.columns.forEach((col) => (col.width = 24));
}

await wb.xlsx.writeFile("sales-planner.xlsx");
console.log("sales-planner.xlsx written");

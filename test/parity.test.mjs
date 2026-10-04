// Parity test: for many input sets, every output of calc.js must equal the workbook's own formulas.
import { test } from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { openWorkbook } from "./workbook-engine.mjs";
import { DEFAULTS, plan, round, roundUp } from "../calc.js";

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(new URL("../sales-planner.xlsx", import.meta.url).pathname);

function toOverrides(i) {
  const cells = { B2: i.goal, B3: i.commission, B4: i.leadToAppointment, B5: i.showRate, B6: i.closeRate, B7: i.weeks, B8: i.daysPerWeek };
  i.weights.forEach((w, m) => (cells[`D${m + 2}`] = w));
  return { Inputs: cells };
}

// calc.js output laid out like the Plan sheet: { "B2": 142, ... }
function webCells(p) {
  const cells = {};
  ["sales", "held", "booked", "leads", "salesPerWeek", "heldPerWeek", "leadsPerWeek", "leadsPerDay", "income", "surplus"].forEach(
    (k, n) => (cells[`B${n + 2}`] = p.summary[k]),
  );
  [...p.months, { month: "Total", ...p.totals }].forEach((m, n) => {
    const r = 14 + n;
    Object.assign(cells, { [`A${r}`]: m.month, [`B${r}`]: m.share, [`C${r}`]: m.income, [`D${r}`]: m.sales, [`E${r}`]: m.held, [`F${r}`]: m.leads });
  });
  return cells;
}

const close = (a, b) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));

// HyperFormula rounds x.x4999999999999996 down where Excel, working with 15 significant digits, sees x.x5 and rounds up
// (Excel: =ROUND(2.675,2) -> 2.68). On such a tie we take the engine's unrounded value and apply Excel's rule,
// then write that value back so dependent cells (totals) are recomputed from it.
function excelTie(hf, sheet, formula, got) {
  const m = formula.match(/^ROUND\((.*),(\d)\)$/);
  if (!m) return false;
  const raw = hf.calculateFormula(`=${m[1]}`, sheet);
  const scaled = Number((Math.abs(raw) * 10 ** Number(m[2])).toPrecision(15));
  if (scaled % 1 !== 0.5) return false;
  return close(got, (Math.sign(raw) * Math.ceil(scaled)) / 10 ** Number(m[2]));
}

const planSheet = wb.getWorksheet("Plan");
const formulaCells = [];
planSheet.eachRow((row) => row.eachCell((cell) => cell.formula && formulaCells.push([cell.address, cell.formula])));

let ties = 0;
function checkCase(name, inputs) {
  const hf = openWorkbook(wb, toOverrides(inputs));
  const sheet = hf.getSheetId("Plan");
  const got = webCells(plan(inputs));
  for (const [ref, formula] of formulaCells) {
    const want = hf.getCellValue(hf.simpleCellAddressFromString(ref, sheet));
    if (typeof want === "number" ? close(got[ref], want) : got[ref] === want) continue;
    if (excelTie(hf, sheet, formula, got[ref])) {
      hf.setCellContents(hf.simpleCellAddressFromString(ref, sheet), got[ref]);
      ties++;
      continue;
    }
    hf.destroy();
    assert.fail(`${name}: Plan!${ref} =${formula} -> web ${got[ref]} vs workbook ${want}`);
  }
  hf.destroy();
}

// Deterministic pseudo-random inputs (mulberry32)
function rng(seed) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomInputs(r) {
  const pick = (min, max, step) => Math.round((min + r() * (max - min)) / step) * step;
  return {
    goal: pick(10000, 500000, 500),
    commission: pick(50, 5000, 5),
    leadToAppointment: pick(0.05, 1, 0.01),
    showRate: pick(0.3, 1, 0.01),
    closeRate: pick(0.05, 1, 0.01),
    weeks: pick(20, 52, 1),
    daysPerWeek: pick(1, 7, 1),
    weights: Array.from({ length: 12 }, () => pick(0, 2, 0.1)),
  };
}

const cases = [
  ["workbook defaults", DEFAULTS],
  ["exact division (no rounding up)", { ...DEFAULTS, goal: 85000, commission: 850 }],
  ["float noise: 0.1+0.2 style rates", { ...DEFAULTS, leadToAppointment: 0.1, showRate: 0.7, closeRate: 0.3 }],
  ["100% rates", { ...DEFAULTS, leadToAppointment: 1, showRate: 1, closeRate: 1 }],
  ["single active month", { ...DEFAULTS, weights: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5] }],
];
const r = rng(20261004);
for (let n = 0; n < 300; n++) {
  const i = randomInputs(r);
  if (i.weights.some((w) => w > 0)) cases.push([`random #${n + 1}`, i]);
}

test(`web calculator matches the workbook (${cases.length} input sets x ${formulaCells.length} formulas)`, () => {
  for (const [name, inputs] of cases) checkCase(name, inputs);
  console.log(`checked ${cases.length * formulaCells.length} cells; ${ties} rounding ties resolved with Excel's 15-digit rule`);
});

test("Excel rounding semantics", () => {
  assert.equal(roundUp(141.00000000000003, 0), 141);
  assert.equal(roundUp(141.2, 0), 142);
  assert.equal(round(2.5, 0), 3);
  assert.equal(round(-2.5, 0), -3);
  assert.equal(round(1.005, 2), 1.01);
});

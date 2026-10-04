# Excel → Web: Sales Plan Calculator

A working example of turning a spreadsheet into a browser app **without changing a single result**.

**Live demo:** https://yasinkilinc.github.io/excel-sales-planner/

- [`sales-planner.xlsx`](sales-planner.xlsx): the "client" workbook, 87 formulas across two sheets (`Inputs`, `Plan`)
- [`calc.js`](calc.js): the same logic in plain JavaScript, no dependencies
- [`index.html`](index.html): the app, a single static page that runs fully in the browser

## How the numbers are verified

`npm test` opens the `.xlsx`, evaluates its own formulas with [HyperFormula](https://hyperformula.handsontable.com/) (an Excel-compatible engine), and compares **every formula cell** with `calc.js` for 305 input sets: the workbook defaults, edge cases (exact divisions, 100% rates, a single active month) and 300 seeded random scenarios.

```
checked 26535 cells; 45 rounding ties resolved with Excel's 15-digit rule
✔ web calculator matches the workbook (305 input sets x 87 formulas)
```

### The rounding detail most ports get wrong

Excel works with 15 significant digits. `504 × 0.5 / 9.6` is `26.249999999999996` in binary floating point, but Excel sees `26.25` and `ROUND(…, 1)` gives **26.3**. A plain JavaScript port gives 26.2. `calc.js` follows Excel's behaviour (the same reason `=ROUND(2.675, 2)` returns 2.68 in Excel). HyperFormula itself rounds these ties down, so the test recognises exact ties, applies Excel's rule to the engine's unrounded value, and feeds the result back so totals are checked too.

## Run locally

```bash
npm install
npm run build   # regenerates sales-planner.xlsx
npm test
npx serve .     # or any static file server
```

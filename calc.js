// Pure JavaScript port of sales-planner.xlsx. No dependencies; runs in the browser and in Node.

export const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export const DEFAULTS = {
  goal: 120000,
  commission: 850,
  leadToAppointment: 0.3,
  showRate: 0.8,
  closeRate: 0.25,
  weeks: 48,
  daysPerWeek: 5,
  weights: [0.7, 0.8, 1, 1.1, 1.1, 1, 0.8, 0.8, 1.1, 1.2, 1.2, 0.9],
};

// Excel works with 15 significant digits: 141.00000000000003 is 141 to Excel, so ROUNDUP gives 141, not 142.
const excelNumber = (x) => Number(x.toPrecision(15));

// Excel ROUND: half away from zero
export function round(x, digits) {
  const f = 10 ** digits;
  const v = excelNumber(Math.abs(x) * f);
  return (Math.sign(x) * Math.floor(v + 0.5)) / f;
}

// Excel ROUNDUP: away from zero
export function roundUp(x, digits) {
  const f = 10 ** digits;
  const v = excelNumber(Math.abs(x) * f);
  return (Math.sign(x) * Math.ceil(v)) / f;
}

export function validate(i) {
  const errors = [];
  if (!(i.goal > 0)) errors.push("Income goal must be greater than 0.");
  if (!(i.commission > 0)) errors.push("Commission per sale must be greater than 0.");
  for (const [key, label] of [
    ["leadToAppointment", "Lead → appointment rate"],
    ["showRate", "Show rate"],
    ["closeRate", "Close rate"],
  ]) {
    if (!(i[key] > 0 && i[key] <= 1)) errors.push(`${label} must be between 0% and 100%.`);
  }
  if (!(i.weeks > 0)) errors.push("Working weeks must be greater than 0.");
  if (!(i.daysPerWeek > 0)) errors.push("Working days per week must be greater than 0.");
  if (!(i.weights.reduce((a, b) => a + b, 0) > 0)) errors.push("At least one month needs a weight above 0.");
  return errors;
}

export function plan(i) {
  const sales = roundUp(i.goal / i.commission, 0);
  const held = roundUp(sales / i.closeRate, 0);
  const booked = roundUp(held / i.showRate, 0);
  const leads = roundUp(booked / i.leadToAppointment, 0);
  const income = sales * i.commission;

  const summary = {
    sales,
    held,
    booked,
    leads,
    salesPerWeek: round(sales / i.weeks, 1),
    heldPerWeek: round(held / i.weeks, 1),
    leadsPerWeek: round(leads / i.weeks, 1),
    leadsPerDay: round(leads / (i.weeks * i.daysPerWeek), 1),
    income,
    surplus: income - i.goal,
  };

  const weightSum = i.weights.reduce((a, b) => a + b, 0);
  const months = MONTHS.map((month, m) => {
    const share = i.weights[m] / weightSum;
    return {
      month,
      share,
      income: round(i.goal * share, 2),
      sales: round(sales * share, 1),
      held: round(held * share, 1),
      leads: round(leads * share, 1),
    };
  });

  const sum = (key) => months.reduce((a, m) => a + m[key], 0);
  const totals = { share: sum("share"), income: sum("income"), sales: sum("sales"), held: sum("held"), leads: sum("leads") };

  return { summary, months, totals };
}

/* Formatters and constants — no demo data */

const PE_INSIGHT_TYPES = [
  { key: "full",      label: "Full Investment Analysis" },
  { key: "fee",       label: "Fee & Cost Analysis" },
  { key: "liquidity", label: "Distribution & Liquidity Planning" },
  { key: "exit",      label: "Exit Strategy Assessment" },
  { key: "pacing",    label: "Commitment Utilization & Pacing" },
];

const HF_INSIGHT_TYPES = [
  { key: "full",      label: "Full HF Performance Analysis" },
  { key: "fee",       label: "Fee & Carry Analysis" },
  { key: "waterfall", label: "Waterfall & Distribution Analysis" },
  { key: "stress",    label: "Stress Test Assessment" },
  { key: "nav",       label: "Unit Price & NAV Attribution" },
];

const fmtMoney = (v, ccy = "USD") => {
  if (v == null || isNaN(v)) return "—";
  const sym = { USD: "$", EUR: "€", GBP: "£" }[ccy] || "$";
  const abs = Math.abs(v);
  let s;
  if (abs >= 1e9) s = (v / 1e9).toFixed(2) + "B";
  else if (abs >= 1e6) s = (v / 1e6).toFixed(2) + "M";
  else if (abs >= 1e3) s = (v / 1e3).toFixed(1) + "K";
  else s = v.toFixed(0);
  return `${sym}${s}`;
};

const fmtMoneyFull = (v, ccy = "USD") => {
  if (v == null || isNaN(v)) return "—";
  const sym = { USD: "$", EUR: "€", GBP: "£" }[ccy] || "$";
  return sym + v.toLocaleString(undefined, { maximumFractionDigits: 0 });
};

const fmtPct  = (v) => (v == null || isNaN(v)) ? "—" : (v * 100).toFixed(1) + "%";
const fmtNum  = (v, d = 2) => (v == null || isNaN(v)) ? "—" : v.toFixed(d);

window.PE_INSIGHT_TYPES  = PE_INSIGHT_TYPES;
window.HF_INSIGHT_TYPES  = HF_INSIGHT_TYPES;
window.fmtMoney          = fmtMoney;
window.fmtMoneyFull      = fmtMoneyFull;
window.fmtPct            = fmtPct;
window.fmtNum            = fmtNum;

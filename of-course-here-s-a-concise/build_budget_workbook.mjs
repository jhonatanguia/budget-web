import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();

const outputDir = path.resolve("outputs/monthly-budget-template");
const outputPath = path.join(outputDir, "monthly_budget_allocation_template.xlsx");

const months = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const allocationBuckets = [
  { label: "TFSA / House", rate: 0.1 },
  { label: "Emergency fund", rate: 0.05 },
  { label: "Other investments", rate: 0.1 },
  { label: "Debt payoff", rate: 0.15 },
  { label: "Household reserve", rate: 0.05 },
  { label: "Travel/Gifts fund", rate: 0.05 },
];
const creditCardCategories = ["CIBC", "RBC", "AMEX", "Wealthsimple", "Triangle", "Other Credit Cards"];
const householdCategories = [
  "Food",
  "Cars",
  "House",
  "Travel",
  "Gift/Education",
  "Custom 1",
  "Custom 2",
  "Custom 3",
  "Business",
  "Monthly Subscription / Utilities",
  "Others",
];
const expenseCategories = [...creditCardCategories, ...householdCategories];
const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const annualCategoryStartRow = 34;
const monthlyCategoryStartRow = 60;

const currencyFormat = '$#,##0;[Red]($#,##0);-';
const percentFormat = '0.0%';
const palette = {
  ink: "#16323a",
  teal: "#0f766e",
  tealDark: "#0f4f4a",
  navy: "#0b1f33",
  sea: "#d8f3ef",
  mint: "#edf9f7",
  amber: "#fff3c4",
  coral: "#f9735b",
  lavender: "#eef2ff",
  steel: "#eef6ff",
  green: "#008000",
  blue: "#0000ff",
  paper: "#fbfcfb",
  line: "#cbd5d1",
  softLine: "#e5e7eb",
  white: "#ffffff",
};

const startHere = workbook.worksheets.add("Start Here");
const overview = workbook.worksheets.add("Annual Overview");
const monthSheets = new Map(months.map((month) => [month, workbook.worksheets.add(month)]));

function mergeTitle(sheet, range, text, fill = palette.tealDark) {
  const title = sheet.getRange(range);
  title.values = [[text]];
  title.merge();
  title.format = {
    fill,
    font: { name: "Aptos Display", size: 18, bold: true, color: palette.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

function sectionHeader(sheet, range, text, fill = palette.teal) {
  const header = sheet.getRange(range);
  header.values = [[text]];
  header.merge();
  header.format = {
    fill,
    font: { name: "Aptos", size: 12, bold: true, color: palette.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
}

function styleBox(range, fill = palette.white) {
  range.format = {
    fill,
    font: { name: "Aptos", size: 10, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.softLine },
    verticalAlignment: "center",
  };
}

function styleHeader(range, fill = palette.sea) {
  range.format = {
    fill,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function styleTotal(range) {
  range.format = {
    fill: palette.amber,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    verticalAlignment: "center",
  };
}

function styleInput(range) {
  range.format = {
    fill: palette.amber,
    font: { name: "Aptos", size: 10, color: palette.blue },
    borders: { preset: "all", style: "thin", color: palette.line },
    verticalAlignment: "center",
  };
}

function styleCard(range, fill = palette.white) {
  range.format = {
    fill,
    font: { name: "Aptos", size: 10, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    verticalAlignment: "top",
    wrapText: true,
  };
}

function setColumnWidths(sheet, widths) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = width;
  }
}

function money(range) {
  range.format.numberFormat = currencyFormat;
  range.format.horizontalAlignment = "right";
}

function percent(range) {
  range.format.numberFormat = percentFormat;
  range.format.horizontalAlignment = "right";
}

function colName(index) {
  let name = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function sumFormula(cells) {
  return `=SUM(${cells.join(",")})`;
}

function buildStartHere() {
  startHere.showGridlines = false;
  setColumnWidths(startHere, {
    A: 24,
    B: 180,
    C: 190,
    D: 210,
    E: 28,
    F: 205,
    G: 205,
    H: 210,
    I: 210,
  });

  mergeTitle(startHere, "B2:I2", "Household Budget System", palette.navy);
  startHere.getRange("B3:I3").values = [["A polished workbook for weekly cash flow, household expenses, rolling trends, and automatic allocation planning."]];
  startHere.getRange("B3:I3").merge();
  startHere.getRange("B3:I3").format = {
    fill: palette.steel,
    font: { name: "Aptos", size: 11, bold: true, color: palette.ink },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };

  sectionHeader(startHere, "B5:D5", "Quick Start", palette.teal);
  startHere.getRange("B6:D10").values = [
    ["1", "Set allocations", "Go to Annual Overview and enter target percentages. Blank buckets automatically share what is left."],
    ["2", "Pick allocation mode", "Use the dropdown for Month-to-date smoothing or Independent weekly."],
    ["3", "Rename custom categories", "Edit Custom 1, Custom 2, and Custom 3 on Annual Overview if you want categories like Gym, Medical, or Pets."],
    ["4", "Enter each week", "Use the month tabs. Add income and expenses line by line; weekly and monthly totals calculate automatically."],
    ["5", "Review dashboard", "Use Annual Overview and the Spending Focus panels to see trends, biggest categories, and cash flow."],
  ];
  styleCard(startHere.getRange("B6:D10"), palette.white);
  startHere.getRange("B6:B10").format = {
    fill: palette.tealDark,
    font: { name: "Aptos", size: 11, bold: true, color: palette.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: palette.line },
  };

  sectionHeader(startHere, "F5:I5", "How The Logic Works", palette.coral);
  startHere.getRange("F6:I10").values = [
    ["Income less expenses", "Each week calculates income minus tracked expenses.", "Allocations", "Positive cash flow is assigned to TFSA / House, investments, debt, and reserves."],
    ["Month-to-date smoothing", "If week 1 is positive and week 2 is negative, week 2 reduces the prior allocation.", "Independent weekly", "Each week stands alone and negative weeks do not pull back prior allocations."],
    ["Effective percentages", "If you type 100% in one bucket, the others become 0%. If you leave blanks, the remaining percentage spreads evenly.", "Checks", "Every month has an OK check to confirm totals reconcile."],
    ["Rolling averages", "The dashboard tracks recent 3-month average income and expenses.", "Spending focus", "The workbook highlights where most money is going each month and across the year."],
    ["Custom categories", "Custom rows are treated as real expense categories.", "Compatibility", "The file is built as an Excel workbook and can be opened in Excel or Apple Numbers."],
  ];
  styleCard(startHere.getRange("F6:I10"), palette.paper);
  startHere.getRange("F6:F10").format.font = { name: "Aptos", size: 10, bold: true, color: palette.ink };
  startHere.getRange("H6:H10").format.font = { name: "Aptos", size: 10, bold: true, color: palette.ink };

  sectionHeader(startHere, "B12:D12", "Input Legend", palette.tealDark);
  startHere.getRange("B13:D17").values = [
    ["Yellow cells", "User inputs", "Percentages, dropdown choices, and custom category names."],
    ["White cells", "Weekly entries", "Enter income and expenses by line item."],
    ["Pale green cells", "Calculated results", "Totals, trends, allocations, and checks."],
    ["Coral headers", "Setup or category areas", "Use these to understand what is being tracked."],
    ["Dark green headers", "Dashboard / summary areas", "Review performance and next decisions."],
  ];
  styleCard(startHere.getRange("B13:D17"), palette.white);
  startHere.getRange("B13:B17").format = {
    fill: palette.lavender,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
  };

  sectionHeader(startHere, "F12:I12", "Dashboard Questions Answered", palette.tealDark);
  startHere.getRange("F13:I17").values = [
    ["Am I spending less?", "Expense direction compares the latest month to the prior month.", "What is normal lately?", "3-month averages smooth income and expense swings."],
    ["Where should I adjust?", "Top category by month shows the biggest spending area for each month.", "What is the yearly pressure point?", "Biggest annual category shows the largest full-year spending bucket."],
    ["How much can be invested?", "Net before allocations and allocation totals show available cash flow.", "What if a week is negative?", "Month-to-date smoothing subtracts it from the month allocation total."],
    ["Can two people use it?", "The categories and reserves are designed for a two-person household.", "Can I personalize it?", "Use the custom categories and allocation percentages to fit your household."],
    ["What should I sell as value?", "Weekly entry, monthly summaries, yearly dashboard, rolling averages, and guided controls.", "What stays automatic?", "Totals, charts, trends, category rankings, and allocation split."],
  ];
  styleCard(startHere.getRange("F13:I17"), palette.paper);
  startHere.getRange("F13:F17").format.font = { name: "Aptos", size: 10, bold: true, color: palette.ink };
  startHere.getRange("H13:H17").format.font = { name: "Aptos", size: 10, bold: true, color: palette.ink };

  sectionHeader(startHere, "B20:I20", "Workbook Map", palette.coral);
  startHere.getRange("B21:D26").values = [
    ["Area", "Purpose", "Use"],
    ["Start Here", "Guide", "Read before using or reselling the workbook."],
    ["Annual Overview", "Dashboard and controls", "Set allocations, rename custom categories, and review full-year trends."],
    ["Month tabs", "Weekly entry", "Enter income and expense line items in the correct month."],
    ["Spending Focus", "Decision support", "See the top category, rolling averages, and whether expenses are rising or falling."],
    ["Checks", "Quality control", "Look for OK in each month after entering data."],
  ];
  startHere.getRange("F21:I26").values = [
    ["Feature", "What it does", "Where", "Automatic"],
    ["Dropdown mode", "Controls how weekly allocation timing works.", "Annual Overview", "Yes"],
    ["Custom categories", "Lets users rename three flexible expense columns.", "Annual Overview", "Labels flow to months"],
    ["Allocation engine", "Splits positive cash flow by effective percentages.", "All months", "Yes"],
    ["3-month trend", "Shows recent average income and expenses.", "Annual + months", "Yes"],
    ["Top category", "Highlights the biggest spending bucket.", "Annual + months", "Yes"],
  ];
  styleHeader(startHere.getRange("B21:D21"), palette.sea);
  styleCard(startHere.getRange("B22:D26"), palette.white);
  styleHeader(startHere.getRange("F21:I21"), palette.sea);
  styleCard(startHere.getRange("F22:I26"), palette.white);
  startHere.getRange("B22:B26").format.font = { name: "Aptos", size: 10, bold: true, color: palette.ink };
  startHere.getRange("F22:F26").format.font = { name: "Aptos", size: 10, bold: true, color: palette.ink };

  startHere.getRange("B28:I28").values = [["Recommended first setup: enter allocation percentages, pick the allocation mode dropdown, rename the three custom categories, then start with January."]];
  startHere.getRange("B28:I28").merge();
  startHere.getRange("B28:I28").format = {
    fill: palette.amber,
    font: { name: "Aptos", size: 11, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function buildOverview() {
  overview.showGridlines = false;
  setColumnWidths(overview, {
    A: 260,
    B: 132,
    C: 132,
    D: 152,
    E: 205,
    F: 132,
    G: 118,
    H: 132,
    I: 136,
    J: 136,
    K: 132,
    L: 18,
    M: 190,
    N: 160,
    O: 130,
    P: 118,
    Q: 118,
    R: 76,
    S: 96,
    T: 96,
    U: 76,
    V: 96,
    W: 96,
    X: 112,
    Y: 112,
    Z: 112,
  });

  mergeTitle(overview, "A1:K1", "Annual Budget Overview");

  sectionHeader(overview, "A3:C3", "Allocation Settings", palette.coral);
  overview.getRange("A4:C12").values = [
    ["Bucket", "Target %", "Effective %"],
    ...allocationBuckets.map((bucket, idx) => [`${bucket.label} %`, idx === 0 ? 0.5 : null, null]),
    ["Effective allocation total %", null, null],
    ["Remaining unallocated %", null, null],
  ];
  const effectiveAllocationFormula = (row) =>
    `=IF(SUM($B$5:$B$10)>1,IF(B${row}<>"",B${row}/SUM($B$5:$B$10),0),IF(B${row}<>"",MIN(MAX(B${row},0),1),IF(COUNTBLANK($B$5:$B$10)>0,MAX(1-SUM($B$5:$B$10),0)/COUNTBLANK($B$5:$B$10),0)))`;
  overview.getRange("C5:C12").formulas = [
    ...[5, 6, 7, 8, 9, 10].map((row) => [effectiveAllocationFormula(row)]),
    ["=SUM(C5:C10)"],
    ["=MAX(1-C11,0)"],
  ];
  styleHeader(overview.getRange("A4:C4"), palette.sea);
  styleBox(overview.getRange("A5:C12"), palette.paper);
  overview.getRange("B5:B10").format = {
    fill: palette.amber,
    font: { name: "Aptos", size: 10, color: palette.blue },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  overview.getRange("B11:C12").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  percent(overview.getRange("B5:C12"));
  overview.getRange("B5:B10").dataValidation = {
    allowBlank: true,
    rule: { type: "decimal", operator: "between", formula1: 0, formula2: 1 },
    errorAlert: {
      style: "stop",
      title: "Enter a percentage",
      message: "Use a percentage between 0% and 100%.",
    },
  };

  sectionHeader(overview, "A14:C14", "Allocation Timing", palette.teal);
  overview.getRange("A15:B15").values = [["Weekly allocation mode", "Month-to-date smoothing"]];
  styleBox(overview.getRange("A15:B15"), palette.paper);
  overview.getRange("B15").format = {
    fill: palette.amber,
    font: { name: "Aptos", size: 10, color: palette.blue },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "left",
  };
  overview.getRange("B15").dataValidation = {
    allowBlank: false,
    list: { inCellDropDown: true, source: ["Month-to-date smoothing", "Independent weekly"] },
    errorAlert: {
      style: "stop",
      title: "Pick an allocation mode",
      message: "Choose Month-to-date smoothing or Independent weekly.",
    },
  };

  sectionHeader(overview, "E3:H3", "Annual Summary");
  overview.getRange("E4:F14").values = [
    ["Total annual income", null],
    ["Total annual expenses", null],
    ["Net before allocations", null],
    ["TFSA / House allocation", null],
    ["Emergency fund allocation", null],
    ["Other investments allocation", null],
    ["Debt payoff allocation", null],
    ["Household reserve allocation", null],
    ["Travel/Gifts fund allocation", null],
    ["Total allocations", null],
    ["Net after allocations", null],
  ];
  overview.getRange("F4:F14").formulas = [
    ["=B30"],
    ["=C30"],
    ["=D30"],
    ["=E30"],
    ["=F30"],
    ["=G30"],
    ["=H30"],
    ["=I30"],
    ["=J30"],
    ["=SUM(F7:F12)"],
    ["=K30"],
  ];
  styleBox(overview.getRange("E4:F14"), palette.paper);
  overview.getRange("F4:F14").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  money(overview.getRange("F4:F14"));

  sectionHeader(overview, "A16:K16", "Monthly Rollup");
  overview.getRange("A17:K17").values = [[
    "Month",
    "Income",
    "Expenses",
    "Net before allocations",
    "TFSA / House",
    "Emergency fund",
    "Investments",
    "Debt payoff",
    "Household reserve",
    "Travel/Gifts fund",
    "Net after allocations",
  ]];
  styleHeader(overview.getRange("A17:K17"));
  overview.getRange("A18:A29").values = months.map((month) => [month]);
  overview.getRange("B18:K29").formulas = months.map((month) => [
    `='${month}'!$B$59`,
    `='${month}'!$B$60`,
    `='${month}'!$B$61`,
    `='${month}'!$B$62`,
    `='${month}'!$B$63`,
    `='${month}'!$B$64`,
    `='${month}'!$B$65`,
    `='${month}'!$B$66`,
    `='${month}'!$B$67`,
    `='${month}'!$B$69`,
  ]);
  overview.getRange("A30:K30").values = [["Annual total", null, null, null, null, null, null, null, null, null, null]];
  overview.getRange("B30:K30").formulas = [[
    "=SUM(B18:B29)",
    "=SUM(C18:C29)",
    "=SUM(D18:D29)",
    "=SUM(E18:E29)",
    "=SUM(F18:F29)",
    "=SUM(G18:G29)",
    "=SUM(H18:H29)",
    "=SUM(I18:I29)",
    "=SUM(J18:J29)",
    "=SUM(K18:K29)",
  ]];
  styleBox(overview.getRange("A18:K29"), palette.white);
  overview.getRange("B18:K29").format.font = { name: "Aptos", size: 10, color: palette.green };
  money(overview.getRange("B18:K30"));
  styleTotal(overview.getRange("A30:K30"));
  money(overview.getRange("B30:K30"));

  const annualCategoryEndRow = annualCategoryStartRow + expenseCategories.length - 1;
  const annualCategoryTotalStartRow = annualCategoryEndRow + 1;

  sectionHeader(overview, "A33:B33", "Annual Category Breakdown", palette.coral);
  overview.getRange(`A${annualCategoryStartRow}:B${annualCategoryTotalStartRow + 2}`).values = [
    ...expenseCategories.map((category) => [category, null]),
    ["Credit cards total", null],
    ["Household spending total", null],
    ["Total tracked expenses", null],
  ];
  overview.getRange(`B${annualCategoryStartRow}:B${annualCategoryEndRow}`).formulas = expenseCategories.map((_, idx) => [
    `=${months.map((month) => `'${month}'!$E$${monthlyCategoryStartRow + idx}`).join("+")}`,
  ]);
  overview.getRange(`B${annualCategoryTotalStartRow}:B${annualCategoryTotalStartRow + 2}`).formulas = [
    [`=SUM(B${annualCategoryStartRow}:B${annualCategoryStartRow + creditCardCategories.length - 1})`],
    [`=SUM(B${annualCategoryStartRow + creditCardCategories.length}:B${annualCategoryEndRow})`],
    [`=SUM(B${annualCategoryStartRow}:B${annualCategoryEndRow})`],
  ];
  styleBox(overview.getRange(`A${annualCategoryStartRow}:B${annualCategoryEndRow}`), palette.white);
  styleTotal(overview.getRange(`A${annualCategoryTotalStartRow}:B${annualCategoryTotalStartRow + 2}`));
  money(overview.getRange(`B${annualCategoryStartRow}:B${annualCategoryTotalStartRow + 2}`));
  const customCategoryStartRow = annualCategoryStartRow + expenseCategories.indexOf("Custom 1");
  styleInput(overview.getRange(`A${customCategoryStartRow}:A${customCategoryStartRow + 2}`));
  overview.getRange(`A${customCategoryStartRow}:A${customCategoryStartRow + 2}`).format.font = {
    name: "Aptos",
    size: 10,
    bold: true,
    color: palette.blue,
  };

  sectionHeader(overview, "M18:N18", "Allocation Mix", palette.teal);
  overview.getRange("M19:N24").values = allocationBuckets.map((bucket) => [bucket.label, null]);
  overview.getRange("N19:N24").formulas = [["=E30"], ["=F30"], ["=G30"], ["=H30"], ["=I30"], ["=J30"]];
  styleBox(overview.getRange("M19:N24"), palette.white);
  money(overview.getRange("N19:N24"));

  overview.getRange("U3:Z3").values = [["Month", "Income", "Expenses", "3-mo avg income", "3-mo avg expenses", "Expense change"]];
  overview.getRange("U4:U15").values = monthLabels.map((label) => [label]);
  overview.getRange("V4:Z15").formulas = months.map((month, idx) => {
    const row = 4 + idx;
    const avgStartRow = Math.max(4, row - 2);
    return [
      `=B${18 + idx}`,
      `=C${18 + idx}`,
      `=AVERAGE(V${avgStartRow}:V${row})`,
      `=AVERAGE(W${avgStartRow}:W${row})`,
      idx === 0 ? "=0" : `=W${row}-W${row - 1}`,
    ];
  });
  styleHeader(overview.getRange("U3:Z3"), palette.sea);
  styleBox(overview.getRange("U4:Z15"), palette.white);
  money(overview.getRange("V4:Z15"));

  const incomeExpenseChart = overview.charts.add("ColumnClustered", overview.getRange("U3:Y15"), "Auto");
  incomeExpenseChart.title.text = "Monthly Income, Expenses, and 3-Month Averages";
  incomeExpenseChart.setPosition(overview.getRange("M2:T15"));
  incomeExpenseChart.width = 630;
  incomeExpenseChart.height = 305;
  incomeExpenseChart.yAxis = { numberFormatCode: currencyFormat };

  const allocationChart = overview.charts.add("pie", overview.getRange("M18:N24"), "Auto");
  allocationChart.title = "Annual Allocation Mix";
  allocationChart.setPosition(overview.getRange("O18:T34"));
  allocationChart.width = 480;
  allocationChart.height = 285;

  sectionHeader(overview, "M36:T36", "Trend & Spending Focus", palette.tealDark);
  overview.getRange("M37:N42").values = [
    ["Last 3 months avg income", null],
    ["Last 3 months avg expenses", null],
    ["Latest expense change", null],
    ["Expense direction", null],
    ["Biggest annual category", null],
    ["Biggest annual amount", null],
  ];
  overview.getRange("N37:N42").formulas = [
    ['=IFERROR(LOOKUP(2,1/(($V$4:$V$15+$W$4:$W$15)>0),$X$4:$X$15),0)'],
    ['=IFERROR(LOOKUP(2,1/(($V$4:$V$15+$W$4:$W$15)>0),$Y$4:$Y$15),0)'],
    ['=IFERROR(LOOKUP(2,1/(($V$4:$V$15+$W$4:$W$15)>0),$Z$4:$Z$15),0)'],
    ['=IF(N39<0,"Reducing",IF(N39>0,"Increasing","Flat / no change"))'],
    [`=IF(MAX($B$${annualCategoryStartRow}:$B$${annualCategoryEndRow})=0,"No spending yet",INDEX($A$${annualCategoryStartRow}:$A$${annualCategoryEndRow},MATCH(MAX($B$${annualCategoryStartRow}:$B$${annualCategoryEndRow}),$B$${annualCategoryStartRow}:$B$${annualCategoryEndRow},0)))`],
    [`=MAX($B$${annualCategoryStartRow}:$B$${annualCategoryEndRow})`],
  ];
  styleBox(overview.getRange("M37:N42"), palette.paper);
  overview.getRange("N37:N39").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  money(overview.getRange("N37:N39"));
  overview.getRange("N40:N41").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
  };
  overview.getRange("N42").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  money(overview.getRange("N42"));

  sectionHeader(overview, "M44:O44", "Top Category By Month", palette.coral);
  overview.getRange("M45:O45").values = [["Month", "Top category", "Amount"]];
  overview.getRange("M46:O57").values = months.map((month) => [month, null, null]);
  overview.getRange("N46:O57").formulas = months.map((month) => [
    `=IF(MAX('${month}'!$E$${monthlyCategoryStartRow}:$E$${monthlyCategoryStartRow + expenseCategories.length - 1})=0,"No spending yet",INDEX('${month}'!$D$${monthlyCategoryStartRow}:$D$${monthlyCategoryStartRow + expenseCategories.length - 1},MATCH(MAX('${month}'!$E$${monthlyCategoryStartRow}:$E$${monthlyCategoryStartRow + expenseCategories.length - 1}),'${month}'!$E$${monthlyCategoryStartRow}:$E$${monthlyCategoryStartRow + expenseCategories.length - 1},0)))`,
    `=MAX('${month}'!$E$${monthlyCategoryStartRow}:$E$${monthlyCategoryStartRow + expenseCategories.length - 1})`,
  ]);
  styleHeader(overview.getRange("M45:O45"), palette.sea);
  styleBox(overview.getRange("M46:O57"), palette.white);
  money(overview.getRange("O46:O57"));
}

function buildMonthSheet(month, sheet) {
  const monthIndex = months.indexOf(month);
  const monthlyCategoryEndRow = monthlyCategoryStartRow + expenseCategories.length - 1;
  const monthlyCategoryTotalStartRow = monthlyCategoryEndRow + 1;

  sheet.showGridlines = false;
  setColumnWidths(sheet, {
    A: 190,
    B: 95,
    C: 95,
    D: 210,
    E: 95,
    F: 150,
    G: 100,
    H: 130,
    I: 95,
    J: 85,
    K: 95,
    L: 150,
    M: 135,
    N: 105,
    O: 105,
    P: 105,
    Q: 120,
    R: 220,
    S: 95,
    T: 112,
    U: 104,
    V: 118,
    W: 118,
    X: 104,
    Y: 128,
    Z: 128,
    AA: 124,
    AB: 18,
    AC: 95,
    AD: 110,
    AE: 110,
    AF: 112,
    AG: 112,
    AH: 112,
    AI: 112,
  });

  mergeTitle(sheet, "A1:AA1", `${month} Budget Planner`);

  sectionHeader(sheet, "A3:D3", "Linked Allocation Rates", palette.coral);
  sheet.getRange("A4:C12").values = [
    ["Bucket", "Target %", "Effective %"],
    ...allocationBuckets.map((bucket) => [`${bucket.label} %`, null, null]),
    ["Effective allocation total %", null, null],
    ["Remaining unallocated %", null, null],
  ];
  sheet.getRange("B5:C12").formulas = [
    [`=IF('Annual Overview'!$B$5="","",'Annual Overview'!$B$5)`, "='Annual Overview'!$C$5"],
    [`=IF('Annual Overview'!$B$6="","",'Annual Overview'!$B$6)`, "='Annual Overview'!$C$6"],
    [`=IF('Annual Overview'!$B$7="","",'Annual Overview'!$B$7)`, "='Annual Overview'!$C$7"],
    [`=IF('Annual Overview'!$B$8="","",'Annual Overview'!$B$8)`, "='Annual Overview'!$C$8"],
    [`=IF('Annual Overview'!$B$9="","",'Annual Overview'!$B$9)`, "='Annual Overview'!$C$9"],
    [`=IF('Annual Overview'!$B$10="","",'Annual Overview'!$B$10)`, "='Annual Overview'!$C$10"],
    ['=""', "='Annual Overview'!$C$11"],
    ['=""', "='Annual Overview'!$C$12"],
  ];
  styleHeader(sheet.getRange("A4:C4"), palette.sea);
  styleBox(sheet.getRange("A5:C12"), palette.paper);
  sheet.getRange("B5:C12").format.font = { name: "Aptos", size: 10, color: palette.green };
  percent(sheet.getRange("B5:C12"));

  sectionHeader(sheet, "F3:J3", "Monthly Snapshot");
  sheet.getRange("F4:G6").values = [["Income", null], ["Expenses", null], ["Net after allocations", null]];
  sheet.getRange("G4:G6").formulas = [["=$B$59"], ["=$B$60"], ["=$B$69"]];
  styleBox(sheet.getRange("F4:G6"), palette.paper);
  sheet.getRange("G4:G6").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  money(sheet.getRange("G4:G6"));

  sheet.getRange("F8:G8").values = [["Allocation mode", null]];
  sheet.getRange("G8").formulas = [["='Annual Overview'!$B$15"]];
  styleBox(sheet.getRange("F8:G8"), palette.paper);
  sheet.getRange("G8").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "left",
  };

  sectionHeader(sheet, "L3:P3", "Spending Focus", palette.tealDark);
  sheet.getRange("L4:M9").values = [
    ["Top category", null],
    ["Top amount", null],
    ["3-month avg income", null],
    ["3-month avg expenses", null],
    ["Expense vs prior month", null],
    ["Expense trend", null],
  ];
  sheet.getRange("M4:M9").formulas = [
    [`=IF(MAX($E$${monthlyCategoryStartRow}:$E$${monthlyCategoryEndRow})=0,"No spending yet",INDEX($D$${monthlyCategoryStartRow}:$D$${monthlyCategoryEndRow},MATCH(MAX($E$${monthlyCategoryStartRow}:$E$${monthlyCategoryEndRow}),$E$${monthlyCategoryStartRow}:$E$${monthlyCategoryEndRow},0)))`],
    [`=MAX($E$${monthlyCategoryStartRow}:$E$${monthlyCategoryEndRow})`],
    [`='Annual Overview'!$X$${4 + monthIndex}`],
    [`='Annual Overview'!$Y$${4 + monthIndex}`],
    [`='Annual Overview'!$Z$${4 + monthIndex}`],
    ['=IF(M8<0,"Reducing",IF(M8>0,"Increasing","Flat / no change"))'],
  ];
  styleBox(sheet.getRange("L4:M9"), palette.paper);
  sheet.getRange("M4").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
  };
  sheet.getRange("M5:M8").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  money(sheet.getRange("M5:M8"));
  sheet.getRange("M9").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
  };

  const headers = [
    "Date / note",
    "Income",
    ...expenseCategories,
    "Total expenses",
    ...allocationBuckets.map((bucket) => `${bucket.label} target`),
    "Net cash flow",
  ];
  const expenseStartCol = 2;
  const expenseEndCol = expenseStartCol + expenseCategories.length - 1;
  const totalExpenseCol = expenseEndCol + 1;
  const allocationStartCol = totalExpenseCol + 1;
  const allocationEndCol = allocationStartCol + allocationBuckets.length - 1;
  const netCol = allocationEndCol + 1;
  const customCategoryIndex = expenseCategories.indexOf("Custom 1");
  const customLabelFormula = (customIdx) => `='Annual Overview'!$A$${annualCategoryStartRow + customCategoryIndex + customIdx}`;
  const expenseEndColName = colName(expenseEndCol);
  const totalExpenseColName = colName(totalExpenseCol);
  const netColName = colName(netCol);
  const lastTableColName = colName(netCol);
  const weekStarts = [13, 24, 35, 46];
  const weeklyTotals = [];

  weekStarts.forEach((startRow, idx) => {
    const week = idx + 1;
    sectionHeader(sheet, `A${startRow}:${lastTableColName}${startRow}`, `Week ${week}`, idx % 2 === 0 ? palette.teal : palette.tealDark);
    sheet.getRange(`A${startRow + 1}:${lastTableColName}${startRow + 1}`).values = [headers];
    styleHeader(sheet.getRange(`A${startRow + 1}:${lastTableColName}${startRow + 1}`));
    [0, 1, 2].forEach((customIdx) => {
      const customCol = colName(expenseStartCol + customCategoryIndex + customIdx);
      sheet.getRange(`${customCol}${startRow + 1}`).formulas = [[customLabelFormula(customIdx)]];
    });

    const entryStart = startRow + 2;
    const entryEnd = startRow + 6;
    const totalRow = startRow + 7;
    weeklyTotals.push(totalRow);

    styleBox(sheet.getRange(`A${entryStart}:${lastTableColName}${entryEnd}`), palette.white);
    sheet.getRange(`A${entryStart}:${expenseEndColName}${entryEnd}`).format.font = { name: "Aptos", size: 10, color: palette.blue };
    sheet.getRange(`B${entryStart}:${lastTableColName}${entryEnd}`).format.numberFormat = currencyFormat;
    for (let row = entryStart; row <= entryEnd; row += 1) {
      sheet.getRange(`${totalExpenseColName}${row}`).formulas = [[`=SUM(C${row}:${expenseEndColName}${row})`]];
      sheet.getRange(`${netColName}${row}`).formulas = [[`=B${row}-${totalExpenseColName}${row}`]];
    }
    sheet.getRange(`${totalExpenseColName}${entryStart}:${totalExpenseColName}${entryEnd}`).format.font = { name: "Aptos", size: 10, color: palette.ink };
    sheet.getRange(`${netColName}${entryStart}:${netColName}${entryEnd}`).format.font = { name: "Aptos", size: 10, color: palette.ink };

    sheet.getRange(`A${totalRow}:${lastTableColName}${totalRow}`).values = [["Weekly total", ...Array(headers.length - 1).fill(null)]];
    const currentAndPriorTotals = weeklyTotals.slice();
    const previousTotals = weeklyTotals.slice(0, -1);
    const cumulativeIncome = `SUM(${currentAndPriorTotals.map((row) => `B${row}`).join(",")})`;
    const cumulativeExpenses = `SUM(${currentAndPriorTotals.map((row) => `${totalExpenseColName}${row}`).join(",")})`;
    const weeklyFormulas = [
      `=SUM(B${entryStart}:B${entryEnd})`,
      ...expenseCategories.map((_, categoryIdx) => {
        const col = colName(expenseStartCol + categoryIdx);
        return `=SUM(${col}${entryStart}:${col}${entryEnd})`;
      }),
      `=SUM(${totalExpenseColName}${entryStart}:${totalExpenseColName}${entryEnd})`,
      ...allocationBuckets.map((_, bucketIdx) => {
        const allocationCol = colName(allocationStartCol + bucketIdx);
        const effectiveRateRef = `'Annual Overview'!$C$${5 + bucketIdx}`;
        const independentAllocation = `MAX(B${totalRow}-${totalExpenseColName}${totalRow},0)*${effectiveRateRef}`;
        const priorAllocation = previousTotals.length > 0
          ? `SUM(${previousTotals.map((row) => `${allocationCol}${row}`).join(",")})`
          : "0";
        const smoothedAllocation = `MAX(${cumulativeIncome}-${cumulativeExpenses},0)*${effectiveRateRef}-${priorAllocation}`;
        return `=IF('Annual Overview'!$B$15="Independent weekly",${independentAllocation},${smoothedAllocation})`;
      }),
      `=B${totalRow}-${totalExpenseColName}${totalRow}-SUM(${colName(allocationStartCol)}${totalRow}:${colName(allocationEndCol)}${totalRow})`,
    ];
    sheet.getRange(`B${totalRow}:${lastTableColName}${totalRow}`).formulas = [weeklyFormulas];
    styleTotal(sheet.getRange(`A${totalRow}:${lastTableColName}${totalRow}`));
    money(sheet.getRange(`B${entryStart}:${lastTableColName}${totalRow}`));
  });

  const totalRefs = (col) => weeklyTotals.map((row) => `${col}${row}`).join(",");
  const allocationTotalRefs = allocationBuckets.map((_, bucketIdx) => totalRefs(colName(allocationStartCol + bucketIdx)));

  sectionHeader(sheet, "A57:B57", "Monthly Totals", palette.tealDark);
  sheet.getRange("A59:B71").values = [
    ["Total income", null],
    ["Total expenses", null],
    ["Net before allocations", null],
    ["TFSA / House allocation", null],
    ["Emergency fund allocation", null],
    ["Other investments allocation", null],
    ["Debt payoff allocation", null],
    ["Household reserve allocation", null],
    ["Travel/Gifts fund allocation", null],
    ["Total allocations", null],
    ["Net cash flow after allocations", null],
    ["Allocation % used", null],
    ["Check", null],
  ];
  sheet.getRange("B59:B71").formulas = [
    [`=SUM(${totalRefs("B")})`],
    [`=SUM(${totalRefs(totalExpenseColName)})`],
    ["=B59-B60"],
    ...allocationTotalRefs.map((refs) => [`=SUM(${refs})`]),
    ["=SUM(B62:B67)"],
    ["=B59-B60-B68"],
    ["='Annual Overview'!$C$11"],
    ['=IF(ROUND(B69-(B59-B60-B68),2)=0,"OK","Review")'],
  ];
  styleBox(sheet.getRange("A59:B71"), palette.paper);
  sheet.getRange("B59:B69").format = {
    fill: palette.mint,
    font: { name: "Aptos", size: 10, bold: true, color: palette.ink },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "right",
  };
  money(sheet.getRange("B59:B69"));
  percent(sheet.getRange("B70"));
  sheet.getRange("B71").format = {
    fill: palette.sea,
    font: { name: "Aptos", size: 10, bold: true, color: palette.tealDark },
    borders: { preset: "all", style: "thin", color: palette.line },
    horizontalAlignment: "center",
  };

  sectionHeader(sheet, "D57:E57", "Category Totals", palette.coral);
  sheet.getRange(`D59:E${monthlyCategoryTotalStartRow + 2}`).values = [
    ["Category", "Monthly total"],
    ...expenseCategories.map((category) => [category, null]),
    ["Credit cards total", null],
    ["Household spending total", null],
    ["Total tracked expenses", null],
  ];
  sheet.getRange(`E${monthlyCategoryStartRow}:E${monthlyCategoryEndRow}`).formulas = expenseCategories.map((_, categoryIdx) => [
    `=SUM(${totalRefs(colName(expenseStartCol + categoryIdx))})`,
  ]);
  sheet.getRange(`E${monthlyCategoryTotalStartRow}:E${monthlyCategoryTotalStartRow + 2}`).formulas = [
    [`=SUM(E${monthlyCategoryStartRow}:E${monthlyCategoryStartRow + creditCardCategories.length - 1})`],
    [`=SUM(E${monthlyCategoryStartRow + creditCardCategories.length}:E${monthlyCategoryEndRow})`],
    [`=SUM(E${monthlyCategoryStartRow}:E${monthlyCategoryEndRow})`],
  ];
  styleHeader(sheet.getRange("D59:E59"), palette.sea);
  styleBox(sheet.getRange(`D${monthlyCategoryStartRow}:E${monthlyCategoryEndRow}`), palette.white);
  [0, 1, 2].forEach((customIdx) => {
    sheet.getRange(`D${monthlyCategoryStartRow + customCategoryIndex + customIdx}`).formulas = [[customLabelFormula(customIdx)]];
  });
  styleTotal(sheet.getRange(`D${monthlyCategoryTotalStartRow}:E${monthlyCategoryTotalStartRow + 2}`));
  money(sheet.getRange(`E${monthlyCategoryStartRow}:E${monthlyCategoryTotalStartRow + 2}`));

  sectionHeader(sheet, "AC2:AF2", "Weekly Chart Data", palette.teal);
  sheet.getRange("AC3:AF3").values = [["Week", "Income", "Expenses", "Allocations"]];
  sheet.getRange("AC4:AF7").values = [["Week 1", null, null, null], ["Week 2", null, null, null], ["Week 3", null, null, null], ["Week 4", null, null, null]];
  sheet.getRange("AD4:AF7").formulas = weeklyTotals.map((row) => [
    `=B${row}`,
    `=${totalExpenseColName}${row}`,
    `=SUM(${colName(allocationStartCol)}${row}:${colName(allocationEndCol)}${row})`,
  ]);
  styleHeader(sheet.getRange("AC3:AF3"));
  styleBox(sheet.getRange("AC4:AF7"), palette.white);
  money(sheet.getRange("AD4:AF7"));

  const chart = sheet.charts.add("ColumnClustered", sheet.getRange("AC3:AF7"), "Auto");
  chart.title.text = `${month} Weekly Income, Expenses, and Allocations`;
  chart.setPosition(sheet.getRange("AC9:AI24"));
  chart.width = 560;
  chart.height = 280;
  chart.yAxis = { numberFormatCode: currencyFormat };

  sectionHeader(sheet, "AC27:AE27", "3-Month Trend Data", palette.tealDark);
  sheet.getRange("AC28:AE28").values = [["Month", "Income", "Expenses"]];
  const trendMonthIndexes = [monthIndex - 2, monthIndex - 1, monthIndex];
  sheet.getRange("AC29:AC31").values = trendMonthIndexes.map((idx) => [idx >= 0 ? monthLabels[idx] : ""]);
  sheet.getRange("AD29:AE31").formulas = trendMonthIndexes.map((idx) => (
    idx >= 0
      ? [`='Annual Overview'!$B$${18 + idx}`, `='Annual Overview'!$C$${18 + idx}`]
      : ['=""', '=""']
  ));
  styleHeader(sheet.getRange("AC28:AE28"), palette.sea);
  styleBox(sheet.getRange("AC29:AE31"), palette.white);
  money(sheet.getRange("AD29:AE31"));

  const trendChart = sheet.charts.add("ColumnClustered", sheet.getRange("AC28:AE31"), "Auto");
  trendChart.title.text = `${month} 3-Month Income vs Expenses`;
  trendChart.setPosition(sheet.getRange("AC33:AI48"));
  trendChart.width = 560;
  trendChart.height = 280;
  trendChart.yAxis = { numberFormatCode: currencyFormat };
}

try {
  buildStartHere();
  console.log("BUILT Start Here");
  buildOverview();
  console.log("BUILT Annual Overview");
  for (const month of months) {
    buildMonthSheet(month, monthSheets.get(month));
    console.log(`BUILT ${month}`);
  }
} catch (error) {
  console.error(`BUILD_FAILED ${error?.message ?? error}`);
  throw error;
}

let inspectAnnual;
let inspectJanuary;
let errors;
try {
  inspectAnnual = await workbook.inspect({
    kind: "table",
    range: "Annual Overview!A1:T60",
    include: "values,formulas",
    tableMaxRows: 65,
    tableMaxCols: 20,
  });
  console.log("ANNUAL_OVERVIEW_INSPECT");
  console.log(inspectAnnual.ndjson);

  const inspectStartHere = await workbook.inspect({
    kind: "table",
    range: "Start Here!B2:I28",
    include: "values,formulas",
    tableMaxRows: 30,
    tableMaxCols: 8,
  });
  console.log("START_HERE_INSPECT");
  console.log(inspectStartHere.ndjson);

  inspectJanuary = await workbook.inspect({
    kind: "table",
    range: "January!A57:E80",
    include: "values,formulas",
    tableMaxRows: 24,
    tableMaxCols: 8,
  });
  console.log("JANUARY_SUMMARY_INSPECT");
  console.log(inspectJanuary.ndjson);

  errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
  });
  console.log("FORMULA_ERROR_SCAN");
  console.log(errors.ndjson);
} catch (error) {
  console.error(`INSPECT_FAILED ${error?.message ?? error}`);
  throw error;
}

try {
  const previewDir = path.join(outputDir, "previews");
  await fs.mkdir(previewDir, { recursive: true });
  const previewTargets = ["Start Here", "Annual Overview", ...months];
  for (const sheetName of previewTargets) {
    const range = sheetName === "Start Here" ? "A1:I29" : sheetName === "Annual Overview" ? "A1:T60" : "A1:AI80";
    const preview = await workbook.render({ sheetName, range, scale: 1.25 });
    await fs.writeFile(path.join(previewDir, `${sheetName.replaceAll(" ", "_")}.png`), Buffer.from(await preview.arrayBuffer()));
    console.log(`RENDERED ${sheetName}`);
  }
} catch (error) {
  console.error(`RENDER_FAILED ${error?.message ?? error}`);
  throw error;
}

try {
  await fs.mkdir(outputDir, { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  console.log(`EXPORTED ${outputPath}`);
} catch (error) {
  console.error(`EXPORT_FAILED ${error?.message ?? error}`);
  throw error;
}

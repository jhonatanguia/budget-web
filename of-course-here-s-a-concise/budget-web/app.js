const STORAGE_KEY = "household-budget-os-v1";

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

const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const allocationBuckets = [
  { id: "tfsaHouse", label: "TFSA / House", defaultTarget: 50 },
  { id: "emergency", label: "Emergency fund", defaultTarget: null },
  { id: "investments", label: "Other investments", defaultTarget: null },
  { id: "debt", label: "Debt payoff", defaultTarget: null },
  { id: "householdReserve", label: "Household reserve", defaultTarget: null },
  { id: "travelGifts", label: "Travel/Gifts fund", defaultTarget: null },
];

const categoryTemplates = [
  { id: "cibc", label: "CIBC", group: "Credit cards" },
  { id: "rbc", label: "RBC", group: "Credit cards" },
  { id: "amex", label: "AMEX", group: "Credit cards" },
  { id: "wealthsimple", label: "Wealthsimple", group: "Credit cards" },
  { id: "triangle", label: "Triangle", group: "Credit cards" },
  { id: "otherCards", label: "Other Credit Cards", group: "Credit cards" },
  { id: "food", label: "Food", group: "Household" },
  { id: "cars", label: "Cars", group: "Household" },
  { id: "house", label: "House", group: "Household" },
  { id: "travel", label: "Travel", group: "Household" },
  { id: "giftEducation", label: "Gift/Education", group: "Household" },
  { id: "custom1", label: "Custom 1", group: "Flexible", custom: true },
  { id: "custom2", label: "Custom 2", group: "Flexible", custom: true },
  { id: "custom3", label: "Custom 3", group: "Flexible", custom: true },
  { id: "business", label: "Business", group: "Household" },
  { id: "subscriptions", label: "Monthly Subscription / Utilities", group: "Household" },
  { id: "others", label: "Others", group: "Household" },
];

const guideSections = [
  {
    title: "Quick start",
    body: `
      <ul>
        <li>Choose the month at the top of the page.</li>
        <li>Go to Weekly Entries, pick a week, and add income or expenses.</li>
        <li>Use Settings to change allocation percentages, allocation mode, and custom category names.</li>
        <li>Return to Dashboard to review cash flow, trends, and top spending categories.</li>
      </ul>
    `,
  },
  {
    title: "How allocation mode works",
    body: `
      <p><strong>Month-to-date smoothing</strong> treats the month like one running balance. If week 1 is positive and week 2 is negative, the negative week reduces the allocation already assigned.</p>
      <p><strong>Independent weekly</strong> lets each week stand alone. Negative weeks do not remove previous allocations.</p>
    `,
  },
  {
    title: "How percentages balance automatically",
    body: `
      <p>Type any percentage from 0 to 100. If you leave buckets blank, the remaining percentage is shared evenly across the blank buckets. If the typed percentages go over 100%, the app normalizes them back down to 100%.</p>
    `,
  },
  {
    title: "How to use custom categories",
    body: `
      <p>Custom 1, Custom 2, and Custom 3 can become anything you want: Gym, Medical, Pets, Events, Kids, or one-time purchases. Rename them in Settings and the category dropdown updates everywhere.</p>
    `,
  },
  {
    title: "Using it on phone and computer",
    body: `
      <p>The app is responsive, so it works on a phone or desktop. Data is saved privately in the browser on each device. Use Export backup and Import backup to move a copy between devices.</p>
    `,
  },
];

const els = {
  menuButton: document.querySelector("#menuButton"),
  sidebar: document.querySelector("#sidebar"),
  globalMonth: document.querySelector("#globalMonth"),
  kpiGrid: document.querySelector("#kpiGrid"),
  monthlyFocusTitle: document.querySelector("#monthlyFocusTitle"),
  monthlyFocusText: document.querySelector("#monthlyFocusText"),
  focusList: document.querySelector("#focusList"),
  categoryRank: document.querySelector("#categoryRank"),
  annualChart: document.querySelector("#annualChart"),
  weeklyChart: document.querySelector("#weeklyChart"),
  categoryChart: document.querySelector("#categoryChart"),
  weekSwitch: document.querySelector("#weekSwitch"),
  entryForm: document.querySelector("#entryForm"),
  entryDate: document.querySelector("#entryDate"),
  entryType: document.querySelector("#entryType"),
  entryCategory: document.querySelector("#entryCategory"),
  entryAmount: document.querySelector("#entryAmount"),
  entryNote: document.querySelector("#entryNote"),
  entrySummary: document.querySelector("#entrySummary"),
  ledgerTitle: document.querySelector("#ledgerTitle"),
  ledgerBody: document.querySelector("#ledgerBody"),
  clearWeekButton: document.querySelector("#clearWeekButton"),
  allocationMode: document.querySelector("#allocationMode"),
  allocationList: document.querySelector("#allocationList"),
  effectiveBox: document.querySelector("#effectiveBox"),
  customCategoryFields: document.querySelector("#customCategoryFields"),
  guideAccordion: document.querySelector("#guideAccordion"),
  exportButton: document.querySelector("#exportButton"),
  settingsExportButton: document.querySelector("#settingsExportButton"),
  importFile: document.querySelector("#importFile"),
  settingsImportFile: document.querySelector("#settingsImportFile"),
  resetButton: document.querySelector("#resetButton"),
  toast: document.querySelector("#toast"),
};

let state = normalizeState(loadState());
let toastTimer = null;

function defaultState() {
  const targets = {};
  allocationBuckets.forEach((bucket) => {
    targets[bucket.id] = bucket.defaultTarget;
  });

  return {
    selectedMonth: new Date().getMonth(),
    selectedWeek: 1,
    allocationMode: "mtd",
    allocationTargets: targets,
    customLabels: {
      custom1: "Custom 1",
      custom2: "Custom 2",
      custom3: "Custom 3",
    },
    entries: [],
  };
}

function normalizeState(raw) {
  const base = defaultState();
  const incoming = raw && typeof raw === "object" ? raw : {};
  return {
    ...base,
    ...incoming,
    selectedMonth: Number.isInteger(incoming.selectedMonth) ? clamp(incoming.selectedMonth, 0, 11) : base.selectedMonth,
    selectedWeek: Number.isInteger(incoming.selectedWeek) ? clamp(incoming.selectedWeek, 1, 4) : base.selectedWeek,
    allocationMode: incoming.allocationMode === "independent" ? "independent" : "mtd",
    allocationTargets: { ...base.allocationTargets, ...(incoming.allocationTargets || {}) },
    customLabels: { ...base.customLabels, ...(incoming.customLabels || {}) },
    entries: Array.isArray(incoming.entries) ? incoming.entries.map(normalizeEntry).filter(Boolean) : [],
  };
}

function normalizeEntry(entry) {
  if (!entry || typeof entry !== "object") return null;
  const type = entry.type === "income" ? "income" : "expense";
  const categoryIds = categoryTemplates.map((category) => category.id);
  const category = categoryIds.includes(entry.category) ? entry.category : "others";
  const amount = Number(entry.amount);
  if (!Number.isFinite(amount) || amount <= 0) return null;
  return {
    id: entry.id || createId(),
    month: clamp(Number(entry.month) || 0, 0, 11),
    week: clamp(Number(entry.week) || 1, 1, 4),
    date: typeof entry.date === "string" ? entry.date : "",
    type,
    category,
    amount,
    note: typeof entry.note === "string" ? entry.note : "",
  };
}

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    return saved ? JSON.parse(saved) : null;
  } catch {
    return null;
  }
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    showToast("Could not save in this browser. Export a backup before closing.");
  }
}

function createId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `entry-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-CA", {
    style: "currency",
    currency: "CAD",
    maximumFractionDigits: Math.abs(value || 0) >= 1000 ? 0 : 2,
  }).format(Number(value) || 0);
}

function formatPercent(value) {
  return `${((Number(value) || 0) * 100).toFixed(1)}%`;
}

function currentCategories() {
  return categoryTemplates.map((category) => ({
    ...category,
    label: category.custom ? state.customLabels[category.id] || category.label : category.label,
  }));
}

function categoryLabel(categoryId) {
  const category = currentCategories().find((item) => item.id === categoryId);
  return category ? category.label : "Other";
}

function categoryGroup(categoryId) {
  const category = currentCategories().find((item) => item.id === categoryId);
  return category ? category.group : "Household";
}

function calculateEffectiveAllocations() {
  const targetValues = allocationBuckets.map((bucket) => {
    const value = state.allocationTargets[bucket.id];
    return value === null || value === "" || Number.isNaN(Number(value)) ? null : clamp(Number(value), 0, 100) / 100;
  });
  const enteredSum = targetValues.reduce((sum, value) => sum + (value === null ? 0 : value), 0);
  const blankCount = targetValues.filter((value) => value === null).length;

  const effective = {};
  allocationBuckets.forEach((bucket, index) => {
    const value = targetValues[index];
    if (enteredSum > 1) {
      effective[bucket.id] = value === null ? 0 : value / enteredSum;
      return;
    }
    if (value !== null) {
      effective[bucket.id] = value;
      return;
    }
    effective[bucket.id] = blankCount > 0 ? Math.max(1 - enteredSum, 0) / blankCount : 0;
  });

  return effective;
}

function calculateMonth(monthIndex) {
  const categories = currentCategories();
  const effective = calculateEffectiveAllocations();
  const weekly = Array.from({ length: 4 }, (_, index) => ({
    week: index + 1,
    income: 0,
    expenses: 0,
    netBefore: 0,
    allocations: {},
    allocationTotal: 0,
    netAfter: 0,
  }));
  const categoryTotals = Object.fromEntries(categories.map((category) => [category.id, 0]));

  state.entries
    .filter((entry) => entry.month === monthIndex)
    .forEach((entry) => {
      const week = weekly[entry.week - 1];
      if (entry.type === "income") {
        week.income += entry.amount;
      } else {
        week.expenses += entry.amount;
        categoryTotals[entry.category] = (categoryTotals[entry.category] || 0) + entry.amount;
      }
    });

  weekly.forEach((week) => {
    week.netBefore = week.income - week.expenses;
  });

  const allocationTotals = Object.fromEntries(allocationBuckets.map((bucket) => [bucket.id, 0]));
  const priorAllocations = Object.fromEntries(allocationBuckets.map((bucket) => [bucket.id, 0]));
  let cumulativeNetBefore = 0;

  weekly.forEach((week) => {
    cumulativeNetBefore += week.netBefore;
    allocationBuckets.forEach((bucket) => {
      let allocation = 0;
      if (state.allocationMode === "independent") {
        allocation = Math.max(week.netBefore, 0) * effective[bucket.id];
      } else {
        const cumulativeTarget = Math.max(cumulativeNetBefore, 0) * effective[bucket.id];
        allocation = cumulativeTarget - priorAllocations[bucket.id];
      }
      week.allocations[bucket.id] = allocation;
      week.allocationTotal += allocation;
      priorAllocations[bucket.id] += allocation;
      allocationTotals[bucket.id] += allocation;
    });
    week.netAfter = week.netBefore - week.allocationTotal;
  });

  const income = weekly.reduce((sum, week) => sum + week.income, 0);
  const expenses = weekly.reduce((sum, week) => sum + week.expenses, 0);
  const netBefore = income - expenses;
  const allocationTotal = Object.values(allocationTotals).reduce((sum, value) => sum + value, 0);
  const netAfter = netBefore - allocationTotal;
  const topCategory = topCategories(categoryTotals, 1)[0] || null;

  return {
    monthIndex,
    income,
    expenses,
    netBefore,
    netAfter,
    weekly,
    categoryTotals,
    allocationTotals,
    allocationTotal,
    topCategory,
  };
}

function calculateAnnual() {
  const monthly = months.map((_, index) => calculateMonth(index));
  const categories = currentCategories();
  const categoryTotals = Object.fromEntries(categories.map((category) => [category.id, 0]));
  const allocationTotals = Object.fromEntries(allocationBuckets.map((bucket) => [bucket.id, 0]));

  monthly.forEach((month) => {
    Object.entries(month.categoryTotals).forEach(([categoryId, value]) => {
      categoryTotals[categoryId] = (categoryTotals[categoryId] || 0) + value;
    });
    Object.entries(month.allocationTotals).forEach(([bucketId, value]) => {
      allocationTotals[bucketId] = (allocationTotals[bucketId] || 0) + value;
    });
  });

  const income = monthly.reduce((sum, month) => sum + month.income, 0);
  const expenses = monthly.reduce((sum, month) => sum + month.expenses, 0);
  const latestDataIndex = findLatestDataMonth(monthly);
  const selectedStats = monthly[state.selectedMonth];
  const selectedTrend = calculateTrendForMonth(monthly, state.selectedMonth);
  const latestTrend = latestDataIndex >= 0 ? calculateTrendForMonth(monthly, latestDataIndex) : selectedTrend;

  return {
    monthly,
    income,
    expenses,
    netBefore: income - expenses,
    categoryTotals,
    allocationTotals,
    topCategory: topCategories(categoryTotals, 1)[0] || null,
    latestDataIndex,
    selectedStats,
    selectedTrend,
    latestTrend,
  };
}

function calculateTrendForMonth(monthly, index) {
  const start = Math.max(0, index - 2);
  const monthsInAverage = monthly.slice(start, index + 1);
  const avgIncome = average(monthsInAverage.map((month) => month.income));
  const avgExpenses = average(monthsInAverage.map((month) => month.expenses));
  const previousExpenses = index > 0 ? monthly[index - 1].expenses : 0;
  const expenseChange = monthly[index].expenses - previousExpenses;
  return {
    avgIncome,
    avgExpenses,
    expenseChange,
    direction: expenseChange < 0 ? "Reducing" : expenseChange > 0 ? "Increasing" : "Flat / no change",
  };
}

function average(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

function findLatestDataMonth(monthly) {
  for (let index = monthly.length - 1; index >= 0; index -= 1) {
    if (monthly[index].income || monthly[index].expenses) return index;
  }
  return -1;
}

function topCategories(categoryTotals, limit = 5) {
  return Object.entries(categoryTotals)
    .map(([id, value]) => ({ id, label: categoryLabel(id), group: categoryGroup(id), value }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, limit);
}

function renderMonthOptions() {
  els.globalMonth.innerHTML = months
    .map((month, index) => `<option value="${index}">${month}</option>`)
    .join("");
  els.globalMonth.value = String(state.selectedMonth);
}

function renderWeekSwitch() {
  els.weekSwitch.innerHTML = [1, 2, 3, 4]
    .map((week) => `<button type="button" class="${week === state.selectedWeek ? "active" : ""}" data-week="${week}">Week ${week}</button>`)
    .join("");
}

function renderCategoryDropdown() {
  const selected = els.entryCategory.value;
  const groups = {};
  currentCategories().forEach((category) => {
    groups[category.group] ||= [];
    groups[category.group].push(category);
  });

  els.entryCategory.innerHTML = Object.entries(groups)
    .map(([group, categories]) => {
      const options = categories.map((category) => `<option value="${category.id}">${escapeHtml(category.label)}</option>`).join("");
      return `<optgroup label="${escapeHtml(group)}">${options}</optgroup>`;
    })
    .join("");

  const categoryIds = currentCategories().map((category) => category.id);
  els.entryCategory.value = categoryIds.includes(selected) ? selected : "food";
  els.entryCategory.disabled = els.entryType.value === "income";
}

function renderDashboard() {
  const annual = calculateAnnual();
  const month = annual.selectedStats;
  const trend = annual.selectedTrend;
  const topCategory = month.topCategory;
  const directionClass = trend.expenseChange < 0 ? "positive" : trend.expenseChange > 0 ? "negative" : "";
  const topLabel = topCategory ? topCategory.label : "No spending yet";

  const kpis = [
    { label: "Month income", value: formatMoney(month.income), sub: months[state.selectedMonth] },
    { label: "Month expenses", value: formatMoney(month.expenses), sub: topLabel, tone: month.expenses > 0 ? "negative" : "" },
    { label: "Net before allocations", value: formatMoney(month.netBefore), sub: "Income minus expenses", tone: month.netBefore >= 0 ? "positive" : "negative" },
    { label: "Allocated", value: formatMoney(month.allocationTotal), sub: state.allocationMode === "mtd" ? "Smoothed month-to-date" : "Independent weekly" },
    { label: "Net after allocations", value: formatMoney(month.netAfter), sub: "Remaining cash flow", tone: month.netAfter >= 0 ? "positive" : "negative" },
    { label: "3-month avg expenses", value: formatMoney(trend.avgExpenses), sub: trend.direction, tone: directionClass },
  ];

  els.kpiGrid.innerHTML = kpis
    .map((kpi) => `
      <article class="kpi-card">
        <span class="kpi-label">${kpi.label}</span>
        <strong class="kpi-value ${kpi.tone || ""}">${kpi.value}</strong>
        <span class="kpi-sub">${kpi.sub}</span>
      </article>
    `)
    .join("");

  els.monthlyFocusTitle.textContent = topCategory ? `${topCategory.label} is the biggest ${months[state.selectedMonth]} expense` : "No spending yet";
  els.monthlyFocusText.textContent = topCategory
    ? `${formatMoney(topCategory.value)} went to ${topCategory.label}. Expense direction is ${trend.direction.toLowerCase()} compared with the prior month.`
    : "Add weekly expenses to see the category that deserves attention.";

  const annualTop = annual.topCategory ? `${annual.topCategory.label} (${formatMoney(annual.topCategory.value)})` : "No spending yet";
  els.focusList.innerHTML = [
    ["3-month avg income", formatMoney(trend.avgIncome)],
    ["3-month avg expenses", formatMoney(trend.avgExpenses)],
    ["Expense change", `${formatMoney(trend.expenseChange)} ${trend.direction}`],
    ["Biggest annual category", annualTop],
  ]
    .map(([label, value]) => `<div class="focus-pill"><strong>${label}</strong><span>${value}</span></div>`)
    .join("");

  const categoryRanks = topCategories(month.categoryTotals, 8);
  els.categoryRank.innerHTML = categoryRanks.length
    ? categoryRanks
        .map((item) => `<div class="rank-row"><strong>${escapeHtml(item.label)}</strong><span>${formatMoney(item.value)}</span></div>`)
        .join("")
    : `<div class="empty-state">No expense categories yet.</div>`;

  drawAnnualChart(annual);
  drawWeeklyChart(month);
  drawCategoryChart(month);
}

function renderEntrySummary() {
  const month = calculateMonth(state.selectedMonth);
  const week = month.weekly[state.selectedWeek - 1];
  const items = [
    ["Income", formatMoney(week.income)],
    ["Expenses", formatMoney(week.expenses)],
    ["Net before", formatMoney(week.netBefore)],
    ["Allocations", formatMoney(week.allocationTotal)],
    ["Net after", formatMoney(week.netAfter)],
  ];
  els.entrySummary.innerHTML = items
    .map(([label, value]) => `<div class="summary-item"><span>${label}</span><strong>${value}</strong></div>`)
    .join("");
}

function renderLedger() {
  const entries = state.entries
    .filter((entry) => entry.month === state.selectedMonth && entry.week === state.selectedWeek)
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  els.ledgerTitle.textContent = `${months[state.selectedMonth]} Week ${state.selectedWeek} entries`;
  if (!entries.length) {
    els.ledgerBody.innerHTML = `<tr><td colspan="6"><div class="empty-state">No entries in this week yet.</div></td></tr>`;
    return;
  }

  els.ledgerBody.innerHTML = entries
    .map((entry) => `
      <tr>
        <td>${entry.date || "Unscheduled"}</td>
        <td>${entry.type === "income" ? "Income" : "Expense"}</td>
        <td>${entry.type === "income" ? "Income" : escapeHtml(categoryLabel(entry.category))}</td>
        <td>${escapeHtml(entry.note || "")}</td>
        <td class="amount-cell ${entry.type === "income" ? "positive" : "negative"}">${formatMoney(entry.amount)}</td>
        <td><button class="icon-delete" type="button" title="Delete entry" data-delete="${entry.id}">x</button></td>
      </tr>
    `)
    .join("");
}

function renderSettings() {
  els.allocationMode.value = state.allocationMode;
  const effective = calculateEffectiveAllocations();
  els.allocationList.innerHTML = allocationBuckets
    .map((bucket) => {
      const rawValue = state.allocationTargets[bucket.id];
      const value = rawValue === null || rawValue === "" ? "" : Number(rawValue);
      return `
        <label class="allocation-row">
          <span>${bucket.label}</span>
          <input type="number" min="0" max="100" step="1" value="${value}" data-allocation="${bucket.id}" placeholder="Auto" />
        </label>
      `;
    })
    .join("");

  const effectiveRows = allocationBuckets
    .map((bucket) => `<div class="effective-row"><strong>${bucket.label}</strong><span>${formatPercent(effective[bucket.id])}</span></div>`)
    .join("");
  const total = Object.values(effective).reduce((sum, value) => sum + value, 0);
  els.effectiveBox.innerHTML = `
    <div class="focus-list">
      ${effectiveRows}
      <div class="effective-row"><strong>Effective total</strong><span>${formatPercent(total)}</span></div>
    </div>
  `;

  els.customCategoryFields.innerHTML = ["custom1", "custom2", "custom3"]
    .map((id, index) => `
      <div class="field">
        <label for="${id}">Custom ${index + 1}</label>
        <input id="${id}" type="text" value="${escapeHtml(state.customLabels[id] || `Custom ${index + 1}`)}" data-custom="${id}" />
      </div>
    `)
    .join("");
}

function renderGuide() {
  els.guideAccordion.innerHTML = guideSections
    .map((section, index) => `
      <article class="accordion-item ${index === 0 ? "open" : ""}">
        <button class="accordion-toggle" type="button">
          ${section.title}
          <span>${index === 0 ? "-" : "+"}</span>
        </button>
        <div class="accordion-content">${section.body}</div>
      </article>
    `)
    .join("");
}

function renderAll() {
  saveState();
  renderMonthOptions();
  renderWeekSwitch();
  renderCategoryDropdown();
  renderDashboard();
  renderEntrySummary();
  renderLedger();
  renderSettings();
}

function drawAnnualChart(annual) {
  const canvas = els.annualChart;
  const { ctx, width, height } = prepareCanvas(canvas);
  const incomes = annual.monthly.map((month) => month.income);
  const expenses = annual.monthly.map((month) => month.expenses);
  const avgExpenses = annual.monthly.map((_, index) => calculateTrendForMonth(annual.monthly, index).avgExpenses);
  const maxValue = Math.max(...incomes, ...expenses, ...avgExpenses, 1);
  clearCanvas(ctx, width, height);

  if (maxValue <= 1) {
    drawNoData(ctx, width, height, "Add entries to build the annual trend.");
    return;
  }

  const pad = { left: 44, right: 18, top: 18, bottom: 48 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  drawGrid(ctx, pad, chartW, chartH, maxValue);

  const band = chartW / 12;
  annual.monthly.forEach((month, index) => {
    const x = pad.left + index * band + band * 0.18;
    const incomeH = (month.income / maxValue) * chartH;
    const expenseH = (month.expenses / maxValue) * chartH;
    ctx.fillStyle = "#176b87";
    ctx.fillRect(x, pad.top + chartH - incomeH, Math.max(band * 0.24, 4), incomeH);
    ctx.fillStyle = "#f9735b";
    ctx.fillRect(x + band * 0.28, pad.top + chartH - expenseH, Math.max(band * 0.24, 4), expenseH);
    ctx.fillStyle = "#60717a";
    ctx.font = "11px system-ui";
    ctx.textAlign = "center";
    ctx.fillText(monthLabels[index], pad.left + index * band + band * 0.5, height - 18);
  });

  drawLine(ctx, avgExpenses, maxValue, pad, chartW, chartH, "#0f766e");
  drawLegend(ctx, [
    ["Income", "#176b87"],
    ["Expenses", "#f9735b"],
    ["3-mo avg expenses", "#0f766e"],
  ], pad.left, height - 8);
}

function drawWeeklyChart(month) {
  const canvas = els.weeklyChart;
  const { ctx, width, height } = prepareCanvas(canvas);
  const maxValue = Math.max(...month.weekly.flatMap((week) => [week.income, week.expenses, Math.abs(week.allocationTotal)]), 1);
  clearCanvas(ctx, width, height);

  if (maxValue <= 1) {
    drawNoData(ctx, width, height, "Add this month’s entries to see weekly movement.");
    return;
  }

  const pad = { left: 44, right: 18, top: 18, bottom: 48 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  drawGrid(ctx, pad, chartW, chartH, maxValue);

  const band = chartW / 4;
  month.weekly.forEach((week, index) => {
    const x = pad.left + index * band + band * 0.18;
    const values = [
      [week.income, "#176b87"],
      [week.expenses, "#f9735b"],
      [Math.abs(week.allocationTotal), "#15803d"],
    ];
    values.forEach(([value, color], valueIndex) => {
      const barH = (value / maxValue) * chartH;
      ctx.fillStyle = color;
      ctx.fillRect(x + valueIndex * band * 0.2, pad.top + chartH - barH, Math.max(band * 0.16, 8), barH);
    });
    ctx.fillStyle = "#60717a";
    ctx.font = "11px system-ui";
    ctx.textAlign = "center";
    ctx.fillText(`Week ${week.week}`, pad.left + index * band + band * 0.5, height - 18);
  });

  drawLegend(ctx, [
    ["Income", "#176b87"],
    ["Expenses", "#f9735b"],
    ["Allocations", "#15803d"],
  ], pad.left, height - 8);
}

function drawCategoryChart(month) {
  const canvas = els.categoryChart;
  const { ctx, width, height } = prepareCanvas(canvas);
  const categories = topCategories(month.categoryTotals, 6);
  clearCanvas(ctx, width, height);

  if (!categories.length) {
    drawNoData(ctx, width, height, "Add expenses to see spending mix.");
    return;
  }

  const total = categories.reduce((sum, item) => sum + item.value, 0);
  const colors = ["#0f766e", "#f9735b", "#176b87", "#d6a525", "#7c3aed", "#15803d"];
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(width, height) * 0.34;
  let startAngle = -Math.PI / 2;

  categories.forEach((item, index) => {
    const slice = (item.value / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, startAngle, startAngle + slice);
    ctx.closePath();
    ctx.fillStyle = colors[index % colors.length];
    ctx.fill();
    startAngle += slice;
  });

  ctx.beginPath();
  ctx.arc(cx, cy, radius * 0.55, 0, Math.PI * 2);
  ctx.fillStyle = "#ffffff";
  ctx.fill();
  ctx.fillStyle = "#172a31";
  ctx.font = "700 14px system-ui";
  ctx.textAlign = "center";
  ctx.fillText("Expenses", cx, cy - 4);
  ctx.font = "800 18px system-ui";
  ctx.fillText(formatMoney(total), cx, cy + 20);
}

function prepareCanvas(canvas) {
  const rect = canvas.getBoundingClientRect();
  const width = Math.max(Math.floor(rect.width || 480), 280);
  const height = Number(canvas.getAttribute("height")) || 280;
  const ratio = window.devicePixelRatio || 1;
  canvas.width = width * ratio;
  canvas.height = height * ratio;
  const ctx = canvas.getContext("2d");
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
  return { ctx, width, height };
}

function clearCanvas(ctx, width, height) {
  ctx.clearRect(0, 0, width, height);
}

function drawNoData(ctx, width, height, text) {
  ctx.fillStyle = "#edf9f7";
  ctx.fillRect(0, 0, width, height);
  ctx.fillStyle = "#60717a";
  ctx.font = "700 14px system-ui";
  ctx.textAlign = "center";
  ctx.fillText(text, width / 2, height / 2);
}

function drawGrid(ctx, pad, chartW, chartH, maxValue) {
  ctx.strokeStyle = "#dbe4e1";
  ctx.lineWidth = 1;
  ctx.fillStyle = "#60717a";
  ctx.font = "11px system-ui";
  ctx.textAlign = "right";

  for (let step = 0; step <= 4; step += 1) {
    const y = pad.top + chartH - (chartH * step) / 4;
    ctx.beginPath();
    ctx.moveTo(pad.left, y);
    ctx.lineTo(pad.left + chartW, y);
    ctx.stroke();
    ctx.fillText(formatShortMoney((maxValue * step) / 4), pad.left - 8, y + 4);
  }
}

function drawLine(ctx, values, maxValue, pad, chartW, chartH, color) {
  ctx.beginPath();
  values.forEach((value, index) => {
    const x = pad.left + (index + 0.5) * (chartW / values.length);
    const y = pad.top + chartH - (value / maxValue) * chartH;
    if (index === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.strokeStyle = color;
  ctx.lineWidth = 3;
  ctx.stroke();
}

function drawLegend(ctx, items, x, y) {
  let cursor = x;
  ctx.font = "11px system-ui";
  ctx.textAlign = "left";
  items.forEach(([label, color]) => {
    ctx.fillStyle = color;
    ctx.fillRect(cursor, y - 9, 10, 10);
    ctx.fillStyle = "#60717a";
    ctx.fillText(label, cursor + 14, y);
    cursor += ctx.measureText(label).width + 38;
  });
}

function formatShortMoney(value) {
  const abs = Math.abs(value);
  if (abs >= 1000000) return `$${(value / 1000000).toFixed(1)}m`;
  if (abs >= 1000) return `$${(value / 1000).toFixed(0)}k`;
  return `$${Math.round(value)}`;
}

function bindEvents() {
  els.menuButton.addEventListener("click", () => {
    els.sidebar.classList.toggle("open");
  });

  document.querySelectorAll(".nav-button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".nav-button").forEach((item) => item.classList.remove("active"));
      document.querySelectorAll(".view").forEach((view) => view.classList.remove("active"));
      button.classList.add("active");
      document.querySelector(`#${button.dataset.view}View`).classList.add("active");
      els.sidebar.classList.remove("open");
      window.setTimeout(() => renderDashboard(), 50);
    });
  });

  els.globalMonth.addEventListener("change", () => {
    state.selectedMonth = Number(els.globalMonth.value);
    renderAll();
  });

  els.weekSwitch.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-week]");
    if (!button) return;
    state.selectedWeek = Number(button.dataset.week);
    renderAll();
  });

  els.entryType.addEventListener("change", renderCategoryDropdown);

  els.entryForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const amount = Number(els.entryAmount.value);
    if (!Number.isFinite(amount) || amount <= 0) {
      showToast("Enter an amount greater than zero.");
      return;
    }

    state.entries.push({
      id: createId(),
      month: state.selectedMonth,
      week: state.selectedWeek,
      date: els.entryDate.value,
      type: els.entryType.value,
      category: els.entryType.value === "income" ? "income" : els.entryCategory.value,
      amount,
      note: els.entryNote.value.trim(),
    });

    els.entryAmount.value = "";
    els.entryNote.value = "";
    saveState();
    renderAll();
    showToast("Entry added.");
  });

  els.ledgerBody.addEventListener("click", (event) => {
    const button = event.target.closest("[data-delete]");
    if (!button) return;
    state.entries = state.entries.filter((entry) => entry.id !== button.dataset.delete);
    renderAll();
    showToast("Entry removed.");
  });

  els.clearWeekButton.addEventListener("click", () => {
    const hasEntries = state.entries.some((entry) => entry.month === state.selectedMonth && entry.week === state.selectedWeek);
    if (!hasEntries) {
      showToast("This week is already empty.");
      return;
    }
    if (!window.confirm(`Clear ${months[state.selectedMonth]} Week ${state.selectedWeek}?`)) return;
    state.entries = state.entries.filter((entry) => entry.month !== state.selectedMonth || entry.week !== state.selectedWeek);
    renderAll();
    showToast("Week cleared.");
  });

  els.allocationMode.addEventListener("change", () => {
    state.allocationMode = els.allocationMode.value;
    renderAll();
  });

  els.allocationList.addEventListener("change", (event) => {
    const input = event.target.closest("[data-allocation]");
    if (!input) return;
    const value = input.value.trim();
    state.allocationTargets[input.dataset.allocation] = value === "" ? null : clamp(Number(value), 0, 100);
    renderAll();
  });

  els.customCategoryFields.addEventListener("input", (event) => {
    const input = event.target.closest("[data-custom]");
    if (!input) return;
    const fallback = input.dataset.custom.replace("custom", "Custom ");
    state.customLabels[input.dataset.custom] = input.value.trim() || fallback;
    saveState();
    renderCategoryDropdown();
    renderDashboard();
    renderLedger();
  });

  els.guideAccordion.addEventListener("click", (event) => {
    const toggle = event.target.closest(".accordion-toggle");
    if (!toggle) return;
    const item = toggle.closest(".accordion-item");
    item.classList.toggle("open");
    toggle.querySelector("span").textContent = item.classList.contains("open") ? "-" : "+";
  });

  [els.exportButton, els.settingsExportButton].forEach((button) => {
    button.addEventListener("click", exportBackup);
  });

  [els.importFile, els.settingsImportFile].forEach((input) => {
    input.addEventListener("change", importBackup);
  });

  els.resetButton.addEventListener("click", () => {
    if (!window.confirm("Reset all entries and settings in this browser?")) return;
    state = defaultState();
    renderAll();
    showToast("Budget reset.");
  });

  window.addEventListener("resize", debounce(() => renderDashboard(), 150));
}

function exportBackup() {
  const payload = {
    exportedAt: new Date().toISOString(),
    version: 1,
    app: "Household Budget OS",
    data: state,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `household-budget-backup-${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  showToast("Backup exported.");
}

function importBackup(event) {
  const file = event.target.files[0];
  event.target.value = "";
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result));
      state = normalizeState(parsed.data || parsed);
      renderAll();
      showToast("Backup imported.");
    } catch {
      showToast("That backup file could not be imported.");
    }
  };
  reader.readAsText(file);
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("visible");
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => {
    els.toast.classList.remove("visible");
  }, 2600);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function debounce(fn, delay) {
  let timer = null;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), delay);
  };
}

function init() {
  renderMonthOptions();
  renderWeekSwitch();
  renderGuide();
  bindEvents();
  renderAll();
}

init();

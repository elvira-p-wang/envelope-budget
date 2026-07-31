/* Envelope — reads Budget and Weekly Log sheets from a linked workbook via SheetJS,
   entirely in the browser, and renders overview cards, insights, charts and a
   category breakdown. Workbook data is processed locally and is never uploaded. */

const WORKBOOK_PATH = "./sample-budget.xlsx";
const TEMPLATE_PATH = "./blank-template.xlsx";

// Used only if the workbook can't be loaded (e.g. opening index.html directly
// without a local server, or offline). Purely illustrative numbers.
const FALLBACK_BUDGET = [
  ["Housing", 350], ["Groceries", 90], ["Dining Out", 60], ["Transport", 40],
  ["Utilities", 45], ["Wellness", 30], ["Entertainment", 35], ["Shopping", 40],
  ["Learning", 15], ["Savings", 100],
];

// A fixed, muted colour per category so the donut, bar chart and category
// list all use the same colour for the same category. Any category not
// listed here falls back to the accent colour, so custom workbooks still work.
const CATEGORY_COLORS = {
  "housing": "#8a6a52",
  "groceries": "#7f9b6b",
  "dining out": "#c98f6b",
  "transport": "#6b8fa3",
  "utilities": "#af7c3f",
  "wellness": "#9b7fb0",
  "entertainment": "#5f7d8c",
  "shopping": "#a35d52",
  "learning": "#b0a48f",
  "savings": "#55806b",
};
const FALLBACK_COLORS = ["#8a6a52", "#7f9b6b", "#c98f6b", "#6b8fa3", "#af7c3f", "#9b7fb0", "#5f7d8c", "#a35d52", "#b0a48f", "#55806b"];
const colorFor = (category, index) => CATEGORY_COLORS[String(category).toLowerCase()] || FALLBACK_COLORS[index % FALLBACK_COLORS.length];

const cssVar = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

const money = (value) => new Intl.NumberFormat("en-US", {
  style: "currency", currency: "USD", maximumFractionDigits: 0,
}).format(Number(value) || 0);

// Keeps letters, spaces, hyphens and apostrophes; drops emoji, variation
// selectors and any other symbols so "🍽️ Dining Out" -> "Dining Out".
const cleanCategory = (value) => String(value ?? "")
  .replace(/[^\p{L}\s'-]/gu, "")
  .replace(/\s+/g, " ")
  .trim();

let charts = {};
let currentPeriod = "month"; // "week" | "month" | "year" — drives the category chart + list
let lastSummary = null;

// ---------- Parsing ----------

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const budgetSheet = workbook.Sheets["Budget"];
  const weeklySheet = workbook.Sheets["Weekly Log"];
  if (!budgetSheet || !weeklySheet) {
    throw new Error("This workbook needs a “Budget” sheet and a “Weekly Log” sheet. See the README for the expected layout.");
  }

  const budgetRows = XLSX.utils.sheet_to_json(budgetSheet, { header: 1, defval: null });
  const weeklyRows = XLSX.utils.sheet_to_json(weeklySheet, { header: 1, defval: null });

  const budget = budgetRows.slice(1)
    .filter((row) => row[0] && String(row[0]).toUpperCase() !== "TOTAL")
    .map((row) => ({
      category: cleanCategory(row[0]),
      weekly: Number(row[1] || 0),
      monthly: Number(row[2] || 0),
      yearly: Number(row[3] || 0),
    }))
    .filter((row) => row.category);

  const headers = (weeklyRows[0] || []).map(cleanCategory);
  const totalIndex = headers.findIndex((h) => h.toLowerCase() === "weekly total");
  const categoryIndexes = budget.map((item) => ({
    category: item.category,
    index: headers.findIndex((h) => h.toLowerCase() === item.category.toLowerCase()),
  }));

  const weeks = weeklyRows.slice(1)
    .filter((row) => row[0] && row.slice(1, 1 + budget.length).some((v) => Number(v || 0) !== 0))
    .map((row) => {
      const total = totalIndex >= 0
        ? Number(row[totalIndex] || 0)
        : categoryIndexes.reduce((sum, item) => sum + (item.index >= 0 ? Number(row[item.index] || 0) : 0), 0);
      const categories = {};
      categoryIndexes.forEach((item) => {
        categories[item.category] = item.index >= 0 ? Number(row[item.index] || 0) : 0;
      });
      return { label: String(row[0]), total, categories };
    });

  if (!budget.length) {
    throw new Error("No categories found on the “Budget” sheet.");
  }

  return { budget, weeks };
}

// ---------- Derived numbers ----------

function summarise(data) {
  const weekBudget = data.budget.reduce((sum, x) => sum + x.weekly, 0);
  const monthBudget = data.budget.reduce((sum, x) => sum + x.monthly, 0);
  const yearBudget = data.budget.reduce((sum, x) => sum + x.yearly, 0);

  const latestWeek = data.weeks.at(-1);
  const weekSpent = latestWeek?.total || 0;

  const recentWeeks = data.weeks.slice(-4);
  const monthSpent = recentWeeks.reduce((sum, w) => sum + w.total, 0);

  // "Year" is a rolling window of the most recent 52 logged weeks, not the
  // whole workbook — same rolling-window approach as "month" above. Without
  // this, a workbook with multiple years of history would lump everything
  // ever logged into "this year".
  const recentYearWeeks = data.weeks.slice(-52);
  const yearSpent = recentYearWeeks.reduce((sum, w) => sum + w.total, 0);

  const categoryTotals = data.budget.map((item) => {
    const weekActual = latestWeek?.categories[item.category] || 0;
    const monthActual = recentWeeks.reduce((sum, w) => sum + (w.categories[item.category] || 0), 0);
    const yearActual = recentYearWeeks.reduce((sum, w) => sum + (w.categories[item.category] || 0), 0);
    return {
      category: item.category,
      weekBudget: item.weekly, monthBudget: item.monthly, yearBudget: item.yearly,
      weekActual, monthActual, yearActual,
      // kept for the "top category" insight, which is always a year-to-date read
      actual: yearActual, used: item.yearly ? yearActual / item.yearly : 0,
    };
  }).sort((a, b) => b.used - a.used);

  return {
    weekBudget, monthBudget, yearBudget,
    weekSpent, monthSpent, yearSpent,
    latestWeek, recentWeeks, recentYearWeeks, categoryTotals,
    weeksLogged: data.weeks.length,
    yearWeeksLogged: recentYearWeeks.length,
    allWeeks: data.weeks,
  };
}

function buildInsights(s) {
  const insights = [];

  if (!s.weeksLogged) {
    insights.push({ text: "Add a week of spending to your log to start seeing insights here.", tone: "" });
    return insights;
  }

  const weekDiff = s.weekBudget - s.weekSpent;
  insights.push(weekDiff >= 0
    ? { text: `You're <span class="amount">${money(weekDiff)}</span> under budget this week.`, tone: "positive" }
    : { text: `You're <span class="amount">${money(-weekDiff)}</span> over budget this week.`, tone: "negative" });

  const monthDiff = s.monthBudget - s.monthSpent;
  const monthSpan = Math.min(4, s.recentWeeks.length);
  const monthLabel = `your last ${monthSpan} logged week${monthSpan === 1 ? "" : "s"}`;
  insights.push(monthDiff >= 0
    ? { text: `You're <span class="amount">${money(monthDiff)}</span> under budget over ${monthLabel}.`, tone: "positive" }
    : { text: `You're <span class="amount">${money(-monthDiff)}</span> over budget over ${monthLabel}.`, tone: "negative" });

  const avgWeekly = s.yearSpent / s.yearWeeksLogged;
  const projectedAnnual = avgWeekly * 52;
  const paceDiff = s.yearBudget - projectedAnnual;
  insights.push(paceDiff >= 0
    ? { text: `At your current pace (~${money(avgWeekly)}/week), you're on track to spend about <span class="amount">${money(projectedAnnual)}</span> this year — ${money(paceDiff)} under your annual budget.`, tone: "positive" }
    : { text: `At your current pace (~${money(avgWeekly)}/week), you're on track to spend about <span class="amount">${money(projectedAnnual)}</span> this year — ${money(-paceDiff)} over your annual budget.`, tone: "negative" });

  const top = s.categoryTotals.find((c) => c.actual > 0);
  if (top) {
    const pct = Math.round(top.used * 100);
    insights.push(top.used >= 1
      ? { text: `<span class="amount">${top.category}</span> has already passed its yearly budget, at ${pct}% used.`, tone: "negative" }
      : { text: `<span class="amount">${top.category}</span> is your highest-usage category so far, at ${pct}% of its yearly budget.`, tone: "warn" });
  }

  return insights;
}

// ---------- Rendering: overview ----------

function renderOverview(s) {
  document.getElementById("weekSpent").textContent = money(s.weekSpent);
  document.getElementById("weekBudget").textContent = money(s.weekBudget);
  document.getElementById("monthSpent").textContent = money(s.monthSpent);
  document.getElementById("monthBudget").textContent = money(s.monthBudget);
  document.getElementById("yearSpent").textContent = money(s.yearSpent);
  document.getElementById("yearBudget").textContent = money(s.yearBudget);

  const weekRatio = s.weekBudget ? s.weekSpent / s.weekBudget : 0;
  const monthRatio = s.monthBudget ? s.monthSpent / s.monthBudget : 0;
  const yearRatio = s.yearBudget ? s.yearSpent / s.yearBudget : 0;

  setProgress("weekProgress", weekRatio);
  setProgress("monthProgress", monthRatio);
  setProgress("yearProgress", yearRatio);

  const weekNote = document.getElementById("weekNote");
  weekNote.textContent = s.latestWeek ? `Latest logged week: ${s.latestWeek.label}` : "No week logged yet";
  weekNote.className = "overview-note " + noteClass(weekRatio, s.weeksLogged);

  const monthNote = document.getElementById("monthNote");
  const span = Math.min(4, s.recentWeeks.length);
  monthNote.textContent = span ? `Based on your last ${span} logged week${span === 1 ? "" : "s"}` : "No weeks logged yet";
  monthNote.className = "overview-note " + noteClass(monthRatio, s.weeksLogged);

  const yearNote = document.getElementById("yearNote");
  yearNote.textContent = `${s.yearWeeksLogged} week${s.yearWeeksLogged === 1 ? "" : "s"} logged this year`;
  yearNote.className = "overview-note " + noteClass(yearRatio, s.yearWeeksLogged);
}

function noteClass(ratio, weeksLogged) {
  if (!weeksLogged) return "";
  return ratio >= 1 ? "negative" : ratio >= 0.8 ? "warn" : "positive";
}

function setProgress(id, ratio) {
  const el = document.getElementById(id);
  const pct = Math.max(0, Math.min(100, ratio * 100));
  el.style.width = pct + "%";
  el.className = "progress-fill " + (ratio >= 1 ? "bad" : ratio >= 0.8 ? "warn" : "");
}

function renderInsights(insights) {
  const list = document.getElementById("insightList");
  list.innerHTML = insights.map((i) => `<li class="${i.tone}">${i.text}</li>`).join("");
}

// ---------- Category period view (week / month / year) ----------

const PERIOD_META = {
  week: { budgetKey: "weekBudget", actualKey: "weekActual", budgetLabel: "Weekly budget", note: "This week's spending by category against the weekly allowance." },
  month: { budgetKey: "monthBudget", actualKey: "monthActual", budgetLabel: "Monthly budget", note: "Spending over your last logged weeks by category against the monthly allowance." },
  year: { budgetKey: "yearBudget", actualKey: "yearActual", budgetLabel: "Annual budget", note: "Year-to-date spending by category against the annual allowance." },
};

function getCategoryView(s, period) {
  const meta = PERIOD_META[period] || PERIOD_META.year;
  return s.categoryTotals.map((c) => {
    const actual = c[meta.actualKey];
    const budget = c[meta.budgetKey];
    return { category: c.category, actual, budget, used: budget ? actual / budget : 0 };
  }).sort((a, b) => b.used - a.used);
}

function updateCategoryNote(period) {
  const note = document.getElementById("categoryNote");
  if (note) note.textContent = (PERIOD_META[period] || PERIOD_META.year).note;
}

// ---------- Rendering: charts ----------

function baseLegend() {
  return { position: "bottom", labels: { boxWidth: 8, usePointStyle: true, padding: 14, font: { size: 11.5 }, color: cssVar("--muted") } };
}

function renderTrendChart(s) {
  const wrap = document.getElementById("trendWrap");
  const weeks = s.allWeeks.slice(-12);

  if (!weeks.length) {
    wrap.innerHTML = '<div class="chart-empty">Log a week of spending to see your trend here.</div>';
    return;
  }

  wrap.innerHTML = '<canvas id="trendChart"></canvas>';
  const ctx = document.getElementById("trendChart");
  const accent = cssVar("--accent");
  const muted = cssVar("--muted");
  const lineSoft = cssVar("--line-soft");

  charts.trend = new Chart(ctx, {
    type: "line",
    data: {
      labels: weeks.map((w) => w.label),
      datasets: [
        {
          label: "Actual", data: weeks.map((w) => w.total),
          borderColor: accent, backgroundColor: "transparent",
          tension: 0.35, borderWidth: 2, pointRadius: 3, pointBackgroundColor: accent,
        },
        {
          label: "Weekly budget", data: weeks.map(() => s.weekBudget),
          borderColor: muted, borderDash: [5, 5], borderWidth: 1.5, pointRadius: 0,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: baseLegend(),
        tooltip: { callbacks: { label: (item) => `${item.dataset.label}: ${money(item.parsed.y)}` } },
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 }, color: muted } },
        y: { beginAtZero: true, grid: { color: lineSoft }, ticks: { font: { size: 11 }, color: muted, callback: (v) => "$" + v } },
      },
    },
  });
}

function renderBudgetShapeChart(data) {
  const ctx = document.getElementById("budgetChart");
  const labels = data.budget.map((b) => b.category);
  const values = data.budget.map((b) => b.weekly);

  charts.budget = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: labels.map((l, i) => colorFor(l, i)),
        borderWidth: 2,
        borderColor: cssVar("--paper"),
        hoverOffset: 4,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: "68%",
      plugins: {
        legend: baseLegend(),
        tooltip: { callbacks: { label: (item) => `${item.label}: ${money(item.parsed)}/wk` } },
      },
    },
  });
}

function renderCategoryBarChart(s, period) {
  if (charts.category) { charts.category.destroy(); charts.category = null; }
  const ctx = document.getElementById("categoryChart");
  const view = getCategoryView(s, period);
  const meta = PERIOD_META[period] || PERIOD_META.year;
  const labels = view.map((c) => c.category);
  const actual = view.map((c) => c.actual);
  const budget = view.map((c) => c.budget);
  const muted = cssVar("--muted");
  const lineSoft = cssVar("--line-soft");

  // Give every category row enough height that Chart.js never has to
  // auto-skip a label — otherwise it silently hides labels that don't fit.
  const wrap = document.getElementById("categoryChartWrap");
  wrap.style.height = Math.max(240, labels.length * 30 + 20) + "px";

  charts.category = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "Actual", data: actual, backgroundColor: labels.map((l, i) => colorFor(l, i)), borderRadius: 4, barPercentage: 0.6 },
        { label: meta.budgetLabel, data: budget, backgroundColor: lineSoft, borderRadius: 4, barPercentage: 0.6 },
      ],
    },
    options: {
      indexAxis: "y", responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: baseLegend(),
        tooltip: { callbacks: { label: (item) => `${item.dataset.label}: ${money(item.parsed.x)}` } },
      },
      scales: {
        x: { beginAtZero: true, grid: { color: lineSoft }, ticks: { font: { size: 11 }, color: muted, callback: (v) => "$" + v } },
        y: { grid: { display: false }, ticks: { font: { size: 11.5 }, color: cssVar("--ink"), autoSkip: false } },
      },
    },
  });
}

function renderCharts(s, data, period) {
  if (charts.trend) { charts.trend.destroy(); charts.trend = null; }
  if (charts.budget) { charts.budget.destroy(); charts.budget = null; }
  renderTrendChart(s);
  renderBudgetShapeChart(data);
  renderCategoryBarChart(s, period);
}

// ---------- Rendering: category list ----------

function statusClass(ratio) {
  if (ratio >= 1) return "bad";
  if (ratio >= 0.8) return "warn";
  return "good";
}

function renderCategories(s, period) {
  const list = document.getElementById("categoryList");
  const view = getCategoryView(s, period);
  list.innerHTML = view.map((c, i) => {
    const pct = Math.min(100, c.used * 100);
    const cls = statusClass(c.used);
    return `
      <div class="category-row">
        <div class="category-row-top">
          <span class="category-name"><span class="category-dot" style="background:${colorFor(c.category, i)}"></span>${c.category}</span>
          <span class="category-amounts">${money(c.actual)} of ${money(c.budget)}</span>
        </div>
        <div class="progress small"><div class="progress-fill ${cls === "good" ? "" : cls}" style="width:${pct}%"></div></div>
      </div>`;
  }).join("");
}

// ---------- Period toggle ----------

function setPeriod(period) {
  currentPeriod = period;
  document.querySelectorAll(".period-btn").forEach((btn) => {
    const active = btn.dataset.period === period;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-selected", String(active));
  });
  updateCategoryNote(period);
  if (!lastSummary) return;
  renderCategoryBarChart(lastSummary, period);
  renderCategories(lastSummary, period);
}

document.querySelectorAll(".period-btn").forEach((btn) => {
  btn.addEventListener("click", () => setPeriod(btn.dataset.period));
});

// ---------- Top-level render ----------

function render(data) {
  const s = summarise(data);
  lastSummary = s;
  renderOverview(s);
  renderInsights(buildInsights(s));
  renderCharts(s, data, currentPeriod);
  renderCategories(s, currentPeriod);
  updateCategoryNote(currentPeriod);
}

function setStatus(text, tone) {
  document.getElementById("statusText").textContent = text;
  document.getElementById("statusDot").className = "status-dot" + (tone ? ` ${tone}` : "");
}

// ---------- Data loading ----------

async function loadLinkedWorkbook() {
  try {
    const response = await fetch(WORKBOOK_PATH, { cache: "no-store" });
    if (!response.ok) throw new Error("Workbook not found next to index.html.");
    const buffer = await response.arrayBuffer();
    render(parseWorkbook(buffer));
    setStatus("Sample workbook loaded", "");
  } catch (error) {
    render({
      budget: FALLBACK_BUDGET.map(([category, weekly]) => ({
        category, weekly, monthly: weekly * 4.33, yearly: weekly * 52,
      })),
      weeks: [],
    });
    setStatus("Preview loaded — open a workbook to see real spending", "warn");
  }
}

document.getElementById("fileInput").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    render(parseWorkbook(buffer));
    setStatus(`${file.name} loaded privately in your browser`, "");
  } catch (error) {
    setStatus(error.message, "bad");
  }
});

document.getElementById("templateBtn").addEventListener("click", () => {
  window.location.href = TEMPLATE_PATH;
});

loadLinkedWorkbook();

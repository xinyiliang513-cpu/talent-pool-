const state = {
  sourceName: "",
  headers: [],
  rows: [],
  people: [],
  charts: {},
  summaries: {},
  currentSummary: "countries",
  chartTypes: {
    hoursChart: "bar",
    projectsChart: "doughnut",
    englishChart: "doughnut",
    timelineChart: "line",
    experienceChart: "bar",
    businessLineChart: "barY",
    degreeChart: "doughnut",
    countryChart: "barY",
    majorChart: "barY",
  },
};

const columnAliases = {
  email: ["contact email", "email", "crowdgen account", "ep account"],
  name: ["name", "姓名"],
  country: ["resident country/region", "country", "国家", "地区"],
  city: ["resident city", "city"],
  pm: ["pm", "recommed to which pm", "recommend to which pm"],
  degree: ["highest degree", "educational background", "学历"],
  major: ["major/specialty", "specified major", "学科专业", "专业"],
  domain: ["domain", "领域"],
  english: ["english level", "英语"],
  hours: ["average daily working hours", "working hours", "工作时长"],
  yearsExperience: ["years of work experience", "work experience", "years experience", "工作经验", "工作年限", "经验年限"],
  projectCount: ["number of projects", "project number", "项目数量"],
  languageCountry: ["语种-国家", "language-country", "language country"],
  nativeLanguage: ["native language", "native language(s)", "母语"],
  foreignLanguage: ["foreign language skills", "foreign language", "外语"],
  completedAt: ["completed at", "提交时间"],
};

const projectColumnBases = [
  "Project Types",
  "Business Line",
  "Business Scenario",
  "Data Types",
  "Annotation Scenarios",
  "Collection Types",
  "Duration",
  "Platform",
  "Role",
  "Accuracy",
  "Efficiency",
  "Last Work Time",
];

const palette = [
  "#1769aa",
  "#15803d",
  "#b45309",
  "#7c3aed",
  "#dc2626",
  "#0891b2",
  "#4f46e5",
  "#be185d",
  "#65a30d",
  "#0f766e",
  "#9333ea",
  "#ca8a04",
];

const MISSING_LABEL = "Missing";
const OTHER_LABEL = "Other";
const PEOPLE_LABEL = "People";
const SUBMISSIONS_LABEL = "Submissions";

const chartLabelsPlugin = {
  id: "chartLabels",
  afterDatasetsDraw(chart) {
    const { ctx, data } = chart;
    const dataset = data.datasets[0];
    const options = chart.options.plugins?.chartLabels || {};
    if (!dataset || options.display === false) return;
    const values = dataset.data.map((value) => Number(value) || 0);
    const total = values.reduce((sum, value) => sum + value, 0);
    if (!total) return;

    if (chart.config.type === "doughnut" || chart.config.type === "pie") {
      drawDoughnutLabels(chart, values, total);
      return;
    }

    ctx.save();
    ctx.font = "700 12px -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#17202f";

    chart.getDatasetMeta(0).data.forEach((element, index) => {
      const value = values[index];
      if (!value) return;
      const label = formatMetricLabel(value, total);
      if (chart.config.type === "line") {
        const point = element.tooltipPosition();
        const width = ctx.measureText(label).width;
        const labelX = clamp(point.x, chart.chartArea.left + width / 2 + 4, chart.chartArea.right - width / 2 - 4);
        const labelY = Math.max(chart.chartArea.top + 12, point.y - 14);
        ctx.textAlign = "center";
        ctx.fillText(label, labelX, labelY);
        return;
      }
      const horizontal = chart.options.indexAxis === "y";
      const point = element.tooltipPosition();
      ctx.textAlign = horizontal ? "left" : "center";
      ctx.fillText(label, horizontal ? point.x + 10 : point.x, horizontal ? point.y : point.y - 12);
    });
    ctx.restore();
  },
};

function drawDoughnutLabels(chart, values, total) {
  const { ctx, chartArea } = chart;
  const meta = chart.getDatasetMeta(0);
  const sides = { left: [], right: [] };

  meta.data.forEach((arc, index) => {
    const value = values[index];
    if (!value) return;
    const angle = (arc.startAngle + arc.endAngle) / 2;
    const side = Math.cos(angle) >= 0 ? "right" : "left";
    const anchorX = arc.x + Math.cos(angle) * arc.outerRadius;
    const anchorY = arc.y + Math.sin(angle) * arc.outerRadius;
    const elbowX = arc.x + Math.cos(angle) * (arc.outerRadius + 18);
    const elbowY = arc.y + Math.sin(angle) * (arc.outerRadius + 18);
    const textX = side === "right" ? chartArea.right + 22 : chartArea.left - 22;
    sides[side].push({
      anchorX,
      anchorY,
      elbowX,
      elbowY,
      textX,
      textY: elbowY,
      label: formatMetricLabel(value, total),
    });
  });

  const bounds = {
    top: chartArea.top + 10,
    bottom: chartArea.bottom - 10,
  };
  adjustLabelSide(sides.left, bounds);
  adjustLabelSide(sides.right, bounds);

  ctx.save();
  ctx.font = "700 12px -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif";
  ctx.textBaseline = "middle";
  ctx.lineWidth = 1;
  ctx.strokeStyle = "#98a2b3";
  ctx.fillStyle = "#17202f";

  ["left", "right"].forEach((side) => {
    sides[side].forEach((item) => {
      const horizontalEndX = side === "right" ? item.textX - 6 : item.textX + 6;
      ctx.beginPath();
      ctx.moveTo(item.anchorX, item.anchorY);
      ctx.lineTo(item.elbowX, item.elbowY);
      ctx.lineTo(horizontalEndX, item.textY);
      ctx.stroke();
      ctx.textAlign = side === "right" ? "left" : "right";
      ctx.fillText(item.label, item.textX, item.textY);
    });
  });
  ctx.restore();
}

function adjustLabelSide(items, bounds) {
  if (!items.length) return;
  const gap = 18;
  items.sort((a, b) => a.textY - b.textY);
  items.forEach((item, index) => {
    item.textY = Math.max(bounds.top, Math.min(bounds.bottom, item.textY));
    if (index > 0) item.textY = Math.max(item.textY, items[index - 1].textY + gap);
  });
  for (let index = items.length - 2; index >= 0; index -= 1) {
    items[index].textY = Math.min(items[index].textY, items[index + 1].textY - gap);
  }
  const overflowTop = bounds.top - items[0].textY;
  if (overflowTop > 0) items.forEach((item) => { item.textY += overflowTop; });
  const overflowBottom = items[items.length - 1].textY - bounds.bottom;
  if (overflowBottom > 0) items.forEach((item) => { item.textY -= overflowBottom; });
}

function formatMetricLabel(value, total) {
  return `${value.toLocaleString()} (${percent(value, total)})`;
}

const el = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", () => {
  if (window.Chart) {
    try {
      Chart.register(chartLabelsPlugin);
    } catch (error) {
      // Chart.js ignores duplicate plugin ids in most builds; this keeps older builds from blocking the page.
    }
  }
  bindFileInput();
  bindFilters();
  bindActions();
});

function bindFileInput() {
  const fileInput = el("fileInput");
  const dropZone = el("dropZone");
  fileInput.addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) loadFile(file);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.remove("dragover");
    });
  });

  dropZone.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files?.[0];
    if (file) loadFile(file);
  });
}

function bindFilters() {
  ["countryFilter", "englishFilter", "degreeFilter", "pmFilter", "searchInput"].forEach((id) => {
    el(id).addEventListener("input", renderDashboard);
  });
  el("resetFilters").addEventListener("click", () => {
    ["countryFilter", "englishFilter", "degreeFilter", "pmFilter"].forEach((id) => {
      el(id).value = "";
    });
    el("searchInput").value = "";
    renderDashboard();
  });
}

function bindActions() {
  el("summarySelect").addEventListener("change", (event) => {
    state.currentSummary = event.target.value;
    renderSummaryTable();
  });
  el("copyTable").addEventListener("click", copyCurrentTable);
  el("downloadCsv").addEventListener("click", downloadCurrentCsv);
  el("copyReport").addEventListener("click", copyReportForFeishu);
  el("downloadHtmlReport").addEventListener("click", downloadHtmlReport);
  el("downloadReport").addEventListener("click", downloadMarkdownReport);
  document.querySelectorAll("[data-download-chart]").forEach((button) => {
    button.addEventListener("click", () => downloadChart(button.dataset.downloadChart));
  });
  document.querySelectorAll("[data-chart-type]").forEach((select) => {
    const chartId = select.dataset.chartType;
    if (state.chartTypes[chartId]) select.value = state.chartTypes[chartId];
    select.addEventListener("change", () => {
      state.chartTypes[chartId] = select.value;
      renderDashboard();
    });
  });
}

async function loadFile(file) {
  if (!window.XLSX || !window.Chart) {
    alert("Chart libraries are still loading. Please refresh after CDN access is available.");
    return;
  }

  state.sourceName = file.name;
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const firstSheet = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheet];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const { headers, rows } = normalizeMatrix(matrix);

  state.rows = rows;
  state.headers = headers;
  state.people = dedupePeople(rows, headers);
  populateFilterOptions();
  el("emptyState").classList.add("hidden");
  el("dashboard").classList.remove("hidden");
  renderDashboard();
}

function syncReportActions() {
  const hasReport = state.people.length > 0;
  ["copyReport", "downloadHtmlReport", "downloadReport"].forEach((id) => {
    el(id).disabled = !hasReport;
  });
}

function normalizeMatrix(matrix) {
  const headerIndex = detectHeaderRow(matrix);
  const headers = (matrix[headerIndex] || []).map((value, index) => {
    const label = clean(value);
    return label || `Column ${index + 1}`;
  });
  const rows = matrix.slice(headerIndex + 1).map((line) => {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = normalizeCell(line[index]);
    });
    return row;
  }).filter((row) => Object.values(row).some((value) => clean(value)));
  return { headers, rows };
}

function detectHeaderRow(matrix) {
  const scores = matrix.slice(0, 10).map((line) => {
    const labels = line.map((value) => clean(value).toLowerCase());
    const hits = Object.values(columnAliases).flat().filter((alias) =>
      labels.some((label) => label.includes(alias)),
    ).length;
    const filled = labels.filter(Boolean).length;
    return hits * 10 + filled;
  });
  const bestScore = Math.max(...scores);
  return Math.max(0, scores.indexOf(bestScore));
}

function dedupePeople(rows, headers) {
  const emailColumn = findColumn(headers, columnAliases.email);
  const seen = new Map();
  rows.forEach((row, rowIndex) => {
    const rawEmail = clean(row[emailColumn]).toLowerCase();
    const fallback = `row-${rowIndex}`;
    const key = rawEmail || fallback;
    if (!seen.has(key)) {
      seen.set(key, buildPerson(row, headers, key));
    } else {
      mergeProjectExperience(seen.get(key), row, headers);
    }
  });
  return Array.from(seen.values());
}

function buildPerson(row, headers, key) {
  const get = (field) => row[findColumn(headers, columnAliases[field])] || "";
  const person = {
    key,
    email: get("email"),
    name: get("name"),
    country: get("country"),
    city: get("city"),
    pm: get("pm"),
    degree: get("degree"),
    major: get("major"),
    domain: get("domain"),
    english: get("english"),
    hours: get("hours"),
    yearsExperience: get("yearsExperience"),
    projectCount: get("projectCount"),
    languageCountry: get("languageCountry"),
    nativeLanguage: get("nativeLanguage"),
    foreignLanguage: get("foreignLanguage"),
    completedAt: get("completedAt"),
    sourceRows: [row],
    projects: [],
  };
  mergeProjectExperience(person, row, headers);
  return person;
}

function mergeProjectExperience(person, row, headers) {
  if (!person.sourceRows.includes(row)) person.sourceRows.push(row);
  for (let group = 0; group < 3; group += 1) {
    const suffix = group === 0 ? "" : `.${group}`;
    const project = {};
    projectColumnBases.forEach((base) => {
      const column = headers.find((header) => header === `${base}${suffix}`);
      if (column) project[base] = row[column] || "";
    });
    if (Object.values(project).some((value) => clean(value))) {
      person.projects.push(project);
    }
  }
}

function findColumn(headers, aliases) {
  const lowered = headers.map((header) => String(header).toLowerCase().trim());
  for (const alias of aliases) {
    const exactIndex = lowered.findIndex((header) => header === alias);
    if (exactIndex >= 0) return headers[exactIndex];
  }
  for (const alias of aliases) {
    const fuzzyIndex = lowered.findIndex((header) => header.includes(alias));
    if (fuzzyIndex >= 0) return headers[fuzzyIndex];
  }
  return "";
}

function populateFilterOptions() {
  setSelectOptions("countryFilter", uniqueSorted(state.people.map((person) => person.country)));
  setSelectOptions("englishFilter", uniqueSorted(state.people.map((person) => person.english)));
  setSelectOptions("degreeFilter", uniqueSorted(state.people.map((person) => normalizeDegree(person.degree))));
  setSelectOptions("pmFilter", uniqueSorted(state.people.map((person) => person.pm)));
}

function setSelectOptions(id, values) {
  const select = el(id);
  const current = select.value;
  select.innerHTML = '<option value="">All</option>';
  values.slice(0, 300).forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  if (values.includes(current)) select.value = current;
}

function renderDashboard() {
  const people = getFilteredPeople();
  const allPeople = state.people;
  const projectRows = people.flatMap((person) => person.projects.map((project) => ({ person, project })));

  syncReportActions();
  state.summaries = buildSummaries(people, projectRows);
  renderKpis(people, allPeople, projectRows);
  renderCharts(people, projectRows);
  renderSummaryTable();
  renderRawMatches(people);
  el("sourceNote").textContent = `${state.sourceName} · Source rows ${state.rows.length.toLocaleString()} · Current selection ${people.length.toLocaleString()} people`;
}

function getFilteredPeople() {
  const country = el("countryFilter").value;
  const english = el("englishFilter").value;
  const degree = el("degreeFilter").value;
  const pm = el("pmFilter").value;
  const keyword = el("searchInput").value.trim().toLowerCase();

  return state.people.filter((person) => {
    if (country && person.country !== country) return false;
    if (english && person.english !== english) return false;
    if (degree && normalizeDegree(person.degree) !== degree) return false;
    if (pm && person.pm !== pm) return false;
    if (keyword) {
      const haystack = [
        person.name,
        person.email,
        person.major,
        person.domain,
        person.yearsExperience,
        person.country,
        person.languageCountry,
      ].join(" ").toLowerCase();
      if (!haystack.includes(keyword)) return false;
    }
    return true;
  });
}

function renderKpis(people, allPeople, projectRows) {
  const total = people.length;
  const projectPeople = people.filter((person) => person.projects.length > 0).length;
  const avgProjects = total ? projectRows.length / total : 0;
  const countryCount = countBy(people, (person) => person.country || MISSING_LABEL).length;
  const highEnglish = people.filter((person) => isHighEnglish(person.english)).length;
  const kpis = [
    ["Unique People", total.toLocaleString(), `Counted by Contact Email, ${percent(total, allPeople.length)} of total pool`],
    ["With Project Experience", projectPeople.toLocaleString(), `${percent(projectPeople, total)} of current talent`],
    ["Avg. Project Records", avgProjects.toFixed(1), "Merged Project No.1-3"],
    ["Country/Region Coverage", countryCount.toLocaleString(), "Based on Resident Country/Region"],
    ["Strong English", highEnglish.toLocaleString(), `${percent(highEnglish, total)} of current talent`],
  ];

  const grid = el("kpiGrid");
  const template = el("kpiTemplate");
  grid.innerHTML = "";
  kpis.forEach(([label, value, note]) => {
    const node = template.content.cloneNode(true);
    node.querySelector("span").textContent = label;
    node.querySelector("strong").textContent = value;
    node.querySelector("small").textContent = note;
    grid.appendChild(node);
  });
}

function renderCharts(people, projectRows) {
  renderFlexibleChart("hoursChart", PEOPLE_LABEL, bucketCounts(people, (person) => person.hours, hourBucketFromText), {
    compactBars: true,
  });
  renderFlexibleChart("projectsChart", PEOPLE_LABEL, projectCountBuckets(people));
  renderFlexibleChart("englishChart", PEOPLE_LABEL, topCounts(countBy(people, (person) => clean(person.english) || MISSING_LABEL), 8));
  renderFlexibleChart("timelineChart", SUBMISSIONS_LABEL, timelineCounts(people));
  renderFlexibleChart("experienceChart", PEOPLE_LABEL, bucketCounts(people, (person) => person.yearsExperience, experienceBucketFromText, experienceBucketOrder()), {
    compactBars: true,
  });
  renderFlexibleChart("businessLineChart", PEOPLE_LABEL, topCounts(countProjectValues(projectRows, "Business Line"), 12));
  renderFlexibleChart("degreeChart", PEOPLE_LABEL, countBy(people, (person) => normalizeDegree(person.degree)));
  renderFlexibleChart("countryChart", PEOPLE_LABEL, topCounts(countBy(people, (person) => person.country || MISSING_LABEL), 15));
  renderFlexibleChart("majorChart", PEOPLE_LABEL, topCounts(countTokenValues(people, (person) => person.major || person.domain), 20, { includeOther: false }));
}

function renderFlexibleChart(canvasId, label, rows, chartOptions = {}) {
  const type = state.chartTypes[canvasId] || "bar";
  if (type === "doughnut" || type === "pie") {
    renderDoughnut(canvasId, rows, type);
    return;
  }
  if (type === "line") {
    renderLine(canvasId, rows, label);
    return;
  }
  renderBar(canvasId, label, rows, {
    ...chartOptions,
    indexAxis: type === "barY" ? "y" : "x",
  });
}

function renderBar(canvasId, label, rows, chartOptions = {}) {
  const labels = rows.map((row) => row.name);
  const values = rows.map((row) => row.count);
  renderChart(canvasId, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label,
        data: values,
        backgroundColor: palette,
        borderRadius: 4,
        barPercentage: chartOptions.compactBars ? 0.42 : 0.74,
        categoryPercentage: chartOptions.compactBars ? 0.62 : 0.82,
      }],
    },
    options: baseChartOptions(chartOptions),
  });
}

function renderDoughnut(canvasId, rows, type = "doughnut") {
  renderChart(canvasId, {
    type,
    data: {
      labels: rows.map((row) => row.name),
      datasets: [{
        data: rows.map((row) => row.count),
        backgroundColor: palette,
        borderWidth: 0,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: type === "doughnut" ? "58%" : 0,
      layout: { padding: { top: 24, right: 120, bottom: 24, left: 120 } },
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
        chartLabels: { display: true },
      },
    },
  });
}

function renderLine(canvasId, rows, label = PEOPLE_LABEL) {
  renderChart(canvasId, {
    type: "line",
    data: {
      labels: rows.map((row) => row.name),
      datasets: [{
        label,
        data: rows.map((row) => row.count),
        borderColor: "#1769aa",
        backgroundColor: "rgba(23, 105, 170, 0.14)",
        fill: true,
        tension: 0.32,
        pointRadius: 3,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 30, right: 44, bottom: 8, left: 26 } },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (context) => formatMetricLabel(Number(context.raw) || 0, rows.reduce((sum, row) => sum + row.count, 0)) } },
        chartLabels: { display: true },
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, ticks: { precision: 0 } },
      },
    },
  });
}

function renderChart(canvasId, config) {
  if (state.charts[canvasId]) state.charts[canvasId].destroy();
  state.charts[canvasId] = new Chart(el(canvasId), config);
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function baseChartOptions(extra = {}) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: extra.indexAxis || "x",
    layout: { padding: extra.indexAxis === "y" ? { top: 10, right: 118, bottom: 8, left: 8 } : { top: 34, right: 18, bottom: 8, left: 8 } },
    plugins: {
      legend: { display: false },
      tooltip: { callbacks: { label: (context) => formatMetricLabel(Number(context.raw) || 0, context.dataset.data.reduce((sum, value) => sum + (Number(value) || 0), 0)) } },
      chartLabels: { display: true },
    },
    scales: {
      x: { beginAtZero: true, ticks: { precision: 0 } },
      y: { beginAtZero: true, ticks: { precision: 0 } },
    },
  };
}

function buildSummaries(people, projectRows) {
  const countries = enrichSummary(countBy(people, (person) => person.country || MISSING_LABEL), people.length);
  const businessLines = enrichSummary(countProjectValues(projectRows, "Business Line"), people.length);
  const domains = enrichSummary(countTokenValues(people, (person) => person.domain), people.length);
  const majors = enrichSummary(countTokenValues(people, (person) => person.major || person.domain), people.length);
  const experience = enrichSummary(bucketCounts(people, (person) => person.yearsExperience, experienceBucketFromText, experienceBucketOrder()), people.length);
  const details = people.map((person) => ({
    name: person.name || MISSING_LABEL,
    email: person.email || MISSING_LABEL,
    country: person.country || MISSING_LABEL,
    english: person.english || MISSING_LABEL,
    degree: normalizeDegree(person.degree),
    major: person.major || MISSING_LABEL,
    hours: person.hours || MISSING_LABEL,
    yearsExperience: person.yearsExperience || MISSING_LABEL,
    projectCount: person.projects.length || parseProjectCount(person.projectCount) || 0,
    pm: person.pm || MISSING_LABEL,
  }));
  return { countries, businessLines, domains, majors, experience, details };
}

function renderSummaryTable() {
  const table = el("summaryTable");
  const rows = state.summaries[state.currentSummary] || [];
  const isDetails = state.currentSummary === "details";
  const headers = isDetails
    ? ["Name", "Email", "Country/Region", "English Level", "Education", "Major", "Working Hours", "Years of Work Experience", "Project Count", "PM"]
    : ["Category", "People", "Share"];

  table.innerHTML = "";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.slice(0, isDetails ? 500 : 80).forEach((row) => {
    const tr = document.createElement("tr");
    const values = isDetails
      ? [row.name, row.email, row.country, row.english, row.degree, row.major, row.hours, row.yearsExperience, row.projectCount, row.pm]
      : [row.name, row.count, row.share];
    values.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

function copyCurrentTable() {
  const csv = currentSummaryCsv();
  navigator.clipboard.writeText(csv).then(() => {
    el("copyTable").textContent = "Copied";
    setTimeout(() => {
      el("copyTable").textContent = "Copy";
    }, 1200);
  });
}

function downloadCurrentCsv() {
  const names = {
    countries: "country-summary",
    businessLines: "business-line-summary",
    domains: "domain-summary",
    majors: "major-summary",
    experience: "work-experience-summary",
    details: "talent-details",
  };
  downloadText(`${names[state.currentSummary] || "summary"}.csv`, currentSummaryCsv(), "text/csv;charset=utf-8");
}

function currentSummaryCsv() {
  const rows = state.summaries[state.currentSummary] || [];
  const isDetails = state.currentSummary === "details";
  const headers = isDetails
    ? ["Name", "Email", "Country", "English Level", "Degree", "Major", "Working Hours", "Years of Work Experience", "Project Count", "PM"]
    : ["Category", "People", "Share"];
  const lines = [headers, ...rows.map((row) => isDetails
    ? [row.name, row.email, row.country, row.english, row.degree, row.major, row.hours, row.yearsExperience, row.projectCount, row.pm]
    : [row.name, row.count, row.share])];
  return lines.map((line) => line.map(csvEscape).join(",")).join("\n");
}

function downloadMarkdownReport() {
  downloadText("talent-pool-analysis-report.md", buildMarkdownReport(), "text/markdown;charset=utf-8");
}

async function copyReportForFeishu() {
  const htmlReport = buildHtmlReport();
  const textReport = buildMarkdownReport();
  try {
    if (window.ClipboardItem && navigator.clipboard.write) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([htmlReport], { type: "text/html" }),
          "text/plain": new Blob([textReport], { type: "text/plain" }),
        }),
      ]);
    } else {
      await navigator.clipboard.writeText(textReport);
    }
    el("copyReport").textContent = "Copied";
    setTimeout(() => {
      el("copyReport").textContent = "Copy Report";
    }, 1400);
  } catch (error) {
    alert("Copy failed. Please allow clipboard permission in the browser.");
  }
}

function downloadHtmlReport() {
  downloadText("talent-pool-analysis-report.html", buildHtmlReport(), "text/html;charset=utf-8");
}

function buildMarkdownReport() {
  const people = getFilteredPeople();
  const projectRows = people.flatMap((person) => person.projects.map((project) => ({ person, project })));
  const sections = buildReportSections(people, projectRows);
  const lines = [
    "# Talent Pool Analytics Report",
    "",
    `- Source file: ${state.sourceName}`,
    `- Counting method: unique people by Contact Email`,
    `- Current selection: ${people.length.toLocaleString()}`,
    `- People with project experience: ${people.filter((person) => person.projects.length > 0).length.toLocaleString()}`,
    `- Country/region coverage: ${countBy(people, (person) => person.country || MISSING_LABEL).length.toLocaleString()}`,
    "",
    ...sections.flatMap((section) => [
      `## ${section.title}`,
      markdownTable(section.rows),
      "",
    ]),
  ];
  return lines.join("\n");
}

function buildHtmlReport() {
  const people = getFilteredPeople();
  const projectRows = people.flatMap((person) => person.projects.map((project) => ({ person, project })));
  const experienced = people.filter((person) => person.projects.length > 0).length;
  const sections = buildReportSections(people, projectRows);
  const cards = [
    ["Current Selection", people.length.toLocaleString()],
    ["With Project Experience", experienced.toLocaleString()],
    ["Country/Region Coverage", countBy(people, (person) => person.country || MISSING_LABEL).length.toLocaleString()],
    ["Strong English", people.filter((person) => isHighEnglish(person.english)).length.toLocaleString()],
  ];

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>Talent Pool Analytics Report</title>
  <style>
    body { margin: 0; color: #17202f; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif; background: #f5f7fb; line-height: 1.55; }
    main { max-width: 1040px; margin: 0 auto; padding: 36px 24px 52px; }
    h1 { margin: 0 0 8px; font-size: 30px; }
    h2 { margin: 30px 0 12px; font-size: 20px; }
    .meta { color: #667085; margin-bottom: 22px; }
    .grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin: 22px 0; }
    .card { background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; padding: 16px; }
    .card span { display: block; color: #667085; font-size: 12px; font-weight: 700; }
    .card strong { display: block; font-size: 28px; margin-top: 4px; }
    .chart-block { background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; padding: 16px; margin: 18px 0 24px; page-break-inside: avoid; }
    .chart-block img { display: block; width: 100%; max-height: 430px; object-fit: contain; border: 1px solid #eef2f7; border-radius: 6px; margin: 8px 0 14px; }
    table { width: 100%; border-collapse: collapse; background: #fff; border: 1px solid #d9e0ea; }
    th, td { border-bottom: 1px solid #e8edf4; padding: 10px 12px; text-align: left; font-size: 14px; }
    th { background: #f0f5fb; }
    @media (max-width: 760px) { .grid { grid-template-columns: 1fr 1fr; } }
  </style>
</head>
<body>
  <main>
    <h1>Talent Pool Analytics Report</h1>
    <div class="meta">Source file: ${escapeHtml(state.sourceName)} · Counting method: unique people by Contact Email</div>
    <section class="grid">${cards.map(([label, value]) => `<article class="card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></article>`).join("")}</section>
    ${sections.map((section) => `<section class="chart-block">
      <h2>${escapeHtml(section.title)}</h2>
      ${chartImageHtml(section.chartId)}
      ${htmlTable(section.rows)}
    </section>`).join("")}
  </main>
</body>
</html>`;
}

function buildReportSections(people, projectRows) {
  const total = people.length;
  return [
    {
      title: "Working Hours",
      chartId: "hoursChart",
      rows: enrichSummary(bucketCounts(people, (person) => person.hours, hourBucketFromText), total),
    },
    {
      title: "Project Count",
      chartId: "projectsChart",
      rows: enrichSummary(projectCountBuckets(people), total),
    },
    {
      title: "English Level",
      chartId: "englishChart",
      rows: enrichSummary(topCounts(countBy(people, (person) => clean(person.english) || MISSING_LABEL), 8), total),
    },
    {
      title: "Submission Trend by Month",
      chartId: "timelineChart",
      rows: enrichSummary(timelineCounts(people), total),
    },
    {
      title: "Years of Work Experience",
      chartId: "experienceChart",
      rows: enrichSummary(bucketCounts(people, (person) => person.yearsExperience, experienceBucketFromText, experienceBucketOrder()), total),
    },
    {
      title: "Business Line Top 12",
      chartId: "businessLineChart",
      rows: enrichSummary(topCounts(countProjectValues(projectRows, "Business Line"), 12), total),
    },
    {
      title: "Education",
      chartId: "degreeChart",
      rows: enrichSummary(countBy(people, (person) => normalizeDegree(person.degree)), total),
    },
    {
      title: "Country/Region Top 15",
      chartId: "countryChart",
      rows: enrichSummary(topCounts(countBy(people, (person) => person.country || MISSING_LABEL), 15), total),
    },
    {
      title: "Major/Domain Top 20",
      chartId: "majorChart",
      rows: enrichSummary(topCounts(countTokenValues(people, (person) => person.major || person.domain), 20, { includeOther: false }), total),
    },
  ];
}

function chartImageHtml(chartId) {
  const chart = state.charts[chartId];
  if (!chart) return "";
  return `<img src="${chart.toBase64Image("image/png", 1)}" alt="${escapeHtml(chartId)} chart" />`;
}

function markdownTable(rows) {
  const tableRows = rows.map((row) => `| ${row.name} | ${row.count} | ${row.share} |`);
  return ["| Category | People | Share |", "| --- | ---: | ---: |", ...tableRows].join("\n");
}

function htmlTable(rows) {
  return `<table><thead><tr><th>Category</th><th>People</th><th>Share</th></tr></thead><tbody>${rows.map((row) =>
    `<tr><td>${escapeHtml(row.name)}</td><td>${escapeHtml(row.count)}</td><td>${escapeHtml(row.share)}</td></tr>`,
  ).join("")}</tbody></table>`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function downloadChart(canvasId) {
  const chart = state.charts[canvasId];
  if (!chart) return;
  const link = document.createElement("a");
  link.download = `${canvasId}.png`;
  link.href = chart.toBase64Image("image/png", 1);
  link.click();
}

function downloadText(filename, text, type) {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function countBy(items, getName) {
  const counts = new Map();
  items.forEach((item) => {
    const name = clean(getName(item)) || MISSING_LABEL;
    counts.set(name, (counts.get(name) || 0) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).sort(sortCountThenName);
}

function countTokenValues(items, getValue) {
  const counts = new Map();
  items.forEach((item) => {
    const tokens = splitMultiValue(getValue(item)).map(normalizeAnalyticCategory).filter(Boolean);
    const unique = new Set(tokens.length ? tokens : [MISSING_LABEL]);
    unique.forEach((token) => counts.set(token, (counts.get(token) || 0) + 1));
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).sort(sortCountThenName);
}

function countProjectValues(projectRows, field) {
  const counts = new Map();
  projectRows.forEach(({ person, project }) => {
    const tokens = splitMultiValue(project[field]).map(normalizeAnalyticCategory).filter(Boolean);
    tokens.forEach((token) => {
      const key = `${token}::${person.key}`;
      counts.set(key, { name: token, person: person.key });
    });
  });
  const peopleByValue = new Map();
  counts.forEach(({ name }) => {
    peopleByValue.set(name, (peopleByValue.get(name) || 0) + 1);
  });
  return Array.from(peopleByValue, ([name, count]) => ({ name, count })).sort(sortCountThenName);
}

function bucketCounts(items, getValue, getBucket, order = [MISSING_LABEL, "0-4h", "4-8h", "8h+"]) {
  const counts = new Map(order.map((name) => [name, 0]));
  items.forEach((item) => {
    const bucket = getBucket(getValue(item));
    counts.set(bucket, (counts.get(bucket) || 0) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).filter((row) => row.count > 0);
}

function projectCountBuckets(people) {
  const counts = new Map([
    ["Missing/0", 0],
    ["1 project", 0],
    ["2 projects", 0],
    ["3 projects", 0],
    ["4-5 projects", 0],
    ["6+ projects", 0],
  ]);
  people.forEach((person) => {
    const count = Math.max(person.projects.length, parseProjectCount(person.projectCount));
    const bucket = count <= 0 ? "Missing/0" : count === 1 ? "1 project" : count === 2 ? "2 projects" : count === 3 ? "3 projects" : count <= 5 ? "4-5 projects" : "6+ projects";
    counts.set(bucket, counts.get(bucket) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).filter((row) => row.count > 0);
}

function timelineCounts(people) {
  const counts = new Map();
  const monthDates = [];
  people.forEach((person) => {
    const date = parseDate(person.completedAt);
    if (!date) return;
    const key = monthKey(date);
    if (date) monthDates.push(new Date(date.getFullYear(), date.getMonth(), 1));
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  if (monthDates.length) {
    const start = new Date(Math.min(...monthDates.map((date) => date.getTime())));
    const latestDataMonth = new Date(Math.max(...monthDates.map((date) => date.getTime())));
    const now = new Date();
    const currentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const end = currentMonth > latestDataMonth ? currentMonth : latestDataMonth;
    for (const cursor = new Date(start); cursor <= end; cursor.setMonth(cursor.getMonth() + 1)) {
      const key = monthKey(cursor);
      if (!counts.has(key)) counts.set(key, 0);
    }
  }
  return Array.from(counts, ([name, count]) => ({ name, count }))
    .sort((a, b) => a.name.localeCompare(b.name, "en"));
}

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function enrichSummary(rows, total) {
  return rows.map((row) => ({
    ...row,
    share: percent(row.count, total),
  }));
}

function topCounts(rows, limit, options = {}) {
  const includeOther = options.includeOther !== false;
  const top = rows.slice(0, limit);
  const rest = rows.slice(limit).reduce((sum, row) => sum + row.count, 0);
  if (includeOther && rest > 0) top.push({ name: OTHER_LABEL, count: rest });
  return top;
}

function sortCountThenName(a, b) {
  return b.count - a.count || String(a.name).localeCompare(String(b.name), "en");
}

function splitMultiValue(value) {
  return clean(value)
    .split(/[,，;；|、\n]+/g)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => item.replace(/\s+/g, " "));
}

function normalizeCategory(value) {
  return normalizeAnalyticCategory(value);
}

function normalizeAnalyticCategory(value) {
  const text = clean(value);
  if (!text) return "";
  if (isPlaceholderCategory(text)) return OTHER_LABEL;
  const key = categoryKey(text);
  const known = {
    adlabeling: "Ads Labelling",
    adlabelling: "Ads Labelling",
    adslabeling: "Ads Labelling",
    adslabelling: "Ads Labelling",
    adslabel: "Ads Labelling",
    annotation: "Annotation",
    collection: "Collection",
    evaluation: "Evaluation",
    coding: "Coding",
    literature: "Literature",
    business: "Business",
    finance: "Finance",
    translation: "Translation",
    llm: "LLM",
    li: "LI (Language Intelligence)",
    languageintelligence: "LI (Language Intelligence)",
    lilanguageintelligence: "LI (Language Intelligence)",
    search: "Search",
    other: OTHER_LABEL,
    multimodalai: "Multimodal AI",
    qa: "QA",
    nlp: "NLP",
    tts: "TTS",
  };
  if (/^li.*languageintelligence$/.test(key)) return "LI (Language Intelligence)";
  if (known[key]) return known[key];
  return smartTitleCase(text);
}

function categoryKey(value) {
  return clean(value)
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "");
}

function smartTitleCase(value) {
  const acronymSet = new Set(["ai", "api", "cv", "id", "li", "llm", "nlp", "ocr", "pm", "qa", "sql", "tts", "ui", "ux"]);
  return clean(value)
    .replace(/[_/\\-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .filter(Boolean)
    .map((word) => {
      const bare = word.replace(/[^a-z0-9]/gi, "");
      const lower = bare.toLowerCase();
      if (acronymSet.has(lower)) return lower.toUpperCase();
      if (/^[A-Z0-9]{2,}$/.test(bare) && bare.length <= 5) return bare;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(" ");
}

function normalizeDegree(value) {
  const text = clean(value);
  const lower = text.toLowerCase();
  if (!text) return OTHER_LABEL;
  if (isPlaceholderCategory(text)) return OTHER_LABEL;
  if (text.includes("未填写")) return OTHER_LABEL;
  if (/^(n\/?a|none|null|unknown|not applicable)$/i.test(text)) return OTHER_LABEL;
  if (lower.includes("phd") || lower.includes("doctor") || text.includes("博士")) return "Doctorate";
  if (lower.includes("master") || text.includes("硕士") || text.includes("研究生")) return "Master's";
  if (lower.includes("bachelor") || text.includes("本科") || text.includes("学士")) return "Bachelor's";
  if (lower.includes("associate") || text.includes("大专") || text.includes("专科") || lower.includes("college")) return "Associate";
  if (lower.includes("high school") || text.includes("高中") || text.includes("中专")) return "High School or Below";
  if (lower.includes("high education") || text.includes("高学历")) return "Master's";
  if (lower.includes("undergraduate")) return "Bachelor's";
  if (lower.includes("postgraduate")) return "Master's";
  return OTHER_LABEL;
}

function isPlaceholderCategory(value) {
  const text = clean(value);
  if (!text) return false;
  const stripped = text.replace(/[\s\-_/\\|.,，。;；:：()[\]{}]+/g, "");
  return stripped.length === 0;
}

function hourBucketFromText(value) {
  const text = clean(value).toLowerCase();
  if (!text) return MISSING_LABEL;
  if (/less[_\s-]*4|<\s*4/.test(text)) return "0-4h";
  if (/4[_\s-]*8|4\s*-\s*8/.test(text)) return "4-8h";
  if (/more[_\s-]*8|>\s*8/.test(text)) return "8h+";
  const match = text.match(/\d+(\.\d+)?/);
  if (!match) return MISSING_LABEL;
  const hours = Number(match[0]);
  if (hours < 4) return "0-4h";
  if (hours <= 8) return "4-8h";
  return "8h+";
}

function experienceBucketOrder() {
  return [MISSING_LABEL, "0-1 years", "1-3 years", "3-5 years", "5-10 years", "10+ years"];
}

function experienceBucketFromText(value) {
  const text = clean(value).toLowerCase();
  if (!text) return MISSING_LABEL;
  if (/no experience|none|无/.test(text)) return "0-1 years";
  const range = text.match(/(\d+(?:\.\d+)?)\s*[-~至到]\s*(\d+(?:\.\d+)?)/);
  if (range) return experienceBucketFromNumber((Number(range[1]) + Number(range[2])) / 2);
  const plus = text.match(/(\d+(?:\.\d+)?)\s*(?:\+|plus|以上|more)/);
  if (plus) return experienceBucketFromNumber(Number(plus[1]) + 0.1);
  const match = text.match(/\d+(\.\d+)?/);
  if (!match) return MISSING_LABEL;
  return experienceBucketFromNumber(Number(match[0]));
}

function experienceBucketFromNumber(years) {
  if (!Number.isFinite(years)) return MISSING_LABEL;
  if (years < 1) return "0-1 years";
  if (years < 3) return "1-3 years";
  if (years < 5) return "3-5 years";
  if (years < 10) return "5-10 years";
  return "10+ years";
}

function parseProjectCount(value) {
  const text = clean(value);
  if (/3[_\s-]*plus|3\s*or\s*more/i.test(text)) return 3;
  if (/1[_\s-]*2|1\s*-\s*2/i.test(text)) return 2;
  if (/no project|^0$/i.test(text)) return 0;
  const match = text.match(/\d+/);
  return match ? Number(match[0]) : 0;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const text = clean(value);
  if (!text) return null;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function isHighEnglish(value) {
  const text = clean(value).toLowerCase();
  return /c1|c2|advanced|fluent|native|professional|流利|母语|高级/.test(text);
}

function percent(numerator, denominator) {
  if (!denominator) return "0%";
  return `${((numerator / denominator) * 100).toFixed(1)}%`;
}

function uniqueSorted(values) {
  return Array.from(new Set(values.map(clean).filter(Boolean))).sort((a, b) => a.localeCompare(b, "en"));
}

function normalizeCell(value) {
  if (value instanceof Date) return value.toLocaleString();
  return clean(value);
}

function clean(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function csvEscape(value) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}
function renderRawMatches(people) {
  const keyword = el("searchInput").value.trim().toLowerCase();
  const panel = el("rawMatchPanel");
  const table = el("rawMatchTable");
  if (!keyword || !state.headers.length) {
    panel.classList.add("hidden");
    table.innerHTML = "";
    return;
  }

  const matchedRows = [];
  people.forEach((person) => {
    person.sourceRows.forEach((row) => {
      if (!matchedRows.includes(row)) matchedRows.push(row);
    });
  });

  panel.classList.remove("hidden");
  el("rawMatchNote").textContent = `Matched ${matchedRows.length.toLocaleString()} source rows`;
  table.innerHTML = "";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  state.headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  matchedRows.slice(0, 100).forEach((row) => {
    const tr = document.createElement("tr");
    state.headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header] || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

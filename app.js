const state = {
  sourceName: "",
  rows: [],
  people: [],
  charts: {},
  summaries: {},
  currentSummary: "countries",
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

const chartLabelsPlugin = {
  id: "chartLabels",
  afterDatasetsDraw(chart) {
    const { ctx, data } = chart;
    const dataset = data.datasets[0];
    if (!dataset) return;
    const values = dataset.data.map((value) => Number(value) || 0);
    const total = values.reduce((sum, value) => sum + value, 0);
    ctx.save();
    ctx.font = "700 12px -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#17202f";

    chart.getDatasetMeta(0).data.forEach((element, index) => {
      const value = values[index];
      if (!value) return;
      const label = chart.config.type === "doughnut"
        ? `${value.toLocaleString()} (${percent(value, total)})`
        : value.toLocaleString();
      if (chart.config.type === "doughnut") {
        const point = element.tooltipPosition();
        ctx.textAlign = "center";
        ctx.fillText(label, point.x, point.y);
        return;
      }
      if (chart.config.type === "line") {
        const point = element.tooltipPosition();
        ctx.textAlign = "center";
        ctx.fillText(label, point.x, point.y - 12);
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
}

async function loadFile(file) {
  if (!window.XLSX || !window.Chart) {
    alert("图表依赖库还没有加载完成，请确认网络可访问 CDN 后刷新页面。");
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
    projectCount: get("projectCount"),
    languageCountry: get("languageCountry"),
    nativeLanguage: get("nativeLanguage"),
    foreignLanguage: get("foreignLanguage"),
    completedAt: get("completedAt"),
    projects: [],
  };
  mergeProjectExperience(person, row, headers);
  return person;
}

function mergeProjectExperience(person, row, headers) {
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
  select.innerHTML = '<option value="">全部</option>';
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
  el("sourceNote").textContent = `${state.sourceName} · 原始行 ${state.rows.length.toLocaleString()} · 当前筛选 ${people.length.toLocaleString()} 人`;
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
  const countryCount = countBy(people, (person) => person.country || "未填写").length;
  const highEnglish = people.filter((person) => isHighEnglish(person.english)).length;
  const kpis = [
    ["去重人数", total.toLocaleString(), `按 Contact Email 统计，占总库 ${percent(total, allPeople.length)}`],
    ["有项目经验", projectPeople.toLocaleString(), `${percent(projectPeople, total)} 的当前人才`],
    ["人均项目记录", avgProjects.toFixed(1), "合并 Project No.1-3"],
    ["国家/地区覆盖", countryCount.toLocaleString(), "按 Resident Country/Region"],
    ["高英语能力", highEnglish.toLocaleString(), `${percent(highEnglish, total)} 的当前人才`],
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
  renderBar("hoursChart", "人数", bucketCounts(people, (person) => person.hours, hourBucketFromText), {
    indexAxis: "x",
  });
  renderDoughnut("projectsChart", projectCountBuckets(people));
  renderDoughnut("englishChart", topCounts(countBy(people, (person) => clean(person.english) || "未填写"), 8));
  renderLine("timelineChart", timelineCounts(people));
  renderBar("projectTypeChart", "人数", topCounts(countProjectValues(projectRows, "Project Types"), 12), {
    indexAxis: "y",
  });
  renderDoughnut("degreeChart", topCounts(countBy(people, (person) => normalizeDegree(person.degree) || "未填写"), 8));
  renderBar("countryChart", "人数", topCounts(countBy(people, (person) => person.country || "未填写"), 10), {
    indexAxis: "y",
  });
  renderBar("majorChart", "人数", topCounts(countTokenValues(people, (person) => person.major || person.domain), 15), {
    indexAxis: "y",
  });
  renderBar("languageChart", "人数", topCounts(countTokenValues(people, (person) => person.languageCountry), 15), {
    indexAxis: "y",
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
      }],
    },
    options: baseChartOptions(chartOptions),
  });
}

function renderDoughnut(canvasId, rows) {
  renderChart(canvasId, {
    type: "doughnut",
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
      layout: { padding: 18 },
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
        chartLabels: { display: true },
      },
    },
  });
}

function renderLine(canvasId, rows) {
  renderChart(canvasId, {
    type: "line",
    data: {
      labels: rows.map((row) => row.name),
      datasets: [{
        label: "提交人数",
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
      layout: { padding: { top: 24, right: 14, bottom: 8, left: 14 } },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (context) => `${context.raw.toLocaleString()} 人` } },
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

function baseChartOptions(extra = {}) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: extra.indexAxis || "x",
    layout: { padding: extra.indexAxis === "y" ? { top: 8, right: 56, bottom: 8, left: 8 } : { top: 26, right: 8, bottom: 8, left: 8 } },
    plugins: {
      legend: { display: false },
      tooltip: { callbacks: { label: (context) => `${context.raw.toLocaleString()} 人` } },
      chartLabels: { display: true },
    },
    scales: {
      x: { beginAtZero: true, ticks: { precision: 0 } },
      y: { beginAtZero: true, ticks: { precision: 0 } },
    },
  };
}

function buildSummaries(people, projectRows) {
  const countries = enrichSummary(countBy(people, (person) => person.country || "未填写"), people.length);
  const projectTypes = enrichSummary(countProjectValues(projectRows, "Project Types"), people.length);
  const domains = enrichSummary(countTokenValues(people, (person) => person.domain), people.length);
  const majors = enrichSummary(countTokenValues(people, (person) => person.major || person.domain), people.length);
  const languages = enrichSummary(countTokenValues(people, (person) => person.languageCountry), people.length);
  const details = people.map((person) => ({
    name: person.name || "未填写",
    email: person.email || "未填写",
    country: person.country || "未填写",
    english: person.english || "未填写",
    degree: normalizeDegree(person.degree) || "未填写",
    major: person.major || "未填写",
    hours: person.hours || "未填写",
    projectCount: person.projects.length || parseProjectCount(person.projectCount) || 0,
    pm: person.pm || "未填写",
  }));
  return { countries, projectTypes, domains, majors, languages, details };
}

function renderSummaryTable() {
  const table = el("summaryTable");
  const rows = state.summaries[state.currentSummary] || [];
  const isDetails = state.currentSummary === "details";
  const headers = isDetails
    ? ["姓名", "邮箱", "国家/地区", "英语水平", "学历", "专业", "工作时长", "项目数", "PM"]
    : ["分类", "人数", "占比"];

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
      ? [row.name, row.email, row.country, row.english, row.degree, row.major, row.hours, row.projectCount, row.pm]
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
    el("copyTable").textContent = "已复制";
    setTimeout(() => {
      el("copyTable").textContent = "复制";
    }, 1200);
  });
}

function downloadCurrentCsv() {
  const names = {
    countries: "country-summary",
    projectTypes: "project-type-summary",
    domains: "domain-summary",
    majors: "major-summary",
    languages: "language-summary",
    details: "talent-details",
  };
  downloadText(`${names[state.currentSummary] || "summary"}.csv`, currentSummaryCsv(), "text/csv;charset=utf-8");
}

function currentSummaryCsv() {
  const rows = state.summaries[state.currentSummary] || [];
  const isDetails = state.currentSummary === "details";
  const headers = isDetails
    ? ["Name", "Email", "Country", "English Level", "Degree", "Major", "Working Hours", "Project Count", "PM"]
    : ["Category", "People", "Share"];
  const lines = [headers, ...rows.map((row) => isDetails
    ? [row.name, row.email, row.country, row.english, row.degree, row.major, row.hours, row.projectCount, row.pm]
    : [row.name, row.count, row.share])];
  return lines.map((line) => line.map(csvEscape).join(",")).join("\n");
}

function downloadMarkdownReport() {
  downloadText("talent-pool-analysis-report.md", buildMarkdownReport(), "text/markdown;charset=utf-8");
}

function copyReportForFeishu() {
  const report = buildMarkdownReport();
  navigator.clipboard.writeText(report).then(() => {
    el("copyReport").textContent = "已复制";
    setTimeout(() => {
      el("copyReport").textContent = "复制报告";
    }, 1400);
  }).catch(() => {
    alert("复制失败，请确认浏览器允许剪贴板权限。");
  });
}

function downloadHtmlReport() {
  downloadText("talent-pool-analysis-report.html", buildHtmlReport(), "text/html;charset=utf-8");
}

function buildMarkdownReport() {
  const people = getFilteredPeople();
  const projectRows = people.flatMap((person) => person.projects.map((project) => ({ person, project })));
  const sections = buildReportSections(people, projectRows);
  const lines = [
    "# Talent Pool 人力画像分析报告",
    "",
    `- 来源文件：${state.sourceName}`,
    `- 统计口径：按 Contact Email 去重统计人数`,
    `- 当前筛选人数：${people.length.toLocaleString()}`,
    `- 有项目经验人数：${people.filter((person) => person.projects.length > 0).length.toLocaleString()}`,
    `- 国家/地区覆盖：${countBy(people, (person) => person.country || "未填写").length.toLocaleString()}`,
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
    ["当前筛选人数", people.length.toLocaleString()],
    ["有项目经验人数", experienced.toLocaleString()],
    ["国家/地区覆盖", countBy(people, (person) => person.country || "未填写").length.toLocaleString()],
    ["高英语能力", people.filter((person) => isHighEnglish(person.english)).length.toLocaleString()],
  ];

  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <title>Talent Pool 人力画像分析报告</title>
  <style>
    body { margin: 0; color: #17202f; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif; background: #f5f7fb; line-height: 1.55; }
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
    <h1>Talent Pool 人力画像分析报告</h1>
    <div class="meta">来源文件：${escapeHtml(state.sourceName)} · 统计口径：按 Contact Email 去重统计人数</div>
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
      title: "工作时长分布",
      chartId: "hoursChart",
      rows: enrichSummary(bucketCounts(people, (person) => person.hours, hourBucketFromText), total),
    },
    {
      title: "项目数量",
      chartId: "projectsChart",
      rows: enrichSummary(projectCountBuckets(people), total),
    },
    {
      title: "英语水平",
      chartId: "englishChart",
      rows: enrichSummary(topCounts(countBy(people, (person) => clean(person.english) || "未填写"), 8), total),
    },
    {
      title: "提交时间趋势",
      chartId: "timelineChart",
      rows: enrichSummary(timelineCounts(people), total),
    },
    {
      title: "项目类型 Top 12",
      chartId: "projectTypeChart",
      rows: enrichSummary(topCounts(countProjectValues(projectRows, "Project Types"), 12), total),
    },
    {
      title: "学历结构",
      chartId: "degreeChart",
      rows: enrichSummary(topCounts(countBy(people, (person) => normalizeDegree(person.degree) || "未填写"), 8), total),
    },
    {
      title: "国家/地区 Top 10",
      chartId: "countryChart",
      rows: enrichSummary(topCounts(countBy(people, (person) => person.country || "未填写"), 10), total),
    },
    {
      title: "专业/领域 Top 15",
      chartId: "majorChart",
      rows: enrichSummary(topCounts(countTokenValues(people, (person) => person.major || person.domain), 15), total),
    },
    {
      title: "语种-国家 Top 15",
      chartId: "languageChart",
      rows: enrichSummary(topCounts(countTokenValues(people, (person) => person.languageCountry), 15), total),
    },
  ];
}

function chartImageHtml(chartId) {
  const chart = state.charts[chartId];
  if (!chart) return "";
  return `<img src="${chart.toBase64Image("image/png", 1)}" alt="${escapeHtml(chartId)} 图表" />`;
}

function markdownTable(rows) {
  const tableRows = rows.map((row) => `| ${row.name} | ${row.count} | ${row.share} |`);
  return ["| 分类 | 人数 | 占比 |", "| --- | ---: | ---: |", ...tableRows].join("\n");
}

function htmlTable(rows) {
  return `<table><thead><tr><th>分类</th><th>人数</th><th>占比</th></tr></thead><tbody>${rows.map((row) =>
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
    const name = clean(getName(item)) || "未填写";
    counts.set(name, (counts.get(name) || 0) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).sort(sortCountThenName);
}

function countTokenValues(items, getValue) {
  const counts = new Map();
  items.forEach((item) => {
    const tokens = splitMultiValue(getValue(item));
    const unique = new Set(tokens.length ? tokens : ["未填写"]);
    unique.forEach((token) => counts.set(token, (counts.get(token) || 0) + 1));
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).sort(sortCountThenName);
}

function countProjectValues(projectRows, field) {
  const counts = new Map();
  projectRows.forEach(({ person, project }) => {
    const tokens = splitMultiValue(project[field]).map(normalizeCategory);
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

function bucketCounts(items, getValue, getBucket) {
  const order = ["未填写", "0-4h", "4-8h", "8h+"];
  const counts = new Map(order.map((name) => [name, 0]));
  items.forEach((item) => {
      const bucket = getBucket(getValue(item));
    counts.set(bucket, (counts.get(bucket) || 0) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).filter((row) => row.count > 0);
}

function projectCountBuckets(people) {
  const counts = new Map([
    ["未填写/0", 0],
    ["1 个", 0],
    ["2 个", 0],
    ["3 个", 0],
    ["4-5 个", 0],
    ["6+ 个", 0],
  ]);
  people.forEach((person) => {
    const count = Math.max(person.projects.length, parseProjectCount(person.projectCount));
    const bucket = count <= 0 ? "未填写/0" : count === 1 ? "1 个" : count === 2 ? "2 个" : count === 3 ? "3 个" : count <= 5 ? "4-5 个" : "6+ 个";
    counts.set(bucket, counts.get(bucket) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count })).filter((row) => row.count > 0);
}

function timelineCounts(people) {
  const counts = new Map();
  people.forEach((person) => {
    const date = parseDate(person.completedAt);
    const key = date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}` : "未填写";
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return Array.from(counts, ([name, count]) => ({ name, count }))
    .sort((a, b) => a.name.localeCompare(b.name, "zh-CN"));
}

function enrichSummary(rows, total) {
  return rows.map((row) => ({
    ...row,
    share: percent(row.count, total),
  }));
}

function topCounts(rows, limit) {
  const top = rows.slice(0, limit);
  const rest = rows.slice(limit).reduce((sum, row) => sum + row.count, 0);
  if (rest > 0) top.push({ name: "其他", count: rest });
  return top;
}

function sortCountThenName(a, b) {
  return b.count - a.count || String(a.name).localeCompare(String(b.name), "zh-CN");
}

function splitMultiValue(value) {
  return clean(value)
    .split(/[,，;；|、\n]+/g)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => item.replace(/\s+/g, " "));
}

function normalizeCategory(value) {
  const text = clean(value);
  if (!text) return "";
  const lower = text.toLowerCase();
  const known = {
    annotation: "Annotation",
    collection: "Collection",
    evaluation: "Evaluation",
    coding: "Coding",
    literature: "Literature",
    business: "Business",
    finance: "Finance",
    translation: "Translation",
  };
  return known[lower] || text;
}

function normalizeDegree(value) {
  const text = clean(value);
  const lower = text.toLowerCase();
  if (!text) return "";
  if (lower.includes("phd") || lower.includes("doctor") || text.includes("博士")) return "博士";
  if (lower.includes("master") || text.includes("硕士") || text.includes("研究生")) return "硕士";
  if (lower.includes("bachelor") || text.includes("本科") || text.includes("学士")) return "本科";
  if (lower.includes("associate") || text.includes("大专") || text.includes("专科")) return "大专";
  if (lower.includes("high school") || text.includes("高中")) return "高中及以下";
  if (text.includes("高学历")) return "硕士/博士";
  return text.length > 42 ? `${text.slice(0, 42)}...` : text;
}

function hourBucketFromText(value) {
  const text = clean(value).toLowerCase();
  if (!text) return "未填写";
  if (/less[_\s-]*4|<\s*4/.test(text)) return "0-4h";
  if (/4[_\s-]*8|4\s*-\s*8/.test(text)) return "4-8h";
  if (/more[_\s-]*8|>\s*8/.test(text)) return "8h+";
  const match = text.match(/\d+(\.\d+)?/);
  if (!match) return "未填写";
  const hours = Number(match[0]);
  if (hours < 4) return "0-4h";
  if (hours <= 8) return "4-8h";
  return "8h+";
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
  return Array.from(new Set(values.map(clean).filter(Boolean))).sort((a, b) => a.localeCompare(b, "zh-CN"));
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

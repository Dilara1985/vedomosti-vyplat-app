const roles = [
  { key: "employee", label: "Работник", id: "employeeId", name: "employeeName" },
  { key: "chief", label: "Главный специалист", id: "chiefId", name: "chiefName" },
  { key: "rm", label: "Региональный менеджер", id: "rmId", name: "rmName" },
  { key: "supervisor", label: "Супервайзер", id: "supervisorId", name: "supervisorName" },
  { key: "head", label: "Руководитель СП", id: "headId", name: "headName" },
];

const defaults = {
  employee: [
    { from: 0, rate: 0 },
    { from: 0.8, rate: 0.03 },
    { from: 1, rate: 0.05 },
    { from: 1.2, rate: 0.07 },
  ],
  chief: [
    { from: 0, rate: 0 },
    { from: 0.85, rate: 0.015 },
    { from: 1, rate: 0.025 },
    { from: 1.15, rate: 0.035 },
  ],
  rm: [
    { from: 0, rate: 0 },
    { from: 0.85, rate: 0.01 },
    { from: 1, rate: 0.018 },
    { from: 1.15, rate: 0.025 },
  ],
  supervisor: [
    { from: 0, rate: 0 },
    { from: 0.85, rate: 0.006 },
    { from: 1, rate: 0.012 },
    { from: 1.15, rate: 0.018 },
  ],
  head: [
    { from: 0, rate: 0 },
    { from: 0.85, rate: 0.004 },
    { from: 1, rate: 0.008 },
    { from: 1.15, rate: 0.012 },
  ],
};

let rawRows = [];
let normalizedRows = [];
let calculated = null;
let rates = structuredClone(defaults);

const columnMap = {
  period: ["период", "месяц", "дата", "period", "month"],
  employeeId: ["id работника", "код работника", "табельный", "таб номер", "employee id", "worker id"],
  employeeName: ["работник", "фио работника", "оператор", "сотрудник", "продавец", "employee", "operator"],
  region: ["регион", "область", "город", "region"],
  channel: ["канал", "канал продаж", "sales channel", "channel"],
  chiefId: ["id глав", "код глав", "id главного", "chief id"],
  chiefName: ["главный специалист", "глав спец", "глав. специалист", "гс", "chief"],
  rmId: ["id рм", "код рм", "rm id"],
  rmName: ["региональный менеджер", "рм", "rm", "regional manager"],
  supervisorId: ["id супервайзера", "код супервайзера", "supervisor id"],
  supervisorName: ["супервайзер", "supervisor"],
  headId: ["id руководителя сп", "код руководителя сп", "head id"],
  headName: ["руководитель сп", "руководитель", "head"],
  plan: ["план", "план продаж", "plan", "target"],
  fact: ["факт", "факт продаж", "продажи", "выручка", "fact", "actual", "sales"],
  outflowPlan: ["план оттока", "план отток", "отток план", "plan outflow"],
  outflowFact: ["факт оттока", "факт отток", "отток факт", "actual outflow"],
  bonusReduction: ["снижение бонусов", "снижение бонуса", "классификация фрода", "фрод", "fraud"],
  kpi: ["kpi", "kpi %", "ключевой показатель"],
  broadbandPlan: ["план шпд", "шпд план", "план broadband"],
  broadbandFact: ["факт шпд", "шпд факт", "факт broadband"],
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  exportBtn: document.querySelector("#exportBtn"),
  emptyState: document.querySelector("#emptyState"),
  content: document.querySelector("#content"),
  periodFilter: document.querySelector("#periodFilter"),
  regionFilter: document.querySelector("#regionFilter"),
  channelFilter: document.querySelector("#channelFilter"),
  ratesEditor: document.querySelector("#ratesEditor"),
};

renderRatesEditor();

els.fileInput.addEventListener("change", async (event) => {
  const files = [...event.target.files];
  rawRows = [];
  normalizedRows = [];
  calculated = null;

  for (const file of files) {
    const rows = await readWorkbook(file);
    rawRows.push(...rows.map((row) => ({ ...row, _sourceFile: file.name })));
  }

  normalizedRows = normalizeRows(rawRows);
  rebuild();
});

for (const select of [els.periodFilter, els.regionFilter, els.channelFilter]) {
  select.addEventListener("change", rebuild);
}

els.exportBtn.addEventListener("click", exportWorkbook);

document.querySelectorAll(".tab").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((tab) => tab.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
    button.classList.add("active");
    document.querySelector(`#${button.dataset.tab}Tab`).classList.add("active");
  });
});

async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const rows = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    json.forEach((row) => rows.push({ ...row, _sheet: sheetName }));
  }

  return rows;
}

function normalizeRows(rows) {
  const normalized = rows
    .map((row, index) => normalizeRow(row, index))
    .filter((row) => row.employeeName || row.employeeId || row.fact || row.plan);

  const hierarchy = new Map();
  for (const row of normalized) {
    const hasHierarchy =
      row.region !== "Не указан" ||
      row.channel !== "Не указан" ||
      row.chiefName !== "Не указан" ||
      row.rmName !== "Не указан" ||
      row.supervisorName !== "Не указан" ||
      row.headName !== "Не указан";
    if (!hasHierarchy) continue;
    for (const key of rowKeys(row)) {
      if (!hierarchy.has(key)) hierarchy.set(key, row);
    }
  }

  const factRows = normalized.filter((row) => row.plan || row.fact);
  const rowsToCalculate = factRows.length ? factRows : normalized;

  return rowsToCalculate.map((row) => {
    const match = rowKeys(row).map((key) => hierarchy.get(key)).find(Boolean);
    if (!match || match === row) return row;
    const merged = { ...row };
    for (const key of [
      "region",
      "channel",
      "chiefId",
      "chiefName",
      "rmId",
      "rmName",
      "supervisorId",
      "supervisorName",
      "headId",
      "headName",
    ]) {
      if (!merged[key] || merged[key] === "Не указан" || merged[key] === "missing") merged[key] = match[key];
    }
    return merged;
  });
}

function normalizeRow(row, index) {
  const normalized = { rowNo: index + 1, sourceFile: row._sourceFile, sheet: row._sheet };
  for (const key of Object.keys(columnMap)) {
    normalized[key] = getByAliases(row, columnMap[key]);
  }

  normalized.period = normalized.period || "Без периода";
  normalized.employeeName = text(normalized.employeeName);
  normalized.employeeId = text(normalized.employeeId) || slug(normalized.employeeName);
  normalized.region = text(normalized.region) || "Не указан";
  normalized.channel = text(normalized.channel) || "Не указан";
  normalized.chiefName = text(normalized.chiefName) || "Не указан";
  normalized.chiefId = text(normalized.chiefId) || slug(normalized.chiefName);
  normalized.rmName = text(normalized.rmName) || "Не указан";
  normalized.rmId = text(normalized.rmId) || slug(normalized.rmName);
  normalized.supervisorName = text(normalized.supervisorName) || "Не указан";
  normalized.supervisorId = text(normalized.supervisorId) || slug(normalized.supervisorName);
  normalized.headName = text(normalized.headName) || "Не указан";
  normalized.headId = text(normalized.headId) || slug(normalized.headName);
  normalized.plan = money(normalized.plan);
  normalized.fact = money(normalized.fact);
  normalized.outflowPlan = money(normalized.outflowPlan);
  normalized.outflowFact = money(normalized.outflowFact);
  normalized.bonusReduction = money(normalized.bonusReduction);
  normalized.kpi = percentOrNumber(normalized.kpi);
  normalized.broadbandPlan = money(normalized.broadbandPlan);
  normalized.broadbandFact = money(normalized.broadbandFact);
  normalized.achievement = normalized.plan ? normalized.fact / normalized.plan : 0;
  normalized.rate = rateFor("employee", normalized.achievement);
  normalized.payout = normalized.fact * normalized.rate;
  return normalized;
}

function getByAliases(row, aliases) {
  const entries = Object.entries(row);
  const found = entries.find(([name]) => {
    const cleaned = cleanHeader(name);
    return aliases.some((alias) => cleaned.includes(cleanHeader(alias)));
  });
  return found ? found[1] : "";
}

function cleanHeader(value) {
  return String(value).toLowerCase().replace(/[._-]/g, " ").replace(/\s+/g, " ").trim();
}

function text(value) {
  return String(value ?? "").trim();
}

function slug(value) {
  const cleaned = text(value).toLowerCase().replace(/[^a-zа-я0-9]+/gi, "-").replace(/^-|-$/g, "");
  return cleaned || "missing";
}

function rowKeys(row) {
  return [row.employeeId, row.employeeName && slug(row.employeeName)].filter(Boolean);
}

function money(value) {
  if (typeof value === "number") return value;
  const cleaned = String(value ?? "")
    .replace(/\s/g, "")
    .replace(/,/g, ".")
    .replace(/[^\d.-]/g, "");
  return Number(cleaned) || 0;
}

function percentOrNumber(value) {
  if (typeof value === "number") return value > 1 ? value / 100 : value;
  const raw = String(value ?? "").trim();
  if (!raw) return 0;
  const number = money(raw);
  return raw.includes("%") || number > 1 ? number / 100 : number;
}

function rebuild() {
  if (!normalizedRows.length) return;
  fillFilters();
  const rows = applyFilters(normalizedRows);
  calculated = calculate(rows);
  render(calculated);
  els.emptyState.classList.add("hidden");
  els.content.classList.remove("hidden");
  els.exportBtn.disabled = false;
}

function fillFilters() {
  fillSelect(els.periodFilter, unique(normalizedRows.map((row) => row.period)));
  fillSelect(els.regionFilter, unique(normalizedRows.map((row) => row.region)));
  fillSelect(els.channelFilter, unique(normalizedRows.map((row) => row.channel)));
}

function fillSelect(select, values) {
  const current = select.value || "Все";
  const next = ["Все", ...values.filter(Boolean).sort((a, b) => a.localeCompare(b, "ru"))];
  select.innerHTML = next.map((value) => `<option>${escapeHtml(value)}</option>`).join("");
  select.value = next.includes(current) ? current : "Все";
}

function applyFilters(rows) {
  return rows.filter((row) => {
    return (
      matchFilter(row.period, els.periodFilter.value) &&
      matchFilter(row.region, els.regionFilter.value) &&
      matchFilter(row.channel, els.channelFilter.value)
    );
  });
}

function matchFilter(value, filter) {
  return !filter || filter === "Все" || value === filter;
}

function calculate(rows) {
  const employees = rows.map((row) => ({
    role: "Работник",
    period: row.period,
    id: row.employeeId,
    name: row.employeeName || row.employeeId,
    region: row.region,
    channel: row.channel,
    plan: row.plan,
    fact: row.fact,
    outflowPlan: row.outflowPlan,
    outflowFact: row.outflowFact,
    bonusReduction: row.bonusReduction,
    kpi: row.kpi,
    broadbandPlan: row.broadbandPlan,
    broadbandFact: row.broadbandFact,
    achievement: row.achievement,
    rate: row.rate,
    payout: row.payout,
    count: 1,
  }));

  const leaders = [
    ...aggregate(rows, roles[1]),
    ...aggregate(rows, roles[2]),
    ...aggregate(rows, roles[3]),
    ...aggregate(rows, roles[4]),
  ];

  const statement = [...employees, ...leaders];
  const totalPlan = sum(rows, "plan");
  const totalFact = sum(rows, "fact");
  const totalPayout = sum(statement, "payout");

  return {
    rows,
    employees,
    leaders,
    statement,
    byRegion: groupSimple(rows, "region"),
    byChannel: groupSimple(rows, "channel"),
    totals: {
      plan: totalPlan,
      fact: totalFact,
      achievement: totalPlan ? totalFact / totalPlan : 0,
      payout: totalPayout,
    },
    issues: findIssues(rows),
  };
}

function aggregate(rows, role) {
  const map = new Map();
  for (const row of rows) {
    const id = row[role.id] || "missing";
    const key = `${role.key}|${id}|${row.period}`;
    if (!map.has(key)) {
      map.set(key, {
        role: role.label,
        period: row.period,
        id,
        name: row[role.name] || id,
        region: row.region,
        channel: "Все каналы",
        plan: 0,
        fact: 0,
        outflowPlan: 0,
        outflowFact: 0,
        bonusReduction: 0,
        kpi: 0,
        broadbandPlan: 0,
        broadbandFact: 0,
        achievement: 0,
        rate: 0,
        payout: 0,
        count: 0,
      });
    }
    const item = map.get(key);
    item.plan += row.plan;
    item.fact += row.fact;
    item.outflowPlan += row.outflowPlan;
    item.outflowFact += row.outflowFact;
    item.bonusReduction += row.bonusReduction;
    item.kpi += row.kpi;
    item.broadbandPlan += row.broadbandPlan;
    item.broadbandFact += row.broadbandFact;
    item.count += 1;
  }

  return [...map.values()].map((item) => {
    item.achievement = item.plan ? item.fact / item.plan : 0;
    item.kpi = item.count ? item.kpi / item.count : 0;
    item.rate = rateFor(role.key, item.achievement);
    item.payout = item.fact * item.rate;
    return item;
  });
}

function groupSimple(rows, field) {
  const map = new Map();
  for (const row of rows) {
    const key = row[field] || "Не указан";
    if (!map.has(key)) map.set(key, { name: key, plan: 0, fact: 0, achievement: 0 });
    const item = map.get(key);
    item.plan += row.plan;
    item.fact += row.fact;
  }
  return [...map.values()].map((item) => {
    item.achievement = item.plan ? item.fact / item.plan : 0;
    return item;
  });
}

function rateFor(roleKey, achievement) {
  return [...rates[roleKey]].sort((a, b) => a.from - b.from).reduce((rate, tier) => {
    return achievement >= tier.from ? tier.rate : rate;
  }, 0);
}

function sum(rows, field) {
  return rows.reduce((total, row) => total + (Number(row[field]) || 0), 0);
}

function findIssues(rows) {
  const issues = [];
  const missingFact = rows.filter((row) => !row.fact).length;
  const missingPlan = rows.filter((row) => !row.plan).length;
  const missingLeader = rows.filter((row) => row.chiefName === "Не указан" || row.rmName === "Не указан").length;

  if (missingFact) issues.push(`${missingFact} строк без факта. Они попадут в ведомость с нулевой выплатой.`);
  if (missingPlan) issues.push(`${missingPlan} строк без плана. Процент выполнения будет равен 0%.`);
  if (missingLeader) issues.push(`${missingLeader} строк без главного специалиста или регионального менеджера.`);
  if (!issues.length) issues.push("Ошибок в загруженных данных не найдено.");
  return issues;
}

function render(data) {
  document.querySelector("#metricPlan").textContent = fmtMoney(data.totals.plan);
  document.querySelector("#metricFact").textContent = fmtMoney(data.totals.fact);
  document.querySelector("#metricAch").textContent = fmtPct(data.totals.achievement);
  document.querySelector("#metricPayout").textContent = fmtMoney(data.totals.payout);

  renderTable("#regionTable", data.byRegion, ["Регион", "План", "Факт", "Выполнение"], (row) => [
    row.name,
    fmtMoney(row.plan),
    fmtMoney(row.fact),
    pctCell(row.achievement),
  ]);
  renderTable("#channelTable", data.byChannel, ["Канал", "План", "Факт", "Выполнение"], (row) => [
    row.name,
    fmtMoney(row.plan),
    fmtMoney(row.fact),
    pctCell(row.achievement),
  ]);
  renderTable("#employeeTable", data.employees, headers(), rowCells);
  renderTable("#leaderTable", data.leaders, headers(true), rowCells);
  renderTable("#statementTable", data.statement, headers(true), rowCells);
  renderIssues(data.issues);
}

function headers(withCount = false) {
  return withCount
    ? ["Роль", "Период", "ID", "ФИО/группа", "Кол-во", "Регион", "Канал", "План", "Факт", "Выполнение", "Ставка", "Выплата"]
    : ["Роль", "Период", "ID", "ФИО", "Регион", "Канал", "План", "Факт", "Выполнение", "Ставка", "Выплата"];
}

function rowCells(row) {
  const base = [row.role, row.period, row.id, row.name];
  if (row.role !== "Работник") base.push(String(row.count));
  return [
    ...base,
    row.region,
    row.channel,
    fmtMoney(row.plan),
    fmtMoney(row.fact),
    pctCell(row.achievement),
    fmtPct(row.rate),
    fmtMoney(row.payout),
  ];
}

function renderTable(selector, rows, tableHeaders, cellBuilder) {
  const table = document.querySelector(selector);
  const body = rows.map((row) => {
    const cells = cellBuilder(row);
    return `<tr>${cells.map((cell, index) => `<td class="${index >= cells.length - 5 ? "num" : ""}">${cell}</td>`).join("")}</tr>`;
  });
  table.innerHTML = `
    <thead><tr>${tableHeaders.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>${body.join("")}</tbody>
  `;
}

function renderIssues(issues) {
  document.querySelector("#issuesList").innerHTML = issues
    .map((issue) => `<div class="issue ${issue.startsWith("Ошибок") ? "good" : ""}">${escapeHtml(issue)}</div>`)
    .join("");
}

function renderRatesEditor() {
  els.ratesEditor.innerHTML = roles
    .flatMap((role) =>
      rates[role.key].slice(1).map((tier, index) => {
        const id = `${role.key}_${index}`;
        return `
          <label class="rate-row">
            <span>${role.label}: от ${fmtPct(tier.from)}</span>
            <input class="rate-input" id="${id}" type="number" step="0.1" min="0" value="${tier.rate * 100}" />
          </label>
        `;
      }),
    )
    .join("");

  roles.forEach((role) => {
    rates[role.key].slice(1).forEach((tier, index) => {
      document.querySelector(`#${role.key}_${index}`).addEventListener("input", (event) => {
        tier.rate = Number(event.target.value) / 100 || 0;
        if (normalizedRows.length) rebuild();
      });
    });
  });
}

function exportWorkbook() {
  if (!calculated) return;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, buildStatementSheet(), "Ведомость");
  XLSX.utils.book_append_sheet(workbook, buildFlatSheet(calculated.employees), "Работники");
  XLSX.utils.book_append_sheet(workbook, buildFlatSheet(calculated.leaders), "Руководители");
  XLSX.utils.book_append_sheet(workbook, buildSummarySheet("Регион", calculated.byRegion), "По регионам");
  XLSX.utils.book_append_sheet(workbook, buildSummarySheet("Канал продаж", calculated.byChannel), "По каналам");
  XLSX.utils.book_append_sheet(workbook, buildRatesSheet(), "Ставки");

  XLSX.writeFile(workbook, "Ведомость_выплат.xlsx");
}

function buildStatementSheet() {
  const selectedPeriod = els.periodFilter.value && els.periodFilter.value !== "Все" ? els.periodFilter.value : "Все периоды";
  const aoa = [];
  const merges = [];
  const rowTypes = [];
  const heads = groupBy(calculated.rows, "headName");

  for (const [headName, headRows] of heads.entries()) {
    const headSummary = summarizeTemplateRows(headRows, "head");
    addLeaderTemplateBlock(aoa, merges, rowTypes, "Руководитель ЦРК", headName, selectedPeriod, headSummary, false);
    aoa.push([]);
    rowTypes.push("gap");

    const chiefs = groupBy(headRows, "chiefName");
    for (const [chiefName, chiefRows] of chiefs.entries()) {
      addLeaderTemplateBlock(aoa, merges, rowTypes, "Главный специалист", chiefName, selectedPeriod, summarizeTemplateRows(chiefRows, "chief"), true, chiefRows);
      aoa.push([]);
      rowTypes.push("gap");
    }
  }

  aoa.push(["Согласовано", "", "", "Проверено", "", "", "Утверждено", "", ""]);
  rowTypes.push("signature");
  aoa.push(["______________", "", "", "______________", "", "", "______________", "", ""]);
  rowTypes.push("signature");

  const sheet = XLSX.utils.aoa_to_sheet(aoa);
  sheet["!cols"] = [
    { wch: 22 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 24 },
    { wch: 9 },
    { wch: 14 },
    { wch: 14 },
  ];
  sheet["!merges"] = merges;
  sheet["!freeze"] = { xSplit: 0, ySplit: 0 };
  styleTemplateSheet(sheet, rowTypes, aoa.length);
  return sheet;
}

function addLeaderTemplateBlock(aoa, merges, rowTypes, title, personName, period, summary, includeWorkers, workers = []) {
  const start = aoa.length;
  aoa.push([title, "", "", "", "", "", "", "", ""]);
  rowTypes.push("purpleTitle");
  merges.push({ s: { r: start, c: 0 }, e: { r: start, c: 8 } });

  aoa.push([personName || "ФИО", "", "", "", "", "", "", "", ""]);
  rowTypes.push("purpleName");
  merges.push({ s: { r: start + 1, c: 0 }, e: { r: start + 1, c: 8 } });

  aoa.push([
    includeWorkers ? "Работники" : "",
    "Факт продаж",
    "План продаж",
    "Факт оттока",
    "План оттока",
    "Снижение бонусов согласно классификации ФРОДа",
    "KPI",
    "Факт ШПД",
    "План ШПД",
  ]);
  rowTypes.push("header");

  aoa.push([period, "", "", "", "", "", "", "", ""]);
  rowTypes.push("month");
  merges.push({ s: { r: start + 3, c: 0 }, e: { r: start + 3, c: 8 } });

  if (includeWorkers) {
    workers.forEach((worker, index) => {
      aoa.push([
        worker.employeeName || `работник №${index + 1}`,
        Math.round(worker.fact || 0),
        Math.round(worker.plan || 0),
        emptyOrNumber(worker.outflowFact),
        emptyOrNumber(worker.outflowPlan),
        emptyOrNumber(worker.bonusReduction),
        emptyOrPercent(worker.kpi),
        emptyOrNumber(worker.broadbandFact),
        emptyOrNumber(worker.broadbandPlan),
      ]);
      rowTypes.push("worker");
    });
  }

  aoa.push(["Общая сумма", ...templateMetricValues(summary)]);
  rowTypes.push("total");
  aoa.push(["% выполнения", ...templatePercentValues(summary)]);
  rowTypes.push("total");
  aoa.push(["Модификатор", summary.rate || 0, "", "", "", "", "", "", ""]);
  rowTypes.push("total");
  aoa.push(["ИТОГ", Math.round(summary.payout || 0), "", "", "", "", "", "", ""]);
  rowTypes.push("grandTotal");
}

function summarizeTemplateRows(rows, roleKey) {
  const plan = sum(rows, "plan");
  const fact = sum(rows, "fact");
  return {
    plan,
    fact,
    outflowPlan: sum(rows, "outflowPlan"),
    outflowFact: sum(rows, "outflowFact"),
    bonusReduction: sum(rows, "bonusReduction"),
    kpi: rows.length ? sum(rows, "kpi") / rows.length : 0,
    broadbandPlan: sum(rows, "broadbandPlan"),
    broadbandFact: sum(rows, "broadbandFact"),
    achievement: plan ? fact / plan : 0,
    rate: rateFor(roleKey, plan ? fact / plan : 0),
    payout: fact * rateFor(roleKey, plan ? fact / plan : 0),
  };
}

function templateMetricValues(summary) {
  return [
    Math.round(summary.fact || 0),
    Math.round(summary.plan || 0),
    emptyOrNumber(summary.outflowFact),
    emptyOrNumber(summary.outflowPlan),
    emptyOrNumber(summary.bonusReduction),
    emptyOrPercent(summary.kpi),
    emptyOrNumber(summary.broadbandFact),
    emptyOrNumber(summary.broadbandPlan),
  ];
}

function templatePercentValues(summary) {
  return [
    summary.plan ? summary.fact / summary.plan : 0,
    "",
    summary.outflowPlan ? summary.outflowFact / summary.outflowPlan : "",
    "",
    "",
    emptyOrPercent(summary.kpi),
    summary.broadbandPlan ? summary.broadbandFact / summary.broadbandPlan : "",
    "",
  ];
}

function emptyOrNumber(value) {
  return Number(value) ? Math.round(value) : "";
}

function emptyOrPercent(value) {
  return Number(value) ? value : "";
}

function buildFlatSheet(rows) {
  const data = rows.map((row) => ({
    Роль: row.role || "",
    Период: row.period || "",
    ID: row.id || "",
    "ФИО/группа": row.name,
    Регион: row.region || "",
    "Канал продаж": row.channel || "",
    "Кол-во работников": row.count || "",
    План: Math.round(row.plan || 0),
    Факт: Math.round(row.fact || 0),
    Выполнение: row.achievement || 0,
    "Ставка выплаты": row.rate || 0,
    Выплата: Math.round(row.payout || 0),
  }));
  const sheet = XLSX.utils.json_to_sheet(data);
  sheet["!cols"] = [
    { wch: 24 },
    { wch: 12 },
    { wch: 16 },
    { wch: 28 },
    { wch: 16 },
    { wch: 18 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
  ];
  applyNumberFormats(sheet, rows.length + 1);
  return sheet;
}

function buildSummarySheet(label, rows) {
  const data = rows.map((row) => ({
    [label]: row.name,
    План: Math.round(row.plan || 0),
    Факт: Math.round(row.fact || 0),
    Выполнение: row.achievement || 0,
  }));
  const sheet = XLSX.utils.json_to_sheet(data);
  sheet["!cols"] = [{ wch: 22 }, { wch: 14 }, { wch: 14 }, { wch: 12 }];
  applyNumberFormats(sheet, rows.length + 1);
  return sheet;
}

function buildRatesSheet() {
  const data = roles.flatMap((role) =>
    rates[role.key].map((tier) => ({
      Роль: role.label,
      "Порог выполнения": tier.from,
      "Процент выплаты от факта": tier.rate,
    })),
  );
  const sheet = XLSX.utils.json_to_sheet(data);
  sheet["!cols"] = [{ wch: 26 }, { wch: 18 }, { wch: 24 }];
  applyNumberFormats(sheet, data.length + 1);
  return sheet;
}

function applyNumberFormats(sheet, rowCount) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let r = 1; r <= rowCount; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || typeof cell.v !== "number") continue;
      const header = String(sheet[XLSX.utils.encode_cell({ r: Math.max(r - 1, 0), c })]?.v || "");
      if (header.includes("Выполнение") || header.includes("Ставка") || header.includes("Порог") || header.includes("Процент")) {
        cell.z = "0.0%";
      } else {
        cell.z = "#,##0";
      }
    }
  }
}

function styleTemplateSheet(sheet, rowTypes, rowCount) {
  const border = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
  };
  const purple = { fgColor: { rgb: "CCC0DA" } };
  const white = { fgColor: { rgb: "FFFFFF" } };

  for (let r = 0; r < rowCount; r += 1) {
    const rowType = rowTypes[r] || "";
    for (let c = 0; c < 9; c += 1) {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (!sheet[ref]) sheet[ref] = { t: "s", v: "" };
      const cell = sheet[ref];
      cell.s = {
        border,
        fill: rowType === "purpleTitle" || rowType === "purpleName" || rowType === "month" ? purple : white,
        font: {
          name: "Times New Roman",
          sz: rowType === "header" ? 11 : 10,
          bold: ["purpleTitle", "purpleName", "header", "month", "total", "grandTotal"].includes(rowType),
          color: { rgb: ["purpleTitle", "purpleName", "month", "worker"].includes(rowType) ? "FF0000" : "000000" },
        },
        alignment: {
          horizontal: ["purpleTitle", "purpleName", "month", "header"].includes(rowType) ? "center" : c === 0 ? "left" : "center",
          vertical: "center",
          wrapText: true,
        },
      };

      if (rowType === "gap") {
        cell.s.border = {};
        cell.s.fill = white;
      }

      if (typeof cell.v === "number") {
        if (rowType === "total" && c === 6) cell.z = "0.0%";
        else if (rowType === "total" && c === 1 && String(sheet[XLSX.utils.encode_cell({ r, c: 0 })]?.v).includes("Модификатор")) cell.z = "0.0%";
        else if (String(sheet[XLSX.utils.encode_cell({ r, c: 0 })]?.v).includes("%")) cell.z = "0.0%";
        else cell.z = "#,##0";
      }
    }
  }

  sheet["!rows"] = rowTypes.map((type) => {
    if (type === "header") return { hpt: 58 };
    if (type === "purpleTitle" || type === "purpleName" || type === "month") return { hpt: 18 };
    if (type === "gap") return { hpt: 18 };
    return { hpt: 16 };
  });
}

function groupBy(rows, field) {
  const map = new Map();
  for (const row of rows) {
    const key = row[field] || "Не указан";
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  return map;
}

function unique(values) {
  return [...new Set(values.map(text).filter(Boolean))];
}

function fmtMoney(value) {
  return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(value || 0);
}

function fmtPct(value) {
  return new Intl.NumberFormat("ru-RU", { style: "percent", maximumFractionDigits: 1 }).format(value || 0);
}

function pctCell(value) {
  const cls = value >= 1 ? "status-ok" : value >= 0.8 ? "status-warn" : "status-bad";
  return `<span class="${cls}">${fmtPct(value)}</span>`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

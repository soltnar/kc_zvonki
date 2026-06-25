const DEFAULT_OPERATOR_MAP = `333 = Власова Анна
317 = Заботкина Дарья
319 = Запорожченко Арина
312 = Кукина Анастасия
315 = Лялина Светлана
305 = Сизова Кристина
304 = Смирнова Ксения
334 = Пазилина Самира
332 = Тарасова Полина
335 = Трифонов Дмитрий
316 = Юдина Екатерина`;

const DOW = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
const MONTHS = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"];
let lastStats = null;

const els = {
  libraryStatus: document.getElementById("libraryStatus"),
  mainFile: document.getElementById("mainFile"),
  sbisFile: document.getElementById("sbisFile"),
  analyzeBtn: document.getElementById("analyzeBtn"),
  operatorMap: document.getElementById("operatorMap"),
  occupancy: document.getElementById("occupancy"),
  buffer: document.getElementById("buffer"),
  replaceSbis: document.getElementById("replaceSbis"),
  includeOutbound: document.getElementById("includeOutbound"),
  notice: document.getElementById("notice"),
  kpis: document.getElementById("kpis"),
  hourMetric: document.getElementById("hourMetric"),
  hourChart: document.getElementById("hourChart"),
  dowChart: document.getElementById("dowChart"),
  monthChart: document.getElementById("monthChart"),
  staffing: document.getElementById("staffing"),
  operatorSearch: document.getElementById("operatorSearch"),
  operatorTable: document.querySelector("#operatorTable tbody"),
  heatmap: document.getElementById("heatmap")
};

els.operatorMap.value = DEFAULT_OPERATOR_MAP;
els.libraryStatus.textContent = window.XLSX ? "XLSX: готово" : "XLSX: библиотека не загружена";
els.analyzeBtn.disabled = !window.XLSX;

els.analyzeBtn.addEventListener("click", analyze);
els.hourMetric.addEventListener("change", () => lastStats && render(lastStats));
els.operatorSearch.addEventListener("input", () => lastStats && renderOperators(lastStats.operatorRows));
[els.occupancy, els.buffer, els.replaceSbis, els.includeOutbound].forEach((el) => {
  el.addEventListener("change", () => lastStats && analyze());
});

async function analyze() {
  if (!els.mainFile.files[0]) {
    setNotice("Нужен основной файл “История внешних звонков”.", true);
    return;
  }

  try {
    els.analyzeBtn.disabled = true;
    setNotice("Читаю Excel и собираю статистику...");
    const mapping = parseMapping(els.operatorMap.value);
    const mainRows = await parseMainWorkbook(els.mainFile.files[0], mapping);
    const sbisRows = els.sbisFile.files[0] ? await parseSbisWorkbook(els.sbisFile.files[0]) : [];
    const rows = mergeSources(mainRows, sbisRows);
    const stats = buildStats(rows, mainRows, sbisRows);
    lastStats = stats;
    render(stats);
  } catch (error) {
    console.error(error);
    setNotice(`Ошибка: ${error.message}`, true);
  } finally {
    els.analyzeBtn.disabled = !window.XLSX;
  }
}

function parseMapping(text) {
  const byNumber = new Map();
  text.split(/\n+/).forEach((line) => {
    const clean = line.trim();
    if (!clean || !clean.includes("=")) return;
    const [left, right] = clean.split("=").map((v) => v.trim());
    right.split(",").map((v) => v.trim()).filter(Boolean).forEach((name) => {
      byNumber.set(String(left).replace(/\D/g, ""), name);
    });
  });
  return byNumber;
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true, dense: false });
}

async function parseMainWorkbook(file, mapping) {
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  const headerIndex = raw.findIndex((row) => row.includes("Тип звонка") && row.includes("Сотрудник"));
  if (headerIndex < 0) throw new Error("В основном файле не найдена строка заголовков.");
  const headers = raw[headerIndex].map(String);
  return raw.slice(headerIndex + 1).map((row) => {
    const rec = Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""]));
    const date = parseMainDate(rec["Дата"], rec["Время"]);
    if (!date) return null;
    const type = String(rec["Тип звонка"] || "").toLowerCase();
    const rawEmployee = String(rec["Сотрудник"] || "").trim() || "Без оператора";
    return {
      source: "Основная телефония",
      date,
      type,
      direction: type.includes("исход") ? "out" : "in",
      missed: type.includes("пропущ") || type.includes("неуспеш"),
      employee: normalizeMainEmployee(rawEmployee, mapping),
      rawEmployee,
      client: String(rec["Клиент"] || ""),
      waitSec: parseDuration(rec["Ожидание"]),
      talkSec: parseDuration(rec["Длительность"])
    };
  }).filter(Boolean);
}

async function parseSbisWorkbook(file) {
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  const headerIndex = raw.findIndex((row) => row.includes("Начало звонка") && row.includes("Сотрудник"));
  if (headerIndex < 0) throw new Error("В файле СБИС не найдена строка заголовков.");
  const headers = raw[headerIndex].map(String);
  return raw.slice(headerIndex + 1).map((row) => {
    const rec = Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""]));
    const date = parseDateTime(rec["Начало звонка"]);
    if (!date) return null;
    const result = String(rec["Результат звонка"] || "").toLowerCase();
    const directionText = String(rec["Направление звонка"] || "").toLowerCase();
    return {
      source: "СБИС",
      date,
      type: `${directionText} ${result}`.trim(),
      direction: directionText.includes("исход") ? "out" : "in",
      missed: /не отвечен|не дождался|занято|вышло время/.test(result),
      employee: shortName(String(rec["Сотрудник"] || "СБИС без оператора")),
      rawEmployee: String(rec["Сотрудник"] || ""),
      client: "",
      waitSec: 0,
      talkSec: parseDuration(rec["Время разговора"])
    };
  }).filter(Boolean);
}

function mergeSources(mainRows, sbisRows) {
  if (!sbisRows.length) return mainRows;
  if (!els.replaceSbis.checked) return mainRows.concat(sbisRows);
  return mainRows.filter((row) => !/сбис/i.test(row.rawEmployee)).concat(sbisRows);
}

function normalizeMainEmployee(value, mapping) {
  if (/сбис/i.test(value)) return "СБИС";
  const number = value.match(/\b(3\d{2})\b/);
  if (number && mapping.has(number[1])) return mapping.get(number[1]);
  return value.replace(/^(\d{3})\s+/, "").replace(/\s+/g, " ").trim();
}

function shortName(full) {
  const parts = full.trim().split(/\s+/);
  if (parts.length < 2) return full.trim() || "Без оператора";
  return `${parts[0]} ${parts[1]}`;
}

function parseMainDate(dateValue, timeValue) {
  const d = parseDateTime(dateValue);
  if (!d) return null;
  const seconds = parseDuration(timeValue);
  d.setHours(Math.floor(seconds / 3600), Math.floor(seconds % 3600 / 60), seconds % 60, 0);
  return d;
}

function parseDateTime(value) {
  if (value instanceof Date && !Number.isNaN(value)) return new Date(value);
  const text = String(value || "").trim();
  if (!text) return null;
  let m = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = text.match(/^(\d{2})[./](\d{2})[./](\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function parseDuration(value) {
  if (value instanceof Date && !Number.isNaN(value)) {
    return value.getHours() * 3600 + value.getMinutes() * 60 + value.getSeconds();
  }
  if (typeof value === "number") return Math.round(value * 86400);
  const text = String(value || "").trim();
  const m = text.match(/(?:(\d+):)?(\d{1,2}):(\d{2})/);
  if (!m) return 0;
  return +(m[1] || 0) * 3600 + +m[2] * 60 + +m[3];
}

function buildStats(rows, mainRows, sbisRows) {
  const activeDays = new Set(rows.map((r) => isoDate(r.date))).size || 1;
  const inboundRows = rows.filter((r) => r.direction === "in");
  const workloadRows = rows.filter((r) => r.direction === "in" || els.includeOutbound.checked);
  const totalTalk = rows.reduce((sum, r) => sum + r.talkSec, 0);
  const totalWait = rows.reduce((sum, r) => sum + r.waitSec, 0);
  const missed = rows.filter((r) => r.missed).length;
  const operators = groupOperators(rows);
  const hour = aggregateRange(24, (r) => r.date.getHours(), workloadRows);
  const dow = aggregateRange(7, (r) => (r.date.getDay() + 6) % 7, rows);
  const month = aggregateRange(12, (r) => r.date.getMonth(), rows);
  const heat = Array.from({ length: 7 }, () => Array.from({ length: 24 }, () => 0));
  rows.forEach((r) => { heat[(r.date.getDay() + 6) % 7][r.date.getHours()]++; });
  const range = dateRange(rows);

  return {
    rows,
    mainRows,
    sbisRows,
    activeDays,
    total: rows.length,
    inbound: inboundRows.length,
    outbound: rows.filter((r) => r.direction === "out").length,
    missed,
    missedRate: rows.length ? missed / rows.length : 0,
    totalTalk,
    totalWait,
    avgTalk: rows.length ? totalTalk / rows.length : 0,
    operatorRows: operators,
    hour,
    dow,
    month,
    heat,
    range,
    sourceRanges: {
      main: dateRange(mainRows),
      sbis: dateRange(sbisRows)
    }
  };
}

function groupOperators(rows) {
  const map = new Map();
  rows.forEach((r) => {
    if (!map.has(r.employee)) {
      map.set(r.employee, { name: r.employee, total: 0, inbound: 0, outbound: 0, missed: 0, talkSec: 0, waitSec: 0, hours: new Set() });
    }
    const item = map.get(r.employee);
    item.total++;
    if (r.direction === "in") item.inbound++;
    if (r.direction === "out") item.outbound++;
    if (r.missed) item.missed++;
    item.talkSec += r.talkSec;
    item.waitSec += r.waitSec;
    item.hours.add(`${isoDate(r.date)} ${r.date.getHours()}`);
  });
  return [...map.values()].map((item) => ({
    ...item,
    activeHours: item.hours.size,
    avgTalk: item.total ? item.talkSec / item.total : 0,
    missedRate: item.total ? item.missed / item.total : 0
  })).sort((a, b) => b.total - a.total);
}

function aggregateRange(size, keyFn, rows) {
  const arr = Array.from({ length: size }, () => ({ calls: 0, missed: 0, talkSec: 0, waitSec: 0, dates: new Set() }));
  rows.forEach((r) => {
    const item = arr[keyFn(r)];
    item.calls++;
    if (r.missed) item.missed++;
    item.talkSec += r.talkSec;
    item.waitSec += r.waitSec;
    item.dates.add(isoDate(r.date));
  });
  const occupancy = clamp(+els.occupancy.value || 75, 35, 95) / 100;
  const buffer = 1 + clamp(+els.buffer.value || 0, 0, 80) / 100;
  return arr.map((item, index) => {
    const days = item.dates.size || 1;
    const avgCalls = item.calls / days;
    const loadSec = (item.talkSec + item.waitSec) / days;
    return {
      index,
      calls: item.calls,
      missed: item.missed,
      avgCalls,
      loadHours: loadSec / 3600,
      operators: Math.max(1, Math.ceil((loadSec / 3600) / occupancy * buffer))
    };
  });
}

function render(stats) {
  renderNotice(stats);
  renderKpis(stats);
  renderBarChart(els.hourChart, stats.hour, hourLabels(), els.hourMetric.value);
  renderBarChart(els.dowChart, stats.dow, DOW, "calls");
  renderBarChart(els.monthChart, stats.month, MONTHS, "calls");
  renderStaffing(stats.hour);
  renderOperators(stats.operatorRows);
  renderHeatmap(stats.heat);
}

function renderNotice(stats) {
  const parts = [`Период анализа: ${formatDate(stats.range.min)} - ${formatDate(stats.range.max)}. Строк: ${formatNumber(stats.total)}.`];
  if (stats.sbisRows.length) {
    parts.push(`СБИС: ${formatDate(stats.sourceRanges.sbis.min)} - ${formatDate(stats.sourceRanges.sbis.max)}.`);
  }
  if (stats.sbisRows.length && els.replaceSbis.checked && stats.sourceRanges.sbis.min > stats.sourceRanges.main.min) {
    parts.push("Внимание: детализация СБИС покрывает не весь период основного файла.");
    setNotice(parts.join(" "), true);
  } else {
    setNotice(parts.join(" "));
  }
}

function renderKpis(stats) {
  const cards = [
    ["Всего звонков", formatNumber(stats.total), `${formatNumber(stats.activeDays)} дней с данными`],
    ["Входящие", formatNumber(stats.inbound), `${percent(stats.inbound / Math.max(stats.total, 1))} от всех`],
    ["Исходящие", formatNumber(stats.outbound), `${percent(stats.outbound / Math.max(stats.total, 1))} от всех`],
    ["Пропущенные", formatNumber(stats.missed), percent(stats.missedRate)],
    ["Разговор", formatDuration(stats.totalTalk), `среднее ${formatDuration(stats.avgTalk)}`]
  ];
  els.kpis.innerHTML = cards.map(([label, value, sub]) => `<div class="panel kpi"><span>${label}</span><strong>${value}</strong><small>${sub}</small></div>`).join("");
}

function renderOperators(rows) {
  const query = els.operatorSearch.value.trim().toLowerCase();
  const filtered = rows.filter((r) => r.name.toLowerCase().includes(query));
  els.operatorTable.innerHTML = filtered.map((r) => `
    <tr>
      <td>${escapeHtml(r.name)}</td>
      <td>${formatNumber(r.total)}</td>
      <td>${formatNumber(r.inbound)}</td>
      <td>${formatNumber(r.outbound)}</td>
      <td>${formatNumber(r.missed)}</td>
      <td>${percent(r.missedRate)}</td>
      <td>${formatDuration(r.talkSec)}</td>
      <td>${formatDuration(r.avgTalk)}</td>
      <td>${formatNumber(r.activeHours)}</td>
    </tr>
  `).join("");
}

function renderStaffing(hourRows) {
  const max = Math.max(...hourRows.map((r) => r.operators), 1);
  els.staffing.innerHTML = hourRows.map((r) => `
    <div class="staff-row">
      <b>${String(r.index).padStart(2, "0")}:00</b>
      <div class="bar" title="${r.avgCalls.toFixed(1)} звонков в активный день"><i style="width:${r.operators / max * 100}%"></i></div>
      <span>${r.operators} опер.</span>
    </div>
  `).join("");
}

function renderHeatmap(heat) {
  const max = Math.max(...heat.flat(), 1);
  const head = [`<div class="heat-label"></div>`].concat(hourLabels().map((h) => `<div class="heat-label">${h}</div>`));
  const rows = heat.flatMap((day, i) => {
    const cells = [`<div class="heat-label">${DOW[i]}</div>`];
    day.forEach((value) => {
      const alpha = 0.08 + value / max * 0.82;
      cells.push(`<div class="heat-cell" style="background:rgba(37,99,235,${alpha})" title="${value}">${value ? formatCompact(value) : ""}</div>`);
    });
    return cells;
  });
  els.heatmap.innerHTML = head.concat(rows).join("");
}

function renderBarChart(canvas, data, labels, metric) {
  const ctx = canvas.getContext("2d");
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  const width = Math.max(320, rect.width);
  const logicalHeight = Number(canvas.getAttribute("height")) || 240;
  canvas.width = width * dpr;
  canvas.height = logicalHeight * dpr;
  const height = logicalHeight;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, width, height);
  const pad = { top: 18, right: 12, bottom: 34, left: 48 };
  const values = data.map((r) => metric === "load" ? r.loadHours : r[metric]);
  const max = Math.max(...values, 1);
  ctx.strokeStyle = "#d9e0e7";
  ctx.fillStyle = "#657282";
  ctx.font = "12px system-ui";
  for (let i = 0; i <= 4; i++) {
    const y = pad.top + (height - pad.top - pad.bottom) * i / 4;
    ctx.beginPath();
    ctx.moveTo(pad.left, y);
    ctx.lineTo(width - pad.right, y);
    ctx.stroke();
    ctx.fillText(formatCompact(max * (4 - i) / 4), 6, y + 4);
  }
  const gap = 5;
  const barW = (width - pad.left - pad.right - gap * (values.length - 1)) / values.length;
  values.forEach((value, i) => {
    const x = pad.left + i * (barW + gap);
    const h = (height - pad.top - pad.bottom) * value / max;
    const y = height - pad.bottom - h;
    ctx.fillStyle = metric === "missed" ? "#c2410c" : metric === "operators" ? "#0f8f61" : "#2563eb";
    ctx.fillRect(x, y, Math.max(3, barW), h);
    ctx.fillStyle = "#657282";
    if (labels.length <= 12 || i % 2 === 0) ctx.fillText(labels[i], x, height - 12);
  });
}

function setNotice(text, warn = false) {
  els.notice.textContent = text;
  els.notice.classList.toggle("warn", warn);
}

function dateRange(rows) {
  if (!rows.length) return { min: null, max: null };
  return rows.reduce((acc, r) => ({
    min: !acc.min || r.date < acc.min ? r.date : acc.min,
    max: !acc.max || r.date > acc.max ? r.date : acc.max
  }), { min: null, max: null });
}

function hourLabels() {
  return Array.from({ length: 24 }, (_, i) => String(i).padStart(2, "0"));
}

function isoDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDate(date) {
  return date ? date.toLocaleDateString("ru-RU") : "нет данных";
}

function formatNumber(value) {
  return Math.round(value).toLocaleString("ru-RU");
}

function formatCompact(value) {
  if (value >= 1000000) return `${(value / 1000000).toFixed(1)}м`;
  if (value >= 1000) return `${Math.round(value / 1000)}к`;
  return String(Math.round(value));
}

function percent(value) {
  return `${(value * 100).toFixed(1).replace(".", ",")}%`;
}

function formatDuration(seconds) {
  seconds = Math.round(seconds || 0);
  const h = Math.floor(seconds / 3600);
  const m = Math.floor(seconds % 3600 / 60);
  const s = seconds % 60;
  if (h > 0) return `${h}ч ${String(m).padStart(2, "0")}м`;
  return `${m}м ${String(s).padStart(2, "0")}с`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;" }[ch]));
}

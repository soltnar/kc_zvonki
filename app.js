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
  workStart: document.getElementById("workStart"),
  workEnd: document.getElementById("workEnd"),
  occupancy: document.getElementById("occupancy"),
  maxWait: document.getElementById("maxWait"),
  buffer: document.getElementById("buffer"),
  replaceSbis: document.getElementById("replaceSbis"),
  includeOutbound: document.getElementById("includeOutbound"),
  notice: document.getElementById("notice"),
  kpis: document.getElementById("kpis"),
  missedThreshold: document.getElementById("missedThreshold"),
  missedSummary: document.getElementById("missedSummary"),
  missedHoursChart: document.getElementById("missedHoursChart"),
  hourMetric: document.getElementById("hourMetric"),
  hourChart: document.getElementById("hourChart"),
  dowChart: document.getElementById("dowChart"),
  monthChart: document.getElementById("monthChart"),
  staffingMonth: document.getElementById("staffingMonth"),
  staffing: document.getElementById("staffing"),
  bonusStart: document.getElementById("bonusStart"),
  bonusEnd: document.getElementById("bonusEnd"),
  bonusRate: document.getElementById("bonusRate"),
  bonusDailyFree: document.getElementById("bonusDailyFree"),
  bonusMinSec: document.getElementById("bonusMinSec"),
  bonusPdfBtn: document.getElementById("bonusPdfBtn"),
  bonusSummary: document.getElementById("bonusSummary"),
  bonusTable: document.querySelector("#bonusTable tbody"),
  operatorSearch: document.getElementById("operatorSearch"),
  operatorTable: document.querySelector("#operatorTable tbody"),
  heatmap: document.getElementById("heatmap")
};

els.operatorMap.value = DEFAULT_OPERATOR_MAP;
els.libraryStatus.textContent = window.XLSX ? "XLSX: готово" : "XLSX: библиотека не загружена";
els.analyzeBtn.disabled = !window.XLSX;

els.analyzeBtn.addEventListener("click", analyze);
els.hourMetric.addEventListener("change", () => lastStats && render(lastStats));
els.missedThreshold.addEventListener("change", () => lastStats && renderMissedFollowup(lastStats));
els.staffingMonth.addEventListener("change", () => lastStats && renderStaffing(lastStats));
els.operatorSearch.addEventListener("input", () => lastStats && renderOperators(lastStats.operatorRows));
[els.workStart, els.workEnd, els.occupancy, els.maxWait, els.buffer, els.replaceSbis, els.includeOutbound].forEach((el) => {
  el.addEventListener("change", () => lastStats && analyze());
});
[els.bonusStart, els.bonusEnd, els.bonusRate, els.bonusDailyFree, els.bonusMinSec].forEach((el) => {
  el.addEventListener("change", () => lastStats && renderBonus(lastStats));
});
els.bonusPdfBtn.addEventListener("click", exportBonusPdf);

async function analyze() {
  if (!els.mainFile.files[0]) {
    setNotice("Нужен основной файл “История внешних звонков”.", true);
    return;
  }

  try {
    els.analyzeBtn.disabled = true;
    setNotice("Готовлю разбор файлов...");
    await nextPaint();
    const mapping = parseMapping(els.operatorMap.value);
    const mainRows = await parseMainWorkbook(els.mainFile.files[0], mapping);
    const sbisRows = els.sbisFile.files[0] ? await parseSbisWorkbook(els.sbisFile.files[0]) : [];
    setNotice("Объединяю телефонии и считаю сводки...");
    await nextPaint();
    const rows = mergeSources(mainRows, sbisRows);
    const stats = await buildStats(rows, mainRows, sbisRows);
    lastStats = stats;
    setNotice("Рисую графики...");
    await nextPaint();
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

function columnIndex(headers) {
  return Object.fromEntries(headers.map((header, index) => [header, index]));
}

async function readWorkbook(file) {
  setNotice(`Открываю файл “${file.name}” (${formatFileSize(file.size)})...`);
  await nextPaint();
  const buffer = await file.arrayBuffer();
  setNotice(`Разбираю структуру Excel “${file.name}”. Для больших файлов это может занять минуту...`);
  await nextPaint();
  return XLSX.read(buffer, { type: "array", cellDates: true, dense: true });
}

async function parseMainWorkbook(file, mapping) {
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  setNotice("Читаю строки основной телефонии...");
  await nextPaint();
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false });
  const headerIndex = raw.findIndex((row) => row.includes("Тип звонка") && row.includes("Сотрудник"));
  if (headerIndex < 0) throw new Error("В основном файле не найдена строка заголовков.");
  const headers = raw[headerIndex].map(String);
  const idx = columnIndex(headers);
  const rows = [];
  for (let i = headerIndex + 1; i < raw.length; i++) {
    const row = raw[i];
    const date = parseMainDate(row[idx["Дата"]], row[idx["Время"]]);
    if (!date) continue;
    const type = String(row[idx["Тип звонка"]] || "").toLowerCase();
    const rawEmployee = String(row[idx["Сотрудник"]] || "").trim() || "Без оператора";
    rows.push({
      source: "Основная телефония",
      date,
      type,
      direction: type.includes("исход") ? "out" : "in",
      missed: type.includes("пропущ") || type.includes("неуспеш"),
      employee: normalizeMainEmployee(rawEmployee, mapping),
      rawEmployee,
      client: String(row[idx["Клиент"]] || ""),
      waitSec: parseDuration(row[idx["Ожидание"]]),
      talkSec: parseDuration(row[idx["Длительность"]])
    });
    if (i % 5000 === 0) {
      setNotice(`Основная телефония: обработано ${formatNumber(i - headerIndex)} из ${formatNumber(raw.length - headerIndex - 1)} строк...`);
      await nextFrame();
    }
  }
  return rows;
}

async function parseSbisWorkbook(file) {
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  setNotice("Читаю строки СБИС...");
  await nextPaint();
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false });
  const headerIndex = raw.findIndex((row) => row.includes("Начало звонка") && row.includes("Сотрудник"));
  if (headerIndex < 0) throw new Error("В файле СБИС не найдена строка заголовков.");
  const headers = raw[headerIndex].map(String);
  const idx = columnIndex(headers);
  const rows = [];
  for (let i = headerIndex + 1; i < raw.length; i++) {
    const row = raw[i];
    const date = parseDateTime(row[idx["Начало звонка"]]);
    if (!date) continue;
    const result = String(row[idx["Результат звонка"]] || "").toLowerCase();
    const directionText = String(row[idx["Направление звонка"]] || "").toLowerCase();
    rows.push({
      source: "СБИС",
      date,
      type: `${directionText} ${result}`.trim(),
      direction: directionText.includes("исход") ? "out" : "in",
      missed: /не отвечен|не дождался|занято|вышло время/.test(result),
      employee: shortName(String(row[idx["Сотрудник"]] || "СБИС без оператора")),
      rawEmployee: String(row[idx["Сотрудник"]] || ""),
      client: "",
      waitSec: 0,
      talkSec: parseDuration(row[idx["Время разговора"]])
    });
    if (i % 5000 === 0) {
      setNotice(`СБИС: обработано ${formatNumber(i - headerIndex)} из ${formatNumber(raw.length - headerIndex - 1)} строк...`);
      await nextFrame();
    }
  }
  return rows;
}

function mergeSources(mainRows, sbisRows) {
  if (!sbisRows.length) return mainRows;
  if (!els.replaceSbis.checked) return mainRows.concat(sbisRows);
  const sbisRange = dateRange(sbisRows);
  const baseRows = mainRows.filter((row) => {
    if (!/сбис/i.test(row.rawEmployee)) return true;
    return !isWithinDateRange(row.date, sbisRange);
  });
  baseRows.forEach((row) => {
    if (/сбис/i.test(row.rawEmployee)) {
      row.employee = "СБИС без детализации";
      row.serviceGroup = true;
    }
  });
  return baseRows.concat(sbisRows);
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

async function buildStats(rows, mainRows, sbisRows) {
  const activeDates = new Set();
  const workloadRows = [];
  let inbound = 0;
  let outbound = 0;
  let missed = 0;
  let unresolvedSbis = 0;
  let totalTalk = 0;
  let totalWait = 0;
  const waitStats = {
    inbound: { total: 0, count: 0 },
    outbound: { total: 0, count: 0 },
    missed: { total: 0, count: 0 }
  };

  rows.forEach((r) => {
    activeDates.add(isoDate(r.date));
    if (r.serviceGroup || r.employee === "СБИС без детализации") unresolvedSbis++;
    if (r.direction === "in") inbound++;
    if (r.direction === "out") outbound++;
    if (r.missed) {
      missed++;
    }
    totalTalk += r.talkSec;
    totalWait += r.waitSec;
    const waitGroup = r.missed ? waitStats.missed : waitStats[r.direction === "out" ? "outbound" : "inbound"];
    waitGroup.total += r.waitSec;
    waitGroup.count++;
    if ((r.direction === "in" || els.includeOutbound.checked) && isWithinWorkHours(r.date)) workloadRows.push(r);
  });

  setNotice("Считаю операторов, часы, дни недели и сезонность...");
  await nextPaint();
  const activeDays = activeDates.size || 1;
  const operators = groupOperators(rows);
  await nextFrame();
  const hour = aggregateRange(24, (r) => r.date.getHours(), workloadRows, true);
  const dow = aggregateRange(7, (r) => (r.date.getDay() + 6) % 7, rows, false);
  const month = aggregateRange(12, (r) => r.date.getMonth(), rows, false);
  const staffing = buildStaffingRecommendations(workloadRows);
  const heat = Array.from({ length: 7 }, () => Array.from({ length: 24 }, () => 0));
  rows.forEach((r) => { heat[(r.date.getDay() + 6) % 7][r.date.getHours()]++; });
  const range = dateRange(rows);

  return {
    rows,
    mainRows,
    sbisRows,
    activeDays,
    total: rows.length,
    inbound,
    outbound,
    missed,
    unresolvedSbis,
    missedRate: rows.length ? missed / rows.length : 0,
    totalTalk,
    totalWait,
    waitStats,
    avgTalk: rows.length ? totalTalk / rows.length : 0,
    operatorRows: operators,
    hour,
    dow,
    month,
    staffing,
    heat,
    range,
    sourceRanges: {
      main: dateRange(mainRows),
      sbis: dateRange(sbisRows)
    }
  };
}

function buildStaffingRecommendations(rows) {
  const cells = Array.from({ length: 12 }, () => Array.from({ length: 7 }, () => Array.from({ length: 24 }, () => emptyStaffCell())));
  rows.forEach((r) => {
    const month = r.date.getMonth();
    const dow = (r.date.getDay() + 6) % 7;
    const hour = r.date.getHours();
    const item = cells[month][dow][hour];
    item.calls++;
    if (r.missed) item.missed++;
    item.talkSec += r.talkSec;
    item.waitSec += r.waitSec;
    item.dates.add(isoDate(r.date));
  });

  const concurrency = calculateSeasonalConcurrency(rows);
  const availableMonths = [];
  const monthTotals = Array.from({ length: 12 }, () => 0);
  const occupancy = clamp(+els.occupancy.value || 75, 35, 95) / 100;
  const buffer = 1 + clamp(+els.buffer.value || 0, 0, 80) / 100;
  const maxWaitSec = clamp(+els.maxWait.value || 120, 15, 600);

  const byMonth = cells.map((monthCells, month) => {
    const table = monthCells.map((dowCells, dow) => dowCells.map((item, hour) => {
      monthTotals[month] += item.calls;
      const days = item.dates.size || 0;
      const avgCalls = days ? item.calls / days : 0;
      const avgTalkSec = item.calls ? item.talkSec / item.calls : 0;
      const staffing = calculateStaffing(avgCalls, avgTalkSec, occupancy, buffer, maxWaitSec);
      const peak = concurrency[month][dow][hour].peak;
      const avgPeak = concurrency[month][dow][hour].avgPeak;
      return {
        hour,
        dow,
        calls: item.calls,
        avgCalls,
        avgTalkSec,
        avgWaitSec: item.calls ? item.waitSec / item.calls : 0,
        missed: item.missed,
        peak,
        avgPeak,
        expectedWaitSec: staffing.expectedWaitSec,
        operators: Math.max(staffing.operators, Math.ceil(avgPeak))
      };
    }));
    if (monthTotals[month] > 0) availableMonths.push(month);
    return table;
  });

  const peak = Array.from({ length: 7 }, (_, dow) => Array.from({ length: 24 }, (_, hour) => {
    return availableMonths.reduce((best, month) => {
      const current = byMonth[month][dow][hour];
      if (!best || current.operators > best.operators || (current.operators === best.operators && current.avgCalls > best.avgCalls)) {
        return { ...current, month };
      }
      return best;
    }, null) || { ...emptyStaffResult(hour, dow), month: null };
  }));

  return { byMonth, peak, availableMonths, monthTotals, slots: getWorkHourSlots() };
}

function emptyStaffCell() {
  return { calls: 0, missed: 0, talkSec: 0, waitSec: 0, dates: new Set() };
}

function emptyStaffResult(hour, dow) {
  return { hour, dow, calls: 0, avgCalls: 0, avgTalkSec: 0, avgWaitSec: 0, missed: 0, peak: 0, avgPeak: 0, expectedWaitSec: 0, operators: 0 };
}

function calculateSeasonalConcurrency(rows) {
  const buckets = Array.from({ length: 12 }, () => Array.from({ length: 7 }, () => Array.from({ length: 24 }, () => new Map())));
  rows.forEach((r) => {
    const durationSec = Math.max(1, r.waitSec + r.talkSec);
    const startMs = r.date.getTime();
    const endMs = startMs + durationSec * 1000;
    let cursor = startOfHour(r.date);
    while (cursor.getTime() <= endMs) {
      const nextHour = new Date(cursor.getTime() + 3600000);
      const segStart = Math.max(startMs, cursor.getTime());
      const segEnd = Math.min(endMs, nextHour.getTime());
      if (segEnd > segStart) {
        const month = cursor.getMonth();
        const dow = (cursor.getDay() + 6) % 7;
        const hour = cursor.getHours();
        const key = `${isoDate(cursor)} ${hour}`;
        if (!buckets[month][dow][hour].has(key)) buckets[month][dow][hour].set(key, []);
        buckets[month][dow][hour].get(key).push([segStart - cursor.getTime(), 1], [segEnd - cursor.getTime(), -1]);
      }
      cursor = nextHour;
    }
  });
  return buckets.map((month) => month.map((dow) => dow.map(concurrencyFromSlotMap)));
}

function groupOperators(rows) {
  const map = new Map();
  rows.forEach((r) => {
    if (!map.has(r.employee)) {
      map.set(r.employee, { name: r.employee, total: 0, inbound: 0, outbound: 0, missed: 0, talkSec: 0, waitSec: 0, hours: new Set(), serviceGroup: false });
    }
    const item = map.get(r.employee);
    if (r.serviceGroup || r.employee === "СБИС без детализации") item.serviceGroup = true;
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
  })).sort((a, b) => Number(a.serviceGroup) - Number(b.serviceGroup) || b.total - a.total);
}

function aggregateRange(size, keyFn, rows, includeConcurrency) {
  const arr = Array.from({ length: size }, () => ({ calls: 0, missed: 0, talkSec: 0, waitSec: 0, dates: new Set() }));
  rows.forEach((r) => {
    const item = arr[keyFn(r)];
    item.calls++;
    if (r.missed) item.missed++;
    item.talkSec += r.talkSec;
    item.waitSec += r.waitSec;
    item.dates.add(isoDate(r.date));
  });
  const concurrency = includeConcurrency ? calculateHourlyConcurrency(rows, size, keyFn) : Array.from({ length: size }, () => ({ peak: 0, avgPeak: 0 }));
  const occupancy = clamp(+els.occupancy.value || 75, 35, 95) / 100;
  const buffer = 1 + clamp(+els.buffer.value || 0, 0, 80) / 100;
  const maxWaitSec = clamp(+els.maxWait.value || 120, 15, 600);
  return arr.map((item, index) => {
    const days = item.dates.size || 1;
    const avgCalls = item.calls / days;
    const avgTalkSec = item.calls ? item.talkSec / item.calls : 0;
    const avgWaitSec = item.calls ? item.waitSec / item.calls : 0;
    const loadSec = item.talkSec / days;
    const staffing = calculateStaffing(avgCalls, avgTalkSec, occupancy, buffer, maxWaitSec);
    return {
      index,
      calls: item.calls,
      missed: item.missed,
      avgCalls,
      avgTalkSec,
      avgWaitSec,
      loadHours: loadSec / 3600,
      concurrent: concurrency[index].peak,
      avgConcurrent: concurrency[index].avgPeak,
      serviceLevel: staffing.serviceLevel,
      expectedWaitSec: staffing.expectedWaitSec,
      traffic: staffing.traffic,
      operators: Math.max(staffing.operators, Math.ceil(concurrency[index].avgPeak))
    };
  });
}

function calculateHourlyConcurrency(rows, size, keyFn) {
  const buckets = Array.from({ length: size }, () => new Map());
  rows.forEach((r) => {
    const durationSec = Math.max(1, r.waitSec + r.talkSec);
    const startMs = r.date.getTime();
    const endMs = startMs + durationSec * 1000;
    let cursor = startOfHour(r.date);
    while (cursor.getTime() <= endMs) {
      const nextHour = new Date(cursor.getTime() + 3600000);
      const segStart = Math.max(startMs, cursor.getTime());
      const segEnd = Math.min(endMs, nextHour.getTime());
      if (segEnd > segStart) {
        const hour = keyFn({ ...r, date: cursor });
        const key = `${isoDate(cursor)} ${cursor.getHours()}`;
        if (!buckets[hour].has(key)) buckets[hour].set(key, []);
        buckets[hour].get(key).push([segStart - cursor.getTime(), 1], [segEnd - cursor.getTime(), -1]);
      }
      cursor = nextHour;
    }
  });
  return buckets.map((slotMap) => {
    return concurrencyFromSlotMap(slotMap);
  });
}

function concurrencyFromSlotMap(slotMap) {
  let peak = 0;
  let peakSum = 0;
  slotMap.forEach((events) => {
    events.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    let active = 0;
    let slotPeak = 0;
    events.forEach((event) => {
      active += event[1];
      if (active > slotPeak) slotPeak = active;
    });
    peak = Math.max(peak, slotPeak);
    peakSum += slotPeak;
  });
  return { peak, avgPeak: slotMap.size ? peakSum / slotMap.size : 0 };
}

function calculateStaffing(avgCalls, avgTalkSec, occupancy, buffer, maxWaitSec) {
  if (!avgCalls || !avgTalkSec) {
    return { operators: 0, traffic: 0, expectedWaitSec: 0, serviceLevel: 1 };
  }
  const traffic = avgCalls * avgTalkSec / 3600;
  let agents = Math.max(1, Math.ceil(traffic / occupancy), Math.floor(traffic) + 1);
  let result = erlangResult(traffic, avgTalkSec, agents, maxWaitSec);
  while ((result.expectedWaitSec > maxWaitSec || agents <= traffic) && agents < 200) {
    agents++;
    result = erlangResult(traffic, avgTalkSec, agents, maxWaitSec);
  }
  return {
    operators: Math.ceil(agents * buffer),
    traffic,
    expectedWaitSec: result.expectedWaitSec,
    serviceLevel: result.serviceLevel
  };
}

function erlangResult(traffic, avgTalkSec, agents, maxWaitSec) {
  if (traffic <= 0 || agents <= 0) return { expectedWaitSec: 0, serviceLevel: 1 };
  if (agents <= traffic) return { expectedWaitSec: Infinity, serviceLevel: 0 };
  const waitProbability = erlangC(traffic, agents);
  const expectedWaitSec = waitProbability * avgTalkSec / (agents - traffic);
  const serviceLevel = 1 - waitProbability * Math.exp(-(agents - traffic) * maxWaitSec / avgTalkSec);
  return { expectedWaitSec, serviceLevel: clamp(serviceLevel, 0, 1) };
}

function erlangC(traffic, agents) {
  let term = 1;
  let sum = 1;
  for (let n = 1; n < agents; n++) {
    term *= traffic / n;
    sum += term;
  }
  const last = term * traffic / agents;
  const queueTerm = last * agents / (agents - traffic);
  return queueTerm / (sum + queueTerm);
}

function render(stats) {
  renderNotice(stats);
  renderKpis(stats);
  renderMissedFollowup(stats);
  renderBarChart(els.hourChart, stats.hour, hourLabels(), els.hourMetric.value);
  renderBarChart(els.dowChart, stats.dow, DOW, "calls");
  renderBarChart(els.monthChart, stats.month, MONTHS, "calls");
  renderStaffingMonthOptions(stats);
  renderStaffing(stats);
  renderBonus(stats);
  renderOperators(stats.operatorRows);
  renderHeatmap(stats.heat);
}

function renderNotice(stats) {
  const parts = [`Период анализа: ${formatDate(stats.range.min)} - ${formatDate(stats.range.max)}. Строк: ${formatNumber(stats.total)}. Расчет смен: ${els.workStart.value} - ${els.workEnd.value}.`];
  if (stats.sbisRows.length) {
    if (els.replaceSbis.checked) {
      parts.push(`СБИС расшифровывает операторов за ${formatDate(stats.sourceRanges.sbis.min)} - ${formatDate(stats.sourceRanges.sbis.max)}; остальной период остается из основной телефонии.`);
    } else {
      parts.push(`СБИС добавлен отдельным источником за ${formatDate(stats.sourceRanges.sbis.min)} - ${formatDate(stats.sourceRanges.sbis.max)}.`);
    }
  }
  setNotice(parts.join(" "));
}

function renderKpis(stats) {
  const cards = [
    ["Всего звонков", formatNumber(stats.total), `${formatNumber(stats.activeDays)} дней с данными`],
    ["Входящие", formatNumber(stats.inbound), `${percent(stats.inbound / Math.max(stats.total, 1))} от всех`],
    ["Исходящие", formatNumber(stats.outbound), `${percent(stats.outbound / Math.max(stats.total, 1))} от всех`],
    ["Пропущенные", formatNumber(stats.missed), percent(stats.missedRate)],
    ["Разговор", formatDuration(stats.totalTalk), `среднее ${formatDuration(stats.avgTalk)}`],
    ["Ожидание входящих", formatDuration(stats.waitStats.inbound.total), `среднее ${formatDuration(stats.waitStats.inbound.count ? stats.waitStats.inbound.total / stats.waitStats.inbound.count : 0)} на звонок`],
    ["Ожидание исходящих", formatDuration(stats.waitStats.outbound.total), `среднее ${formatDuration(stats.waitStats.outbound.count ? stats.waitStats.outbound.total / stats.waitStats.outbound.count : 0)} на звонок`],
    ["Ожидание пропущенных", formatDuration(stats.waitStats.missed.total), `среднее ${formatDuration(stats.waitStats.missed.count ? stats.waitStats.missed.total / stats.waitStats.missed.count : 0)} на звонок`],
    ["СБИС без детализации", formatNumber(stats.unresolvedSbis), `${percent(stats.unresolvedSbis / Math.max(stats.total, 1))} строк`]
  ];
  els.kpis.innerHTML = cards.map(([label, value, sub]) => `<div class="panel kpi"><span>${label}</span><strong>${value}</strong><small>${sub}</small></div>`).join("");
}

function renderMissedFollowup(stats) {
  const threshold = clamp(Math.round(+els.missedThreshold.value || 10), 1, 20);
  els.missedThreshold.value = threshold;
  const inboundMissed = stats.rows.filter((r) => r.direction === "in" && r.missed);
  const qualified = inboundMissed.filter((r) => r.waitSec > threshold);
  const outboundByClientDay = new Map();

  stats.rows.forEach((r) => {
    if (r.direction !== "out") return;
    const client = normalizeClientNumber(r.client);
    if (!client) return;
    const key = `${isoDate(r.date)}|${client}`;
    if (!outboundByClientDay.has(key)) outboundByClientDay.set(key, []);
    outboundByClientDay.get(key).push(r.date.getTime());
  });

  const missedByClientDay = new Map();
  let recognizedCalls = 0;
  qualified.forEach((r) => {
    const client = normalizeClientNumber(r.client);
    if (!client) return;
    recognizedCalls++;
    const key = `${isoDate(r.date)}|${client}`;
    const current = missedByClientDay.get(key);
    if (!current || r.date > current.date) missedByClientDay.set(key, { date: r.date, client });
  });

  const unresolved = Array.from(missedByClientDay.entries()).filter(([key, item]) => {
    return !(outboundByClientDay.get(key) || []).some((time) => time > item.date.getTime());
  });
  const hourly = Array.from({ length: 24 }, (_, hour) => ({ unique: 0, hour }));
  unresolved.forEach(([, item]) => hourly[item.date.getHours()].unique++);

  els.missedSummary.innerHTML = `
    <div class="missed-card"><span>Пропущено &gt; ${threshold} сек</span><strong>${formatNumber(qualified.length)}</strong><small>${percent(qualified.length / Math.max(inboundMissed.length, 1))} от входящих пропущенных</small></div>
    <div class="missed-card alert"><span>Уникальных без перезвона</span><strong>${formatNumber(unresolved.length)}</strong><small>до конца дня</small></div>
    <div class="missed-card"><span>Распознано номеров</span><strong>${formatNumber(missedByClientDay.size)}</strong><small>${formatNumber(recognizedCalls)} звонков для проверки</small></div>
  `;
  renderBarChart(els.missedHoursChart, hourly, hourLabels(), "unique");
}

function normalizeClientNumber(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (digits.length < 7) return "";
  return digits.slice(-10);
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

function renderBonus(stats) {
  ensureBonusPeriod(stats.range);
  const result = calculateBonus(stats.rows);
  els.bonusSummary.innerHTML = `
    <div class="bonus-card"><span>Всего премия</span><strong>${formatMoney(result.totalBonus)}</strong></div>
    <div class="bonus-card"><span>Звонков к премии</span><strong>${formatNumber(result.totalPaidCalls)}</strong></div>
    <div class="bonus-card"><span>Учтено звонков</span><strong>${formatNumber(result.totalEligibleCalls)}</strong></div>
    <div class="bonus-card"><span>Период</span><strong>${formatDateInput(els.bonusStart.value)} - ${formatDateInput(els.bonusEnd.value)}</strong></div>
  `;
  els.bonusTable.innerHTML = result.rows.map((r) => `
    <tr>
      <td>${escapeHtml(r.name)}</td>
      <td>${formatNumber(r.days)}</td>
      <td>${formatNumber(r.eligibleCalls)}</td>
      <td>${formatNumber(r.bonusDays)}</td>
      <td>${formatNumber(r.paidCalls)}</td>
      <td>${formatMoney(r.bonus)}</td>
      <td>${formatNumber(r.avgEligiblePerDay)}</td>
    </tr>
  `).join("");
}

function ensureBonusPeriod(range) {
  if (!els.bonusStart.value && range.min) els.bonusStart.value = isoDate(range.min);
  if (!els.bonusEnd.value && range.max) els.bonusEnd.value = isoDate(range.max);
}

function calculateBonus(rows) {
  const start = parseDateInput(els.bonusStart.value, false);
  const end = parseDateInput(els.bonusEnd.value, true);
  const rate = Math.max(0, Number(els.bonusRate.value) || 6.5);
  const dailyFree = Math.max(0, Math.floor(Number(els.bonusDailyFree.value) || 80));
  const minSec = Math.max(0, Math.floor(Number(els.bonusMinSec.value) || 10));
  const byOperator = new Map();

  rows.forEach((r) => {
    if (r.serviceGroup || r.employee === "СБИС без детализации") return;
    if (r.date < start || r.date > end) return;
    if (r.talkSec < minSec) return;
    if (!byOperator.has(r.employee)) byOperator.set(r.employee, new Map());
    const byDay = byOperator.get(r.employee);
    const day = isoDate(r.date);
    byDay.set(day, (byDay.get(day) || 0) + 1);
  });

  const resultRows = [...byOperator.entries()].map(([name, daysMap]) => {
    let eligibleCalls = 0;
    let bonusDays = 0;
    let paidCalls = 0;
    daysMap.forEach((count) => {
      eligibleCalls += count;
      if (count > dailyFree) {
        bonusDays++;
        paidCalls += count;
      }
    });
    return {
      name,
      days: daysMap.size,
      eligibleCalls,
      bonusDays,
      paidCalls,
      bonus: paidCalls * rate,
      avgEligiblePerDay: daysMap.size ? eligibleCalls / daysMap.size : 0
    };
  }).filter((r) => r.eligibleCalls > 0).sort((a, b) => b.bonus - a.bonus || b.eligibleCalls - a.eligibleCalls);

  return {
    rows: resultRows,
    totalBonus: resultRows.reduce((sum, r) => sum + r.bonus, 0),
    totalPaidCalls: resultRows.reduce((sum, r) => sum + r.paidCalls, 0),
    totalEligibleCalls: resultRows.reduce((sum, r) => sum + r.eligibleCalls, 0)
  };
}

function exportBonusPdf() {
  if (!lastStats) {
    setNotice("Сначала загрузите файлы и посчитайте отчет.", true);
    return;
  }
  ensureBonusPeriod(lastStats.range);
  const result = calculateBonus(lastStats.rows);
  const reportWindow = window.open("", "_blank");
  if (!reportWindow) {
    setNotice("Браузер заблокировал окно PDF. Разрешите всплывающие окна для сайта.", true);
    return;
  }
  reportWindow.document.open();
  reportWindow.document.write(buildBonusPrintHtml(result));
  reportWindow.document.close();
  window.setTimeout(() => {
    reportWindow.focus();
    reportWindow.print();
  }, 350);
}

function buildBonusPrintHtml(result) {
  const reportName = `Премия_${formatDateInput(els.bonusStart.value)}-${formatDateInput(els.bonusEnd.value)}`;
  const rows = result.rows.map((r) => `
    <tr>
      <td>${escapeHtml(r.name)}</td>
      <td>${formatNumber(r.days)}</td>
      <td>${formatNumber(r.eligibleCalls)}</td>
      <td>${formatNumber(r.bonusDays)}</td>
      <td>${formatNumber(r.paidCalls)}</td>
      <td>${formatMoney(r.bonus)}</td>
      <td>${formatNumber(r.avgEligiblePerDay)}</td>
    </tr>
  `).join("");
  return `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>${reportName}</title>
  <style>
    @page { size: A4 landscape; margin: 10mm; }
    @media print {
      @page { size: A4 landscape; margin: 10mm; }
      html, body { width: 277mm; }
    }
    * { box-sizing: border-box; }
    body { margin: 0; color: #17202a; font: 12px Arial, sans-serif; }
    h1 { margin: 0 0 6px; font-size: 22px; }
    .meta { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin: 14px 0; }
    .card { border: 1px solid #d9e0e7; border-radius: 6px; padding: 8px; }
    .card span { display: block; color: #657282; font-size: 10px; }
    .card strong { display: block; margin-top: 3px; font-size: 16px; }
    .rules { margin: 8px 0 14px; color: #657282; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #d9e0e7; padding: 6px 7px; text-align: right; white-space: nowrap; }
    th { background: #eef2f7; color: #334155; }
    td:first-child, th:first-child { text-align: left; white-space: normal; }
    tfoot td { font-weight: 700; background: #f8fafc; }
  </style>
</head>
<body>
  <h1>${reportName.replace("_", " ")}</h1>
  <div class="rules">
    Период: ${formatDateInput(els.bonusStart.value)} - ${formatDateInput(els.bonusEnd.value)}.
    Учитываются звонки от ${formatNumber(Number(els.bonusMinSec.value) || 10)} сек.
    Если за день больше ${formatNumber(Number(els.bonusDailyFree.value) || 80)} звонков, оплачиваются все звонки этого дня по ${formatMoney(Number(els.bonusRate.value) || 6.5)} за звонок.
  </div>
  <div class="meta">
    <div class="card"><span>Всего премия</span><strong>${formatMoney(result.totalBonus)}</strong></div>
    <div class="card"><span>Звонков к премии</span><strong>${formatNumber(result.totalPaidCalls)}</strong></div>
    <div class="card"><span>Учтено звонков</span><strong>${formatNumber(result.totalEligibleCalls)}</strong></div>
    <div class="card"><span>Дата выгрузки</span><strong>${new Date().toLocaleDateString("ru-RU")}</strong></div>
  </div>
  <table>
    <thead>
      <tr>
        <th>Оператор</th>
        <th>Дней</th>
        <th>Звонков</th>
        <th>Дней с премией</th>
        <th>К премии</th>
        <th>Премия</th>
        <th>Средн./день</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
    <tfoot>
      <tr>
        <td>Итого</td>
        <td></td>
        <td>${formatNumber(result.totalEligibleCalls)}</td>
        <td></td>
        <td>${formatNumber(result.totalPaidCalls)}</td>
        <td>${formatMoney(result.totalBonus)}</td>
        <td></td>
      </tr>
    </tfoot>
  </table>
</body>
</html>`;
}

function renderStaffingMonthOptions(stats) {
  const current = els.staffingMonth.value || "peak";
  const options = [`<option value="peak">Пиковый сезон</option>`].concat(
    stats.staffing.availableMonths.map((month) => `<option value="${month}">${MONTHS[month]}</option>`)
  );
  els.staffingMonth.innerHTML = options.join("");
  els.staffingMonth.value = [...els.staffingMonth.options].some((option) => option.value === current) ? current : "peak";
}

function renderStaffing(stats) {
  const selected = els.staffingMonth.value || "peak";
  const table = selected === "peak" ? stats.staffing.peak : stats.staffing.byMonth[+selected];
  const maxOperators = Math.max(...stats.staffing.slots.flatMap((hour) => DOW.map((_, dow) => table[dow][hour].operators)), 1);
  const head = `<div class="staff-cell staff-head">Время</div>${DOW.map((day) => `<div class="staff-cell staff-head">${day}</div>`).join("")}`;
  const rows = stats.staffing.slots.map((hour) => {
    const cells = DOW.map((_, dow) => {
      const item = table[dow][hour];
      const level = maxOperators ? item.operators / maxOperators : 0;
      const monthLabel = selected === "peak" && item.month !== null ? ` · ${MONTHS[item.month]}` : "";
      return `<div class="staff-cell staff-plan" style="--level:${level}" title="${item.avgCalls.toFixed(1)} звонков/день${monthLabel}; пик ${item.peak}; средний разговор ${formatDuration(item.avgTalkSec)}; ожидание ${formatDuration(item.expectedWaitSec)}">
        <strong>${item.operators || ""}</strong>
        <span>${item.avgCalls ? item.avgCalls.toFixed(1) : ""}${monthLabel}</span>
      </div>`;
    }).join("");
    return `<div class="staff-cell staff-time">${formatWorkHourLabel(hour)}</div>${cells}`;
  }).join("");
  els.staffing.innerHTML = `<div class="staff-grid">${head}${rows}</div>`;
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
  if (!canvas.dataset.logicalHeight) {
    canvas.dataset.logicalHeight = String(Number(canvas.getAttribute("height")) || 240);
  }
  const logicalHeight = Number(canvas.dataset.logicalHeight);
  canvas.style.height = `${logicalHeight}px`;
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
  els.notice.classList.toggle("busy", text.includes("..."));
}

function nextPaint() {
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

function nextFrame() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function dateRange(rows) {
  if (!rows.length) return { min: null, max: null };
  return rows.reduce((acc, r) => ({
    min: !acc.min || r.date < acc.min ? r.date : acc.min,
    max: !acc.max || r.date > acc.max ? r.date : acc.max
  }), { min: null, max: null });
}

function isWithinDateRange(date, range) {
  if (!range.min || !range.max) return false;
  const day = startOfDay(date).getTime();
  return day >= startOfDay(range.min).getTime() && day <= startOfDay(range.max).getTime();
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function startOfHour(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours());
}

function isWithinWorkHours(date) {
  const start = parseClock(els.workStart.value || "09:30");
  const end = parseClock(els.workEnd.value || "23:00");
  const minute = date.getHours() * 60 + date.getMinutes();
  if (start <= end) return minute >= start && minute < end;
  return minute >= start || minute < end;
}

function parseClock(value) {
  const [h, m] = String(value || "00:00").split(":").map(Number);
  return (h || 0) * 60 + (m || 0);
}

function getWorkHourSlots() {
  const start = parseClock(els.workStart.value || "09:30");
  const end = parseClock(els.workEnd.value || "23:00");
  return Array.from({ length: 24 }, (_, hour) => hour).filter((hour) => {
    const slotStart = hour * 60;
    const slotEnd = slotStart + 60;
    if (start <= end) return slotEnd > start && slotStart < end;
    return slotEnd > start || slotStart < end;
  });
}

function formatWorkHourLabel(hour) {
  const start = parseClock(els.workStart.value || "09:30");
  if (Math.floor(start / 60) === hour && start % 60) {
    return `${String(hour).padStart(2, "0")}:${String(start % 60).padStart(2, "0")}`;
  }
  return `${String(hour).padStart(2, "0")}:00`;
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

function parseDateInput(value, endOfDay) {
  const parsed = parseDateTime(value);
  if (!parsed) return endOfDay ? new Date(8640000000000000) : new Date(-8640000000000000);
  parsed.setHours(endOfDay ? 23 : 0, endOfDay ? 59 : 0, endOfDay ? 59 : 0, endOfDay ? 999 : 0);
  return parsed;
}

function formatDateInput(value) {
  const parsed = parseDateTime(value);
  return parsed ? parsed.toLocaleDateString("ru-RU") : "не задан";
}

function formatNumber(value) {
  return Math.round(value).toLocaleString("ru-RU");
}

function formatMoney(value) {
  return `${Math.round(value).toLocaleString("ru-RU")} ₽`;
}

function formatFileSize(bytes) {
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1).replace(".", ",")} МБ`;
  if (bytes >= 1024) return `${Math.round(bytes / 1024)} КБ`;
  return `${bytes} Б`;
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
  if (!Number.isFinite(seconds)) return ">10м";
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

// ================================================================
// 설정
// ================================================================
const STORAGE_KEY = "근태대시보드_v1";
const DRIVE_XLSX_URL = "https://docs.google.com/spreadsheets/d/1grIcwPHx4XanTASz9UGmANC8L6bNAIMdH5D2h6wP73Q/export?format=xlsx";
const HOURS_PER_DAY = 8;

// ================================================================
// 전역 상태
// ================================================================
const state = {
  records: [],
  months: [],
  departments: [],
  byDeptMonth: new Map(),      // "dept__month"     → records[]
  byDeptMonthName: new Map(),  // "dept__month__name" → records[]
  employees: new Map(),        // "dept__month"     → Map(name → empId)
};

// ================================================================
// DOM 헬퍼
// ================================================================
const $ = (id) => document.getElementById(id);

const loadDriveBtn  = $("loadDriveBtn");
const clearBtn      = $("clearBtn");
const printBtn      = $("printBtn");
const localFileBtn  = $("localFileBtn");
const fileInput     = $("fileInput");
const deptSelect    = $("departmentSelect");
const monthSelect   = $("monthSelect");
const empSelect     = $("employeeSelect");

// ================================================================
// 유틸: 날짜·시간 파싱
// ================================================================

/** 엑셀 시리얼 숫자 → Date */
function excelSerial(n) {
  return new Date(Date.UTC(1899, 11, 30) + n * 86400000);
}

/** 날짜 파싱 */
function parseDate(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") return excelSerial(v);
  if (typeof v === "string") {
    const s = v.replace(/\(.*?\)/g, "").trim();
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }
  return null;
}

/** 날짜 → "YYYY-MM-DD" */
function toDateStr(v) {
  const d = parseDate(v);
  return d ? d.toISOString().slice(0, 10) : "-";
}

/** 숫자·문자열 → 시간(소수) */
function parseHours(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  const s = String(v).trim();
  const hm = s.match(/^(\d{1,2}):(\d{2})$/);
  if (hm) return +hm[1] + +hm[2] / 60;
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/** "N일 N시간 N분" 문자열 → 시간(소수) */
function parseDuration(text) {
  const s = String(text ?? "");
  const days  = Number(s.match(/(\d+)\s*일/)?.[1] ?? 0);
  const hours = Number(s.match(/(\d+)\s*시간/)?.[1] ?? 0);
  const mins  = Number(s.match(/(\d+)\s*분/)?.[1] ?? 0);
  return days * HOURS_PER_DAY + hours + mins / 60;
}

/** 시간(소수) → "N일 N시간 N분" */
function fmtHours(hours) {
  const total = Math.round((Number(hours) || 0) * 60);
  if (total <= 0) return "0시간";
  const days = Math.floor(total / (HOURS_PER_DAY * 60));
  const rem  = total % (HOURS_PER_DAY * 60);
  const h    = Math.floor(rem / 60);
  const m    = rem % 60;
  const parts = [];
  if (days) parts.push(`${days}일`);
  if (h)    parts.push(`${h}시간`);
  if (m)    parts.push(`${m}분`);
  return parts.join(" ") || "0시간";
}

/** 월간/연간 포맷 */
function fmtMA(monthly, annual, overtime = false) {
  if (overtime) return `${Number(monthly).toFixed(2)}시간 / ${Number(annual).toFixed(2)}시간`;
  return `${fmtHours(monthly)} / ${fmtHours(annual)}`;
}

/** "2026-01-15 09:00" 형식 → Date */
function parseDatetimeLine(line) {
  const m = line.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
}

/** 출장 시간 계산 */
function calcTripHours(fromLine, toLine) {
  const f = parseDatetimeLine(fromLine);
  const t = parseDatetimeLine(toLine);
  if (!f || !t) return 0;
  const h = (t - f) / 3600000;
  return h > 0 ? h : 0;
}

// ================================================================
// 유틸: 월 정렬
// ================================================================
const monthOrder = (name) => Number(String(name).match(/\d+/)?.[0] ?? 999);
const sortMonths = (arr) => arr.slice().sort((a, b) => monthOrder(a) - monthOrder(b) || a.localeCompare(b, "ko"));

// ================================================================
// 엑셀 셀 파싱
// ================================================================

/** 시간외 셀 파싱 */
function parseOvertimeCell(cell) {
  if (typeof cell === "number") return cell;
  const s = String(cell ?? "");
  if (!s.trim()) return 0;
  const labeled = s.match(/(?:총시간외?|시간외시간|실근무)\s*[:：]\s*([^\n\r]+)/);
  if (labeled) return parseHours(labeled[1]);
  if (/신청시각|실근무|종별/.test(s)) return 0;
  return parseHours(s);
}

/** 휴가관리 셀 파싱 → 엔트리 배열 */
function parseLeaveCell(cell) {
  const text = String(cell ?? "").trim();
  if (!text) return [];
  const lines = text.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  const entries = [];
  let cur = null;

  for (const line of lines) {
    const typeM = line.match(/^종별\s*[:：]\s*(.+)$/);
    if (typeM) {
      if (cur) entries.push(cur);
      cur = { typeRaw: typeM[1].trim(), durationHours: 0, reason: "", fromLine: "", toLine: "" };
      continue;
    }
    if (!cur) continue;
    const durM = line.match(/^일수\/시간\s*[:：]\s*(.+)$/);
    if (durM)  { cur.durationHours = parseDuration(durM[1]); continue; }
    const reaM = line.match(/^사유\s*[:：]\s*(.+)$/);
    if (reaM)  { cur.reason = reaM[1].trim(); continue; }
    if (/^부터\s*[:：]/.test(line)) cur.fromLine = line;
    if (/^까지\s*[:：]/.test(line)) cur.toLine   = line;
  }
  if (cur) entries.push(cur);
  return entries;
}

/** 출장관리 셀 파싱 → 엔트리 배열 */
function parseTripCell(cell) {
  const text = String(cell ?? "").trim();
  if (!text) return [];
  const lines = text.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  const entries = [];
  let cur = null;

  for (const line of lines) {
    const typeM = line.match(/^종별\s*[:：]\s*(.+)$/);
    if (typeM) {
      if (cur) entries.push(cur);
      cur = { typeRaw: typeM[1].trim(), fromLine: "", toLine: "", reason: "" };
      continue;
    }
    if (!cur) continue;
    if (/^부터\s*[:：]/.test(line)) cur.fromLine = line;
    if (/^까지\s*[:：]/.test(line)) cur.toLine   = line;
    const reaM = line.match(/^사유\s*[:：]\s*(.+)$/);
    if (reaM) cur.reason = reaM[1].trim();
  }
  if (cur) entries.push(cur);
  return entries;
}

// ================================================================
// 분류
// ================================================================

function classifyLeave(typeRaw, durationHours) {
  const t = typeRaw.replace(/\s+/g, "");
  if (t.includes("병가"))                   return { category: "병가",          subType: typeRaw };
  if (t.includes("산전후휴가"))              return { category: "산전후휴가",    subType: typeRaw };
  if (t.includes("임산부정기검진"))          return { category: "임산부정기검진", subType: typeRaw };
  if (t.includes("임신기단축"))              return { category: "임신기단축",    subType: "임신기 단축" };
  if (t.includes("근속휴가"))               return { category: "근속휴가",      subType: "근속휴가" };
  if (t.includes("법인발전유공휴가") || t.includes("포상휴가") || t.includes("포상"))
                                             return { category: "포상휴가",      subType: "포상휴가" };
  if (t.includes("조퇴"))                   return { category: "조퇴",          subType: typeRaw };
  if (t.includes("대체휴무"))               return { category: "대휴",          subType: durationHours > 0 ? `대체휴무(${fmtHours(durationHours)})` : "대체휴무", isDayOff: true };
  if (t.includes("연차"))                   return { category: "연차",          subType: durationHours > 0 ? `연차(${fmtHours(durationHours)})` : "연차", isDayOff: true };
  if (t.includes("휴무") || t.includes("휴가")) return { category: "휴무",     subType: durationHours > 0 ? `휴무(${fmtHours(durationHours)})` : "휴무", isDayOff: true };
  return { category: "기타", subType: typeRaw };
}

function classifyTrip(typeRaw) {
  const t = typeRaw.replace(/\s+/g, "");
  if (t.includes("관내출장")) return { category: "출장", subType: "관내출장" };
  if (t.includes("관외출장")) return { category: "출장", subType: "관외출장" };
  if (t.includes("국외출장")) return { category: "출장", subType: "국외출장" };
  if (t.includes("국내출장")) return { category: "출장", subType: "국내출장" };
  if (t.includes("출장"))     return { category: "출장", subType: typeRaw };
  return { category: "기타", subType: typeRaw };
}

// ================================================================
// 엑셀 파싱
// ================================================================

/** 헤더 행 인덱스 자동 탐지 */
function findHeaderRow(sheet) {
  const keywords = ["성명", "이름", "날짜", "일자", "부서", "소속", "팀", "휴가관리", "시간외관리", "출장관리"];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
  const limit = Math.min(raw.length, 30);
  for (let i = 0; i < limit; i++) {
    const cells = (raw[i] || []).map((v) => String(v ?? "").replace(/\s+/g, "").trim());
    const hits  = keywords.filter((k) => cells.includes(k));
    if (hits.length >= 2 && (cells.includes("성명") || cells.includes("이름"))) return i;
  }
  return 0;
}

/** 행(row 객체) → 근태 레코드[] */
function parseRow(row, sheetName) {
  const name = String(row["성명"] ?? row["이름"] ?? "").trim();
  if (!name) return [];

  const dept    = String(row["부서"] ?? row["소속"] ?? row["팀"] ?? "미지정").trim() || "미지정";
  const empId   = String(row["사원번호"] ?? row["사번"] ?? "-").trim() || "-";
  const dateStr = toDateStr(row["날짜"] ?? row["일자"]);

  const records = [];

  // 시간외
  const otHours = parseOvertimeCell(row["시간외관리"] ?? row["시간외(시간)"] ?? row["시간외"]);
  if (otHours > 0) {
    records.push({ month: sheetName, dept, name, empId, date: dateStr, category: "시간외", subType: "시간외근무", tripType: "", overtimeHours: otHours, durationHours: 0, dedupeKey: "" });
  }

  // 조기퇴근/조퇴
  const earlyH = parseHours(row["조기퇴근"] ?? row["조퇴"]);
  if (earlyH > 0) {
    records.push({ month: sheetName, dept, name, empId, date: dateStr, category: "조퇴", subType: "조기퇴근", tripType: "", overtimeHours: 0, durationHours: earlyH, dedupeKey: "" });
  }

  // 휴가관리
  for (const entry of parseLeaveCell(row["휴가관리"] ?? row["근태유형"] ?? row["유형"])) {
    const c           = classifyLeave(entry.typeRaw, entry.durationHours);
    const fromDt      = parseDatetimeLine(entry.fromLine);
    const recordDate  = fromDt ? fromDt.toISOString().slice(0, 10) : dateStr;
    const dedupeKey   = (entry.fromLine && entry.toLine)
      ? `${name}|${dept}|${c.category}|${entry.reason}|${entry.durationHours}|${entry.fromLine}|${entry.toLine}`
      : "";
    records.push({ month: sheetName, dept, name, empId, date: recordDate, category: c.category, subType: entry.reason || c.subType, tripType: "", overtimeHours: 0, durationHours: entry.durationHours, dedupeKey });
  }

  // 출장관리
  for (const entry of parseTripCell(row["출장관리"])) {
    const c     = classifyTrip(entry.typeRaw);
    const tripH = calcTripHours(entry.fromLine, entry.toLine);
    records.push({ month: sheetName, dept, name, empId, date: dateStr, category: c.category, subType: entry.reason || c.subType, tripType: c.subType, overtimeHours: 0, durationHours: tripH, dedupeKey: "" });
  }

  return records;
}

/** ArrayBuffer → { records, months, departments } */
function parseWorkbook(buffer) {
  const wb  = XLSX.read(buffer, { type: "array" });
  const all = [];

  for (const sheetName of wb.SheetNames) {
    const sheet  = wb.Sheets[sheetName];
    const hdrRow = findHeaderRow(sheet);
    const rows   = XLSX.utils.sheet_to_json(sheet, { defval: "", range: hdrRow });
    for (const row of rows) all.push(...parseRow(row, sheetName));
  }

  // 중복 제거
  const seen    = new Set();
  const deduped = all.filter((r) => {
    if (!r.dedupeKey) return true;
    if (seen.has(r.dedupeKey)) return false;
    seen.add(r.dedupeKey);
    return true;
  });

  const departments = [...new Set(deduped.map((r) => r.dept).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
  return { records: deduped, months: wb.SheetNames.slice(), departments };
}

// ================================================================
// 상태 관리
// ================================================================

function applyData({ records, months, departments }) {
  state.records     = records;
  state.months      = months;
  state.departments = departments;

  state.byDeptMonth.clear();
  state.byDeptMonthName.clear();
  state.employees.clear();

  for (const r of records) {
    const dmKey  = `${r.dept}__${r.month}`;
    const dmnKey = `${r.dept}__${r.month}__${r.name}`;

    if (!state.byDeptMonth.has(dmKey))     state.byDeptMonth.set(dmKey, []);
    if (!state.byDeptMonthName.has(dmnKey)) state.byDeptMonthName.set(dmnKey, []);
    if (!state.employees.has(dmKey))        state.employees.set(dmKey, new Map());

    state.byDeptMonth.get(dmKey).push(r);
    state.byDeptMonthName.get(dmnKey).push(r);
    state.employees.get(dmKey).set(r.name, r.empId);
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    records: state.records,
    months: state.months,
    departments: state.departments,
    savedAt: new Date().toISOString(),
  }));
}

function loadSavedState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed.records) || !parsed.records.length) return false;
    applyData(parsed);
    return true;
  } catch {
    return false;
  }
}

// ================================================================
// 요약 계산
// ================================================================

function buildSummary(records) {
  const s = { overtime: 0, dayOff: 0, annual: 0, compOff: 0, sick: 0, earlyLeave: 0, maternity: 0, pregnancyCheck: 0, longService: 0, reward: 0, pregnancyShorter: 0, localTrip: 0, outsideTrip: 0, intlTrip: 0 };
  for (const r of records) {
    if (r.category === "시간외")           s.overtime         += r.overtimeHours;
    if (r.category === "연차")             { s.annual += r.durationHours; s.dayOff += r.durationHours; }
    if (r.category === "대휴")             { s.compOff += r.durationHours; s.dayOff += r.durationHours; }
    if (r.category === "휴무")             s.dayOff           += r.durationHours;
    if (r.category === "병가")             s.sick             += r.durationHours;
    if (r.category === "조퇴")             s.earlyLeave       += r.durationHours;
    if (r.category === "산전후휴가")       s.maternity        += r.durationHours;
    if (r.category === "임산부정기검진")   s.pregnancyCheck   += r.durationHours;
    if (r.category === "근속휴가")         s.longService      += r.durationHours;
    if (r.category === "포상휴가")         s.reward           += r.durationHours;
    if (r.category === "임신기단축")       s.pregnancyShorter += r.durationHours;
    if (r.tripType === "관내출장")         s.localTrip        += 1;
    if (r.tripType === "관외출장")         s.outsideTrip      += 1;
    if (r.tripType === "국외출장")         s.intlTrip         += 1;
  }
  return s;
}

// ================================================================
// UI 렌더링
// ================================================================

function populateDepts() {
  deptSelect.innerHTML = '<option value="">부서를 선택하세요</option>';
  state.departments.forEach((d) => deptSelect.appendChild(new Option(d, d)));
}

function populateMonths() {
  monthSelect.innerHTML = '<option value="">월을 선택하세요</option>';
  sortMonths(state.months).forEach((m) => monthSelect.appendChild(new Option(m, m)));
}

function populateEmployees() {
  const key    = `${deptSelect.value}__${monthSelect.value}`;
  const empMap = state.employees.get(key) || new Map();
  empSelect.innerHTML = '<option value="">직원을 선택하세요</option>';
  [...empMap.keys()]
    .sort((a, b) => a.localeCompare(b, "ko"))
    .forEach((name) => empSelect.appendChild(new Option(name, name)));
}

/** 선택한 부서의 해당 월 레코드 */
function getMonthRecords(dept, month) {
  return state.byDeptMonth.get(`${dept}__${month}`) || [];
}

/** 선택한 부서의 해당 월까지 누적 레코드 */
function getAnnualRecords(dept, month) {
  const target = monthOrder(month);
  const result = [];
  for (const m of state.months) {
    if (monthOrder(m) <= target) result.push(...(state.byDeptMonth.get(`${dept}__${m}`) || []));
  }
  return result;
}

function renderTeamOverview() {
  const dept  = deptSelect.value;
  const month = monthSelect.value;

  if (!dept || !month) {
    $("teamDayOff").textContent   = "0시간 / 0시간";
    $("teamSickLeave").textContent = "0시간 / 0시간";
    $("teamOvertime").textContent  = "0.00시간 / 0.00시간";
    return;
  }

  const ms = buildSummary(getMonthRecords(dept, month));
  const as = buildSummary(getAnnualRecords(dept, month));

  $("teamDayOff").textContent   = fmtMA(ms.dayOff,  as.dayOff);
  $("teamSickLeave").textContent = fmtMA(ms.sick,   as.sick);
  $("teamOvertime").textContent  = fmtMA(ms.overtime, as.overtime, true);
}

function renderTeamSummary() {
  const tbody = $("teamRows");
  if (!tbody) return;
  const dept  = deptSelect.value;
  const month = monthSelect.value;

  if (!dept || !month) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">부서와 월을 선택하면 통합 데이터가 표시됩니다.</td></tr>';
    return;
  }

  const mRecords = getMonthRecords(dept, month);
  const aRecords = getAnnualRecords(dept, month);
  const names    = [...new Set(aRecords.map((r) => r.name).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));

  if (!names.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">표시할 통합 데이터가 없습니다.</td></tr>';
    return;
  }

  // 이름별 레코드 묶기
  const mByName = new Map();
  const aByName = new Map();
  for (const r of mRecords) { if (!mByName.has(r.name)) mByName.set(r.name, []); mByName.get(r.name).push(r); }
  for (const r of aRecords) { if (!aByName.has(r.name)) aByName.set(r.name, []); aByName.get(r.name).push(r); }

  const rows = names.map((name) => {
    const ms = buildSummary(mByName.get(name) || []);
    const as = buildSummary(aByName.get(name) || []);
    return `<tr>
      <td>${name}</td>
      <td>${fmtMA(ms.dayOff, as.dayOff)}</td>
      <td>${fmtMA(ms.sick, as.sick)}</td>
      <td>${fmtMA(ms.overtime, as.overtime, true)}</td>
    </tr>`;
  });

  // 팀 합계
  const mt = buildSummary(mRecords);
  const at = buildSummary(aRecords);
  const totalRow = `<tr class="team-total-row">
    <td>팀 전체</td>
    <td>${fmtMA(mt.dayOff, at.dayOff)}</td>
    <td>${fmtMA(mt.sick, at.sick)}</td>
    <td>${fmtMA(mt.overtime, at.overtime, true)}</td>
  </tr>`;

  tbody.innerHTML = rows.join("") + totalRow;
}

function renderDetails(records) {
  const tbody = $("detailRows");
  if (!tbody) return;

  if (!records.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">표시할 데이터가 없습니다.</td></tr>';
    return;
  }

  // 날짜+분류+세부 기준 그루핑
  const grouped = new Map();
  for (const r of records) {
    const key = `${r.date}__${r.category}__${r.subType}`;
    const h   = (r.overtimeHours || 0) + (r.durationHours || 0);
    if (!grouped.has(key)) grouped.set(key, { date: r.date, category: r.category, subType: r.subType, hours: 0, isOvertime: r.category === "시간외" });
    grouped.get(key).hours += h;
  }

  tbody.innerHTML = [...grouped.values()]
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((r) => `<tr>
      <td>${r.date}</td>
      <td>${r.category}</td>
      <td>${r.subType || "-"}</td>
      <td>${r.isOvertime ? `${r.hours.toFixed(2)}시간` : fmtHours(r.hours)}</td>
    </tr>`)
    .join("");
}

function updateDashboard() {
  const dept  = deptSelect.value;
  const month = monthSelect.value;
  const name  = empSelect.value;

  const records = (dept && month && name)
    ? (state.byDeptMonthName.get(`${dept}__${month}__${name}`) || [])
    : [];

  const empId = state.employees.get(`${dept}__${month}`)?.get(name) || "-";
  const s     = buildSummary(records);

  // null-safe setter
  const set = (id, val) => { const el = $(id); if (el) el.textContent = val; };

  set("empNo",             empId);
  set("overtime",          s.overtime.toFixed(2));
  set("dayOff",            fmtHours(s.dayOff));
  set("sickLeave",         fmtHours(s.sick));
  set("earlyLeave",        fmtHours(s.earlyLeave));
  set("localTrip",         String(s.localTrip));
  set("outsideTrip",       String(s.outsideTrip));
  set("internationalTrip", String(s.intlTrip));
  set("maternityLeave",    fmtHours(s.maternity));
  set("pregnancyCheckup",  fmtHours(s.pregnancyCheck));
  set("longServiceLeave",  fmtHours(s.longService));
  set("rewardLeave",       fmtHours(s.reward));
  set("pregnancyShorter",  fmtHours(s.pregnancyShorter));

  renderTeamOverview();
  renderTeamSummary();
  renderDetails(records);
}

// ================================================================
// Google Drive 불러오기
// ================================================================

function makeFetchTargets(url) {
  const id     = url.match(/\/d\/([^/?&]+)/)?.[1] || "";
  const noProto = url.replace(/^https?:\/\//, "");
  return [
    { label: "google-direct",      url,                                                                          cred: true  },
    { label: "google-authuser",    url: `${url}&authuser=0`,                                                     cred: true  },
    { label: "google-alt",         url: id ? `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx` : url, cred: true },
    { label: "corsproxy.io",       url: `https://corsproxy.io/?${encodeURIComponent(url)}`,                     cred: false },
    { label: "allorigins",         url: `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,        cred: false },
    { label: "cors.isomorphic-git", url: `https://cors.isomorphic-git.org/${noProto}`,                         cred: false },
  ];
}

async function fetchBuffer(url) {
  const targets = makeFetchTargets(url);
  const errors  = [];

  for (const t of targets) {
    try {
      const ctrl  = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 15000);

      const res = await fetch(t.url, {
        signal:      ctrl.signal,
        credentials: t.cred ? "include" : "omit",
        cache:       "no-store",
      });
      clearTimeout(timer);

      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const ct = res.headers.get("content-type") || "";
      if (ct.includes("text/html")) throw new Error("HTML 응답 (로그인 페이지 또는 권한 오류)");

      return await res.arrayBuffer();
    } catch (e) {
      errors.push(`[${t.label}] ${e.message}`);
    }
  }

  throw new Error(`모든 경로에서 실패했습니다:\n${errors.join("\n")}`);
}

async function loadFromDrive() {
  loadDriveBtn.disabled    = true;
  loadDriveBtn.textContent = "불러오는 중…";

  try {
    const buffer = await fetchBuffer(DRIVE_XLSX_URL);
    const parsed = parseWorkbook(buffer);

    if (!parsed.records.length) {
      throw new Error("파싱된 데이터가 0건입니다.\n엑셀 시트 컬럼명(날짜, 성명, 부서 등)을 확인해주세요.");
    }

    applyData(parsed);
    saveState();
    populateDepts();
    populateMonths();

    if (state.departments.length) deptSelect.value = state.departments[0];
    const sorted = sortMonths(state.months);
    if (sorted.length) monthSelect.value = sorted[0];

    populateEmployees();
    updateDashboard();

    alert(`✅ 완료! 총 ${state.records.length}건 불러왔습니다.`);
  } catch (e) {
    alert(`❌ 불러오기 실패\n\n${e.message}\n\n─────────────────\n※ 점검 사항\n1. 스프레드시트 공유: '링크가 있는 모든 사용자 - 뷰어'\n2. 크롬에 구글 계정으로 로그인된 상태인지 확인`);
  } finally {
    loadDriveBtn.disabled    = false;
    loadDriveBtn.textContent = "최신 데이터 새로고침";
  }
}

// ================================================================
// 이벤트 연결
// ================================================================

loadDriveBtn.addEventListener("click", loadFromDrive);

clearBtn.addEventListener("click", () => {
  if (!confirm("저장된 데이터를 초기화하시겠습니까?")) return;
  localStorage.removeItem(STORAGE_KEY);
  applyData({ records: [], months: [], departments: [] });
  populateDepts();
  populateMonths();
  populateEmployees();
  updateDashboard();
  alert("초기화했습니다.");
});

if (printBtn) {
  printBtn.addEventListener("click", () => window.print());
}

if (localFileBtn && fileInput) {
  localFileBtn.addEventListener("click", () => fileInput.click());

  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(buffer);
      if (!parsed.records.length) throw new Error("파싱된 데이터가 0건입니다.");
      applyData(parsed);
      saveState();
      populateDepts();
      populateMonths();
      if (state.departments.length) deptSelect.value = state.departments[0];
      const sorted = sortMonths(state.months);
      if (sorted.length) monthSelect.value = sorted[0];
      populateEmployees();
      updateDashboard();
      alert(`✅ 완료! 총 ${state.records.length}건 불러왔습니다.`);
    } catch (err) {
      alert(`❌ 로컬 파일 오류\n${err.message}`);
    } finally {
      fileInput.value = "";
    }
  });
}

deptSelect.addEventListener("change",  () => { populateEmployees(); updateDashboard(); });
monthSelect.addEventListener("change", () => { populateEmployees(); updateDashboard(); });
empSelect.addEventListener("change",   updateDashboard);

// ================================================================
// 초기화 실행
// ================================================================

if (loadSavedState()) {
  // 로컬스토리지에 저장된 데이터가 있으면 바로 표시
  populateDepts();
  populateMonths();
  if (state.departments.length) deptSelect.value = state.departments[0];
  const sorted = sortMonths(state.months);
  if (sorted.length) monthSelect.value = sorted[0];
  populateEmployees();
  updateDashboard();
} else {
  // 없으면 드라이브에서 자동 로드 시도
  loadFromDrive();
}

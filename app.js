/* =========================================================================
   รายงานเยี่ยมร้านค้าเทียบแผน — app.js
   ========================================================================= */

const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let planFile = null;
let actualFile = null;
let lastComparison = null;
let lastBatchMeta = null;

/* ---------------------- helpers: parsing values ---------------------- */

function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function toInt(v) {
  const n = toNum(v);
  return n === null ? null : Math.round(n);
}

function toStr(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === '' ? null : s;
}

// รับค่าเป็น string "dd/mm/yyyy", Date object, หรือ excel serial number -> คืน "yyyy-mm-dd"
function excelValueToISODate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    return v.toISOString().slice(0, 10);
  }
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
}

// รับค่าเป็น string "dd/mm/yyyy hh:mm:ss", Date object, หรือ excel serial number -> คืน ISO datetime + เขตเวลา +07:00
function excelValueToISODateTime(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    const pad = (n) => String(Math.floor(n)).padStart(2, '0');
    return `${d.y}-${pad(d.m)}-${pad(d.d)}T${pad(d.H || 0)}:${pad(d.M || 0)}:${pad(d.S || 0)}+07:00`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) {
    const [, d, mo, y, hh, mm, ss] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}T${hh.padStart(2, '0')}:${mm}:${(ss || '00').padStart(2, '0')}+07:00`;
  }
  return null;
}

function fmtDt(iso) {
  if (!iso) return '';
  return String(iso).replace('T', ' ').replace(/\+07:00$/, '').slice(0, 19);
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ---------------------- helpers: sheet header detection ---------------------- */

function findHeaderRowIndex(matrix, mustInclude) {
  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const row = matrix[i] || [];
    const rowSet = new Set(row.map((c) => (c === null || c === undefined ? '' : String(c).trim())));
    if (mustInclude.every((h) => rowSet.has(h))) return i;
  }
  return 0;
}

/* ---------------------- parse: ไฟล์แผนเยี่ยม ---------------------- */

function parsePlanWorkbook(wb) {
  const sheetName = wb.SheetNames.includes('CallPlanData') ? 'CallPlanData' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const headerIdx = findHeaderRowIndex(matrix, ['รหัสร้านค้า', 'วันที่เข้าเยี่ยม']);
  const rows = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: null, raw: true });

  return rows
    .map((r) => ({
      company: toStr(r['บริษัท']),
      channel: toStr(r['ช่องทาง']),
      visit_order: toInt(r['ลำดับการเข้าเยี่ยม']),
      box_name: toStr(r['ชื่อกล่อง']),
      store_code: toStr(r['รหัสร้านค้า']),
      store_name: toStr(r['ชื่อร้านค้า']),
      region: toStr(r['ภาค']),
      sales_unit: toStr(r['หน่วยขาย']),
      store_type: toStr(r['ประเภทร้านค้า']),
      employee_code: toStr(r['รหัสพนักงาน']),
      employee_firstname: toStr(r['ชื่อพนักงาน']),
      employee_lastname: toStr(r['นามสกุลพนักงาน']),
      plan_visit_date: excelValueToISODate(r['วันที่เข้าเยี่ยม']),
      latitude: toNum(r['Latitude']),
      longitude: toNum(r['Longitude']),
    }))
    .filter((r) => r.store_code && r.plan_visit_date);
}

/* ---------------------- parse: ไฟล์เยี่ยมร้านค้าจริง ---------------------- */

function parseActualWorkbook(wb) {
  const sheetName = wb.SheetNames.includes('Sheet1') ? 'Sheet1' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const headerIdx = findHeaderRowIndex(matrix, ['รหัสร้านค้า', 'วันที่เข้าเยี่ยม', 'ในแผน/นอกแผน']);
  const rows = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: null, raw: true });

  return rows
    .map((r) => ({
      seq_no: toInt(r['ลำดับที่']),
      employee_code: toStr(r['รหัสพนักงาน']),
      employee_name: toStr(r['ชื่อพนักงาน']),
      box_name: toStr(r['ชื่อกล่อง']),
      store_code: toStr(r['รหัสร้านค้า']),
      store_name: toStr(r['ชื่อร้านค้า']),
      lat_in: toNum(r['ละติจูดเข้า']),
      lng_in: toNum(r['ลองจิจูดเข้า']),
      distance_in_m: toNum(r['ระยะห่างเข้า (เมตร)']),
      time_in: excelValueToISODateTime(r['เวลาเข้า']),
      reason_in: toStr(r['เหตุผลการเข้านอกพิกัด']),
      lat_out: toNum(r['ละติจูดออก']),
      lng_out: toNum(r['ลองจิจูดออก']),
      distance_out_m: toNum(r['ระยะห่างออก (เมตร)']),
      time_out: excelValueToISODateTime(r['เวลาออก']),
      reason_out: toStr(r['เหตุผลการออกนอกพิกัด']),
      duration_minutes: toInt(r['ใช้เวลา']),
      note: toStr(r['Note']),
      source_plan_status: toStr(r['ในแผน/นอกแผน']),
      visit_date: excelValueToISODate(r['วันที่เข้าเยี่ยม']),
      photo_url: toStr(r['รูปถ่าย Check-In']),
      closed_status: toStr(r['สถานะร้านปิด/เลิกกิจการ']),
      closed_note: toStr(r['หมายเหตุสถานะร้านปิด/เลิกกิจการ']),
      closed_photo_url: toStr(r['รูปสถานะร้านปิด/เลิกกิจการ']),
      lat_record: toNum(r['ละติจูด(บันทึกพิกัด)']),
      lng_record: toNum(r['ลองจิจูด(บันทึกพิกัด)']),
      distance_record_m: toNum(r['ระยะห่างบันทึกพิกัด(เมตร)']),
      qty_whiskey_white: toNum(r['สุราขาว']),
      qty_whiskey_color: toNum(r['สุราสี']),
      qty_rtd: toNum(r['RTD']),
      qty_beer: toNum(r['เบียร์']),
      qty_oishi: toNum(r['โออิชิ']),
      qty_est: toNum(r['เอส']),
      qty_soda: toNum(r['โซดา']),
      qty_water: toNum(r['น้ำ']),
      qty_other: toNum(r['อื่นๆ']),
      qty_total: toNum(r['total']),
    }))
    .filter((r) => r.store_code && r.visit_date);
}

/* ---------------------- เปรียบเทียบแผน vs เยี่ยมจริง ---------------------- */
// จับคู่ด้วย รหัสร้านค้า + วันที่ ต้องตรงกันเป๊ะ

function computeComparison(planRows, actualRows) {
  const planMap = new Map();
  planRows.forEach((p) => {
    const key = p.store_code + '||' + p.plan_visit_date;
    if (!planMap.has(key)) planMap.set(key, []);
    planMap.get(key).push(p);
  });

  const actualMap = new Map();
  actualRows.forEach((a) => {
    const key = a.store_code + '||' + a.visit_date;
    if (!actualMap.has(key)) actualMap.set(key, []);
    actualMap.get(key).push(a);
  });

  const visited = [];
  const notVisited = [];
  planRows.forEach((p) => {
    const key = p.store_code + '||' + p.plan_visit_date;
    const matches = actualMap.get(key);
    if (matches && matches.length) {
      matches.forEach((a) => visited.push({ plan: p, actual: a }));
    } else {
      notVisited.push(p);
    }
  });

  const offPlan = [];
  actualRows.forEach((a) => {
    const key = a.store_code + '||' + a.visit_date;
    if (!planMap.has(key)) offPlan.push(a);
  });

  return { visited, notVisited, offPlan };
}

/* ---------------------- UI: log ---------------------- */

function logReset() {
  const el = document.getElementById('log');
  el.textContent = '';
  el.classList.remove('error');
}
function logAppend(msg) {
  const el = document.getElementById('log');
  el.textContent += (el.textContent ? '\n' : '') + msg;
}
function logError(msg) {
  const el = document.getElementById('log');
  el.classList.add('error');
  el.textContent += (el.textContent ? '\n' : '') + '❌ ' + msg;
}

/* ---------------------- Supabase I/O ---------------------- */

async function checkConnection() {
  const el = document.getElementById('connStatus');
  try {
    const { error } = await sb.from('upload_batches').select('id', { count: 'exact', head: true });
    if (error) throw error;
    el.textContent = '● เชื่อมต่อ Supabase สำเร็จ';
    el.className = 'status-chip ok';
  } catch (e) {
    el.textContent = '● เชื่อมต่อ Supabase ไม่สำเร็จ - ตรวจสอบ config.js และรัน supabase-schema.sql (' + (e.message || '') + ')';
    el.className = 'status-chip warn';
  }
}

async function insertChunked(table, rows, batchId, chunkSize = 300) {
  const withBatch = rows.map((r) => ({ ...r, batch_id: batchId }));
  for (let i = 0; i < withBatch.length; i += chunkSize) {
    const chunk = withBatch.slice(i, i + chunkSize);
    const { error } = await sb.from(table).insert(chunk);
    if (error) throw error;
    logAppend(`บันทึก ${table}: ${Math.min(i + chunkSize, withBatch.length)}/${withBatch.length}`);
  }
}

async function fetchAllRows(table, batchId) {
  let all = [];
  let from = 0;
  const pageSize = 1000;
  // eslint-disable-next-line no-constant-condition
  while (true) {
    const { data, error } = await sb.from(table).select('*').eq('batch_id', batchId).range(from, from + pageSize - 1);
    if (error) throw error;
    all = all.concat(data);
    if (data.length < pageSize) break;
    from += pageSize;
  }
  return all;
}

/* ---------------------- file inputs ---------------------- */

document.getElementById('file-plan').addEventListener('change', (e) => {
  planFile = e.target.files[0] || null;
  document.getElementById('fn-plan').textContent = planFile ? planFile.name : '';
  document.getElementById('dz-plan').classList.toggle('filled', !!planFile);
  updateProcessBtn();
});

document.getElementById('file-actual').addEventListener('change', (e) => {
  actualFile = e.target.files[0] || null;
  document.getElementById('fn-actual').textContent = actualFile ? actualFile.name : '';
  document.getElementById('dz-actual').classList.toggle('filled', !!actualFile);
  updateProcessBtn();
});

function updateProcessBtn() {
  document.getElementById('btn-process').disabled = !(planFile && actualFile);
}

/* ---------------------- process button ---------------------- */

document.getElementById('btn-process').addEventListener('click', async () => {
  const btn = document.getElementById('btn-process');
  btn.disabled = true;
  logReset();
  try {
    logAppend('กำลังอ่านไฟล์แผนเยี่ยม...');
    const planWb = XLSX.read(await planFile.arrayBuffer(), { type: 'array', cellDates: false });
    const planRows = parsePlanWorkbook(planWb);
    logAppend(`อ่านแผนเยี่ยมสำเร็จ: ${planRows.length} แถว`);

    logAppend('กำลังอ่านไฟล์เยี่ยมร้านค้าจริง...');
    const actualWb = XLSX.read(await actualFile.arrayBuffer(), { type: 'array', cellDates: false });
    const actualRows = parseActualWorkbook(actualWb);
    logAppend(`อ่านเยี่ยมจริงสำเร็จ: ${actualRows.length} แถว`);

    if (!planRows.length) throw new Error('ไม่พบข้อมูลในไฟล์แผนเยี่ยม (ตรวจสอบคอลัมน์ รหัสร้านค้า / วันที่เข้าเยี่ยม)');
    if (!actualRows.length) throw new Error('ไม่พบข้อมูลในไฟล์เยี่ยมร้านค้าจริง (ตรวจสอบคอลัมน์ รหัสร้านค้า / วันที่เข้าเยี่ยม)');

    const allDates = [...planRows.map((r) => r.plan_visit_date), ...actualRows.map((r) => r.visit_date)]
      .filter(Boolean)
      .sort();

    const batchMetaInsert = {
      company: planRows[0]?.company || null,
      channel: planRows[0]?.channel || null,
      region: planRows[0]?.region || null,
      period_start: allDates[0] || null,
      period_end: allDates[allDates.length - 1] || null,
      plan_filename: planFile.name,
      actual_filename: actualFile.name,
      plan_row_count: planRows.length,
      actual_row_count: actualRows.length,
    };

    logAppend('กำลังสร้างชุดข้อมูล (batch) ใน Supabase...');
    const { data: batchData, error: batchErr } = await sb.from('upload_batches').insert(batchMetaInsert).select().single();
    if (batchErr) throw batchErr;
    const batchId = batchData.id;

    logAppend('กำลังบันทึกแผนเยี่ยมลง Supabase...');
    await insertChunked('visit_plans', planRows, batchId);

    logAppend('กำลังบันทึกเยี่ยมจริงลง Supabase...');
    await insertChunked('actual_visits', actualRows, batchId);

    logAppend('กำลังคำนวณเปรียบเทียบ...');
    const cmp = computeComparison(planRows, actualRows);
    lastComparison = cmp;
    lastBatchMeta = batchData;
    renderResults(batchData, cmp);

    logAppend('เสร็จสิ้น ✅ บันทึกและสรุปผลเรียบร้อย');
    loadHistory();
  } catch (err) {
    console.error(err);
    logError(err.message || String(err));
  } finally {
    btn.disabled = !(planFile && actualFile);
  }
});

/* ---------------------- results rendering ---------------------- */

function renderResults(batchMeta, cmp) {
  document.getElementById('results-section').style.display = '';
  document.getElementById('stat-visited').textContent = cmp.visited.length;
  document.getElementById('stat-notvisited').textContent = cmp.notVisited.length;
  document.getElementById('stat-offplan').textContent = cmp.offPlan.length;
  document.getElementById('period-hint').textContent = `— ${batchMeta.period_start || ''} ถึง ${batchMeta.period_end || ''}`;
  document.querySelectorAll('.tab').forEach((x) => x.classList.remove('active'));
  document.querySelector('.tab[data-tab="visited"]').classList.add('active');
  renderTable('visited');
}

function renderTable(kind) {
  const container = document.getElementById('table-container');
  if (!lastComparison) {
    container.innerHTML = '<div class="empty-note">ยังไม่มีข้อมูล</div>';
    return;
  }
  let rows, headers, mapper;
  if (kind === 'visited') {
    rows = lastComparison.visited;
    headers = ['รหัสร้านค้า', 'ชื่อร้านค้า', 'ภาค/หน่วยขาย', 'พนักงานตามแผน', 'วันที่แผน', 'พนักงานที่เยี่ยมจริง', 'เวลาเข้า', 'เวลาออก'];
    mapper = ({ plan, actual }) => [
      plan.store_code,
      plan.store_name,
      `${plan.region || ''}/${plan.sales_unit || ''}`,
      `${plan.employee_firstname || ''} ${plan.employee_lastname || ''}`.trim(),
      plan.plan_visit_date,
      actual.employee_name || '',
      fmtDt(actual.time_in),
      fmtDt(actual.time_out),
    ];
  } else if (kind === 'notvisited') {
    rows = lastComparison.notVisited;
    headers = ['รหัสร้านค้า', 'ชื่อร้านค้า', 'ภาค/หน่วยขาย', 'พนักงานตามแผน', 'วันที่แผน'];
    mapper = (p) => [
      p.store_code,
      p.store_name,
      `${p.region || ''}/${p.sales_unit || ''}`,
      `${p.employee_firstname || ''} ${p.employee_lastname || ''}`.trim(),
      p.plan_visit_date,
    ];
  } else {
    rows = lastComparison.offPlan;
    headers = ['รหัสร้านค้า', 'ชื่อร้านค้า', 'พนักงานที่เยี่ยม', 'วันที่เยี่ยม', 'เวลาเข้า', 'เวลาออก', 'หมายเหตุ'];
    mapper = (a) => [a.store_code, a.store_name, a.employee_name || '', a.visit_date, fmtDt(a.time_in), fmtDt(a.time_out), a.note || ''];
  }

  if (!rows.length) {
    container.innerHTML = '<div class="empty-note">ไม่มีข้อมูลในหมวดนี้</div>';
    return;
  }

  const maxRows = 500;
  const shown = rows.slice(0, maxRows);
  let html = '<table class="data"><thead><tr>' + headers.map((h) => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  shown.forEach((r) => {
    html += '<tr>' + mapper(r).map((c) => `<td>${escapeHtml(c ?? '')}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  if (rows.length > maxRows) {
    html += `<div class="empty-note">แสดง ${maxRows} จาก ${rows.length} แถว — ดาวน์โหลด Excel เพื่อดูข้อมูลทั้งหมด</div>`;
  }
  container.innerHTML = html;
}

document.querySelectorAll('.tab').forEach((t) => {
  t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((x) => x.classList.remove('active'));
    t.classList.add('active');
    renderTable(t.dataset.tab);
  });
});

/* ---------------------- download report ---------------------- */

document.getElementById('btn-download').addEventListener('click', () => {
  if (!lastComparison || !lastBatchMeta) return;
  downloadReport(lastBatchMeta, lastComparison);
});

function downloadReport(batchMeta, cmp) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    { A: 'บริษัท', B: batchMeta.company || '' },
    { A: 'ช่องทาง', B: batchMeta.channel || '' },
    { A: 'ช่วงเวลา', B: `${batchMeta.period_start || ''} ถึง ${batchMeta.period_end || ''}` },
    { A: 'ไฟล์แผนเยี่ยม', B: batchMeta.plan_filename || '' },
    { A: 'ไฟล์เยี่ยมจริง', B: batchMeta.actual_filename || '' },
    { A: '', B: '' },
    { A: 'จำนวนครั้งที่เยี่ยมตามแผน (ร้าน+วันที่ตรงกัน)', B: cmp.visited.length },
    { A: 'จำนวนแผนที่ยังไม่ถูกเยี่ยม (ตามวันที่แผนกำหนด)', B: cmp.notVisited.length },
    { A: 'จำนวนครั้งที่เยี่ยมนอกแผน/ผิดวันที่แผน', B: cmp.offPlan.length },
  ];
  const wsSummary = XLSX.utils.json_to_sheet(summaryRows, { skipHeader: true });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'สรุป');

  const visitedRows = cmp.visited.map(({ plan, actual }) => ({
    รหัสร้านค้า: plan.store_code,
    ชื่อร้านค้า: plan.store_name,
    ภาค: plan.region,
    หน่วยขาย: plan.sales_unit,
    ประเภทร้านค้า: plan.store_type,
    พนักงานตามแผน: `${plan.employee_firstname || ''} ${plan.employee_lastname || ''}`.trim(),
    วันที่ตามแผน: plan.plan_visit_date,
    พนักงานที่เยี่ยมจริง: actual.employee_name,
    วันที่เยี่ยมจริง: actual.visit_date,
    เวลาเข้า: fmtDt(actual.time_in),
    เวลาออก: fmtDt(actual.time_out),
    'ใช้เวลา (นาที)': actual.duration_minutes,
    หมายเหตุ: actual.note,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(visitedRows), 'เยี่ยมตามแผน');

  const notVisitedRows = cmp.notVisited.map((p) => ({
    รหัสร้านค้า: p.store_code,
    ชื่อร้านค้า: p.store_name,
    ภาค: p.region,
    หน่วยขาย: p.sales_unit,
    ประเภทร้านค้า: p.store_type,
    พนักงานตามแผน: `${p.employee_firstname || ''} ${p.employee_lastname || ''}`.trim(),
    วันที่ตามแผน: p.plan_visit_date,
    ลำดับการเข้าเยี่ยมตามแผน: p.visit_order,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(notVisitedRows), 'ยังไม่เยี่ยม');

  const offPlanRows = cmp.offPlan.map((a) => ({
    รหัสร้านค้า: a.store_code,
    ชื่อร้านค้า: a.store_name,
    พนักงานที่เยี่ยม: a.employee_name,
    ชื่อกล่อง: a.box_name,
    วันที่เยี่ยมจริง: a.visit_date,
    เวลาเข้า: fmtDt(a.time_in),
    เวลาออก: fmtDt(a.time_out),
    'ใช้เวลา (นาที)': a.duration_minutes,
    สถานะจากไฟล์ต้นฉบับ: a.source_plan_status,
    หมายเหตุ: a.note,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(offPlanRows), 'นอกแผน');

  const fname = `รายงานเยี่ยมร้านค้าเทียบแผน_${batchMeta.period_start || ''}_${batchMeta.period_end || ''}.xlsx`.replace(/\s+/g, '');
  XLSX.writeFile(wb, fname);
}

/* ---------------------- history ---------------------- */

async function loadHistory() {
  const container = document.getElementById('history-list');
  try {
    const { data, error } = await sb.from('upload_batches').select('*').order('uploaded_at', { ascending: false }).limit(20);
    if (error) throw error;
    if (!data.length) {
      container.innerHTML = '<div class="empty-note">ยังไม่มีประวัติการอัปโหลด</div>';
      return;
    }
    container.innerHTML = data
      .map(
        (b) => `
      <div class="history-item">
        <div class="meta">
          <div class="period">${b.period_start || '?'} ถึง ${b.period_end || '?'} ${b.company ? '· ' + escapeHtml(b.company) : ''} ${
          b.channel ? '· ' + escapeHtml(b.channel) : ''
        }</div>
          <div class="sub">อัปโหลดเมื่อ ${new Date(b.uploaded_at).toLocaleString('th-TH')} · แผน ${b.plan_row_count} แถว / เยี่ยมจริง ${b.actual_row_count} แถว</div>
        </div>
        <div class="actions">
          <button class="btn secondary" data-batch="${b.id}">ดูสรุป + ดาวน์โหลด</button>
        </div>
      </div>
    `
      )
      .join('');
    container.querySelectorAll('button[data-batch]').forEach((btn) => {
      const meta = data.find((b) => b.id === btn.dataset.batch);
      btn.addEventListener('click', () => reloadBatch(btn.dataset.batch, meta));
    });
  } catch (e) {
    container.innerHTML = `<div class="empty-note">โหลดประวัติไม่สำเร็จ: ${escapeHtml(e.message)}</div>`;
  }
}

async function reloadBatch(batchId, batchMeta) {
  logReset();
  logAppend('กำลังดึงข้อมูลจาก Supabase สำหรับชุดนี้...');
  try {
    const [planRows, actualRows] = await Promise.all([fetchAllRows('visit_plans', batchId), fetchAllRows('actual_visits', batchId)]);
    const cmp = computeComparison(planRows, actualRows);
    lastComparison = cmp;
    lastBatchMeta = batchMeta;
    renderResults(batchMeta, cmp);
    logAppend('โหลดสำเร็จ ✅');
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (e) {
    logError(e.message);
  }
}

/* ---------------------- init ---------------------- */

checkConnection();
loadHistory();

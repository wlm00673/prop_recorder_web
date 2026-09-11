/* ============================================================
 * 螺旋桨测试数据记录器 —— 纯前端版
 * 数据存在浏览器 localStorage，不上传任何服务器；
 * 导出 Excel 为客户端自建 .xlsx（无第三方库、无需联网）。
 * ============================================================ */
(function () {
  'use strict';

  /* ---------------- 基础选项 ---------------- */
  const USAGES = ['固定翼', '多旋翼'];
  const POWERS = ['电动', '油动'];
  const MATERIALS = ['尼龙玻纤', '碳纤维', '榉木', '尼龙', '其他'];

  const SIZES_FIXED = ['5×3', '6×3', '6×4', '7×4', '7×5', '7×6', '8×4', '8×5', '8×6', '9×4', '9×5', '9×6', '9×7',
    '10×4', '10×5', '10×6', '10×7', '10×8', '11×5', '11×5.5', '11×6', '11×7', '11×8',
    '12×4', '12×5', '12×6', '12×8', '13×6', '13×8', '14×5', '14×6', '14×7', '14×8', '14×10',
    '15×5', '15×6', '15×8', '15×10', '16×5', '16×6', '16×8', '16×10', '17×6', '17×8',
    '18×6', '18×8', '18×10', '19×8', '20×8', '20×10', '22×10', '24×10', '24×12'];

  const SIZES_MULTI = ['8×4.5', '9×4.5', '9×5', '10×4.5', '10×5', '11×5', '12×4', '12×4.5', '13×4.5', '13×5.5',
    '14×4.8', '15×5', '15×5.5', '16×5.5', '17×5.8', '18×5.5', '18×6.1', '20×6.5', '21×6.6',
    '22×7', '24×7.2', '26×8.5', '28×9.2', '30×10'];

  const CUSTOM_SIZE = '✏️ 自定义（手动填直径 / 螺距）';
  const MANUAL = '✏️ 手动输入 / 未收录（在下方填写）';

  /* ---------------- 内置电机库（参考数据） ---------------- */
  const BUILTIN_MOTORS = [
    // 朗宇 SunnySky
    { brand: '朗宇', model: 'X2212 III', kv: 980, type: '多旋翼/固定翼', props: '8~10寸', maxA: 20, weight: 55, note: '经典练习机 / 四轴' },
    { brand: '朗宇', model: 'X2212 III', kv: 1400, type: '多旋翼/固定翼', props: '7~8寸', maxA: 25, weight: 55, note: '' },
    { brand: '朗宇', model: 'X2216 III', kv: 880, type: '多旋翼/固定翼', props: '10~11寸', maxA: 28, weight: 72, note: '' },
    { brand: '朗宇', model: 'X2216 III', kv: 1100, type: '多旋翼/固定翼', props: '9~10寸', maxA: 32, weight: 72, note: '' },
    { brand: '朗宇', model: 'X2814 III', kv: 900, type: '固定翼', props: '11~13寸', maxA: 40, weight: 105, note: '' },
    { brand: '朗宇', model: 'X2814 III', kv: 1000, type: '固定翼', props: '10~12寸', maxA: 45, weight: 105, note: '' },
    { brand: '朗宇', model: 'X2820 III', kv: 1050, type: '多旋翼', props: '10~13寸', maxA: 45, weight: 110, note: 'CUADC 专用（用户指定）' },
    { brand: '朗宇', model: 'X2820 III', kv: 1250, type: '固定翼', props: '10~12寸', maxA: 50, weight: 110, note: '' },
    { brand: '朗宇', model: 'X3120 III', kv: 900, type: '固定翼/多旋翼', props: '12~14寸', maxA: 55, weight: 150, note: '' },
    { brand: '朗宇', model: 'X3520 III', kv: 520, type: '固定翼', props: '15~17寸', maxA: 70, weight: 200, note: '' },
    { brand: '朗宇', model: 'X3520 III', kv: 700, type: '固定翼', props: '13~15寸', maxA: 80, weight: 200, note: '' },
    { brand: '朗宇', model: 'X3520 III', kv: 850, type: '固定翼', props: '12~14寸', maxA: 85, weight: 200, note: '' },
    { brand: '朗宇', model: 'X4120 III', kv: 465, type: '固定翼', props: '16~19寸', maxA: 90, weight: 280, note: '' },
    { brand: '朗宇', model: 'X4120 III', kv: 650, type: '固定翼', props: '14~16寸', maxA: 100, weight: 280, note: '' },
    // T-Motor
    { brand: 'T-Motor', model: 'MN3110', kv: 470, type: '多旋翼', props: '15~17寸', maxA: 20, weight: 95, note: '' },
    { brand: 'T-Motor', model: 'MN3110', kv: 700, type: '多旋翼', props: '13~15寸', maxA: 25, weight: 95, note: '' },
    { brand: 'T-Motor', model: 'MN3508', kv: 380, type: '多旋翼', props: '15~17寸', maxA: 25, weight: 130, note: '' },
    { brand: 'T-Motor', model: 'MN3508', kv: 415, type: '多旋翼', props: '15~17寸', maxA: 28, weight: 130, note: '' },
    { brand: 'T-Motor', model: 'MN3508', kv: 580, type: '多旋翼', props: '13~15寸', maxA: 30, weight: 130, note: '' },
    { brand: 'T-Motor', model: 'MN3508', kv: 700, type: '多旋翼', props: '12~14寸', maxA: 32, weight: 130, note: '' },
    { brand: 'T-Motor', model: 'MN4010', kv: 370, type: '多旋翼', props: '17~20寸', maxA: 35, weight: 175, note: '' },
    { brand: 'T-Motor', model: 'MN4010', kv: 475, type: '多旋翼', props: '15~18寸', maxA: 40, weight: 175, note: '' },
    { brand: 'T-Motor', model: 'MN4014', kv: 330, type: '多旋翼', props: '18~22寸', maxA: 45, weight: 240, note: '' },
    { brand: 'T-Motor', model: 'MN4014', kv: 400, type: '多旋翼', props: '16~20寸', maxA: 50, weight: 240, note: '' },
    { brand: 'T-Motor', model: 'MN4120', kv: 320, type: '多旋翼/长航时固定翼', props: '20~24寸', maxA: 55, weight: 300, note: '' },
    { brand: 'T-Motor', model: 'MN4120', kv: 400, type: '多旋翼/长航时固定翼', props: '18~22寸', maxA: 60, weight: 300, note: '' },
    { brand: 'T-Motor', model: 'MN4120', kv: 465, type: '多旋翼/长航时固定翼', props: '16~20寸', maxA: 65, weight: 300, note: '' },
    { brand: 'T-Motor', model: 'MN5212', kv: 340, type: '多旋翼/重载', props: '22~28寸', maxA: 80, weight: 480, note: '' },
    { brand: 'T-Motor', model: 'AT2306', kv: 2300, type: '多旋翼(FPV)', props: '5~6寸', maxA: 35, weight: 30, note: '' },
    { brand: 'T-Motor', model: 'F60 PRO V', kv: 2200, type: '多旋翼(FPV)', props: '5~6寸', maxA: 45, weight: 36, note: '' },
    // 新西达 XXD
    { brand: '新西达', model: 'A2212', kv: 1000, type: '固定翼', props: '9~10寸', maxA: 20, weight: 50, note: '经典练习机' },
    { brand: '新西达', model: 'A2212', kv: 1400, type: '固定翼', props: '8~9寸', maxA: 22, weight: 50, note: '' },
    { brand: '新西达', model: 'A2212', kv: 2200, type: '固定翼/多旋翼', props: '6~7寸', maxA: 25, weight: 50, note: '' },
    // 银燕 EMAX
    { brand: '银燕', model: 'MT2213', kv: 935, type: '多旋翼', props: '9~10寸', maxA: 20, weight: 55, note: 'DJI 2212 级别' },
    { brand: '银燕', model: 'GT2215', kv: 1100, type: '固定翼', props: '9~11寸', maxA: 30, weight: 70, note: '' },
    { brand: '银燕', model: 'MT3510', kv: 600, type: '多旋翼', props: '14~16寸', maxA: 30, weight: 120, note: '' }
  ];

  /* ---------------- 内置电调库 ---------------- */
  const BUILTIN_ESCS = [
    { brand: 'T-Motor', model: '66A ESC', amp: 66, peak: 80, volt: '2~6S', bec: '5V/3A(参考)', weight: '', note: '用户指定预置电调' },
    { brand: 'T-Motor', model: 'ALPHA 60A', amp: 60, peak: 80, volt: '3~6S', bec: '5V/5A', weight: 50, note: '' },
    { brand: 'T-Motor', model: 'FLAME 60A', amp: 60, peak: 80, volt: '2~6S', bec: '5V/3A', weight: 45, note: '' },
    { brand: 'T-Motor', model: 'AIR 40A', amp: 40, peak: 60, volt: '2~4S', bec: '5V/3A', weight: 26, note: '' },
    { brand: '好盈', model: '天行者 Skywalker 40A', amp: 40, peak: 60, volt: '2~4S', bec: '5V/2A', weight: 37, note: '固定翼常用' },
    { brand: '好盈', model: '天行者 Skywalker 60A', amp: 60, peak: 80, volt: '2~6S', bec: '5V/3A', weight: 60, note: '' },
    { brand: '好盈', model: '天行者 Skywalker 80A', amp: 80, peak: 100, volt: '2~6S', bec: '5V/3A', weight: 82, note: '' },
    { brand: '好盈', model: 'Platinum 60A V4', amp: 60, peak: 80, volt: '3~6S', bec: '5V/7A', weight: 63, note: '固定翼 / 直升机' },
    { brand: '中特威', model: 'ZTW Beatles 60A', amp: 60, peak: 80, volt: '2~6S', bec: '5V/3A', weight: 65, note: '' },
    { brand: '新西达', model: 'XXD 30A', amp: 30, peak: 40, volt: '2~4S', bec: '5V/2A', weight: 26, note: '' },
    { brand: '新西达', model: 'XXD 40A', amp: 40, peak: 50, volt: '2~4S', bec: '5V/3A', weight: 32, note: '' },
    { brand: '银燕', model: 'EMAX BLHeli 30A', amp: 30, peak: 40, volt: '2~4S', bec: '无', weight: 10, note: '多旋翼' },
    { brand: '朗宇', model: 'SunnySky 40A', amp: 40, peak: 60, volt: '2~6S', bec: '5V/3A', weight: 40, note: '' }
  ];

  /* ---------------- 记录字段 ---------------- */
  const COLUMNS = ['记录ID', '记录时间', '用途', '动力类型', '材质', '直径(inch)', '螺距(inch)', '规格',
    '电机品牌', '电机型号', 'KV', '电调品牌', '电调型号', '电调额定电流(A)', '电池',
    '拉力(kg)', '电流(A)', '转速(rpm)', '电压(V)', '油门(%)', '环境温度(℃)',
    '测试者', '测试地点', '备注'];

  const LS = { records: 'prop_records_v1', motors: 'prop_motors_v1', escs: 'prop_escs_v1' };

  /* ============================================================
   *  存储
   * ============================================================ */
  const store = {
    load(key, fallback) {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : fallback;
      } catch (e) { return fallback; }
    },
    save(key, val) {
      try { localStorage.setItem(key, JSON.stringify(val)); return true; }
      catch (e) { return false; }
    }
  };

  let records = store.load(LS.records, []);
  let customMotors = store.load(LS.motors, []);
  let customEscs = store.load(LS.escs, []);

  function persistRecords() {
    if (!store.save(LS.records, records)) toast('浏览器存储写入失败（可能是隐私模式或空间已满）', true);
  }
  function persistMotors() { store.save(LS.motors, customMotors); }
  function persistEscs() { store.save(LS.escs, customEscs); }

  const allMotors = () => BUILTIN_MOTORS.map(m => Object.assign({ custom: false }, m))
    .concat(customMotors.map(m => Object.assign({ custom: true }, m)));
  const allEscs = () => BUILTIN_ESCS.map(e => Object.assign({ custom: false }, e))
    .concat(customEscs.map(e => Object.assign({ custom: true }, e)));

  /* ============================================================
   *  小工具
   * ============================================================ */
  const $ = id => document.getElementById(id);

  function fmtNum(v) {
    const n = Number(v);
    if (!isFinite(n)) return '';
    return String(parseFloat(n.toFixed(2)));
  }
  function nowStr() {
    const d = new Date(), p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }
  function stamp() {
    const d = new Date(), p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
  }
  function nextId() {
    let max = 0;
    records.forEach(r => {
      const m = String(r['记录ID'] || '').match(/(\d+)/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    return 'P' + String(max + 1).padStart(4, '0');
  }
  function valueOrNull(id) {
    const el = $(id);
    if (!el) return null;
    const raw = String(el.value || '').trim();
    if (raw === '') return null;
    const n = Number(raw);
    return isFinite(n) ? n : null;
  }
  function cellText(v) {
    if (v === null || v === undefined) return '';
    return String(v);
  }
  let toastTimer = null;
  function toast(msg, isErr) {
    const t = $('toast');
    if (!t) return;
    t.textContent = msg;
    t.className = 'toast show' + (isErr ? ' err' : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.className = 'toast'; }, 2400);
  }

  /* ============================================================
   *  表单
   * ============================================================ */
  function fillSelect(el, items, selected) {
    el.innerHTML = '';
    items.forEach(text => {
      const o = document.createElement('option');
      o.value = text;
      o.textContent = text;
      el.appendChild(o);
    });
    if (selected !== undefined && items.indexOf(selected) >= 0) el.value = selected;
  }

  function refreshSizeOptions() {
    const usage = $('usage').value;
    const list = (usage === '多旋翼' ? SIZES_MULTI : SIZES_FIXED).concat([CUSTOM_SIZE]);
    fillSelect($('sizeSel'), list);
    if (list.length) applySize(list[0]);
  }

  function applySize(spec) {
    if (spec === CUSTOM_SIZE) { $('propD').focus(); return; }
    const parts = spec.split('×');
    if (parts.length === 2) {
      $('propD').value = parts[0];
      $('propP').value = parts[1];
    }
  }

  function refreshMotorOptions() {
    const list = allMotors().map(m => `${m.brand} ${m.model} ${m.kv || ''}KV`.trim());
    fillSelect($('motorSel'), [MANUAL].concat(list));
    $('motorSel').value = MANUAL;
  }
  function refreshEscOptions() {
    const list = allEscs().map(e => `${e.brand} ${e.model}（${e.amp || '-'}A）`);
    fillSelect($('escSel'), [MANUAL].concat(list));
    $('escSel').value = MANUAL;
  }

  function onMotorPick() {
    const idx = $('motorSel').selectedIndex;
    if (idx <= 0) return;
    const m = allMotors()[idx - 1];
    $('motorBrand').value = m.brand || '';
    $('motorModel').value = m.model || '';
    $('motorKv').value = (m.kv === 0 || m.kv) ? m.kv : '';
  }
  function onEscPick() {
    const idx = $('escSel').selectedIndex;
    if (idx <= 0) return;
    const e = allEscs()[idx - 1];
    $('escBrand').value = e.brand || '';
    $('escModel').value = e.model || '';
    $('escAmp').value = (e.amp === 0 || e.amp) ? e.amp : '';
  }

  function rememberMotor(rec) {
    if (!rec['电机型号']) return;
    const exists = allMotors().some(m => m.brand === rec['电机品牌'] && m.model === rec['电机型号']);
    if (exists) return;
    customMotors.push({ brand: rec['电机品牌'] || '自定义', model: rec['电机型号'], kv: rec['KV'] || '', type: '', props: '', maxA: '', weight: '', note: '录入时自动记住' });
    persistMotors();
    refreshMotorOptions();
  }
  function rememberEsc(rec) {
    if (!rec['电调型号']) return;
    const exists = allEscs().some(e => e.brand === rec['电调品牌'] && e.model === rec['电调型号']);
    if (exists) return;
    customEscs.push({ brand: rec['电调品牌'] || '自定义', model: rec['电调型号'], amp: rec['电调额定电流(A)'] || '', peak: '', volt: '', bec: '', weight: '', note: '录入时自动记住' });
    persistEscs();
    refreshEscOptions();
  }

  /* ============================================================
   *  保存记录
   * ============================================================ */
  function buildRecord() {
    const d = valueOrNull('propD');
    const p = valueOrNull('propP');
    return {
      '记录ID': nextId(),
      '记录时间': nowStr(),
      '用途': $('usage').value,
      '动力类型': $('power').value,
      '材质': $('material').value,
      '直径(inch)': d,
      '螺距(inch)': p,
      '规格': (d !== null && p !== null) ? `${fmtNum(d)}×${fmtNum(p)}` : '',
      '电机品牌': $('motorBrand').value.trim(),
      '电机型号': $('motorModel').value.trim(),
      'KV': valueOrNull('motorKv'),
      '电调品牌': $('escBrand').value.trim(),
      '电调型号': $('escModel').value.trim(),
      '电调额定电流(A)': valueOrNull('escAmp'),
      '电池': $('battery').value.trim(),
      '拉力(kg)': valueOrNull('thrust'),
      '电流(A)': valueOrNull('current'),
      '转速(rpm)': valueOrNull('rpm'),
      '电压(V)': valueOrNull('voltage'),
      '油门(%)': valueOrNull('throttle'),
      '环境温度(℃)': valueOrNull('temp'),
      '测试者': $('tester').value.trim(),
      '测试地点': $('location').value.trim(),
      '备注': $('note').value.trim()
    };
  }

  function saveRecord() {
    const thrust = valueOrNull('thrust');
    if (thrust === null || thrust <= 0) {
      toast('请先填写「螺旋桨拉力 (kg)」', true);
      $('thrust').focus();
      return;
    }
    const rec = buildRecord();
    records.push(rec);
    rememberMotor(rec);
    rememberEsc(rec);
    persistRecords();
    renderAll();
    ['thrust', 'current', 'rpm', 'voltage', 'throttle', 'temp'].forEach(id => { $(id).value = ''; });
    $('thrust').focus();
    toast(`已保存 ${rec['记录ID']} ｜ ${rec['规格']} ｜ ${thrust} kg`);
  }

  /* ============================================================
   *  渲染：记录表 / 统计 / 汇总
   * ============================================================ */
  function filteredRecords() {
    const kw = $('filterKeyword').value.trim().toLowerCase();
    const usage = $('filterUsage').value;
    const material = $('filterMaterial').value;
    const motor = $('filterMotor').value;
    return records.filter(r => {
      if (usage && usage !== '全部用途' && r['用途'] !== usage) return false;
      if (material && material !== '全部材质' && r['材质'] !== material) return false;
      if (motor && motor !== '全部电机' && r['电机型号'] !== motor) return false;
      if (kw) {
        const text = COLUMNS.map(c => cellText(r[c])).join(' ').toLowerCase();
        if (text.indexOf(kw) < 0) return false;
      }
      return true;
    });
  }

  function renderRecords() {
    const body = $('recordsBody');
    const list = filteredRecords().slice().reverse();
    body.innerHTML = '';
    $('recordsEmpty').style.display = list.length ? 'none' : 'block';
    $('recordsEmpty').textContent = records.length ? '没有符合筛选条件的记录。' : '还没有记录，先在上面录入一条吧。';
    const frag = document.createDocumentFragment();
    list.forEach(r => {
      const tr = document.createElement('tr');
      const cells = [
        r['记录ID'], r['记录时间'], r['用途'], r['规格'], r['材质'],
        [r['电机品牌'], r['电机型号']].filter(Boolean).join(' '),
        [r['电调品牌'], r['电调型号']].filter(Boolean).join(' '),
        cellText(r['拉力(kg)']), cellText(r['电流(A)']), cellText(r['转速(rpm)']), cellText(r['备注'])
      ];
      cells.forEach((v, i) => {
        const td = document.createElement('td');
        td.textContent = cellText(v);
        if (i === 7 || i === 8 || i === 9) td.className = 'num';
        tr.appendChild(td);
      });
      const td = document.createElement('td');
      const btn = document.createElement('button');
      btn.className = 'del';
      btn.textContent = '删除';
      btn.title = '删除这条记录';
      btn.dataset.id = r['记录ID'];
      td.appendChild(btn);
      tr.appendChild(td);
      frag.appendChild(tr);
    });
    body.appendChild(frag);
  }

  function renderStats() {
    $('statTotal').textContent = records.length;
    $('statProp').textContent = new Set(records.map(r => r['规格']).filter(Boolean)).size;
    $('statMotor').textContent = new Set(records.map(r => r['电机型号']).filter(Boolean)).size;
    const thrusts = records.map(r => Number(r['拉力(kg)'])).filter(v => isFinite(v));
    $('statMax').textContent = thrusts.length ? Math.max.apply(null, thrusts).toFixed(2) : '—';
  }

  function renderSummary() {
    const map = new Map();
    records.forEach(r => {
      const key = `${r['电机品牌'] || '未填'} ${r['电机型号'] || ''}`.trim();
      if (!key) return;
      if (!map.has(key)) map.set(key, { brand: r['电机品牌'] || '', model: r['电机型号'] || '', n: 0, thrusts: [], currents: [] });
      const g = map.get(key);
      g.n++;
      const t = Number(r['拉力(kg)']); if (isFinite(t)) g.thrusts.push(t);
      const c = Number(r['电流(A)']); if (isFinite(c)) g.currents.push(c);
    });
    const rows = Array.from(map.values()).sort((a, b) => Math.max.apply(null, b.thrusts.concat([0])) - Math.max.apply(null, a.thrusts.concat([0])));
    const body = $('summaryBody');
    body.innerHTML = '';
    if (!rows.length) { body.innerHTML = '<tr><td colspan="6" class="empty">暂无数据</td></tr>'; return; }
    const avg = a => a.length ? (a.reduce((x, y) => x + y, 0) / a.length).toFixed(2) : '—';
    rows.forEach(g => {
      const tr = document.createElement('tr');
      [g.brand, g.model, g.n, g.thrusts.length ? Math.max.apply(null, g.thrusts).toFixed(2) : '—', avg(g.thrusts), avg(g.currents)]
        .forEach(v => { const td = document.createElement('td'); td.textContent = cellText(v); tr.appendChild(td); });
      body.appendChild(tr);
    });
  }

  function refreshFilterOptions() {
    const cur = { usage: $('filterUsage').value, material: $('filterMaterial').value, motor: $('filterMotor').value };
    const usages = Array.from(new Set(records.map(r => r['用途']).filter(Boolean)));
    const mats = Array.from(new Set(records.map(r => r['材质']).filter(Boolean)));
    const motors = Array.from(new Set(records.map(r => r['电机型号']).filter(Boolean)));
    fillSelect($('filterUsage'), ['全部用途'].concat(usages), '全部用途');
    fillSelect($('filterMaterial'), ['全部材质'].concat(mats), '全部材质');
    fillSelect($('filterMotor'), ['全部电机'].concat(motors), '全部电机');
    $('filterUsage').value = usages.indexOf(cur.usage) >= 0 ? cur.usage : '全部用途';
    $('filterMaterial').value = mats.indexOf(cur.material) >= 0 ? cur.material : '全部材质';
    $('filterMotor').value = motors.indexOf(cur.motor) >= 0 ? cur.motor : '全部电机';
  }

  function renderLibraries() {
    const mb = $('motorLibBody');
    mb.innerHTML = '';
    allMotors().forEach((m, i) => {
      const tr = document.createElement('tr');
      [m.custom ? '自定义' : '内置', m.brand, m.model, m.kv, m.type, m.props, m.maxA, m.weight, m.note]
        .forEach(v => { const td = document.createElement('td'); td.textContent = cellText(v); tr.appendChild(td); });
      const td = document.createElement('td');
      if (m.custom) {
        const btn = document.createElement('button');
        btn.className = 'del';
        btn.textContent = '删除';
        btn.dataset.idx = i - BUILTIN_MOTORS.length;
        td.appendChild(btn);
      }
      tr.appendChild(td);
      mb.appendChild(tr);
    });

    const eb = $('escLibBody');
    eb.innerHTML = '';
    allEscs().forEach((e, i) => {
      const tr = document.createElement('tr');
      [e.custom ? '自定义' : '内置', e.brand, e.model, e.amp, e.peak, e.volt, e.bec, e.weight, e.note]
        .forEach(v => { const td = document.createElement('td'); td.textContent = cellText(v); tr.appendChild(td); });
      const td = document.createElement('td');
      if (e.custom) {
        const btn = document.createElement('button');
        btn.className = 'del';
        btn.textContent = '删除';
        btn.dataset.idx = i - BUILTIN_ESCS.length;
        td.appendChild(btn);
      }
      tr.appendChild(td);
      eb.appendChild(tr);
    });
  }

  function renderAll() {
    refreshFilterOptions();
    renderRecords();
    renderStats();
    renderSummary();
    renderLibraries();
  }

  /* ============================================================
   *  导出：CSV / JSON
   * ============================================================ */
  function downloadBytes(filename, bytes, mime) {
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
  }

  function exportCsv() {
    if (!records.length) { toast('还没有记录', true); return; }
    const lines = [COLUMNS.join(',')];
    records.forEach(r => {
      lines.push(COLUMNS.map(c => {
        const v = cellText(r[c]);
        return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
      }).join(','));
    });
    downloadBytes(`螺旋桨测试数据_${stamp()}.csv`, new TextEncoder().encode('\uFEFF' + lines.join('\r\n')), 'text/csv');
    toast('CSV 已导出');
  }

  function exportJson() {
    const data = { version: 1, exportedAt: nowStr(), records: records, customMotors: customMotors, customEscs: customEscs };
    downloadBytes(`螺旋桨测试数据_备份_${stamp()}.json`, new TextEncoder().encode(JSON.stringify(data, null, 2)), 'application/json');
    toast('JSON 备份已导出');
  }

  function importJson(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        const incoming = Array.isArray(data) ? data : (data.records || []);
        if (!incoming.length) { toast('文件里没有记录', true); return; }
        records = records.concat(incoming);
        if (Array.isArray(data.customMotors)) {
          data.customMotors.forEach(m => {
            if (!allMotors().some(x => x.brand === m.brand && x.model === m.model)) customMotors.push(m);
          });
          persistMotors();
        }
        if (Array.isArray(data.customEscs)) {
          data.customEscs.forEach(e => {
            if (!allEscs().some(x => x.brand === e.brand && x.model === e.model)) customEscs.push(e);
          });
          persistEscs();
        }
        persistRecords();
        renderAll();
        toast(`已导入 ${incoming.length} 条记录`);
      } catch (err) {
        toast('导入失败：不是有效的 JSON 备份文件', true);
      }
    };
    reader.readAsText(file, 'utf-8');
  }

  /* ============================================================
   *  导出 Excel：纯 JS 生成 .xlsx（zip + xml，无第三方库）
   * ============================================================ */
  let CRC_TABLE = null;
  function crc32(u8) {
    if (!CRC_TABLE) {
      const t = new Uint32Array(256);
      for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        t[n] = c >>> 0;
      }
      CRC_TABLE = t;
    }
    let c = 0xFFFFFFFF;
    for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function zipFiles(files) {
    const enc = new TextEncoder();
    const chunks = [];
    const central = [];
    let offset = 0;
    const d = new Date();
    const dosTime = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF;
    const dosDate = (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;

    files.forEach(f => {
      const nameBytes = enc.encode(f.name);
      const data = f.data;
      const crc = crc32(data);

      const local = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0x0800, true);
      lv.setUint16(8, 0, true);
      lv.setUint16(10, dosTime, true);
      lv.setUint16(12, dosDate, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, data.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      local.set(nameBytes, 30);
      chunks.push(local, data);

      const cen = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(cen.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, dosTime, true);
      cv.setUint16(14, dosDate, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      cen.set(nameBytes, 46);
      central.push(cen);

      offset += local.length + data.length;
    });

    const centralSize = central.reduce((a, b) => a + b.length, 0);
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    const all = chunks.concat(central, [end]);
    const total = all.reduce((a, b) => a + b.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    all.forEach(a => { out.set(a, pos); pos += a.length; });
    return out;
  }

  function xmlEsc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  const STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<fonts count="2">'
    + '<font><sz val="11"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    + '</fonts>'
    + '<fills count="3">'
    + '<fill><patternFill patternType="none"/></fill>'
    + '<fill><patternFill patternType="gray125"/></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/><bgColor indexed="64"/></patternFill></fill>'
    + '</fills>'
    + '<borders count="1"><border/></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="2">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    + '</cellXfs>'
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';

  function sheetXml(rows, widths) {
    let cols = '';
    if (widths && widths.length) {
      cols = '<cols>' + widths.map((w, i) =>
        `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
    }
    let body = '';
    rows.forEach((row, ri) => {
      const r = ri + 1;
      let cells = '';
      row.forEach((val, ci) => {
        if (val === null || val === undefined || val === '') return;
        const ref = colLetter(ci + 1) + r;
        const style = ri === 0 ? ' s="1"' : '';
        if (typeof val === 'number' && isFinite(val)) {
          cells += `<c r="${ref}"${style}><v>${val}</v></c>`;
        } else {
          cells += `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(val)}</t></is></c>`;
        }
      });
      body += `<row r="${r}">${cells}</row>`;
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<sheetViews><sheetView workbookViewId="0">'
      + '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
      + '</sheetView></sheetViews>'
      + cols
      + `<sheetData>${body}</sheetData>`
      + '</worksheet>';
  }

  function buildXlsx(sheets) {
    const enc = new TextEncoder();
    const files = [];
    const overrides = sheets.map((s, i) =>
      `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');

    files.push({
      name: '[Content_Types].xml', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + overrides
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + '</Types>')
    });
    files.push({
      name: '_rels/.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        + '</Relationships>')
    });

    const sheetTags = sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
    files.push({
      name: 'xl/workbook.xml', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + `<sheets>${sheetTags}</sheets></workbook>`)
    });

    const rels = sheets.map((s, i) =>
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
      + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    files.push({
      name: 'xl/_rels/workbook.xml.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels + '</Relationships>')
    });

    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc.encode(sheetXml(s.rows, s.widths)) }));
    files.push({ name: 'xl/styles.xml', data: enc.encode(STYLES_XML) });
    return zipFiles(files);
  }

  const RECORD_WIDTHS = [9, 18, 9, 10, 11, 11, 11, 10, 11, 16, 7, 11, 13, 14, 14, 10, 9, 10, 9, 9, 12, 10, 10, 26];
  const SUMMARY_WIDTHS = [12, 16, 10, 12, 12, 12, 13];

  function recordsToSheets(list) {
    const rows = [COLUMNS.slice()];
    list.forEach(r => rows.push(COLUMNS.map(c => {
      const v = r[c];
      if (v === null || v === undefined || v === '') return '';
      return (typeof v === 'number' && isFinite(v)) ? v : String(v);
    })));

    const map = new Map();
    list.forEach(r => {
      const key = `${r['电机品牌'] || '未填'} ${r['电机型号'] || ''}`.trim();
      if (!map.has(key)) map.set(key, { brand: r['电机品牌'] || '', model: r['电机型号'] || '', n: 0, t: [], c: [], rpm: [] });
      const g = map.get(key);
      g.n++;
      const t = Number(r['拉力(kg)']); if (isFinite(t)) g.t.push(t);
      const c = Number(r['电流(A)']); if (isFinite(c)) g.c.push(c);
      const m = Number(r['转速(rpm)']); if (isFinite(m)) g.rpm.push(m);
    });
    const mean = a => a.length ? Number((a.reduce((x, y) => x + y, 0) / a.length).toFixed(2)) : '';
    const summary = [['电机品牌', '电机型号', '测试次数', '最大拉力kg', '平均拉力kg', '平均电流A', '最高转速rpm']];
    Array.from(map.values()).forEach(g => {
      summary.push([g.brand, g.model, g.n,
        g.t.length ? Number(Math.max.apply(null, g.t).toFixed(2)) : '',
        mean(g.t), mean(g.c),
        g.rpm.length ? Math.max.apply(null, g.rpm) : '']);
    });

    return [
      { name: '测试记录', rows: rows, widths: RECORD_WIDTHS },
      { name: '按电机汇总', rows: summary, widths: SUMMARY_WIDTHS }
    ];
  }

  function exportXlsx() {
    if (!records.length) { toast('还没有记录', true); return; }
    try {
      const bytes = buildXlsx(recordsToSheets(records));
      downloadBytes(`螺旋桨测试数据_${stamp()}.xlsx`, bytes,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      toast('Excel 已导出');
    } catch (e) {
      toast('导出失败：' + e.message, true);
    }
  }

  /* ============================================================
   *  事件绑定 / 初始化
   * ============================================================ */
  function init() {
    fillSelect($('usage'), USAGES);
    fillSelect($('power'), POWERS);
    fillSelect($('material'), MATERIALS);
    refreshSizeOptions();
    refreshMotorOptions();
    refreshEscOptions();
    refreshFilterOptions();
    renderAll();

    $('usage').addEventListener('change', refreshSizeOptions);
    $('sizeSel').addEventListener('change', e => applySize(e.target.value));
    $('motorSel').addEventListener('change', onMotorPick);
    $('escSel').addEventListener('change', onEscPick);

    $('saveBtn').addEventListener('click', saveRecord);
    $('thrust').addEventListener('keydown', e => {
      if (e.key === 'Enter' && $('autoSave').checked) { e.preventDefault(); saveRecord(); }
    });
    document.addEventListener('keydown', e => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); saveRecord(); }
    });

    ['filterKeyword', 'filterUsage', 'filterMaterial', 'filterMotor'].forEach(id =>
      $(id).addEventListener('input', renderRecords));
    $('filterUsage').addEventListener('change', renderRecords);
    $('filterMaterial').addEventListener('change', renderRecords);
    $('filterMotor').addEventListener('change', renderRecords);
    $('clearFilter').addEventListener('click', () => {
      $('filterKeyword').value = '';
      $('filterUsage').value = '全部用途';
      $('filterMaterial').value = '全部材质';
      $('filterMotor').value = '全部电机';
      renderRecords();
    });

    $('recordsBody').addEventListener('click', e => {
      const btn = e.target.closest('button.del');
      if (!btn) return;
      const id = btn.dataset.id;
      if (!confirm(`确定删除记录 ${id} 吗？`)) return;
      records = records.filter(r => r['记录ID'] !== id);
      persistRecords();
      renderAll();
      toast(`已删除 ${id}`);
    });

    $('dangerClear').addEventListener('click', () => {
      if (!records.length) { toast('没有记录可清空', true); return; }
      if (!confirm(`确定清空全部 ${records.length} 条记录吗？此操作不可恢复，建议先导出备份。`)) return;
      records = [];
      persistRecords();
      renderAll();
      toast('已清空全部记录');
    });

    $('exportXlsx').addEventListener('click', exportXlsx);
    $('exportCsv').addEventListener('click', exportCsv);
    $('exportJson').addEventListener('click', exportJson);
    $('importJson').addEventListener('change', e => {
      const f = e.target.files && e.target.files[0];
      if (f) importJson(f);
      e.target.value = '';
    });

    $('addMotor').addEventListener('click', () => {
      const brand = $('nmBrand').value.trim();
      const model = $('nmModel').value.trim();
      if (!brand || !model) { toast('品牌和型号不能为空', true); return; }
      customMotors.push({
        brand: brand, model: model, kv: valueOrNull('nmKv'), type: $('nmType').value.trim(),
        props: $('nmProps').value.trim(), maxA: valueOrNull('nmMaxA'),
        weight: valueOrNull('nmWeight'), note: $('nmNote').value.trim()
      });
      persistMotors();
      ['nmBrand', 'nmModel', 'nmKv', 'nmType', 'nmProps', 'nmMaxA', 'nmWeight', 'nmNote'].forEach(id => { $(id).value = ''; });
      refreshMotorOptions();
      renderLibraries();
      toast('已添加到电机库');
    });

    $('addEsc').addEventListener('click', () => {
      const brand = $('neBrand').value.trim();
      const model = $('neModel').value.trim();
      if (!brand || !model) { toast('品牌和型号不能为空', true); return; }
      customEscs.push({
        brand: brand, model: model, amp: valueOrNull('neAmp'), peak: valueOrNull('nePeak'),
        volt: $('neVolt').value.trim(), bec: $('neBec').value.trim(),
        weight: valueOrNull('neWeight'), note: $('neNote').value.trim()
      });
      persistEscs();
      ['neBrand', 'neModel', 'neAmp', 'nePeak', 'neWeight', 'neNote'].forEach(id => { $(id).value = ''; });
      refreshEscOptions();
      renderLibraries();
      toast('已添加到电调库');
    });

    $('motorLibBody').addEventListener('click', e => {
      const btn = e.target.closest('button.del');
      if (!btn) return;
      if (!confirm('确定从电机库删除这条自定义电机吗？')) return;
      customMotors.splice(Number(btn.dataset.idx), 1);
      persistMotors();
      refreshMotorOptions();
      renderLibraries();
      toast('已删除');
    });
    $('escLibBody').addEventListener('click', e => {
      const btn = e.target.closest('button.del');
      if (!btn) return;
      if (!confirm('确定从电调库删除这条自定义电调吗？')) return;
      customEscs.splice(Number(btn.dataset.idx), 1);
      persistEscs();
      refreshEscOptions();
      renderLibraries();
      toast('已删除');
    });
  }

  /* 供 node 单元测试导出（浏览器里忽略） */
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { buildXlsx, zipFiles, crc32, sheetXml, recordsToSheets, COLUMNS };
  }
  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
  }
})();

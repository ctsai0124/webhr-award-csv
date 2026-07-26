/* 敘獎 CSV 產製工具 — 介面邏輯 */

const state = {
  roster: [],
  corpus: [],
  docs: [],          // { file, text, meta, block, assess, rows[] }
  signers: [],       // 批核軌跡抽出的職稱沿革原始紀錄
  history: null,
  orgCode: '397085000Y',
  excludePrincipal: true,
  withBasis: true,
  encoding: 'big5',
};

const $ = sel => document.querySelector(sel);
const el = (tag, props = {}, ...kids) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === 'class') n.className = v;
    else if (k === 'text') n.textContent = v;
    else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
    else if (v !== null && v !== false) n.setAttribute(k, v);
  }
  for (const c of kids) if (c) n.append(c);
  return n;
};

/* ---------- 檔案投放區 ---------- */
function wireDrop(id, accept, handler) {
  const zone = $(id), input = zone.querySelector('input');
  input.accept = accept;
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });
  input.addEventListener('change', () => { handler([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach(ev =>
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev =>
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove('over'); }));
  zone.addEventListener('drop', e => handler([...e.dataTransfer.files]));
}

function say(id, msg, isErr = false) {
  const box = $(id);
  box.hidden = false;
  box.className = 'loaded' + (isErr ? ' err' : '');
  // 檔名由使用者提供，必須當純文字顯示，避免特製檔名被當成 HTML 執行。
  box.textContent = msg;
}

/* ---------- 1. 人員基本資料 ---------- */
async function loadRoster(files) {
  const f = files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
    state.roster = parseRoster(wb);
    const edu = state.roster.filter(p => p.kind === 'edu').length;
    const sub = state.roster.filter(p => p.isSubstitute).length;
    say('#rosterMsg',
      `已讀取 ${state.roster.length} 位人員 — 教育人員 ${edu} 位、公務人員 ${state.roster.length - edu} 位` +
      (sub ? `（其中代理教師 ${sub} 位）` : ''));
    if (state.docs.length) reanalyzeDocs();
    $('#step2').hidden = false;
    render();
  } catch (err) {
    say('#rosterMsg', '讀取失敗：' + err.message, true);
  }
}

/* ---------- 2. 歷史獎懲明細（選用） ---------- */
async function loadCorpus(files) {
  const f = files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
    state.corpus = parseCorpus(wb);
    say('#corpusMsg', `已讀取 ${state.corpus.length} 筆歷史事由，撰寫事由時會列出相似寫法供參考`);
    render();
  } catch (err) {
    say('#corpusMsg', '讀取失敗：' + err.message, true);
  }
}

/* ---------- 3. 公文 PDF ---------- */
async function loadDocs(files) {
  const keyOf = f => `${f.name}|${f.size}|${f.lastModified}`;
  const existing = new Set(state.docs.map(doc => doc.key).filter(Boolean));
  const pdfs = files.filter(f => /\.pdf$/i.test(f.name) && !existing.has(keyOf(f)));
  if (!pdfs.length) {
    say('#docMsg', '沒有新增公文；相同的 PDF 已經載入。');
    return;
  }

  // 第一階段：讀完所有 PDF 的文字，同時從批核軌跡蒐集職稱沿革。
  // 必須先讀完才能比對 —— 某份公文的「時任教導主任是誰」，
  // 答案可能藏在另一份公文的簽核紀錄裡。
  for (let i = 0; i < pdfs.length; i++) {
    say('#docMsg', `讀取中… ${i + 1} / ${pdfs.length}　${pdfs[i].name}`);
    try {
      const text = await readPdfText(pdfs[i]);
      state.docs.push({ key: keyOf(pdfs[i]), file: pdfs[i].name, text, rows: [] });
      state.signers.push(...extractSigners(text));
    } catch (err) {
      state.docs.push({
        key: keyOf(pdfs[i]), file: pdfs[i].name, text: '', meta: {}, block: '', rows: [],
        assess: { level: 'fail', note: 'PDF 讀取失敗：' + err.message },
      });
    }
  }

  // 第二階段：合成沿革後，把所有公文重新比對一次
  state.history = mergeSigners(state.signers);
  reanalyzeDocs({ preserve: false });

  renderHistory();
  render();

  const n = state.docs.reduce((s, d) => s + d.rows.length, 0);
  say('#docMsg', `已載入 ${state.docs.length} 份公文，解析出 ${n} 筆敘獎資料`);
  $('#step3').hidden = false;
  $('#step4').hidden = false;
}

async function readPdfText(file) {
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  let text = '';
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    text += itemsToText((await page.getTextContent()).items) + '\n';
  }
  return text;
}

function analyzeDoc(doc) {
  const text = doc.text;
  const meta = parseDocMeta(text);
  const block = extractProposal(text);
  const era = detectEra(meta.subject, meta.year);

  let hits = matchAll(block, state.roster, state.history, era);
  if (state.excludePrincipal) hits = hits.filter(h => !h.person.isPrincipal);

  const awards = findAwards(block);
  let paired = pairNameAward(block, hits, awards);
  const fb = findFallbackAward(text);
  if (fb) paired = paired.map(p => (p.award ? p : { ...p, award: fb }));

  const assess = assessDoc(block, hits, awards);
  const reason = draftReason(meta, block, { withBasis: state.withBasis });
  const category = guessCategory(meta.subject + block);

  Object.assign(doc, {
    meta, block, assess, era,
    rows: paired.map(p => ({
      include: p.confidence === 'exact' && !!p.award && assess.level === 'ok',
      person: p.person,
      written: p.written,
      confidence: p.confidence,
      candidates: p.candidates || null,
      tenure: p.tenure || null,
      awardCode: p.award ? p.award.code : '',
      reason,
      category,
    })),
  });
}

/** 設定或名冊變更後重跑全部公文，盡量保留已核對的人工修改。 */
function reanalyzeDocs(opts = {}) {
  const preserve = opts.preserve !== false;
  const preserveReason = opts.preserveReason !== false;
  for (const doc of state.docs) {
    if (!doc.text) continue;
    const old = new Map((doc.rows || []).map(row => [row.person.id, row]));
    analyzeDoc(doc);
    if (!preserve) continue;
    for (const row of doc.rows) {
      const prior = old.get(row.person.id);
      if (!prior) continue;
      row.include = prior.include;
      row.awardCode = prior.awardCode;
      row.category = prior.category;
      if (preserveReason) row.reason = prior.reason;
    }
  }
}

/** 把整理出來的職稱沿革顯示出來，這是判斷「時任」的依據 */
function renderHistory() {
  const box = $('#history');
  const rows = [];
  for (const [title, list] of (state.history || new Map())) {
    const known = list.filter(r => state.roster.some(p => p.name === r.name));
    if (known.length < 2) continue;
    rows.push(el('div', { class: 'hist-row' },
      el('span', { class: 'hist-title', text: title }),
      el('span', { text: known.map(r => `${r.name}（${fmtRoc(r.first)}–${fmtRoc(r.last)}）`).join('　→　') })));
  }
  box.replaceChildren();
  box.hidden = !rows.length;
  if (rows.length) {
    box.append(el('div', { class: 'hist-head', text: '從批核軌跡整理出的職務交接，公文若涉及過去年度會依此排序候選人' }));
    for (const r of rows) box.append(r);
  }
}

/* ---------- 畫面繪製 ---------- */
function render() {
  const host = $('#docs');
  host.replaceChildren();

  for (const doc of state.docs) host.append(renderDoc(doc));
  refreshSummary();
}

function refreshSummary() {
  const rows = allRows();
  const on = rows.filter(r => r.include);
  $('#summary').innerHTML = on.length
    ? `已核章 <b>${on.length}</b> 筆，共 <b>${new Set(on.map(r => r.person.id)).size}</b> 人`
    : '尚未核章任何資料';
  $('#exportBtn').disabled = !on.length;

  const over = on.filter(r => r.reason.length > 100);
  $('#warn').hidden = !over.length;
  if (over.length) {
    $('#warn').textContent =
      `有 ${over.length} 筆事由超過 100 字上限（${over.map(r => r.person.name).join('、')}），` +
      `WebHR 會拒絕匯入，請先縮短。`;
  }
}

function allRows() {
  return state.docs.flatMap(d => d.rows);
}

function renderDoc(doc) {
  const m = doc.meta || {};
  const head = el('div', { class: 'doc-head' },
    el('h3', { text: m.subject || doc.file }),
    el('div', { class: 'doc-meta', text: [
      m.agency,
      m.year ? `${m.year}.${m.month}.${m.day}` : '',
      m.docNo,
    ].filter(Boolean).join('　') || doc.file }),
  );
  if (doc.assess && doc.assess.level !== 'ok') {
    head.append(el('div', {
      class: 'doc-note' + (doc.assess.level === 'fail' ? ' fail' : ''),
      text: doc.assess.note,
    }));
  }

  const body = el('div', { class: 'doc-body' });
  if (!doc.rows.length) {
    body.append(el('div', { class: 'fields', style: 'padding:14px 18px;color:#6B6E68;font-size:13.5px' },
      el('span', { text: '這份公文沒有解析出可用的名單，請改用 WebHR 手動建檔。' })));
  }
  for (const row of doc.rows) body.append(renderRow(row));

  return el('div', { class: 'doc' }, head, body);
}

const CONF_LABEL = {
  exact:     ['ok',    '姓名相符'],
  typo:      ['typo',  '疑似錯字'],
  partial:   ['check', '需確認'],
  role:      ['typo',  '職稱推定'],
  title:     ['typo',  '職稱對應'],
  ambiguous: ['check', '多人符合'],
};

function renderRow(row) {
  const wrap = el('div', { class: 'row' + (row.include ? '' : ' off') });

  // 核章
  const seal = el('button', {
    type: 'button',
    'aria-pressed': String(row.include),
    'aria-label': `核章 ${row.person.name}`,
    title: '核章（確認納入匯出）',
    text: '核',
    onclick: () => { row.include = !row.include; render(); },
  });
  wrap.append(el('div', { class: 'seal' }, seal));

  const f = el('div', { class: 'fields' });

  // 第一行：狀態、姓名、身分證號、職稱、獎懲結果
  const [cls, label] = CONF_LABEL[row.confidence] || CONF_LABEL.exact;
  const line1 = el('div', { class: 'line1' },
    el('span', { class: 'badge ' + cls, text: label }),
    personSelect(row),
    el('span', { class: 'pid', text: row.person.id }),
    el('span', { class: 'title-txt', text: row.person.title }),
    awardSelect(row),
  );
  if (row.confidence !== 'exact') {
    line1.append(el('span', { class: 'written', text: `公文寫「${row.written}」` }));
  }
  if (row.candidates && row.candidates.length > 1) {
    const names = row.candidates.map(c =>
      c.name + (row.tenure && row.tenure[c.id] ? `（${row.tenure[c.id]}）` : ''));
    line1.append(el('span', { class: 'written', text: `符合的有 ${names.join('、')}，請確認` }));
  }
  if (row.confidence === 'title') {
    line1.append(el('span', { class: 'written', text: '依現職職稱對應' }));
  }
  if (row.person.isSubstitute) {
    line1.append(el('span', { class: 'written', text: '代理教師，請確認適用法規' }));
  }
  f.append(line1);

  // 獎懲事由
  const count = el('span', { class: 'count' });
  const ta = el('textarea', {
    'aria-label': `${row.person.name} 的獎懲事由`,
    oninput: e => { row.reason = e.target.value; updateCount(); refreshSummary(); },
  });
  ta.value = row.reason;
  const updateCount = () => {
    const n = row.reason.length;
    count.textContent = `${n}/100`;
    count.className = 'count' + (n > 100 ? ' over' : '');
  };
  updateCount();
  f.append(el('div', { class: 'reason' }, ta, count));

  f.append(renderMore(row));
  wrap.append(f);
  return wrap;
}

function personSelect(row) {
  const s = el('select', {
    'aria-label': '對應人員',
    onchange: e => {
      row.person = state.roster.find(p => p.id === e.target.value);
      row.confidence = 'exact';
      render();
    },
  });
  const cand = row.candidates || [];
  const ordered = cand.length > 1
    ? [...cand, ...state.roster.filter(p => !cand.includes(p))]
    : state.roster;
  for (const p of ordered) {
    const mark = cand.length > 1 && cand.includes(p) ? '◆ ' : '';
    const ten = row.tenure && row.tenure[p.id] ? `　${row.tenure[p.id]}` : '';
    const o = el('option', { value: p.id, text: `${mark}${p.name}　${p.title}${ten}` });
    if (p.id === row.person.id) o.selected = true;
    s.append(o);
  }
  return s;
}

function awardSelect(row) {
  const s = el('select', {
    'aria-label': '獎懲結果',
    onchange: e => { row.awardCode = e.target.value; render(); },
  });
  s.append(el('option', { value: '', text: '— 請選擇獎度 —' }));
  for (const [label, code] of Object.entries(AWARD_CODES)) {
    const o = el('option', { value: code, text: `${label}　${code}` });
    if (code === row.awardCode) o.selected = true;
    s.append(o);
  }
  if (!row.awardCode) s.style.borderColor = 'var(--vermil)';
  return s;
}

function renderMore(row) {
  const kind = row.person.kind;
  const law = LAW[kind];
  const cl = (CLAUSE[kind] && CLAUSE[kind][row.awardCode]) || {};
  const v = k => (cl[k] === null || cl[k] === undefined ? SENTINEL[k] : cl[k]);

  const catSel = el('select', {
    'aria-label': '獎懲類別',
    onchange: e => { row.category = e.target.value; },
  });
  for (const [code, name] of CATEGORY_CODES) {
    const o = el('option', { value: code, text: `${code} ${name}` });
    if (code === row.category) o.selected = true;
    catSel.append(o);
  }

  const nums = ['條', '點', '項', '款', '目'].map(k =>
    el('label', {}, el('span', { text: k }),
      el('input', { type: 'text', class: 'num', value: String(v(k)), readonly: true })));

  const grid = el('div', { class: 'more-grid' },
    el('label', {}, el('span', { text: '獎懲類別' }), catSel),
    el('span', { class: 'law', text: `適用法規　${law.code}　${law.name}` }),
    ...nums,
    el('span', { class: 'law', text: `教示條款　${row.person.teachClause}（${kind === 'edu' ? '教育人員' : '公務人員'}）` }),
  );

  const det = el('details', { class: 'more' }, el('summary', { text: '法規欄位' }), grid);

  // 歷史事由參考
  if (state.corpus.length) {
    const sims = similarReasons(row.reason, state.corpus, 3);
    if (sims.length) {
      const list = el('div', { class: 'more-grid', style: 'flex-direction:column;align-items:stretch;gap:6px' });
      list.append(el('span', { class: 'law', text: '過去相似寫法（點擊套用）' }));
      for (const s of sims) {
        list.append(el('button', {
          type: 'button', class: 'ghost',
          style: 'text-align:left;font-size:12.5px;padding:5px 9px;line-height:1.5',
          text: s.text,
          onclick: () => { row.reason = s.text; render(); },
        }));
      }
      det.append(list);
    }
  }
  return det;
}

/** 用共同字元比例找相似的歷史事由 */
function similarReasons(reason, corpus, n) {
  const key = new Set((reason || '').replace(/依據[\s\S]*$/, ''));
  return corpus
    .map(c => {
      let hit = 0;
      for (const ch of new Set(c.text)) if (key.has(ch)) hit++;
      return { ...c, score: hit / Math.max(6, new Set(c.text).size) };
    })
    .filter(c => c.score > 0.3)
    .sort((a, b) => b.score - a.score)
    .slice(0, n);
}

/* ---------- 匯出 ---------- */
function exportCsv() {
  const selected = allRows().filter(r => r.include);
  const missingAward = selected.filter(r => !r.awardCode);
  if (missingAward.length) {
    alert(`有 ${missingAward.length} 筆已核章資料尚未選擇獎度，請先補齊。`);
    return;
  }
  const over = selected.filter(r => r.reason.length > 100);
  if (over.length) {
    alert(`有 ${over.length} 筆獎懲事由超過 100 字，WebHR 會拒絕匯入，請先縮短。`);
    return;
  }

  const rows = selected.filter(r => r.awardCode)
    .map(r => buildRow(r.person, r.awardCode, r.reason, r.category, state.orgCode));

  if (!rows.length) return;

  const csv = rowsToCsv(rows);
  let bytes;

  if (state.encoding === 'utf8') {
    bytes = new Uint8Array([0xEF, 0xBB, 0xBF, ...new TextEncoder().encode(csv)]);
  } else {
    const res = Big5.encode(csv);
    if (res.unmapped.length && !confirmMissing(res.unmapped)) return;
    bytes = res.bytes;
  }

  const blob = new Blob([bytes], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = el('a', { href: url, download: `CPAA0118B_${stamp()}.csv` });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Big5 缺字：政府系統的老問題，不能靜默丟掉，要讓使用者知道並選擇怎麼辦 */
function confirmMissing(chars) {
  const who = allRows()
    .filter(r => r.include && chars.some(c => r.person.name.includes(c)))
    .map(r => r.person.name);
  const names = who.length ? `\n受影響人員：${[...new Set(who)].join('、')}` : '';

  return confirm(
    `這些字不在 Big5 字集裡：${chars.join('　')}${names}\n\n` +
    `繼續匯出的話，這些字會變成「?」。\n\n` +
    `兩個處理方向：\n` +
    `1. 取消，改選 UTF-8 編碼再匯出\n` +
    `2. 直接匯出 — WebHR 的姓名欄位是從個人基本資料表2現職檔帶出的，` +
    `CSV 只要身分證號正確，姓名通常會被系統覆蓋回正確的字\n\n` +
    `要直接匯出嗎？`
  );
}

function stamp() {
  const d = new Date();
  const roc = d.getFullYear() - 1911;
  const p = n => String(n).padStart(2, '0');
  return `${roc}${p(d.getMonth() + 1)}${p(d.getDate())}`;
}

/* ---------- 啟動 ---------- */
window.addEventListener('DOMContentLoaded', () => {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  wireDrop('#rosterDrop', '.xls,.xlsx', loadRoster);
  wireDrop('#corpusDrop', '.xls,.xlsx', loadCorpus);
  wireDrop('#docDrop', '.pdf', loadDocs);

  $('#orgCode').addEventListener('input', e => { state.orgCode = e.target.value.trim(); });
  $('#excludePrincipal').addEventListener('change', e => {
    state.excludePrincipal = e.target.checked;
    reanalyzeDocs();
    render();
  });
  $('#withBasis').addEventListener('change', e => {
    state.withBasis = e.target.checked;
    reanalyzeDocs({ preserveReason: false });
    render();
  });
  $('#encoding').addEventListener('change', e => { state.encoding = e.target.value; });

  $('#exportBtn').addEventListener('click', exportCsv);
  $('#sealAll').addEventListener('click', () => {
    const rows = allRows();
    const target = !rows.every(r => r.include);
    rows.forEach(r => { r.include = target && !!r.awardCode; });
    render();
  });
  $('#clearBtn').addEventListener('click', () => {
    state.docs = [];
    state.signers = [];
    state.history = null;
    renderHistory();
    $('#docMsg').hidden = true;
    $('#step3').hidden = true;
    $('#step4').hidden = true;
    render();
  });

  // 先把 Big5 對照表建起來，匯出時就不會卡頓
  if (typeof requestIdleCallback === 'function') requestIdleCallback(() => Big5.build());
  else setTimeout(() => Big5.build(), 500);
});

/* 敘獎 CSV 產製工具 — 介面邏輯 */

/* ---------- 到訪計數（選用） ----------
   這是整個工具唯一會對外發出的請求，內容只有一個固定的計數 key，
   不含任何人事資料、檔名或解析結果。人事資料一律留在瀏覽器。
   把 VISIT_KEY 設成空字串，就完全不會發出任何外部請求。
   服務：Abacus（https://abacus.jasoncameron.dev），免註冊、計數器不公開列出。 */
const VISIT_NAMESPACE = 'ctsai0124.github.io';
const VISIT_KEY = 'webhr-award-csv';

async function countVisit() {
  if (!VISIT_KEY) return;
  const box = $('#visits');
  try {
    const res = await fetch(
      `https://abacus.jasoncameron.dev/hit/${VISIT_NAMESPACE}/${VISIT_KEY}`,
      { cache: 'no-store' },
    );
    if (!res.ok) return;
    const data = await res.json();
    if (typeof data.value !== 'number') return;
    box.textContent = `到訪 ${data.value.toLocaleString('zh-TW')} 次`;
    box.hidden = false;
  } catch {
    // 服務掛掉、被防火牆擋、離線使用都可能失敗。
    // 計數本來就不是必要功能，靜靜略過就好。
  }
}

const state = {
  roster: [],
  corpus: [],
  docs: [],          // { file, text, meta, block, assess, rows[] }
  signers: [],       // 批核軌跡抽出的職稱沿革原始紀錄
  history: null,
  orgCode: '397085000Y',
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
  if (!state.roster.length) {
    say('#docMsg', '請先上傳左側的人員基本資料表，再選取敘獎公文。', true);
    return;
  }
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
      const result = await readPdfText(pdfs[i], detail => {
        say('#docMsg', `${i + 1} / ${pdfs.length}　${pdfs[i].name}　${detail}`);
      });
      state.docs.push({
        key: keyOf(pdfs[i]),
        file: pdfs[i].name,
        text: result.text,
        pageTexts: result.pageTexts,
        ocrUsed: result.ocrUsed,
        ocrPages: result.ocrPages,
        ocrConfidence: result.ocrConfidence,
        rows: [],
      });
      state.signers.push(...extractSigners(result.text));
    } catch (err) {
      state.docs.push({
        key: keyOf(pdfs[i]), file: pdfs[i].name, text: '', pageTexts: [],
        meta: {}, block: '', rows: [],
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
  const ocrCount = state.docs.filter(d => d.ocrUsed).length;
  const perDoc = state.docs.map(doc => {
    const people = new Set(doc.rows.map(row => row.person.id)).size;
    return `• ${doc.file}：${people} 人（${doc.rows.length} 筆）`;
  }).join('\n');
  say('#docMsg',
    `已載入 ${state.docs.length} 份公文，解析出 ${n} 筆敘獎資料` +
    (ocrCount ? `；其中 ${ocrCount} 份使用瀏覽器 OCR，請逐筆核對` : '') +
    (perDoc ? `\n${perDoc}` : ''));
  $('#step3').hidden = false;
  $('#step4').hidden = false;
}

function hasUsablePdfText(text) {
  const compact = (text || '').replace(/\s/g, '');
  const han = ((text || '').match(/[\u3400-\u9FFF]/g) || []).length;
  return compact.length >= 80 && (han >= 15 || compact.length >= 180);
}

function normalizeOcrText(text) {
  return (text || '')
    .replace(/([\u3400-\u9FFF])[ \t]+(?=[\u3400-\u9FFF])/g, '$1')
    .replace(/([\u3400-\u9FFF])[ \t]+(?=[：:，,。；;、])/g, '$1')
    .replace(/([：:，,。；;、])[ \t]+(?=[\u3400-\u9FFF])/g, '$1')
    .replace(/[ \t]+/g, ' ');
}

async function readPdfText(file, onProgress = () => {}) {
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    pages.push({
      page,
      text: itemsToText((await page.getTextContent()).items),
    });
  }

  const nativeText = pages.map(p => p.text).join('\n');
  if (hasUsablePdfText(nativeText)) {
    return {
      text: nativeText,
      pageTexts: pages.map(p => p.text),
      ocrUsed: false,
      ocrPages: [],
      ocrConfidence: null,
    };
  }

  if (!window.Tesseract || typeof window.Tesseract.createWorker !== 'function') {
    throw new Error('OCR 元件載入失敗，請確認網路連線後重新整理頁面');
  }

  let currentPage = 0;
  const worker = await window.Tesseract.createWorker(['chi_tra', 'eng'], 1, {
    workerPath: new URL('tesseract-worker.min.js', window.location.href).href,
    corePath: new URL('tesseract-core-lstm.wasm.js', window.location.href).href,
    langPath: new URL('.', window.location.href).href,
    logger: message => {
      if (message.status === 'recognizing text') {
        onProgress(`OCR 第 ${currentPage}/${pages.length} 頁　${Math.round(message.progress * 100)}%`);
      } else if (message.status === 'loading language traineddata') {
        onProgress('首次載入繁體中文 OCR 模型…');
      }
    },
  });

  const ocrTexts = [];
  const confidences = [];
  try {
    for (let i = 0; i < pages.length; i++) {
      currentPage = i + 1;
      onProgress(`準備 OCR 第 ${currentPage}/${pages.length} 頁…`);
      const base = pages[i].page.getViewport({ scale: 1 });
      const scale = Math.min(2, 2800 / Math.max(base.width, base.height));
      const viewport = pages[i].page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const context = canvas.getContext('2d', { alpha: false });
      context.fillStyle = '#FFFFFF';
      context.fillRect(0, 0, canvas.width, canvas.height);
      await pages[i].page.render({ canvasContext: context, viewport }).promise;
      const result = await worker.recognize(canvas);
      ocrTexts.push(normalizeOcrText(result.data.text));
      if (Number.isFinite(result.data.confidence)) confidences.push(result.data.confidence);
      canvas.width = 0;
      canvas.height = 0;
    }
  } finally {
    await worker.terminate();
  }

  return {
    text: ocrTexts.join('\n'),
    pageTexts: ocrTexts,
    ocrUsed: true,
    ocrPages: pages.map((_, i) => i + 1),
    ocrConfidence: confidences.length
      ? Math.round(confidences.reduce((sum, value) => sum + value, 0) / confidences.length)
      : null,
  };
}

function bestPagesForText(pageTexts, target, limit = 2) {
  if (!target || !pageTexts?.length) return [];
  return pageTexts
    .map((text, index) => ({ page: index + 1, score: similarityScore(text, target) }))
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map(item => item.page);
}

function suggestedReviewPages(doc, meta, block, hits, awards) {
  const pages = new Set([
    ...bestPagesForText(doc.pageTexts, block, 1),
    ...bestPagesForText(doc.pageTexts, meta.subject, 1),
  ]);

  if (!hits.length || !awards.length) {
    (doc.pageTexts || []).forEach((text, index) => {
      if (/(?:敘|敍)[獎奬]|嘉[獎奬]|記功|獎勵額度|敘獎名單/.test(text)) {
        pages.add(index + 1);
      }
    });
  }
  return [...pages].sort((a, b) => a - b).slice(0, 6);
}

/**
 * 哪些比對結果可以自動核章。
 * 除了完全相符，還有兩種可確定的情形不必再勞煩人工：
 *   唯一對得上的筆誤（江雨薇 → 江宇薇，名冊上只有一個人差一字）
 *   批核軌跡明寫「承辦」的「本人」
 */
function autoSeal(p) {
  if (p.confidence === 'exact') return true;
  // 弱比對（只寫名又錯一個字）不自動放行，實際只吻合一個字
  if (p.confidence === 'typo' && p.unique !== false && !p.weak) return true;
  // 只寫名沒寫姓，但名冊上唯一對得上
  if (p.confidence === 'partial' && p.unique !== false) return true;
  if (p.confidence === 'self' && p.selfCertain) return true;
  // 職稱完整命中且唯一，等同寫出姓名。
  // 只有批核軌跡顯示這個職務交接過，才需要人工確認是哪一任。
  if (p.confidence === 'titleExact' && !p.titleHandover) return true;
  // 批核意見只抓「批核意見：」之後的內容，簽核者姓名寫在那行之前不會被撈進來，
  // 所以這裡的姓名完全相符加上明寫的獎度，可靠度等同從擬辦讀到的。
  if (p.confidence === 'approval' && p.baseConfidence === 'exact') return true;
  // 姓、職稱、單位三個條件同時吻合，且名冊上只有一人（多人會是 ambiguous）
  if (p.confidence === 'role') return true;
  return false;
}

function analyzeDoc(doc) {
  const text = doc.text;
  const meta = parseDocMeta(text, doc.file);
  const block = extractProposal(text, doc.file);
  const era = detectEra(meta.subject, meta.year);

  const submitter = findSubmitter(text);
  let hits = matchAll(block, state.roster, state.history, era, submitter);

  let awards = findAwards(block);
  let paired = pairNameAward(block, hits, awards);
  const fb = findFallbackAward(text);
  if (fb && !awards.length) paired = paired.map(p => (p.award ? p : { ...p, award: fb }));
  let approvalUsed = null;
  if (!hits.length) {
    approvalUsed = findApprovalAwards(text, state.roster);
    if (approvalUsed) {
      hits = approvalUsed.paired;
      awards = approvalUsed.awards;
      paired = approvalUsed.paired;
    }
  }

  // info 等級：把資料來源講清楚，但不阻擋自動核章。
  // 批核意見的擷取只取「批核意見：」之後的內容，簽核者姓名不會被誤收。
  let assess = approvalUsed
    ? {
        level: 'info',
        blockAll: false,
        note: '擬辦未列受獎人，姓名與獎度取自批核意見。',
      }
    : assessDoc(block, hits, awards, fb, paired);
  if (meta.subjectCount > 1 || meta.proposalCount > 1) {
    const sectionNote =
      `PDF 內含 ${meta.subjectCount} 組主旨、${meta.proposalCount} 組擬辦，` +
      '已依檔名選擇最相近案件，請確認主旨與名單。';
    assess = {
      level: assess.level === 'fail' ? 'fail' : 'warn',
      blockAll: true,   // 案件可能挑錯，整份都要人工看過
      note: assess.note ? `${sectionNote} ${assess.note}` : sectionNote,
    };
  }
  const duplicateGroups = new Map();
  for (const p of paired) {
    const list = duplicateGroups.get(p.person.id) || [];
    list.push(p);
    duplicateGroups.set(p.person.id, list);
  }
  const duplicates = [...duplicateGroups.values()].filter(list => list.length > 1);
  const duplicateIds = new Set(duplicates.map(list => list[0].person.id));
  if (duplicates.length) {
    const details = duplicates.map(list => {
      const scopes = [...new Set(list.map(p =>
        [p.duty, p.assignment].filter(Boolean).join('／')).filter(Boolean))];
      return `${list[0].person.name}${scopes.length ? `（${scopes.join('、')}）` : ''}${list.length}筆`;
    }).join('、');
    const missingScope = duplicates.some(list =>
      new Set(list.map(p =>
        [p.duty, p.assignment].filter(Boolean).join('／')).filter(Boolean)).size < list.length);
    // 同一人多筆通常代表不同工作或事由，本身不是異常。
    // 只有範圍不足、會讓產生的事由相同時，才要求人工區分。
    if (missingScope) {
      const duplicateNote =
        `同一人因不同工作有多筆敘獎，已全部保留：${details}。` +
        '部分工作範圍未辨識，產生的事由可能相同，請修改事由後再核章。';
      assess = {
        level: assess.level === 'fail' ? 'fail' : 'warn',
        note: assess.note ? `${assess.note} ${duplicateNote}` : duplicateNote,
      };
    }
  }
  if (doc.ocrUsed) {
    const confidence = doc.ocrConfidence === null ? '' : `，平均信心值 ${doc.ocrConfidence}%`;
    const ocrNote =
      `這份掃描公文使用瀏覽器 OCR（第 ${doc.ocrPages.join('、')} 頁${confidence}）。` +
      '姓名、獎度與事由均須人工核對，不會預設核章。';
    assess = {
      level: assess.level === 'fail' ? 'fail' : 'warn',
      blockAll: true,   // OCR 辨識結果不保證正確
      note: assess.note ? `${ocrNote} ${assess.note}` : ocrNote,
    };
  }
  if (assess.level !== 'ok' || !paired.length) {
    const reviewPages = suggestedReviewPages(doc, meta, block, hits, awards);
    if (approvalUsed) {
      for (const page of bestPagesForText(doc.pageTexts, approvalUsed.block, 2)) {
        if (!reviewPages.includes(page)) reviewPages.push(page);
      }
      reviewPages.sort((a, b) => a - b);
    }
    if (reviewPages.length) {
      const pageNote = `建議查看第 ${reviewPages.join('、')} 頁的擬辦、敘獎名單或獎度說明。`;
      assess.note = assess.note ? `${assess.note} ${pageNote}` : pageNote;
    }
    doc.reviewPages = reviewPages;
  } else {
    doc.reviewPages = [];
  }
  const reason = draftReason(meta, block, { withBasis: state.withBasis });
  const category = guessCategory(meta.subject + block);

  const built = paired.map(p => ({
    include: false,
    person: p.person,
    written: p.written,
    confidence: p.confidence,
    titleTier: p.titleTier || '',
    titleHandover: !!p.titleHandover,
    eraPast: !!(era && era.past),
    unique: p.unique !== false,
    weak: !!p.weak,
    selfCertain: !!p.selfCertain,
    selfTitle: p.selfTitle || '',
    candidates: p.candidates || null,
    tenure: p.tenure || null,
    duty: p.duty || '',
    assignment: p.assignment || '',
    sourcePos: p.pos,
    occurrenceKey: `${p.person.id}@${p.pos}`,
    awardCode: p.award ? p.award.code : '',
    reason: applyOccurrenceToReason(
      reason,
      p.duty,
      duplicateIds.has(p.person.id) ? p.assignment : '',
    ),
    category,
    _auto: autoSeal(p) && !!p.award,
  }));

  // 同一人多筆而事由撞在一起時，WebHR 會拒絕匯入。
  // 先自動從擬辦的句段抓出區別寫進事由，真的抓不出來才回頭麻煩人工。
  const stillSame = differentiateReasons(block, built);
  if (stillSame.length) {
    const note = `${stillSame.join('、')} 有多筆敘獎，但擬辦看不出工作範圍的差異，` +
      '請自行補上區別後再核章。';
    assess = {
      level: assess.level === 'fail' ? 'fail' : 'warn',
      blockAll: !!assess.blockAll,   // 只擋事由撞在一起的那幾位
      note: assess.note ? `${assess.note} ${note}` : note,
    };
  }
  for (const row of built) {
    row.include = row._auto && !assess.blockAll && !stillSame.includes(row.person.name);
    delete row._auto;
  }

  Object.assign(doc, {
    meta, block, assess, era, submitter,
    baseReason: reason,
    baseCategory: category,
    rows: built,
  });
}

/** 設定或名冊變更後重跑全部公文，盡量保留已核對的人工修改。 */
function reanalyzeDocs(opts = {}) {
  const preserve = opts.preserve !== false;
  const preserveReason = opts.preserveReason !== false;
  for (const doc of state.docs) {
    if (!doc.text) continue;
    const old = new Map((doc.rows || []).map(row => [
      row.occurrenceKey || `${row.person.id}@${row.sourcePos ?? ''}`,
      row,
    ]));
    // 手動新增的列不是解析出來的，重跑後要原樣接回去
    const manual = (doc.rows || []).filter(row => row.manual);
    analyzeDoc(doc);
    if (manual.length) doc.rows.push(...manual);
    if (!preserve) continue;
    for (const row of doc.rows) {
      const prior = old.get(row.occurrenceKey);
      if (!prior) continue;
      row.include = prior.include;
      if (prior.picked) { row.person = prior.person; row.picked = true; }
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
  renderTally(on);
  refreshFlow(on.length);
}

/** 本次核章的獎度分布，匯入 WebHR 前可以先對一次額度 */
function renderTally(rows) {
  const box = $('#tally');
  box.hidden = !rows.length;
  box.replaceChildren();
  if (!rows.length) return;

  const count = new Map();
  for (const r of rows) count.set(r.awardCode, (count.get(r.awardCode) || 0) + 1);

  for (const [label, code] of Object.entries(AWARD_CODES)) {
    const n = count.get(code) || 0;
    if (!n) continue;
    box.append(el('div', { class: 'tally-item' },
      el('span', { class: 'tally-label', text: label }),
      el('span', { class: 'tally-num', text: String(n) }),
      el('span', { class: 'tally-label', text: '筆' })));
  }
  const docs = state.docs.filter(d => d.rows.some(r => r.include)).length;
  box.append(el('div', { class: 'tally-item' },
    el('span', { class: 'tally-label', text: '來自公文' }),
    el('span', { class: 'tally-num', text: String(docs) }),
    el('span', { class: 'tally-label', text: '份' })));
}

/* ---------- 累計統計 ----------
   只存在這台電腦的瀏覽器裡，不會外傳。記錄總筆數與公文份數，
   讓你知道這個工具實際幫忙處理掉多少件。 */
const LIFETIME_KEY = 'webhr-award-lifetime';

function readLifetime() {
  try {
    const raw = localStorage.getItem(LIFETIME_KEY);
    const v = raw ? JSON.parse(raw) : null;
    return v && typeof v === 'object' ? { rows: v.rows | 0, docs: v.docs | 0, runs: v.runs | 0 } : null;
  } catch { return null; }
}

function bumpLifetime(rows, docs) {
  const cur = readLifetime() || { rows: 0, docs: 0, runs: 0 };
  const next = { rows: cur.rows + rows, docs: cur.docs + docs, runs: cur.runs + 1 };
  try { localStorage.setItem(LIFETIME_KEY, JSON.stringify(next)); } catch { /* 無痕模式會擋，忽略 */ }
  renderLifetime();
}

function renderLifetime() {
  const v = readLifetime();
  const box = $('#lifetime');
  const btn = $('#resetLifetime');
  if (!v || !v.rows) { box.textContent = ''; btn.hidden = true; return; }
  box.textContent = `累計匯出 ${v.rows} 筆・${v.docs} 份公文・${v.runs} 次`;
  btn.hidden = false;
}

function allRows() {
  return state.docs.flatMap(d => d.rows);
}

function refreshFlow(selectedCount = 0) {
  const active = selectedCount ? 4 : state.docs.length ? 3 : state.roster.length ? 2 : 1;
  document.querySelectorAll('[data-flow-step]').forEach(item => {
    const step = Number(item.dataset.flowStep);
    item.classList.toggle('active', step === active);
    item.classList.toggle('done', step < active);
  });
}

/**
 * 把事由拆成「本文」和「依據…函辦理」。
 * 依據子句由公文表頭機械產生，同一份公文每個人都一樣，
 * 畫面上講一次就好，但匯出時仍要接回去。
 */
function splitReason(reason) {
  const i = (reason || '').indexOf('\n依據');
  if (i < 0) return { main: reason || '', basis: '' };
  return { main: reason.slice(0, i), basis: reason.slice(i + 1) };
}

function docBasis(doc) {
  for (const row of doc.rows || []) {
    const { basis } = splitReason(row.reason);
    if (basis) return basis;
  }
  return '';
}

/** 這份公文核了幾個人，蓋章時要能立刻看到數字變化 */
function docCount(doc) {
  const total = doc.rows.length;
  if (!total) return el('span', { class: 'doc-count none', text: '無名單' });

  const on = doc.rows.filter(r => r.include);
  const people = new Set(on.map(r => r.person.id)).size;
  // 同一人可能因不同分工有多筆，人數和筆數不一致時兩個都要講清楚
  const detail = people && people !== on.length ? `（${people} 人）` : '';

  const state = on.length === 0 ? 'none' : (on.length === total ? 'all' : 'part');
  return el('span', {
    class: 'doc-count ' + state,
    title: `共 ${total} 筆，已核章 ${on.length} 筆`,
    text: `已核章 ${on.length} / ${total}${detail}`,
  });
}

function renderDoc(doc) {
  const m = doc.meta || {};
  const meta = [
    m.agency,
    m.year ? `${m.year}.${m.month}.${m.day}` : '',
    m.docNo,
  ].filter(Boolean).join('　');

  const head = el('div', { class: 'doc-head' },
    el('div', { class: 'doc-head-main' },
      el('div', { class: 'doc-titles' },
        // 檔名開頭是案件編號，捲動時那個才是一眼認出公文的識別，放在主旨之上
        el('div', { class: 'doc-file', title: doc.file },
          el('i', { class: 'file-mark', 'aria-hidden': 'true' }),
          el('span', { text: doc.file })),
        el('h3', { text: m.subject || doc.file }),
        meta ? el('div', { class: 'doc-meta', text: meta }) : null,
      ),
      docCount(doc),
    ),
  );

  // 依據子句每個人都一樣，在這裡講一次，下面各列就不重複了
  const basis = docBasis(doc);
  if (basis) {
    head.append(el('div', { class: 'doc-basis' },
      el('span', { class: 'doc-basis-tag', text: '各筆事由後自動接' }),
      el('span', { text: `${basis}（${basis.length} 字）` })));
  }
  if (doc.assess && doc.assess.level !== 'ok') {
    head.append(el('div', {
      class: 'doc-note' + (doc.assess.level === 'fail' ? ' fail'
        : doc.assess.level === 'info' ? ' info' : ''),
      text: doc.assess.note,
    }));
  }

  const body = el('div', { class: 'doc-body' });
  if (!doc.rows.length) {
    body.append(el('div', { class: 'fields', style: 'padding:14px 18px;color:#6B6E68;font-size:13.5px' },
      el('span', { text: '這份公文沒有解析出名單。可以用下面的「新增一筆」自行指定人員。' })));
  }
  for (const row of doc.rows) body.append(renderRow(row));

  // 解析不出來或漏抓時，讓使用者自己補一筆
  body.append(el('div', { class: 'add-row' },
    el('button', {
      type: 'button', class: 'ghost add-btn', text: '＋ 新增一筆',
      onclick: () => { addRow(doc); },
    }),
    el('span', { class: 'add-hint', text: '從名冊挑人，適用於系統沒抓到或抓錯人的情形' }),
  ));

  return el('div', { class: 'doc' }, head, body);
}

const CONF_LABEL = {
  exact:     ['ok',    '姓名相符'],
  typo:      ['typo',  '疑似錯字'],
  typoSure:  ['ok',    '錯字已更正'],
  manual:    ['check', '手動新增'],
  self:      ['ok',    '本人＝承辦人'],
  selfGuess: ['check', '本人待確認'],
  partial:   ['check', '需確認'],
  partialSure: ['ok',  '省略姓氏'],
  approval:    ['ok',  '批核意見'],
  role:      ['ok',    '姓＋職稱'],
  title:     ['typo',  '職稱對應'],
  titleExact:['ok',    '職稱相符'],
  titlePast: ['typo',  '職稱相符'],
  history:   ['check', '時任推定'],
  ambiguous: ['check', '多人符合'],
};

/** 手動補一筆敘獎資料。預設不核章，人員與獎度都要自己挑。 */
function addRow(doc) {
  if (!state.roster.length) return;
  const person = state.roster[0];
  doc.rows.push({
    include: false,
    person,
    written: '',
    confidence: 'manual',
    unique: true,
    selfCertain: false,
    selfTitle: '',
    candidates: null,
    tenure: null,
    duty: '',
    assignment: '',
    sourcePos: Number.MAX_SAFE_INTEGER,
    occurrenceKey: `manual@${Date.now()}`,
    awardCode: '',
    reason: doc.baseReason || '',
    category: doc.baseCategory || 'A02',
    manual: true,
  });
  render();
}

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
  let key = row.confidence;
  if (key === 'typo' && row.unique !== false && !row.weak) key = 'typoSure';
  if (key === 'partial' && row.unique !== false) key = 'partialSure';
  if (key === 'self' && !row.selfCertain) key = 'selfGuess';
  if (key === 'titleExact' && row.titleHandover) key = 'titlePast';
  if (key === 'title' && row.titleTier === 'history') key = 'history';
  const [cls, label] = CONF_LABEL[key] || CONF_LABEL.exact;
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
    // 既然已經算出候選人，就讓使用者直接點，不必再去下拉裡找
    const pick = el('span', { class: 'picks' },
      el('span', { class: 'picks-label', text: '請選一位' }));
    for (const c of row.candidates) {
      const tenure = row.tenure && row.tenure[c.id] ? row.tenure[c.id] : '現職';
      pick.append(el('button', {
        type: 'button',
        class: 'pick' + (c.id === row.person.id ? ' on' : ''),
        'aria-pressed': String(c.id === row.person.id),
        onclick: () => {
          row.person = c;
          row.picked = true;
          render();
        },
      },
        el('span', { class: 'pick-name', text: c.name }),
        el('span', { class: 'pick-sub', text: `${c.title}　${tenure}` }),
      ));
    }
    line1.append(pick);
  }
  if (row.manual) {
    line1.append(el('button', {
      type: 'button', class: 'link-btn',
      text: '移除這筆',
      onclick: () => {
        const doc = state.docs.find(d => d.rows.includes(row));
        if (doc) { doc.rows = doc.rows.filter(r => r !== row); render(); }
      },
    }));
  }
  if (row.confidence === 'self') {
    line1.append(el('span', { class: 'written',
      text: row.selfCertain
        ? `公文寫「${row.written}」，批核軌跡的承辦人是${row.selfTitle} ${row.person.name}`
        : `公文寫「${row.written}」，軌跡沒明寫承辦人，這是依第 1 筆簽核推的` }));
  }
  if (row.confidence === 'typo' && row.unique !== false && !row.weak) {
    line1.append(el('span', { class: 'written', text: `公文誤寫「${row.written}」，已更正` }));
  }
  if (row.confidence === 'typo' && row.weak) {
    line1.append(el('span', { class: 'written',
      text: `公文只寫「${row.written}」且與名冊差一字，可能是名冊外的人，請確認` }));
  }
  if (row.confidence === 'approval') {
    line1.append(el('span', { class: 'written', text: '擬辦沒列名單，取自批核意見' }));
  }
  if (row.confidence === 'partial' && row.unique !== false) {
    line1.append(el('span', { class: 'written',
      text: `公文只寫「${row.written}」，名冊上僅此一人` }));
  }
  if (row.confidence === 'titleExact') {
    line1.append(el('span', { class: 'written',
      text: row.titleHandover ? '名冊職稱完全相符，但這個職務交接過，請確認是哪一任'
                              : '名冊職稱完全相符' }));
  }
  if (row.confidence === 'title') {
    line1.append(el('span', { class: 'written',
      text: row.titleTier === 'history' ? '現職職稱對不上，是依批核軌跡的沿革推定'
          : row.titleTier === 'unit' ? '職稱沒完全對上，是靠單位分辨的'
          : '依現職職稱對應' }));
  }
  if (row.person.isSubstitute) {
    line1.append(el('span', { class: 'written', text: '代理教師，請確認適用法規' }));
  }
  f.append(line1);

  // 獎懲事由
  const count = el('span', { class: 'count' });
  const { basis } = splitReason(row.reason);
  const ta = el('textarea', {
    'aria-label': `${row.person.name} 的獎懲事由`,
    rows: '1',
    oninput: e => {
      // 畫面上只編輯本文，依據子句在存回去時接上
      row.reason = basis ? `${e.target.value}\n${basis}` : e.target.value;
      updateCount();
      refreshSummary();
    },
  });
  ta.value = splitReason(row.reason).main;
  const updateCount = () => {
    // 上限算的是含依據的完整字數，那才是 WebHR 檢查的長度
    const n = row.reason.length;
    count.textContent = `${n}/100`;
    count.title = basis ? `本文 ${n - basis.length - 1} 字 ＋ 依據 ${basis.length} 字` : '';
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
  // 畫面以空白表示不適用；CSV 匯出時仍由 buildRow 轉為 WebHR 哨兵值。
  const v = k => (cl[k] === null || cl[k] === undefined ? '' : cl[k]);

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
    el('span', {
      class: 'law',
      text: kind === 'civil'
        ? '適用法規　公務人員預設留空，請於 WebHR 確認'
        : `適用法規　${law.code}　${law.name}`,
    }),
    ...nums,
    ...(kind === 'edu' && (row.awardCode === '4001' || row.awardCode === '4002')
      ? [el('span', { class: 'law', text: '教師嘉獎預設：第6條第2項第3款第10目（無法細分時採第10目；「點」不適用）' })]
      : []),
    ...(kind === 'edu' && row.awardCode === '4010'
      ? [el('span', { class: 'law', text: '教師記功預設：第6條第2項第2款第9目（無法細分時採第9目；「點」不適用）' })]
      : []),
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
  bumpLifetime(rows.length, state.docs.filter(d => d.rows.some(r => r.include)).length);
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
    new URL('pdf.worker.min.js', window.location.href).href;

  $('#version').textContent = `v${APP_VERSION}　${APP_DATE}`;
  countVisit();
  renderLifetime();
  $('#resetLifetime').addEventListener('click', () => {
    if (!confirm('要把累計數字歸零嗎？這不會影響已經匯出的檔案。')) return;
    try { localStorage.removeItem(LIFETIME_KEY); } catch { /* 忽略 */ }
    renderLifetime();
  });

  wireDrop('#rosterDrop', '.xls,.xlsx', loadRoster);
  wireDrop('#corpusDrop', '.xls,.xlsx', loadCorpus);
  wireDrop('#docDrop', '.pdf', loadDocs);

  $('#orgCode').addEventListener('input', e => { state.orgCode = e.target.value.trim(); });
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

/* ===========================================================
   敘獎 CSV 產製工具
   =========================================================== */

const APP_VERSION = '1.16.0';
const APP_DATE = '2026-07-27';


/* -----------------------------------------------------------
   異體字正規化
   人事資料和公文常混用同一個字的不同寫法，肉眼幾乎分不出來，
   對程式卻是兩個字元。不處理的話「陳佳姸」和「佳妍」會被判成錯字，
   而畫面上兩個名字看起來完全一樣，使用者根本不知道差在哪。
   ----------------------------------------------------------- */
const VARIANT_PAIRS = [
  '妍姸', '峰峯', '群羣', '裡裏', '為爲', '彥彦', '曉暁', '台臺',
  '珏玨', '秘祕', '溫温', '青靑', '真眞', '姊姉', '冊册', '兌兑',
  '匯滙', '敘敍', '床牀', '啟啓', '鵝鵞', '線綫', '污汙', '佩珮',
  '蒓蓴', '踪蹤', '嘆歎', '毀燬', '裸躶', '恆恒', '游遊', '喆哲',
];

const VARIANT_MAP = (() => {
  const m = new Map();
  for (const pair of VARIANT_PAIRS) {
    for (const ch of pair) m.set(ch, pair[0]);
  }
  return m;
})();

/** 把異體字統一成同一種寫法再比對 */
function normVariant(s) {
  let out = '';
  for (const ch of String(s || '')) out += VARIANT_MAP.get(ch) || ch;
  return out;
}


/* -----------------------------------------------------------
   1. Big5 編碼器
   瀏覽器原生只有 TextDecoder 支援 big5，沒有 TextEncoder。
   解法：用 TextDecoder 反查，在執行時建出「字元 -> 位元組」對照表。
   ----------------------------------------------------------- */
const Big5 = (() => {
  let table = null;

  // WHATWG 規範：這幾個字要取「最後」一個 pointer，不是第一個
  const PREFER_LAST = new Set(['\u2550', '\u255E', '\u2561', '\u256A', '\u5341', '\u5345']);

  function build() {
    if (table) return table;
    const map = new Map();
    const dec = new TextDecoder('big5');
    const buf = new Uint8Array(2);
    for (let lead = 0x81; lead <= 0xFE; lead++) {
      for (let trail = 0x40; trail <= 0xFE; trail++) {
        if (trail > 0x7E && trail < 0xA1) continue;
        buf[0] = lead; buf[1] = trail;
        const ch = dec.decode(buf);
        if (ch.length !== 1 || ch === '\uFFFD') continue;
        // 跳過私用區（造字）。這些碼位在每台電腦代表的字不一樣，
        // 寫進 CSV 匯入 WebHR 只會變成亂碼，要當成缺字提醒使用者。
        const cp = ch.codePointAt(0);
        if (cp >= 0xE000 && cp <= 0xF8FF) continue;
        if (PREFER_LAST.has(ch)) { map.set(ch, [lead, trail]); continue; }
        if (!map.has(ch)) map.set(ch, [lead, trail]);
      }
    }
    table = map;
    return table;
  }

  /** 把字串編成 Big5 位元組。無法對應的字元回報出來，不靜默丟棄。 */
  function encode(str) {
    const t = build();
    const out = [];
    const unmapped = new Set();
    for (const ch of str) {
      const code = ch.codePointAt(0);
      if (code < 0x80) { out.push(code); continue; }
      const pair = t.get(ch);
      if (pair) { out.push(pair[0], pair[1]); }
      else { unmapped.add(ch); out.push(0x3F); } // '?'
    }
    return { bytes: new Uint8Array(out), unmapped: [...unmapped] };
  }

  return { encode, build };
})();


/* -----------------------------------------------------------
   2. 代碼常數
   ----------------------------------------------------------- */
const AWARD_CODES = {
  '嘉獎一次': '4001',
  '嘉獎二次': '4002',
  '記功一次': '4010',
};

const CATEGORY_CODES = [
  ['A01', '領導有方'], ['A02', '工作績優'], ['A03', '操守廉潔'],
  ['A04', '全勤獎勵'], ['A05', '研究發展績優'], ['A06', '訓練進修績優'],
  ['A07', '競技比賽績優'], ['A08', '特殊功績'], ['A09', '優良事蹟'],
  ['A10', '服務績優'],
];

const LAW = {
  edu: { code: 'CA', name: '公立高級中等以下學校教師成績考核辦法' },
  civil: { code: 'BT', name: '高雄市政府及所屬各機關公務人員平時獎懲標準表' },
};

/**
 * 教師成績考核辦法第6條第2項各目（115.1.12 修正版）。
 * 嘉獎與記功各自的目，供使用者依實際情形挑選。
 */
const EDU_ITEMS = {
  // 第3款 嘉獎
  '4001': [
    [1, '課業編排得當，課程調配妥善，經實施確具成效'],
    [2, '進行課程研發，有具體績效，在校內進行分享'],
    [3, '編撰教材、自製教具或教學媒體，成績優良'],
    [4, '教學優良，評量認真，確能提高學生程度'],
    [5, '對學生之輔導管教，熱心負責，成績優良'],
    [6, '辦理教學演示、分享或研習活動，表現優異'],
    [7, '教師本人或指導學生參加各項活動、比賽，成績優良'],
    [8, '擔任導師能有效進行品格教育、生活教育足堪表率'],
    [9, '在課程研發、教學創新、多元評量等方面著有績效，促進團隊合作'],
    [10, '其他辦理有關教育工作，成績優良'],
  ],
  // 第2款 記功
  '4010': [
    [1, '革新改進教育業務，且努力推行，著有成效'],
    [2, '對學校校務、設施，有長期發展計畫，且能切實執行，績效卓著'],
    [3, '研究改進教材教法，確能增進教學效果，提高學生程度'],
    [4, '自願輔導學生課業，並能注意學生身心健康，而教學成績優良'],
    [5, '推展輔導管教工作，確能變化學生氣質，造成優良學風'],
    [6, '輔導畢業學生就業，著有成績'],
    [7, '對偶發事件之預防或處理適當，因而避免或減少可能發生之損害'],
    [8, '教師本人或指導學生代表學校參加各級主管機關認定之全國校際比賽，成績卓著'],
    [9, '其他優良事蹟，足資表率'],
  ],
};
EDU_ITEMS['4002'] = EDU_ITEMS['4001'];   // 嘉獎二次與一次同屬第3款

// 條 / 點 / 項 / 款 / 目。null 代表該欄不適用，輸出時留空。
const CLAUSE = {
  edu: {
    '4001': { 條: 6, 點: null, 項: 2, 款: 3, 目: 10 },
    '4002': { 條: 6, 點: null, 項: 2, 款: 3, 目: 10 },
    '4010': { 條: 6, 點: null, 項: 2, 款: 2, 目: 9 },
  },
  // 公務人員不自動推定條點項款目，一律輸出哨兵值交由承辦人於 WebHR 調整。
  // 但「適用法規名稱」必填，留空 WebHR 會拒絕匯入。
  civil: {},
};

// 哨兵值：沿用你既有 CSV 的慣例
// 不適用的欄位一律留空，不填哨兵值。
// WebHR 接受空白，填 999 反而讓人以為真的引用了第 999 條。
const SENTINEL = { 條: '', 點: '', 項: '', 款: '', 目: '' };


/* -----------------------------------------------------------
   3. 人員基本資料表解析 (CPAB1207R)
   ----------------------------------------------------------- */
function parseRoster(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // 找標題列：同時含「姓名」與「身分證號」
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const r = rows[i].map(c => String(c).trim());
    if (r.includes('姓名') && r.includes('身分證號')) { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) throw new Error('找不到標題列。請確認上傳的是 CPAB1207R 人員基本資料表。');

  // 標題列之前通常是學校名稱，抓來當機關代碼的對照依據
  let school = '';
  for (let i = 0; i < hdrIdx; i++) {
    const line = rows[i].map(c => String(c).trim()).filter(Boolean).join('');
    const m = line.match(/([\u4e00-\u9fa5]{2,20}?(?:國民小學|國民中學|國小|國中|高級中學|高中|高工|高商|幼兒園|實驗學校|學校))/);
    if (m) { school = m[1]; break; }
  }

  const hdr = rows[hdrIdx].map(c => String(c).trim());
  const col = name => hdr.indexOf(name);
  const cName = col('姓名'), cId = col('身分證號'), cTitle = col('職稱');
  const cUnit = col('單位'), cRank = col('現支官職等');

  const people = [];
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const name = String(r[cName] ?? '').trim();
    const id = String(r[cId] ?? '').trim().toUpperCase();
    if (!name || !/^[A-Z][0-9]{9}$/.test(id)) continue;

    const title = String(r[cTitle] ?? '').trim();
    const rank = String(r[cRank] ?? '').trim();
    const unit = String(r[cUnit] ?? '').trim();

    people.push({
      name, id, title, unit,
      ...classifyPerson(title, rank),
    });
  }
  if (!people.length) throw new Error('沒有讀到任何人員資料，請確認檔案內容。');
  people.school = school;
  return people;
}

/**
 * 判斷人員類別。
 * 教育人員：職稱含「教師」，或現支官職等是薪額制（教師薪級）。
 * 其餘（護理師、幹事、書記等職員）一律歸公務人員。
 */
function classifyPerson(title, rank) {
  const isEdu = /教師/.test(title) || /薪額|薪點/.test(rank);
  const isSubstitute = /代理|代課|兼任/.test(title);
  const isPrincipal = /校長/.test(title);
  return {
    kind: isEdu ? 'edu' : 'civil',
    teachClause: isEdu ? '2' : '1',   // 教示條款：2 教育人員 / 1 公務人員
    isSubstitute,
    isPrincipal,
  };
}


/* -----------------------------------------------------------
   4. 歷史獎懲明細表解析 (CPAB1209R) — 事由語料庫，選用
   ----------------------------------------------------------- */
function parseCorpus(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const r = rows[i].map(c => String(c).trim());
    if (r.includes('獎懲事由')) { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) return [];
  const hdr = rows[hdrIdx].map(c => String(c).trim());
  const cReason = hdr.indexOf('獎懲事由');
  const cResult = hdr.indexOf('獎懲結果');

  const out = [];
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const raw = String(rows[i][cReason] ?? '').trim();
    if (!raw) continue;
    // Excel 的軟換行會被存成 _x000d_
    const main = raw.split(/_x000d_|\r|\n/)[0].trim();
    if (main.length < 6) continue;
    out.push({ text: main, result: String(rows[i][cResult] ?? '').trim() });
  }
  return out;
}


/* -----------------------------------------------------------
   5. 公文文字擷取與欄位解析
   ----------------------------------------------------------- */

/** 從 pdf.js 的 textContent 還原成帶換行的純文字 */
/**
 * 把 Mac 端 Vision OCR 回傳的 token 還原成文字。
 *
 * token 帶 x/y 座標，所以能照實際版面排序 —— 這正是 pdf.js 抽表格時
 * 會弄丟的資訊。同一列的字由左至右接起來，列與列之間換行。
 *
 * y 是正規化過的（0 在頂端 1 在底端），同列容差取 0.008，
 * 約等於一行文字高度的一半。
 */
function tokensToText(tokens, opts = {}) {
  const rowTol = opts.rowTol ?? 0.008;
  const byPage = new Map();
  for (const t of tokens || []) {
    if (!t || typeof t.text !== 'string') continue;
    const p = t.page ?? 0;
    if (!byPage.has(p)) byPage.set(p, []);
    byPage.get(p).push(t);
  }

  const pages = [];
  for (const page of [...byPage.keys()].sort((a, b) => a - b)) {
    const items = byPage.get(page).slice().sort((a, b) => a.y - b.y || a.x - b.x);
    const rows = [];
    for (const it of items) {
      const row = rows[rows.length - 1];
      if (row && Math.abs(it.y - row.y) <= rowTol) {
        row.items.push(it);
        // 用平均值當基準，長列才不會因為第一個字的位置偏掉而愈飄愈遠
        row.y = (row.y * (row.items.length - 1) + it.y) / row.items.length;
      } else {
        rows.push({ y: it.y, items: [it] });
      }
    }
    pages.push(rows.map(r =>
      r.items.sort((a, b) => a.x - b.x).map(i => i.text).join('')).join('\n'));
  }
  return pages.join('\n');
}

/** 只在同一列的 token 之間插入分隔符，供表格解析辨識欄位邊界 */
function tokensToRows(tokens) {
  const text = tokensToText(tokens);
  return text ? text.split('\n') : [];
}


function itemsToText(items) {
  let out = '', lastY = null;
  for (const it of items) {
    const y = it.transform ? Math.round(it.transform[5]) : null;
    if (lastY !== null && y !== null && Math.abs(y - lastY) > 3) out += '\n';
    out += it.str;
    lastY = y;
  }
  return out;
}

function compareKey(s) {
  return String(s || '')
    .replace(/\.pdf$/i, '')
    .replace(/^[\d-]+/, '')
    .replace(/[，。；：、（）()「」『』【】\sˇ\-—─_]/g, '')
    .replace(/(有關|檢送|本市|貴校|本校|請查照|請依說明辦理|詳如說明)/g, '');
}

function similarityScore(a, b) {
  const x = compareKey(a), y = compareKey(b);
  if (!x || !y) return 0;
  const grams = s => {
    const out = new Set();
    if (s.length === 1) out.add(s);
    for (let i = 0; i < s.length - 1; i++) out.add(s.slice(i, i + 2));
    return out;
  };
  const gx = grams(x), gy = grams(y);
  let common = 0;
  for (const g of gx) if (gy.has(g)) common++;
  return common / Math.max(1, Math.min(gx.size, gy.size));
}

function subjectMatches(text) {
  return [...String(text || '').matchAll(
    /主旨[：:]([\s\S]{0,300}?)(?=說明[：:]|正本[：:]|$)/g,
  )];
}

function selectSubjectMatch(text, filename = '') {
  const matches = subjectMatches(text);
  if (!matches.length) return null;
  if (!filename || matches.length === 1) return matches[0];
  return matches.reduce((best, m) =>
    similarityScore(filename, m[1]) > similarityScore(filename, best[1]) ? m : best);
}

function nearestMatch(matches, pos) {
  if (!matches.length) return null;
  if (!Number.isFinite(pos)) return matches[0];
  return matches.reduce((best, m) =>
    Math.abs(m.index - pos) < Math.abs(best.index - pos) ? m : best);
}

/** 抽公文表頭欄位；PDF 內若有多案，以檔名最接近的主旨為準。 */
/**
 * 判斷這段文字是不是可讀的中文主旨。
 * OCR 失敗時的典型症狀是冒出大量連續的拉丁字母，中文比例掉下來。
 */
function looksReadable(s) {
  const t = String(s || '');
  if (t.length < 8) return false;
  const han = (t.match(/[\u4e00-\u9fa5]/g) || []).length;
  if (han / t.length < 0.6) return false;
  // 連續 6 個以上的大寫拉丁字母，正常公文不會出現
  if (/[A-Za-z]{6,}/.test(t)) return false;
  return true;
}

/** 從檔名還原主旨。檔名開頭常有案件編號，要去掉 */
function subjectFromFilename(filename) {
  if (!filename) return '';
  let s = String(filename).replace(/\.[a-z0-9]+$/i, '');

  // 只砍「編號＋分隔符」的組合（43-47-、04-06-、12_），
  // 不能無差別砍開頭數字，否則「113學年度…」的年度會被吃掉
  s = s.replace(/^\d+(?:[-–~]\d+)*\s*[-–_.、]\s*/, '')
       .replace(/[_]+/g, '')
       .replace(/(?<=[\u4e00-\u9fa5」』）)])\d{1,2}$/, '')   // 結尾流水號，前面要是中文
       .trim();
  return s.length >= 8 && looksReadable(s) ? s : '';
}

function parseDocMeta(text, filename = '') {
  const flat = text.replace(/[ \t]+/g, '');
  const mSub = selectSubjectMatch(text, filename);
  const subjectPos = mSub ? mSub.index : Number.NaN;

  const mDate = nearestMatch(
    [...flat.matchAll(/發文日期[：:]\s*中華民國(\d{2,3})年(\d{1,2})月(\d{1,2})日/g)],
    subjectPos,
  );
  const mNo = nearestMatch(
    [...flat.matchAll(/發文字號[：:]\s*([^\s\n]{6,40}?號)/g)],
    subjectPos,
  );

  // 發文機關：找「XXX 函」那一行
  let agency = '';
  const agencyMatches = [];
  let linePos = 0;
  for (const rawLine of text.split('\n')) {
    const l = rawLine.trim();
    const m = l.match(/^([\u4e00-\u9fa5]{4,20}?)\s*函$/);
    if (m) agencyMatches.push({ 1: m[1], index: linePos });
    linePos += rawLine.length + 1;
  }
  const mAgency = nearestMatch(agencyMatches, subjectPos);
  if (mAgency) agency = mAgency[1];
  if (!agency) {
    const m = nearestMatch(
      [...flat.matchAll(/([\u4e00-\u9fa5]{4,20}?)函[\s\S]{0,80}?地址[：:]/g)],
      subjectPos,
    );
    if (m) agency = m[1];
  }

  // 主旨：到「說明」或「正本」為止
  let subject = '';
  if (mSub) subject = mSub[1].replace(/\s+/g, '').replace(/[。．]+$/, '');

  // OCR 把主旨讀壞時（大量拉丁亂碼、中文比例過低），改用檔名。
  // 檔名是承辦人自己打的，通常完整正確；但公文內容仍優先，只在明顯壞掉時才換。
  let subjectFromFile = false;
  const fromName = subjectFromFilename(filename);
  if (fromName && (!subject || !looksReadable(subject))) {
    subject = fromName;
    subjectFromFile = true;
  }

  // 說明段：主旨講的是「這份公文要幹嘛」，說明才寫「同仁做了什麼」
  let body = '';
  const bodySource = mSub ? text.slice(mSub.index + mSub[0].length) : text;
  const mBody = bodySource.match(/說明[：:]([\s\S]{0,900}?)(?=正本[：:]|副本[：:]|$)/);
  if (mBody) body = cleanBlock(mBody[1]);

  return {
    agency,
    year: mDate ? mDate[1] : '',
    month: mDate ? mDate[2] : '',
    day: mDate ? mDate[3] : '',
    docNo: mNo ? mNo[1] : '',
    subject,
    subjectFromFile,
    body,
    subjectCount: subjectMatches(text).length,
    proposalCount: proposalMatches(text).length,
  };
}

/**
 * 擷取「擬辦」區塊 — 只在這裡找敘獎名單。
 * 這一步很關鍵：批核軌跡裡會出現校長、人事主任等簽核者，
 * 他們不是受獎人，必須排除在解析範圍外。
 */
function proposalMatches(text) {
  const END = /(批核軌跡|批示\s|會辦\s|承辦\s高雄|欄位批核紀錄|貼紙備註資訊|意見[：:])/;
  const out = [];
  const source = String(text || '');
  const proposalMarkers = [...source.matchAll(/擬辦[：:]/g)];
  const markers = proposalMarkers.length
    ? proposalMarkers
    : [...source.matchAll(/承辦意見[：:]/g)];

  for (const m of markers) {
    let block = text.slice(m.index + m[0].length);
    const e = block.search(END);
    if (e >= 0) block = block.slice(0, e);
    out.push({ index: m.index, block: cleanBlock(block).slice(0, 6000) });
  }
  if (!out.length) {
    // 格式二：電子公文「意見及流程 / 文號：xxxxx」之後直接接內容，沒有擬辦標籤
    for (const m of String(text || '').matchAll(
      /意見及流程[\s\S]{0,40}?文號[：:]\s*\d+\s*/g,
    )) {
      let block = text.slice(m.index + m[0].length);
      const e = block.search(END);
      if (e >= 0) block = block.slice(0, e);
      out.push({ index: m.index, block: cleanBlock(block).slice(0, 6000) });
    }
  }
  return out;
}

function extractProposal(text, filename = '') {
  const proposals = proposalMatches(text);
  if (!proposals.length) return '';
  if (proposals.length === 1) return proposals[0].block;

  const subject = selectSubjectMatch(text, filename);
  if (!subject) return proposals[0].block;
  return nearestMatch(proposals, subject.index).block;
}

/** 清掉 PDF 裝訂線的點狀雜訊，這些點會把詞切開（例：承裝﹒﹒﹒辦人員） */
function cleanBlock(s) {
  return s
    .split('\n')
    // 整行只有裝訂線標記的就整行丟掉，否則「嘉獎\n裝\n一次」會被切成無法辨識的字串
    .filter(l => !/^[\s裝訂線﹒·.]*$/.test(l))
    .join('\n')
    .replace(/[裝訂線]?[﹒·]{2,}/g, '')
    .replace(/[﹒]/g, '')
    .replace(/[奬奖]/g, '獎')
    .replace(/[敍叙]/g, '敘')
    .replace(/\s+/g, '');
}

/**
 * 取姓名附近最近的分工標記。
 * 同一人在同一案可能分別以主辦、承辦、協辦身分各敘一次，
 * 這個標記會用來保留多筆資料並把事由寫成不同內容。
 */
function detectDuty(block, pos, written = '') {
  const beforeStart = Math.max(
    block.lastIndexOf('。', pos - 1),
    block.lastIndexOf('；', pos - 1),
    block.lastIndexOf('\n', pos - 1),
    pos - 30,
  ) + 1;
  let before = block.slice(Math.max(0, beforeStart), pos);
  // 獎度通常是上一列的結尾或本列的開頭；不要把更前一列的「承辦」沿用到下一人。
  const lastAward = Math.max(before.lastIndexOf('嘉獎'), before.lastIndexOf('記功'));
  if (lastAward >= 0) before = before.slice(lastAward + 2);
  const dutyRe = /(主要承辦|協助辦理|主辦|承辦|協辦|協助)(?:人員|單位)?/g;
  let duty = '', m;
  while ((m = dutyRe.exec(before)) !== null) duty = m[1];

  if (!duty) {
    const after = block.slice(pos + written.length, pos + written.length + 16);
    const afterMatch = after.match(
      /^(?:老師|教師|主任|組長|護理師|幹事)?(?:(?:[（(：:、，,]+(主要承辦|協助辦理|主辦|承辦|協辦|協助))|(?:(主要承辦|協助辦理|主辦|承辦|協辦|協助)(?:人員|單位)))/,
    );
    if (afterMatch) duty = afterMatch[1] || afterMatch[2];
  }

  if (duty === '主要承辦') return '承辦';
  return duty;
}

/**
 * 取姓名前緊鄰的班級或職務標記。
 * 例如同一位教師分別列在「一忠導師李雅玲」與「四忠導師李雅玲」，
 * 兩筆都要保留，且事由需能分辨其負責範圍。
 */
function detectAssignment(block, pos) {
  const before = block.slice(Math.max(0, pos - 16), pos);
  const qualifier =
    '(?:[一二三四五六七八九十\\d]{1,2}[忠孝仁愛信義和平]?|幼兒園|特教班|資源班|' +
    '教務|學務|訓導|輔導|總務|教導|人事|主計|會計|資訊)?';
  const title = '(?:導師|主任|組長|護理師|幹事|執秘|秘書|觀察員)';
  const m = before.match(new RegExp(`(${qualifier}${title})$`));
  if (m) return m[1];

  // 群組式名單：「國:甲、乙」「授課教師:甲乙」中的每一位都屬於同一範圍。
  // 名單中可能夾有「(校長)」等註記，不能只允許純中文姓名。
  const groupBefore = block.slice(Math.max(0, pos - 90), pos);
  const markerRe = /(授課教師|行政人員|國語文?|數學?|英語?|國|數|英)[：:]/g;
  let marker = null, markerMatch;
  while ((markerMatch = markerRe.exec(groupBefore)) !== null) marker = markerMatch;
  if (!marker) return '';
  const listed = groupBefore.slice(marker.index + marker[0].length);
  if (listed.length > 80 || /[。；\n]/.test(listed)) return '';
  const assignment = ({ 國: '國語', 數: '數學', 英: '英語' })[marker[1]] || marker[1];

  // 同一科目可能分屬不同工作，例如「113年成長測驗」及「114年篩選測驗」。
  // 將最近的測驗名稱併入範圍，才能為同一人的兩筆資料產生不同事由。
  if (/^(?:國語文?|數學|英語)$/.test(assignment)) {
    const scopeBefore = block.slice(Math.max(0, pos - 300), pos);
    const scopes = [...scopeBefore.matchAll(/(11\d年(?:成長測驗|篩選測驗))/g)];
    if (scopes.length) return scopes[scopes.length - 1][1] + assignment;
  }
  return assignment;
}


/* -----------------------------------------------------------
   6. 姓名比對
   以人員名冊為錨點掃描擬辦文字，而不是硬猜姓名格式。
   ----------------------------------------------------------- */
function matchNames(block, roster) {
  const hits = [];
  const seen = new Set();

  const push = (person, pos, written, confidence, unique = true, weak = false) => {
    const key = person.id + '@' + pos;
    if (seen.has(key)) return;
    seen.add(key);
    hits.push({ person, pos, written, confidence, unique, weak });
  };

  // 一個錯寫的姓名，名冊上只有一個人對得起來才算「可確定的筆誤」。
  // 兩個人都差一字就不能自動放行，要人工挑。
  const soleTypoMatch = (written) => roster.filter(p =>
    p.name.length === written.length &&
    p.name[0] === written[0] &&
    [...written].filter((c, k) => c !== p.name[k]).length === 1).length === 1;

  for (let i = 0; i < block.length; i++) {
    for (const p of roster) {
      const n = p.name;
      // 完全相符
      // 異體字（陳佳姸／陳佳妍）視為同一個名字，不能判成錯字
      const seg = block.slice(i, i + n.length);
      if (seg === n || (seg.length === n.length && normVariant(seg) === normVariant(n))) {
        push(p, i, seg, 'exact'); continue;
      }
      // 同長度、差一字（抓錯字：江雨薇/江宇薇、陳玟純/陳玟蒓）
      if (n.length >= 3) {
        const w = block.slice(i, i + n.length);
        if (w.length === n.length) {
          const wn = normVariant(w), nn = normVariant(n);
          let diff = 0;
          for (let k = 0; k < n.length; k++) if (wn[k] !== nn[k]) diff++;
          if (diff === 1 && w[0] === n[0] && /^[\u4e00-\u9fa5]+$/.test(w)) {
            push(p, i, w, 'typo', soleTypoMatch(w));
          }
        }
      }
    }
  }

  // 只有名沒有姓：「怡慈老師」→ 林怡慈
  const TITLE = /(老師|主任|組長|護理師|幹事|執秘|秘書)/;
  for (const p of roster) {
    if (p.name.length < 3) continue;
    const given = p.name.slice(1);
    const givenNorm = normVariant(given);
    const uniqueGiven = roster.filter(
      other => normVariant(other.name.slice(1)) === givenNorm).length === 1;
    // 在正規化後的文字上找位置，異體字（佳妍／佳姸）才對得起來
    const blockNorm = normVariant(block);
    let idx = -1;
    while ((idx = blockNorm.indexOf(givenNorm, idx + 1)) >= 0) {
      if (idx > 0 && block[idx - 1] === p.name[0]) continue; // 已被完整比對抓到
      const after = block.slice(idx + given.length, idx + given.length + 3);
      const before = block.slice(Math.max(0, idx - 12), idx);
      const structured =
        uniqueGiven &&
        (/(?:人員|執秘|主任|組長|名單|[：:、，,])$/.test(before) ||
          /^(?:記功|嘉獎|[、，,])/.test(after));
      // 省略姓氏時，名冊上只有一個人的名字對得起來就等同寫全名；
      // 兩個人同名才需要人工挑。
      if (TITLE.test(after) || structured) {
        push(p, idx, block.slice(idx, idx + given.length), 'partial', uniqueGiven);
      }
    }

    // 只有名且疑似一字錯誤：「佳妍」→ 名冊「佳姸」。
    if (uniqueGiven) {
      for (let i = 0; i <= block.length - given.length; i++) {
        if (i > 0 && block[i - 1] === p.name[0]) continue;
        const written = block.slice(i, i + given.length);
        const writtenNorm = normVariant(written);
        if (!/^[\u4e00-\u9fa5]+$/.test(written) || writtenNorm === givenNorm) continue;
        if (roster.some(other =>
          other !== p && normVariant(other.name.slice(1)) === writtenNorm)) continue;
        if (hits.some(hit =>
          i < hit.pos + hit.written.length && i + written.length > hit.pos)) continue;
        let diff = 0;
        for (let k = 0; k < given.length; k++) if (writtenNorm[k] !== givenNorm[k]) diff++;
        if (diff !== 1 || writtenNorm[0] !== givenNorm[0]) continue;
        const before = block.slice(Math.max(0, i - 12), i);
        const after = block.slice(i + written.length, i + written.length + 4);
        if (/(?:人員|執秘|主任|組長|名單|[：:、，,])$/.test(before) ||
            /^(?:記功|嘉獎|[、，,])/.test(after)) {
          // 沒有姓氏定錨，兩個字還錯一個 —— 實際只吻合一個字，太弱。
          // 名冊上查無此人的離職者（例如寫「怡如」但蔡怡如已離職），
          // 會被誤配成名冊上的「怡慈」，所以標為弱比對，一律人工確認。
          push(p, i, written, 'typo', true, true);
        }
      }
    }
  }

  // 不再依身分證號去重。同一人可能因主辦、承辦、協辦等不同分工敘獎多次。
  return hits.sort((a, b) => a.pos - b.pos);
}


/* -----------------------------------------------------------
   6d. 職稱沿革（取自批核軌跡）
   人員基本資料表只有現職，查不到「113學年度當時的教導主任是誰」。
   但公文後半的批核軌跡會寫「…國民小學教導主任 侯進昇(2025.09.23)」，
   把多份公文的簽核紀錄彙整起來，就看得出職務什麼時候交接。
   ----------------------------------------------------------- */

const SIGNER_RE =
  /(?:小學|中學|國中|國小|學校|幼兒園)([\u4e00-\u9fa5]{2,8}?)[ \u3000]+([\u4e00-\u9fa5]{2,4})[ \u3000]*[（(：:]\s*(\d{3,4})[./](\d{1,2})[./](\d{1,2})/g;

function extractSigners(text) {
  const out = [];
  let m;
  SIGNER_RE.lastIndex = 0;
  while ((m = SIGNER_RE.exec(text)) !== null) {
    const [, title, name, y, mo, d] = m;
    // 西元年換算民國年，兩種寫法的公文都有
    const roc = Number(y) > 1911 ? Number(y) - 1911 : Number(y);
    out.push({ title, name, roc, month: Number(mo), day: Number(d) });
  }
  return out;
}

/** 把多份公文的簽核紀錄合成一份職稱沿革 */
function mergeSigners(all) {
  const map = new Map();
  for (const s of all) {
    const key = s.title;
    if (!map.has(key)) map.set(key, new Map());
    const byName = map.get(key);
    const cur = byName.get(s.name);
    const stamp = s.roc * 10000 + s.month * 100 + s.day;
    if (!cur) byName.set(s.name, { name: s.name, first: stamp, last: stamp });
    else { cur.first = Math.min(cur.first, stamp); cur.last = Math.max(cur.last, stamp); }
  }
  const out = new Map();
  for (const [title, byName] of map) {
    out.set(title, [...byName.values()].sort((a, b) => a.first - b.first));
  }
  return out;
}

const fmtRoc = (n) => `${Math.floor(n / 10000)}/${String(Math.floor(n / 100) % 100).padStart(2, '0')}`;


/* -----------------------------------------------------------
   6e. 「本人」代稱
   擬辦有人會寫「本人擬敘嘉獎1次」，本人指的就是承辦人。
   承辦人在批核軌跡上標得很清楚，可以直接對回名冊。
   ----------------------------------------------------------- */

const SIGNER_TAIL =
  '(?:小學|中學|國中|國小|學校|幼兒園)([\\u4e00-\\u9fa5]{2,8}?)[ \\u3000]+([\\u4e00-\\u9fa5]{2,4})[ \\u3000]*[（(：:]';

/**
 * 找出承辦人。兩種批核軌跡格式標的位置不一樣：
 *   「意見及流程」—— 行首直接標 承辦，後面接機關全名和姓名
 *   「批核軌跡及意見」—— 編號式，簽核人的「下一行」才標 承辦意見
 * 兩種都是公文明寫的，不是猜的。
 */
function findSubmitter(text) {
  let m = text.match(new RegExp(`(?:^|\\n)[ \\u3000]*承辦[ \\u3000][^\\n]{0,60}?${SIGNER_TAIL}`));
  if (m) return { title: m[1], name: m[2], certain: true };

  m = text.match(new RegExp(`${SIGNER_TAIL}[^\\n]*\\n[ \\u3000]*承辦意見`));
  if (m) return { title: m[1], name: m[2], certain: true };

  // 都找不到才退回「軌跡第 1 筆通常是承辦人」，這種要人工確認
  const idx = text.search(/批核軌跡|意見及流程/);
  if (idx >= 0) {
    m = text.slice(idx).match(new RegExp(`(?:^|\\n)[ \\u3000]*1[ \\u3000]*[.、][^\\n]*?${SIGNER_TAIL}`));
    if (m) return { title: m[1], name: m[2], certain: false };
  }
  return null;
}

/** 把擬辦裡的「本人」對應到承辦人 */
function matchSelfReference(block, roster, submitter) {
  if (!submitter) return [];
  const person = roster.find(p => p.name === submitter.name);
  if (!person) return [];

  const hits = [];
  const re = /本人|本席/g;
  let m;
  while ((m = re.exec(block)) !== null) {
    hits.push({
      person,
      pos: m.index,
      written: m[0],
      confidence: 'self',
      selfCertain: !!submitter.certain,
      selfTitle: submitter.title,
    });
  }
  return hits;
}


/* -----------------------------------------------------------
   6c. 純職稱比對
   「教導主任，教學組長，訓導組長各嘉獎1次」連姓都沒有。
   名冊上的職稱多半是唯一的，仍然查得到人。

   名冊就是敘獎範圍：不在名冊上的人不列入，他曾任什麼職務也不構成問題。
   所以職稱完整命中且名冊上唯一時直接核章；只有名冊上兩人以上符合，
   或現職職稱對不上、要靠沿革才查得到時，才需要人工挑。
   ----------------------------------------------------------- */

// 「敬會人事主任」是會辦程序不是受獎人，這些前置詞要擋掉
const PROCEDURAL = /(敬會|會辦|文會|陳送|陳核|送請|知會|簽會|轉陳|呈)[^。；，,]{0,8}$/;

function matchByTitle(block, roster, taken = new Set(), history = null, era = null) {
  const hits = [];
  const re = /([\u4e00-\u9fa5]{2})?(主任|組長|護理師|幹事|校長|執秘|秘書|書記|管理員)/g;
  let m;

  while ((m = re.exec(block)) !== null) {
    const [full, qualifier, role] = m;
    if (taken.has(m.index)) continue;
    if (PROCEDURAL.test(block.slice(Math.max(0, m.index - 12), m.index))) continue;

    const roleKeys = ROLE_ALIAS[role] || [role];
    const hasRole = p => roleKeys.some(k => p.title.includes(k));
    let cands = [], written = full;
    // 記錄是靠哪一層對上的，確認的必要性差很多：
    //   full —— 名冊職稱完整包含公文寫的職稱，跟寫姓名一樣確定
    //   unit —— 職稱對不上，靠單位才分辨出來
    //   role —— 只剩職務別可比
    //   history —— 現職對不上，只有批核軌跡查得到
    let tier = 'role';

    if (qualifier) {
      const keys = UNIT_ALIAS[qualifier] || [qualifier];

      // 第一層：職稱本身就寫全了（教學組長、訓導組長）
      cands = roster.filter(p => keys.some(k => p.title.includes(k + role)));
      if (cands.length) { written = qualifier + role; tier = 'full'; }

      // 第二層：職稱沒寫全，改看單位。
      // 例：公文寫「教導主任」，名冊職稱是「教師兼教務主任」但單位是「教導處」，
      //     用單位就能和總務主任區分開來。
      if (!cands.length) {
        cands = roster.filter(p => hasRole(p) && keys.some(k => p.title.includes(k)));
        if (!cands.length) {
          cands = roster.filter(p => hasRole(p) && keys.some(k => p.unit.includes(k)));
        }
        if (cands.length) { written = qualifier + role; tier = 'unit'; }
      }
    }

    // 第三層：只剩職務別可比，會有一串候選人讓使用者挑
    if (!cands.length) {
      cands = roster.filter(hasRole);
      written = full;
    }

    // 第四層：補上批核軌跡查到的歷任者。
    // 名冊只有現職，職務交接過的話現職者不見得是公文所指的人。
    const past = [];
    if (history) {
      const key = qualifier ? qualifier + role : full;
      const rec = history.get(key) || history.get(full) || history.get(role);
      for (const r of rec || []) {
        const p = roster.find(x => x.name === r.name);
        if (p && !cands.includes(p)) { cands.push(p); past.push({ p, ...r }); }
        else if (p) past.push({ p, ...r });
      }
    }
    if (!cands.length) continue;

    // 公文講的是更早的年度，優先排在任較早的人
    if (past.length > 1 && era && era.past) {
      const order = new Map(past.map((x, i) => [x.p.id, i]));
      cands.sort((a, b) => (order.has(a.id) ? order.get(a.id) : 99) -
                           (order.has(b.id) ? order.get(b.id) : 99));
    }

    // 選中的人現職職稱對不上，表示是靠沿革查到的 —— 這種一定要人工確認
    const chosen = cands[0];
    const chosenTier = hasRole(chosen) ? tier : 'history';

    // 這個職務在批核軌跡上換過人嗎？
    // 敘獎公文幾乎都是在敘去年的事，光看年度早晚沒有鑑別力；
    // 真正該警戒的是「有證據顯示這個職務交接過」。
    const handover = past.length > 1;

    hits.push({
      person: chosen,
      candidates: cands,
      titleTier: chosenTier,
      titleHandover: handover,
      tenure: Object.fromEntries(past.map(x => [x.p.id, `${fmtRoc(x.first)}–${fmtRoc(x.last)} 時任`])),
      pos: m.index,
      written,
      // 職稱完整命中且只有一個人，確定性等同寫出姓名
      confidence: cands.length > 1 ? 'ambiguous'
        : (chosenTier === 'full' ? 'titleExact' : 'title'),
    });
  }
  return hits;
}

/**
 * 判斷公文講的是不是更早的年度。
 * 「113學年度推動喜閱網」但公文發於 114 年 —— 敘的是去年的事，
 * 這時現職者不一定是當時的人。
 */
function detectEra(subject, docYear) {
  const m = (subject || '').match(/(\d{2,3})\s*學?年度/);
  if (!m || !docYear) return null;
  const y = Number(m[1]);
  return { year: y, past: y < Number(docYear) };
}

/** 完整比對跑完後，依序用姓+職稱、純職稱補沒抓到的人 */
function matchAll(block, roster, history = null, era = null, submitter = null) {
  const found = matchNames(block, roster);
  const cover = (list) => {
    const s = new Set();
    for (const h of list) for (let i = 0; i < h.written.length; i++) s.add(h.pos + i);
    return s;
  };
  // 明確姓名可以在不同位置重複出現；推定的職稱則不可再製造同人的假重複。
  const hasPerson = (list, h) => list.some(d => d.person.id === h.person.id);

  const byRole = matchByRole(block, roster, cover(found)).filter(h => !hasPerson(found, h));
  const stage2 = [...found, ...byRole];

  // 職稱旁邊如果已經寫了姓名，那個職稱是在描述那個人，不是另一個受獎人。
  // 例：「時任資訊執秘為:吳孟謙老師」—— 資訊執秘講的就是吳孟謙，
  // 不該再把現任的資訊執秘也算進來。
  const named = (h) => stage2.some(d => {
    const gapAfter = d.pos - (h.pos + h.written.length);
    const gapBefore = h.pos - (d.pos + d.written.length);
    return (gapAfter >= 0 && gapAfter <= 12) || (gapBefore >= 0 && gapBefore <= 4);
  });

  const byTitle = matchByTitle(block, roster, cover(stage2), history, era)
    .filter(h => !hasPerson(stage2, h) && !named(h));

  // 「本人」指的是承辦人。已經有明確姓名的位置就不用再解一次。
  const stage3 = [...stage2, ...byTitle];
  const bySelf = matchSelfReference(block, roster, submitter)
    .filter(h => !hasPerson(stage3, h));

  return [...stage3, ...bySelf]
    .sort((a, b) => a.pos - b.pos)
    .map(h => ({
      ...h,
      duty: detectDuty(block, h.pos, h.written) || '',
      assignment: detectAssignment(block, h.pos) || '',
    }));
}


/* -----------------------------------------------------------
   6b. 姓 + 職稱 推定
   「教務林主任」「輔導謝組長」這種只有姓沒有名的寫法，
   靠「單位 + 姓 + 職稱」三個條件回名冊反查。
   ----------------------------------------------------------- */

// 學務處和訓導處是同一個單位的新舊稱呼，要互通
const UNIT_ALIAS = {
  教務: ['教務'], 學務: ['學務', '訓導'], 訓導: ['訓導', '學務'],
  輔導: ['輔導'], 總務: ['總務'], 教導: ['教導'],
  人事: ['人事'], 主計: ['主計', '會計'], 會計: ['會計', '主計'],
  圖書: ['圖書'], 幼兒園: ['幼兒園', '幼稚園'],
};

const ROLE_ALIAS = {
  主任: ['主任'], 組長: ['組長'], 老師: ['教師'], 教師: ['教師'],
  護理師: ['護理師', '護士'], 幹事: ['幹事'], 校長: ['校長'],
  執秘: ['執秘', '執行秘書'], 秘書: ['秘書'], 書記: ['書記'], 管理員: ['管理員'],
};

const ROLE_RE = /(主任|組長|老師|教師|護理師|幹事|校長|執秘|秘書|書記|管理員)/;
const UNIT_RE = /(教務|學務|訓導|輔導|總務|教導|人事|主計|會計|圖書|幼兒園)/;

/**
 * 掃描「(單位)姓職稱」樣式。
 * 姓必須真的是名冊上某人的姓，否則不成立 —— 這一條同時擋掉了
 * 「訓導組長」被誤讀成姓「導」的情況。
 */
function matchByRole(block, roster, taken = new Set()) {
  const hits = [];
  // UNIT_RE / ROLE_RE 本身已經帶括號，外層要用非捕獲群組，
  // 否則群組編號會位移，抓到的「姓」其實是單位名。
  const re = new RegExp(`(?:${UNIT_RE.source})?([\\u4e00-\\u9fa5])${ROLE_RE.source}`, 'g');
  let m;

  while ((m = re.exec(block)) !== null) {
    const [, unit, surname, role] = m;
    const surnamePos = m.index + (unit ? unit.length : 0);
    if (taken.has(surnamePos)) continue;

    const roleKeys = ROLE_ALIAS[role] || [role];
    const unitKeys = unit ? (UNIT_ALIAS[unit] || [unit]) : null;

    let cands = roster.filter(p =>
      p.name[0] === surname && roleKeys.some(k => p.title.includes(k)));

    // 單位要分層比。職稱寫明單位（教師兼教務處教務主任）比單位欄位精確，
    // 混在一起比會讓「幼稚園主任但單位掛在教務處」的人也符合「教務林主任」。
    if (unitKeys && cands.length > 1) {
      const byTitle = cands.filter(p => unitKeys.some(k => p.title.includes(k)));
      if (byTitle.length) cands = byTitle;
      else {
        const byUnit = cands.filter(p => unitKeys.some(k => p.unit.includes(k)));
        if (byUnit.length) cands = byUnit;
      }
    }
    if (!cands.length) continue;

    hits.push({
      person: cands[0],
      candidates: cands,
      pos: surnamePos,
      written: (unit || '') + surname + role,
      confidence: cands.length === 1 ? 'role' : 'ambiguous',
    });
  }
  return hits;
}

/* -----------------------------------------------------------
   7. 獎度解析
   ----------------------------------------------------------- */
const CN_NUM = {
  '一': 1, '乙': 1, '二': 2, '兩': 2, '三': 3,
  '１': 1, '２': 2, '1': 1, '2': 2, '3': 3,
};

function findAwards(block) {
  const out = [];
  // 「嘉獎2次」「嘉一次」（漏字）「嘉獎1支」「記功乙次」都要抓得到
  const re = /(嘉獎?|記功|記大功)[裝訂線]?\s*([一乙二兩三１２1-3])\s*[次支個]/g;
  let m;
  while ((m = re.exec(block)) !== null) {
    const kind = m[1].startsWith('嘉') ? '嘉獎' : (m[1] === '記功' ? '記功' : '記大功');
    const n = CN_NUM[m[2]] ?? 1;
    let code = null;
    if (kind === '嘉獎') code = n === 2 ? '4002' : '4001';
    else if (kind === '記功' && n === 1) code = '4010';

    // 「各嘉獎1支」「均記功1次」— 這個獎度是整份名單共用的，不屬於某一個人
    const lead = block.slice(Math.max(0, m.index - 8), m.index);
    const shared = /各|均|每人|分別|同敘/.test(lead);

    // 「嘉獎2次1人：蕭宗仁主任」— 獎度寫在前面當標題，後面才接名單。
    // 分隔號限定冒號和破折號，不能收「、」「，」，
    // 否則「…敘嘉獎2次、林怡慈老師…」會被誤判成標題，把下一個人吃掉。
    const afterPos = m.index + m[0].length;
    const mh = block.slice(afterPos, afterPos + 10).match(/^\s*(?:\d+\s*人)?\s*[：:\-－—]\s*/);

    out.push({
      pos: m.index,
      code,
      label: `${kind}${n === 2 ? '二' : n === 3 ? '三' : '一'}次`,
      raw: m[0],
      shared,
      heading: !!mh,
      headingEnd: mh ? afterPos + mh[0].length : null,
    });
  }
  return out;
}

/**
 * 後備：擬辦沒寫獎度時，從來文本文找。
 * 例：「茲同意每人敘嘉獎1次」— 獎度在函裡，擬辦只寫誰獲獎。
 */
function findFallbackAward(text) {
  const normalized = String(text || '')
    .replace(/[奬奖]/g, '獎')
    .replace(/[敍叙]/g, '敘');
  let m = normalized.match(/(?:敘|核予|同意|各)[^。\n]{0,24}?(嘉獎?|記功)\s*([一二三１２1-3])\s*次/);
  // 評鑑公文常在擬辦只寫「本校成績為優等」，獎度另列於來文規則。
  // 這種是照本校等第精準查到的，即使來文列了多種等第也不算不確定。
  let precise = false;
  let grade = '';
  if (!m) {
    const g = normalized.match(/成績為[：:]?(特優|優等|甲等)/);
    if (g) {
      const rule = normalized.match(
        new RegExp(`${g[1]}[：:]?[^。\\n]{0,50}?(嘉獎?|記功)\\s*([一二三１２1-3])\\s*次`),
      );
      if (rule) { m = rule; precise = true; grade = g[1]; }
    }
  }
  if (!m) return null;
  const kind = m[1].startsWith('嘉') ? '嘉獎' : '記功';
  const n = CN_NUM[m[2]] ?? 1;
  const code = kind === '嘉獎' ? (n === 2 ? '4002' : '4001') : (n === 1 ? '4010' : null);
  // 全文只出現一種獎度時，套用給誰都不會錯；出現多種才有選錯的風險。
  const levels = new Set();
  for (const a of findAwards(normalized)) if (a.code) levels.add(a.code);
  return {
    pos: -1, code, label: `${kind}${n === 2 ? '二' : '一'}次`,
    raw: m[0], fallback: true, precise, grade,
    sole: levels.size <= 1,
    levels: [...levels],
  };
}

/**
 * 擬辦完全沒有名單時，保守檢查「批核意見」中的明確敘獎句。
 * 只取批核意見冒號後的文字，不讀簽核者職稱／姓名；
 * 且必須同時命中名冊姓名與獎度，才提供未核章的人工確認資料。
 */
function findApprovalAwards(text, roster) {
  const candidates = [];
  for (const m of String(text || '').matchAll(
    /批核意見[：:]([\s\S]{0,500}?)(?=\n\s*\d+\.\s|欄位批核紀錄|貼紙備註資訊|$)/g,
  )) {
    const block = cleanBlock(m[1]);
    const awards = findAwards(block);
    if (!awards.length) continue;
    const hits = matchNames(block, roster);
    if (!hits.length) continue;
    const paired = pairNameAward(block, hits, awards)
      .filter(item => item.award)
      // 保留原本的比對強度：姓名完全相符才能自動核章
      .map(item => ({ ...item, baseConfidence: item.confidence, confidence: 'approval' }));
    if (paired.length) candidates.push({ block, awards, paired, index: m.index });
  }
  if (!candidates.length) return null;
  return candidates.sort((a, b) => b.paired.length - a.paired.length)[0];
}

/**
 * 把姓名和獎度配對。
 *
 * 四階段認領：
 *   1. 標題式獎度
 *   2. 共用獎度（各／均）
 *   3. 一般敘述依版面方向認領
 *   4. 尚未配到的再從反方向補找
 *
 * 單純取最近距離會出錯：名單連寫時，後一個人的姓名離「前一個人的獎度」
 * 反而比離自己的獎度更近。
 *
 * 一般公文多是「姓名…獎度」，但表格式擬辦常是「獎度／職稱／姓名」。
 * 先比較整段中獎度落在姓名前後的命中數，再決定一般認領方向，
 * 可避免學生健檢表格把郭喜淑的嘉獎二次錯配成下一列的嘉獎一次。
 */
function segmentAt(block, pos) {
  // 句段邊界：句號、分號、換行，以及「1.」「一、」這種條列編號
  // 數字條列前一字不得也是數字，避免把 90.91% 的小數點當成新段落。
  const re = /[。；\n]|(?<![0-9])[0-9]\s*[.、]|[一二三四五六七八九十]\s*、/g;
  let start = 0, end = block.length, m;
  while ((m = re.exec(block)) !== null) {
    if (m.index < pos) start = m.index + m[0].length;
    else { end = m.index; break; }
  }
  return [start, end];
}

function pairNameAward(block, hits, awards, opts = {}) {
  const FWD = opts.forward ?? 22;
  const BACK = opts.backward ?? 25;

  const claimed = new Set();
  const result = hits.map(h => ({ ...h, award: null }));

  // 第一階段：標題式獎度「嘉獎2次1人：甲主任、嘉獎1次3人：乙老師、丙主任」
  // 一個標題管到下一個獎度出現為止，中間列的人全部適用。
  awards.forEach((a, i) => {
    if (!a.heading) return;
    const next = awards.find(b => b.pos > a.pos);
    const end = next ? next.pos : segmentAt(block, a.pos)[1];
    const inSpan = result.filter(r => !r.award && r.pos >= a.headingEnd && r.pos < end);
    if (!inSpan.length) return;
    for (const r of inSpan) r.award = a;
    claimed.add(i);
  });

  // 第二階段：共用獎度「…各嘉獎1支」不歸某一個人，而是同一句裡列出的每個人都有。
  // 範圍限縮在同一個句段，否則會外溢到後面不相干的名單。
  for (const a of awards) {
    if (!a.shared) continue;
    const [s0, s1] = segmentAt(block, a.pos);
    // 同一句前面若已有另一個獎度，共用範圍從那個獎度之後才開始。
    // 例：「蘇宇祥嘉獎2次，教導主任、總務主任各嘉獎1次」，
    // 後面的「各嘉獎1次」不能回頭覆蓋蘇宇祥。
    const previous = awards
      .filter(other => other !== a && other.pos >= s0 && other.pos < a.pos)
      .sort((x, y) => y.pos - x.pos)[0];
    const scopeStart = previous ? previous.pos + previous.raw.length : s0;
    // 「各／每人」獎度通常寫在名單後面，只套用獎度之前列出的人；
    // 避免「甲乙各嘉獎1次，丙嘉獎2次」把丙也先吃成嘉獎1次。
    const inSeg = result.filter(r => r.pos >= scopeStart && r.pos < a.pos);
    for (const r of inSeg) if (!r.award) r.award = a;
  }

  const ordinary = awards
    .map((award, index) => ({ award, index }))
    .filter(({ award }) => !award.heading && !award.shared);

  // 沒寫「各」但明顯是群組名單：「協助人員甲、乙、丙嘉獎1次」。
  // 只在同句、前一個獎度之後至少列出兩人時才視為共用，避免吃到前一組名單。
  for (const { award: a, index } of ordinary) {
    const [s0] = segmentAt(block, a.pos);
    const previous = awards
      .filter(other => other !== a && other.pos >= s0 && other.pos < a.pos)
      .sort((x, y) => y.pos - x.pos)[0];
    const scopeStart = previous ? previous.pos + previous.raw.length : s0;
    const inGroup = result.filter(r => !r.award && r.pos >= scopeStart && r.pos < a.pos);
    if (inGroup.length < 2) continue;
    const lead = block.slice(scopeStart, inGroup[0].pos);
    if (!/(人員|名單|導師|教師|主任|組長|為|[:：])/.test(lead)) continue;
    for (const r of inGroup) r.award = a;
    a.group = true;
    claimed.add(index);
  }

  const hasNear = (r, direction) => ordinary.some(({ award }) => {
    const d = direction === 'forward' ? award.pos - r.pos : r.pos - award.pos;
    return d >= 0 && d <= (direction === 'forward' ? FWD : BACK);
  });
  const forwardHits = result.filter(r => !r.award && hasNear(r, 'forward')).length;
  const backwardHits = result.filter(r => !r.award && hasNear(r, 'backward')).length;
  const firstDirection = backwardHits > forwardHits ? 'backward' : 'forward';

  const claimNearest = (direction) => {
    const limit = direction === 'forward' ? FWD : BACK;
    for (const r of result) {
      if (r.award) continue;
      let best = null, bestD = Infinity;
      for (const { award, index } of ordinary) {
        if (claimed.has(index)) continue;
        const d = direction === 'forward' ? award.pos - r.pos : r.pos - award.pos;
        if (d < 0 || d > limit) continue;
        if (d < bestD) { bestD = d; best = index; }
      }
      if (best !== null) { claimed.add(best); r.award = awards[best]; }
    }
  };

  // 第三、四階段：依整段版面方向認領，再從反方向補找。
  claimNearest(firstDirection);
  claimNearest(firstDirection === 'forward' ? 'backward' : 'forward');

  return result;
}


/**
 * 評估整份公文的解析可信度。
 * 姓名數和獎度數對不起來，通常表示 PDF 的表格排版在抽文字時被打亂了，
 * 這種情況配對結果不可信，要整份標記人工核對。
 */
/**
 * blockAll 表示問題會動搖整份公文的配對可信度，所有列都不自動核章。
 * 沒標的警告只是在描述個別列的狀況 —— 那些列本身就過不了自動核章的門檻，
 * 不需要把同一份公文裡其他已經確定的人一起扣住。
 */
function assessDoc(block, hits, awards, fallback = null, paired = null) {
  const hasFallbackAward = fallback && typeof fallback === 'object' ? true : !!fallback;
  const soleFallback = fallback && typeof fallback === 'object'
    ? (fallback.sole !== false || fallback.precise === true)
    : false;

  if (!block) return { level: 'fail', blockAll: true, note: '找不到擬辦區塊，無法解析名單' };
  if (!hits.length) {
    return /主任|組長|導師|護理師|校長/.test(block)
      ? { level: 'fail', blockAll: true, note: '擬辦只寫職稱沒有姓名，需自行指定人員' }
      : { level: 'fail', blockAll: true, note: '擬辦區塊中找不到名冊上的人員' };
  }
  if (!awards.length) {
    if (!hasFallbackAward) {
      return { level: 'fail', blockAll: true, note: '擬辦與來文全文都找不到明確獎度，請人工確認' };
    }
    // 全文只有一種獎度，套用給誰都一樣，不必逐筆確認
    if (soleFallback) {
      const why = fallback && fallback.precise
        ? `擬辦沒寫獎度，已依本校成績「${fallback.grade}」對到來文的獎度。`
        : '擬辦沒寫獎度，已採用來文本文的獎度（全文僅此一種）。';
      return { level: 'info', blockAll: false, note: why };
    }
    // 來文列了多種獎度又查不到本校適用哪一種，把選項講出來讓人比較好挑
    const opts = (fallback && fallback.levels || [])
      .map(c => Object.entries(AWARD_CODES).find(([, v]) => v === c))
      .filter(Boolean).map(([label]) => label);
    return {
      level: 'warn', blockAll: true,
      note: `擬辦沒寫獎度，來文有${opts.length ? `${opts.join('、')}` : '多種獎度'}，請確認每位適用哪一種。`,
    };
  }
  if (paired) {
    const unmatched = paired.filter(item => !item.award).length;
    if (unmatched) {
      return {
        level: 'warn', blockAll: true,
        note: `有 ${unmatched} 位人員未能配到獎度，請查看原公文後人工指定`,
      };
    }
  }
  const shared = awards.some(a => a.shared || a.heading || a.group);
  if (awards.length && !shared && hits.length !== awards.length) {
    return { level: 'warn', blockAll: true, note: `找到 ${hits.length} 位人員但 ${awards.length} 個獎度，配對可能錯位，請逐筆核對` };
  }
  // 真的需要你做決定的，是「名冊上有兩個以上的人都符合」。
  // 離職者不在名冊上就不列入，他曾任什麼職務也不構成問題，不必為此發警告。
  const pickNeeded = hits.filter(h => h.confidence === 'ambiguous').length;
  if (pickNeeded) {
    return { level: 'warn', blockAll: false, note:
      `有 ${pickNeeded} 個職稱在名冊上不只一人符合，請點選正確的人員。` };
  }
  const byTitle = hits.filter(h => h.confidence === 'title').length;
  if (byTitle) {
    return { level: 'warn', blockAll: false, note: `有 ${byTitle} 位公文只寫職稱，職稱沒完全對上，請確認人員。` };
  }
  // 姓＋職稱＋單位三個條件都吻合且名冊上唯一，確定性足夠，
  // 只說明比對依據，不要求人工確認。
  const byRole = hits.filter(h => h.confidence === 'role').length;
  if (byRole) {
    return { level: 'info', blockAll: false, note: `有 ${byRole} 位公文寫的是「單位＋姓＋職稱」，已對回名冊人員。` };
  }
  return { level: 'ok', blockAll: false, note: '' };
}


/* -----------------------------------------------------------
   8. 獎懲事由草擬
   ----------------------------------------------------------- */
const SUBJECT_NOISE = [
  /^有關/, /^為/, /^關於/,
  /貴校|貴屬|本市各級學校|本市公私立|本局/g,
];

/**
 * 把公文主旨清成可以當事由的核心描述。
 *
 * 公文主旨是「教育局對學校講話」的語氣，事由是「描述這個人做了什麼」，
 * 兩者視角不同。檢送、請逕依權責辦理、詳如說明這些是行文用語，
 * 不是在描述工作內容，全部要砍掉。
 */
function cleanSubject(subject) {
  let s = String(subject || '');

  // PDF 抽文字時，特殊符號有時會被讀成注音符號或全形記號。
  // 這些不是內容，出現在事由裡是雜訊。
  s = s.replace(/[ˇˊˋ˙ˉ｀´]/g, '');

  // 一、開頭的發語詞
  s = s.replace(/^(有關|為|關於|茲)/, '');

  // 二、「檢送…(如附件)」這種轉發語氣。附件標註連同括號一起砍
  s = s.replace(/^檢送/, '')
       .replace(/[（(](?:如)?附件[0-9]*[）)]/g, '')
       .replace(/[（(]詳如[^）)]{0,12}[）)]/g, '');

  // 三、句尾的請求語。這些之後的內容全部不要
  s = s.replace(
    /[，,、]?\s*(?:詳如說明|請查照|請依說明[^，。]*|請依敘獎清冊[^，。]*|請各校[^，。]*|請逕依[^，。]*|請惠予[^，。]*|請依權責[^，。]*|請貴校[^，。]*|以資鼓勵|如說明)[\s\S]*$/g,
    '');

  // 四、說明受獎對象的子句，連同後面的「敘獎N案」一起砍。
  // 必須排在案件後綴之前 —— 否則「敘獎1案」先被砍掉，
  // 「…之現職教師及行政人員」的「人員」也會被吃掉，留下半截子句。
  const TAIL = '(?:敘獎|獎勵|敘獎獎勵)?\\s*\\d*\\s*(?:案|一案|乙案|事宜)?$';
  const TARGETS = [
    `[之]?(?:現職)?(?:教師|人員|同仁)(?:及|與|暨)(?:相關)?(?:行政)?(?:人員|同仁|教師)${TAIL}`,
    `[之]?(?:相關)?(?:工作|承辦|協辦|業務)(?:人員|同仁)${TAIL}`,
    `[之]?(?:現職|相關|有功)(?:教師|人員|同仁)${TAIL}`,
    `[之]?(?:資訊|午餐|衛生|體育)?教育人員${TAIL}`,
  ];
  for (const src of TARGETS) {
    const next = s.replace(new RegExp(src), '');
    if (next !== s) { s = next; break; }
  }

  // 說明產出而非工作內容的子句。事由要講的是辦了什麼，不是產出什麼文件，
  // 必須排在名詞後綴之前，否則「結果」「名單」先被砍掉會留下「統計」「優勝」。
  const OUTPUT_TAILS = [
    /[之]?(?:競賽|比賽|評鑑|考核|測驗|統計)?(?:結果|成果)(?:及|與|暨)(?:達標|獲獎|優勝|得獎)(?:學校|名單|單位)?$/,
    /[之]?(?:競賽|比賽|評鑑|考核|測驗|統計)(?:結果|成果)$/,
    /[之]?(?:達標|獲獎|優勝|得獎)(?:學校|名單|單位)$/,
    /[之]?(?:結果|成果)$/,
  ];
  for (const re of OUTPUT_TAILS) {
    const next = s.replace(re, '');
    if (next !== s) { s = next; break; }
  }

  // 五、剩下的案件性質後綴。公文常寫「敘獎1案」表示件數，
  // 數字要連同「案」一起砍，否則會留下孤零零的數字。
  s = s.replace(/[之]?(?:有功|相關)?(?:人員|教師|同仁)?(?:敘獎|獎勵|敘獎獎勵)\s*\d*\s*(?:案|一案|事宜|乙案)?$/, '')
       .replace(/(?:成績|名單|清冊|一覽表|成果報告)$/, '')
       .replace(/\s*\d+\s*案$/, '')
       // 「案」單獨結尾才砍。「實施方案」「專案」「方案」是計畫名稱的一部分，
       // 前面是這些字時不能動。
       .replace(/(?<![方專議提個本該此草結專])案$/, '');

  // 六、指稱行文雙方的詞。事由是描述工作內容，不需要這些
  s = s.replace(/貴校|貴屬|貴單位|本市公私立高級中等以下學校|本局所屬|所屬機關學校/g, '')
       .replace(/^本局|^本校/, '');

  // 七、開頭殘留的動詞。外層還會加「協助辦理」，留著會變成「協助辦理辦理…」
  s = s.replace(/^(?:協助)?(?:辦理|承辦|協辦|推動|舉辦)/, '');

  // 八、末尾殘留的名詞化後綴。語料顯示事由講的是「做了什麼事」，
  // 不是「產出什麼文件」，所以優勝名單、成績一覽表這類要砍
  s = s.replace(/[之]?(?:優勝|得獎|獲獎)?(?:名單|清冊|一覽表|統計表|成績|結果)$/, '')
       .replace(/[之]?(?:相關)?(?:事宜|事項|工作)$/, '');

  return s.replace(/[，。、；;之及與暨和]+$/, '').trim();
}


function draftReason(meta, block, opts = {}) {
  const { withBasis = true, limit = 100 } = opts;

  // 優先取主旨裡「」引號內的計畫名稱 — 這通常就是事由核心
  let core = '';
  // 引號要成對抓。公文常見巢狀寫法「114年度『高雄數位學園』推廣計畫」，
  // 若用 [「『]...[」』] 會從外引號開始卻停在內引號的收尾，切出殘缺的名稱。
  const quoted =
    (meta.subject || '').match(/「([^「」]{4,60})」/) ||
    (meta.subject || '').match(/『([^『』]{4,60})』/);
  if (quoted) {
    core = quoted[1];
  } else {
    core = cleanSubject(meta.subject || '');
  }
  core = core
    .replace(/^[「『]|[」』]$/g, '')
    .replace(/統計結果$|相關事宜$|事宜$/g, '')
    .replace(/[及與暨和，、．。]+$/g, '')   // 截掉後留下的懸空連接詞
    .trim();

  // 主旨清完只剩「…資訊教育人員」這種人物名詞，接上動詞會不通順。
  // 改從說明段找真正描述做了什麼的句子。
  let verbFromBody = '';
  if (!quoted && /(人員|教師|學校|人|者)$/.test(core)) {
    const m = (meta.body || '').match(
      /(協助|辦理|推動|參與|承辦|協辦)[^，。、）\n]{4,40}?(業務|工作|計畫|活動|方案|任務|事項)/);
    if (m) { core = m[0]; verbFromBody = m[1]; }
  }

  // 動詞：擬辦裡自稱承辦就用「辦理」，否則用「協助辦理」
  const verb = verbFromBody ? '' : (/承辦人員|業務承辦|主要業務承辦/.test(block) ? '辦理' : '協助辦理');

  let main = core ? `${verb}${core}，辛勞得力` : '辦理相關業務，辛勞得力';

  let basis = '';
  if (withBasis && meta.agency && meta.year && meta.docNo) {
    basis = `依據${meta.agency}${meta.year}年${meta.month}月${meta.day}日${meta.docNo}函辦理`;
  }

  // 100 字上限：優先保住依據子句，前段不夠就截
  let full = basis ? `${main}\n${basis}` : main;
  if (full.length > limit) {
    const room = limit - (basis ? basis.length + 1 : 0) - 5;
    if (room > 12) {
      main = `${verb}${core.slice(0, Math.max(6, room - verb.length - 5))}，辛勞得力`;
      full = basis ? `${main}\n${basis}` : main;
    }
    if (full.length > limit) full = full.slice(0, limit);
  }
  return full;
}

/** 將同案不同分工寫進事由，避免同一人的多筆資料完全相同。 */
function applyDutyToReason(reason, duty) {
  if (!duty) return reason;
  const parts = String(reason || '').split('\n');
  const main = parts.shift() || '';
  const core = main.replace(/^(?:協助辦理|辦理|主辦|承辦|協辦|協助)/, '');
  const verb = duty === '協助' ? '協助辦理' : duty;
  return [`${verb}${core}`, ...parts].join('\n');
}

/**
 * 將同一人的每次分工／負責範圍寫進事由，並維持 WebHR 的 100 字限制。
 */
/**
 * 同一人多筆時自動把事由寫成不同的。
 *
 * WebHR 不接受同一人有獎懲類別、事由、核定日期完全相同的兩筆，
 * 所以事由沒區別開就匯不進去。分工（承辦／協辦）和範圍（一忠導師、國語）
 * 都辨識不出來時，改用各筆在擬辦裡所處的句段當區別依據 ——
 * 同一人被列兩次，本來就是寫在兩個不同的地方。
 */
function detectScope(block, pos, written = '') {
  const [s0, s1] = segmentAt(block, pos);
  const AWARD = /(嘉獎?|記功|記大功)[裝訂線]?\s*[一乙二兩三１２1-3]\s*[次支個]/;

  // 競賽類公文同一人多筆，差別多半在班級、組別或名次，優先抓這些
  const near = block.slice(Math.max(0, pos - 24), pos + written.length + 24);
  const CONTEST = [
    /([一二三四五六1-6]\s*年?[甲乙丙丁戊己庚忠孝仁愛信義和平]\s*班?)/,
    /((?:低|中|高)年級組|幼兒?組|團體組|個人組)/,
    /(特優|優等|甲等|乙等|第[一二三四五六1-6]名|冠軍|亞軍|季軍)/,
    /((?:國語文?|數學|英語|自然|社會|資訊|體育|音樂|美術)科?)/,
  ];
  for (const re of CONTEST) {
    const m = near.match(re);
    if (m) return m[1].replace(/\s+/g, '');
  }

  let lead = block.slice(s0, pos)
    .replace(new RegExp(AWARD.source, 'g'), '')
    .replace(/[（(][^）)]{0,20}[）)]/g, '')
    .replace(/^[\s\d.、）)]+/, '');

  // 冒號後面接的是名單，冒號前面那句才是這一組的工作名稱
  const byColon = lead.match(/([\u4e00-\u9fa5][^：:；;。]{3,24})[：:]\s*[^：:]*$/);
  if (byColon) return trimScope(byColon[1]);

  // 描述也可能寫在姓名後面：「秦桂屏辦理線上課程建置嘉獎1次」
  let after = block.slice(pos + written.length, s1);
  const am = after.match(AWARD);
  if (am) after = after.slice(0, am.index);
  after = after
    .replace(/^(?:老師|教師|主任|組長|護理師|幹事|工友)?[：:、，,]*/, '')
    .replace(/[、，,：:；;]+$/, '');
  if (after.length >= 4) return trimScope(after);

  // 都沒有就取句段開頭的描述，砍掉尾端殘留的姓名與標點
  const plain = lead.replace(/[、，,]\s*[\u4e00-\u9fa5]{2,4}\s*$/, '')
                    .replace(/[、，,：:；;\s]+$/, '');
  return plain.length >= 4 ? trimScope(plain) : '';
}

/**
 * 跨公文的重複事由檢查。
 *
 * 單份公文內的區別是各自做的，不同公文的（其1）會撞在一起。
 * WebHR 以「身分證號＋獎懲類別＋事由」判定重複，一筆撞號會讓整批匯入失敗，
 * 所以所有公文都解析完之後要再做一次全域檢查。
 */
function dedupeAcrossDocs(rows, limit = 100) {
  const seen = new Map();
  const fixed = [];

  for (const row of rows) {
    const key = `${row.person.id}|${row.category}|${row.reason}`;
    if (!seen.has(key)) { seen.set(key, 1); continue; }

    // 已經出現過，加上流水號讓它不一樣
    const n = seen.get(key) + 1;
    seen.set(key, n);

    const parts = String(row.reason).split('\n');
    const head = parts.shift() || '';
    const suffix = parts.length ? `\n${parts.join('\n')}` : '';
    // 已經有（其N）就換掉，沒有才加
    const base = head.replace(/（其\d+）/g, '');
    let main = insertScope(base, `其${n}`);
    const room = Math.max(0, limit - suffix.length);
    if (main.length > room) main = main.slice(0, room);
    row.reason = main + suffix;
    row.crossDup = true;
    fixed.push(row.person.name);
  }
  return [...new Set(fixed)];
}

/**
 * 區別詞能不能用。
 * 抓到的如果是擬辦的行文說明而不是工作範圍，寧可不用 ——
 * 括號裡出現別人的姓名或「敘獎人員為…」這種句子很不妥。
 */
function usableScope(scope, roster) {
  const t = String(scope || '').trim();
  if (t.length < 2 || t.length > 16) return false;

  // 含有名冊上其他人的姓名
  if (roster && roster.some(p => t.includes(p.name))) return false;

  // 行文說明的特徵：敘獎、名單、人員為、以及…
  if (/敘獎|獎勵|名單|人員為|以及|如下|等人|各敘|每人/.test(t)) return false;

  // 職稱本身不算工作範圍
  if (/^(?:老師|教師|主任|組長|導師|護理師|幹事|工友|校長)$/.test(t)) return false;

  return /[\u4e00-\u9fa5]/.test(t);
}

function trimScope(s) {
  return s.replace(/^(協助辦理|協助|辦理|承辦|協辦|擬|建議|敘獎名單如下)/, '')
          .replace(/(人員|名單|如下)$/, '')
          .replace(/[、，,：:；;]+$/, '')
          .trim()
          .slice(0, 24);
}

/** 把區別詞插在「，辛勞得力」之前的括號裡，符合既有事由的寫法 */
function insertScope(head, scope) {
  const i = head.lastIndexOf('，');
  return i < 0 ? `${head}（${scope}）` : `${head.slice(0, i)}（${scope}）${head.slice(i)}`;
}

/** 砍掉事由裡已經講過的字，避免「…高雄數位學園…（數位學園教師研習）」這種重複 */
function trimOverlap(scope, head) {
  let cut = 0;
  for (let n = scope.length; n >= 2; n--) {
    if (head.includes(scope.slice(0, n))) { cut = n; break; }
  }
  if (!cut) return scope;
  const rest = scope.slice(cut).replace(/^[之的與及、，,]+/, '');
  return rest.length >= 2 ? rest : '';
}

/**
 * 把同一人的重複事由改寫成互不相同。
 * 只在真的會撞在一起時才動手，正常情形不改。
 */
function differentiateReasons(block, rows, limit = 100, roster = null) {
  const byPerson = new Map();
  for (const row of rows) {
    const list = byPerson.get(row.person.id) || [];
    list.push(row);
    byPerson.set(row.person.id, list);
  }

  const unresolved = [];
  for (const list of byPerson.values()) {
    if (list.length < 2) continue;
    if (new Set(list.map(r => r.reason)).size === list.length) continue;

    // 先看看能不能抓到堪用又互不相同的區別詞
    const scopes = list.map(row => {
      const raw = detectScope(block, row.sourcePos, row.written || '');
      return usableScope(raw, roster) ? raw : '';
    });
    const good = scopes.every(Boolean) && new Set(scopes).size === scopes.length;

    list.forEach((row, i) => {
      const parts = String(row.reason).split('\n');
      const head = parts.shift() || '';
      const suffix = parts.length ? `\n${parts.join('\n')}` : '';

      // 抓不到好的區別詞就用序號。事由必須互不相同，
      // 否則 WebHR 會以「重複資料」拒絕整批匯入。
      const mark = good ? trimOverlap(scopes[i], head) || scopes[i] : `其${i + 1}`;

      let main = insertScope(head, mark);
      const room = Math.max(0, limit - suffix.length);
      if (main.length > room) main = main.slice(0, room);
      row.reason = main + suffix;
      row.scopeAuto = !good;   // 用序號代替的要提醒使用者補寫
    });

    if (new Set(list.map(r => r.reason)).size !== list.length) {
      unresolved.push(list[0].person.name);
    }
  }
  return unresolved;
}

function applyOccurrenceToReason(reason, duty, assignment, limit = 100) {
  const parts = String(applyDutyToReason(reason, duty) || '').split('\n');
  let main = parts.shift() || '';
  if (assignment) main = `${assignment}${main}`;

  const suffix = parts.length ? `\n${parts.join('\n')}` : '';
  const room = Math.max(0, limit - suffix.length);
  if (main.length > room) main = main.slice(0, room);
  return main + suffix;
}


/* -----------------------------------------------------------
   9. 獎懲類別判斷
   ----------------------------------------------------------- */
function guessCategory(text) {
  const t = text || '';
  if (/比賽|競賽|錦標|球賽|闖關|競技/.test(t)) return 'A07';
  if (/研習|培訓|訓練|進修|工作坊/.test(t)) return 'A06';
  if (/研究|研發|課程發展|教材/.test(t)) return 'A05';
  if (/防疫|疫苗|健檢|健康檢查|衛生/.test(t)) return 'A02';
  return 'A02'; // 預設：工作績優
}


/* -----------------------------------------------------------
   10. 組出一列 CSV 資料
   ----------------------------------------------------------- */
function buildRow(person, awardCode, reason, category, orgCode, opts = {}) {
  const kind = person.kind;
  const law = LAW[kind];
  const cl = (CLAUSE[kind] && CLAUSE[kind][awardCode]) || { 條: null, 點: null, 項: null, 款: null, 目: null };
  const v = (key) => {
    // 使用者挑過的「目」優先，其餘依法條對應表
    if (key === '目' && opts.item) return opts.item;
    return cl[key] === null || cl[key] === undefined ? SENTINEL[key] : cl[key];
  };

  return {
    身分證號: person.id,
    姓名: person.name,
    機關代碼: orgCode,
    獎懲事由: reason,
    獎懲事由長: '',
    獎懲結果: awardCode || '',
    其他事項: '',
    獎懲類別: category,
    適用法規: law.code,
    適用法規名稱: law.name,
    條: v('條'), 點: v('點'), 項: v('項'), 款: v('款'), 目: v('目'),
    教示條款: person.teachClause,
  };
}

const CSV_HEADER = [
  '身分證號', '姓名', '機關代碼', '獎懲事由(更新資料庫)',
  '獎懲事由(長)(非必填，建議不使用)', '獎懲結果', '其他事項', '獎懲類別',
  '適用法規', '適用法規名稱', '條', '點', '項', '款', '目', '教示條款',
];

function rowsToCsv(rows) {
  const esc = (s) => {
    const t = String(s ?? '');
    // 事由裡的換行會讓 WebHR 匯入出錯，一律轉成空白
    const clean = t.replace(/[\r\n]+/g, ' ').replace(/'/g, '');
    return /[",]/.test(clean) ? `"${clean.replace(/"/g, '""')}"` : clean;
  };
  const lines = [CSV_HEADER.join(',')];
  for (const r of rows) {
    lines.push([
      r.身分證號, r.姓名, r.機關代碼, r.獎懲事由, r.獎懲事由長,
      r.獎懲結果, r.其他事項, r.獎懲類別, r.適用法規, r.適用法規名稱,
      r.條, r.點, r.項, r.款, r.目, r.教示條款,
    ].map(esc).join(','));
  }
  return lines.join('\r\n') + '\r\n';
}

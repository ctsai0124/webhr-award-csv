/* ===========================================================
   本機記憶
   記住你在這台電腦上做過的修正，下次遇到相同情況自動套用。

   分兩層存放：
     本機（這個檔案）—— 含人名的：職稱對應、姓名更正、多筆區別
     雲端（memory-cloud.js）—— 不含人名的：事由寫法、獎懲類別

   含人名的資料不上傳，換電腦請用匯出匯入。
   =========================================================== */

const MEMORY_KEY = 'webhr-award-memory';
const MEMORY_VERSION = 1;

const emptyMemory = () => ({
  version: MEMORY_VERSION,
  updated: '',
  // 職稱 → 身分證號。「教導主任」在這所學校指的是誰
  titles: {},
  // 公文寫的字 → 身分證號。錯字、異體字、只寫名的情況
  names: {},
  // 「身分證號@區別詞」→ 事由本文，同一人多筆時各自的寫法
  scopes: {},
});

function loadMemory() {
  try {
    const raw = localStorage.getItem(MEMORY_KEY);
    if (!raw) return emptyMemory();
    const v = JSON.parse(raw);
    if (!v || v.version !== MEMORY_VERSION) return emptyMemory();
    return { ...emptyMemory(), ...v };
  } catch {
    return emptyMemory();
  }
}

function saveMemory(mem) {
  mem.updated = new Date().toISOString();
  try {
    localStorage.setItem(MEMORY_KEY, JSON.stringify(mem));
  } catch {
    // 無痕模式或容量滿了。記憶只是加分功能，失敗不影響主要流程。
  }
  return mem;
}

function memoryCount(mem) {
  return Object.keys(mem.titles).length
    + Object.keys(mem.names).length
    + Object.keys(mem.scopes).length;
}

/* ---------- 記錄修正 ---------- */

/**
 * 記住一筆人員修正。
 * written 是公文上實際寫的字，person 是使用者最後選定的人。
 */
function rememberPerson(mem, written, person) {
  if (!written || !person || !person.id) return mem;
  const key = String(written).trim();
  if (!key || key === person.name) return mem;   // 完全相符不用記

  // 職稱型（教導主任、教務林主任）和姓名型分開放，
  // 因為職稱會隨人事異動改變，姓名不會。
  const bucket = /主任|組長|護理師|幹事|校長|執秘|秘書|書記|管理員|老師|教師/.test(key)
    ? 'titles' : 'names';
  mem[bucket][key] = { id: person.id, name: person.name, at: Date.now() };
  return saveMemory(mem);
}

/** 記住同一人多筆時某個區別詞底下的事由寫法 */
function rememberScope(mem, person, scope, reason) {
  if (!person || !scope || !reason) return mem;
  const main = String(reason).split('\n')[0].trim();
  if (!main) return mem;
  mem.scopes[`${person.id}@${scope}`] = { text: main, at: Date.now() };
  return saveMemory(mem);
}

/* ---------- 套用記憶 ---------- */

/**
 * 用記憶修正一批解析結果。
 * 只在「系統不確定」時介入 —— 姓名完全相符的不動，
 * 否則記憶一旦記錯，會蓋掉本來正確的判斷。
 */
function applyMemory(mem, hits, roster) {
  const findPerson = (rec) => roster.find(p => p.id === rec.id);
  let used = 0;

  for (const h of hits) {
    if (h.confidence === 'exact') continue;
    const rec = mem.titles[h.written] || mem.names[h.written];
    if (!rec) continue;
    const person = findPerson(rec);
    if (!person) continue;              // 已離職，記憶失效

    if (person.id !== h.person.id) {
      h.person = person;
      h.candidates = h.candidates
        ? [person, ...h.candidates.filter(c => c.id !== person.id)]
        : null;
    }
    h.fromMemory = true;
    h.memoryLabel = `上次「${h.written}」對到 ${person.name}`;
    used++;
  }
  return used;
}

/** 記憶裡有這個人這個區別詞的寫法就套用 */
function applyScopeMemory(mem, person, scope) {
  const rec = mem.scopes[`${person.id}@${scope}`];
  return rec ? rec.text : '';
}

/* ---------- 匯出匯入 ---------- */

function exportMemory(mem) {
  return JSON.stringify({ ...mem, exported: new Date().toISOString() }, null, 2);
}

/**
 * 匯入時採合併而非覆蓋，兩台電腦各自累積的記憶都會保留。
 * 同一個 key 衝突時取比較新的那筆。
 */
function importMemory(mem, json) {
  let incoming;
  try {
    incoming = JSON.parse(json);
  } catch {
    throw new Error('檔案格式不正確，請確認是這個工具匯出的記憶檔');
  }
  if (!incoming || incoming.version !== MEMORY_VERSION) {
    throw new Error('記憶檔版本不符，無法匯入');
  }

  let added = 0;
  for (const bucket of ['titles', 'names', 'scopes']) {
    for (const [k, v] of Object.entries(incoming[bucket] || {})) {
      const cur = mem[bucket][k];
      if (!cur || (v.at || 0) > (cur.at || 0)) {
        mem[bucket][k] = v;
        added++;
      }
    }
  }
  saveMemory(mem);
  return added;
}

'use strict';
/* Từ vựng – PWA học từ vựng từ file Word (.docx). Không cần thư viện ngoài, chạy offline. */

const KEY = 'tuvung.v1';
const DAY = 86400000;
const INTERVAL = [0, 1, 2, 4, 8, 16]; // số ngày ôn lại theo hộp Leitner 0..5
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();
const shuffle = a => { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const ANSWER = /^\d+\s*[.)]?\s*\(?[a-e]\)?\.?$/i; // dạng "192 (A)" trong bảng đáp án
const ICON_SPK = '<svg viewBox="0 0 24 24"><path d="M11 5 6 9H3v6h3l5 4z"/><path d="M15.5 8.5a5 5 0 0 1 0 7M18.5 5.5a9 9 0 0 1 0 13"/></svg>';
const ICON_UP = '<svg viewBox="0 0 24 24"><path d="M12 15V3m0 0L7.5 7.5M12 3l4.5 4.5"/><path d="M4 14v4a3 3 0 0 0 3 3h10a3 3 0 0 0 3-3v-4"/></svg>';

/* ---------- Lưu trữ ---------- */
let S = load();
function load() {
  try { const s = JSON.parse(localStorage.getItem(KEY)); if (s && s.words) return s; } catch (e) {}
  return { decks: [], words: [], cur: 'all' };
}
function save() {
  try { localStorage.setItem(KEY, JSON.stringify(S)); }
  catch (e) { toast('Không lưu được dữ liệu (bộ nhớ trình duyệt đầy?)'); }
}
// Dọn các dòng đáp án đã lỡ nhập ở bản cũ
S.words = S.words.filter(w => !ANSWER.test(w.w));
const words = () => S.cur === 'all' ? S.words : S.words.filter(w => w.deck === S.cur);
const due = () => words().filter(w => (w.next || 0) <= Date.now());
const curDeck = () => S.decks.find(d => d.id === S.cur);

let toastT;
function toast(t) { const el = $('#toast'); el.textContent = t; el.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => el.classList.remove('show'), 2200); }

/* ---------- Đọc file .docx (giải nén ZIP bằng DecompressionStream) ---------- */
async function readZipEntry(buf, name) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf), td = new TextDecoder();
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('File không phải .docx hợp lệ.');
  const n = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  for (let k = 0; k < n; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const fnl = dv.getUint16(p + 28, true), exl = dv.getUint16(p + 30, true), cml = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    const fn = td.decode(u8.subarray(p + 46, p + 46 + fnl));
    if (fn === name) {
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      const data = u8.subarray(start, start + csize);
      if (method === 0) return td.decode(data);
      if (method === 8) {
        const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        return td.decode(await new Response(stream).arrayBuffer());
      }
      throw new Error('Kiểu nén không hỗ trợ.');
    }
    p += 46 + fnl + exl + cml;
  }
  throw new Error('Không tìm thấy nội dung văn bản trong file.');
}

const COLS = [ // thứ tự kiểm tra quan trọng: "từ loại" phải xét trước "từ"
  ['skip', /^(stt|so tt|#|no\.?|tt)$/],
  ['t', /^(tu loai|loai|loai tu|type|pos|part of speech)/],
  ['m', /^(nghia|nghia tieng viet|meaning|dich|vietnamese)/],
  ['c', /^(ngu canh|vi du|context|example|cau vi du|cau)/],
  ['w', /^(tu|tu vung|tu moi|word|vocabulary|vocab|english)/],
];
function headerMap(cells) {
  const map = {}; let hits = 0;
  cells.forEach((c, i) => {
    const n = norm(c);
    for (const [k, re] of COLS) if (re.test(n)) { if (k !== 'skip' && map[k] === undefined) map[k] = i; hits++; break; }
  });
  return hits >= 2 && map.w !== undefined ? map : null;
}
function cellText(tc) {
  return [...tc.getElementsByTagNameNS(W, 'p')]
    .map(p => [...p.getElementsByTagNameNS(W, '*')].map(e => e.localName === 't' ? e.textContent : (e.localName === 'tab' || e.localName === 'br') ? ' ' : '').join('').trim())
    .filter(Boolean).join('\n');
}
function parseDocument(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const out = []; let last = null;
  for (const tbl of doc.getElementsByTagNameNS(W, 'tbl')) {
    let map = null;
    for (const tr of tbl.getElementsByTagNameNS(W, 'tr')) {
      if (tr.parentNode !== tbl) continue; // bỏ qua bảng lồng
      const cells = [...tr.childNodes].filter(e => e.localName === 'tc').map(cellText);
      if (!cells.some(Boolean)) continue;
      const h = headerMap(cells);
      if (h) { map = last = h; continue; }
      let m = map || last;
      if (!m) { // không có dòng tiêu đề: đoán Từ | Loại | Nghĩa | Ngữ cảnh (bỏ cột STT nếu có)
        const s = /^\d+\.?$/.test(cells[0]) && cells.length >= 3 ? 1 : 0;
        m = cells.length - s >= 4 ? { w: s, t: s + 1, m: s + 2, c: s + 3 } : { w: s, m: s + 1, c: s + 2 };
      }
      const g = k => m[k] !== undefined ? (cells[m[k]] || '').trim() : '';
      const w = g('w').replace(/\s*\n\s*/g, ' ');
      if (!w || /^\d+\.?$/.test(w) || ANSWER.test(w) || ANSWER.test(g('m'))) continue;
      out.push({ w, t: g('t'), m: g('m'), c: g('c') });
    }
  }
  return out;
}
async function importFile(file) {
  if (!/\.docx$/i.test(file.name)) throw new Error('Chỉ nhận file .docx (file .doc cũ hãy mở bằng Word rồi lưu lại thành .docx).');
  const rows = parseDocument(await readZipEntry(await file.arrayBuffer(), 'word/document.xml'));
  if (!rows.length) throw new Error('Không thấy bảng từ vựng nào trong file.');
  const name = file.name.replace(/\.docx$/i, '');
  let deck = S.decks.find(d => d.name === name);
  if (!deck) { deck = { id: 'd' + Date.now().toString(36), name }; S.decks.push(deck); }
  const have = new Map(S.words.filter(w => w.deck === deck.id).map(w => [norm(w.w), w]));
  let added = 0, updated = 0;
  for (const r of rows) {
    const k = norm(r.w), old = have.get(k);
    if (old) { Object.assign(old, r); updated++; }
    else { const nw = { id: Math.random().toString(36).slice(2, 10), deck: deck.id, ...r, box: 0, next: 0, ok: 0, ng: 0 }; S.words.push(nw); have.set(k, nw); added++; }
  }
  S.cur = deck.id; save();
  return { name, added, updated, total: rows.length };
}
function deleteDeck(id) {
  const d = S.decks.find(x => x.id === id); if (!d) return;
  const n = S.words.filter(w => w.deck === id).length;
  if (!confirm(`Xóa bộ “${d.name}” (${n} từ) và toàn bộ tiến độ học?\nKhông thể hoàn tác.`)) return;
  S.words = S.words.filter(w => w.deck !== id); S.decks = S.decks.filter(x => x.id !== id);
  if (S.cur === id) S.cur = 'all';
  save(); toast(`Đã xóa bộ “${d.name}”`); go(tab);
}

/* ---------- Phát âm ---------- */
function speak(text) {
  if (!('speechSynthesis' in window)) return;
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(text.split('\n')[0]);
  u.lang = 'en-US'; u.rate = 0.9;
  const vs = speechSynthesis.getVoices();
  const v = vs.find(v => /^en[-_]US/i.test(v.lang)) || vs.find(v => /^en/i.test(v.lang));
  if (v) u.voice = v;
  speechSynthesis.speak(u);
}
const spk = w => `<button class="speak" data-say="${esc(w)}" aria-label="Phát âm">${ICON_SPK}</button>`;
function pos(t) {
  if (!t) return '';
  const n = norm(t).replace(/[().]/g, '');
  const c = /^(adv|trang)/.test(n) ? 'adv' : /^(adj|a$|tinh)/.test(n) ? 'adj' : /^(n|noun|danh)/.test(n) ? 'n' : /^(v|verb|dong)/.test(n) ? 'v' : 'x';
  return `<span class="badge pos-${c}">${esc(t)}</span>`;
}
const dots = b => `<span class="dots" title="Hộp ${b || 0}/5">${[1, 2, 3, 4, 5].map(i => `<i class="${i <= (b || 0) ? 'f' : ''}"></i>`).join('')}</span>`;

/* ---------- Điều hướng ---------- */
let tab = 'learn';
function renderDecks() {
  const sel = $('#deck');
  if (S.cur !== 'all' && !curDeck()) S.cur = 'all';
  sel.innerHTML = [`<option value="all">Tất cả bộ từ · ${S.words.length}</option>`]
    .concat(S.decks.map(d => `<option value="${d.id}">${esc(d.name)} · ${S.words.filter(w => w.deck === d.id).length}</option>`)).join('');
  sel.value = S.cur;
}
function go(t) {
  tab = t; session = null;
  $$('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  renderDecks();
  ({ learn: viewHome, quiz: viewQuiz, list: viewList, import: viewImport })[t]();
  $('#view').classList.remove('fade'); void $('#view').offsetWidth; $('#view').classList.add('fade');
  window.scrollTo(0, 0);
}
function empty() {
  $('#view').innerHTML = `<div class="card empty stack"><div class="big">📚</div><h3 style="margin:8px 0 0">Chưa có từ nào</h3>
  <p class="muted" style="margin:6px 0 0">Nhập file Word có bảng các cột <b>Từ / Loại / Nghĩa / Ngữ cảnh</b> để bắt đầu.</p>
  <button class="btn pri" onclick="go('import')">${ICON_UP} Nhập file .docx</button></div>`;
}

/* ---------- Tổng quan ---------- */
function viewHome() {
  const ws = words();
  if (!ws.length) return empty();
  const d = due().length, learned = ws.filter(w => (w.box || 0) >= 3).length;
  const pct = Math.round(learned / ws.length * 100), C = 2 * Math.PI * 40;
  const boxes = [0, 0, 0, 0, 0, 0]; ws.forEach(w => boxes[w.box || 0]++);
  const mx = Math.max(...boxes, 1);
  const ok = ws.reduce((a, w) => a + (w.ok || 0), 0), ng = ws.reduce((a, w) => a + (w.ng || 0), 0);
  const nxt = d ? 0 : Math.min(...ws.map(w => w.next || 0));
  $('#view').innerHTML = `
  <div class="hero">
    <div class="row">
      <div class="grow">
        <h2>Cần ôn hôm nay</h2>
        <div class="num">${d}</div>
        <div class="sub">${d ? 'từ đang chờ bạn' : 'Lượt tiếp: ' + new Date(nxt).toLocaleString('vi-VN', { hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit' })}</div>
      </div>
      <svg class="ring" viewBox="0 0 96 96"><circle class="bg" cx="48" cy="48" r="40"/><circle class="fg" cx="48" cy="48" r="40" stroke-dasharray="${C}" stroke-dashoffset="${C * (1 - pct / 100)}"/><text x="48" y="48">${pct}%</text></svg>
    </div>
    <div class="actions">
      ${d ? `<button class="btn white grow" id="start">Bắt đầu học ${Math.min(d, 20)} từ</button>` : `<button class="btn white grow" id="rand">Ôn ngẫu nhiên 20 từ</button>`}
      ${d ? `<button class="btn ghost" id="rand">Ngẫu nhiên</button>` : ''}
    </div>
  </div>
  <div class="stats">
    <div class="stat"><b>${ws.length}</b><span>Tổng số từ</span></div>
    <div class="stat"><b>${learned}</b><span>Đã thuộc</span></div>
    <div class="stat"><b>${ok + ng ? Math.round(ok / (ok + ng) * 100) + '%' : '–'}</b><span>Tỉ lệ đúng</span></div>
  </div>
  <div class="title">Mức độ ghi nhớ</div>
  <div class="card">
    <div class="boxes">${boxes.map((n, i) => `<div><small>${n}</small><i style="height:${Math.max(4, n / mx * 64)}px;opacity:${.35 + i * .13}"></i><small>${i === 0 ? 'Mới' : 'Hộp ' + i}</small></div>`).join('')}</div>
    <p class="muted small" style="margin:12px 0 0">Mỗi lần nhớ đúng, từ lên một hộp và được hẹn ôn lại xa hơn (1, 2, 4, 8, 16 ngày). Từ ở hộp 3 trở lên được tính là đã thuộc.</p>
  </div>`;
  if ($('#start')) $('#start').onclick = () => startSession(false);
  $('#rand').onclick = () => startSession(true);
}

/* ---------- Flashcard + lặp lại ngắt quãng ---------- */
let session = null, si = 0, flipped = false, sessOk = 0;
function startSession(rand) {
  session = shuffle((rand ? words() : due()).slice()).slice(0, 20); si = 0; flipped = false; sessOk = 0;
  drawCard();
}
function drawCard() {
  if (si >= session.length) {
    const n = session.length; session = null;
    $('#view').innerHTML = `<div class="card empty stack fade"><div class="big">🎉</div><h3 style="margin:8px 0 0">Hoàn thành!</h3>
    <p class="muted" style="margin:6px 0 0">Bạn đã ôn ${n} thẻ, nhớ ngay ${sessOk} thẻ.</p>
    <button class="btn pri block" onclick="go('learn')">Về tổng quan</button></div>`;
    return;
  }
  const w = session[si];
  $('#view').innerHTML = `
  <div class="sessbar"><button class="iconbtn" id="quit" aria-label="Thoát">✕</button><div class="prog"><i style="width:${si / session.length * 100}%"></i></div><span class="muted small" style="font-weight:700">${si + 1}/${session.length}</span></div>
  <div class="scene"><div class="flip${flipped ? ' on' : ''}" id="fc">
    <div class="face front">
      ${pos(w.t)}
      <div class="word" style="margin-top:12px">${esc(w.w)}</div>
      <div style="margin-top:14px">${spk(w.w)}</div>
      <div class="hint">Chạm để lật thẻ</div>
    </div>
    <div class="face back">
      <div class="muted small" style="font-weight:700">${esc(w.w)} ${w.t ? '· ' + esc(w.t) : ''}</div>
      <div class="mean">${esc(w.m) || '<span class="muted">(chưa có nghĩa)</span>'}</div>
      ${w.c ? `<div class="ctx">${esc(w.c)}</div>` : ''}
    </div>
  </div></div>
  <div class="grade" id="gr" style="visibility:${flipped ? 'visible' : 'hidden'}">
    <button class="btn ng" id="no">Chưa nhớ</button><button class="btn ok" id="yes">Đã nhớ</button>
  </div>`;
  const fc = $('#fc');
  fc.onclick = e => { if (e.target.closest('.speak')) return; flipped = !flipped; fc.classList.toggle('on', flipped); $('#gr').style.visibility = 'visible'; };
  $('#yes').onclick = () => grade(w, true);
  $('#no').onclick = () => grade(w, false);
  $('#quit').onclick = () => go('learn');
}
function grade(w, ok) {
  if (ok) { w.box = Math.min((w.box || 0) + 1, 5); w.ok = (w.ok || 0) + 1; sessOk++; }
  else { w.box = 0; w.ng = (w.ng || 0) + 1; session.push(w); }
  w.next = Date.now() + INTERVAL[w.box] * DAY;
  save(); si++; flipped = false; drawCard();
}
document.addEventListener('keydown', e => {
  if (!session || e.target.matches('input,select')) return;
  if (e.code === 'Space') { e.preventDefault(); $('#fc')?.click(); }
  else if (flipped && e.key === '1') $('#no')?.click();
  else if (flipped && e.key === '2') $('#yes')?.click();
});

/* ---------- Kiểm tra trắc nghiệm ---------- */
let qdir = 'wm', score = { ok: 0, n: 0, streak: 0 };
function viewQuiz() {
  const ws = words().filter(w => w.m);
  if (ws.length < 4) {
    if (!words().length) return empty();
    $('#view').innerHTML = `<div class="card empty"><div class="big">🧩</div><p class="muted">Cần ít nhất 4 từ có nghĩa để làm bài kiểm tra.</p></div>`; return;
  }
  const w = ws[Math.floor(Math.random() * ws.length)];
  const key = qdir === 'wm' ? 'm' : 'w';
  const seen = new Set([norm(w[key])]);
  const pool = shuffle(ws.slice()).filter(x => { const k = norm(x[key]); if (x === w || seen.has(k)) return false; seen.add(k); return true; }).slice(0, 3);
  const opts = shuffle([w, ...pool]);
  $('#view').innerHTML = `
  <div class="row" style="margin-bottom:14px">
    <div class="seg grow"><button data-d="wm" class="${qdir === 'wm' ? 'on' : ''}">Từ → Nghĩa</button><button data-d="mw" class="${qdir === 'mw' ? 'on' : ''}">Nghĩa → Từ</button></div>
  </div>
  <div class="row" style="margin-bottom:14px">
    <span class="pill">✅ ${score.ok}/${score.n}</span><span class="pill">🔥 ${score.streak}</span><span class="grow"></span>
  </div>
  <div class="card q">
    ${qdir === 'wm'
      ? `${pos(w.t)}<div class="word" style="margin-top:10px">${esc(w.w)}</div><div style="margin-top:12px">${spk(w.w)}</div>`
      : `${pos(w.t)}<div class="mean" style="margin-top:10px">${esc(w.m)}</div>`}
  </div>
  <div id="opts">${opts.map((o, i) => `<button class="opt" data-i="${i}"><span class="k">${'ABCD'[i]}</span><span>${esc(o[key])}</span></button>`).join('')}</div>
  <div id="after"></div>`;
  $$('.seg button').forEach(b => b.onclick = () => { qdir = b.dataset.d; viewQuiz(); });
  $$('.opt').forEach(b => b.onclick = () => {
    const right = opts[+b.dataset.i] === w;
    $$('.opt').forEach(x => { x.disabled = true; if (opts[+x.dataset.i] === w) x.classList.add('right'); });
    if (!right) b.classList.add('wrong', 'shake');
    score.n++;
    if (right) { score.ok++; score.streak++; w.ok = (w.ok || 0) + 1; }
    else { score.streak = 0; w.ng = (w.ng || 0) + 1; w.box = 0; w.next = 0; }
    save(); if (qdir === 'mw' || !right) speak(w.w);
    $('#after').innerHTML = `${w.c ? `<div class="ctx fade">${esc(w.c)}</div>` : ''}<button class="btn pri block" style="margin-top:14px" id="nq">Câu tiếp theo →</button>`;
    $('#nq').onclick = viewQuiz;
  });
}

/* ---------- Danh sách ---------- */
function viewList() {
  if (!words().length) return empty();
  const deck = curDeck(), ws0 = words();
  const learned = ws0.filter(w => (w.box || 0) >= 3).length;
  $('#view').innerHTML = `
  <div class="card deckhead" style="margin-bottom:12px">
    <div class="ic">📘</div>
    <div class="grow"><b>${deck ? esc(deck.name) : 'Tất cả bộ từ'}</b><span class="muted small">${ws0.length} từ · đã thuộc ${learned}</span></div>
    ${deck ? `<button class="btn danger" id="del">Xóa bộ</button>` : ''}
  </div>
  <div class="search"><input type="search" id="q" placeholder="Tìm từ hoặc nghĩa…"></div>
  <div class="card" id="ls"></div>`;
  const draw = () => {
    const q = norm($('#q').value);
    const ws = ws0.filter(w => !q || norm(w.w + ' ' + w.m).includes(q));
    $('#ls').innerHTML = ws.length ? ws.slice(0, 400).map(w => `<div class="item"><div class="grow">
        <div><span class="w">${esc(w.w)}</span> ${pos(w.t)}${dots(w.box)}</div>
        <div class="m">${esc(w.m)}</div>${w.c ? `<div class="c">${esc(w.c)}</div>` : ''}</div>${spk(w.w)}</div>`).join('')
      + (ws.length > 400 ? `<p class="muted small" style="text-align:center">… còn ${ws.length - 400} từ, hãy gõ để tìm.</p>` : '')
      : '<p class="muted" style="text-align:center;margin:8px 0">Không tìm thấy.</p>';
  };
  $('#q').oninput = draw; draw();
  if (deck) $('#del').onclick = () => deleteDeck(deck.id);
}

/* ---------- Bộ từ: nhập file, quản lý, sao lưu ---------- */
function viewImport(lines) {
  $('#view').innerHTML = `
  <label class="drop" id="drop">
    <div class="ic">${ICON_UP}</div>
    <b>Chọn file Word (.docx)</b>
    <div class="muted small" style="margin-top:4px">Bảng có các cột Từ / Loại / Nghĩa / Ngữ cảnh.<br>Nhập lại cùng tên file sẽ cập nhật và giữ tiến độ.</div>
    <input type="file" id="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple hidden>
  </label>
  <div class="msg" id="msg">${(lines || []).map(l => `<div class="${l.ok ? 'good' : 'bad'}">${esc(l.t)}</div>`).join('')}</div>
  <div class="title">Bộ từ của bạn</div>
  <div class="card decks">${S.decks.length ? S.decks.map(d => {
    const ws = S.words.filter(w => w.deck === d.id), l = ws.filter(w => (w.box || 0) >= 3).length;
    return `<div class="d"><div class="deckhead grow"><div class="ic">📘</div><div class="grow"><b>${esc(d.name)}</b><span class="muted small">${ws.length} từ · đã thuộc ${l}</span></div></div>
      <button class="btn danger" data-del="${d.id}">Xóa</button></div>`;
  }).join('') : '<p class="muted" style="margin:0;text-align:center">Chưa có bộ từ nào.</p>'}</div>
  <div class="title">Sao lưu tiến độ</div>
  <div class="card stack">
    <p class="muted small" style="margin:0">Dữ liệu chỉ lưu trên trình duyệt của máy này. Hãy xuất file sao lưu để chuyển sang máy khác hoặc phòng khi xóa dữ liệu trình duyệt.</p>
    <div class="row"><button class="btn grow" id="exp">Xuất sao lưu</button><button class="btn grow" id="impb">Khôi phục</button></div>
    <input type="file" id="bk" accept=".json,application/json" hidden>
  </div>`;
  const handle = async files => {
    $('#msg').innerHTML = '<div>Đang đọc…</div>';
    const out = [];
    for (const f of files) {
      try { const r = await importFile(f); out.push({ ok: 1, t: `✓ ${r.name}: ${r.total} từ (mới ${r.added}, cập nhật ${r.updated})` }); }
      catch (e) { out.push({ ok: 0, t: `✕ ${f.name}: ${e.message}` }); }
    }
    renderDecks(); viewImport(out);
  };
  $('#file').onchange = e => handle([...e.target.files]);
  const drop = $('#drop');
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); handle([...e.dataTransfer.files]); };
  $$('[data-del]').forEach(b => b.onclick = () => deleteDeck(b.dataset.del));
  $('#exp').onclick = () => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(S)], { type: 'application/json' }));
    a.download = `tuvung-saoluu-${new Date().toISOString().slice(0, 10)}.json`; a.click();
    toast('Đã xuất file sao lưu');
  };
  $('#impb').onclick = () => $('#bk').click();
  $('#bk').onchange = async e => {
    try { const s = JSON.parse(await e.target.files[0].text()); if (!s.words) throw 0; if (confirm('Thay toàn bộ dữ liệu hiện tại bằng bản sao lưu?')) { S = s; save(); toast('Đã khôi phục'); go('learn'); } }
    catch (_) { toast('File sao lưu không hợp lệ'); }
  };
}

/* ---------- Khởi động ---------- */
$$('.tabs button').forEach(b => b.onclick = () => go(b.dataset.tab));
$('#deck').onchange = e => { S.cur = e.target.value; save(); go(tab); };
document.addEventListener('click', e => { const b = e.target.closest('.speak'); if (b) { e.stopPropagation(); speak(b.dataset.say); } }, true);
if ('speechSynthesis' in window) speechSynthesis.getVoices();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
save();
go(S.words.length ? 'learn' : 'import');

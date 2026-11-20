'use strict';
/* Từ vựng – PWA học từ vựng từ file Word (.docx). Không cần thư viện ngoài, chạy offline. */

const KEY = 'tuvung.v1';
const DAY = 86400000;
const INTERVAL = [0, 1, 2, 4, 8, 16]; // số ngày ôn lại theo hộp Leitner 0..5
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();
const shuffle = a => { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };

/* ---------- Lưu trữ ---------- */
let S = load();
function load() {
  try { const s = JSON.parse(localStorage.getItem(KEY)); if (s && s.words) return s; } catch (e) {}
  return { decks: [], words: [], cur: 'all' };
}
function save() {
  try { localStorage.setItem(KEY, JSON.stringify(S)); }
  catch (e) { alert('Không lưu được dữ liệu (bộ nhớ trình duyệt đầy?)'); }
}
const words = () => S.cur === 'all' ? S.words : S.words.filter(w => w.deck === S.cur);
const due = () => words().filter(w => (w.next || 0) <= Date.now());

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
      if (!w || /^\d+\.?$/.test(w)) continue;
      out.push({ w, t: g('t'), m: g('m'), c: g('c') });
    }
  }
  return out;
}
async function importFile(file) {
  if (!/\.docx$/i.test(file.name)) throw new Error('Chỉ nhận file Word .docx (file .doc cũ hãy mở bằng Word rồi lưu lại thành .docx).');
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

/* ---------- Phát âm ---------- */
function speak(text) {
  if (!('speechSynthesis' in window)) return;
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(text.split('\n')[0]);
  u.lang = 'en-US'; u.rate = 0.9;
  const v = speechSynthesis.getVoices().find(v => /^en[-_]US/i.test(v.lang)) || speechSynthesis.getVoices().find(v => /^en/i.test(v.lang));
  if (v) u.voice = v;
  speechSynthesis.speak(u);
}
const spk = w => `<button class="speak" data-say="${esc(w)}" aria-label="Phát âm">🔊</button>`;

/* ---------- Giao diện ---------- */
let tab = 'learn';
function renderDecks() {
  const sel = $('#deck');
  const opts = [`<option value="all">Tất cả (${S.words.length})</option>`]
    .concat(S.decks.map(d => `<option value="${d.id}">${esc(d.name)} (${S.words.filter(w => w.deck === d.id).length})</option>`));
  sel.innerHTML = opts.join('');
  if (S.cur !== 'all' && !S.decks.some(d => d.id === S.cur)) S.cur = 'all';
  sel.value = S.cur;
}
function go(t) {
  tab = t;
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  renderDecks();
  ({ learn: viewLearn, quiz: viewQuiz, list: viewList, import: viewImport })[t]();
  window.scrollTo(0, 0);
}
function empty() {
  $('#view').innerHTML = `<div class="card stack"><p>Chưa có từ nào. Hãy nhập file Word có bảng từ vựng (các cột <b>Từ / Loại / Nghĩa / Ngữ cảnh</b>).</p><button class="btn pri" onclick="go('import')">Nhập file .docx</button></div>`;
}

/* Học bằng flashcard + lặp lại ngắt quãng */
let session = [], si = 0, flipped = false;
function viewLearn(all) {
  if (!words().length) return empty();
  const d = due();
  session = shuffle((all ? words().slice() : d).slice()).slice(0, 20); si = 0; flipped = false;
  if (!session.length) {
    const nxt = Math.min(...words().map(w => w.next || 0));
    $('#view').innerHTML = `<div class="card stack"><p>🎉 Đã ôn hết các từ đến hạn.</p><p class="muted">Lượt ôn tiếp theo: ${new Date(nxt).toLocaleString('vi-VN')}</p><button class="btn pri" id="all">Ôn thêm ngẫu nhiên 20 từ</button></div>${stats()}`;
    $('#all').onclick = () => viewLearn(true);
    return;
  }
  drawCard();
}
function stats() {
  const ws = words(), boxes = [0, 0, 0, 0, 0, 0];
  ws.forEach(w => boxes[w.box || 0]++);
  const learned = ws.filter(w => (w.box || 0) >= 3).length;
  return `<div class="card stack" style="margin-top:12px"><div class="row"><span class="grow">Đã thuộc (hộp ≥ 3)</span><b>${learned}/${ws.length}</b></div>
  <div class="bar"><i style="width:${ws.length ? learned / ws.length * 100 : 0}%"></i></div>
  <div class="muted" style="font-size:14px">Hộp 0→5: ${boxes.join(' · ')} &nbsp;|&nbsp; Đến hạn hôm nay: ${due().length}</div></div>`;
}
function drawCard() {
  if (si >= session.length) { $('#view').innerHTML = `<div class="card stack"><p>✅ Xong lượt ${session.length} từ.</p><button class="btn pri" onclick="viewLearn()">Học tiếp</button></div>${stats()}`; return; }
  const w = session[si];
  $('#view').innerHTML = `
  <div class="muted" style="margin-bottom:8px">Thẻ ${si + 1}/${session.length} · hộp ${w.box || 0}</div>
  <div class="card flash" id="fc">
    <div class="word">${esc(w.w)} ${spk(w.w)}</div>
    ${w.t ? `<div class="type">${esc(w.t)}</div>` : ''}
    ${flipped ? `<div class="mean">${esc(w.m)}</div>${w.c ? `<div class="ctx">${esc(w.c)}</div>` : ''}` : `<div class="muted" style="margin-top:18px">Chạm để xem nghĩa</div>`}
  </div>
  <div class="row" style="margin-top:12px">
    ${flipped ? `<button class="btn ng grow" id="no">✗ Chưa nhớ</button><button class="btn ok grow" id="yes">✓ Đã nhớ</button>` : `<button class="btn pri grow" id="flip">Lật thẻ</button>`}
  </div>`;
  const flip = () => { flipped = true; drawCard(); };
  $('#fc').onclick = e => { if (!e.target.closest('.speak') && !flipped) flip(); };
  if (!flipped) $('#flip').onclick = flip;
  else { $('#yes').onclick = () => grade(w, true); $('#no').onclick = () => grade(w, false); }
}
function grade(w, ok) {
  if (ok) { w.box = Math.min((w.box || 0) + 1, 5); w.ok = (w.ok || 0) + 1; }
  else { w.box = 0; w.ng = (w.ng || 0) + 1; session.push(w); }
  w.next = Date.now() + INTERVAL[w.box] * DAY;
  save(); si++; flipped = false; drawCard();
}

/* Kiểm tra trắc nghiệm */
let qdir = 'wm', score = { ok: 0, n: 0 };
function viewQuiz() {
  const ws = words().filter(w => w.m);
  if (ws.length < 4) { if (!words().length) return empty(); $('#view').innerHTML = `<div class="card">Cần ít nhất 4 từ có nghĩa để làm bài kiểm tra.</div>`; return; }
  const w = ws[Math.floor(Math.random() * ws.length)];
  const key = qdir === 'wm' ? 'm' : 'w';
  const pool = shuffle(ws.filter(x => x !== w && norm(x[key]) !== norm(w[key]))).slice(0, 3);
  const opts = shuffle([w, ...pool]);
  $('#view').innerHTML = `
  <div class="row" style="margin-bottom:10px"><select id="qd" class="grow"><option value="wm">Từ → Nghĩa</option><option value="mw">Nghĩa → Từ</option></select>
  <span class="muted">Đúng ${score.ok}/${score.n}</span></div>
  <div class="card">
    ${qdir === 'wm' ? `<div class="word" style="font-size:26px">${esc(w.w)} ${spk(w.w)}</div>${w.t ? `<div class="type">${esc(w.t)}</div>` : ''}`
      : `<div class="mean" style="margin:0">${esc(w.m)}</div>${w.t ? `<div class="type">${esc(w.t)}</div>` : ''}`}
    <div id="opts">${opts.map((o, i) => `<button class="btn opt" data-i="${i}">${esc(o[key])}</button>`).join('')}</div>
    <div id="after"></div>
  </div>`;
  $('#qd').value = qdir;
  $('#qd').onchange = e => { qdir = e.target.value; viewQuiz(); };
  document.querySelectorAll('.opt').forEach(b => b.onclick = () => {
    const pick = opts[+b.dataset.i], right = pick === w;
    document.querySelectorAll('.opt').forEach(x => { x.disabled = true; if (opts[+x.dataset.i] === w) x.classList.add('right'); });
    if (!right) b.classList.add('wrong');
    score.n++; if (right) { score.ok++; w.ok = (w.ok || 0) + 1; } else { w.ng = (w.ng || 0) + 1; w.box = 0; w.next = 0; }
    save();
    if (qdir === 'mw') speak(w.w);
    $('#after').innerHTML = `${w.c ? `<div class="ctx">${esc(w.c)}</div>` : ''}<button class="btn pri" style="width:100%;margin-top:12px" id="nq">Câu tiếp →</button>`;
    $('#nq').onclick = viewQuiz;
  });
}

/* Danh sách + tìm kiếm */
function viewList() {
  if (!words().length) return empty();
  const deck = S.decks.find(d => d.id === S.cur);
  $('#view').innerHTML = `
  <input type="search" id="q" placeholder="Tìm từ hoặc nghĩa…" style="width:100%">
  <div class="card list" id="ls" style="margin-top:12px"></div>
  ${deck ? `<button class="btn" style="margin-top:12px;width:100%;color:var(--ng)" id="del">Xóa bộ “${esc(deck.name)}”</button>` : ''}`;
  const draw = () => {
    const q = norm($('#q').value);
    const ws = words().filter(w => !q || norm(w.w + ' ' + w.m).includes(q));
    $('#ls').innerHTML = ws.slice(0, 500).map(w => `<div class="it"><b>${esc(w.w)}</b> ${w.t ? `<i class="muted">(${esc(w.t)})</i>` : ''} ${spk(w.w)}<span class="box">hộp ${w.box || 0}</span>
      <div style="white-space:pre-line">${esc(w.m)}</div>${w.c ? `<div class="ctx" style="margin-top:4px">${esc(w.c)}</div>` : ''}</div>`).join('')
      + (ws.length > 500 ? `<p class="muted">… và ${ws.length - 500} từ nữa, hãy gõ để tìm.</p>` : '') || '<p class="muted">Không tìm thấy.</p>';
  };
  $('#q').oninput = draw; draw();
  if (deck) $('#del').onclick = () => {
    if (!confirm(`Xóa bộ “${deck.name}” và toàn bộ tiến độ học?`)) return;
    S.words = S.words.filter(w => w.deck !== deck.id); S.decks = S.decks.filter(d => d !== deck); S.cur = 'all'; save(); go('list');
  };
}

/* Nhập file + sao lưu */
function viewImport() {
  $('#view').innerHTML = `
  <div class="card stack">
    <label class="drop" id="drop" style="display:block;cursor:pointer">
      <div style="font-size:36px">📄</div>
      <b>Chọn file Word (.docx)</b>
      <div class="muted" style="font-size:14px">Bảng có các cột Từ / Loại / Nghĩa / Ngữ cảnh. Nhập lại cùng tên file sẽ cập nhật và giữ tiến độ.</div>
      <input type="file" id="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple hidden>
    </label>
    <div id="msg" class="msg"></div>
  </div>
  <div class="card stack" style="margin-top:12px">
    <b>Sao lưu tiến độ</b>
    <div class="row"><button class="btn grow" id="exp">Xuất file sao lưu</button><button class="btn grow" id="impb">Khôi phục</button></div>
    <input type="file" id="bk" accept=".json,application/json" hidden>
  </div>`;
  const handle = async files => {
    const msg = $('#msg'); msg.textContent = 'Đang đọc…';
    const lines = [];
    for (const f of files) {
      try { const r = await importFile(f); lines.push(`✅ ${r.name}: ${r.total} từ (mới ${r.added}, cập nhật ${r.updated})`); }
      catch (e) { lines.push(`❌ ${f.name}: ${e.message}`); }
    }
    msg.textContent = lines.join('\n'); renderDecks();
  };
  $('#file').onchange = e => handle([...e.target.files]);
  const drop = $('#drop');
  drop.ondragover = e => { e.preventDefault(); };
  drop.ondrop = e => { e.preventDefault(); handle([...e.dataTransfer.files]); };
  $('#exp').onclick = () => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(S)], { type: 'application/json' }));
    a.download = `tuvung-saoluu-${new Date().toISOString().slice(0, 10)}.json`; a.click();
  };
  $('#impb').onclick = () => $('#bk').click();
  $('#bk').onchange = async e => {
    try { const s = JSON.parse(await e.target.files[0].text()); if (!s.words) throw 0; if (confirm('Thay toàn bộ dữ liệu hiện tại bằng bản sao lưu?')) { S = s; save(); go('learn'); } }
    catch (_) { alert('File sao lưu không hợp lệ.'); }
  };
}

/* ---------- Khởi động ---------- */
document.querySelectorAll('nav button').forEach(b => b.onclick = () => go(b.dataset.tab));
$('#deck').onchange = e => { S.cur = e.target.value; save(); go(tab); };
document.addEventListener('click', e => { const b = e.target.closest('.speak'); if (b) { e.stopPropagation(); speak(b.dataset.say); } }, true);
if ('speechSynthesis' in window) speechSynthesis.getVoices();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
go(S.words.length ? 'learn' : 'import');

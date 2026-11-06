"use strict";
/* Коллекция банок — данные хранятся в браузере (IndexedDB) на устройстве пользователя. */

// SheetJS 0.18.5 (Apache-2.0), loaded only for Excel import/export: local copy first, then public CDNs
const XLSX_URLS = [
  "xlsx.full.min.js",
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
];

const FIELDS = [
  { key: "name",    label: "Название", q: "Как называется напиток?", type: "text", def: "coca-cola" },
  { key: "country", label: "Страна",   q: "Из какой страны банка?", hint: "Обычно указано мелким шрифтом рядом с составом или адресом производителя.", type: "text", required: true },
  { key: "flavor",  label: "Вкус",     q: "Какой вкус?", type: "text", required: true },
  { key: "event",   label: "Событие",  q: "Банка посвящена какому-то событию?", hint: "Чемпионат, фильм, праздник. Если нет, выберите «нет».", type: "text", def: "нет" },
  { key: "volume",  label: "Объём, л", q: "Какой объём банки?", type: "text", def: "0.33", fixed: ["0.15", "0.25", "0.33", "0.355", "0.5"] },
  { key: "height",  label: "Высота",   q: "Банка высокая или низкая?", hint: "Высокая — узкая и вытянутая, низкая — классическая широкая.", type: "choice", fixed: ["выс", "низ"], required: true },
  { key: "date",    label: "Год",      q: "Какая дата на банке?", hint: "Дата с донышка банки. Можно пропустить.", type: "date" },
  { key: "notes",   label: "Примечания", q: "Что ещё стоит запомнить?", hint: "Где купили, особенности дизайна. Можно пропустить.", type: "textarea" },
];
const HEIGHT_LABEL = { "выс": "высокая", "низ": "низкая" };
const ABBR = { "сша": "США", "оаэ": "ОАЭ", "юар": "ЮАР", "кндр": "КНДР" };

/* ---------- helpers ---------- */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const cap = (s) => String(s ?? "").trim();
const norm = (v) => cap(v).toLowerCase();
const capFirst = (s) => (s ? s[0].toUpperCase() + s.slice(1) : s);
const dispCountry = (v) => ABBR[norm(v)] || capFirst(cap(v));
const dispHeight = (v) => HEIGHT_LABEL[norm(v)] || cap(v);
const hasEvent = (c) => !!norm(c.event) && norm(c.event) !== "нет";
function fmtDate(d) {
  if (!d) return "—";
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(d);
  return m ? `${m[3]}.${m[2]}.${m[1]}` : d;
}
function plural(n, a, b, c) {
  const m10 = n % 10, m100 = n % 100;
  return n + " " + (m10 === 1 && m100 !== 11 ? a : m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? b : c);
}
function toast(msg, ms = 2600) {
  const t = $("toast"); t.textContent = msg; t.hidden = false;
  clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), ms);
}
const uid = () => "can-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

/* ---------- storage (IndexedDB) ---------- */
const DB_NAME = "can-collection", STORE = "cans";
let dbPromise = null;
function openDB() {
  if (!dbPromise) dbPromise = new Promise((res, rej) => {
    if (!("indexedDB" in window)) return rej(new Error("no-idb"));
    const r = indexedDB.open(DB_NAME, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: "id" });
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbPromise;
}
async function run(mode, fn) {
  if (run.disabled) return undefined; // storage unavailable: keep data in memory only
  const db = await openDB();
  return new Promise((res, rej) => {
    const t = db.transaction(STORE, mode), s = t.objectStore(STORE);
    const req = fn(s);
    t.oncomplete = () => res(req && "result" in req ? req.result : undefined);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error || new Error("abort"));
  });
}
const dbAll = () => run("readonly", (s) => s.getAll());
const dbPut = (c) => run("readwrite", (s) => s.put(c));
const dbPutMany = (list) => run("readwrite", (s) => { list.forEach((c) => s.put(c)); return null; });
const dbDelete = (id) => run("readwrite", (s) => s.delete(id));

let cans = [];
let lastShown = [], lastFiltering = false; // what the list currently shows, used by the Excel export

/* object URLs for stored photos */
const urlCache = new Map();
function blobURL(key, blob) {
  if (!blob) return "";
  const e = urlCache.get(key);
  if (e && e.blob === blob) return e.url;
  if (e) URL.revokeObjectURL(e.url);
  const url = URL.createObjectURL(blob);
  urlCache.set(key, { blob, url });
  return url;
}
const thumbSrc = (c) => blobURL(c.id + ":t", c.thumb || c.photo);
const photoSrc = (c) => blobURL(c.id + ":p", c.photo || c.thumb);

/* ---------- list ---------- */
function fillSelect(id, key, all) {
  const el = $(id), cur = el.value, counts = {};
  cans.forEach((c) => { const v = norm(c[key]); if (v) counts[v] = (counts[v] || 0) + 1; });
  const vals = Object.keys(counts).sort((a, b) => a.localeCompare(b, "ru"));
  el.innerHTML = `<option value="">${all}</option>` +
    vals.map((v) => `<option value="${esc(v)}">${esc(key === "country" ? dispCountry(v) : capFirst(v))} (${counts[v]})</option>`).join("");
  el.value = vals.includes(cur) ? cur : "";
}
const SORTS = {
  "num-desc": (a, b) => (b.num || 0) - (a.num || 0),
  "num-asc": (a, b) => (a.num || 0) - (b.num || 0),
  "country": (a, b) => norm(a.country).localeCompare(norm(b.country), "ru") || (a.num || 0) - (b.num || 0),
  "date-desc": (a, b) => cap(b.date).localeCompare(cap(a.date)) || (b.num || 0) - (a.num || 0),
};
function render() {
  fillSelect("fCountry", "country", "Все страны");
  fillSelect("fFlavor", "flavor", "Все вкусы");
  const q = norm($("search").value);
  const F = { country: $("fCountry").value, flavor: $("fFlavor").value, event: $("fEvent").value, height: $("fHeight").value };
  ["fCountry", "fFlavor", "fEvent", "fHeight"].forEach((id) => $(id).classList.toggle("on", !!$(id).value));
  const filtering = !!(q || F.country || F.flavor || F.event || F.height);
  $("fReset").hidden = !filtering;

  const shown = cans.filter((c) =>
    (!q || [c.country, dispCountry(c.country), c.flavor, c.event, c.notes, c.name].join(" ").toLowerCase().includes(q)) &&
    (!F.country || norm(c.country) === F.country) &&
    (!F.flavor || norm(c.flavor) === F.flavor) &&
    (!F.event || (F.event === "yes") === hasEvent(c)) &&
    (!F.height || norm(c.height) === F.height)
  ).sort(SORTS[$("fSort").value] || SORTS["num-desc"]);
  lastShown = shown; lastFiltering = filtering;
  $("exportBtn").hidden = cans.length === 0;
  $("exportBtn").title = filtering ? `Выгрузить в Excel показанные банки (${shown.length})` : "Выгрузить весь список в Excel";

  const countries = new Set(cans.map((c) => norm(c.country)).filter(Boolean));
  $("count").textContent = `${plural(cans.length, "банка", "банки", "банок")} · ${plural(countries.size, "страна", "страны", "стран")}` +
    (filtering ? ` · показано ${shown.length}` : "");
  $("empty").hidden = cans.length > 0;
  $("filters").hidden = cans.length === 0;
  $("search").hidden = cans.length === 0;

  $("list").innerHTML = shown.map((c) => {
    const src = thumbSrc(c);
    const img = src
      ? `<img class="thumb" src="${src}" alt="" loading="lazy">`
      : `<span class="thumb" aria-hidden="true"><svg width="26" height="34" viewBox="0 0 26 34" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3" y="3" width="20" height="28" rx="4"/><path d="M3 8h20M3 26h20"/></svg></span>`;
    const ev = hasEvent(c) ? `<span class="tag">${esc(c.event)}</span>` : "";
    return `<li><button class="item" type="button" data-id="${esc(c.id)}">${img}
      <span class="txt"><div class="country">${esc(dispCountry(c.country) || "Без страны")}</div><div class="flavor">${esc(cap(c.flavor) || "вкус не указан")}</div>${ev}</span>
      <span class="num">№${esc(c.num)}</span></button></li>`;
  }).join("");
  if (cans.length && !shown.length) $("list").innerHTML = `<li class="empty">Ничего не нашлось. Измените фильтры или нажмите «Сбросить».</li>`;
}
$("search").addEventListener("input", render);
["fCountry", "fFlavor", "fEvent", "fHeight", "fSort"].forEach((id) => $(id).addEventListener("change", render));
$("fReset").onclick = () => { ["fCountry", "fFlavor", "fEvent", "fHeight", "search"].forEach((id) => ($(id).value = "")); render(); };
$("list").addEventListener("click", (e) => { const b = e.target.closest(".item"); if (b) openDetail(b.dataset.id); });

/* ---------- sheets ---------- */
let lastFocus = null;
function openSheet(id) { lastFocus = document.activeElement; $(id).hidden = false; $(id).scrollTop = 0; document.body.style.overflow = "hidden"; }
function closeSheet(id) {
  $(id).hidden = true;
  if ($("wizard").hidden && $("detail").hidden && $("menu").hidden) document.body.style.overflow = "";
  if (lastFocus && document.contains(lastFocus)) lastFocus.focus({ preventScroll: true });
}
["wizard", "detail", "menu"].forEach((id) => $(id).addEventListener("click", (e) => {
  if (e.target === $(id) && matchMedia("(min-width: 720px)").matches) id === "wizard" ? closeWizard() : closeSheet(id);
}));
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!$("menu").hidden) closeSheet("menu");
  else if (!$("detail").hidden) closeSheet("detail");
  else if (W) closeWizard();
});

/* ---------- detail ---------- */
function openDetail(id) {
  const c = cans.find((x) => x.id === id); if (!c) return;
  const src = photoSrc(c);
  const rows = FIELDS.map((f) => {
    let v = c[f.key];
    if (f.key === "date") v = fmtDate(v);
    if (f.key === "height") v = dispHeight(v);
    if (f.key === "country") v = dispCountry(v);
    return `<dt>${f.label}</dt><dd>${esc(cap(v) || "—")}</dd>`;
  }).join("");
  $("detailBody").innerHTML = `
    <div class="bar"><button class="ghost" id="dClose" type="button">← К списку</button>
      <span><button class="ghost" id="dEdit" type="button">Изменить</button><button class="ghost danger" id="dDel" type="button">Удалить</button></span></div>
    ${src ? `<img class="photo-big" src="${src}" alt="Фото банки">` : `<label class="photo-slot" id="dAddPhoto">Фото нет.<br>Нажмите «Изменить»,<br>чтобы добавить</label>`}
    <div class="detail-title">${esc(dispCountry(c.country))} · ${esc(cap(c.flavor))}</div>
    <dl><dt>Номер</dt><dd>${esc(c.num)}</dd>${rows}</dl>
    <div id="dConfirm"></div>`;
  openSheet("detail");
  $("dClose").onclick = () => closeSheet("detail");
  $("dEdit").onclick = () => { closeSheet("detail"); openWizard(c); };
  $("dDel").onclick = () => {
    $("dConfirm").innerHTML = `<div class="confirm">Удалить банку №${esc(c.num)} из коллекции?
      <button class="btn primary" id="dYes" type="button">Удалить</button><button class="btn" id="dNo" type="button">Отмена</button></div>`;
    $("dNo").onclick = () => ($("dConfirm").innerHTML = "");
    $("dYes").onclick = async () => {
      try { await dbDelete(c.id); cans = cans.filter((x) => x.id !== c.id); render(); closeSheet("detail"); toast("Удалено"); }
      catch (e) { toast("Не удалось удалить. Попробуйте ещё раз."); }
    };
  };
}

/* ---------- photos ---------- */
function loadImage(blob) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(blob), img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); res(img); };
    img.onerror = () => { URL.revokeObjectURL(url); rej(new Error("Не получилось открыть фото. Попробуйте другой снимок.")); };
    img.src = url;
  });
}
function toJpeg(img, max, quality) {
  const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
  const cv = document.createElement("canvas");
  cv.width = Math.round(img.naturalWidth * k); cv.height = Math.round(img.naturalHeight * k);
  cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
  return new Promise((res, rej) => cv.toBlob((b) => (b ? res(b) : rej(new Error("photo"))), "image/jpeg", quality));
}
async function preparePhoto(file) {
  const img = await loadImage(file);
  return { photo: await toJpeg(img, 1400, 0.85), thumb: await toJpeg(img, 360, 0.78) };
}

/* ---------- add / edit questionnaire ---------- */
let W = null;
function suggestions(key) {
  const f = FIELDS.find((x) => x.key === key);
  if (f.fixed) return f.fixed;
  const counts = {};
  cans.forEach((c) => { const v = norm(c[key]); if (v) counts[v] = (counts[v] || 0) + 1; });
  if (f.def) counts[f.def] = (counts[f.def] || 0) + 1e6;
  return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 12).map((x) => x[0]);
}
function openWizard(existing) {
  const nextNum = cans.reduce((m, c) => Math.max(m, +c.num || 0), 0) + 1;
  const draft = existing ? { ...existing } : { id: uid(), num: nextNum };
  if (!existing) FIELDS.forEach((f) => { if (f.def) draft[f.key] = f.def; });
  W = { draft, step: 0, editing: !!existing, preview: existing ? photoSrc(existing) : "", newPhoto: null };
  openSheet("wizard"); drawStep();
}
function closeWizard() { if (W?.newPreview) URL.revokeObjectURL(W.newPreview); W = null; closeSheet("wizard"); }
const totalSteps = () => FIELDS.length + 2;

function drawStep() {
  const s = W.step, body = $("wizardBody");
  const pct = Math.round((s / (totalSteps() - 1)) * 100);
  const bar = `<div class="bar"><button class="ghost" id="wBack" type="button">${s === 0 ? "Отмена" : "← Назад"}</button>
    <div class="progress" aria-hidden="true"><i style="width:${pct}%"></i></div><button class="ghost" id="wClose" type="button">Закрыть</button></div>`;

  if (s === 0) {
    body.innerHTML = bar + `
      <div class="q" style="text-align:center">${W.editing ? "Фото банки" : "Сфотографируйте банку"}</div>
      <p class="hint" style="text-align:center">${W.editing ? "Можно оставить старое фото или сделать новое." : "Лучше вертикально, чтобы банка заняла весь кадр."}</p>
      <label class="photo-slot" for="wCam">${W.preview ? `<img src="${W.preview}" alt="Фото банки">` : `<span>Нажмите, чтобы<br>сделать фото</span>`}</label>
      <div class="row">
        <label class="btn"><input id="wCam" type="file" accept="image/*" capture="environment">Камера</label>
        <label class="btn"><input id="wGal" type="file" accept="image/*">Выбрать файл</label>
        <button class="btn primary" id="wNext" type="button">${W.preview ? "Дальше" : "Без фото"}</button>
      </div>`;
    const onFile = async (e) => {
      const f = e.target.files[0]; if (!f) return;
      try {
        W.newPhoto = await preparePhoto(f);
        if (W.newPreview) URL.revokeObjectURL(W.newPreview);
        W.newPreview = W.preview = URL.createObjectURL(W.newPhoto.thumb);
        drawStep();
      } catch (err) { toast(err.message); }
    };
    $("wCam").onchange = onFile; $("wGal").onchange = onFile;
  } else if (s <= FIELDS.length) {
    const f = FIELDS[s - 1], val = W.draft[f.key] ?? "";
    const sug = f.type === "textarea" || f.type === "date" ? [] : suggestions(f.key);
    let input = "";
    if (f.type === "textarea") input = `<textarea class="answer" id="wIn" rows="3">${esc(val)}</textarea>`;
    else if (f.type !== "choice") input = `<input class="answer" id="wIn" type="${f.type}" value="${esc(val)}" autocomplete="off" ${f.type === "text" ? 'enterkeyhint="next"' : ""}>`;
    const label = (v) => (f.key === "height" ? capFirst(HEIGHT_LABEL[v]) : f.key === "country" ? dispCountry(v) : v);
    body.innerHTML = bar + `
      ${W.preview ? `<div class="mini"><img src="${W.preview}" alt=""><span class="num">Вопрос ${s} из ${FIELDS.length}</span></div>` : `<span class="num">Вопрос ${s} из ${FIELDS.length}</span>`}
      <div class="q">${f.q}</div>${f.hint ? `<p class="hint">${f.hint}</p>` : ""}
      ${input}
      ${sug.length ? `<div class="chips">${sug.map((v) => `<button class="chip" type="button" aria-pressed="${norm(val) === v}" data-v="${esc(v)}">${esc(label(v))}</button>`).join("")}</div>` : ""}
      <div class="row end"><button class="btn primary" id="wNext" type="button">${f.required || cap(val) ? "Дальше" : "Пропустить"}</button></div>`;
    const inp = $("wIn");
    if (inp) {
      inp.oninput = () => {
        W.draft[f.key] = inp.value;
        $("wNext").textContent = f.required || cap(inp.value) ? "Дальше" : "Пропустить";
        body.querySelectorAll(".chip").forEach((c) => c.setAttribute("aria-pressed", c.dataset.v === norm(inp.value)));
      };
      inp.onkeydown = (e) => { if (e.key === "Enter" && f.type !== "textarea") { e.preventDefault(); $("wNext").click(); } };
      if (!matchMedia("(pointer: coarse)").matches) inp.focus();
    }
    body.querySelectorAll(".chip").forEach((ch) => (ch.onclick = () => {
      W.draft[f.key] = ch.dataset.v;
      if (inp) { inp.value = ch.dataset.v; inp.dispatchEvent(new Event("input")); }
      else { W.step++; drawStep(); }
    }));
  } else {
    const rows = FIELDS.map((f) => {
      let v = W.draft[f.key];
      if (f.key === "date") v = fmtDate(v);
      if (f.key === "height") v = dispHeight(v);
      if (f.key === "country") v = dispCountry(v);
      return `<dt>${f.label}</dt><dd>${esc(cap(v) || "—")}</dd>`;
    }).join("");
    body.innerHTML = bar + `
      ${W.preview ? `<img class="photo-big" src="${W.preview}" alt="Фото банки">` : ""}
      <div class="q">Всё верно?</div><p class="hint">Нажмите «Назад», чтобы поправить ответ.</p>
      <dl><dt>Номер</dt><dd>${esc(W.draft.num)}</dd>${rows}</dl>
      <div class="row end"><button class="btn primary" id="wNext" type="button">${W.editing ? "Сохранить изменения" : "Добавить в коллекцию"}</button></div>`;
  }

  $("wClose").onclick = closeWizard;
  $("wBack").onclick = () => { if (W.step === 0) closeWizard(); else { W.step--; drawStep(); } };
  $("wNext").onclick = async () => {
    if (s >= 1 && s <= FIELDS.length) {
      const f = FIELDS[s - 1];
      W.draft[f.key] = cap(W.draft[f.key]);
      if (f.required && !W.draft[f.key]) { toast("Это поле нужно заполнить"); return; }
    }
    if (s < totalSteps() - 1) { W.step++; drawStep(); return; }
    await finish();
  };
  $("wizard").scrollTop = 0;
}

async function finish() {
  const btn = $("wNext"); btn.disabled = true; btn.textContent = "Сохраняю…";
  try {
    const c = { ...W.draft };
    if (W.newPhoto) { c.photo = W.newPhoto.photo; c.thumb = W.newPhoto.thumb; }
    c.updatedAt = new Date().toISOString();
    if (!c.createdAt) c.createdAt = c.updatedAt;
    await dbPut(c);
    const i = cans.findIndex((x) => x.id === c.id);
    if (i >= 0) cans[i] = c; else cans.push(c);
    const editing = W.editing;
    closeWizard(); render();
    toast(editing ? "Сохранено" : `Банка №${c.num} добавлена`);
  } catch (e) {
    btn.disabled = false; btn.textContent = "Попробовать снова";
    toast(e && e.name === "QuotaExceededError" ? "На устройстве закончилось место для фото." : "Не удалось сохранить. Попробуйте ещё раз.");
  }
}

["addBtn", "addBtnTop", "emptyAdd"].forEach((id) => ($(id).onclick = () => openWizard(null)));

/* ---------- files: Excel and backup ---------- */
function loadXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  const tryAt = (k) => new Promise((res, rej) => {
    if (k >= XLSX_URLS.length) return rej(new Error("Не удалось загрузить модуль для Excel. Проверьте подключение к интернету."));
    const s = document.createElement("script");
    s.src = XLSX_URLS[k];
    s.onload = () => (window.XLSX ? res(window.XLSX) : tryAt(k + 1).then(res, rej));
    s.onerror = () => { s.remove(); tryAt(k + 1).then(res, rej); };
    document.head.appendChild(s);
  });
  return tryAt(0);
}
async function saveFile(blob, name) {
  const file = new File([blob], name, { type: blob.type });
  const mobile = matchMedia("(pointer: coarse)").matches;
  if (mobile && navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return; }
    catch (e) { if (e.name === "AbortError") return; }
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(file); a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
const today = () => new Date().toISOString().slice(0, 10);

function excelSerial(d) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || "");
  return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000 : (d || "");
}
// list: which cans to export (defaults to the whole collection, by number)
async function exportExcel(list) {
  const part = Array.isArray(list);
  const items = part ? list : [...cans].sort(SORTS["num-asc"]);
  if (!items.length) { toast("Нет банок для выгрузки."); return; }
  try {
    const X = await loadXLSX();
    const rows = items.map((c) => ({
      "номер": c.num, "название": c.name || "", "страна": c.country || "", "примечания": c.notes || "",
      "вкус": c.flavor || "", "событие": c.event || "", "объём": c.volume || "", "высота": c.height || "",
      "год": excelSerial(c.date),
    }));
    const ws = X.utils.json_to_sheet(rows);
    // the date column holds Excel day numbers; show them as dates
    for (let r = 1; r <= rows.length; r++) { const cell = ws[X.utils.encode_cell({ r, c: 8 })]; if (cell && cell.t === "n") cell.z = "dd.mm.yyyy"; }
    ws["!cols"] = [6, 12, 20, 22, 22, 22, 8, 8, 12].map((w) => ({ wch: w }));
    const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "Коллекция");
    const out = X.write(wb, { bookType: "xlsx", type: "array" });
    await saveFile(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `kollekciya-banok-${part ? "vyborka-" : ""}${today()}.xlsx`);
    toast(`Выгружено в Excel: ${plural(items.length, "банка", "банки", "банок")}`);
  } catch (e) { toast(e.message || "Не удалось выгрузить таблицу."); }
}

const COLS = [
  ["num", /^(номер|№|n\b|#)/], ["name", /^назв/], ["country", /^стран/], ["notes", /^прим/], ["flavor", /^вкус/],
  ["event", /^событ/], ["volume", /^о[бъ]*[её]м/], ["height", /^высот/], ["date", /^(год|дата)/],
];
function cellDate(v, X) {
  if (v == null || v === "") return "";
  const pad = (n) => String(n).padStart(2, "0");
  if (v instanceof Date && !isNaN(v)) { const d = new Date(v.getTime() + 12 * 3600e3); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  if (typeof v === "number") { const p = X.SSF.parse_date_code(v); return p ? `${p.y}-${pad(p.m)}-${pad(p.d)}` : String(v); }
  const s = cap(v), m = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(s);
  return m ? `${m[3]}-${pad(m[2])}-${pad(m[1])}` : s;
}
function cleanHeight(v) { const s = norm(v); return s.startsWith("выс") ? "выс" : s.startsWith("низ") ? "низ" : s; }

async function importExcel(file) {
  try {
    const X = await loadXLSX();
    const wb = X.read(await file.arrayBuffer(), { cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const grid = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    const hi = grid.findIndex((r) => r.some((h) => /^стран/.test(norm(h))));
    if (hi < 0) { toast("В таблице не нашлась колонка «страна». Проверьте заголовки.", 4000); return; }
    const idx = {};
    grid[hi].forEach((h, i) => { const k = COLS.find(([, re]) => re.test(norm(h))); if (k && !(k[0] in idx)) idx[k[0]] = i; });
    let added = 0, updated = 0, maxNum = cans.reduce((m, c) => Math.max(m, +c.num || 0), 0);
    const byNum = new Map(cans.map((c) => [String(c.num), c]));
    const changed = [];
    for (const r of grid.slice(hi + 1)) {
      const get = (k) => (k in idx ? r[idx[k]] : "");
      if (!cap(get("country")) && !cap(get("flavor"))) continue;
      const fields = {
        name: cap(get("name")) || "coca-cola", country: cap(get("country")), notes: cap(get("notes")),
        flavor: cap(get("flavor")).replace(/\s+/g, " "), event: cap(get("event")) || "нет",
        volume: cap(get("volume")).replace(",", "."), height: cleanHeight(get("height")), date: cellDate(get("date"), X),
      };
      const n = parseInt(get("num"), 10);
      const ex = Number.isFinite(n) ? byNum.get(String(n)) : null;
      if (ex) { Object.assign(ex, fields, { updatedAt: new Date().toISOString() }); changed.push(ex); updated++; }
      else {
        const num = Number.isFinite(n) ? n : ++maxNum; maxNum = Math.max(maxNum, num);
        const c = { id: uid(), num, ...fields, createdAt: new Date().toISOString() };
        cans.push(c); byNum.set(String(num), c); changed.push(c); added++;
      }
    }
    await dbPutMany(changed);
    render();
    toast(`Загружено из Excel: добавлено ${added}, обновлено ${updated}. Фото сохранены.`, 4000);
  } catch (e) { toast(e.message || "Не удалось прочитать таблицу.", 4000); }
}

const blobToDataURL = (b) => new Promise((r) => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); });
async function exportBackup() {
  try {
    toast("Готовлю копию…");
    const list = [];
    for (const c of cans) {
      const o = { ...c };
      if (c.photo) o.photo = await blobToDataURL(c.photo);
      if (c.thumb) o.thumb = await blobToDataURL(c.thumb);
      list.push(o);
    }
    const json = JSON.stringify({ app: "can-collection", version: 1, exportedAt: new Date().toISOString(), cans: list });
    await saveFile(new Blob([json], { type: "application/json" }), `kollekciya-banok-kopiya-${today()}.json`);
  } catch (e) { toast("Не удалось сохранить копию."); }
}
async function importBackup(file) {
  try {
    const data = JSON.parse(await file.text());
    if (data.app !== "can-collection" || !Array.isArray(data.cans)) { toast("Это не файл резервной копии коллекции.", 4000); return; }
    const list = [];
    for (const o of data.cans) {
      const c = { ...o };
      if (typeof o.photo === "string") c.photo = await (await fetch(o.photo)).blob();
      if (typeof o.thumb === "string") c.thumb = await (await fetch(o.thumb)).blob();
      list.push(c);
    }
    await dbPutMany(list);
    const ids = new Set(list.map((c) => c.id));
    cans = cans.filter((c) => !ids.has(c.id)).concat(list);
    render();
    toast(`Восстановлено банок: ${list.length}`);
  } catch (e) { toast("Не удалось прочитать файл копии.", 4000); }
}

$("importXlsx").onchange = (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) importExcel(f); };
$("importBackup").onchange = (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) importBackup(f); };
$("emptyImport").onclick = () => $("importXlsx").click();

/* ---------- menu & install ---------- */
let installEvt = null;
window.addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); installEvt = e; });
const isStandalone = () => matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);

function openMenu() {
  const install = isStandalone() ? "" : installEvt
    ? `<button class="menu-item" id="mInstall" type="button"><b>Установить приложение</b><span>Иконка на главном экране или рабочем столе</span></button>`
    : `<div class="menu-item"><b>Установить на телефон</b><span>${isIOS()
        ? "В Safari нажмите «Поделиться», затем «На экран „Домой“»."
        : "В меню браузера выберите «Установить приложение» или «Добавить на главный экран»."}</span></div>`;
  $("menuBody").innerHTML = `
    <div class="bar"><div class="menu-title">Меню</div><button class="ghost" id="mClose" type="button">Закрыть</button></div>
    <div class="menu-list">
      <button class="menu-item" id="mExport" type="button"><b>Выгрузить всю коллекцию в Excel</b><span>Таблица с теми же колонками, без фото. Кнопка «Excel» вверху учитывает фильтры</span></button>
      <button class="menu-item" id="mImport" type="button"><b>Загрузить из Excel</b><span>Добавит банки из таблицы. Совпадающие номера обновятся, фото останутся</span></button>
      <button class="menu-item" id="mBackup" type="button"><b>Сохранить резервную копию</b><span>Один файл со всеми банками и фото</span></button>
      <button class="menu-item" id="mRestore" type="button"><b>Восстановить из копии</b><span>Например, чтобы перенести коллекцию на другое устройство</span></button>
      ${install}
    </div>
    <p class="menu-note">Коллекция хранится только на этом устройстве, в этом браузере. Время от времени сохраняйте резервную копию.</p>`;
  openSheet("menu");
  $("mClose").onclick = () => closeSheet("menu");
  $("mExport").onclick = () => { closeSheet("menu"); exportExcel(); };
  $("mImport").onclick = () => { closeSheet("menu"); $("importXlsx").click(); };
  $("mBackup").onclick = () => { closeSheet("menu"); exportBackup(); };
  $("mRestore").onclick = () => { closeSheet("menu"); $("importBackup").click(); };
  const mi = $("mInstall");
  if (mi) mi.onclick = async () => { installEvt.prompt(); await installEvt.userChoice.catch(() => {}); installEvt = null; closeSheet("menu"); };
}
$("menuBtn").onclick = openMenu;
// with filters on, the button exports exactly what is shown, in the shown order; otherwise the whole collection
$("exportBtn").onclick = () => exportExcel(lastFiltering ? lastShown : undefined);

/* ---------- start ---------- */
(async () => {
  try {
    cans = await dbAll();
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  } catch (e) {
    $("status").hidden = false;
    $("status").textContent = "Браузер не даёт сохранять данные (возможно, включён приватный режим). Банки пропадут после закрытия страницы.";
    cans = [];
    run.disabled = true;
  }
  render();
})();

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
}

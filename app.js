"use strict";
/* Коллекция банок.
   Без входа коллекция хранится в браузере (IndexedDB). После входа — в облаке (cloud.js),
   а в браузере остаётся копия для быстрого запуска и работы без интернета. */

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

let records = []; // everything on this device: the local collection and cached cloud collections
let cans = [];    // the collection on screen
let lastShown = [], lastFiltering = false; // what the list currently shows, used by the Excel export
let storageWarning = "";

const owner = () => (Cloud.user() ? Cloud.user().id : "local");
const isCloud = () => !!Cloud.user();
const ownerOf = (r) => r.owner || "local";
function recompute() { const o = owner(); cans = records.filter((r) => ownerOf(r) === o && !r._deleted); }
function upsertRecords(list) {
  const byId = new Map(list.map((r) => [r.id, r]));
  records = records.filter((r) => !byId.has(r.id)).concat(list);
}
// every change goes through here: saved on the device, then sent to the cloud when signed in
async function saveCans(list) {
  const o = owner(), cloud = isCloud();
  list.forEach((c) => { c.owner = o; c._rev = (c._rev || 0) + 1; if (cloud) c._dirty = true; });
  await dbPutMany(list);
  upsertRecords(list);
  render();
  if (cloud) scheduleSync();
}
async function removeCan(c) {
  if (isCloud()) {
    const d = { ...c, _deleted: true, _dirty: true, _rev: (c._rev || 0) + 1 };
    await dbPut(d); upsertRecords([d]); scheduleSync();
  } else {
    await dbDelete(c.id); records = records.filter((r) => r.id !== c.id);
  }
  render();
}

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
  recompute();
  renderStatus();
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
  if (["wizard", "detail", "menu", "auth", "ask"].every((x) => $(x).hidden)) document.body.style.overflow = "";
  if (lastFocus && document.contains(lastFocus)) lastFocus.focus({ preventScroll: true });
}
["wizard", "detail", "menu", "auth", "ask"].forEach((id) => $(id).addEventListener("click", (e) => {
  if (e.target === $(id) && matchMedia("(min-width: 720px)").matches) id === "wizard" ? closeWizard() : closeSheet(id);
}));
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!$("ask").hidden) closeSheet("ask");
  else if (!$("auth").hidden) closeSheet("auth");
  else if (!$("menu").hidden) closeSheet("menu");
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
    ${src ? `<img class="photo-big" src="${src}" alt="Фото банки">` : (c.photo_path ? `<div class="photo-slot">Фото загружается…</div>` : `<div class="photo-slot">Фото нет.<br>Нажмите «Изменить»,<br>чтобы добавить</div>`)}
    <div class="detail-title">${esc(dispCountry(c.country))} · ${esc(cap(c.flavor))}</div>
    <dl><dt>Номер</dt><dd>${esc(c.num)}</dd>${rows}</dl>
    <div id="dConfirm"></div>`;
  openSheet("detail");
  if (!c.photo && c.photo_path && isCloud()) {
    loadFullPhoto(c).then((b) => {
      const img = $("detailBody").querySelector(".photo-big");
      if (b && img && !$("detail").hidden) img.src = photoSrc(c);
      else if (b && !img && !$("detail").hidden) openDetail(c.id);
    }).catch(() => {});
  }
  $("dClose").onclick = () => closeSheet("detail");
  $("dEdit").onclick = () => { closeSheet("detail"); openWizard(c); };
  $("dDel").onclick = () => {
    $("dConfirm").innerHTML = `<div class="confirm">Удалить банку №${esc(c.num)} из коллекции?
      <button class="btn primary" id="dYes" type="button">Удалить</button><button class="btn" id="dNo" type="button">Отмена</button></div>`;
    $("dNo").onclick = () => ($("dConfirm").innerHTML = "");
    $("dYes").onclick = async () => {
      try { await removeCan(c); closeSheet("detail"); toast("Удалено"); }
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
    if (W.newPhoto) { c.photo = W.newPhoto.photo; c.thumb = W.newPhoto.thumb; c._photoDirty = true; }
    c.updatedAt = new Date().toISOString();
    if (!c.createdAt) c.createdAt = c.updatedAt;
    await saveCans([c]);
    const editing = W.editing;
    closeWizard();
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
      if (ex) { changed.push({ ...ex, ...fields, updatedAt: new Date().toISOString() }); updated++; }
      else {
        const num = Number.isFinite(n) ? n : ++maxNum; maxNum = Math.max(maxNum, num);
        const c = { id: uid(), num, ...fields, createdAt: new Date().toISOString() };
        cans.push(c); byNum.set(String(num), c); changed.push(c); added++;
      }
    }
    await saveCans(changed);
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
      ["owner", "_dirty", "_photoDirty", "_rev", "photo_path", "thumb_path"].forEach((k) => delete o[k]);
      if (!c.photo && c.photo_path && isCloud()) { try { await loadFullPhoto(c); } catch (e) { /* keep the small copy */ } }
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
    const list = [], me = owner();
    for (const o of data.cans) {
      const c = { ...o };
      ["owner", "_dirty", "_deleted", "_photoDirty", "_rev", "photo_path", "thumb_path"].forEach((k) => delete c[k]);
      // an id already used by another collection on this device gets a new one
      if (records.some((r) => r.id === c.id && ownerOf(r) !== me)) c.id = uid();
      if (typeof o.photo === "string") c.photo = await (await fetch(o.photo)).blob();
      if (typeof o.thumb === "string") c.thumb = await (await fetch(o.thumb)).blob();
      if (c.photo || c.thumb) c._photoDirty = true;
      list.push(c);
    }
    await saveCans(list);
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
  const u = Cloud.user();
  const localCount = records.filter((r) => ownerOf(r) === "local" && !r._deleted).length;
  const account = u
    ? `<div class="menu-item"><b>Аккаунт</b><span>${esc(u.email)} · коллекция в облаке</span></div>
       <button class="menu-item" id="mSync" type="button"><b>Синхронизировать сейчас</b><span>Отправить изменения и загрузить новое с других устройств</span></button>
       ${localCount ? `<button class="menu-item" id="mMigrate" type="button"><b>Перенести банки с устройства (${localCount})</b><span>Банки, добавленные без входа, попадут в облако</span></button>` : ""}`
    : `<button class="menu-item" id="mLogin" type="button"><b>Войти или создать аккаунт</b><span>Коллекция будет в облаке и на всех ваших устройствах</span></button>`;
  $("menuBody").innerHTML = `
    <div class="bar"><div class="menu-title">Меню</div><button class="ghost" id="mClose" type="button">Закрыть</button></div>
    <div class="menu-list">
      ${account}
      <button class="menu-item" id="mExport" type="button"><b>Выгрузить всю коллекцию в Excel</b><span>Таблица с теми же колонками, без фото. Кнопка «Excel» вверху учитывает фильтры</span></button>
      <button class="menu-item" id="mImport" type="button"><b>Загрузить из Excel</b><span>Добавит банки из таблицы. Совпадающие номера обновятся, фото останутся</span></button>
      <button class="menu-item" id="mBackup" type="button"><b>Сохранить резервную копию</b><span>Один файл со всеми банками и фото</span></button>
      <button class="menu-item" id="mRestore" type="button"><b>Восстановить из копии</b><span>Например, чтобы перенести коллекцию на другое устройство</span></button>
      ${install}
      ${u ? `<button class="menu-item" id="mLogout" type="button"><b>Выйти из аккаунта</b><span>Коллекция останется в облаке</span></button>` : ""}
    </div>
    <p class="menu-note">${u
      ? "Коллекция хранится в облаке. На этом устройстве есть её копия для работы без интернета."
      : "Сейчас коллекция хранится только на этом устройстве, в этом браузере. Войдите, чтобы хранить её в облаке, или время от времени сохраняйте резервную копию."}</p>`;
  openSheet("menu");
  $("mClose").onclick = () => closeSheet("menu");
  $("mExport").onclick = () => { closeSheet("menu"); exportExcel(); };
  $("mImport").onclick = () => { closeSheet("menu"); $("importXlsx").click(); };
  $("mBackup").onclick = () => { closeSheet("menu"); exportBackup(); };
  $("mRestore").onclick = () => { closeSheet("menu"); $("importBackup").click(); };
  if ($("mLogin")) $("mLogin").onclick = () => { closeSheet("menu"); openAuth("signin"); };
  if ($("mSync")) $("mSync").onclick = () => { closeSheet("menu"); scheduleSync(); };
  if ($("mMigrate")) $("mMigrate").onclick = () => { closeSheet("menu"); offerMigration(true); };
  if ($("mLogout")) $("mLogout").onclick = () => { closeSheet("menu"); confirmSignOut(); };
  const mi = $("mInstall");
  if (mi) mi.onclick = async () => { installEvt.prompt(); await installEvt.userChoice.catch(() => {}); installEvt = null; closeSheet("menu"); };
}
$("menuBtn").onclick = openMenu;
// with filters on, the button exports exactly what is shown, in the shown order; otherwise the whole collection
$("exportBtn").onclick = () => exportExcel(lastFiltering ? lastShown : undefined);

/* ---------- cloud sync ---------- */
const FIELD_KEYS = ["num", "name", "country", "notes", "flavor", "event", "volume", "height", "date"];
function toRow(r, me) {
  const now = new Date().toISOString();
  const o = { id: r.id, user_id: me, photo_path: r.photo_path || null, thumb_path: r.thumb_path || null, created_at: r.createdAt || now, updated_at: r.updatedAt || now };
  FIELD_KEYS.forEach((k) => (o[k] = k === "num" ? (Number.isFinite(+r.num) ? +r.num : null) : String(r[k] ?? "")));
  return o;
}
function fromRow(row) {
  const o = { id: row.id, photo_path: row.photo_path || "", thumb_path: row.thumb_path || "", createdAt: row.created_at, updatedAt: row.updated_at };
  FIELD_KEYS.forEach((k) => (o[k] = row[k] ?? ""));
  return o;
}
const sameRec = (a, b) => ["updatedAt", "photo_path", "thumb_path", ...FIELD_KEYS].every((k) => String(a[k] ?? "") === String(b[k] ?? ""));

let sync = { state: "idle", msg: "" };
let syncRun = null, syncQueued = false;
function setSync(state, msg = "") { sync = { state, msg }; renderStatus(); }
function scheduleSync() {
  if (!isCloud()) return Promise.resolve();
  if (syncRun) { syncQueued = true; return syncRun; }
  syncRun = doSync().finally(() => {
    syncRun = null;
    if (syncQueued) { syncQueued = false; scheduleSync(); }
  });
  return syncRun;
}

async function doSync() {
  const me = owner();
  setSync("busy");
  try {
    // 1. deletions made on this device
    const gone = records.filter((r) => r.owner === me && r._deleted);
    if (gone.length) {
      await Cloud.remove(gone.map((r) => r.id));
      const paths = gone.flatMap((r) => [r.photo_path, r.thumb_path]).filter(Boolean);
      if (paths.length) await Cloud.removeFiles(paths).catch(() => {});
      for (const r of gone) await dbDelete(r.id);
      const ids = new Set(gone.map((r) => r.id));
      records = records.filter((r) => !(ids.has(r.id) && r._deleted));
    }

    // 2. new and changed cans: photos first, then the rows
    const dirty = records.filter((r) => r.owner === me && r._dirty && !r._deleted);
    const revs = new Map(dirty.map((r) => [r.id, r._rev]));
    for (const r of dirty) {
      if (!r._photoDirty) continue;
      if (r.photo) { const p = `${me}/${r.id}.jpg`; await Cloud.upload(p, r.photo); r.photo_path = p; }
      if (r.thumb) { const p = `${me}/${r.id}_t.jpg`; await Cloud.upload(p, r.thumb); r.thumb_path = p; }
    }
    for (let i = 0; i < dirty.length; i += 100) await Cloud.upsert(dirty.slice(i, i + 100).map((r) => toRow(r, me)));
    const done = [];
    for (const r of dirty) {
      const cur = records.find((x) => x.id === r.id);
      if (!cur) continue;
      if (cur === r && r._rev === revs.get(r.id)) { delete r._dirty; delete r._photoDirty; done.push(r); }
      else if (cur !== r && !cur._photoDirty) {
        // edited again while sending: keep the uploaded photo links, send the new edit next round
        cur.photo_path = cur.photo_path || r.photo_path; cur.thumb_path = cur.thumb_path || r.thumb_path; done.push(cur);
      }
    }
    if (done.length) await dbPutMany(done);

    // 3. the cloud copy is the source of truth for everything not waiting to be sent
    const rows = await Cloud.list();
    if (owner() !== me) return;
    const mine = new Map(records.filter((r) => r.owner === me).map((r) => [r.id, r]));
    const seen = new Set(), put = [];
    for (const row of rows) {
      seen.add(row.id);
      const loc = mine.get(row.id);
      if (loc && (loc._dirty || loc._deleted)) continue;
      const rec = { ...(loc || {}), ...fromRow(row), owner: me };
      if (loc && loc.photo_path !== rec.photo_path) delete rec.photo;
      if (loc && loc.thumb_path !== rec.thumb_path) delete rec.thumb;
      if (!loc || !sameRec(loc, rec)) put.push(rec);
    }
    const stale = [...mine.values()].filter((r) => !seen.has(r.id) && !r._dirty && !r._deleted);
    if (put.length) { await dbPutMany(put); upsertRecords(put); }
    if (stale.length) {
      for (const r of stale) await dbDelete(r.id);
      const ids = new Set(stale.map((r) => r.id));
      records = records.filter((r) => !ids.has(r.id));
    }
    setSync("ok");
    render();

    // 4. small photos for the list
    await fetchThumbs(me);
  } catch (e) {
    if (e.code === "offline") setSync("offline");
    else if (!Cloud.user()) setSync("idle");
    else setSync("error", e.message);
  }
}

async function fetchThumbs(me) {
  const need = records.filter((r) => r.owner === me && !r.thumb && r.thumb_path);
  let i = 0, n = 0;
  const worker = async () => {
    while (i < need.length) {
      const r = need[i++];
      try {
        r.thumb = await Cloud.download(r.thumb_path);
        if (owner() !== me) return;
        await dbPut(r);
        if (++n % 8 === 0) render();
      } catch (e) { if (e.code === "offline") return; }
    }
  };
  await Promise.all([worker(), worker(), worker(), worker()]);
  if (n) render();
}
async function loadFullPhoto(c) {
  if (c.photo || !c.photo_path || !isCloud()) return null;
  const b = await Cloud.download(c.photo_path);
  c.photo = b;
  await dbPut(c);
  return b;
}

/* ---------- status line ---------- */
function renderStatus() {
  const el = $("status");
  const u = Cloud.user();
  const pending = u ? records.filter((r) => r.owner === u.id && (r._dirty || r._deleted)).length : 0;
  let html = "";
  if (!u) {
    html = `<span class="dot"></span><span>Коллекция только на этом устройстве.</span><button class="linkbtn" id="stLogin" type="button">Войти и хранить в облаке</button>`;
  } else {
    const text = {
      busy: "Синхронизация…",
      ok: pending ? `Ждут отправки: ${pending}` : "Сохранено в облаке",
      offline: `Нет интернета${pending ? ` · ждут отправки: ${pending}` : ""}. Изменения отправятся позже.`,
      error: `Не удалось синхронизировать: ${sync.msg}`,
      idle: pending ? `Ждут отправки: ${pending}` : "Облако",
    }[sync.state] || "";
    const cls = sync.state === "ok" && !pending ? "ok" : sync.state === "error" ? "err" : sync.state === "offline" ? "warn" : "";
    html = `<span class="dot ${cls}"></span><span>${esc(text)}</span><span class="who">${esc(u.email)}</span>` +
      (sync.state === "error" || sync.state === "offline" ? `<button class="linkbtn" id="stRetry" type="button">Повторить</button>` : "");
  }
  if (storageWarning) html += `<span class="warn-text">${esc(storageWarning)}</span>`;
  el.innerHTML = html;
  el.hidden = false;
  if ($("stLogin")) $("stLogin").onclick = () => openAuth("signin");
  if ($("stRetry")) $("stRetry").onclick = () => scheduleSync();
}

/* ---------- confirm dialog ---------- */
function ask({ title, text, buttons }) {
  $("askBody").innerHTML = `<div class="menu-title">${esc(title)}</div><p class="hint">${esc(text)}</p>
    <div class="row end">${buttons.map((b, i) => `<button class="btn ${b.primary ? "primary" : ""}" data-i="${i}" type="button">${esc(b.label)}</button>`).join("")}</div>`;
  $("askBody").querySelectorAll("button[data-i]").forEach((el) => (el.onclick = async () => {
    closeSheet("ask");
    const b = buttons[+el.dataset.i];
    if (b.run) await b.run();
  }));
  openSheet("ask");
}

/* ---------- sign in ---------- */
let authMode = "signin", authEmail = "";
function openAuth(mode = "signin", msg = "", isError = false) { authMode = mode; drawAuth(msg, isError); openSheet("auth"); }
function drawAuth(msg = "", isError = false) {
  const m = authMode;
  const title = { signin: "Вход", signup: "Новый аккаунт", reset: "Восстановление пароля", newpass: "Новый пароль" }[m];
  const intro = {
    signin: "Войдите, чтобы коллекция хранилась в облаке и была одинаковой на телефоне и компьютере.",
    signup: "Аккаунт нужен, чтобы хранить коллекцию в облаке. Ваши банки видите только вы.",
    reset: "Пришлём на почту ссылку для смены пароля.",
    newpass: "Придумайте новый пароль для входа.",
  }[m];
  const submit = { signin: "Войти", signup: "Создать аккаунт", reset: "Отправить ссылку", newpass: "Сохранить пароль" }[m];
  const needEmail = m !== "newpass", needPass = m !== "reset";
  $("authBody").innerHTML = `
    <div class="bar"><div class="menu-title">${title}</div><button class="ghost" id="aClose" type="button">Закрыть</button></div>
    <p class="hint">${intro}</p>
    <form id="aForm" class="form" novalidate>
      ${needEmail ? `<label class="field"><span>Email</span><input class="answer" id="aEmail" type="email" autocomplete="email" inputmode="email" autocapitalize="off" spellcheck="false" value="${esc(authEmail)}"></label>` : ""}
      ${needPass ? `<label class="field"><span>Пароль</span><input class="answer" id="aPass" type="password" autocomplete="${m === "signin" ? "current-password" : "new-password"}"></label>` : ""}
      ${m === "signup" || m === "newpass" ? `<p class="field-hint">Не меньше 6 символов.</p>` : ""}
      <p class="form-msg ${isError ? "err" : ""}" id="aMsg" ${msg ? "" : "hidden"}>${esc(msg)}</p>
      <button class="btn primary wide" id="aSubmit" type="submit">${submit}</button>
    </form>
    <div class="auth-links">
      ${m === "signin" ? `<button class="linkbtn" data-mode="signup" type="button">Создать аккаунт</button><button class="linkbtn" data-mode="reset" type="button">Забыли пароль?</button>` : ""}
      ${m === "signup" || m === "reset" ? `<button class="linkbtn" data-mode="signin" type="button">У меня уже есть аккаунт</button>` : ""}
    </div>`;
  $("aClose").onclick = () => closeSheet("auth");
  $("authBody").querySelectorAll("[data-mode]").forEach((b) => (b.onclick = () => {
    if ($("aEmail")) authEmail = cap($("aEmail").value);
    authMode = b.dataset.mode; drawAuth();
  }));
  const say = (t, err) => { const el = $("aMsg"); el.textContent = t; el.hidden = !t; el.classList.toggle("err", !!err); };
  $("aForm").onsubmit = async (e) => {
    e.preventDefault();
    const email = needEmail ? cap($("aEmail").value) : "", pass = needPass ? $("aPass").value : "";
    if (needEmail) authEmail = email;
    if (needEmail && !/^\S+@\S+\.\S+$/.test(email)) return say("Проверьте, правильно ли написан email.", true);
    if (needPass && pass.length < 6) return say("Пароль должен быть не короче 6 символов.", true);
    const btn = $("aSubmit"); btn.disabled = true; btn.textContent = "Подождите…";
    try {
      if (m === "signin") { await Cloud.signIn(email, pass); closeSheet("auth"); toast("Вы вошли"); }
      else if (m === "signup") {
        const needConfirm = await Cloud.signUp(email, pass);
        if (needConfirm) { authMode = "signin"; drawAuth(`Мы отправили письмо на ${email}. Откройте ссылку из письма, затем войдите здесь с этим паролем.`); }
        else { closeSheet("auth"); toast("Аккаунт создан, вы вошли"); }
      } else if (m === "reset") { await Cloud.sendReset(email); say("Если такой аккаунт есть, письмо со ссылкой уже отправлено. Проверьте почту, в том числе папку «Спам»."); btn.disabled = false; btn.textContent = submit; }
      else if (m === "newpass") { await Cloud.setPassword(pass); closeSheet("auth"); toast("Пароль изменён"); }
    } catch (err) {
      say(Cloud.explain(err), true);
      if ($("aSubmit")) { $("aSubmit").disabled = false; $("aSubmit").textContent = submit; }
    }
  };
  if (!matchMedia("(pointer: coarse)").matches) setTimeout(() => ($("aEmail") || $("aPass"))?.focus(), 50);
}

/* ---------- cans added before signing in ---------- */
function offerMigration(fromMenu = false) {
  const locals = records.filter((r) => ownerOf(r) === "local" && !r._deleted);
  if (!isCloud() || !locals.length) return;
  const inCloud = cans.length;
  ask({
    title: "Перенести банки в облако?",
    text: `На этом устройстве ${plural(locals.length, "банка", "банки", "банок")}, добавленных без входа. Перенести их в ваш аккаунт? ` +
      (inCloud ? `В облаке уже ${plural(inCloud, "банка", "банки", "банок")}, перенесённые получат номера после них.` : "Номера сохранятся."),
    buttons: [{ label: fromMenu ? "Отмена" : "Не сейчас" }, { label: "Перенести", primary: true, run: migrateLocal }],
  });
}
async function migrateLocal() {
  const locals = records.filter((r) => ownerOf(r) === "local" && !r._deleted).sort((a, b) => (a.num || 0) - (b.num || 0));
  let max = cans.reduce((m, c) => Math.max(m, +c.num || 0), 0);
  const renumber = cans.length > 0;
  locals.forEach((r) => { if (renumber) r.num = ++max; if (r.photo || r.thumb) r._photoDirty = true; });
  toast(`Переношу в облако: ${plural(locals.length, "банка", "банки", "банок")}`, 3500);
  await saveCans(locals);
}

function confirmSignOut() {
  const me = owner();
  const pending = records.filter((r) => r.owner === me && (r._dirty || r._deleted)).length;
  ask({
    title: "Выйти из аккаунта?",
    text: (pending ? `${plural(pending, "изменение ещё не отправлено", "изменения ещё не отправлены", "изменений ещё не отправлены")} в облако и пропадут. ` : "") +
      "Коллекция останется в облаке, а её копия с этого устройства удалится. Чтобы снова её увидеть, войдите в аккаунт.",
    buttons: [{ label: "Отмена" }, { label: "Выйти", primary: true, run: async () => {
      const mine = records.filter((r) => r.owner === me);
      await Cloud.signOut();
      for (const r of mine) await dbDelete(r.id).catch(() => {});
      records = records.filter((r) => r.owner !== me);
      render();
      toast("Вы вышли из аккаунта");
    } }],
  });
}

let currentUser = Cloud.user() ? Cloud.user().id : null;
Cloud.onChange(async (s) => {
  const now = s ? s.user.id : null;
  if (now === currentUser) { renderStatus(); return; } // token refreshed
  currentUser = now;
  setSync("idle");
  render();
  if (now) { await scheduleSync(); offerMigration(); }
});
window.addEventListener("online", () => scheduleSync());
let lastWake = 0;
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && Date.now() - lastWake > 30000) { lastWake = Date.now(); scheduleSync(); }
});

/* ---------- start ---------- */
(async () => {
  try {
    records = await dbAll();
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  } catch (e) {
    storageWarning = "Браузер не даёт сохранять данные на устройстве (возможно, включён приватный режим).";
    records = [];
    run.disabled = true;
  }
  render();
  try {
    const link = await Cloud.fromLink();
    if (link && link.error) openAuth("signin", Cloud.explain({ code: link.error, message: link.message }), true);
    else if (link && link.type === "recovery") openAuth("newpass");
    else if (link) toast("Email подтверждён, вы вошли");
  } catch (e) { openAuth("signin", Cloud.explain(e), true); }
  if (isCloud()) { lastWake = Date.now(); scheduleSync(); }
})();

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
}

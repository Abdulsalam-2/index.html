/* المُوثِّق الميداني — client-side logic */
'use strict';

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const xmlEsc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const pad = (n) => String(n).padStart(2, '0');
const fmtDate = (d) => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const fmtTime = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const fmtCoords = (lat, lng) => `${lat.toFixed(6)}, ${lng.toFixed(6)}`;

let gps = null;           // { lat, lng, acc, ts, source }
let watchId = null;
let reportPhotos = [];    // { id, base, lat, lng, acc, coordSource, date, caption, wm, rendered }
let letterhead = null;    // dataURL of page background
let docxTemplate = null;  // { name, data: ArrayBuffer }

/* ===================== Local storage (IndexedDB — settings & templates only) ===================== */
const DB = {
  db: null,
  open() {
    if (this.db) return Promise.resolve(this.db);
    return new Promise((res, rej) => {
      const r = indexedDB.open('field-reports', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => { this.db = r.result; res(this.db); };
      r.onerror = () => rej(r.error);
    });
  },
  async get(k) {
    const db = await this.open();
    return new Promise((res) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(undefined); });
  },
  async set(k, v) {
    const db = await this.open();
    return new Promise((res) => { const t = db.transaction('kv', 'readwrite'); if (v === undefined || v === null) t.objectStore('kv').delete(k); else t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => res(); });
  },
};

const DEFAULT_LABELS = {
  lblDate: 'التاريخ',
  lblLocation: 'الموقع',
  lblCoords: 'الإحداثيات',
  lblPreparedBy: 'مُعِدّ التقرير',
  lblPhotoCount: 'عدد الصور',
  lblNotesTitle: 'ملاحظات عامة:',
  lblNoCaption: 'بدون ملاحظات',
};

const SETTING_IDS = [
  'projectName', 'tplPerPage', 'tplAccent', 'tplMt', 'tplMb', 'tplMs', 'tplShowInfo', 'tplRepeatHeader', 'tplShowMeta', 'tplPageNo', 'docxKeepBody',
  'wmEnabled', 'wmDate', 'wmTime', 'wmCoords', 'wmProject', 'wmLocation',
  'lblDate', 'lblLocation', 'lblCoords', 'lblPreparedBy', 'lblPhotoCount', 'lblNotesTitle', 'lblNoCaption',
];

function readSettings() {
  const num = (id, d) => { const v = parseFloat($(id).value); return Number.isFinite(v) ? Math.max(0, v) : d; };
  const txt = (id, d) => { const el = $(id); const v = el ? el.value.trim() : ''; return v || d; };
  return {
    perPage: parseInt($('tplPerPage').value, 10) || 2,
    accent: $('tplAccent').value || '#1e3a8a',
    mt: num('tplMt', 42), mb: num('tplMb', 32), ms: num('tplMs', 15),
    showInfo: $('tplShowInfo').checked,
    repeatHeader: $('tplRepeatHeader').checked,
    showMeta: $('tplShowMeta').checked,
    pageNo: $('tplPageNo').checked,
    keepBody: $('docxKeepBody').checked,
    lblDate: txt('lblDate', DEFAULT_LABELS.lblDate),
    lblLocation: txt('lblLocation', DEFAULT_LABELS.lblLocation),
    lblCoords: txt('lblCoords', DEFAULT_LABELS.lblCoords),
    lblPreparedBy: txt('lblPreparedBy', DEFAULT_LABELS.lblPreparedBy),
    lblPhotoCount: txt('lblPhotoCount', DEFAULT_LABELS.lblPhotoCount),
    lblNotesTitle: txt('lblNotesTitle', DEFAULT_LABELS.lblNotesTitle),
    lblNoCaption: txt('lblNoCaption', DEFAULT_LABELS.lblNoCaption),
  };
}

function resetFixedLabels() {
  Object.entries(DEFAULT_LABELS).forEach(([id, val]) => {
    const el = $(id);
    if (el) el.value = val;
  });
  saveSettings();
}

let saveTimer = null;
function saveSettings() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    const o = {};
    SETTING_IDS.forEach((id) => {
      const el = $(id);
      if (!el) return;
      o[id] = el.type === 'checkbox' ? el.checked : el.value;
    });
    DB.set('settings', o).catch(() => {});
  }, 300);
}

async function loadSettings() {
  try {
    const o = await DB.get('settings');
    if (o) {
      SETTING_IDS.forEach((id) => {
        if (id in o) {
          const el = $(id);
          if (!el) return;
          if (el.type === 'checkbox') el.checked = !!o[id];
          else el.value = o[id];
        }
      });
    }
    letterhead = (await DB.get('letterhead')) || null;
    docxTemplate = (await DB.get('docxTemplate')) || null;
  } catch (e) { console.warn('storage unavailable', e); }
  updateTemplateUI();
  updateWmFieldsUI();
}

/* ===================== GPS ===================== */
function setGpsStatus(text, color, detail) {
  const el = $('gps-status');
  el.textContent = text;
  el.className = 'font-bold ' + color;
  $('gps-detail').innerHTML = detail || '';
  $('gps-status-mini').textContent = text;
  $('gps-status-mini').className = 'text-[10px] ' + color;
}

function showGps() {
  if (!gps) return;
  const ago = Math.round((Date.now() - gps.ts) / 1000);
  const src = gps.source === 'manual' ? 'إدخال يدوي' : `دقة ±${Math.round(gps.acc || 0)} م`;
  setGpsStatus(`GPS: ${gps.lat.toFixed(5)}, ${gps.lng.toFixed(5)}`, 'text-emerald-400',
    `${src}${gps.source === 'manual' ? '' : ' — آخر تحديث قبل ' + ago + ' ث'} · <a class="underline text-indigo-300" target="_blank" rel="noopener noreferrer" href="https://maps.google.com/?q=${gps.lat},${gps.lng}">عرض على الخريطة</a>`);
}

function gpsErrorMessage(err) {
  const ua = navigator.userAgent;
  const ios = /iPhone|iPad|iPod/i.test(ua);
  if (err.code === 1) {
    return ['تم رفض إذن الموقع', ios
      ? 'فعّل الموقع: الإعدادات ← الخصوصية والأمان ← خدمات الموقع (تشغيل) ← Safari/المتصفح ← "أثناء الاستخدام"، ثم أعد تحميل الصفحة واضغط "تحديد موقعي الآن".'
      : 'اضغط على رمز القفل 🔒 بجانب عنوان الموقع ← الأذونات ← الموقع ← سماح، وتأكد أن الموقع (GPS) مفعّل في الجهاز، ثم اضغط "تحديد موقعي الآن".'];
  }
  if (err.code === 2) return ['تعذر تحديد الموقع', 'تأكد من تفعيل خدمة الموقع (GPS) في الجهاز وأنك في مكان مكشوف أو متصل بالإنترنت.'];
  if (err.code === 3) return ['انتهت مهلة تحديد الموقع', 'الإشارة ضعيفة. انتقل لمكان مكشوف واضغط "تحديد موقعي الآن" مرة أخرى، أو أدخل الإحداثيات يدوياً.'];
  return ['تعذر تحديد الموقع', err.message || ''];
}

function onPosition(pos) {
  const c = pos.coords;
  if (gps && gps.source === 'gps' && Date.now() - gps.ts < 15000 && c.accuracy > gps.acc * 1.5) return;
  gps = { lat: c.latitude, lng: c.longitude, acc: c.accuracy, ts: Date.now(), source: 'gps' };
  showGps();
}

function onPositionError(err) {
  if (gps) return;
  const [t, d] = gpsErrorMessage(err);
  setGpsStatus(t, 'text-red-400', d);
}

function startGPS(userInitiated) {
  if (!window.isSecureContext) {
    setGpsStatus('الموقع يتطلب اتصالاً آمناً', 'text-red-400', 'افتح التطبيق عبر رابط يبدأ بـ https://');
    return;
  }
  if (!('geolocation' in navigator)) {
    setGpsStatus('المتصفح لا يدعم تحديد الموقع', 'text-red-400', 'أدخل الإحداثيات يدوياً.');
    return;
  }
  if (!gps || userInitiated) setGpsStatus('جاري تحديد موقعك...', 'text-amber-400', 'إذا ظهرت رسالة طلب الإذن اختر "سماح".');
  navigator.geolocation.getCurrentPosition(onPosition, onPositionError, { enableHighAccuracy: false, timeout: 15000, maximumAge: 60000 });
  if (watchId !== null) navigator.geolocation.clearWatch(watchId);
  watchId = navigator.geolocation.watchPosition(onPosition, onPositionError, { enableHighAccuracy: true, timeout: 30000, maximumAge: 5000 });
}

function getFreshPosition(timeoutMs = 8000) {
  if (gps && (gps.source === 'manual' || Date.now() - gps.ts < 60000)) return Promise.resolve(gps);
  if (!('geolocation' in navigator) || !window.isSecureContext) return Promise.resolve(gps);
  return new Promise((resolve) => {
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(gps); } };
    setTimeout(finish, timeoutMs);
    navigator.geolocation.getCurrentPosition((p) => { onPosition(p); finish(); }, () => finish(),
      { enableHighAccuracy: true, timeout: timeoutMs, maximumAge: 30000 });
  });
}

function validateLatLng(lat, lng) {
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
  if (lat === 0 && lng === 0) return null;
  return { lat, lng };
}

function parseCoords(text) {
  const raw = String(text || '').trim();
  if (!raw) return null;
  let decoded = raw;
  try { decoded = decodeURIComponent(raw); } catch {}

  const pairRe = '(-?\\d{1,2}(?:\\.\\d+)?)\\s*[,،\\s]+\\s*(-?\\d{1,3}(?:\\.\\d+)?)';
  const patterns = [
    new RegExp(`@${pairRe}`),
    new RegExp(`[?&](?:q|query|ll|center|destination)=${pairRe}`, 'i'),
    new RegExp(`!3d(-?\\d{1,2}(?:\\.\\d+)?)!4d(-?\\d{1,3}(?:\\.\\d+)?)`),
    new RegExp(`^\\s*${pairRe}\\s*$`),
    new RegExp(pairRe),
  ];

  for (const re of patterns) {
    const m = decoded.match(re);
    if (m) {
      const valid = validateLatLng(parseFloat(m[1]), parseFloat(m[2]));
      if (valid) return valid;
    }
  }
  return null;
}

function applyManualCoords() {
  const c = parseCoords($('manualCoords').value);
  if (!c) { alert('صيغة غير صحيحة. أدخل إحداثيات مثل: 24.7136, 46.6753 أو الصق رابط خرائط Google يحتوي على الإحداثيات.'); return; }
  if (watchId !== null) { navigator.geolocation.clearWatch(watchId); watchId = null; }
  gps = { lat: c.lat, lng: c.lng, acc: 0, ts: Date.now(), source: 'manual' };
  showGps();
}

/* ===================== Photos ===================== */
function loadImage(src) {
  return new Promise((res, rej) => { const img = new Image(); img.onload = () => res(img); img.onerror = rej; img.src = src; });
}

async function fileToBase(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url);
    const max = 1600;
    let w = img.naturalWidth, h = img.naturalHeight;
    const s = Math.min(1, max / Math.max(w, h));
    w = Math.round(w * s); h = Math.round(h * s);
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    c.getContext('2d').drawImage(img, 0, 0, w, h);
    return c.toDataURL('image/jpeg', 0.88);
  } finally { URL.revokeObjectURL(url); }
}

async function readExif(file) {
  try {
    if (!window.exifr) return {};
    const d = await exifr.parse(file, { tiff: true, exif: true, gps: true });
    if (!d) return {};
    const out = {};
    if (Number.isFinite(d.latitude) && Number.isFinite(d.longitude) && !(d.latitude === 0 && d.longitude === 0)) { out.lat = d.latitude; out.lng = d.longitude; }
    const dt = d.DateTimeOriginal || d.CreateDate;
    if (dt instanceof Date && !isNaN(dt)) out.date = dt;
    return out;
  } catch { return {}; }
}

async function handleImageUpload(e, kind) {
  const files = Array.from(e.target.files || []);
  e.target.value = '';
  if (!files.length) return;
  $('busy').classList.remove('hidden');
  try {
    const pos = await getFreshPosition();
    for (const file of files) {
      const exif = await readExif(file);
      let date = exif.date;
      if (!date) date = kind === 'gallery' && file.lastModified ? new Date(file.lastModified) : new Date();
      const photo = {
        id: Date.now() + Math.random(),
        base: await fileToBase(file),
        caption: '', wm: true, date,
        lat: null, lng: null, acc: null, coordSource: 'none',
      };
      if (exif.lat != null) Object.assign(photo, { lat: exif.lat, lng: exif.lng, coordSource: 'exif' });
      else if (pos) Object.assign(photo, { lat: pos.lat, lng: pos.lng, acc: pos.acc, coordSource: pos.source });
      await renderWatermark(photo);
      reportPhotos.push(photo);
    }
  } catch (err) {
    console.error(err);
    alert('تعذر معالجة إحدى الصور.');
  } finally {
    $('busy').classList.add('hidden');
  }
  renderPhotos();
}

function wmOptions() {
  return {
    enabled: $('wmEnabled').checked,
    date: $('wmDate').checked, time: $('wmTime').checked, coords: $('wmCoords').checked,
    project: $('wmProject').checked, location: $('wmLocation').checked,
    projectName: $('projectName').value.trim(), locationName: $('locationName').value.trim(),
  };
}

let fontReady = null;
async function renderWatermark(photo) {
  const o = wmOptions();
  if (!o.enabled || !photo.wm) { photo.rendered = photo.base; return; }
  if (!fontReady) fontReady = document.fonts ? document.fonts.load('bold 24px Cairo').catch(() => {}) : Promise.resolve();
  await fontReady;
  const img = await loadImage(photo.base);
  const w = img.naturalWidth, h = img.naturalHeight;
  const c = $('watermarkCanvas');
  c.width = w; c.height = h;
  const ctx = c.getContext('2d');
  ctx.drawImage(img, 0, 0, w, h);

  const lines = [];
  const top = [];
  if (o.project && o.projectName) top.push(`المشروع: ${o.projectName}`);
  if (o.location && o.locationName) top.push(`الموقع: ${o.locationName}`);
  if (top.length) lines.push({ text: top.join('  |  '), color: '#ffffff' });
  const dt = [];
  if (o.date) dt.push(`📅 ${fmtDate(photo.date)}`);
  if (o.time) dt.push(`🕒 ${fmtTime(photo.date)}`);
  if (dt.length) lines.push({ text: dt.join('    '), color: '#ffffff', ltr: true });
  if (o.coords) {
    lines.push(photo.lat != null
      ? { text: `📍 ${fmtCoords(photo.lat, photo.lng)}`, color: '#34d399', ltr: true }
      : { text: '📍 الإحداثيات غير متوفرة', color: '#fbbf24' });
  }
  if (!lines.length) { photo.rendered = photo.base; return; }

  const fs = Math.max(16, Math.round(Math.min(w, h * 1.3) * 0.032));
  const lh = Math.round(fs * 1.5);
  const padY = Math.round(fs * 0.6), padX = Math.round(fs * 0.9);
  const barH = lines.length * lh + padY * 2;
  const grad = ctx.createLinearGradient(0, h - barH, 0, h);
  grad.addColorStop(0, 'rgba(15,23,42,0.55)');
  grad.addColorStop(1, 'rgba(15,23,42,0.85)');
  ctx.fillStyle = grad;
  ctx.fillRect(0, h - barH, w, barH);
  ctx.font = `bold ${fs}px Cairo, Arial, sans-serif`;
  ctx.textBaseline = 'middle';
  ctx.shadowColor = 'rgba(0,0,0,0.6)';
  ctx.shadowBlur = 3;
  lines.forEach((l, i) => {
    const y = h - barH + padY + lh * i + lh / 2;
    ctx.fillStyle = l.color;
    ctx.direction = l.ltr ? 'ltr' : 'rtl';
    ctx.textAlign = 'right';
    ctx.fillText(l.text, w - padX, y, w - padX * 2);
  });
  ctx.shadowBlur = 0;
  photo.rendered = c.toDataURL('image/jpeg', 0.88);
}

function updateWmFieldsUI() {
  $('wm-fields').style.opacity = $('wmEnabled').checked ? '1' : '0.4';
}

let wmFeedbackTimer = null;
function showWmFeedback() {
  const el = $('wm-feedback');
  if (!el) return;
  el.classList.remove('hidden');
  clearTimeout(wmFeedbackTimer);
  wmFeedbackTimer = setTimeout(() => el.classList.add('hidden'), 2000);
}

async function refreshWatermarksNow(showToast = false) {
  if (!reportPhotos.length) return;
  for (const p of reportPhotos) await renderWatermark(p);
  renderPhotos();
  if (showToast) showWmFeedback();
}

function onWatermarkChange() {
  updateWmFieldsUI();
  saveSettings();
  refreshWatermarksNow(false);
}

function coordLabel(p) {
  if (p.lat == null) return '<span class="text-amber-400">📍 بدون إحداثيات — اضغط للإضافة</span>';
  const src = { exif: 'من بيانات الصورة', gps: 'GPS', manual: 'يدوي' }[p.coordSource] || '';
  return `<span class="text-emerald-400" dir="ltr">📍 ${fmtCoords(p.lat, p.lng)}</span> <span class="text-slate-500">(${src})</span>`;
}

function renderPhotos() {
  const container = $('photos-container');
  $('photo-count').textContent = reportPhotos.length;
  container.innerHTML = '';
  reportPhotos.forEach((photo, idx) => {
    const card = document.createElement('div');
    card.className = 'bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-md';
    card.innerHTML = `
      <div class="relative">
        <img src="${photo.rendered}" class="w-full max-h-80 object-contain bg-black" />
        <button data-act="del" class="absolute top-2 left-2 bg-red-600 text-white w-8 h-8 rounded-full flex items-center justify-center text-xs">✕</button>
        <span class="absolute top-2 right-2 bg-slate-950/80 text-white text-[10px] px-2.5 py-1 rounded-full">صورة #${idx + 1}</span>
      </div>
      <div class="p-3 space-y-2">
        <input type="text" data-act="cap" placeholder="اكتب ملاحظة للصورة..." value="${esc(photo.caption)}" class="inp text-xs" />
        <div class="flex items-center justify-between gap-2 text-[11px]">
          <button data-act="coords" class="text-right">${coordLabel(photo)}</button>
          <span class="text-slate-400 whitespace-nowrap" dir="ltr">${fmtDate(photo.date)} ${fmtTime(photo.date)}</span>
        </div>
        <label class="flex items-center gap-2 text-[11px] text-slate-300"><input type="checkbox" class="chk" data-act="wm" ${photo.wm ? 'checked' : ''} ${$('wmEnabled').checked ? '' : 'disabled'} /> إضافة البصمة على هذه الصورة</label>
      </div>`;
    card.querySelector('[data-act=del]').onclick = () => { if (confirm('حذف الصورة؟')) { reportPhotos.splice(idx, 1); renderPhotos(); } };
    card.querySelector('[data-act=cap]').oninput = (ev) => { photo.caption = ev.target.value; };
    card.querySelector('[data-act=wm]').onchange = async (ev) => { photo.wm = ev.target.checked; await renderWatermark(photo); renderPhotos(); };
    card.querySelector('[data-act=coords]').onclick = async () => {
      const cur = photo.lat != null ? fmtCoords(photo.lat, photo.lng) : (gps ? fmtCoords(gps.lat, gps.lng) : '');
      const v = prompt('إحداثيات الصورة (خط العرض, خط الطول أو رابط Google Maps) — اتركه فارغاً للحذف:', cur);
      if (v === null) return;
      if (!v.trim()) Object.assign(photo, { lat: null, lng: null, coordSource: 'none' });
      else { const c = parseCoords(v); if (!c) { alert('صيغة غير صحيحة'); return; } Object.assign(photo, { lat: c.lat, lng: c.lng, coordSource: 'manual' }); }
      await renderWatermark(photo); renderPhotos();
    };
    container.appendChild(card);
  });
}

function clearAllData() {
  if (confirm('هل تريد البدء بتقرير جديد؟ (يتم الإبقاء على القالب واسم مُعِدّ التقرير)')) {
    reportPhotos = [];
    ['locationName', 'generalNotes'].forEach((id) => { $(id).value = ''; });
    renderPhotos();
  }
}

/* ===================== Template uploads ===================== */
let pdfjsLoading = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (!pdfjsLoading) {
    pdfjsLoading = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      s.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'; res(window.pdfjsLib); };
      s.onerror = () => { pdfjsLoading = null; rej(new Error('pdf.js failed')); };
      document.head.appendChild(s);
    });
  }
  return pdfjsLoading;
}

async function pdfToImage(file) {
  const pdfjs = await loadPdfJs();
  const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
  const page = await pdf.getPage(1);
  const vp1 = page.getViewport({ scale: 1 });
  const scale = 1800 / vp1.width;
  const vp = page.getViewport({ scale });
  const c = document.createElement('canvas');
  c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: ctx, viewport: vp }).promise;
  return c.toDataURL('image/jpeg', 0.9);
}

async function imageFileToLetterhead(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url);
    const s = Math.min(1, 2000 / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.round(img.naturalWidth * s); c.height = Math.round(img.naturalHeight * s);
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(img, 0, 0, c.width, c.height);
    return c.toDataURL('image/jpeg', 0.9);
  } finally { URL.revokeObjectURL(url); }
}

async function handleLetterheadUpload(e) {
  const file = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
    if (/\.docx?$/i.test(file.name)) { alert('لملفات Word استخدم زر "رفع قالب Word" بالأسفل.'); return; }
    letterhead = isPdf ? await pdfToImage(file) : await imageFileToLetterhead(file);
    await DB.set('letterhead', letterhead);
    updateTemplateUI();
  } catch (err) {
    console.error(err);
    alert('تعذر قراءة الملف. استخدم صورة PNG/JPG أو ملف PDF.');
  }
}

function removeLetterhead() { letterhead = null; DB.set('letterhead', null); updateTemplateUI(); }

async function handleDocxUpload(e) {
  const file = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!file) return;
  if (!/\.docx$/i.test(file.name)) { alert('يرجى رفع ملف بصيغة ‎.docx (إذا كان الملف ‎.doc افتحه في Word واحفظه بصيغة ‎.docx).'); return; }
  try {
    const data = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(data);
    if (!zip.file('word/document.xml')) throw new Error('invalid docx');
    docxTemplate = { name: file.name, data };
    await DB.set('docxTemplate', docxTemplate);
    updateTemplateUI();
  } catch (err) {
    console.error(err);
    alert('الملف ليس مستند Word صالحاً.');
  }
}

function removeDocx() { docxTemplate = null; DB.set('docxTemplate', null); updateTemplateUI(); }

function updateTemplateUI() {
  $('lh-preview').classList.toggle('hidden', !letterhead);
  $('lh-remove').classList.toggle('hidden', !letterhead);
  if (letterhead) $('lh-img').src = letterhead;
  $('docx-name').classList.toggle('hidden', !docxTemplate);
  $('docx-remove').classList.toggle('hidden', !docxTemplate);
  $('docx-keep-wrap').classList.toggle('hidden', !docxTemplate);
  $('docx-keep-wrap').classList.toggle('flex', !!docxTemplate);
  if (docxTemplate) $('docx-name').textContent = `✔ القالب الحالي: ${docxTemplate.name}`;
}

/* ===================== Report data ===================== */
function reportInfo() {
  const s = readSettings();
  const rows = [];
  const loc = $('locationName').value.trim();
  const by = $('preparedBy').value.trim();
  rows.push([s.lblDate, fmtDate(new Date())]);
  if (loc) rows.push([s.lblLocation, loc]);
  const withCoords = reportPhotos.find((p) => p.lat != null);
  const c = gps || withCoords;
  if (c) rows.push([s.lblCoords, fmtCoords(c.lat, c.lng)]);
  if (by) rows.push([s.lblPreparedBy, by]);
  rows.push([s.lblPhotoCount, String(reportPhotos.length)]);
  return {
    title: $('projectName').value.trim() || 'تقرير فحص ميداني',
    rows,
    notes: $('generalNotes').value.trim(),
    notesTitle: s.lblNotesTitle,
    noCaption: s.lblNoCaption,
  };
}

function photoMeta(p) {
  const parts = [`${fmtDate(p.date)} ${fmtTime(p.date)}`];
  if (p.lat != null) parts.push(`GPS: ${fmtCoords(p.lat, p.lng)}`);
  return parts.join('   •   ');
}

/* ===================== PDF / print pages ===================== */
const LAYOUTS = { 1: [1, 1], 2: [1, 2], 4: [2, 2], 6: [2, 3] };

function buildPages() {
  const s = readSettings();
  const info = reportInfo();
  const [cols, rows] = LAYOUTS[s.perPage] || LAYOUTS[2];
  const per = cols * rows;
  const chunks = [];
  for (let i = 0; i < reportPhotos.length; i += per) chunks.push(reportPhotos.slice(i, i + per));
  if (!chunks.length) chunks.push([]);
  const total = chunks.length;
  const bg = letterhead ? `<img class="lh" src="${letterhead}" />` : '';

  const html = chunks.map((chunk, pi) => {
    let head = '';
    if (pi === 0 || s.repeatHeader) {
      head += `<h1 style="color:${s.accent};font-size:16pt;font-weight:700;margin:0 0 2.5mm 0;text-align:center">${esc(info.title)}</h1>`;
      if (s.showInfo) {
        head += `<table class="info">${info.rows.map(([k, v]) => `<tr><td class="k" style="color:${s.accent}">${esc(k)}</td><td>${esc(v)}</td></tr>`).join('')}</table>`;
      }
      if (info.notes && pi === 0) head += `<div class="notes"><b style="color:${s.accent}">${esc(info.notesTitle)} </b>${esc(info.notes)}</div>`;
    }
    const cells = chunk.map((p, i) => {
      const n = pi * per + i + 1;
      return `<div class="cell">
        <div class="imgbox"><img src="${p.rendered}" /></div>
        <div class="cap"><span style="color:${s.accent}">${n}.</span> ${esc(p.caption || info.noCaption)}</div>
        ${s.showMeta ? `<div class="meta"><span dir="ltr">${esc(photoMeta(p))}</span></div>` : ''}
      </div>`;
    }).join('');
    return `<div class="page-wrap"><div class="page">${bg}
      <div class="content" style="top:${s.mt}mm;bottom:${s.mb}mm;right:${s.ms}mm;left:${s.ms}mm">
        ${head}
        <div class="grid-photos" style="grid-template-columns:repeat(${cols},1fr);grid-template-rows:repeat(${rows},minmax(0,1fr))">${cells}</div>
        ${s.pageNo ? `<div class="pno">صفحة ${pi + 1} من ${total}</div>` : ''}
      </div></div></div>`;
  }).join('');
  $('pages').innerHTML = html;
  fitPages();
}

function fitPages() {
  const pxPerMm = 96 / 25.4;
  const pageW = 210 * pxPerMm;
  const pageH = 297 * pxPerMm;
  const z = Math.min(1, (window.innerWidth - 24) / pageW);
  document.querySelectorAll('#pages .page-wrap').forEach((wrap) => {
    const p = wrap.firstElementChild;
    if (!p) return;
    if (z < 1) {
      wrap.style.width = `${Math.round(pageW * z)}px`;
      wrap.style.height = `${Math.round(pageH * z)}px`;
      p.style.transform = `scale(${z})`;
    } else {
      wrap.style.width = '';
      wrap.style.height = '';
      p.style.transform = '';
    }
  });
}

function openPreview() {
  if (!reportPhotos.length) { alert('يرجى التقاط صورة واحدة على الأقل!'); return; }
  buildPages();
  $('preview').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closePreview() {
  $('preview').classList.remove('open');
  document.body.style.overflow = '';
}

window.addEventListener('resize', () => { if ($('preview').classList.contains('open')) fitPages(); });
window.addEventListener('beforeprint', () => { if (!$('preview').classList.contains('open') && reportPhotos.length) buildPages(); });

/* ===================== Word (.docx) export ===================== */
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
};
const NS_DECL = Object.entries(NS).map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');
const REL_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const REL_HDR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const mm2tw = (mm) => Math.round(mm * 56.6929);
const TW2EMU = 635;

function dataUrlToBytes(url) {
  const b64 = url.split(',')[1];
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function hexColor(c) { return String(c || '#1e3a8a').replace('#', '').toUpperCase(); }

function wRun(text, o = {}) {
  return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>${o.b ? '<w:b/><w:bCs/>' : ''}${o.color ? `<w:color w:val="${o.color}"/>` : ''}<w:sz w:val="${o.sz || 22}"/><w:szCs w:val="${o.sz || 22}"/>${o.ltr ? '' : '<w:rtl/>'}</w:rPr><w:t xml:space="preserve">${xmlEsc(text)}</w:t></w:r>`;
}

function wPara(runs, o = {}) {
  return `<w:p><w:pPr>${o.keepNext ? '<w:keepNext/>' : ''}${o.pageBreakBefore ? '<w:pageBreakBefore/>' : ''}<w:bidi/><w:spacing w:before="${o.before || 0}" w:after="${o.after ?? 80}"/>${o.center ? '<w:jc w:val="center"/>' : ''}</w:pPr>${runs}</w:p>`;
}

function wInlineImage(rid, cx, cy, id) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Photo ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="photo${id}.jpeg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function wInfoTable(rows, widthTw, accent) {
  const k = Math.round(widthTw * 0.3), v = widthTw - k;
  const b = 'w:val="single" w:sz="4" w:space="0" w:color="94A3B8"';
  const cell = (w, content, fill) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}</w:tcPr>${content}</w:tc>`;
  return `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="${widthTw}" w:type="dxa"/><w:tblBorders><w:top ${b}/><w:left ${b}/><w:bottom ${b}/><w:right ${b}/><w:insideH ${b}/><w:insideV ${b}/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="${k}"/><w:gridCol w:w="${v}"/></w:tblGrid>` +
    rows.map(([kk, vv]) => `<w:tr>${cell(k, wPara(wRun(kk, { b: true, color: accent, sz: 20 }), { after: 0 }), 'F1F5F9')}${cell(v, wPara(wRun(vv, { sz: 20, color: '0F172A' }), { after: 0 }), '')}</w:tr>`).join('') +
    `</w:tbl>`;
}

async function imageSize(url) {
  try {
    const img = await loadImage(url);
    return { w: img.naturalWidth || 1200, h: img.naturalHeight || 900 };
  } catch {
    return { w: 1200, h: 900 };
  }
}

async function buildDocxBody(addImage, usableW, usableH, startId) {
  const s = readSettings();
  const info = reportInfo();
  const accent = hexColor(s.accent);
  let xml = wPara(wRun(info.title, { b: true, sz: 34, color: accent }), { center: true, after: 200 });
  if (s.showInfo) xml += wInfoTable(info.rows, usableW, accent) + wPara('', { after: 120 });
  if (info.notes) xml += wPara(wRun(info.notesTitle + ' ', { b: true, color: accent }) + wRun(info.notes, { color: '0F172A' }), { after: 200 });

  const maxH = s.perPage === 1 ? usableH * 0.78 : s.perPage === 2 ? usableH * 0.38 : usableH * 0.3;
  const maxW = s.perPage >= 4 ? usableW * 0.75 : usableW;
  let id = startId;
  for (let i = 0; i < reportPhotos.length; i++) {
    const p = reportPhotos[i];
    const { w, h } = await imageSize(p.rendered);
    const scale = Math.min(maxW / w, maxH / h);
    const cx = Math.round(w * scale * TW2EMU), cy = Math.round(h * scale * TW2EMU);
    const rid = addImage(dataUrlToBytes(p.rendered), 'jpeg');
    const pageBreak = s.perPage === 1 && i > 0;
    xml += wPara(wInlineImage(rid, cx, cy, id++), { center: true, keepNext: true, after: 60, before: pageBreak ? 0 : 120, pageBreakBefore: pageBreak });
    xml += wPara(wRun(`${i + 1}. `, { b: true, color: accent }) + wRun(p.caption || info.noCaption, { b: true, color: '0F172A' }), { keepNext: s.showMeta, after: s.showMeta ? 0 : 160 });
    if (s.showMeta) xml += wPara(wRun(photoMeta(p), { sz: 16, color: '475569', ltr: true }), { after: 160 });
  }
  return xml;
}

function addContentTypeDefault(ct, ext, type) {
  if (new RegExp(`Extension="${ext}"`, 'i').test(ct)) return ct;
  return ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${type}"/></Types>`);
}

function addRelationship(rels, id, type, target) {
  return rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
}

async function exportDocx() {
  if (!reportPhotos.length) { alert('يرجى التقاط صورة واحدة على الأقل!'); return; }
  if (!window.JSZip) { alert('تعذر تحميل مكتبة Word. تحقق من الاتصال بالإنترنت.'); return; }
  $('busy').classList.remove('hidden');
  try {
    const blob = docxTemplate ? await docxFromTemplate() : await docxFromScratch();
    const name = (($('projectName').value.trim() || 'تقرير-ميداني').replace(/[\\/:*?"<>|]+/g, '-')) + '.docx';
    downloadBlob(blob, name);
  } catch (err) {
    console.error(err);
    alert('تعذر إنشاء ملف Word: ' + (err.message || err));
  } finally {
    $('busy').classList.add('hidden');
  }
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

async function docxFromTemplate() {
  const zip = await JSZip.loadAsync(docxTemplate.data);
  const s = readSettings();
  const docPath = 'word/document.xml';
  const relsPath = 'word/_rels/document.xml.rels';
  const docXml = await zip.file(docPath).async('string');
  let rels = zip.file(relsPath) ? await zip.file(relsPath).async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let ct = await zip.file('[Content_Types].xml').async('string');

  // Parse as XML while preserving exact original string structure around body & sectPr
  const doc = new DOMParser().parseFromString(docXml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('تعذر قراءة القالب');
  const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];

  // Find direct child sectPr at end of body
  let sect = null;
  for (let n = body.lastChild; n; n = n.previousSibling) {
    if (n.nodeType === 1 && n.localName === 'sectPr') { sect = n; break; }
  }
  // If sectPr was inside a paragraph, extract it cleanly
  if (!sect) {
    const embedded = body.getElementsByTagNameNS(NS.w, 'sectPr');
    if (embedded.length) {
      sect = embedded[embedded.length - 1].cloneNode(true);
      body.appendChild(sect);
    }
  }

  let pgW = 11906, pgH = 16838, mL = 1440, mR = 1440, mT = 1440, mB = 1440;
  if (sect) {
    const sz = sect.getElementsByTagNameNS(NS.w, 'pgSz')[0];
    const mar = sect.getElementsByTagNameNS(NS.w, 'pgMar')[0];
    const g = (el, a, d) => { const v = el && parseInt(el.getAttributeNS(NS.w, a) || el.getAttribute('w:' + a), 10); return Number.isFinite(v) ? Math.abs(v) : d; };
    pgW = g(sz, 'w', pgW); pgH = g(sz, 'h', pgH);
    mL = g(mar, 'left', mL); mR = g(mar, 'right', mR); mT = g(mar, 'top', mT); mB = g(mar, 'bottom', mB);
  }
  const usableW = Math.max(3000, pgW - mL - mR);
  const usableH = Math.max(4000, pgH - mT - mB);

  let counter = 0;
  const addImage = (bytes, ext) => {
    counter++;
    const rid = `rIdFR${counter}`;
    const target = `media/fr_photo_${counter}.${ext}`;
    zip.file('word/' + target, bytes);
    rels = addRelationship(rels, rid, REL_IMG, target);
    return rid;
  };
  const generatedXml = await buildDocxBody(addImage, usableW, usableH, 9000);

  // Serialize original sectPr intact so Word header/footer rIds & prefixes stay 100% untouched
  const sectXml = sect ? new XMLSerializer().serializeToString(sect) : '';

  let keptPrefix = '';
  if (s.keepBody) {
    const ser = new XMLSerializer();
    Array.from(body.childNodes).forEach((n) => {
      if (n !== sect && n.nodeType === 1) keptPrefix += ser.serializeToString(n);
    });
  }

  // Replace only the inner contents of <w:body>...</w:body> in the original XML string
  const newDocXml = docXml.replace(/<w:body\b[^>]*>[\s\S]*<\/w:body>/, (match) => {
    const openTag = match.slice(0, match.indexOf('>') + 1);
    return `${openTag}${keptPrefix}${generatedXml}${sectXml}</w:body>`;
  });

  ct = addContentTypeDefault(ct, 'jpeg', 'image/jpeg');
  zip.file(docPath, newDocXml);
  zip.file(relsPath, rels);
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

async function docxFromScratch() {
  const s = readSettings();
  const zip = new JSZip();
  const pgW = 11906, pgH = 16838;
  const mT = mm2tw(letterhead ? s.mt : 20), mB = mm2tw(letterhead ? s.mb : 20), mS = mm2tw(letterhead ? s.ms : 18);
  const usableW = pgW - mS * 2, usableH = pgH - mT - mB;
  let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let counter = 0;
  const addImage = (bytes, ext) => {
    counter++;
    const rid = `rIdImg${counter}`;
    zip.file(`word/media/photo_${counter}.${ext}`, bytes);
    rels = addRelationship(rels, rid, REL_IMG, `media/photo_${counter}.${ext}`);
    return rid;
  };
  const bodyXml = await buildDocxBody(addImage, usableW, usableH, 100);

  let headerRef = '';
  let ctOverrides = '';
  if (letterhead) {
    zip.file('word/media/letterhead.jpeg', dataUrlToBytes(letterhead));
    zip.file('word/_rels/header1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLh" Type="${REL_IMG}" Target="media/letterhead.jpeg"/></Relationships>`);
    const cx = pgW * TW2EMU, cy = pgH * TW2EMU;
    zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS_DECL}><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="14" w:lineRule="exact"/></w:pPr><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Letterhead"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="letterhead.jpeg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLh"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>`);
    rels = addRelationship(rels, 'rIdHdr1', REL_HDR, 'header1.xml');
    headerRef = '<w:headerReference w:type="default" r:id="rIdHdr1"/>';
    ctOverrides += '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
  }

  const sectPr = `<w:sectPr>${headerRef}<w:pgSz w:w="${pgW}" w:h="${pgH}"/><w:pgMar w:top="${mT}" w:right="${mS}" w:bottom="${mB}" w:left="${mS}" w:header="0" w:footer="0" w:gutter="0"/><w:bidi/></w:sectPr>`;
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS_DECL}><w:body>${bodyXml}${sectPr}</w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', rels);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${ctOverrides}</Types>`);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

/* ===================== Init ===================== */
window.addEventListener('load', async () => {
  await loadSettings();
  try {
    if (navigator.permissions && navigator.permissions.query) {
      const st = await navigator.permissions.query({ name: 'geolocation' });
      if (st.state === 'denied') { onPositionError({ code: 1 }); }
      else startGPS(false);
      st.onchange = () => { if (st.state === 'granted') startGPS(true); };
    } else startGPS(false);
  } catch { startGPS(false); }
});

/* Shared helpers for TOW Apartment pages */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const store = { get: k => { try { return localStorage.getItem(k) } catch (e) { return null } }, set: (k, v) => { try { localStorage.setItem(k, v) } catch (e) { } }, del: k => { try { localStorage.removeItem(k) } catch (e) { } } };
const qs = k => new URLSearchParams(location.search).get(k);
const STATUS_TEXT = { list: "Loading…", headInit: "Logging in…", createApproval: "Creating link…", getApproval: "Loading your form…", submitRequest: "Sending your request…", getHosting: "Loading…", saveHosting: "Saving…", save: "Saving…", cancel: "Cancelling…", markAck: "Updating…", createHosting: "Creating…", updateHosting: "Saving…", makePdf: "Preparing PDF…", backup: "Backing up…" };
const DONE_TEXT = { list: "", headInit: "Logged in", createApproval: "Link created", submitRequest: "Sent", saveHosting: "Saved", save: "Saved", cancel: "Cancelled", createHosting: "Created", updateHosting: "Saved", makePdf: "PDF ready", backup: "Backup saved" };
let _busy = 0, _stT = null;
function showStatus(text, kind) {
  let el = document.getElementById("statusBar");
  if (!el) { el = document.createElement("div"); el.id = "statusBar"; document.body.appendChild(el) }
  clearTimeout(_stT);
  el.className = "show" + (kind ? " " + kind : "");
  el.innerHTML = (kind === "busy" ? '<span class="spin"></span>' : kind === "ok" ? "✓" : kind === "err" ? "⚠" : "") + `<span>${esc(text)}</span>`;
  if (kind !== "busy") _stT = setTimeout(() => el.className = "", kind === "err" ? 5000 : 1800);
}
async function api(action, body = {}) {
  const url = window.TOW_API;
  if (!url || url.indexOf("PASTE") >= 0) { showStatus("This page isn't connected yet (config.js)", "err"); throw new Error("This page isn't connected yet (config.js).") }
  _busy++; showStatus(STATUS_TEXT[action] || "Working…", "busy");
  try {
    const res = await fetch(url, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ action, ...body }) });
    let j; try { j = await res.json() } catch (e) { throw new Error("Couldn't reach the server. Check the web app is deployed with access set to Anyone.") }
    if (!j.ok) throw new Error(j.error === "Wrong passcode" ? "Wrong passcode. Check capital letters and try again." : (j.error || "Something went wrong"));
    _busy--; if (!_busy) { DONE_TEXT[action] ? showStatus(DONE_TEXT[action], "ok") : (document.getElementById("statusBar") || {}).className = "" }
    return j;
  } catch (e) {
    _busy = Math.max(0, _busy - 1);
    const m = /Failed to fetch|NetworkError|Load failed/i.test(e.message) ? "No connection to the server. Check your internet and try again." : e.message;
    showStatus(m, "err"); throw new Error(m);
  }
}
const EYE_ON = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"/><circle cx="12" cy="12" r="3"/></svg>';
const EYE_OFF_ICON = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/><line x1="1" y1="1" x2="23" y2="23"/></svg>';
function pwToggle(id) {
  const inp = document.getElementById(id); if (!inp) return;
  const wrap = document.createElement("div"); wrap.className = "pw"; inp.parentNode.insertBefore(wrap, inp); wrap.appendChild(inp);
  const b = document.createElement("button"); b.type = "button"; b.className = "pwbtn"; b.title = "Show / hide"; b.innerHTML = EYE_ON; wrap.appendChild(b);
  b.onclick = () => { const show = inp.type === "password"; inp.type = show ? "text" : "password"; b.innerHTML = show ? EYE_OFF_ICON : EYE_ON; inp.focus() };
  inp.addEventListener("keydown", e => { if (e.key === "Enter") { const btn = document.getElementById(id === "hkey" ? "loginBtn" : "loginBtn"); btn && btn.click() } });
}
const toD = s => { if (!s) return null; const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/); return m ? new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : null };
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const fmtT = d => d.toLocaleTimeString("en-GB", { hour: "numeric", minute: "2-digit", hour12: true }).replace(" ", "").toLowerCase();
const fmt = s => { const d = toD(s); if (!d) return ""; return d.toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short", year: "numeric" }) + ", " + fmtT(d) };
const fmtShort = s => { const d = toD(s); if (!d) return ""; return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) + (d.getHours() || d.getMinutes() ? " " + fmtT(d) : "") };
const nightsOf = (a, b) => { const A = toD(a), B = toD(b); if (!A || !B) return 0; A.setHours(0, 0, 0, 0); B.setHours(0, 0, 0, 0); return Math.max(0, Math.round((B - A) / 864e5)) };
const accompanied = s => { const x = toD(s); if (!x) return null; const day = x.getDay(), h = x.getHours() + x.getMinutes() / 60; return (day >= 2 && day <= 5 && h >= 9 && h < 18) || (day === 0 && h >= 8 && h < 11) };
const keyText = (s, w) => { const a = accompanied(s); return a === null ? "" : a ? `With Ministry Head / Church Admin` : `Self ${w}: keys at the guard house (sealed envelope)` };
const keyShort = (s, w) => { const a = accompanied(s); return a === null ? "" : a ? "🤝 With Ministry Head / Church Admin" : `🔑 Self ${w} · keys at guard house` };
// Unit 1 is kept for guest speakers: offered last for paid stays
const unitOrder = cat => (RATES[cat] === 0 ? [1, 2, 3, 4] : [3, 4, 2, 1]);
const unitsOf = u => String(u || "").match(/[1-4]/g) || [];
const waNum = p => { let d = String(p || "").replace(/\D/g, ""); if (d.startsWith("0")) d = "6" + d; return d };
const waLink = (phone, text) => `https://wa.me/${waNum(phone)}?text=${encodeURIComponent(text)}`;
const earliestBooking = () => { const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() + 7); return d };
const RATES = { "Pastors, Guest Speakers & Mission Pastors": 0, "Member serving that weekend": 50, "Member, personal or non-serving stay": 200, "Non-member": 250 };
const PAID_CATS = ["Member serving that weekend", "Member, personal or non-serving stay", "Non-member"];
const ALL_CATS = ["Pastors, Guest Speakers & Mission Pastors", ...PAID_CATS];
const CHURCH_BANK = "TABERNACLE OF WORSHIP · Hong Leong Bank (Malaysia) · 371 000 128 43";
const TERMS = [
  "Bookings must be made at least 1 week in advance.",
  "Check-in &amp; check-out: Tue–Fri 9am–6pm and Sun 8am–11am, accompanied by the Ministry Head / Church Admin. At other times it is a self check-in/out, with keys left at the guard house in a sealed envelope.",
  "Please take time to read the House Rules posted in the unit.",
  "Keep the unit tidy. Do not move furniture or remove any items.",
  "At check-out: throw all rubbish into the large bin near the security booth, switch off all appliances, lock up and return the keys as arranged.",
  "Please report any damage at once so it can be attended to. We kindly ask guests to cover any repair or replacement needed.",
  "Love gifts go towards cleaning and upkeep. They are offerings, not rental fees."
];
const WHY = "These apartments are a blessing entrusted to our church. Your love gift helps us steward them well, keeping every unit clean, cared for and ready to welcome guest speakers, missionaries and fellow serving personnel. Thank you for sowing into this ministry of hospitality.";
const VERSE = "“Use hospitality one to another without grudging.” — 1 Peter 4:9";

function sigPad(box) {
  const cv = box.querySelector("canvas"), ctx = cv.getContext("2d"), ph = box.querySelector(".ph");
  const pad = { empty: true, data: () => pad.empty ? "" : cv.toDataURL("image/png") };
  const size = () => { const r = cv.getBoundingClientRect(), d = window.devicePixelRatio || 1; if (!r.width) return; const img = pad.empty ? null : cv.toDataURL(); cv.width = r.width * d; cv.height = r.height * d; ctx.setTransform(d, 0, 0, d, 0, 0); ctx.lineWidth = 2.2; ctx.lineCap = "round"; ctx.lineJoin = "round"; ctx.strokeStyle = "#2f251b"; if (img) { const i = new Image(); i.onload = () => ctx.drawImage(i, 0, 0, r.width, r.height); i.src = img } };
  size(); window.addEventListener("resize", size); pad.resize = size;
  let on = false, last = null; const pt = e => { const r = cv.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top } };
  cv.addEventListener("pointerdown", e => { on = true; last = pt(e); cv.setPointerCapture(e.pointerId); pad.empty = false; ph.style.display = "none" });
  cv.addEventListener("pointermove", e => { if (!on) return; const p = pt(e); ctx.beginPath(); ctx.moveTo(last.x, last.y); ctx.lineTo(p.x, p.y); ctx.stroke(); last = p });
  ["pointerup", "pointercancel", "pointerleave"].forEach(t => cv.addEventListener(t, () => on = false));
  box.querySelector(".clr").addEventListener("click", e => { e.preventDefault(); ctx.clearRect(0, 0, cv.width, cv.height); pad.empty = true; ph.style.display = "" });
  return pad;
}
async function copyText(t, btn) { try { await navigator.clipboard.writeText(t); const o = btn.textContent; btn.textContent = "Copied"; setTimeout(() => btn.textContent = o, 1200) } catch (e) { prompt("Copy this:", t) } }

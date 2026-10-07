// Midas Ipswich — Phone AI receptionist + messaging server (v2)
// Twilio ConversationRelay + Claude + Supabase. Runs on Railway.
import Fastify from "fastify";
import websocket from "@fastify/websocket";
import formbody from "@fastify/formbody";
import cors from "@fastify/cors";
import Anthropic from "@anthropic-ai/sdk";
import { createClient } from "@supabase/supabase-js";
import twilio from "twilio";
import crypto from "node:crypto";
import XLSX from "xlsx";

// ---------------------------------------------------------------- config
const env = process.env;
const PORT = Number(env.PORT || 8080);
const HOST = env.PUBLIC_HOST;
const SHOP_ID = Number(env.SHOP_ID || 1);
const MODEL = env.CLAUDE_MODEL || "claude-haiku-4-5-20251001";
const VOICE_ID = env.ELEVENLABS_VOICE_ID || "bnr31VMIPcsgqWJQh7Fs";
const VOICE_SETTINGS = env.VOICE_SETTINGS || "1.05_0.6_0.8"; // speed_stability_similarity
const STT_MODEL = env.STT_MODEL || "nova-3-general";       // or "flux"
const AUDIO_BASE = (env.AUDIO_BASE_URL || "").replace(/\/$/, "");
const SMS_FROM = env.SMS_FROM || env.TWILIO_NUMBER;
const APP_ORIGINS = (env.APP_ORIGINS ||
  "https://midasipswich.app,https://www.midasipswich.app,https://midasipswich.com.au,https://www.midasipswich.com.au").split(",");
const OFFSET_H = 10; // Australia/Brisbane, no daylight saving

const claude = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY });
const db = createClient(env.SUPABASE_URL, env.SUPABASE_SERVICE_KEY, { auth: { persistSession: false } });
const tw = twilio(env.TWILIO_SID, env.TWILIO_TOKEN);

// ---------------------------------------------------------------- helpers
const pad = (n) => String(n).padStart(2, "0");
const bne = () => new Date(Date.now() + OFFSET_H * 3600e3); // read with getUTC* = Brisbane wall time
const isoDay = (d) => d.toISOString().slice(0, 10);
const todayISO = () => isoDay(bne());
const addDays = (iso, n) => isoDay(new Date(Date.parse(iso + "T00:00:00Z") + n * 864e5));
const nowHHMM = () => { const d = bne(); return `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`; };
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const ord = (n) => n + ((n % 100 >= 11 && n % 100 <= 13) ? "th" : ({ 1: "st", 2: "nd", 3: "rd" }[n % 10] || "th"));
const longDate = (iso) => { const d = new Date(iso + "T00:00:00Z"); return `${DAYS[d.getUTCDay()]} ${ord(d.getUTCDate())} ${MONTHS[d.getUTCMonth()]}`; };
const time12 = (t) => { if (!t) return ""; const [h, m] = t.split(":").map(Number); return `${((h + 11) % 12) + 1}${m ? ":" + pad(m) : ""}${h < 12 ? "am" : "pm"}`; };

const normPhone = (p) => {
  const d = String(p || "").replace(/\D/g, "");
  if (!d) return null;
  if (/^61\d{9}$/.test(d)) return "0" + d.slice(2);
  if (/^[2-478]\d{8}$/.test(d)) return "0" + d;
  return d;
};
const toE164 = (p) => { const n = normPhone(p); return n && /^0\d{9}$/.test(n) ? "+61" + n.slice(1) : null; };
const titleCase = (s) => String(s || "").trim().toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
const cleanRego = (r) => String(r || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
const xml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const TWIML = (inner) => `<?xml version="1.0" encoding="UTF-8"?><Response>${inner}</Response>`;
const SAY_VOICE = "Google.en-AU-Neural2-C";

// ---------------------------------------------------------------- cached data
let settings = {}; let settingsAt = 0;
async function getSettings(force = false) {
  if (!force && Date.now() - settingsAt < 60e3) return settings;
  const { data, error } = await db.from("settings").select("*").order("id").limit(1).single();
  if (!error && data) { settings = data; settingsAt = Date.now(); }
  return settings;
}
let jobTypes = []; let jobTypesAt = 0;
async function getJobTypes() {
  if (Date.now() - jobTypesAt < 60e3 && jobTypes.length) return jobTypes;
  const { data } = await db.from("job_types").select("code,name,hours,keywords").eq("shop_id", SHOP_ID).eq("active", true).order("sort_order");
  if (data) { jobTypes = data; jobTypesAt = Date.now(); }
  return jobTypes;
}

// ---------------------------------------------------------------- capacity
async function getCapacity(from, to) {
  const { data, error } = await db.rpc("get_day_capacity", { p_from: from, p_to: to, p_shop: SHOP_ID });
  if (error) { console.error("capacity", error.message); return []; }
  return data.map((r) => ({ ...r, free: Math.max(0, Number(r.capacity_hours) - Number(r.booked_hours)) }));
}
const dayFits = (r, hours) => r && !r.blocked && r.free + 1e-9 >= hours && (r.max_bookings == null || r.booking_count < r.max_bookings);
function diaryText(cap) {
  return cap.map((r) => {
    const label = `${longDate(r.day)} (${r.day})`;
    if (r.blocked) return `${label}: CLOSED${r.block_reason && r.block_reason !== "Closed" ? " - " + r.block_reason : ""}`;
    const full = r.max_bookings != null && r.booking_count >= r.max_bookings;
    return `${label}: ${full ? "FULL" : r.free.toFixed(1) + " hours free"}`;
  }).join("\n");
}
function calendarText() {
  const t = todayISO();
  return Array.from({ length: 21 }, (_, i) => { const d = addDays(t, i); return `${longDate(d)} = ${d}`; }).join("\n");
}
async function findDays(jobCodes, preferred, after) {
  const jt = await getJobTypes();
  const picked = jt.filter((j) => (jobCodes || []).map((c) => String(c).toUpperCase()).includes(j.code));
  const hours = picked.reduce((a, j) => a + Number(j.hours), 0) || 1;
  const today = todayISO();
  const isDate = (d) => /^\d{4}-\d{2}-\d{2}$/.test(d || "");
  const start = isDate(after) ? addDays(after, 1) : isDate(preferred) && preferred > today ? preferred : today;
  const cap = await getCapacity(start < today ? today : start, addDays(start < today ? today : start, 45));
  const out = {};
  if (isDate(preferred) && !isDate(after)) {
    const r = cap.find((x) => x.day === preferred);
    if (preferred < today) out.preferred = { date: preferred, available: false, reason: "That date has passed." };
    else out.preferred = { date: preferred, day: longDate(preferred), available: dayFits(r, hours), reason: (r?.closed ?? r?.blocked) ? `Closed${r.block_reason && r.block_reason !== "Closed" ? " - " + r.block_reason : ""}` : dayFits(r, hours) ? null : "Fully booked" };
  }
  const s = await getSettings();
  const drop = (d) => dayHours(s, d).dropoff.map(time12).join(", ");
  if (out.preferred && out.preferred.available) out.preferred.dropoff_times = drop(preferred);
  out.next_available = cap.filter((x) => dayFits(x, hours) && x.day !== preferred).slice(0, 3).map((x) => ({ date: x.day, day: longDate(x.day), dropoff_times: drop(x.day) }));
  return out;
}

async function transferTargets() {
  const s = await getSettings();
  const { data, error } = await db.from("phone_contacts").select("number").eq("shop_id", SHOP_ID).eq("ring", true);
  const list = (!error && data ? data.map((x) => toE164(x.number)) : []).filter(Boolean);
  if (!list.length && s.ai_transfer_number) { const one = toE164(s.ai_transfer_number); if (one) list.push(one); }
  return [...new Set(list)];
}

// Opening hours for a date: per-weekday override (e.g. Saturday) or the normal day
const WD = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
function dayHours(s, iso) {
  const o = (s.booking_day_hours || {})[WD[new Date(iso + "T00:00:00Z").getUTCDay()]] || {};
  const dropoff = String(o.dropoff || s.booking_dropoff_times || "07:30,08:00,08:30").split(",").map((x) => x.trim()).filter(Boolean);
  return {
    start: o.start || s.booking_day_start || "07:30",
    end: o.end || s.booking_day_end || "17:30",
    dropoff,
    last: o.last || s.booking_last_time || "14:00",
    special: !!o.start,
  };
}
function hoursText(s) {
  const n = dayHours(s, "2026-10-05"); // a Monday = normal hours
  const lines = [`Monday to Friday: open ${time12(n.start)} to ${time12(n.end)}, drop-off from ${time12(n.dropoff[0])}, no bookings after ${time12(n.last)}.`];
  for (const [k, o] of Object.entries(s.booking_day_hours || {})) {
    const name = { sun: "Sunday", mon: "Monday", tue: "Tuesday", wed: "Wednesday", thu: "Thursday", fri: "Friday", sat: "Saturday" }[k];
    const drops = String(o.dropoff || "").split(",").map((x) => x.trim()).filter(Boolean);
    lines.push(`${name}: open ${time12(o.start)} to ${time12(o.end)}; drop-off must be between ${time12(drops[0] || o.start)} and ${time12(o.last || drops[drops.length - 1] || o.start)} so the work can be finished the same morning.`);
  }
  return lines.join("\n");
}

async function inHoursNow() {
  const s = await getSettings();
  const t = todayISO();
  const [r] = await getCapacity(t, t);
  if (!r || (r.closed ?? r.blocked)) return false;
  const now = nowHHMM();
  const h = dayHours(s, t);
  return now >= h.start && now < h.end;
}

async function isOpenNow() {
  const s = await getSettings();
  const t = todayISO();
  const [r] = await getCapacity(t, t);
  if (!r || (r.closed ?? r.blocked)) return false; // "no more bookings" days still put calls through
  const now = nowHHMM();
  const h = dayHours(s, t);
  if (!(now >= h.start && now < h.end)) return false;
  // part-day blocks marked "don't put calls through" (meetings, trainers)
  const { data: tb } = await db.from("time_blocks").select("start_time,end_time").eq("shop_id", SHOP_ID).eq("block_date", t).eq("no_calls", true);
  if ((tb || []).some((x) => now >= x.start_time && now < x.end_time)) return false;
  return true;
}

// When do we next open? e.g. "tomorrow, Wednesday, at 7:30am" / "on Monday at 7:30am" / "today at 7:30am"
async function nextOpenText() {
  const s = await getSettings();
  const t = todayISO(), now = nowHHMM();
  const cap = await getCapacity(t, addDays(t, 14)).catch(() => []);
  for (const r of cap) {
    if (r.closed ?? r.blocked) continue;
    const h = dayHours(s, r.day);
    if (r.day === t && now >= h.start) continue;
    const day = r.day === t ? "today" : r.day === addDays(t, 1) ? `tomorrow, ${DAYS[new Date(r.day + "T00:00:00Z").getUTCDay()]},` : `on ${longDate(r.day)}`;
    return `${day} at ${time12(h.start)}`;
  }
  return "";
}

// Blacklisted numbers: never put through, message only (entries can expire)
async function blacklisted(phone) {
  const n = normPhone(phone); if (!n) return null;
  const { data } = await db.from("phone_blacklist").select("name,reason,expires_on").eq("phone_norm", n).eq("active", true).maybeSingle();
  if (!data || (data.expires_on && data.expires_on < todayISO())) return null;
  return data;
}

// Standard price list for the phone AI
let priceList = []; let priceListAt = 0;
async function getPriceList() {
  if (Date.now() - priceListAt < 60e3) return priceList;
  const { data, error } = await db.from("price_list").select("item,price_from,price_to,includes,notes").eq("shop_id", SHOP_ID).eq("active", true).order("sort_order").order("item");
  if (!error) { priceList = data || []; priceListAt = Date.now(); }
  return priceList;
}
const money0 = (n) => "$" + Number(n).toLocaleString("en-AU", { minimumFractionDigits: Number(n) % 1 ? 2 : 0, maximumFractionDigits: 2 });
function priceListText(list) {
  if (!list.length) return "(no price list set – offer an estimate for any price question)";
  return list.map((p) => `- ${p.item}: ${p.price_from != null && p.price_to && Number(p.price_to) > Number(p.price_from) ? `${money0(p.price_from)} to ${money0(p.price_to)}` : p.price_from != null ? `from ${money0(p.price_from)}` : p.price_to ? `up to ${money0(p.price_to)}` : "price on request"}${p.includes ? `. Includes ${p.includes}` : ""}${p.notes ? `. ${p.notes}` : ""}`).join("\n");
}

// ---------------------------------------------------------------- customers
async function lookupCustomer(phone) {
  const n = normPhone(phone);
  if (!n) return null;
  const { data } = await db.from("customers")
    .select("id, given, surname, vehicles(rego, make, model, description)")
    .eq("shop_id", SHOP_ID).eq("phone_norm", n).maybeSingle();
  if (!data) return data;
  // has their car actually been in before? (a past booking that wasn't cancelled / a no-show) – and any booking coming up
  try {
    const t = todayISO();
    const { data: bks } = await db.from("bookings").select("booking_date,time_in,status,state,service")
      .eq("shop_id", SHOP_ID).eq("phone_norm", n).in("state", ["booking", "inbox"]).order("booking_date");
    const ok = (bks || []).filter((b) => !["cancelled", "no_show"].includes(b.status || ""));
    data.visits = ok.filter((b) => b.booking_date < t || (b.booking_date === t && b.status && b.status !== "booked")).length;
    if (!data.visits) {
      const { count } = await db.from("checkins").select("id", { count: "exact", head: true })
        .or(`phone.eq.${n},phone.eq.${n.slice(0, 4)} ${n.slice(4, 7)} ${n.slice(7)},phone.eq.+61${n.slice(1)},phone.eq.61${n.slice(1)}`);
      data.visits = count || 0;
    }
    data.upcoming = ok.find((b) => b.booking_date > t) || null;
    data.history = await invoiceHistory({ phone: n, regos: (data.vehicles || []).map((v) => v.rego) }, 3).catch(() => []);
    if (data.history.length && !data.visits) data.visits = data.history.length;
  } catch { data.visits = data.visits || 1; }
  return data;
}
// what is this caller most likely ringing about? (car with us today > estimate they said they'd call to book > estimate sent recently)
async function callerContext(phone) {
  const n = normPhone(phone);
  if (!n) return null;
  const today = todayISO();
  const { data: bks } = await db.from("bookings")
    .select("id, given, customer_name, phone, contact2_phone, vehicle, make, model, rego, time_in, time_out, status, state, checkin_id")
    .eq("shop_id", SHOP_ID).eq("booking_date", today).in("state", ["booking", "inbox"]);
  const bk = (bks || []).find((b) => !["cancelled", "no_show"].includes(b.status || "") && (normPhone(b.phone) === n || normPhone(b.contact2_phone) === n));
  if (bk) {
    let stage = "";
    if (bk.checkin_id) {
      const { data: ci } = await db.from("checkins").select("job_status, technician_assigned").eq("id", bk.checkin_id).maybeSingle();
      stage = ci?.job_status || "";
    }
    return {
      kind: "today", booking: bk, given: bk.given || String(bk.customer_name || "").split(/\s+/)[0] || "",
      car: bk.vehicle || [bk.make, bk.model].filter(Boolean).join(" ") || "",
      rego: bk.rego || "", stage: stage || (bk.checkin_id ? "Checked in" : "Not dropped off yet"),
    };
  }
  const since = new Date(Date.now() - 120 * 86400000).toISOString();
  const { data: qs } = await db.from("quotes").select("*").gte("created_at", since).not("phone", "is", null).order("created_at", { ascending: false }).limit(2000);
  const mine = (qs || []).filter((q) => normPhone(q.phone) === n && !q.accepted_booking_id && String(q.status || "") !== "Declined");
  const live = (q) => !q.valid_until || q.valid_until >= today;
  const q = mine.find((x) => x.accepted_at && live(x)) || mine.find((x) => !x.accepted_at && (x.sent_at || x.status === "Sent") && live(x));
  if (!q) return null;
  return { kind: "estimate", quote: q, given: firstName(q), car: q.vehicle || "", rego: q.rego || "", callToBook: !!q.accepted_at };
}
const ctxCar = (c) => (c?.car ? `your ${c.car}` : c?.rego ? `your car, rego ${String(c.rego).toUpperCase().split("").join(" ")}` : "your car");
function contextQuestion(c, name) {
  if (!c) return "";
  const who = name ? `Is this ${name}, calling` : "Is this";
  if (c.kind === "today") return `${who} about ${ctxCar(c)} that's booked in with us today?`;
  return `${who} about your estimate on ${c.car ? `the ${c.car}` : "your car"}?`;
}

async function upsertCustomer({ given, surname, mobile, rego, vehicle }) {
  const n = normPhone(mobile);
  if (!n) return null;
  let { data: c } = await db.from("customers").select("id").eq("shop_id", SHOP_ID).eq("phone_norm", n).maybeSingle();
  if (!c) {
    const ins = await db.from("customers")
      .insert({ shop_id: SHOP_ID, given: titleCase(given) || null, surname: titleCase(surname) || null, mobile: n, phone_norm: n })
      .select("id").single();
    c = ins.data;
  }
  if (c && rego) {
    const [make, ...model] = String(vehicle || "").trim().split(/\s+/);
    await db.from("vehicles").upsert(
      { customer_id: c.id, rego: cleanRego(rego), make: titleCase(make) || null, model: titleCase(model.join(" ")) || null, description: vehicle || null },
      { onConflict: "customer_id,rego", ignoreDuplicates: true });
  }
  return c?.id || null;
}

// ---------------------------------------------------------------- messaging
async function getTemplate(code) {
  const { data } = await db.from("sms_templates").select("*").eq("code", code).maybeSingle();
  return data;
}
const fill = (body, v) => String(body).replace(/\{(\w+)\}/g, (_, k) => (v[k] ?? ""));
function smsSegments(text) {
  const gsm = /^[\x20-\x7E\n\r£¥èéùìòÇØøÅåÆæßÉÄÖÑÜ§¿äöñüà]*$/.test(text);
  const len = text.length;
  return gsm ? (len <= 160 ? 1 : Math.ceil(len / 153)) : (len <= 70 ? 1 : Math.ceil(len / 67));
}
async function sendSms({ to, body, bookingId = null, customerId = null, template = null, sentBy = "system", withHeader = true }) {
  const s = await getSettings();
  const text = ((withHeader && s.sms_header ? `${s.sms_header} ` : "") + body).trim();
  const row = {
    phone: normPhone(to), direction: "outbound", message: text, template_label: template, sent_by: sentBy,
    status: "queued", channel: "sms", booking_id: bookingId, customer_id: customerId, segments: smsSegments(text),
  };
  const e164 = toE164(to);
  if (!e164 || !SMS_FROM) {
    row.status = "failed";
    row.error_message = !e164 ? "Invalid phone number" : "SMS_FROM not set";
    await db.from("sms_log").insert(row);
    return { ok: false, error: row.error_message };
  }
  const { data: log } = await db.from("sms_log").insert(row).select("id").single();
  try {
    const m = await tw.messages.create({ from: SMS_FROM, to: e164, body: text, statusCallback: `https://${HOST}/sms/status` });
    if (log) await db.from("sms_log").update({ provider_sid: m.sid, status: m.status }).eq("id", log.id);
    return { ok: true, sid: m.sid, id: log?.id };
  } catch (e) {
    console.error("SMS failed", e.message);
    if (log) await db.from("sms_log").update({ status: "failed", error_message: e.message }).eq("id", log.id);
    return { ok: false, error: e.message };
  }
}
function bookingVars(b, s) {
  return {
    first_name: titleCase(b.given || String(b.customer_name || "").split(" ")[0]),
    name: b.customer_name || "",
    rego: b.rego || "",
    date: b.booking_date ? longDate(b.booking_date) : "",
    time: time12(b.time_in),
    shop_name: s.name || "Midas Ipswich",
    shop_phone: s.phone || "",
    ref: b.booking_ref ? String(b.booking_ref) : "",
  };
}
async function sendBookingTemplate(booking, code, sentBy) {
  const t = await getTemplate(code);
  if (!t || !t.enabled) return { ok: false, error: `Template ${code} is missing or switched off` };
  const s = await getSettings();
  const r = await sendSms({ to: booking.phone, body: fill(t.body, bookingVars(booking, s)), bookingId: booking.id, customerId: booking.customer_id, template: code, sentBy });
  const stamp = new Date().toISOString();
  if (r.ok && code === "BOOKIN") await db.from("bookings").update({ bookin_sent_at: stamp }).eq("id", booking.id);
  if (r.ok && code === "REMIND") await db.from("bookings").update({ reminder_sent_at: stamp }).eq("id", booking.id);
  return r;
}
async function sendEmail({ to, subject, html, attachments }) {
  if (!env.RESEND_API_KEY || !env.EMAIL_FROM || !to) { console.warn("Email not configured / no recipient"); return { ok: false }; }
  const res = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({ from: env.EMAIL_FROM, to: String(to).split(/[,;\s]+/).filter(Boolean), subject, html, attachments }),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok) console.error("Email failed", j);
  await db.from("sms_log").insert({
    channel: "email", email: to, subject, direction: "outbound", sent_by: "system",
    message: html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").slice(0, 2000),
    status: res.ok ? "sent" : "failed", error_message: res.ok ? null : JSON.stringify(j).slice(0, 500), provider_sid: j.id || null,
  });
  return { ok: res.ok };
}

// ---------------------------------------------------------------- twilio webhook security
function twilioOk(req) {
  if (env.TWILIO_SKIP_VALIDATION === "1") return true;
  const sig = req.headers["x-twilio-signature"];
  if (!sig) return false;
  return twilio.validateRequest(env.TWILIO_TOKEN, sig, `https://${HOST}${req.url}`, req.body || {});
}

// ---------------------------------------------------------------- cached voice phrases
const PHRASES = {
  greeting: "Thanks for calling Midas Ipswich. Are you looking to make a booking, get an estimate, or speak with a member of our team?",
  transfer: "Certainly, I'll put you through to the team now.",
  voicemail: "I'm sorry, the team is unavailable at the moment. Please leave your name, number and a short message after the tone, and we'll get back to you.",
  quote_done: "Thank you. The team will contact you with an estimate shortly.",
  error: "I'm sorry, I'm having some trouble. I'll put you through to the team.",
  goodbye: "Thank you for calling Midas Ipswich. Goodbye.",
  voicemail_closed: "Please leave your name, number and a short message after the tone, and the team will get back to you when we open.",
};
const partOfDay = () => { const h = bne().getUTCHours(); return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening"; };
function greetingFor(cust, open = true, nextOpen = "", ctx = null) {
  if (ctx) {
    const closed = open === false ? ` We're closed at the moment${nextOpen ? ` and open again ${nextOpen}` : ""}.` : "";
    return `${partOfDay()}, welcome to Midas Ipswich.${closed} ${contextQuestion(ctx, cust?.given || ctx.given)}`;
  }
  if (cust?.given) return `${partOfDay()}, welcome to Midas Ipswich. Am I speaking with ${cust.given}?`;
  if (open === "busy") return `${partOfDay()}, thanks for calling Midas Ipswich. The team is tied up at the moment, but I can help you make a booking, get an estimate, or take a message for them.`;
  if (!open) return `${partOfDay()}, thanks for calling Midas Ipswich. We're closed at the moment${nextOpen ? ` and open again ${nextOpen}` : ""}. I can take a message for the team, or help you make a booking.`;
  return `${partOfDay()}, thanks for calling Midas Ipswich. Are you looking to make a booking, get an estimate, or speak with a member of our team?`;
}
const audioUrl = (k) => (AUDIO_BASE ? `${AUDIO_BASE}/${k}.mp3` : null);

// ---------------------------------------------------------------- AI prompt + tools
function buildSystem(s, jt, cap, cust, from, prices = [], nextOpen = "", ctx = null) {
  const shop = s.name || "Midas Ipswich";
  const lastTime = s.booking_last_time || "14:00";
  const staticText = `You are the phone receptionist for ${shop}, a car service and repair workshop${s.address ? ` at ${s.address}` : ""}.

TONE
- Professional, warm and clear. Plain Australian English with no slang: never say "G'day", "mate", "no worries" or "reckon".
- Always use the caller's name exactly as they give it. Never shorten it (Matthew stays Matthew).

YOU HANDLE THREE THINGS. Keep them separate and finish one before starting another.

1. BOOKINGS
- Collect, one at a time: name, mobile, rego, make and model, and what the vehicle needs.
- The rego is optional. If they don't know it, say "That's fine, you can give it to us when you drop the car off" and carry on. Never ask for it twice or hold up the booking for it.
- Ask what day suits them. Don't suggest a day first. Then call find_day with their preferred date. If it's available, use it. If not, say "The next available day is ..." using the first date find_day returns.
- If they turn a day down and ask for the next one, call find_day again with after_date set to the day they turned down. Never work out dates yourself; always use find_day.
- How we work: vehicles are dropped off in the morning and collected when we call to say they're ready. Encourage them to leave the vehicle with us, and ask what time they need it back.
- OPENING HOURS AND DROP-OFF:
${hoursText(s)}
- find_day tells you the allowed drop-off times for each date. Only offer those times.
- Only book a wait-while-you-wait if they insist, at an agreed time no later than ${time12(lastTime)}.
- We don't take drop-offs or bookings after ${time12(lastTime)}. If they need a later time, offer to put them through to the team to arrange it.
- They're welcome to drop off the night before or early. Our key drop box is on the wall of the building, beside the air conditioning unit, and has our street number, 325, on it. Mention this when they want to drop off early or the night before.
- Never tell the caller how long the work will take, and never mention hours.
- Before saving, read back the day, the drop-off, when they need it back, the rego (if they gave one) and the mobile, and get a yes. Then call book_job and tell them what the result says.

2. ESTIMATES
- Always call it an "estimate". Never say the word "quote" to the caller.
- Collect name, mobile, make and model, the rego if they know it, and what they'd like priced, then call request_quote. It plays its own confirmation.

3. SPEAK TO THE TEAM
- If they ask for a person, are upset, it's urgent, or you can't help, call transfer_to_staff straight away without saying anything first. It plays its own message.

PRICING
- Only give prices that are on the PRICE LIST (in the call details below). Never make up or estimate a price.
- Say "from" prices as "from" and ranges as a range, and mention what's included. Prices can vary by vehicle; we'll contact them before doing anything extra.
- If what they ask about isn't on the list, offer to arrange an estimate.

CUSTOM EXHAUST WORK
- We don't do custom exhaust work: custom or performance exhausts, fabrication, pipe bending, extractors or modified systems.
- If a caller asks for that, politely tell them we don't do custom exhaust work and that an exhaust specialist is their best bet, then call text_no_custom_exhaust so they get it in writing. Don't book it or take an estimate request for it.
- Fixing their existing standard exhaust (a blowing muffler, a cracked pipe, a broken hanger) is normal work – handle it like any other booking or estimate.

ESSENTIAL SERVICE OR LOGBOOK SERVICE (if asked)
- A logbook service follows the manufacturer's logbook for that service interval. We can only do exactly what's listed in the book, otherwise we can't stamp it.
- An essential service covers the oil and filter, labour, the environmental levy and an inspection of the vehicle.
- A logbook service can cost more than an essential service, but it needs to be done to the logbook to maintain the manufacturer's warranty.
- If asked: yes, we can complete and stamp logbook services to maintain the manufacturer's warranty.

PRANK OR ABUSIVE CALLERS
- Callers are often near a TV, radio or other people. If what you hear clearly isn't said to you (random words, TV dialogue, a conversation that has nothing to do with the call), call ignore_background and say nothing.
- Unclear or garbled words are usually a bad phone line or background noise, not swearing. If you can't make sense of what they said, say "Sorry, I didn't quite catch that. Could you say that again?"
- Only if the caller clearly swears at you or is obviously making a prank call, say exactly "I'm happy to help, but please keep the language polite, otherwise I'll have to end the call." and carry on.
- If they do it again, call end_call with abusive set to true. Don't say anything else.
- Someone venting about their car ("the bloody thing won't start") is not abuse; keep helping them.

RULES
- This is a live phone call. Reply in one or two short sentences, then stop and let them talk. No lists, symbols or abbreviations.
- Ask one question at a time.
- Ask for the mobile number in one go, then read it back in groups, like oh-four-one-two, three-four-five, six-seven-eight.
- Read regos back letter by letter to confirm.
- If an answer sounds cut off or unfinished, say "Sorry, please go on." and wait.
- Say dates as the weekday and date, like "Monday the sixteenth". Only offer dates that find_day says are available.
- Never diagnose faults or guess what's wrong.
- Keypad: "(pressed 1)" means booking, "(pressed 2)" estimate, "(pressed 3)" speak to the team.
- When the caller has nothing else, call end_call without saying goodbye yourself. It plays its own goodbye.`;

  const jtText = jt.map((j) => `${j.code}: ${j.name}, ${j.hours} hours${j.keywords ? ` (e.g. ${j.keywords})` : ""}`).join("\n");
  let callerText;
  if (cust?.given) {
    const name = [cust.given, cust.surname].filter(Boolean).join(" ");
    const cars = (cust.vehicles || []).map((v) => `${v.rego} ${v.description || [v.make, v.model].filter(Boolean).join(" ")}`.trim()).join("; ");
    callerText = `The caller has just been asked "Am I speaking with ${cust.given}?" because we have ${name} on file for this number${cars ? `, with vehicles: ${cars}` : ""}.
${cust.visits > 0
  ? `- They've been in to see us before. If they say yes: say "Thanks for coming back to us." and ask whether they'd like to make a booking, get an estimate, or speak with the team (if the shop is closed, say it's currently closed and offer a booking, an estimate, or to take a message instead).`
  : `- Their car has NOT been in to us yet – they've only been in touch before (a booking, enquiry or estimate). Never say "welcome back", "thanks for coming back" or call them a returning customer. If they say yes: say "Thanks ${cust.given}." and ask how you can help.`}
${cust.upcoming ? `- They already have a booking coming up: ${longDate(cust.upcoming.booking_date)}, drop-off ${time12(cust.upcoming.time_in || "08:00")}${cust.upcoming.service ? `, for ${cust.upcoming.service}` : ""}. If it fits, ask if they're calling about that booking. To change or cancel it, use transfer_to_staff.` : ""}
- You don't need their name or mobile again, and can ask if it's for the vehicle on file.
${cust.history?.length ? `SERVICE HISTORY from their invoices (only use once they've confirmed who they are):
${historyText(cust.history)}
- If they ask what was done last time, what's due, or how many kms were on it, answer from this. Never read out prices from old invoices.
- When they book, if any RECOMMENDED work above hasn't been done since, mention it briefly and offer to look at it on the day (add it to the booking description if they say yes).` : ""}
- If they say no: say "Sorry, we had that name attached to this phone number. Who am I speaking with?" Then treat them as a new caller and never mention the name or vehicles on file.`;
  } else {
    const n = normPhone(from);
    callerText = `The caller has just heard a greeting asking if they'd like to make a booking, get an estimate, or speak with the team. ` +
      (n && /^04/.test(n)
        ? `Caller ID is mobile ${n}. Ask if that's the best number to reach them rather than asking them to recite it.`
        : `Caller ID: ${n || "withheld"}. You'll need to ask for their mobile.`);
  }
  if (ctx) {
    const name = cust?.given || ctx.given;
    const asked = `${partOfDay()}, welcome to Midas Ipswich. ${contextQuestion(ctx, name)}`;
    const onFile = cust?.given ? `We have ${[cust.given, cust.surname].filter(Boolean).join(" ")} on file for this number. ` : "";
    if (ctx.kind === "today") {
      const b = ctx.booking;
      callerText = `The caller has just been greeted with: "${asked}" ${onFile}This number matches a booking TODAY: ${ctx.car || "vehicle not recorded"}${ctx.rego ? `, rego ${ctx.rego}` : ""}, drop-off ${time12(b.time_in || "08:00")}${b.time_out ? `, needed back ${b.time_out}` : ""}. Current job status on the workshop sheet: "${ctx.stage}".
- If they say yes, they're most likely after an update. Tell them the status in plain words: "Not dropped off yet" = we're expecting it in today; "Checked in" or waiting statuses = it's in the queue; in-progress statuses = the team is working on it now; anything with "ready" or "complete" = it's ready to pick up; "Authorisation Required" = the team needs to talk to them about some work, so put them through. Never guess a finish time, price or what's wrong with the car.
- For anything more (how long, what was found, cost, approving work, running late, changing the drop-off), use transfer_to_staff.
- If they say no or it's about something else: carry on as normal (booking, estimate, or the team). You already have their name and mobile.`;
    } else {
      const q = ctx.quote;
      const items = custLines(q).map((i) => `${i.description} ${$2(i.amount)}`).filter(Boolean).join("; ");
      callerText = `The caller has just been greeted with: "${asked}" ${onFile}This number has ESTIMATE #${q.id}${q.created_at ? ` from ${longDate(String(q.created_at).slice(0, 10))}` : ""}: ${items || "work as discussed"}. Total ${$2(qTotal(q))} including GST${q.valid_until ? `, valid until ${longDate(q.valid_until)}` : ""}. Vehicle: ${vehicleOf(q) || "not recorded"}.${ctx.callToBook ? " They have ALREADY ACCEPTED this estimate online and said they'd call to book it in." : " It has been sent to them but not accepted yet."}
- If they say yes: ${ctx.callToBook ? "thank them for accepting and" : "check they're happy to go ahead, then"} offer to book it in. Pick the matching JOB TYPES (use OTHER if none fit), use find_day, then book_job with estimate_id ${q.id} and the vehicle and rego above. You already have their name and mobile, so don't ask again.
- If they ask what's on it or the total, you can read out the items and the total above. For questions about the work itself, or to change the estimate, use transfer_to_staff.
- Always call it an "estimate", never a quote.
- If they say no or it's about something else: carry on as normal.`;
    }
  }
  const t0 = todayISO(); const r0 = cap.find((x) => x.day === t0); const n0 = nowHHMM();
  const h0 = dayHours(s, t0);
  const openNow = r0 && !(r0.closed ?? r0.blocked) && n0 >= h0.start && n0 < h0.end;
  const dynamicText = `TODAY: ${longDate(todayISO())} (${todayISO()}), time ${time12(nowHHMM())}.
TEAM AVAILABLE TO TAKE CALLS RIGHT NOW: ${openNow ? "yes" : `no. THE SHOP IS CLOSED (see OPENING HOURS).${nextOpen ? ` We open again ${nextOpen}; tell the caller this if they ask or want to speak to someone.` : ""}
- Never offer to put them through or to speak to someone, and never call transfer_to_staff.
- Offer to make a booking, arrange a quote, or take a message for the team. If they ask for a person, say "The shop is currently closed. Would you like me to take a message for the team?" If yes, call take_message without saying anything else; it plays its own instructions.
- Keypad "(pressed 3)" means take a message.`}`;
  const dynamicText2 = `
CALLER:
${callerText}
PRICE LIST:
${priceListText(prices)}
JOB TYPES (never mention the hours):
${jtText}
CALENDAR (for working out dates like "next Tuesday"; availability always comes from find_day):
${calendarText()}`;
  return [
    { type: "text", text: staticText, cache_control: { type: "ephemeral" } },
    { type: "text", text: dynamicText + dynamicText2 },
  ];
}

const TOOLS = [
  {
    name: "find_day",
    description: "Check if a preferred date is available for the job, and get the next available dates. Use after_date to get the next available after a day the caller turned down.",
    input_schema: {
      type: "object",
      properties: {
        job_type_codes: { type: "array", items: { type: "string" }, description: "Codes from JOB TYPES" },
        preferred_date: { type: "string", description: "YYYY-MM-DD the caller asked for" },
        after_date: { type: "string", description: "YYYY-MM-DD; return the next available days after this date" },
      },
      required: ["job_type_codes"],
    },
  },
  {
    name: "book_job",
    description: "Save the booking once the caller has confirmed the day, drop-off time, rego and mobile.",
    input_schema: {
      type: "object",
      properties: {
        given_name: { type: "string" },
        surname: { type: "string" },
        mobile: { type: "string", description: "Caller's mobile number, digits only" },
        rego: { type: "string", description: "Leave empty if they don't know it" },
        vehicle: { type: "string", description: "Make and model, e.g. Ford Ranger" },
        job_type_codes: { type: "array", items: { type: "string" }, description: "Codes from JOB TYPES" },
        description: { type: "string", description: "What the caller said the car needs, in their words" },
        date: { type: "string", description: "YYYY-MM-DD from the DIARY" },
        dropoff_time: { type: "string", description: "HH:MM 24 hour, the morning drop-off or the agreed wait time" },
        night_before: { type: "boolean", description: "true if they'll drop it off the night before using the key box" },
        waiting: { type: "boolean", description: "true only if they insist on waiting while it's done" },
        needed_back: { type: "string", description: "When they need the vehicle back, e.g. 3pm, end of day" },
        estimate_id: { type: "integer", description: "Only when booking in the caller's estimate from CALLER: its number" },
      },
      required: ["given_name", "mobile", "job_type_codes", "date"],
    },
  },
  {
    name: "request_quote",
    description: "Log an estimate request for the team. Plays its own confirmation to the caller.",
    input_schema: {
      type: "object",
      properties: {
        given_name: { type: "string" },
        surname: { type: "string" },
        mobile: { type: "string" },
        rego: { type: "string" },
        vehicle: { type: "string" },
        work_requested: { type: "string" },
      },
      required: ["given_name", "mobile", "work_requested"],
    },
  },
  {
    name: "transfer_to_staff",
    description: "Put the caller through to staff, or to voicemail if nobody's available. Plays its own message, so say nothing before calling it.",
    input_schema: { type: "object", properties: { reason: { type: "string" } } },
  },
  {
    name: "ignore_background",
    description: "Use when what you heard is background noise (TV, radio, other people) and not the caller talking to you. Say nothing; you'll keep listening.",
    input_schema: { type: "object", properties: {} },
  },
  {
    name: "text_no_custom_exhaust",
    description: "After telling the caller we don't do custom exhaust work, call this to text them a note confirming it. Then tell them what the result says and ask if there's anything else.",
    input_schema: { type: "object", properties: { given_name: { type: "string", description: "Caller's first name if you know it" }, mobile: { type: "string", description: "Only if caller ID isn't a mobile and they gave one" } } },
  },
  {
    name: "take_message",
    description: "When the shop is closed: take the caller to voicemail to leave a message for the team. Plays its own instructions, so say nothing before calling it.",
    input_schema: { type: "object", properties: {} },
  },
  {
    name: "end_call",
    description: "Hang up when the caller has nothing else. Plays its own goodbye, so don't say goodbye yourself. Set abusive to true to hang up on a prank or abusive caller with no goodbye.",
    input_schema: { type: "object", properties: { abusive: { type: "boolean" } } },
    cache_control: { type: "ephemeral" },
  },
];

// Holds short or unfinished-sounding answers (half a phone number, "um", "and") for a moment
function looksUnfinished(t) {
  const s = t.toLowerCase().replace(/[.,!?]+$/, "").trim();
  if (/\b(um+|uh+|er+|and|so|it's|its|my|the|is|a|an|it)$/.test(s)) return true;
  const NUMW = ["zero", "oh", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "double", "triple"];
  const digits = (s.match(/\d/g) || []).length;
  const words = s.replace(/[^a-z\s]/g, " ").split(/\s+/).filter((w) => NUMW.includes(w)).length;
  const n = digits + words;
  return n >= 2 && n < 10 && (/\d$/.test(s) || NUMW.some((w) => s.endsWith(w)));
}

// ---------------------------------------------------------------- server
const app = Fastify({ logger: { level: "info" }, bodyLimit: 30 * 1024 * 1024 }); // allow head office report uploads
await app.register(cors, { origin: APP_ORIGINS, methods: ["GET", "POST"], allowedHeaders: ["Authorization", "Content-Type"] });
await app.register(formbody);
app.removeContentTypeParser("application/json");
app.addContentTypeParser("application/json", { parseAs: "string" }, (req, body, done) => {
  req.rawBody = body;
  try { done(null, body ? JSON.parse(body) : {}); } catch (e) { e.statusCode = 400; done(e, undefined); }
});
await app.register(websocket);

app.get("/health", async () => {
  const t = todayISO();
  const cap = await db.rpc("get_day_capacity", { p_from: t, p_to: addDays(t, 6), p_shop: SHOP_ID });
  const hol = await db.from("calendar_blocks").select("id", { count: "exact", head: true }).eq("source", "holiday");
  return {
    ok: !cap.error, model: MODEL, stt: STT_MODEL, voice: VOICE_ID, cachedAudio: !!AUDIO_BASE,
    smsFrom: SMS_FROM ? SMS_FROM.slice(0, 6) + "..." : null, email: !!(env.RESEND_API_KEY && env.EMAIL_FROM),
    time: `${t} ${nowHHMM()}`,
    database: cap.error ? `ERROR: ${cap.error.message} (check SUPABASE_SERVICE_KEY is the service_role / secret key)` : `ok, ${cap.data.length} diary days`,
    holidays: hol.error ? `ERROR: ${hol.error.message}` : hol.count,
  };
});

// Incoming call
app.post("/voice", async (req, reply) => {
  if (!twilioOk(req)) return reply.code(403).send("Forbidden");
  // Blacklisted callers: never put through – a short message, then the message bank
  try {
    const bl = await blacklisted(req.body?.From);
    if (bl) {
      db.from("voice_calls").insert({ shop_id: SHOP_ID, call_sid: req.body.CallSid, from_number: normPhone(req.body.From), started_at: new Date().toISOString(), outcome: "blacklist_message", transcript: `Blacklisted caller${bl.name ? ": " + bl.name : ""}${bl.reason ? " (" + bl.reason + ")" : ""}` }).then(() => {}, () => {});
      const line = "Thanks for calling Midas Ipswich. Please leave your name, number and a short message after the beep, and the team will get back to you.";
      return reply.type("text/xml").send(TWIML(
        `<Connect action="https://${HOST}/handoff">` +
        `<ConversationRelay url="wss://${HOST}/ws" language="en-AU" transcriptionProvider="Deepgram" speechModel="${STT_MODEL}"` +
        ` ttsProvider="ElevenLabs" voice="${VOICE_ID}-flash_v2_5-${VOICE_SETTINGS}" elevenlabsTextNormalization="on" interruptible="none"` +
        ` welcomeGreeting="${xml(line)}"><Parameter name="wl" value="record"/><Parameter name="wlms" value="${Math.round(1200 + line.split(/\s+/).length * 360)}"/>` +
        `</ConversationRelay></Connect>`));
    }
  } catch (e) { console.error("blacklist check", e); }
  // Whitelisted callers (suppliers etc.) skip the AI: straight to the desk when open, message bank when closed
  try {
    const from = normPhone(req.body?.From);
    if (from) {
      const { data: wl } = await db.from("phone_whitelist").select("name").eq("phone_norm", from).eq("active", true).maybeSingle();
      if (wl) {
        const open = await isOpenNow();
        const targets = open ? await transferTargets() : [];
        db.from("voice_calls").insert({ shop_id: SHOP_ID, call_sid: req.body.CallSid, from_number: req.body.From, started_at: new Date().toISOString(), outcome: targets.length ? "whitelist_put_through" : "whitelist_voicemail", transcript: `Whitelisted caller: ${wl.name}` }).then(() => {}, () => {});
        const first = String(wl.name || "").trim().split(/\s+/)[0];
        const hi = first ? `Hi ${first}, ` : "Hi, ";
        const inHrs = targets.length ? true : await inHoursNow().catch(() => false);
        const line = targets.length ? hi + "just putting you through to the team now."
          : hi + (inHrs ? "the team can't get to the phone right now." : "the team has finished for the day.") + " Please leave a message after the beep and we'll get back to you.";
        // Spoken by ConversationRelay so it's the same ElevenLabs voice as the AI; the server then hands off straight away
        return reply.type("text/xml").send(TWIML(
          `<Connect action="https://${HOST}/handoff">` +
          `<ConversationRelay url="wss://${HOST}/ws" language="en-AU" transcriptionProvider="Deepgram" speechModel="${STT_MODEL}"` +
          ` ttsProvider="ElevenLabs" voice="${VOICE_ID}-flash_v2_5-${VOICE_SETTINGS}" elevenlabsTextNormalization="on" interruptible="none"` +
          ` welcomeGreeting="${xml(line)}">` +
          `<Parameter name="wl" value="${targets.length ? "transfer" : "record"}"/><Parameter name="wlms" value="${Math.round(1200 + line.split(/\s+/).length * 360)}"/>` +
          `</ConversationRelay></Connect>`));
      }
    }
  } catch (e) { console.error("whitelist check", e); }
  const [cust, openNow, inHours, ctx] = await Promise.all([lookupCustomer(req.body?.From).catch(() => null), isOpenNow().catch(() => true), inHoursNow().catch(() => true), callerContext(req.body?.From).catch(() => null)]);
  const open = openNow ? true : inHours ? "busy" : false;
  const nextOpen = open === false ? await nextOpenText().catch(() => "") : "";
  const useCached = AUDIO_BASE && !cust?.given && !ctx && open === true;
  const greet = useCached ? "" : ` welcomeGreeting="${xml(greetingFor(cust, open, nextOpen, ctx))}"`;
  reply.type("text/xml").send(TWIML(
    `<Connect action="https://${HOST}/handoff">` +
    `<ConversationRelay url="wss://${HOST}/ws" language="en-AU"` +
    ` transcriptionProvider="Deepgram" speechModel="${STT_MODEL}"` +
    ` ttsProvider="ElevenLabs" voice="${VOICE_ID}-flash_v2_5-${VOICE_SETTINGS}" elevenlabsTextNormalization="on"` +
    ` interruptible="speech" interruptSensitivity="low" ignoreBackchannel="true" dtmfDetection="true"${greet} />` +
    `</Connect>`));
});

function voicemailTwiml(closed = false) {
  const key = closed ? "voicemail_closed" : "voicemail";
  const intro = audioUrl(key) ? `<Play>${audioUrl(key)}</Play>` : `<Say voice="${SAY_VOICE}">${xml(PHRASES[key])}</Say>`;
  return TWIML(
    intro +
    `<Record maxLength="180" timeout="6" playBeep="true" transcribe="true" transcribeCallback="https://${HOST}/voicemail" />` +
    `<Say voice="${SAY_VOICE}">Thanks, we'll be in touch. Bye.</Say>`);
}

// AI session finished: transfer, voicemail or hang up
app.post("/handoff", async (req, reply) => {
  if (!twilioOk(req)) return reply.code(403).send("Forbidden");
  let h = {};
  try { h = JSON.parse(req.body.HandoffData || "{}"); } catch {}
  const s = await getSettings();
  if (h.action === "transfer") {
    const open = await isOpenNow();
    const targets = open ? await transferTargets() : [];
    if (targets.length) {
      const ring = Number(s.ai_transfer_ring_seconds || 35);
      const callerId = toE164(req.body.From) || env.TWILIO_NUMBER;
      return reply.type("text/xml").send(TWIML(
        `<Dial timeout="${ring}" answerOnBridge="true" callerId="${xml(callerId)}" action="https://${HOST}/dial-done">` +
        targets.map((t) => `<Number>${xml(t)}</Number>`).join("") + `</Dial>`));
    }
    return reply.type("text/xml").send(voicemailTwiml(!open));
  }
  if (h.action === "transfer_direct") {
    const targets = await transferTargets();
    const ring = Number(s.ai_transfer_ring_seconds || 35);
    if (targets.length) return reply.type("text/xml").send(TWIML(
      `<Dial timeout="${ring}" answerOnBridge="true" callerId="${xml(toE164(req.body.From) || env.TWILIO_NUMBER)}" action="https://${HOST}/dial-done">` +
      targets.map((t) => `<Number>${xml(t)}</Number>`).join("") + `</Dial>`));
    return reply.type("text/xml").send(voicemailTwiml(false));
  }
  if (h.action === "record_only") return reply.type("text/xml").send(TWIML(
    `<Record maxLength="180" timeout="6" playBeep="true" transcribe="true" transcribeCallback="https://${HOST}/voicemail" /><Hangup/>`));
  if (h.action === "voicemail") return reply.type("text/xml").send(voicemailTwiml(!!h.closed));
  reply.type("text/xml").send(TWIML("<Hangup/>"));
});

// Staff didn't answer -> voicemail
app.post("/dial-done", async (req, reply) => {
  if (!twilioOk(req)) return reply.code(403).send("Forbidden");
  const talk = Number(req.body.DialCallDuration) || 0;
  db.from("voice_calls").update({ transfer_seconds: talk }).eq("call_sid", req.body.CallSid).then(() => {}, () => {});
  if (req.body.DialCallStatus === "completed") {
    db.from("voice_calls").update({ outcome: "answered" }).eq("call_sid", req.body.CallSid).not("outcome", "like", "whitelist%").then(() => {}, () => {});
    return reply.type("text/xml").send(TWIML("<Hangup/>"));
  }
  reply.type("text/xml").send(voicemailTwiml());
});

// Voicemail transcribed -> store, log, email
app.post("/voicemail", async (req, reply) => {
  if (!twilioOk(req)) return reply.code(403).send("Forbidden");
  reply.send("ok");
  handleVoicemail(req.body).catch((e) => console.error("voicemail", e));
});
async function handleVoicemail(b) {
  const s = await getSettings(true);
  const from = normPhone(b.From);
  const cust = await lookupCustomer(b.From);
  const callerName = cust ? [cust.given, cust.surname].filter(Boolean).join(" ") : "";
  const transcript = b.TranscriptionStatus === "completed" && b.TranscriptionText ? b.TranscriptionText : "(transcription unavailable - listen to the recording)";
  let audio = null; let path = null;
  try {
    const r = await fetch(`${b.RecordingUrl}.mp3`, {
      headers: { Authorization: "Basic " + Buffer.from(`${env.TWILIO_SID}:${env.TWILIO_TOKEN}`).toString("base64") },
    });
    if (r.ok) {
      audio = Buffer.from(await r.arrayBuffer());
      path = `${todayISO()}/${b.RecordingSid || b.CallSid}.mp3`;
      const up = await db.storage.from("voicemails").upload(path, audio, { contentType: "audio/mpeg", upsert: true });
      if (up.error) { console.error("voicemail upload", up.error.message); path = null; }
    }
  } catch (e) { console.error("recording download", e.message); }

  const { error } = await db.from("phone_messages").insert({
    for_staff_name: "All staff", caller_name: callerName || "Unknown caller", caller_phone: from || "withheld", regarding: transcript, transcript,
    recording_path: path, source: "phone_ai_voicemail", call_sid: b.CallSid, taken_by: "Phone AI", status: "new",
  });
  if (error) console.error("phone_messages insert", error.message);
  await db.from("voice_calls").update({ outcome: "voicemail" }).eq("call_sid", b.CallSid).not("outcome", "like", "blacklist%").not("outcome", "like", "whitelist%");

  await sendEmail({
    to: s.notify_email,
    subject: `Voicemail from ${callerName || from || "unknown caller"}`,
    html: `<p><b>From:</b> ${esc(callerName)} ${esc(from || "withheld")}</p><p><b>Message:</b><br>${esc(transcript)}</p><p>${audio ? "Recording attached." : "Recording unavailable."}</p>`,
    attachments: audio ? [{ filename: "voicemail.mp3", content: audio.toString("base64") }] : undefined,
  });
}

// Inbound SMS (needs a two-way mobile number pointed here)
app.post("/sms/inbound", async (req, reply) => {
  if (!twilioOk(req)) return reply.code(403).send("Forbidden");
  const { From, Body, MessageSid } = req.body;
  const n = normPhone(From);
  const cust = await lookupCustomer(From);
  const { data: bk } = await db.from("bookings").select("id")
    .eq("phone_norm", n).gte("booking_date", addDays(todayISO(), -14))
    .order("booking_date", { ascending: false }).limit(1).maybeSingle();
  await db.from("sms_log").insert({
    phone: n, direction: "inbound", message: Body, status: "received", provider_sid: MessageSid,
    channel: "sms", customer_id: cust?.id || null, booking_id: bk?.id || null, sent_by: "customer",
  });
  if (bk) await db.from("bookings").update({ unread_reply: true }).eq("id", bk.id);
  if (/^\s*(stop|unsubscribe|opt ?out)\s*$/i.test(Body || "")) {
    await db.from("sms_opt_outs").insert({ phone: n, opted_out_at: new Date().toISOString(), note: "Replied STOP" });
    if (cust) await db.from("customers").update({ marketing_sms_optout: true }).eq("id", cust.id);
  }
  reply.type("text/xml").send("<Response></Response>");
});

// SMS delivery status
app.post("/sms/status", async (req, reply) => {
  if (!twilioOk(req)) return reply.code(403).send("Forbidden");
  const { MessageSid, MessageStatus, ErrorMessage, ErrorCode } = req.body;
  const upd = { status: MessageStatus };
  if (MessageStatus === "delivered") upd.delivered_at = new Date().toISOString();
  const WHY = { 30003: "Phone off or out of range", 30004: "Blocked by the recipient", 30005: "Number doesn't exist", 30006: "Landline or unreachable number", 30007: "Blocked by the phone carrier's spam filter (often caused by a link)", 30008: "Unknown carrier error", 30034: "Sender not registered" };
  if (ErrorCode || ErrorMessage) upd.error_message = [ErrorCode ? `Error ${ErrorCode}` : "", WHY[ErrorCode] || ErrorMessage || ""].filter(Boolean).join(" – ");
  await db.from("sms_log").update(upd).eq("provider_sid", MessageSid);
  reply.send("ok");
});

// ---------------------------------------------------------------- app API (signed-in admins)
async function requireAdmin(req, reply) {
  const token = (req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  if (!token) { reply.code(401).send({ error: "Not signed in" }); return null; }
  const { data, error } = await db.auth.getUser(token);
  if (error || !data?.user) { reply.code(401).send({ error: "Session expired" }); return null; }
  const { data: role } = await db.from("staff_roles").select("role, display_name, suspended").eq("user_id", data.user.id).maybeSingle();
  if (!role || role.suspended || !["admin", "super_admin", "god"].includes(role.role)) { reply.code(403).send({ error: "Not allowed" }); return null; }
  return { user: data.user, name: role.display_name || data.user.email, role: role.role };
}

// any signed-in staff member (not suspended) – used for sending quotes
async function requireStaffUser(req, reply) {
  const token = (req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  if (!token) { reply.code(401).send({ error: "Not signed in" }); return null; }
  const { data, error } = await db.auth.getUser(token);
  if (error || !data?.user) { reply.code(401).send({ error: "Session expired" }); return null; }
  const { data: role } = await db.from("staff_roles").select("role, display_name, suspended").eq("user_id", data.user.id).maybeSingle();
  if (!role || role.suspended) { reply.code(403).send({ error: "Not allowed" }); return null; }
  return { user: data.user, name: role.display_name || data.user.email, role: role.role };
}

// ---------------------------------------------------------------- simple markdown -> email-safe HTML
function mdToHtml(md) {
  const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const inline = (t) => esc(t).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/(^|\W)\*(?!\s)(.+?)\*(?=\W|$)/g, "$1<i>$2</i>");
  const lines = String(md || "").replace(/\r/g, "").split("\n");
  let html = "", list = null;
  const close = () => { if (list) { html += `</${list}>`; list = null; } };
  for (const raw of lines) {
    const l = raw.trimEnd();
    let m;
    if (!l.trim()) { close(); continue; }
    if ((m = l.match(/^#{1,2}\s+(.*)/))) { close(); html += `<h2 style="font:700 19px Arial,sans-serif;color:#111;margin:22px 0 8px;padding:6px 10px;background:#ffd400;border-radius:4px">${inline(m[1])}</h2>`; continue; }
    if ((m = l.match(/^#{3,6}\s+(.*)/))) { close(); html += `<h3 style="font:700 16px Arial,sans-serif;color:#da291c;margin:16px 0 6px">${inline(m[1])}</h3>`; continue; }
    if ((m = l.match(/^\s*[-*•]\s+(.*)/))) { if (list !== "ul") { close(); html += '<ul style="margin:4px 0 10px 0;padding-left:22px">'; list = "ul"; } html += `<li style="margin:3px 0">${inline(m[1])}</li>`; continue; }
    if ((m = l.match(/^\s*\d+[.)]\s+(.*)/))) { if (list !== "ol") { close(); html += '<ol style="margin:4px 0 10px 0;padding-left:22px">'; list = "ol"; } html += `<li style="margin:4px 0">${inline(m[1])}</li>`; continue; }
    close();
    html += `<p style="margin:6px 0">${inline(l)}</p>`;
  }
  close();
  return html;
}
function reportEmail(title, subtitle, bodyMd) {
  // Midas brand: yellow band + red rule, logo, Impact-style headings with the red full stop, black footer
  const body = mdToHtml(bodyMd)
    .replace(/<h2 style="[^"]*">/g, '<h2 style="font:700 22px Impact,\'Arial Narrow Bold\',\'Arial Narrow\',Arial,sans-serif;text-transform:uppercase;letter-spacing:.5px;color:#000;margin:24px 0 8px;padding:0 0 4px;border-bottom:3px solid #FFEB00">')
    .replace(/<\/h2>/g, '<span style="color:#DA291C">.</span></h2>')
    .replace(/<h3 style="[^"]*">/g, '<h3 style="font:700 15px Arial,sans-serif;color:#DA291C;margin:16px 0 6px;text-transform:uppercase;letter-spacing:.04em">');
  return `<div style="background:#f3f3f3;padding:16px 0"><div style="max-width:720px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;font:14px/1.55 Arial,sans-serif;color:#222">
    <div style="background:#FFEB00;border-bottom:6px solid #DA291C;padding:12px 20px"><table role="presentation" style="width:100%;border-collapse:collapse"><tr>
      <td style="width:120px"><img src="${PUBLIC_BASE}/brand/midas-logo.png" alt="Midas" width="110" style="display:block;width:110px;height:auto"></td>
      <td style="font:800 18px Arial,sans-serif;color:#000">Ipswich<span style="color:#DA291C">.</span><div style="font:700 11px Arial,sans-serif;letter-spacing:.08em;text-transform:uppercase">You can rely on us</div></td>
    </tr></table></div>
    <div style="padding:18px 20px 0;font:700 32px Impact,'Arial Narrow Bold','Arial Narrow',Arial,sans-serif;text-transform:uppercase;letter-spacing:.5px;line-height:1">${title}<span style="color:#DA291C">.</span></div>
    <div style="padding:6px 20px 0;color:#DA291C;font:700 13px Arial,sans-serif;text-transform:uppercase;letter-spacing:.05em">${subtitle}</div>
    <div style="padding:6px 20px 18px">${body}</div>
    <div style="background:#000;color:#fff;padding:14px 20px;font-size:12px"><b style="color:#FFEB00">Midas Ipswich</b> · Full charts and history: midasipswich.app → Bookings → Reports<div style="font:700 15px Impact,'Arial Narrow',Arial,sans-serif;text-transform:uppercase;color:#FFEB00;margin-top:6px;letter-spacing:.5px">You can rely on us<span style="color:#DA291C">.</span></div></div>
  </div></div>`;
}

// ---------------------------------------------------------------- head office / register report analysis
const REPORT_MODEL = env.REPORT_MODEL || "claude-sonnet-5-5";
const num = (v) => (v == null || v === "" || isNaN(Number(v)) ? null : Number(v));

// Turn files into Claude content blocks. files: [{ name, type, buf }]
function filesToContent(files) {
  const out = [];
  for (const f of files) {
    const name = f.name || "report";
    if (/pdf/i.test(f.type) || /\.pdf$/i.test(name)) out.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: f.buf.toString("base64") } });
    else if (/^image\//i.test(f.type)) out.push({ type: "image", source: { type: "base64", media_type: f.type, data: f.buf.toString("base64") } });
    else if (/\.(xlsx|xls)$/i.test(name) || /spreadsheet|excel/i.test(f.type)) {
      const wb = XLSX.read(f.buf, { type: "buffer" });
      out.push({ type: "text", text: `FILE ${name} (spreadsheet, converted to CSV):\n` + wb.SheetNames.map((n) => `### Sheet: ${n}\n` + XLSX.utils.sheet_to_csv(wb.Sheets[n])).join("\n\n").slice(0, 180000) });
    } else out.push({ type: "text", text: `FILE ${name}:\n` + f.buf.toString("utf8").slice(0, 180000) });
  }
  return out;
}


async function systemMetrics(days = 28) {
  const to = todayISO(), fromD = new Date(Date.now() - days * 864e5).toISOString();
  const from = fromD.slice(0, 10);
  const [calls, bks, cis, cap] = await Promise.all([
    db.from("voice_calls").select("outcome", { count: "exact" }).gte("started_at", fromD),
    db.from("bookings").select("state,status,source").gte("booking_date", from).lte("booking_date", to),
    db.from("checkins").select("id", { count: "exact", head: true }).gte("created_at", fromD),
    getCapacity(from, to).catch(() => []),
  ]);
  const c = calls.data || [], b = bks.data || [];
  const live = b.filter((x) => x.state === "booking" && !["cancelled", "no_show"].includes(x.status));
  const capH = cap.reduce((a, r) => a + Number(r.capacity_hours || 0), 0), useH = cap.reduce((a, r) => a + Number(r.booked_hours || 0), 0);
  return {
    period: `last ${days} days (${from} to ${to})`, calls: c.length, ai_bookings: c.filter((x) => x.outcome === "booking").length,
    put_through: c.filter((x) => x.outcome === "transfer").length, voicemails: c.filter((x) => x.outcome === "voicemail").length,
    bookings: live.length, phone_ai_bookings: live.filter((x) => x.source === "phone_ai").length, cancellations: b.filter((x) => x.status === "cancelled").length,
    no_shows: b.filter((x) => x.status === "no_show").length, kiosk_checkins: cis.count ?? null,
    capacity_used_pct: capH ? Math.round((useH / capH) * 1000) / 10 : null,
    note: "The phone AI and booking system only went live in early October 2026, so earlier weeks have little or no system data.",
  };
}

async function analyseReports({ files, extraText, period, question, metrics, uploadedBy, source, sender }) {
  const s = await getSettings();
  const { data: hist } = await db.from("ho_metrics").select("period_type,period_end,sales,invoices,avg_invoice,gp_pct,cust_per_day").order("period_end", { ascending: false }).limit(40);
  if (!metrics || !Object.keys(metrics).length) metrics = await systemMetrics().catch(() => ({}));
  const content = filesToContent(files);
  if (extraText) content.push({ type: "text", text: "EMAIL TEXT:\n" + String(extraText).slice(0, 20000) });
  content.push({ type: "text", text: `Our own numbers from our booking / phone / kiosk system (JSON): ${JSON.stringify(metrics || {}).slice(0, 20000)}
Head office figures we already have on file (most recent first): ${JSON.stringify(hist || [])}
Current average invoice set in our system: ${s.avg_invoice ? "$" + s.avg_invoice : "not set"}.
Period: ${period || "work it out from the report"}.
${question ? "Owner's question: " + question : ""}

These are reports for Midas Ipswich (automotive service & repair workshop, West Ipswich QLD, 3 technicians, Mon–Fri 7:30–5:30, Sat 8–12). They may be a head office Weekly Store Report, Monthly Store Report (incl. category mix: car service, brakes, suspension, drivetrain, tyres, other) or a daily sales report from the register.
Write it for a busy owner reading on a phone, in plain Australian English. Use EXACTLY this layout in Markdown:
## Headline
One or two sentences: the single most important thing this report says.
## Key numbers
- 4 to 6 bullets, each "Label: this period vs last year (± %)" – numbers only, no commentary
## Trend
- 2 to 4 short bullets comparing with the figures already on file
## Going well
- 2 to 3 short bullets
## Needs attention
- 2 to 4 short bullets
## Our system vs head office
- 2 to 3 short bullets comparing with our booking / call / check-in numbers (skip this section if we gave you no numbers)
## Actions this week
1. 3 to 5 numbered actions, most valuable first, each one line with the expected impact in brackets
Keep every bullet to one line (under 25 words). No long paragraphs, no tables, no extra sections. Never invent figures. Then, on the very last line, output one JSON object exactly like this (numbers only, no $ or %; null where not in the report):
EXTRACT {"report_type":"weekly_store|monthly_store|daily_sales|other","period":"text","average_invoice":n,"invoice_count":n,"total_sales":n,
"periods":[{"period_type":"week|month|day","period_end":"YYYY-MM-DD","label":"text","work_days":n,"sales":n,"sales_prior":n,"invoices":n,"invoices_prior":n,"cust_per_day":n,"cust_per_day_prior":n,"avg_invoice":n,"avg_invoice_prior":n,"gp_pct":n,"gp_pct_prior":n}],
"state_avg":{"sales":n,"invoices":n,"cust_per_day":n,"avg_invoice":n,"gp_pct":n},
"category":{"month_end":"YYYY-MM-DD","car_service":{"sales":n,"jobs":n},"brakes":{"sales":n,"jobs":n},"suspension":{"sales":n,"jobs":n},"drivetrain":{"sales":n,"jobs":n},"tyres":{"sales":n,"jobs":n},"other":{"sales":n,"jobs":n},"sales_mix_pct":{},"state_mix_pct":{}}}
Every row in a weekly or monthly table is its own entry in "periods" (week = week ending date, month = last day of the month).` });

  const res = await claude.messages.create({
    model: REPORT_MODEL, max_tokens: 8000,
    system: "You are a sharp, practical automotive workshop business analyst helping a Midas franchise owner understand head office and sales reports and grow the business.",
    messages: [{ role: "user", content }],
  });
  let analysis = res.content.filter((c) => c.type === "text").map((c) => c.text).join("\n").trim();
  let extracted = null;
  const i = analysis.lastIndexOf("EXTRACT");
  if (i >= 0) {
    const js = analysis.slice(i + 7).trim().replace(/^```(json)?|```$/g, "").trim();
    try { extracted = JSON.parse(js); } catch { /* ignore */ }
    analysis = analysis.slice(0, i).trim();
  }

  // file the numbers so trends build up over time
  if (extracted?.periods?.length) {
    const rows = extracted.periods.filter((p) => p.period_end && /^\d{4}-\d{2}-\d{2}$/.test(p.period_end)).map((p) => ({
      period_type: ["week", "month", "day"].includes(p.period_type) ? p.period_type : "month", period_end: p.period_end, label: p.label || null,
      work_days: num(p.work_days), sales: num(p.sales), sales_prior: num(p.sales_prior), invoices: num(p.invoices), invoices_prior: num(p.invoices_prior),
      cust_per_day: num(p.cust_per_day), cust_per_day_prior: num(p.cust_per_day_prior), avg_invoice: num(p.avg_invoice), avg_invoice_prior: num(p.avg_invoice_prior),
      gp_pct: num(p.gp_pct), gp_pct_prior: num(p.gp_pct_prior), source: extracted.report_type || source, updated_at: new Date().toISOString(),
    }));
    if (rows.length) await db.from("ho_metrics").upsert(rows, { onConflict: "period_type,period_end" });
  }
  if (extracted?.category?.month_end) await db.from("ho_categories").upsert({ month_end: extracted.category.month_end, data: extracted.category, state_avg: extracted.state_avg || null }, { onConflict: "month_end" });

  // keep the original files
  const paths = [];
  for (const f of files) {
    try {
      const fp = `${new Date().toISOString().slice(0, 10)}/${Date.now()}_${String(f.name || "report").replace(/[^\w.\-]+/g, "_")}`;
      const up = await db.storage.from("ho-reports").upload(fp, f.buf, { contentType: f.type || "application/octet-stream", upsert: false });
      if (!up.error) paths.push(fp);
    } catch { /* ignore */ }
  }
  const { data: row } = await db.from("ho_reports").insert({
    uploaded_by: uploadedBy, file_name: files.map((f) => f.name).join(", ") || null, period: period || extracted?.period || null,
    file_path: paths.join(", ") || null, extracted, analysis, model: REPORT_MODEL, source, sender: sender || null,
  }).select().single();
  return { analysis, extracted, id: row?.id };
}


// ---------------------------------------------------------------- register End Of Day reports (sales list, banking, exceptions)
const isEod = (name) => /end[\s_-]*of[\s_-]*day/i.test(name || "");
// daily reports are only read for the numbers (cheap model); the thinking happens once a week in weeklyEodReport()
const EOD_MODEL = env.EOD_MODEL || "claude-haiku-4-5-20251001";
async function ingestEod(files, uploadedBy, source) {
  const content = filesToContent(files);
  content.push({ type: "text", text: `These are End Of Day reports from the register (Marlin) at Midas Ipswich: the Customer Sales Transaction List and/or the Bank Transaction List for one register date. IGNORE any Exceptions Report completely (return "exceptions": []).
Note: banking is often entered the next day, so takings banked on a date can belong to the previous day's sales – don't treat a gap between sales and banking as a problem.
Return ONLY one JSON object, no other text:
{"register_date":"YYYY-MM-DD",
 "invoices":[{"invoice_no":"I026338","customer":"text","salesperson":"text","total_inc":n,"gst":n,"revenue":n,"gross_profit":n,"gpm":n}],
 "totals":{"invoices":n,"sales_inc":n,"gst":n,"revenue":n,"gross_profit":n,"gpm":n},
 "banking":{"cash":n,"cheque":n,"card":n,"eftpos":n,"direct_eft":n,"other":n,"total":n},
 "exceptions":[]}
Use null for anything not in these reports. Dates in the reports are Australian (DD/MM/YYYY).` });
  const res = await claude.messages.create({ model: EOD_MODEL, max_tokens: 8000, messages: [{ role: "user", content }] });
  const txt = res.content.filter((c) => c.type === "text").map((c) => c.text).join("").trim().replace(/^```(json)?|```$/g, "").trim();
  const j = JSON.parse(txt.slice(txt.indexOf("{"), txt.lastIndexOf("}") + 1));
  j.notes = [];
  const d = j.register_date;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d || "")) throw new Error("Couldn't find the register date in the report");

  const day = { register_date: d, updated_at: new Date().toISOString() };
  if (j.invoices?.length || j.totals?.invoices != null) {
    const t = j.totals || {};
    Object.assign(day, { invoices: num(t.invoices) ?? j.invoices.length, sales_inc: num(t.sales_inc), gst: num(t.gst), revenue: num(t.revenue), gross_profit: num(t.gross_profit), gpm: num(t.gpm) });
    if (j.invoices?.length) {
      await db.from("eod_invoices").upsert(j.invoices.filter((x) => x.invoice_no).map((x) => ({
        invoice_no: String(x.invoice_no), register_date: d, customer: x.customer || null, salesperson: x.salesperson || null,
        total_inc: num(x.total_inc), gst: num(x.gst), revenue: num(x.revenue), gross_profit: num(x.gross_profit), gpm: num(x.gpm),
      })), { onConflict: "invoice_no" });
    }
  }
  if (j.banking && Object.values(j.banking).some((v) => v != null)) day.banking = j.banking;
  if (j.notes?.length) day.notes = j.notes;
  await db.from("eod_days").upsert(day, { onConflict: "register_date" });
  // also feed the trend charts as a daily figure
  if (day.revenue != null || day.sales_inc != null) {
    const inv = day.invoices || null;
    await db.from("ho_metrics").upsert({
      period_type: "day", period_end: d, label: "Register " + d.split("-").reverse().join("/"), sales: day.revenue ?? day.sales_inc, invoices: inv,
      avg_invoice: inv ? Math.round((day.revenue ?? day.sales_inc) / inv) : null, gp_pct: day.gpm, source: "eod", updated_at: new Date().toISOString(),
    }, { onConflict: "period_type,period_end" });
  }
  const { data: full } = await db.from("eod_days").select("*").eq("register_date", d).single();
  const summary = [
    `## End of day ${d.split("-").reverse().join("/")}`,
    full?.invoices != null ? `- **Sales:** ${full.invoices} invoices · $${Number(full.sales_inc || 0).toLocaleString("en-AU")} inc GST · revenue $${Number(full.revenue || 0).toLocaleString("en-AU")} · GP $${Number(full.gross_profit || 0).toLocaleString("en-AU")} (${full.gpm ?? "–"}%)` : null,
    full?.banking?.total != null ? `- **Banked:** $${Number(full.banking.total).toLocaleString("en-AU")}` : null,
    j.notes?.length ? "### Notes" : null,
    ...(j.notes || []).map((n) => "- " + n),
  ].filter(Boolean).join("\n");
  await db.from("ho_reports").insert({ uploaded_by: uploadedBy, file_name: files.map((f) => f.name).join(", "), period: "End of day " + d, extracted: { report_type: "daily_sales", register_date: d }, analysis: summary, model: EOD_MODEL, source });
  return { analysis: summary, extracted: { report_type: "daily_sales", register_date: d } };
}


// ---------------------------------------------------------------- customer invoices (Megabus/Marlin tax invoices emailed to reports@…)
// Each invoice is read by Claude, filed against the customer + car, and kept as service history (phone AI + booking form).
const INVOICE_MODEL = env.INVOICE_MODEL || "claude-haiku-4-5-20251001";
const isInvoice = (name) => !isEod(name) && (/\bI\d{6,8}\b/i.test(name || "") || /tax[\s_-]*invoice|^invoice/i.test(name || ""));
async function ingestInvoice(file, source = "email") {
  const content = filesToContent([file]);
  content.push({ type: "text", text: `This is a customer tax invoice from Midas Ipswich (an automotive workshop). Read it carefully and return ONLY one JSON object, no other text:
{"invoice_no":"exactly as printed, e.g. I0026358","invoice_date":"YYYY-MM-DD","customer_name":"text","phone":"text","email":"text","address":"text",
 "rego":"text","vehicle":"year make model as printed","make":"text","model":"text","vehicle_year":"text","odometer":integer_kms,"vin":"text",
 "lines":[{"code":"product code","description":"text","qty":n,"unit_price":n,"total":n}],
 "guarantee":"the Midas guarantee / warranty wording printed on the invoice, word for word",
 "recommendations":["each piece of work recommended / advised / due next time, short and plain, e.g. 'Rear brake pads at 3mm – replace next service'"],
 "next_service":"next service due (kms and/or date) if printed",
 "total_inc":n,"gst":n,"revenue":n,"paid_by":"cash / card / eftpos / account …","salesperson":"text","technician":"text"}
RULES for "lines": ONLY the actual products and labour charged – rows with a product code and an amount (e.g. ".MO-0540 PROULT  PENRITE 5W40 PRO ULTRA  3.60  22.23  80.03", "/LMAINT  SERVICE LABOUR - MAINTENANCE", "/*WHEEL_ALIGN  SERVICE CHARGE - ALIGNMENT").
SKIP the service description / checklist text (lines starting with "~", e.g. from "Our Essential Minor Vehicle Service…" down to "Complete vehicle inspection report"), "Sub Total" rows and blank rows.
Odometer is the kms in, as printed. Use null for anything not on the invoice. Dates on the invoice are Australian (DD/MM/YYYY). Money as plain numbers.` });
  const res = await claude.messages.create({ model: INVOICE_MODEL, max_tokens: 4000, messages: [{ role: "user", content }] });
  const txt = res.content.filter((c) => c.type === "text").map((c) => c.text).join("").trim().replace(/^```(json)?|```$/g, "").trim();
  const j = JSON.parse(txt.slice(txt.indexOf("{"), txt.lastIndexOf("}") + 1));
  const no = String(j.invoice_no || "").trim().toUpperCase();
  if (!no) throw new Error("Couldn't find the invoice number");
  let fp = null;
  try {
    const p = `invoices/${no.replace(/[^\w-]/g, "_")}.pdf`;
    const up = await db.storage.from("ho-reports").upload(p, file.buf, { contentType: file.type || "application/pdf", upsert: true });
    if (!up.error) fp = p;
  } catch { /* ignore */ }
  const phone = normPhone(j.phone);
  const rego = cleanRego(j.rego) || null;
  let customerId = null;
  try {
    const [given, ...rest] = String(j.customer_name || "").trim().split(/\s+/);
    if (phone && /^04\d{8}$/.test(phone) && given && !/^cash/i.test(given)) customerId = await upsertCustomer({ given, surname: rest.join(" "), mobile: phone, rego, vehicle: j.vehicle });
    if (!customerId && rego) { const { data: v } = await db.from("vehicles").select("customer_id").eq("rego", rego).limit(1).maybeSingle(); customerId = v?.customer_id || null; }
  } catch { /* ignore */ }
  const row = {
    invoice_no: no, invoice_date: /^\d{4}-\d{2}-\d{2}$/.test(j.invoice_date || "") ? j.invoice_date : null,
    customer_name: j.customer_name || null, phone: j.phone || null, phone_norm: phone, email: j.email || null, address: j.address || null,
    rego, vehicle: j.vehicle || null, make: j.make || null, model: j.model || null, vehicle_year: j.vehicle_year ? String(j.vehicle_year) : null,
    odometer: Number.isFinite(Number(j.odometer)) && Number(j.odometer) > 0 ? Math.round(Number(j.odometer)) : null, vin: j.vin || null,
    lines: (Array.isArray(j.lines) ? j.lines : []).filter((l) => l && l.description && !/^\s*~/.test(l.description) && !/^sub\s*total/i.test(l.description) && (l.code || l.total != null)),
    notes: j.guarantee || null, recommendations: Array.isArray(j.recommendations) ? j.recommendations.filter(Boolean) : [],
    next_service: j.next_service || null, total_inc: num(j.total_inc), gst: num(j.gst), revenue: num(j.revenue) ?? (num(j.total_inc) != null && num(j.gst) != null ? Math.round((num(j.total_inc) - num(j.gst)) * 100) / 100 : null),
    paid_by: j.paid_by || null, salesperson: j.salesperson || null, technician: j.technician || null,
    customer_id: customerId, file_path: fp, source, raw: j, updated_at: new Date().toISOString(),
  };
  const { error } = await db.from("invoices").upsert(row, { onConflict: "invoice_no" });
  if (error) throw new Error(error.message);
  return row;
}
// last few invoices for a phone number and/or regos – for the phone AI and the booking form
async function invoiceHistory({ phone, regos = [] }, limit = 4) {
  const n = normPhone(phone);
  const ors = [];
  if (n) ors.push(`phone_norm.eq.${n}`);
  regos.map(cleanRego).filter(Boolean).slice(0, 6).forEach((r) => ors.push(`rego.eq.${r}`));
  if (!ors.length) return [];
  const { data } = await db.from("invoices").select("invoice_no,invoice_date,rego,vehicle,odometer,lines,notes,recommendations,next_service,total_inc")
    .or(ors.join(",")).order("invoice_date", { ascending: false }).limit(limit);
  return data || [];
}
function historyText(list) {
  return list.map((v) => {
    const work = (v.lines || []).map((l) => l.description).filter(Boolean).slice(0, 8).join(", ");
    return `- ${v.invoice_date ? longDate(v.invoice_date) : "Date unknown"}${v.rego ? ` · ${v.rego}` : ""}${v.vehicle ? ` ${v.vehicle}` : ""}${v.odometer ? ` · ${v.odometer.toLocaleString("en-AU")} km` : ""}: ${work || "work not listed"}${v.recommendations?.length ? `. RECOMMENDED: ${v.recommendations.join("; ")}` : ""}${v.next_service ? `. Next service due: ${v.next_service}` : ""}`;
  }).join("\n");
}
// staff: history + a short-lived PDF link
app.post("/api/invoices/history", async (req, reply) => {
  const who = await requireStaffUser(req, reply); if (!who) return;
  const b = req.body || {};
  reply.send({ ok: true, invoices: await invoiceHistory({ phone: b.phone, regos: [b.rego, ...(b.regos || [])].filter(Boolean) }, 30) });
});
app.post("/api/invoices/pdf", async (req, reply) => {
  const who = await requireStaffUser(req, reply); if (!who) return;
  const { data: inv } = await db.from("invoices").select("file_path").eq("invoice_no", String(req.body?.invoice_no || "")).maybeSingle();
  if (!inv?.file_path) return reply.code(404).send({ error: "No PDF saved for that invoice" });
  const { data, error } = await db.storage.from("ho-reports").createSignedUrl(inv.file_path, 600);
  if (error) return reply.code(500).send({ error: error.message });
  reply.send({ ok: true, url: data.signedUrl });
});
// one-off: file any invoices that were emailed before this existed (they were saved as "head office reports")
async function backfillInvoices() {
  try {
    const { data: rows } = await db.from("ho_reports").select("id,file_name,file_path").eq("source", "email").not("file_path", "is", null).order("id", { ascending: false }).limit(200);
    for (const r of rows || []) {
      if (!isInvoice(r.file_name)) continue;
      for (const fp of String(r.file_path).split(/,\s*/)) {
        const name = fp.split("/").pop();
        if (!isInvoice(name)) continue;
        const m = name.match(/I\d{6,8}/i);
        if (m) { const { data: ex } = await db.from("invoices").select("invoice_no").eq("invoice_no", m[0].toUpperCase()).maybeSingle(); if (ex) continue; }
        const dl = await db.storage.from("ho-reports").download(fp);
        if (dl.error || !dl.data) continue;
        const inv = await ingestInvoice({ name, type: "application/pdf", buf: Buffer.from(await dl.data.arrayBuffer()) }, "email").catch((e) => { console.warn("invoice backfill", name, e.message); return null; });
        if (inv) { await db.from("ho_reports").delete().eq("id", r.id); console.log("filed invoice", inv.invoice_no); }
      }
    }
  } catch (e) { console.warn("invoice backfill", e.message); }
}
setTimeout(backfillInvoices, 20000);
// several end-of-day PDFs in one email are usually different days – read each one on its own
async function ingestEodMany(files, uploadedBy, source) {
  if (files.length < 2) return ingestEod(files, uploadedBy, source);
  const out = [];
  for (const f of files) { try { out.push(await ingestEod([f], uploadedBy, source)); } catch (e) { out.push({ analysis: `**${f.name}:** couldn't read it – ${e.message}`, extracted: {} }); } }
  const dates = out.map((r) => r.extracted?.register_date).filter(Boolean).sort();
  return { analysis: out.map((r) => r.analysis).join("\n\n"), extracted: { report_type: "daily_sales", register_date: dates[dates.length - 1] || todayISO(), dates } };
}

// upload from the app (God / Super Admin)
app.post("/api/reports/analyse", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  if (!["god", "super_admin"].includes(who.role)) return reply.code(403).send({ error: "Only God / Super Admin users can analyse reports" });
  const { period, metrics, question } = req.body || {};
  const list = req.body?.files?.length ? req.body.files : [req.body || {}];
  const files = list.filter((f) => f.data_base64 || f.text).map((f) => (f.data_base64
    ? { name: f.file_name, type: f.media_type, buf: Buffer.from(f.data_base64, "base64") }
    : { name: (f.file_name || "report").replace(/\.(xlsx|xls)$/i, ".csv"), type: "text/csv", buf: Buffer.from(String(f.text), "utf8") }));
  if (!files.length) return reply.code(400).send({ error: "No report attached" });
  try {
    const eodFiles = files.filter((f) => isEod(f.name)), other = files.filter((f) => !isEod(f.name));
    let r = null;
    if (eodFiles.length) r = await ingestEodMany(eodFiles, who.name, "upload");
    if (other.length) r = await analyseReports({ files: other, period, question, metrics, uploadedBy: who.name, source: "upload" });
    reply.send({ ok: true, ...r });
  } catch (err) {
    req.log.error(err);
    reply.code(500).send({ error: "AI analysis failed: " + (err?.message || err) });
  }
});

// ---------------------------------------------------------------- reports emailed to reports@in.midasipswich.com.au (Resend Inbound)
function resendSigOk(req) {
  const secret = env.RESEND_WEBHOOK_SECRET;
  if (!secret) return true; // not set yet – rely on the sender allow-list
  const id = req.headers["svix-id"], ts = req.headers["svix-timestamp"], sig = String(req.headers["svix-signature"] || "");
  if (!id || !ts || !req.rawBody) return false;
  if (Math.abs(Date.now() / 1000 - Number(ts)) > 600) return false;
  const key = Buffer.from(secret.replace(/^whsec_/, ""), "base64");
  const expected = crypto.createHmac("sha256", key).update(`${id}.${ts}.${req.rawBody}`).digest("base64");
  return sig.split(" ").some((p) => p.split(",")[1] === expected);
}
const resendGet = async (path) => {
  const r = await fetch("https://api.resend.com" + path, { headers: { Authorization: `Bearer ${env.RESEND_API_KEY}` } });
  if (!r.ok) throw new Error(`Resend ${path} → ${r.status}`);
  return r.json();
};

app.post("/email/inbound", async (req, reply) => {
  if (!resendSigOk(req)) return reply.code(401).send({ error: "bad signature" });
  const ev = req.body || {};
  reply.send({ ok: true }); // answer Resend straight away; do the work after
  if (ev.type !== "email.received") return;
  const d = ev.data || {};
  try {
    // midas.com.au booking / enquiry forms forwarded from the store mailbox → Inbox (handled by /api/inbound/resend too; the dedupe key stops doubles)
    if (/new booking request from midas|new submission from midas|store enquiry form/i.test(String(d.subject || ""))) {
      const mail = await resendGet(`/emails/receiving/${d.email_id}`);
      const text = mail.text || htmlToText(mail.html);
      const res = await saveOnlineRequest(parseMidasForm(mail.subject || d.subject, text), { emailId: d.email_id, rawText: text });
      if (!res.ok) console.warn("online request not saved:", res.error);
      return;
    }
    const s = await getSettings();
    const sender = String(d.from || "").toLowerCase().replace(/.*<|>.*/g, "").trim();
    const allowed = String(s.report_senders || "").toLowerCase().split(/[\s,;]+/).filter(Boolean);
    const ok = allowed.some((a) => sender === a || sender.endsWith("@" + a.replace(/^@/, "")) || sender.endsWith("." + a.replace(/^@/, "")));
    if (!ok) {
      await db.from("ho_reports").insert({ uploaded_by: "Email (ignored)", file_name: d.subject || null, sender, source: "email_rejected", analysis: `Ignored – ${sender} is not on the allowed senders list (Booking & Phone Settings → Reports email).` });
      return;
    }
    const [mail, atts] = await Promise.all([resendGet(`/emails/receiving/${d.email_id}`), resendGet(`/emails/receiving/${d.email_id}/attachments`)]);
    const files = [];
    for (const a of atts.data || []) {
      if (a.content_disposition === "inline" && /^image\//.test(a.content_type || "")) continue; // signatures / logos
      if ((a.size || 0) > 20 * 1024 * 1024) continue;
      const r = await fetch(a.download_url);
      if (r.ok) files.push({ name: a.filename, type: a.content_type, buf: Buffer.from(await r.arrayBuffer()) });
    }
    if (!files.length && !(mail.text || "").trim()) return;
    const invFiles = files.filter((f) => isInvoice(f.name));
    const eodFiles = files.filter((f) => isEod(f.name)), other = files.filter((f) => !isEod(f.name) && !isInvoice(f.name));
    for (const f of invFiles) {
      try { const inv = await ingestInvoice(f, "email"); console.log("filed invoice", inv.invoice_no, inv.customer_name); }
      catch (e) { console.error("invoice", f.name, e.message); await db.from("ho_reports").insert({ uploaded_by: "Email from " + sender, file_name: f.name, source: "email_error", analysis: "Couldn't read this invoice: " + e.message }).then(() => {}, () => {}); }
    }
    if (eodFiles.length) {
      await ingestEodMany(eodFiles, "Email from " + sender, "email"); // filed quietly – the summary comes in the Monday weekly report
    }
    if (!other.length) return;
    const result = await analyseReports({ files: other, extraText: mail.text || "", period: d.subject || "", uploadedBy: "Email from " + sender, source: "email", sender });
    const repTo2 = s.report_email_to || s.notify_email;
    if (repTo2 && s.ho_email_summary !== false) {
      const html = reportEmail("Head office report", (d.subject || "Report").replace(/</g, "&lt;"), result.analysis);
      await sendEmail({ to: repTo2, subject: `Report analysed: ${d.subject || "head office report"}`, html });
    }
  } catch (err) {
    console.error("inbound email failed", err);
    try { await db.from("ho_reports").insert({ uploaded_by: "Email", file_name: d.subject || null, source: "email_error", analysis: "Couldn't process this email: " + (err?.message || err) }); } catch { /* ignore */ }
  }
});

app.post("/api/sms/send", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const { to, body, booking_id, customer_id, template_code } = req.body || {};
  if (!to || !body) return reply.code(400).send({ error: "to and body are required" });
  reply.send(await sendSms({ to, body, bookingId: booking_id || null, customerId: customer_id || null, template: template_code || null, sentBy: who.name }));
});

app.post("/api/email/send", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const { to, subject, body } = req.body || {};
  if (!to || !subject || !body) return reply.code(400).send({ error: "to, subject and body are required" });
  reply.send(await sendEmail({ to, subject, html: esc(body).replace(/\n/g, "<br>") }));
});

app.post("/api/bookings/:id/accept", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const { data: b, error } = await db.from("bookings").update({ state: "booking" }).eq("id", req.params.id).select("*").single();
  if (error) return reply.code(400).send({ error: error.message });
  let sms = null;
  if (!b.bookin_sent_at && req.body?.sms !== false) sms = await sendBookingTemplate(b, "BOOKIN", who.name);
  reply.send({ ok: true, sms });
});

app.post("/api/bookings/:id/template", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const { code } = req.body || {};
  const { data: b } = await db.from("bookings").select("*").eq("id", req.params.id).single();
  if (!b) return reply.code(404).send({ error: "Booking not found" });
  reply.send(await sendBookingTemplate(b, String(code || "").toUpperCase(), who.name));
});

// Test ring the ticked "speak to the team" numbers
app.post("/api/phone/test", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const targets = await transferTargets();
  if (!targets.length) return reply.code(400).send({ error: "No numbers are ticked to ring" });
  const from = env.TWILIO_NUMBER;
  for (const to of targets) {
    await tw.calls.create({ to, from, twiml: TWIML(`<Pause length="1"/><Say voice="${SAY_VOICE}">This is a test call from the Midas Ipswich phone system. Customer calls will be put through to this number.</Say>`) });
  }
  reply.send({ ok: true, called: targets.map((t) => "0" + t.slice(3)) });
});

// One-off: generate cached voice phrases. Visit /admin/generate-audio?token=YOUR_ADMIN_TOKEN
app.get("/admin/generate-audio", async (req, reply) => {
  if (!env.ADMIN_TOKEN || req.query.token !== env.ADMIN_TOKEN) return reply.code(403).send({ error: "Forbidden" });
  if (!env.ELEVENLABS_API_KEY) return reply.code(400).send({ error: "ELEVENLABS_API_KEY not set in Railway" });
  const [speed, stability, similarity] = VOICE_SETTINGS.split("_").map(Number);
  const done = [];
  for (const [key, text] of Object.entries(PHRASES)) {
    const r = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${VOICE_ID}?output_format=mp3_44100_128`, {
      method: "POST",
      headers: { "xi-api-key": env.ELEVENLABS_API_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ text, model_id: "eleven_flash_v2_5", voice_settings: { stability, similarity_boost: similarity, speed } }),
    });
    if (!r.ok) return reply.code(500).send({ error: `${key}: ${await r.text()}`, done });
    const buf = Buffer.from(await r.arrayBuffer());
    const up = await db.storage.from("voice-audio").upload(`${key}.mp3`, buf, { contentType: "audio/mpeg", upsert: true });
    if (up.error) return reply.code(500).send({ error: up.error.message, done });
    done.push(key);
  }
  reply.send({ ok: true, done, next: "Add this Railway variable, then redeploy", AUDIO_BASE_URL: `${env.SUPABASE_URL}/storage/v1/object/public/voice-audio` });
});

// ---------------------------------------------------------------- live call (ConversationRelay websocket)
app.register(async (f) => {
  f.get("/ws", { websocket: true }, (socket) => {
    const call = {
      sid: null, from: "", system: null, messages: [], queue: [], busy: false, ended: false,
      holdTimer: null, held: "", outcome: null, bookingId: null, quoteId: null, ready: null,
    };
    const send = (o) => { if (socket.readyState === 1) socket.send(JSON.stringify(o)); };
    const speak = (key) => {
      const url = audioUrl(key);
      if (url) send({ type: "play", source: url, loop: 1, preemptible: false, interruptible: true });
      else send({ type: "text", token: PHRASES[key], last: true });
    };
    const end = (handoff, delayMs = 0) => {
      if (call.ended) return;
      call.ended = true;
      setTimeout(() => send({ type: "end", handoffData: JSON.stringify(handoff) }), delayMs);
    };

    async function setup(msg) {
      call.sid = msg.callSid;
      call.from = msg.from || "";
      const wlMode = msg.customParameters?.wl;
      if (wlMode) { // whitelisted caller: greeting already playing, then straight to the desk / message bank
        call.whitelist = true;
        end({ action: wlMode === "transfer" ? "transfer_direct" : "record_only" }, Number(msg.customParameters.wlms || 4000));
        return;
      }
      const today = todayISO();
      const [s, jt, cap, cust, prices, ctx] = await Promise.all([getSettings(), getJobTypes(), getCapacity(today, addDays(today, 20)), lookupCustomer(call.from), getPriceList().catch(() => []), callerContext(call.from).catch(() => null)]);
      call.ctx = ctx;
      call.openNow = await isOpenNow();
      call.inHours = await inHoursNow();
      const nextOpen = call.openNow ? "" : await nextOpenText().catch(() => "");
      if (AUDIO_BASE && !cust?.given && !ctx && call.openNow) speak("greeting");
      call.system = buildSystem(s, jt, cap, cust, call.from, prices, nextOpen, ctx);
      db.from("voice_calls").insert({ shop_id: SHOP_ID, call_sid: call.sid, from_number: normPhone(call.from), customer_id: cust?.id || null })
        .then(() => {}, () => {});
    }

    function addUser(text) {
      const last = call.messages[call.messages.length - 1];
      if (last && last.role === "user") {
        if (typeof last.content === "string") last.content = [{ type: "text", text: last.content }];
        last.content.push({ type: "text", text });
      } else call.messages.push({ role: "user", content: text });
    }

    // Caller cut the AI off: keep only what they actually heard
    function trimHeard(heard) {
      for (let i = call.messages.length - 1; i >= 0; i--) {
        const m = call.messages[i];
        if (m.role !== "assistant") continue;
        if (Array.isArray(m.content)) {
          const tb = m.content.find((c) => c.type === "text");
          if (tb) tb.text = String(heard || "").trim() || "...";
        }
        break;
      }
    }

    async function bookJob(inp) {
      const jt = await getJobTypes();
      const picked = jt.filter((j) => (inp.job_type_codes || []).map((c) => String(c).toUpperCase()).includes(j.code));
      const est = call.ctx?.kind === "estimate" && Number(inp.estimate_id) === Number(call.ctx.quote.id) ? call.ctx.quote : null;
      const hours = Math.max(picked.reduce((a, j) => a + Number(j.hours), 0) || 1, est ? await quoteHours(est) : 0);
      const date = String(inp.date || "");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: "Date must be YYYY-MM-DD from the diary." };
      if (date < todayISO()) return { ok: false, error: "That date has already passed." };
      const mobile = normPhone(inp.mobile || call.from);
      if (!mobile || mobile.length < 10) return { ok: false, error: "Need a valid mobile number." };

      // Re-check the diary right before saving
      const cap = await getCapacity(date, addDays(date, 14));
      if (!dayFits(cap[0], hours)) {
        const alternatives = cap.slice(1).filter((r) => dayFits(r, hours)).slice(0, 3).map((r) => `${longDate(r.day)} (${r.day})`);
        return { ok: false, error: cap[0]?.blocked ? "We're closed that day." : "That day has just filled up.", alternatives };
      }

      const s = await getSettings();
      const given = titleCase(inp.given_name);
      const surname = titleCase(inp.surname);
      const customerId = await upsertCustomer({ given, surname, mobile, rego: inp.rego, vehicle: inp.vehicle });
      const [make, ...model] = String(inp.vehicle || "").trim().split(/\s+/);
      const dh = dayHours(s, date);
      const time = /^\d{2}:\d{2}$/.test(inp.dropoff_time || "") ? inp.dropoff_time : dh.dropoff[0];
      if (time < dh.start) return { ok: false, error: `On that day we open at ${time12(dh.start)}. Drop-off times are ${dh.dropoff.map(time12).join(", ")}.` };
      if (time > dh.last) return { ok: false, error: dh.special ? `On that day drop-off must be by ${time12(dh.last)} so we can finish the work. Drop-off times are ${dh.dropoff.map(time12).join(", ")}.` : `We don't take bookings after ${time12(dh.last)} without special arrangement. Offer to put them through to the team.` };
      const confirmed = !!s.ai_auto_confirm;

      const { data: b, error } = await db.from("bookings").insert({
        shop_id: SHOP_ID, customer_id: customerId,
        customer_name: `${given} ${surname}`.trim(), given, surname, phone: mobile,
        rego: cleanRego(inp.rego) || null, rego_state: "QLD", vehicle: inp.vehicle || null,
        make: titleCase(make) || null, model: titleCase(model.join(" ")) || null,
        service: picked.map((j) => j.name).join(", ") || "Other", job_types: picked.map((j) => j.code),
        hours, duration_mins: Math.round(hours * 60),
        description: est ? `Accepted estimate #${est.id} (${$2(qTotal(est))} inc GST)${inp.description ? ` – ${inp.description}` : ""}` : inp.description || null,
        booking_date: date, time_in: time, time_out_mode: inp.waiting ? "waiting" : "all_day",
        time_out: inp.needed_back || null,
        start_time: `${date}T${time}:00+10:00`,
        state: confirmed ? "booking" : "inbox", status: "booked", source: "phone_ai",
        salesperson: "PHONE AI", call_sid: call.sid, created_by: "Phone AI",
        notes: [
          !cleanRego(inp.rego) ? "Rego to be provided at drop-off." : null,
          inp.night_before ? "Dropping off the night before (key box)." : null,
          inp.needed_back ? `Needed back: ${inp.needed_back}.` : null,
          `Booked by phone AI. Caller ID ${normPhone(call.from) || "withheld"}.`,
        ].filter(Boolean).join(" "),
      }).select("*").single();
      if (error) { console.error("booking insert", error.message); return { ok: false, error: "Couldn't save the booking. Offer to put them through to the team." }; }

      call.outcome = "booking";
      call.bookingId = b.id;
      if (est) db.from("quotes").update({ accepted_at: est.accepted_at || new Date().toISOString(), status: "Accepted", accepted_booking_id: b.id }).eq("id", est.id).then(() => {}, (e) => console.error("estimate link", e));
      const refSay = b.booking_ref ? ` Their booking number is ${String(b.booking_ref).split("").join(" ")}.` : "";
      if (confirmed) sendBookingTemplate(b, "BOOKIN", "Phone AI").catch((e) => console.error("bookin", e));
      return {
        ok: true,
        tell_caller: confirmed
          ? `Booked in for ${longDate(date)}${inp.waiting ? ` at ${time12(time)}` : inp.night_before ? ", dropping off the night before" : `, drop off from ${time12(time)}`}. A confirmation text is on its way.${refSay}`
          : `Booked in for ${longDate(date)}${inp.waiting ? ` at ${time12(time)}` : inp.night_before ? ", dropping off the night before" : `, drop off from ${time12(time)}`}. The team will send a text to confirm shortly.${refSay}`,
      };
    }

    async function requestQuote(inp) {
      const given = titleCase(inp.given_name);
      const surname = titleCase(inp.surname);
      const mobile = normPhone(inp.mobile || call.from);
      const { data: q, error } = await db.from("quotes").insert({
        given, surname, name: `${given} ${surname}`.trim(), phone: mobile,
        vehicle: inp.vehicle || null, rego: cleanRego(inp.rego) || null, rego_state: "QLD",
        status: "Pending", notes: `PHONE AI ESTIMATE REQUEST: ${inp.work_requested || ""}`,
        line_items: [], total: 0, source: "phone_ai", call_sid: call.sid, created_by_name: "Phone AI",
      }).select("id").single();
      if (error) { console.error("quote insert", error.message); return { ok: false, error: "Couldn't save the request. Offer to put them through to the team." }; }
      upsertCustomer({ given, surname, mobile, rego: inp.rego, vehicle: inp.vehicle }).catch(() => {});
      call.outcome = "quote";
      call.quoteId = q.id;
      speak("quote_done");
      return { ok: true, note: "The caller has just been told the team will contact them with an estimate shortly. Don't repeat that. Ask if there's anything else." };
    }

    async function runTool(name, input) {
      try {
        if (name === "find_day") return await findDays(input?.job_type_codes, input?.preferred_date, input?.after_date);
        if (name === "book_job") return await bookJob(input || {});
        if (name === "request_quote") return await requestQuote(input || {});
        if (name === "transfer_to_staff") {
          call.outcome = "transfer";
          const open = await isOpenNow();
          if (open && (await transferTargets()).length) { speak("transfer"); end({ action: "transfer", reason: input?.reason || "" }, 3500); }
          else end({ action: "voicemail", closed: !(await inHoursNow()) }, 300);
          return { __end: true };
        }
        if (name === "ignore_background") return { ok: true };
        if (name === "text_no_custom_exhaust") {
          const to = normPhone(input?.mobile) || normPhone(call.from);
          call.outcome = call.outcome || "no_custom_exhaust";
          if (!to || !/^04\d{8}$/.test(to)) return { ok: false, tell_caller: "Don't promise a text – there's no mobile number to send it to." };
          if (call.exhaustTexted) return { ok: true, tell_caller: "The text has already been sent." };
          const t = await getTemplate("NOEXHAUST");
          if (!t || !t.enabled) return { ok: false, tell_caller: "Texting is switched off for this – just confirm it verbally." };
          const s = await getSettings();
          const r = await sendSms({ to, body: fill(t.body, { first_name: titleCase(input?.given_name || "") || "there", shop_phone: s.phone || "07 3444 2200", shop_name: s.name || "Midas Ipswich" }), template: "NOEXHAUST", sentBy: "Phone AI" });
          if (!r.ok) return { ok: false, tell_caller: "The text didn't go through – just confirm it verbally." };
          call.exhaustTexted = true;
          return { ok: true, tell_caller: "I've sent you a text confirming that." };
        }
        if (name === "take_message") {
          call.outcome = "voicemail";
          end({ action: "voicemail", closed: !(await inHoursNow()) }, 200);
          return { __end: true };
        }
        if (name === "end_call") {
          if (input?.abusive) { call.outcome = "abusive"; end({ action: "hangup" }, 200); return { __end: true }; }
          speak("goodbye");
          end({ action: "hangup" }, 3500);
          return { __end: true };
        }
        return { ok: false, error: "Unknown tool" };
      } catch (e) {
        console.error("tool", name, e);
        return { ok: false, error: "System error. Offer to put them through to the team." };
      }
    }

    // One AI turn, streamed, including any tool calls. Returns true if the call is ending.
    async function runTurn() {
      for (let i = 0; i < 6; i++) {
        if (call.ended) return true;
        const stream = claude.messages.stream({ model: MODEL, max_tokens: 200, system: call.system, tools: TOOLS, messages: call.messages });
        stream.on("text", (t) => send({ type: "text", token: t, last: false }));
        const res = await stream.finalMessage();
        send({ type: "text", token: "", last: true });
        call.messages.push({ role: "assistant", content: res.content });

        const uses = res.content.filter((c) => c.type === "tool_use");
        if (!uses.length) return false;
        if (uses.every((u) => u.name === "ignore_background")) {
          call.messages.push({ role: "user", content: uses.map((u) => ({ type: "tool_result", tool_use_id: u.id, content: "Ignored. Keep listening." })) });
          return false; // stay silent and wait for the caller
        }
        const results = [];
        for (const u of uses) {
          if (call.ended) return true;
          const out = await runTool(u.name, u.input);
          if (out.__end) return true;
          results.push({ type: "tool_result", tool_use_id: u.id, content: JSON.stringify(out) });
        }
        call.messages.push({ role: "user", content: results });
      }
      return false;
    }

    async function pump() {
      if (call.busy) return;
      call.busy = true;
      try {
        if (call.ready) await call.ready;
        while (call.queue.length && !call.ended) {
          addUser(call.queue.splice(0).join(" "));
          if (await runTurn()) break;
        }
      } catch (e) {
        console.error("turn", e);
        speak("error");
        end({ action: "transfer", reason: "error" }, 4500);
      } finally {
        call.busy = false;
      }
    }
    const push = (text) => { call.queue.push(text); pump(); };

    // "operator", or asking to speak/talk to a real person – not "are you a real person?"
    const OPERATOR = /\boperator\b|\b(speak|talk|put me through)\b[^.?!]*\b(real person|a human|human being|actual person|live person|receptionist)\b/i;
    function enqueue(text, immediate) {
      if (call.ended) return;
      text = String(text || "").trim();
      if (!text) return;
      if (OPERATOR.test(text) || text === "(pressed 0)") { // straight through, no questions (closed rule still applies inside the tool)
        if (call.holdTimer) { clearTimeout(call.holdTimer); call.holdTimer = null; call.held = ""; }
        call.queue = [];
        if (!call.busy) addUser(text);
        (call.ready || Promise.resolve()).then(() => runTool("transfer_to_staff", { reason: "Caller asked for the operator" })).catch((e) => console.error("operator", e));
        return;
      }
      if (call.holdTimer) {
        clearTimeout(call.holdTimer);
        call.holdTimer = null;
        text = `${call.held} ${text}`.trim();
        call.held = "";
      }
      if (!immediate && looksUnfinished(text)) {
        call.held = text;
        call.holdTimer = setTimeout(() => { call.holdTimer = null; const t = call.held; call.held = ""; push(t); }, 1200);
        return;
      }
      push(text);
    }

    socket.on("message", (raw) => {
      let msg;
      try { msg = JSON.parse(raw); } catch { return; }
      if (msg.type === "setup") { call.ready = setup(msg).catch((e) => console.error("setup", e)); return; }
      if (msg.type === "prompt") { if (msg.last === false || call.whitelist) return; enqueue(msg.voicePrompt, false); return; }
      if (msg.type === "interrupt") { trimHeard(msg.utteranceUntilInterrupt); return; }
      if (msg.type === "dtmf") { enqueue(`(pressed ${msg.digit})`, true); return; }
      if (msg.type === "error") console.error("relay error", msg.description);
    });

    socket.on("close", () => {
      call.ended = true;
      if (call.holdTimer) clearTimeout(call.holdTimer);
      if (!call.sid || call.whitelist) return; // whitelist / blacklist rows keep the outcome /voice gave them
      const transcript = call.messages.map((m) => ({
        role: m.role,
        text: typeof m.content === "string" ? m.content : m.content.filter((c) => c.type === "text").map((c) => c.text).join(" "),
      })).filter((m) => m.text);
      db.from("voice_calls").update({
        ended_at: new Date().toISOString(), outcome: call.outcome || "hangup",
        booking_id: call.bookingId, quote_id: call.quoteId, transcript,
      }).eq("call_sid", call.sid).then(() => {}, () => {});
    });
  });
});


// ---------------------------------------------------------------- online bookings & enquiries from midas.com.au (forwarded emails)
// Head office's website forms (Gravity Forms) email the store. Forward them to the Resend inbound address and they land in the app Inbox.
const FORM_LABELS = ["Choose your store", "State", "Store", "Service Details", "Service Required", "Preferred Date", "Drop Off Time", "Additional Comments",
  "Vehicle Details", "Make", "Model", "Year", "Registration No.", "Registration", "Personal Details", "First Name", "Last Name", "Email", "Phone", "Telephone",
  "Preferred Method of Contact", "Postcode", "Enquiry", "Megabus Id", "Comments", "Message"];
const SECTION_ONLY = new Set(["Choose your store", "Service Details", "Vehicle Details", "Personal Details"]);
function htmlToText(html) {
  return String(html || "")
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|tr|li|h\d|table)>/gi, "\n").replace(/<\/t[dh]>/gi, "\t")
    .replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&#39;|&apos;/g, "'").replace(/&quot;/g, '"');
}
function parseMidasForm(subject, text) {
  const lines = String(text || "").replace(/\r/g, "").split("\n").map((l) => l.replace(/\t/g, " ").replace(/\s+/g, " ").trim());
  const isLabel = (l) => { const c = l.replace(/\*$/, "").trim(); return FORM_LABELS.find((x) => x.toLowerCase() === c.toLowerCase()) || null; };
  const f = {}; let cur = null;
  for (const l of lines) {
    if (!l) continue;
    const lab = isLabel(l);
    if (lab) { cur = SECTION_ONLY.has(lab) ? null : lab; if (cur && f[cur] == null) f[cur] = ""; continue; }
    if (/^(map it|australia)\b/i.test(l) || /gravityforms|^<https?:/i.test(l) || /^I (have read|consent)/i.test(l) || /REQUESTED DATE SUBJECT TO AVAILABILITY/i.test(l)) { if (!/^australia/i.test(l)) cur = null; continue; }
    if (cur) f[cur] = (f[cur] ? f[cur] + "\n" : "") + l.replace(/\s*<mailto:[^>]*>/i, "").replace(/<https?:[^>]*>/g, "").trim();
  }
  const g = (...k) => { for (const x of k) if (f[x]) return f[x].trim(); return ""; };
  const store = g("Store");
  const kind = /enquiry/i.test(subject || "") || (f.Enquiry != null && f["Service Required"] == null) ? "enquiry" : "booking";
  const first = titleCase(g("First Name").toLowerCase()), last = titleCase(g("Last Name").toLowerCase());
  const phone = normPhone(g("Phone", "Telephone")) || g("Phone", "Telephone");
  if (!first && !phone) return null;
  const dm = g("Preferred Date").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  const date = dm ? `${dm[3]}-${pad(+dm[2])}-${pad(+dm[1])}` : null;
  const tm = g("Drop Off Time").match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i);
  const time = tm ? `${pad(((+tm[1]) % 12) + (/p/i.test(tm[3] || "") ? 12 : 0))}:${tm[2]}` : null;
  const make = titleCase(g("Make").toLowerCase()), model = g("Model").length <= 4 ? g("Model").toUpperCase() : titleCase(g("Model").toLowerCase());
  return {
    kind, store, name: [first, last].filter(Boolean).join(" "), phone, email: g("Email").toLowerCase(),
    rego: cleanRego(g("Registration No.", "Registration")) || null, rego_state: g("State") || "QLD",
    service: g("Service Required") || (kind === "enquiry" ? "Enquiry" : ""), preferred_date: date, preferred_time: time,
    message: g("Additional Comments", "Enquiry", "Comments", "Message") || null,
    make: make || null, model: model || null, vehicle_year: g("Year") || null,
    vehicle: [g("Year"), make, model].filter(Boolean).join(" ") || null,
    contact_method: g("Preferred Method of Contact") || null, postcode: g("Postcode").replace(/\D/g, "").slice(0, 4) || null,
  };
}
async function saveOnlineRequest(p, { emailId = null, receivedAt = null, rawText = "", source = null } = {}) {
  if (!p) return { ok: false, error: "Couldn't find a booking or enquiry in that email" };
  if (p.store && !/ipswich/i.test(p.store)) return { ok: false, error: `That one is for another store (${p.store})` };
  const key = crypto.createHash("sha1").update([p.kind, normPhone(p.phone) || p.phone, p.rego || "", p.preferred_date || "", p.preferred_time || "", String(p.message || "").slice(0, 300)].join("|")).digest("hex");
  const row = {
    kind: p.kind, source: source || (p.kind === "enquiry" ? "midas_enquiry_form" : "midas_booking_form"),
    name: p.name, phone: p.phone, email: p.email || null, rego: p.rego, rego_state: p.rego_state, service: p.service || null,
    preferred_date: p.preferred_date, preferred_time: p.preferred_time, message: p.message, make: p.make, model: p.model,
    vehicle_year: p.vehicle_year, vehicle: p.vehicle, contact_method: p.contact_method, postcode: p.postcode,
    email_id: emailId, dedupe_key: key, raw_text: String(rawText || "").slice(0, 8000), status: "new",
    ...(receivedAt ? { created_at: receivedAt } : {}),
  };
  const { data, error } = await db.from("booking_requests").insert(row).select("id").single();
  if (error) {
    if (/duplicate|unique/i.test(error.message)) return { ok: true, duplicate: true };
    console.error("online request insert", error.message); return { ok: false, error: error.message };
  }
  return { ok: true, id: data.id, kind: p.kind, name: p.name };
}
// Resend "email.received" webhook (signed with Svix – set RESEND_WEBHOOK_SECRET from the Resend dashboard)
function svixOk(req) {
  const secret = env.RESEND_WEBHOOK_SECRET;
  if (!secret) return false;
  const id = req.headers["svix-id"], ts = req.headers["svix-timestamp"], sigs = String(req.headers["svix-signature"] || "");
  if (!id || !ts || !sigs || Math.abs(Date.now() / 1000 - Number(ts)) > 600) return false;
  const key = Buffer.from(secret.replace(/^whsec_/, ""), "base64");
  const expected = crypto.createHmac("sha256", key).update(`${id}.${ts}.${req.rawBody || ""}`).digest("base64");
  return sigs.split(" ").some((s) => { const v = s.split(",")[1] || ""; return v.length === expected.length && crypto.timingSafeEqual(Buffer.from(v), Buffer.from(expected)); });
}
app.post("/api/inbound/resend", async (req, reply) => {
  if (!svixOk(req)) return reply.code(401).send({ error: "bad signature" });
  const ev = req.body || {};
  if (ev.type !== "email.received" || !ev.data?.email_id) return reply.send({ ok: true, ignored: true });
  try {
    const r = await fetch(`https://api.resend.com/emails/receiving/${encodeURIComponent(ev.data.email_id)}`, { headers: { Authorization: `Bearer ${env.RESEND_API_KEY}` } });
    const em = await r.json();
    if (!r.ok) throw new Error(em?.message || "HTTP " + r.status);
    const subject = em.subject || ev.data.subject || "";
    const text = em.text || htmlToText(em.html);
    if (!/midas/i.test(subject + " " + text) || !/(booking request|store enquiry|submission)/i.test(subject + " " + text)) return reply.send({ ok: true, ignored: "not a Midas form" });
    const res = await saveOnlineRequest(parseMidasForm(subject, text), { emailId: em.id || ev.data.email_id, rawText: text });
    if (!res.ok) console.warn("inbound email not saved:", res.error, subject);
    reply.send({ ok: true, saved: res.ok && !res.duplicate });
  } catch (e) { console.error("inbound email", e); reply.code(500).send({ error: "could not read email" }); }
});
// staff: paste an email's text straight into the Inbox (backup if forwarding isn't set up)
app.post("/api/inbound/paste", async (req, reply) => {
  const who = await requireStaffUser(req, reply); if (!who) return;
  const b = req.body || {};
  const text = String(b.text || "").slice(0, 20000);
  const p = parseMidasForm(b.subject || text.slice(0, 400), text);
  const res = await saveOnlineRequest(p, { rawText: text, source: p ? (p.kind === "enquiry" ? "midas_enquiry_form" : "midas_booking_form") : null });
  if (!res.ok) return reply.code(400).send(res);
  reply.send(res);
});

// ---------------------------------------------------------------- quotes: send by SMS / email, customer accept + online booking
const PUBLIC_BASE = (env.PUBLIC_QUOTE_BASE || "https://go.midasipswich.com.au").replace(/\/$/, ""); // custom domain on Railway (CNAME go -> Railway)
const quoteLink = (q) => `${PUBLIC_BASE}/q/${q.public_token}`;
const firstName = (q) => titleCase(q.given || String(q.name || "").split(" ")[0] || "");
const qItems = (q) => (Array.isArray(q.line_items) ? q.line_items : []).filter((i) => i && (i.description || i.price));
const qTotal = (q) => Number(q.total) || qItems(q).reduce((a, i) => a + Number(i.price || 0) * Number(i.quantity || 1), 0);
const $2 = (n) => "$" + Number(n || 0).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
// what the customer sees: packages are one line with the package total; our own lines stay itemised; remarks left out
const groupTitle = (g) => titleCase(String(g || "").toLowerCase());
function custLines(q) {
  const out = [], byG = {};
  for (const i of qItems(q)) {
    if (i.isRemark) continue;
    const amt = Number(i.price || 0) * Number(i.quantity || 1);
    const g = i.packageGroup && i.packageGroup !== "Additional Items" ? String(i.packageGroup) : "";
    if (g) {
      if (!byG[g]) { byG[g] = { description: groupTitle(g), quantity: 1, amount: 0, brand: "" }; out.push(byG[g]); }
      byG[g].amount += amt;
    } else out.push({ description: String(i.description || ""), quantity: Number(i.quantity || 1), amount: amt, brand: i.brand || "" });
  }
  return out;
}
const vehicleOf = (q) => [q.vehicle, q.rego ? `(${String(q.rego).toUpperCase()})` : ""].filter(Boolean).join(" ");

function quoteSmsText(q, s) {
  const items = custLines(q).map((i) => `- ${i.description}${i.quantity > 1 ? ` x${i.quantity}` : ""} ${$2(i.amount)}`).join("\n");
  return `Hi ${firstName(q)}, here is your estimate${vehicleOf(q) ? ` for your ${vehicleOf(q)}` : ""}:\n${items}\nTotal ${$2(qTotal(q))} inc GST\nTo accept your estimate and book online: ${quoteLink(q)}\nQuestions? Call us on ${s.phone || "07 3444 2200"}.`;
}
function quoteEmailHtml(q, s, intro) {
  const rows = custLines(q).map((i) => `<tr><td style="padding:8px 10px;border-bottom:1px solid #eee">${esc(i.description)}${i.quantity > 1 ? ` × ${esc(i.quantity)}` : ""}${i.brand ? `<div style="color:#888;font-size:12px">${esc(i.brand)}</div>` : ""}</td><td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap">${$2(i.amount)}</td></tr>`).join("");
  return `<div style="background:#f3f3f3;padding:16px 0"><div style="max-width:620px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;font:15px/1.5 Arial,sans-serif;color:#222">
    <div style="background:#FFEB00;border-bottom:6px solid #DA291C;padding:12px 20px"><table role="presentation" style="width:100%;border-collapse:collapse"><tr><td style="width:120px"><img src="${PUBLIC_BASE}/brand/midas-logo.png" alt="Midas" width="110" style="display:block;width:110px;height:auto"></td><td style="font:800 18px Arial,sans-serif;color:#000">Ipswich<span style="color:#DA291C">.</span><div style="font:600 11px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase">You can rely on us</div></td></tr></table></div>
    <div style="padding:18px 20px 0;font:700 30px Impact,'Arial Narrow',Arial,sans-serif;text-transform:uppercase;letter-spacing:.5px">Your estimate<span style="color:#DA291C">.</span></div>
    <div style="padding:16px 20px">
      <p>${esc(intro || `Hi ${firstName(q)}, thanks for your enquiry. Here is your estimate${vehicleOf(q) ? ` for your ${vehicleOf(q)}` : ""}.`)}</p>
      <table style="width:100%;border-collapse:collapse;margin:10px 0">${rows}
        <tr><td style="padding:10px;font-weight:700">Total (inc GST)</td><td style="padding:10px;text-align:right;font-weight:700;font-size:18px">${$2(qTotal(q))}</td></tr></table>
      ${q.valid_until ? `<p style="color:#666;font-size:13px">Estimate valid until ${esc(longDate(q.valid_until))}.</p>` : ""}
      <p style="text-align:center;margin:22px 0"><a href="${quoteLink(q)}" style="background:#DA291C;color:#fff;text-decoration:none;font-weight:700;padding:15px 30px;border-radius:8px;display:inline-block;text-transform:uppercase;letter-spacing:.04em">Accept &amp; book online</a></p>
      <p style="color:#666;font-size:13px">Questions? Reply to this email or call us on ${esc(s.phone || "07 3444 2200")}. Prices can vary once we inspect the vehicle; we'll always contact you before doing anything extra.</p>
    </div>
    <div style="background:#fafafa;border-top:1px solid #eee;padding:10px 20px;color:#888;font-size:12px">${esc(s.name || "Midas Ipswich")}${s.address ? " · " + esc(s.address) : ""}</div>
  </div></div>`;
}

// staff: preview or send a quote
app.post("/api/quotes/:id/send", async (req, reply) => {
  const who = await requireStaffUser(req, reply); if (!who) return;
  const s = await getSettings();
  const { data: q, error } = await db.from("quotes").select("*").eq("id", req.params.id).maybeSingle();
  if (error || !q) return reply.code(404).send({ error: "Estimate not found" });
  if (!q.public_token) { const t = crypto.randomUUID().replace(/-/g, ""); await db.from("quotes").update({ public_token: t }).eq("id", q.id); q.public_token = t; }
  if (!q.valid_until) q.valid_until = addDays(todayISO(), Number(s.quote_valid_days || 30));
  const b = req.body || {};
  const via = ["sms", "email", "both"].includes(b.via) ? b.via : (q.email ? "both" : "sms");
  const smsText = String(b.sms_text || quoteSmsText(q, s));
  const subject = b.subject || `Your estimate from Midas Ipswich${vehicleOf(q) ? " – " + vehicleOf(q) : ""}`;
  const html = quoteEmailHtml(q, s, b.email_intro);
  if (b.preview) return reply.send({ ok: true, via, to_phone: normPhone(q.phone), to_email: q.email || "", sms_text: smsText, sms_segments: smsSegments((s.sms_header ? s.sms_header + " " : "") + smsText), subject, email_html: html, link: quoteLink(q) });
  const sent = [];
  if ((via === "sms" || via === "both") && q.phone) {
    const r = await sendSms({ to: q.phone, body: smsText, template: "ESTIMATE", sentBy: who.name });
    if (!r.ok) return reply.code(500).send({ error: r.error || "SMS failed" });
    sent.push("sms");
  }
  if ((via === "email" || via === "both") && q.email) {
    const r = await sendEmail({ to: q.email, subject, html });
    if (!r.ok) return reply.code(500).send({ error: "Email failed" });
    sent.push("email");
  }
  if (!sent.length) return reply.code(400).send({ error: "No mobile number or email on this estimate" });
  await db.from("quotes").update({ status: "Sent", sent_at: new Date().toISOString(), sent_via: sent.join("+"), valid_until: q.valid_until }).eq("id", q.id);
  reply.send({ ok: true, sent, link: quoteLink(q) });
});

// ---- public (customer) side, found only by the private link
async function quoteByToken(token) {
  if (!/^[a-f0-9]{16,64}$/i.test(String(token || ""))) return null;
  const { data } = await db.from("quotes").select("*").eq("public_token", token).maybeSingle();
  return data || null;
}
async function quoteHours(q) {
  const jt = await getJobTypes();
  const text = qItems(q).map((i) => i.description).join(" ").toLowerCase() + " " + String(q.notes || "").toLowerCase();
  let h = 0;
  for (const j of jt) {
    const keys = [j.name, ...String(j.keywords || "").split(",")].map((k) => k.trim().toLowerCase()).filter((k) => k.length > 2);
    if (keys.some((k) => text.includes(k))) h += Number(j.hours) || 0;
  }
  const s = await getSettings();
  return Math.min(8, h || Number(s.quote_default_hours || 2));
}
async function quoteDays(q) {
  const s = await getSettings();
  const hours = await quoteHours(q);
  const from = addDays(todayISO(), 1);
  const cap = await getCapacity(from, addDays(from, 34));
  return cap.filter((r) => !(r.closed ?? r.blocked) && dayFits(r, hours)).slice(0, 14).map((r) => ({ day: r.day, label: longDate(r.day), times: dayHours(s, r.day).dropoff }));
}
app.get("/api/public/quote/:token", async (req, reply) => {
  const q = await quoteByToken(req.params.token);
  if (!q) return reply.code(404).send({ error: "This estimate link isn't valid. Please call us." });
  const s = await getSettings();
  if (!q.viewed_at && !req.query?.staff) db.from("quotes").update({ viewed_at: new Date().toISOString() }).eq("id", q.id).then(() => {}, () => {});
  let booking = null;
  if (q.accepted_booking_id) { const { data } = await db.from("bookings").select("booking_date,time_in,state").eq("id", q.accepted_booking_id).maybeSingle(); booking = data ? { day: data.booking_date, label: longDate(data.booking_date), time: time12(data.time_in) } : null; }
  const expired = q.valid_until && q.valid_until < todayISO() && !q.accepted_at;
  reply.send({
    quote: { first_name: firstName(q), vehicle: q.vehicle || "", rego: q.rego || "", items: custLines(q).map((i) => ({ description: i.description, quantity: i.quantity, amount: i.amount })), total: qTotal(q), valid_until: q.valid_until ? longDate(q.valid_until) : null, accepted: !!q.accepted_at, expired, booking },
    shop: { name: s.name || "Midas Ipswich", phone: s.phone || "07 3444 2200", address: s.address || "" },
    days: q.accepted_at || expired ? [] : await quoteDays(q),
  });
});
app.post("/api/public/quote/:token/accept", async (req, reply) => {
  const q = await quoteByToken(req.params.token);
  if (!q) return reply.code(404).send({ error: "This estimate link isn't valid. Please call us." });
  if (q.accepted_at) return reply.send({ ok: true, already: true });
  if (q.valid_until && q.valid_until < todayISO()) return reply.code(410).send({ error: "This estimate has expired. Please call us for an updated price." });
  const s = await getSettings();
  const b = req.body || {};
  // claim the quote first so a double tap can't make two bookings
  const { data: claimed } = await db.from("quotes").update({ accepted_at: new Date().toISOString(), status: "Accepted" }).eq("id", q.id).is("accepted_at", null).select("id");
  if (!claimed?.length) return reply.send({ ok: true, already: true });
  const unclaim = () => db.from("quotes").update({ accepted_at: null, status: q.status || "Sent" }).eq("id", q.id).then(() => {}, () => {});
  let booking = null;
  if (b.day) {
    const day = String(b.day), time = String(b.time || "");
    const hours = await quoteHours(q);
    const days = await quoteDays(q);
    const d = days.find((x) => x.day === day);
    if (!d) { await unclaim(); return reply.code(409).send({ error: "Sorry, that day has just filled up. Please pick another day." }); }
    if (!d.times.includes(time)) { await unclaim(); return reply.code(400).send({ error: "Please pick a drop-off time." }); }
    const confirmed = !!s.ai_auto_confirm;
    const mobile = normPhone(q.phone);
    const { data: bk, error } = await db.from("bookings").insert({
      shop_id: SHOP_ID, customer_name: q.name || firstName(q), given: firstName(q), surname: q.surname || null, phone: mobile, email: q.email || null,
      rego: cleanRego(q.rego) || null, rego_state: q.rego_state || "QLD", vehicle: q.vehicle || null,
      service: qItems(q).map((i) => i.description).join(", ").slice(0, 300) || "Quoted work", job_types: [], hours, duration_mins: Math.round(hours * 60),
      description: `Accepted online estimate #${q.id} (${$2(qTotal(q))} inc GST)`,
      booking_date: day, time_in: time, time_out_mode: "all_day", time_out: b.needed_back ? String(b.needed_back).slice(0, 40) : null,
      start_time: `${day}T${time}:00+10:00`, state: confirmed ? "booking" : "inbox", status: "booked", source: "quote_online",
      salesperson: "ONLINE", created_by: "Customer (online estimate)",
      notes: [`Booked online from estimate #${q.id}.`, b.notes ? `Customer note: ${String(b.notes).slice(0, 300)}` : null].filter(Boolean).join(" "),
    }).select("*").single();
    if (error) { console.error("quote booking", error.message); await unclaim(); return reply.code(500).send({ error: "Sorry, we couldn't save the booking. Please call us." }); }
    booking = bk;
    if (confirmed) sendBookingTemplate(bk, "BOOKIN", "Online estimate").catch(() => {});
  }
  if (booking) await db.from("quotes").update({ accepted_booking_id: booking.id }).eq("id", q.id);
  sendEmail({
    to: s.notify_email, subject: `Estimate accepted${booking ? " and booked" : ""} – ${q.name || firstName(q)} ${vehicleOf(q)}`,
    html: `<p><b>${esc(q.name || "")}</b> ${esc(normPhone(q.phone) || "")} accepted estimate #${q.id} (${$2(qTotal(q))} inc GST)${vehicleOf(q) ? " for their " + esc(vehicleOf(q)) : ""}.</p>${booking ? `<p>Booked online for <b>${esc(longDate(booking.booking_date))}</b>, drop off ${esc(time12(booking.time_in))}${booking.state === "inbox" ? " – <b>waiting in the Inbox to be accepted</b>" : ""}.</p>` : "<p>They didn't pick a day – give them a call to book in.</p>"}`,
  }).catch(() => {});
  reply.send({ ok: true, booked: !!booking, confirmed: booking?.state === "booking", label: booking ? `${longDate(booking.booking_date)}, drop off from ${time12(booking.time_in)}` : null });
});

// the customer's page (self-contained, no app login) – Midas brand: Bebas Neue + Montserrat, yellow / red / black, oval price point
app.get("/q/:token", async (req, reply) => {
  reply.header("Cache-Control", "no-store").header("X-Robots-Tag", "noindex").type("text/html").send(QUOTE_PAGE);
});
const MIDAS_LOGO_PNG = Buffer.from("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", "base64");
const FRANK_WEBP = Buffer.from("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", "base64");
app.get("/brand/frank.webp", async (req, reply) => {
  reply.header("Cache-Control", "public, max-age=604800").type("image/webp").send(FRANK_WEBP);
});
app.get("/brand/midas-logo.png", async (req, reply) => {
  reply.header("Cache-Control", "public, max-age=604800").type("image/png").send(MIDAS_LOGO_PNG);
});
const QUOTE_PAGE = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex">
<title>Your estimate – Midas Ipswich</title>
<link rel="icon" href="/brand/midas-logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--y:#FFEB00;--r:#DA291C;--k:#000;--g:#373A36;--soft:#F5F5F2;--line:#E2E2DC;--muted:#5d5f5a}
*{box-sizing:border-box}html,body{margin:0}
body{background:var(--soft);color:var(--k);font:16px/1.55 Montserrat,Arial,sans-serif;-webkit-font-smoothing:antialiased}
.h{font-family:"Bebas Neue","Arial Narrow",Arial,sans-serif;font-weight:400;text-transform:uppercase;letter-spacing:.5px;line-height:.95;margin:0}
.dot{color:var(--r)}
.top{background:var(--y);border-bottom:6px solid var(--r)}
.top-in{max-width:640px;margin:0 auto;padding:12px 16px;display:flex;align-items:center;gap:12px}
.top img{height:48px;width:auto;display:block}
.store{font:800 17px Montserrat,Arial,sans-serif;line-height:1}
.store small{display:block;font-weight:600;font-size:11px;letter-spacing:.06em;text-transform:uppercase;margin-top:3px}
.call{margin-left:auto;background:var(--k);color:var(--y);text-decoration:none;font:700 14px Montserrat,Arial,sans-serif;padding:10px 14px;border-radius:999px;white-space:nowrap}
.hero{background:var(--y);padding:18px 16px 72px;overflow:hidden}
.hero-in{max-width:640px;margin:0 auto;position:relative;min-height:190px;padding-right:150px}
.frank{position:absolute;right:-8px;bottom:-72px;width:170px;height:auto;pointer-events:none}
.bubble{display:inline-block;position:relative;background:#fff;color:var(--r);font-family:"Bebas Neue",Arial,sans-serif;font-size:24px;letter-spacing:.3px;padding:6px 14px;border-radius:14px;margin-bottom:10px;line-height:1.1}
.bubble:after{content:"";position:absolute;right:-12px;top:12px;border:8px solid transparent;border-left:14px solid #fff}
.hero .h{font-size:54px}
.sub{color:var(--r);font-family:"Bebas Neue",Arial,sans-serif;font-size:26px;letter-spacing:.4px;margin-top:6px;line-height:1.05}
main{max-width:640px;margin:-52px auto 0;padding:0 16px 24px;position:relative;z-index:2}
.card{background:#fff;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.12);padding:20px;margin-bottom:16px}
.veh{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-weight:600}
.plate{font:700 15px Montserrat,Arial,sans-serif;letter-spacing:2px;background:#fff;border:2px solid var(--k);border-radius:6px;padding:2px 9px}
table{width:100%;border-collapse:collapse;margin-top:12px}
td{padding:12px 0;border-bottom:1px solid var(--line);vertical-align:top;font-size:15px}td.a{text-align:right;white-space:nowrap;padding-left:12px;font-weight:700}
.tot{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:16px}
.tot .lbl{font-weight:700}.tot .lbl small{display:block;font-weight:500;color:var(--muted);font-size:12px}
.oval{position:relative;background:var(--y);border:4px solid var(--r);border-radius:50%;padding:14px 26px;transform:rotate(-4deg);text-align:center;min-width:150px}
.oval .h{font-size:44px;line-height:1}
.note{font-size:13px;color:var(--muted);margin-top:14px}
.sec{font-size:38px;margin-bottom:4px}
.lead{color:var(--muted);font-size:14px;margin:0 0 6px}
label{display:block;font-weight:700;font-size:13px;text-transform:uppercase;letter-spacing:.05em;margin-top:18px}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}
.chip{border:2px solid var(--k);background:#fff;color:var(--k);border-radius:10px;padding:10px 12px;font:600 14px Montserrat,Arial,sans-serif;cursor:pointer;text-align:left;line-height:1.2}
.chip b{display:block;font-family:"Bebas Neue",Arial,sans-serif;font-weight:400;font-size:21px;letter-spacing:.4px}
.chip.on{background:var(--k);color:var(--y)}
input,textarea{width:100%;padding:12px;border:2px solid var(--line);border-radius:10px;font:16px Montserrat,Arial,sans-serif;margin-top:6px;background:#fff}
input:focus,textarea:focus{outline:none;border-color:var(--k)}
.btn{display:block;width:100%;border:0;border-radius:10px;padding:15px 16px;cursor:pointer;margin-top:12px;font-family:"Bebas Neue",Arial,sans-serif;font-size:26px;letter-spacing:.6px;line-height:1}
.btn.red{background:var(--r);color:#fff;box-shadow:0 6px 0 #9e1d14}
.btn.red:active{transform:translateY(3px);box-shadow:0 3px 0 #9e1d14}
.btn.out{background:#fff;color:var(--k);border:3px solid var(--k);font-size:22px}a.btn{text-decoration:none;text-align:center}
.btn:disabled{opacity:.45;cursor:default;box-shadow:none}
.err{color:var(--r);font-weight:700;margin-top:10px}
.trust{display:grid;gap:10px;margin:4px 0 16px}
.trust div{display:flex;gap:10px;align-items:flex-start;font-size:14px;font-weight:500}
.trust i{flex:none;width:22px;height:22px;border-radius:50%;background:var(--r);color:#fff;font-style:normal;font-weight:800;display:grid;place-items:center;font-size:13px;margin-top:1px}
.done{text-align:center}.done .h{font-size:52px}.done .big{width:76px;height:76px;margin:6px auto 10px;border-radius:50%;background:var(--r);color:#fff;display:grid;place-items:center;font-size:44px;font-weight:800;line-height:1}
.tag{font-family:"Bebas Neue",Arial,sans-serif;font-size:30px;text-align:center;margin:18px 0 6px}
footer{background:var(--k);color:#fff;padding:22px 16px 30px;font-size:14px}
footer .in{max-width:640px;margin:0 auto;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
footer b{color:var(--y)}footer a{color:#fff}
.hide{display:none}
@media(max-width:420px){.hero-in{padding-right:118px;min-height:170px}.frank{width:140px}.bubble{font-size:20px}.hero .h{font-size:46px}.oval .h{font-size:36px}.oval{min-width:120px;padding:12px 18px}}
</style></head><body>
<div class="top"><div class="top-in"><img src="/brand/midas-logo.png" alt="Midas"><div class="store">Ipswich<span class="dot">.</span><small>You can rely on us</small></div><a class="call" id="call" href="tel:0734442200">📞 Call us</a></div></div>
<div class="hero"><div class="hero-in"><div class="bubble" id="bub">Let’s be frank about…</div><h1 class="h">Here is your estimate<span class="dot">.</span></h1><div class="sub" id="hi">Loading…</div><img class="frank" src="/brand/frank.webp" alt="Frank from Midas"></div></div>
<main id="m"><div class="card">Loading your estimate…</div></main>
<footer><div class="in" id="ft"><div><b>Midas Ipswich</b><br><span id="ad"></span></div></div></footer>
<script>
const T=location.pathname.split("/").pop(),M=document.getElementById("m");
const $=(n)=>"$"+Number(n||0).toLocaleString("en-AU",{minimumFractionDigits:2,maximumFractionDigits:2});
const e=(s)=>String(s??"").replace(/[&<>"]/g,(c)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const t12=(t)=>{const[h,m]=t.split(":").map(Number);return((h+11)%12+1)+(m?":"+String(m).padStart(2,"0"):"")+(h<12?"am":"pm")};
const tel=(p)=>"tel:"+String(p||"").replace(/[^0-9+]/g,"");
let D=null,pick={day:null,time:null};
async function load(){const r=await fetch("/api/public/quote/"+T+(location.search.includes("staff=1")?"?staff=1":""));D=await r.json().catch(()=>({error:"Something went wrong – please call us."}));
 if(!r.ok){document.getElementById("hi").textContent="";M.innerHTML='<div class="card"><h2 class="h sec">Estimate not found<span class="dot">.</span></h2><p>'+e(D.error)+'</p></div>';return}
 const s=D.shop;document.getElementById("call").href=tel(s.phone);document.getElementById("ad").innerHTML=e(s.address||"")+(s.address?"<br>":"")+'<a href="'+tel(s.phone)+'">'+e(s.phone)+'</a>';render()}
function dayLabel(l){const p=String(l).split(" ");return{d:p[0],rest:p.slice(1).join(" ")}}
function render(){const q=D.quote,s=D.shop;
 if(q.accepted)document.getElementById("bub").textContent="Good as done!";
 document.getElementById("hi").textContent=q.accepted?"Thanks "+q.first_name+" – you're all sorted":"Hi "+q.first_name+" – no jargon, no surprises";
 const car=[q.vehicle].filter(Boolean).join(" ");
 let h='<div class="card"><div class="veh">'+(car?e(car):"Your vehicle")+(q.rego?' <span class="plate">'+e(q.rego)+'</span>':"")+'</div><table>'+q.items.map(i=>'<tr><td>'+e(i.description)+(i.quantity>1?" × "+i.quantity:"")+'</td><td class="a">'+$(i.amount)+'</td></tr>').join("")+'</table>'
  +'<div class="tot"><div class="lbl">Total<small>inc GST</small></div><div class="oval"><div class="h">'+$(q.total)+'</div></div></div>'
  +'<div class="note">'+(q.valid_until&&!q.accepted?"Estimate valid until "+e(q.valid_until)+". ":"")+'If we find anything else once we start, we\\'ll call you before doing any extra work.</div></div>';
 if(q.accepted){h+='<div class="card done"><div class="big">✓</div><h2 class="h">You\\'re booked in<span class="dot">.</span></h2>'+(q.booking?'<p style="font-size:18px">'+e(q.booking.label)+'<br>Drop off from <b>'+e(q.booking.time)+'</b></p><p class="note">Running early or dropping off the night before? Use the key drop box on the wall of the building.</p>':'<p>Thanks for accepting your estimate. We\\'ll be in touch shortly to book you in.</p>')+'<a class="btn out" href="'+tel(s.phone)+'">Questions? Call '+e(s.phone)+'</a></div>'+trust();M.innerHTML=h;return}
 if(q.expired){h+='<div class="card"><h2 class="h sec">This estimate has expired<span class="dot">.</span></h2><p>Prices can change, so give us a call for an updated estimate.</p><a class="btn red" href="'+tel(s.phone)+'">Call '+e(s.phone)+'</a></div>';M.innerHTML=h;return}
 h+='<div class="card"><h2 class="h sec">Pick a day<span class="dot">.</span></h2><p class="lead">Drop your car off in the morning and we\\'ll call you when it\\'s ready.</p>';
 if(!D.days.length)h+='<p>We don\\'t have any spots online right now. Accept the estimate below and we\\'ll call you to book in.</p>';
 else{h+='<label>Day</label><div class="chips">'+D.days.map(d=>{const x=dayLabel(d.label);return '<button class="chip'+(pick.day===d.day?" on":"")+'" data-d="'+d.day+'"><b>'+e(x.d)+'</b>'+e(x.rest)+'</button>'}).join("")+'</div>';
  const day=D.days.find(d=>d.day===pick.day);
  if(day)h+='<label>Drop-off time</label><div class="chips">'+day.times.map(t=>'<button class="chip'+(pick.time===t?" on":"")+'" data-t="'+t+'"><b>'+t12(t)+'</b></button>').join("")+'</div>';
  h+='<label for="nb">When do you need it back? <span style="font-weight:500;text-transform:none">(optional)</span></label><input id="nb" placeholder="e.g. 3pm, end of day">';
  h+='<label for="nt">Anything we should know? <span style="font-weight:500;text-transform:none">(optional)</span></label><textarea id="nt" rows="2"></textarea>';
  h+='<button class="btn red" id="bk"'+(pick.day&&pick.time?"":" disabled")+'>'+(day&&pick.time?"Accept estimate & book "+e(dayLabel(day.label).d)+" "+t12(pick.time):"Accept estimate & book")+'</button>'}
 h+='<button class="btn out" id="ac">Accept estimate – I\\'ll call to book</button><div class="err" id="er"></div></div>'+trust();
 M.innerHTML=h}
function trust(){return '<div class="trust"><div><i>✓</i>No extra work without your OK – we always call first.</div><div><i>✓</i>Logbook services that keep your new-car warranty.</div><div><i>✓</i>Our work is guaranteed at Midas stores nationwide.</div></div><div class="tag">You can rely on us<span class="dot">.</span></div>'}
async function accept(withDay){const er=document.getElementById("er");er.textContent="";document.querySelectorAll(".btn").forEach(b=>b.disabled=true);
 const body=withDay?{day:pick.day,time:pick.time,needed_back:(document.getElementById("nb")||{}).value||"",notes:(document.getElementById("nt")||{}).value||""}:{};
 const r=await fetch("/api/public/quote/"+T+"/accept",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)}).catch(()=>({ok:false,status:0,json:()=>({})}));const j=await Promise.resolve(r.json()).catch(()=>({}));
 if(!r.ok){er.textContent=j.error||"Something went wrong – please call us.";document.querySelectorAll(".btn").forEach(b=>b.disabled=false);if(r.status===409)load();return}
 window.scrollTo({top:0,behavior:"smooth"});load()}
M.addEventListener("click",(ev)=>{const d=ev.target.closest("[data-d]"),t=ev.target.closest("[data-t]");
 const keep=()=>({nb:document.getElementById("nb")?.value,nt:document.getElementById("nt")?.value}),put=(k)=>{if(k.nb)document.getElementById("nb").value=k.nb;if(k.nt)document.getElementById("nt").value=k.nt};
 if(d){const k=keep();pick.day=d.dataset.d;pick.time=null;render();put(k)}else if(t){const k=keep();pick.time=t.dataset.t;render();put(k)}
 else if(ev.target.closest("#bk"))accept(true);else if(ev.target.closest("#ac"))accept(false)});
load();
</script></body></html>`;

// ---------------------------------------------------------------- "pick another day" page for online booking requests / enquiries
const offerLink = (r) => `${PUBLIC_BASE}/b/${r.offer_token}`;
async function reqHours(r) {
  const jt = await getJobTypes();
  const text = `${r.service || ""} ${r.message || ""}`.toLowerCase();
  let h = 0;
  for (const j of jt) {
    const keys = [j.name, ...String(j.keywords || "").split(",")].map((k) => k.trim().toLowerCase()).filter((k) => k.length > 2);
    if (keys.some((k) => text.includes(k))) h = Math.max(h, Number(j.hours) || 0);
  }
  if (!h) { const svc = jt.find((j) => j.code === "SERVICE"); h = svc ? Number(svc.hours) : 1.5; }
  return Math.min(8, h);
}
async function offerDays(r, n = 14) {
  const s = await getSettings();
  const hours = await reqHours(r);
  const from = addDays(todayISO(), 1);
  const cap = await getCapacity(from, addDays(from, 45));
  return cap.filter((x) => !(x.closed ?? x.blocked) && dayFits(x, hours)).slice(0, n).map((x) => ({ day: x.day, label: longDate(x.day), times: dayHours(s, x.day).dropoff }));
}
async function reqByOfferToken(token) {
  if (!/^[a-f0-9]{16,64}$/i.test(String(token || ""))) return null;
  const { data } = await db.from("booking_requests").select("*").eq("offer_token", token).maybeSingle();
  return data || null;
}
const reqFirst = (r) => titleCase(String(r.name || "").trim().split(/\s+/)[0] || "there");
const shortLabel = (iso) => longDate(iso).replace(/\s+\d{4}$/, "");
function offerSmsText(r, s, days) {
  const wanted = r.kind !== "enquiry" && r.preferred_date ? r.preferred_date : null;
  const list = days.slice(0, 3).map((d) => shortLabel(d.day).replace(/^(\w{3})\w*/, "$1"));
  const next = list.length ? ` Our next available days are ${list.length > 1 ? list.slice(0, -1).join(", ") + " or " + list[list.length - 1] : list[0]}.` : "";
  return (wanted
    ? `Hi ${reqFirst(r)}, thanks for your booking request for ${shortLabel(wanted).replace(/^(\w{3})\w*/, "$1")}. Sorry, we're fully booked that day.${next}`
    : `Hi ${reqFirst(r)}, thanks for your enquiry.${next}`) + `\nPick the day that suits and book online here: ${offerLink(r)}\nQuestions? Call us on ${s.phone || "07 3444 2200"}.`;
}
// staff: preview / send the "pick another day" text
app.post("/api/requests/:id/offer", async (req, reply) => {
  const who = await requireStaffUser(req, reply); if (!who) return;
  const s = await getSettings();
  const { data: r } = await db.from("booking_requests").select("*").eq("id", req.params.id).maybeSingle();
  if (!r) return reply.code(404).send({ error: "Request not found" });
  if (!r.offer_token) { r.offer_token = crypto.randomUUID().replace(/-/g, ""); await db.from("booking_requests").update({ offer_token: r.offer_token }).eq("id", r.id); }
  const days = await offerDays(r, 3);
  const b = req.body || {};
  const smsText = String(b.sms_text || offerSmsText(r, s, days));
  if (b.preview) return reply.send({ ok: true, to_phone: normPhone(r.phone), sms_text: smsText, sms_segments: smsSegments((s.sms_header ? s.sms_header + " " : "") + smsText), link: offerLink(r), days: days.map((d) => d.label) });
  if (!normPhone(r.phone)) return reply.code(400).send({ error: "No mobile number on this request" });
  const x = await sendSms({ to: r.phone, body: smsText, template: "OFFER", sentBy: who.name });
  if (!x.ok) return reply.code(500).send({ error: x.error || "SMS failed" });
  await db.from("booking_requests").update({ offer_sent_at: new Date().toISOString() }).eq("id", r.id);
  reply.send({ ok: true, link: offerLink(r) });
});
app.get("/api/public/booking/:token", async (req, reply) => {
  const r = await reqByOfferToken(req.params.token);
  if (!r) return reply.code(404).send({ error: "This link isn't valid. Please call us." });
  const s = await getSettings();
  if (!r.offer_viewed_at && !req.query?.staff) db.from("booking_requests").update({ offer_viewed_at: new Date().toISOString() }).eq("id", r.id).then(() => {}, () => {});
  let booked = null;
  if (r.booked_id) { const { data: b } = await db.from("bookings").select("booking_date,time_in,state,status").eq("id", r.booked_id).maybeSingle(); if (b && b.status !== "cancelled") booked = { label: longDate(b.booking_date), time: time12(b.time_in), confirmed: b.state === "booking" }; }
  reply.header("Cache-Control", "no-store").send({
    first_name: reqFirst(r), vehicle: r.vehicle || "", rego: r.rego || "", service: r.kind === "enquiry" ? "" : r.service || "",
    wanted: r.kind !== "enquiry" && r.preferred_date ? longDate(r.preferred_date) : "", enquiry: r.kind === "enquiry", message: r.kind === "enquiry" ? r.message || "" : "",
    days: booked ? [] : await offerDays(r, 14), booked,
    shop: { name: s.name || "Midas Ipswich", phone: s.phone || "07 3444 2200", address: s.address || "" },
  });
});
app.post("/api/public/booking/:token", async (req, reply) => {
  const r = await reqByOfferToken(req.params.token);
  if (!r) return reply.code(404).send({ error: "This link isn't valid. Please call us." });
  if (r.booked_id) return reply.send({ ok: true, already: true });
  const b = req.body || {};
  const day = String(b.day || ""), time = String(b.time || "");
  const days = await offerDays(r, 30);
  const d = days.find((x) => x.day === day);
  if (!d) return reply.code(409).send({ error: "Sorry, that day has just filled up. Please pick another day." });
  if (!d.times.includes(time)) return reply.code(400).send({ error: "Please pick a drop-off time." });
  // claim first so a double tap can't make two bookings
  const claimId = crypto.randomUUID();
  const { data: claimed } = await db.from("booking_requests").update({ booked_id: claimId }).eq("id", r.id).is("booked_id", null).select("id");
  if (!claimed?.length) return reply.send({ ok: true, already: true });
  const s = await getSettings();
  const hours = await reqHours(r);
  const confirmed = !!s.ai_auto_confirm;
  const [given, ...rest] = String(r.name || "").trim().split(/\s+/);
  const { data: bk, error } = await db.from("bookings").insert({
    shop_id: SHOP_ID, customer_name: r.name || given, given: titleCase(given || ""), surname: titleCase(rest.join(" ")) || null,
    phone: normPhone(r.phone) || r.phone, email: r.email || null, rego: cleanRego(r.rego) || null, rego_state: r.rego_state || "QLD",
    vehicle: r.vehicle || null, make: r.make || null, model: r.model || null,
    service: (r.kind === "enquiry" ? "" : r.service) || "Online booking", job_types: [], hours, duration_mins: Math.round(hours * 60),
    description: [r.kind === "enquiry" ? "" : r.service, r.message].filter(Boolean).join(" – ") || null,
    booking_date: day, time_in: time, time_out_mode: "all_day", time_out: b.needed_back ? String(b.needed_back).slice(0, 40) : null,
    start_time: `${day}T${time}:00+10:00`, state: confirmed ? "booking" : "inbox", status: "booked", source: "website",
    salesperson: "ONLINE", created_by: "Customer (online)",
    notes: [`Booked online from ${r.kind === "enquiry" ? "an enquiry" : "a booking request"} (midas.com.au).`, b.notes ? `Customer note: ${String(b.notes).slice(0, 300)}` : null].filter(Boolean).join(" "),
  }).select("*").single();
  if (error) {
    console.error("offer booking", error.message);
    await db.from("booking_requests").update({ booked_id: null }).eq("id", r.id);
    return reply.code(500).send({ error: "Sorry, we couldn't save the booking. Please call us." });
  }
  await db.from("booking_requests").update({ booked_id: bk.id, status: "converted", handled_at: new Date().toISOString(), handled_by: "Customer (online)" }).eq("id", r.id);
  if (confirmed) sendBookingTemplate(bk, "BOOKIN", "Online").catch(() => {});
  sendEmail({
    to: s.notify_email, subject: `Booked online – ${r.name} – ${longDate(day)} ${time12(time)}`,
    html: `<p><b>${esc(r.name || "")}</b> ${esc(normPhone(r.phone) || "")} picked a day from our text and booked online.</p><p><b>${esc(longDate(day))}</b>, drop off from <b>${esc(time12(time))}</b>${b.needed_back ? ` · needed back ${esc(b.needed_back)}` : ""}</p><p>${esc(r.vehicle || "")} ${esc(r.rego || "")} · ${esc(r.service || r.message || "")}</p>${b.notes ? `<p>Note: ${esc(b.notes)}</p>` : ""}<p>${confirmed ? "Confirmed and the customer has been texted." : "It's in the Inbox for you to accept."}</p>`,
  }).catch(() => {});
  reply.send({ ok: true, confirmed, label: `${longDate(day)}, drop off from ${time12(time)}` });
});
app.get("/b/:token", async (req, reply) => {
  reply.header("Cache-Control", "no-store").header("X-Robots-Tag", "noindex").type("text/html").send(BOOK_PAGE);
});

const BOOK_PAGE = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex">
<title>Book your car in – Midas Ipswich</title>
<link rel="icon" href="/brand/midas-logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--y:#FFEB00;--r:#DA291C;--k:#000;--g:#373A36;--soft:#F5F5F2;--line:#E2E2DC;--muted:#5d5f5a}
*{box-sizing:border-box}html,body{margin:0}
body{background:var(--soft);color:var(--k);font:16px/1.55 Montserrat,Arial,sans-serif;-webkit-font-smoothing:antialiased}
.h{font-family:"Bebas Neue","Arial Narrow",Arial,sans-serif;font-weight:400;text-transform:uppercase;letter-spacing:.5px;line-height:.95;margin:0}
.dot{color:var(--r)}
.top{background:var(--y);border-bottom:6px solid var(--r)}
.top-in{max-width:640px;margin:0 auto;padding:12px 16px;display:flex;align-items:center;gap:12px}
.top img{height:48px;width:auto;display:block}
.store{font:800 17px Montserrat,Arial,sans-serif;line-height:1}
.store small{display:block;font-weight:600;font-size:11px;letter-spacing:.06em;text-transform:uppercase;margin-top:3px}
.call{margin-left:auto;background:var(--k);color:var(--y);text-decoration:none;font:700 14px Montserrat,Arial,sans-serif;padding:10px 14px;border-radius:999px;white-space:nowrap}
.hero{background:var(--y);padding:18px 16px 72px;overflow:hidden}
.hero-in{max-width:640px;margin:0 auto;position:relative;min-height:190px;padding-right:150px}
.frank{position:absolute;right:-8px;bottom:-72px;width:170px;height:auto;pointer-events:none}
.bubble{display:inline-block;position:relative;background:#fff;color:var(--r);font-family:"Bebas Neue",Arial,sans-serif;font-size:24px;letter-spacing:.3px;padding:6px 14px;border-radius:14px;margin-bottom:10px;line-height:1.1}
.bubble:after{content:"";position:absolute;right:-12px;top:12px;border:8px solid transparent;border-left:14px solid #fff}
.hero .h{font-size:54px}
.sub{color:var(--r);font-family:"Bebas Neue",Arial,sans-serif;font-size:26px;letter-spacing:.4px;margin-top:6px;line-height:1.05}
main{max-width:640px;margin:-52px auto 0;padding:0 16px 24px;position:relative;z-index:2}
.card{background:#fff;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.12);padding:20px;margin-bottom:16px}
.veh{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-weight:600}
.plate{font:700 15px Montserrat,Arial,sans-serif;letter-spacing:2px;background:#fff;border:2px solid var(--k);border-radius:6px;padding:2px 9px}
table{width:100%;border-collapse:collapse;margin-top:12px}
td{padding:12px 0;border-bottom:1px solid var(--line);vertical-align:top;font-size:15px}td.a{text-align:right;white-space:nowrap;padding-left:12px;font-weight:700}
.tot{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:16px}
.tot .lbl{font-weight:700}.tot .lbl small{display:block;font-weight:500;color:var(--muted);font-size:12px}
.oval{position:relative;background:var(--y);border:4px solid var(--r);border-radius:50%;padding:14px 26px;transform:rotate(-4deg);text-align:center;min-width:150px}
.oval .h{font-size:44px;line-height:1}
.note{font-size:13px;color:var(--muted);margin-top:14px}
.sec{font-size:38px;margin-bottom:4px}
.lead{color:var(--muted);font-size:14px;margin:0 0 6px}
label{display:block;font-weight:700;font-size:13px;text-transform:uppercase;letter-spacing:.05em;margin-top:18px}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}
.chip{border:2px solid var(--k);background:#fff;color:var(--k);border-radius:10px;padding:10px 12px;font:600 14px Montserrat,Arial,sans-serif;cursor:pointer;text-align:left;line-height:1.2}
.chip b{display:block;font-family:"Bebas Neue",Arial,sans-serif;font-weight:400;font-size:21px;letter-spacing:.4px}
.chip.on{background:var(--k);color:var(--y)}
input,textarea{width:100%;padding:12px;border:2px solid var(--line);border-radius:10px;font:16px Montserrat,Arial,sans-serif;margin-top:6px;background:#fff}
input:focus,textarea:focus{outline:none;border-color:var(--k)}
.btn{display:block;width:100%;border:0;border-radius:10px;padding:15px 16px;cursor:pointer;margin-top:12px;font-family:"Bebas Neue",Arial,sans-serif;font-size:26px;letter-spacing:.6px;line-height:1}
.btn.red{background:var(--r);color:#fff;box-shadow:0 6px 0 #9e1d14}
.btn.red:active{transform:translateY(3px);box-shadow:0 3px 0 #9e1d14}
.btn.out{background:#fff;color:var(--k);border:3px solid var(--k);font-size:22px}a.btn{text-decoration:none;text-align:center}
.btn:disabled{opacity:.45;cursor:default;box-shadow:none}
.err{color:var(--r);font-weight:700;margin-top:10px}
.trust{display:grid;gap:10px;margin:4px 0 16px}
.trust div{display:flex;gap:10px;align-items:flex-start;font-size:14px;font-weight:500}
.trust i{flex:none;width:22px;height:22px;border-radius:50%;background:var(--r);color:#fff;font-style:normal;font-weight:800;display:grid;place-items:center;font-size:13px;margin-top:1px}
.done{text-align:center}.done .h{font-size:52px}.done .big{width:76px;height:76px;margin:6px auto 10px;border-radius:50%;background:var(--r);color:#fff;display:grid;place-items:center;font-size:44px;font-weight:800;line-height:1}
.tag{font-family:"Bebas Neue",Arial,sans-serif;font-size:30px;text-align:center;margin:18px 0 6px}
footer{background:var(--k);color:#fff;padding:22px 16px 30px;font-size:14px}
footer .in{max-width:640px;margin:0 auto;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
footer b{color:var(--y)}footer a{color:#fff}
.hide{display:none}
@media(max-width:420px){.hero-in{padding-right:118px;min-height:170px}.frank{width:140px}.bubble{font-size:20px}.hero .h{font-size:46px}.oval .h{font-size:36px}.oval{min-width:120px;padding:12px 18px}}
</style></head><body>
<div class="top"><div class="top-in"><img src="/brand/midas-logo.png" alt="Midas"><div class="store">Ipswich<span class="dot">.</span><small>You can rely on us</small></div><a class="call" id="call" href="tel:0734442200">📞 Call us</a></div></div>
<div class="hero"><div class="hero-in"><div class="bubble" id="bub">Let’s be frank about…</div><h1 class="h" id="h1">Pick a day that suits<span class="dot">.</span></h1><div class="sub" id="hi">Loading…</div><img class="frank" src="/brand/frank.webp" alt="Frank from Midas"></div></div>
<main id="m"><div class="card">Loading…</div></main>
<footer><div class="in" id="ft"><div><b>Midas Ipswich</b><br><span id="ad"></span></div></div></footer>
<script>
const T=location.pathname.split("/").pop(),M=document.getElementById("m");
const e=(s)=>String(s??"").replace(/[&<>"]/g,(c)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const t12=(t)=>{const[h,m]=t.split(":").map(Number);return((h+11)%12+1)+(m?":"+String(m).padStart(2,"0"):"")+(h<12?"am":"pm")};
const tel=(p)=>"tel:"+String(p||"").replace(/[^0-9+]/g,"");
let D=null,pick={day:null,time:null};
async function load(){const r=await fetch("/api/public/booking/"+T+(location.search.includes("staff=1")?"?staff=1":""));D=await r.json().catch(()=>({error:"Something went wrong – please call us."}));
 if(!r.ok){document.getElementById("hi").textContent="";M.innerHTML='<div class="card"><h2 class="h sec">Link not found<span class="dot">.</span></h2><p>'+e(D.error)+'</p></div>';return}
 const s=D.shop;document.getElementById("call").href=tel(s.phone);document.getElementById("ad").innerHTML=e(s.address||"")+(s.address?"<br>":"")+'<a href="'+tel(s.phone)+'">'+e(s.phone)+'</a>';render()}
function dayLabel(l){const p=String(l).split(" ");return{d:p[0],rest:p.slice(1,3).join(" ")}}
function trust(){return '<div class="trust"><div><i>✓</i>No extra work without your OK – we always call first.</div><div><i>✓</i>Logbook services that keep your new-car warranty.</div><div><i>✓</i>Our work is guaranteed at Midas stores nationwide.</div></div><div class="tag">You can rely on us<span class="dot">.</span></div>'}
function render(){const s=D.shop;
 if(D.booked){document.getElementById("bub").textContent="Good as done!";document.getElementById("h1").innerHTML='You\\'re booked in<span class="dot">.</span>';document.getElementById("hi").textContent="Thanks "+D.first_name+" – see you soon";
  M.innerHTML='<div class="card done"><div class="big">✓</div><h2 class="h">See you then<span class="dot">.</span></h2><p style="font-size:18px">'+e(D.booked.label)+'<br>Drop off from <b>'+e(D.booked.time)+'</b></p><p class="note">'+(D.booked.confirmed?"Your booking is confirmed.":"We\\'ll text you shortly to confirm.")+' Running early or dropping off the night before? Use the key drop box on the wall of the building.</p><a class="btn out" href="'+tel(s.phone)+'">Questions? Call '+e(s.phone)+'</a></div>'+trust();return}
 document.getElementById("hi").textContent=D.wanted?"Sorry "+D.first_name+" – "+dayLabel(D.wanted).d+" is full":"Hi "+D.first_name+" – thanks for your enquiry";
 let h='<div class="card"><div class="veh">'+(D.vehicle?e(D.vehicle):"Your vehicle")+(D.rego?' <span class="plate">'+e(D.rego)+'</span>':"")+'</div>'+(D.service?'<p style="margin:10px 0 0;font-weight:700">'+e(D.service)+'</p>':"")+(D.message?'<p class="note" style="margin-top:8px">“'+e(D.message)+'”</p>':"")+(D.wanted?'<p class="note">We\\'re fully booked on '+e(D.wanted)+'. Here are the next days we can fit you in.</p>':"")+'</div>';
 h+='<div class="card"><h2 class="h sec">Pick a day<span class="dot">.</span></h2><p class="lead">Drop your car off in the morning and we\\'ll call you when it\\'s ready.</p>';
 if(!D.days.length)h+='<p>We don\\'t have any spots online right now – please give us a call.</p><a class="btn red" href="'+tel(s.phone)+'">Call '+e(s.phone)+'</a>';
 else{h+='<label>Day</label><div class="chips">'+D.days.map(d=>{const x=dayLabel(d.label);return '<button class="chip'+(pick.day===d.day?" on":"")+'" data-d="'+d.day+'"><b>'+e(x.d)+'</b>'+e(x.rest)+'</button>'}).join("")+'</div>';
  const day=D.days.find(d=>d.day===pick.day);
  if(day)h+='<label>Drop-off time</label><div class="chips">'+day.times.map(t=>'<button class="chip'+(pick.time===t?" on":"")+'" data-t="'+t+'"><b>'+t12(t)+'</b></button>').join("")+'</div>';
  h+='<label for="nb">When do you need it back? <span style="font-weight:500;text-transform:none">(optional)</span></label><input id="nb" placeholder="e.g. 3pm, end of day">';
  h+='<label for="nt">Anything we should know? <span style="font-weight:500;text-transform:none">(optional)</span></label><textarea id="nt" rows="2"></textarea>';
  h+='<button class="btn red" id="bk"'+(pick.day&&pick.time?"":" disabled")+'>'+(day&&pick.time?"Book "+e(dayLabel(day.label).d)+" "+t12(pick.time):"Book my car in")+'</button><div class="err" id="er"></div>'}
 h+='</div>'+trust();M.innerHTML=h}
async function book(){const er=document.getElementById("er");er.textContent="";document.querySelectorAll(".btn").forEach(b=>b.disabled=true);
 const body={day:pick.day,time:pick.time,needed_back:(document.getElementById("nb")||{}).value||"",notes:(document.getElementById("nt")||{}).value||""};
 const r=await fetch("/api/public/booking/"+T,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)}).catch(()=>({ok:false,status:0,json:()=>({})}));const j=await Promise.resolve(r.json()).catch(()=>({}));
 if(!r.ok){er.textContent=j.error||"Something went wrong – please call us.";document.querySelectorAll(".btn").forEach(b=>b.disabled=false);if(r.status===409)load();return}
 window.scrollTo({top:0,behavior:"smooth"});load()}
M.addEventListener("click",(ev)=>{const d=ev.target.closest("[data-d]"),t=ev.target.closest("[data-t]");
 const keep=()=>({nb:document.getElementById("nb")?.value,nt:document.getElementById("nt")?.value}),put=(k)=>{if(k.nb)document.getElementById("nb").value=k.nb;if(k.nt)document.getElementById("nt").value=k.nt};
 if(d){const k=keep();pick.day=d.dataset.d;pick.time=null;render();put(k)}else if(t){const k=keep();pick.time=t.dataset.t;render();put(k)}
 else if(ev.target.closest("#bk"))book()});
load();
</script></body></html>`;


// ---------------------------------------------------------------- authorisations: branded online approval (urgent / recommended / can do today)
const authLink = (q) => `${PUBLIC_BASE}/a/${q.auth_token}`;
const PRIORITIES = ["booked", "urgent", "recommended", ""];
// group the estimate lines into the things a customer can tick: each package is one unit, every other line is its own unit
function authUnits(q) {
  const lines = Array.isArray(q.line_items) ? q.line_items : [];
  const units = [], byKey = {}, notes = [];
  lines.forEach((l, idx) => {
    if (!l) return;
    const g = l.packageGroup && l.packageGroup !== "Additional Items" ? String(l.packageGroup) : "";
    if (l.isRemark) {
      if (!l.description) return;
      if (g && byKey["g:" + g]) byKey["g:" + g].notes.push(String(l.description)); else notes.push(String(l.description));
      return;
    }
    if (!l.description && !Number(l.price)) return;
    const key = g ? "g:" + g : "i:" + idx;
    let u = byKey[key];
    if (!u) {
      u = byKey[key] = { key, title: g ? groupTitle(g) : String(l.description || "Item"), isPackage: !!g, lines: [], idx: [], notes: [], total: 0,
        priority: PRIORITIES.includes(l.priority) ? l.priority : "", today: !!l.today };
      units.push(u);
    }
    const qty = Number(l.quantity) || 1, amt = qty * (Number(l.price) || 0);
    u.lines.push({ description: String(l.description || ""), quantity: qty, amount: amt });
    u.idx.push(idx); u.total += amt;
    if (l.priority && !u.priority) u.priority = l.priority;
    if (l.today) u.today = true;
  });
  const order = { booked: 0, urgent: 1, recommended: 2, "": 3 };
  units.sort((a, b) => order[a.priority] - order[b.priority]);
  return { units, notes };
}
// write staff choices (priority / today) back onto the estimate's lines
function applyAuthMeta(q, meta) {
  if (!Array.isArray(meta)) return q.line_items;
  const lines = (Array.isArray(q.line_items) ? q.line_items : []).map((l) => (l ? { ...l } : l));
  const { units } = authUnits({ ...q, line_items: lines });
  for (const m of meta) {
    const u = units.find((x) => x.key === m.key); if (!u) continue;
    for (const i of u.idx) { lines[i].priority = PRIORITIES.includes(m.priority) ? m.priority : ""; lines[i].today = !!m.today; }
    // keep package remarks in step too
    if (u.isPackage) lines.forEach((l) => { if (l && l.isRemark && l.packageGroup === u.key.slice(2)) { l.priority = lines[u.idx[0]].priority; l.today = lines[u.idx[0]].today; } });
  }
  return lines;
}
function authCounts(units) {
  return { urgent: units.filter((u) => u.priority === "urgent").length, rec: units.filter((u) => u.priority === "recommended").length, today: units.filter((u) => u.today).length };
}
function authSmsText(q, s, units) {
  const c = authCounts(units), car = vehicleOf(q);
  const parts = [];
  if (c.urgent) parts.push(`${c.urgent} item${c.urgent > 1 ? "s need" : " needs"} urgent attention`);
  if (c.rec) parts.push(`${c.rec} ${c.rec > 1 ? "are" : "is"} recommended`);
  const found = parts.length ? ` ${parts.join(" and ")}.` : "";
  const today = c.today ? ` We can do ${c.today === units.length ? (units.length > 1 ? "all of it" : "it") : c.today + " of them"} today${q.auth_ready_by ? `, ready by ${q.auth_ready_by}` : ""}.` : "";
  return `Hi ${firstName(q)}, we've checked your ${car || "vehicle"}.${found}${today}\nPlease review and approve the work here: ${authLink(q)}\nQuestions? Call us on ${s.phone || "07 3444 2200"}.`;
}
function authEmailHtml(q, s, units, notes) {
  const c = authCounts(units);
  const sec = (title, colour, sub, list) => !list.length ? "" : `
    <tr><td colspan="2" style="padding:16px 0 0"><div style="background:${colour};color:#fff;font:700 15px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;padding:9px 12px;border-radius:6px 6px 0 0">${title}</div>
      <div style="border:2px solid ${colour};border-top:0;border-radius:0 0 6px 6px;padding:4px 12px 8px">${sub ? `<div style="color:#666;font-size:13px;padding:6px 0 2px">${sub}</div>` : ""}
      <table role="presentation" style="width:100%;border-collapse:collapse">${list.map((u) => `<tr><td style="padding:9px 0;border-bottom:1px solid #eee">${esc(u.title)}${u.today ? ' <span style="background:#FFEB00;color:#000;font:700 11px Arial,sans-serif;padding:2px 7px;border-radius:99px;white-space:nowrap">⚡ CAN DO TODAY</span>' : ""}</td><td style="padding:9px 0 9px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:700;white-space:nowrap">${$2(u.total)}</td></tr>`).join("")}</table></div></td></tr>`;
  const total = units.reduce((a, u) => a + u.total, 0);
  return `<div style="background:#f3f3f3;padding:16px 0"><div style="max-width:620px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;font:15px/1.5 Arial,sans-serif;color:#222">
    <div style="background:#FFEB00;border-bottom:6px solid #DA291C;padding:12px 20px"><table role="presentation" style="width:100%;border-collapse:collapse"><tr><td style="width:120px"><img src="${PUBLIC_BASE}/brand/midas-logo.png" alt="Midas" width="110" style="display:block;width:110px;height:auto"></td><td style="font:800 18px Arial,sans-serif;color:#000">Ipswich<span style="color:#DA291C">.</span><div style="font:700 11px Arial,sans-serif;letter-spacing:.08em;text-transform:uppercase">You can rely on us</div></td></tr></table></div>
    <div style="padding:18px 20px 0;font:700 30px Impact,'Arial Narrow',Arial,sans-serif;text-transform:uppercase;letter-spacing:.5px">Your approval needed<span style="color:#DA291C">.</span></div>
    <div style="padding:12px 20px 16px">
      <p>Hi ${esc(firstName(q))}, we've checked your ${esc(vehicleOf(q) || "vehicle")}. Here's what we found${c.today ? ` – the items marked <b>⚡ Can do today</b> we can finish today${q.auth_ready_by ? `, ready by <b>${esc(q.auth_ready_by)}</b>` : ""}` : ""}.</p>
      ${q.auth_note ? `<p style="background:#fafafa;border-left:4px solid #DA291C;padding:8px 12px">${esc(q.auth_note)}</p>` : ""}
      <table role="presentation" style="width:100%;border-collapse:collapse">
        ${sec("🔒 Already booked in – price locked in", "#373A36", "This work is already booked – no need to approve it again.", units.filter((u) => u.priority === "booked"))}
        ${sec("🔴 Needs urgent attention", "#DA291C", "Safety or reliability – we strongly recommend fixing these now.", units.filter((u) => u.priority === "urgent"))}
        ${sec("🟠 Recommended", "#EE7D00", "Not urgent today, but worth doing soon.", units.filter((u) => u.priority === "recommended"))}
        ${sec("Your service", "#111111", "", units.filter((u) => !u.priority))}
        <tr><td style="padding:14px 0 0;font-weight:700">Total if you approve everything (inc GST)</td><td style="padding:14px 0 0;text-align:right;font-weight:700;font-size:18px">${$2(total)}</td></tr>
      </table>
      ${notes.length ? `<p style="color:#666;font-size:13px">${notes.map(esc).join("<br>")}</p>` : ""}
      <p style="text-align:center;margin:22px 0"><a href="${authLink(q)}" style="background:#DA291C;color:#fff;text-decoration:none;font-weight:700;padding:15px 30px;border-radius:8px;display:inline-block;text-transform:uppercase;letter-spacing:.04em">Review &amp; approve</a></p>
      <p style="color:#666;font-size:13px">You choose exactly which items we go ahead with. Questions? Reply to this email or call us on ${esc(s.phone || "07 3444 2200")}.</p>
    </div>
    <div style="background:#fafafa;border-top:1px solid #eee;padding:10px 20px;color:#888;font-size:12px">${esc(s.name || "Midas Ipswich")}${s.address ? " · " + esc(s.address) : ""}</div>
  </div></div>`;
}

// staff: preview or send the approval request (also saves the urgent / recommended / today marks)
app.post("/api/quotes/:id/send-auth", async (req, reply) => {
  const who = await requireStaffUser(req, reply); if (!who) return;
  const s = await getSettings();
  const { data: q, error } = await db.from("quotes").select("*").eq("id", req.params.id).maybeSingle();
  if (error || !q) return reply.code(404).send({ error: "Estimate not found" });
  if (q.auth_signed_at) return reply.code(409).send({ error: "The customer has already signed this one." });
  const b = req.body || {};
  if (!q.auth_token) { q.auth_token = crypto.randomUUID().replace(/-/g, ""); await db.from("quotes").update({ auth_token: q.auth_token }).eq("id", q.id); }
  if (Array.isArray(b.units)) q.line_items = applyAuthMeta(q, b.units);
  if (b.ready_by !== undefined) q.auth_ready_by = String(b.ready_by || "").slice(0, 40) || null;
  if (b.note !== undefined) q.auth_note = String(b.note || "").slice(0, 400) || null;
  const { units, notes } = authUnits(q);
  const via = ["sms", "email", "both"].includes(b.via) ? b.via : (q.email ? "both" : "sms");
  const smsText = String(b.sms_text || authSmsText(q, s, units));
  const subject = `Your approval needed – ${vehicleOf(q) || "your vehicle"} – Midas Ipswich`;
  const html = authEmailHtml(q, s, units, notes);
  // save the marks as soon as they change, so the customer's link always matches what staff picked (even before sending)
  if (Array.isArray(b.units) || b.ready_by !== undefined || b.note !== undefined)
    await db.from("quotes").update({ line_items: q.line_items, auth_ready_by: q.auth_ready_by || null, auth_note: q.auth_note || null }).eq("id", q.id);
  if (b.preview) return reply.send({ ok: true, via, to_phone: normPhone(q.phone), to_email: q.email || "", sms_text: smsText, auto_sms: authSmsText(q, s, units), subject, email_html: html, link: authLink(q),
    units: units.map((u) => ({ key: u.key, title: u.title, total: u.total, priority: u.priority, today: u.today, lines: u.lines.length })), ready_by: q.auth_ready_by || "", note: q.auth_note || "", sent_at: q.auth_sent_at, viewed_at: q.auth_viewed_at });
  // save the marks first so the customer's page matches what was sent
  await db.from("quotes").update({ line_items: q.line_items, auth_ready_by: q.auth_ready_by || null, auth_note: q.auth_note || null }).eq("id", q.id);
  const sent = [];
  if ((via === "sms" || via === "both") && q.phone) {
    const r = await sendSms({ to: q.phone, body: smsText, template: "AUTH", sentBy: who.name });
    if (!r.ok) return reply.code(500).send({ error: r.error || "SMS failed" });
    sent.push("sms");
  }
  if ((via === "email" || via === "both") && q.email) {
    const r = await sendEmail({ to: q.email, subject, html });
    if (!r.ok) return reply.code(500).send({ error: "Email failed" });
    sent.push("email");
  }
  if (!sent.length) return reply.code(400).send({ error: "No mobile number or email on this estimate" });
  const nowIso = new Date().toISOString();
  await db.from("quotes").update({ auth_sent_at: nowIso, auth_sent_via: sent.join("+"),
    ...(!q.status || q.status === "Pending" ? { status: "Sent", sent_at: q.sent_at || nowIso, sent_via: q.sent_via || sent.join("+") } : {}) }).eq("id", q.id);
  reply.send({ ok: true, sent, link: authLink(q) });
});

// ---- customer side, found only by the private link
async function quoteByAuthToken(token) {
  if (!/^[a-f0-9]{16,64}$/i.test(String(token || ""))) return null;
  const { data } = await db.from("quotes").select("*").eq("auth_token", token).maybeSingle();
  return data || null;
}
app.get("/api/public/auth/:token", async (req, reply) => {
  const q = await quoteByAuthToken(req.params.token);
  if (!q) return reply.code(404).send({ error: "This link isn't valid. Please call us." });
  const s = await getSettings();
  if (!q.auth_viewed_at && !req.query?.staff) db.from("quotes").update({ auth_viewed_at: new Date().toISOString() }).eq("id", q.id).then(() => {}, () => {});
  const { units, notes } = authUnits(q);
  let signed = null;
  if (q.auth_signed_at && q.auth_id) {
    const { data: w } = await db.from("work_authorizations").select("authorized_items,authorized_total,declined_items,name").eq("id", q.auth_id).maybeSingle();
    if (w) {
      const dk = new Set((w.declined_items || []).map((u) => u.key));
      signed = { at: q.auth_signed_at, total: Number(w.authorized_total || 0), approved: Array.isArray(w.declined_items) ? units.filter((u) => !dk.has(u.key)).map((u) => u.title) : (w.authorized_items || []).filter((l) => !l.isRemark).map((l) => l.description), declined: (w.declined_items || []).map((u) => u.title || u) };
    }
  } else if (q.auth_signed_at) signed = { at: q.auth_signed_at, total: null, approved: [], declined: [] };
  reply.header("Cache-Control", "no-store").send({
    first_name: firstName(q), given: q.given || "", surname: q.surname || "", phone: q.phone || "", email: q.email || "",
    vehicle: q.vehicle || "", rego: q.rego ? String(q.rego).toUpperCase() : "",
    units: units.map((u) => ({ key: u.key, title: u.title, isPackage: u.isPackage, notes: u.notes, total: u.total, priority: u.priority, today: u.today })),
    notes, ready_by: q.auth_ready_by || "", note: q.auth_note || "", show_line_prices: !!s.publish_line_item_prices,
    terms: s.works_terms || "", signed,
    shop: { name: s.name || "Midas Ipswich", phone: s.phone || "07 3444 2200", address: s.address || "" },
  });
});
app.post("/api/public/auth/:token/sign", { bodyLimit: 2 * 1024 * 1024 }, async (req, reply) => {
  const q = await quoteByAuthToken(req.params.token);
  if (!q) return reply.code(404).send({ error: "This link isn't valid. Please call us." });
  if (q.auth_signed_at) return reply.send({ ok: true, already: true });
  const b = req.body || {};
  const given = String(b.given || "").trim().slice(0, 60), surname = String(b.surname || "").trim().slice(0, 60);
  const phone = String(b.phone || "").trim().slice(0, 30), email = String(b.email || "").trim().slice(0, 120);
  if (!given || !surname || !phone) return reply.code(400).send({ error: "Please fill in your name and phone." });
  if (!(b.ack_quote && b.ack_faults && b.ack_terms)) return reply.code(400).send({ error: "Please tick the three acknowledgements." });
  if (!b.age_ack) return reply.code(400).send({ error: "Please confirm you are 18 or older." });
  if (!b.esign_consent) return reply.code(400).send({ error: "Please agree to signing electronically." });
  const sig = String(b.signature || "");
  if (!/^data:image\/png;base64,[A-Za-z0-9+/=]{200,}$/.test(sig) || sig.length > 1500000) return reply.code(400).send({ error: "Please sign in the box." });
  const { units } = authUnits(q);
  const keys = new Set(Array.isArray(b.keys) ? b.keys.map(String) : []);
  units.forEach((u) => { if (u.priority === "booked") keys.add(u.key); }); // already booked work is always included
  const ok = units.filter((u) => keys.has(u.key)), no = units.filter((u) => !keys.has(u.key));
  const lines = Array.isArray(q.line_items) ? q.line_items : [];
  const approvedLines = ok.flatMap((u) => u.idx.map((i) => lines[i]));
  const total = ok.reduce((a, u) => a + u.total, 0);
  // claim first so a double tap can't sign twice
  const { data: claimed } = await db.from("quotes").update({ auth_signed_at: new Date().toISOString() }).eq("id", q.id).is("auth_signed_at", null).select("id");
  if (!claimed?.length) return reply.send({ ok: true, already: true });
  const s = await getSettings();
  const rec = {
    ts: Date.now(), given, surname, name: `${given} ${surname}`.trim(), phone, email,
    ack_quote: true, ack_faults: true, ack_terms: true, age_ack: true, esign_consent: true, signature: sig,
    terms_text_snapshot: s.works_terms || "", business_name_snapshot: s.location || s.name || "Midas Ipswich",
    business_address_snapshot: s.address || "", business_phone_snapshot: s.phone || "",
    quote_id: q.id, authorized_items: approvedLines, authorized_total: total, source: "online",
    declined_items: no.map((u) => ({ key: u.key, title: u.title, total: u.total, priority: u.priority })),
  };
  const { data: w, error } = await db.from("work_authorizations").insert(rec).select("id").single();
  if (error) {
    console.error("auth insert", error.message);
    await db.from("quotes").update({ auth_signed_at: null }).eq("id", q.id);
    return reply.code(500).send({ error: "Sorry, we couldn't save that. Please try again or call us." });
  }
  await db.from("quotes").update({ auth_id: w.id }).eq("id", q.id);
  const li = (list) => list.map((u) => `<li>${esc(u.title)} – ${$2(u.total)}${u.priority ? ` <small>(${u.priority})</small>` : ""}${u.today ? " ⚡ today" : ""}</li>`).join("");
  if (s.notify_email) sendEmail({
    to: s.notify_email, subject: `✍️ Work approved – ${rec.name} ${vehicleOf(q)} – ${$2(total)}`,
    html: `<p><b>${esc(rec.name)}</b> (${esc(phone)}) signed the Authority to Commence Works online for ${esc(vehicleOf(q) || "their vehicle")}.</p>
      <p><b>Approved – ${$2(total)} inc GST</b></p><ul>${li(ok) || "<li>Nothing ticked – terms only</li>"}</ul>
      ${no.length ? `<p><b>Not approved</b></p><ul>${li(no)}</ul>` : ""}<p>The signed copy is in Authorisations in the app.</p>`,
  }).catch(() => {});
  reply.send({ ok: true });
});
app.get("/a/:token", async (req, reply) => {
  reply.header("Cache-Control", "no-store").header("X-Robots-Tag", "noindex").type("text/html").send(AUTH_PAGE);
});
const AUTH_PAGE = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex">
<title>Your approval – Midas Ipswich</title>
<link rel="icon" href="/brand/midas-logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--y:#FFEB00;--r:#DA291C;--o:#EE7D00;--k:#000;--g:#373A36;--soft:#F5F5F2;--line:#E2E2DC;--muted:#5d5f5a}
*{box-sizing:border-box}html,body{margin:0}
body{background:var(--soft);color:var(--k);font:16px/1.55 Montserrat,Arial,sans-serif;-webkit-font-smoothing:antialiased;padding-bottom:84px}
.h{font-family:"Bebas Neue","Arial Narrow",Arial,sans-serif;font-weight:400;text-transform:uppercase;letter-spacing:.5px;line-height:.95;margin:0}
.dot{color:var(--r)}
.top{background:var(--y);border-bottom:6px solid var(--r)}
.top-in{max-width:640px;margin:0 auto;padding:12px 16px;display:flex;align-items:center;gap:12px}
.top img{height:48px;width:auto;display:block}
.store{font:800 17px Montserrat,Arial,sans-serif;line-height:1}
.store small{display:block;font-weight:600;font-size:11px;letter-spacing:.06em;text-transform:uppercase;margin-top:3px}
.call{margin-left:auto;background:var(--k);color:var(--y);text-decoration:none;font:700 14px Montserrat,Arial,sans-serif;padding:10px 14px;border-radius:999px;white-space:nowrap}
.hero{background:var(--y);padding:18px 16px 72px;overflow:hidden}
.hero-in{max-width:640px;margin:0 auto;position:relative;min-height:190px;padding-right:150px}
.frank{position:absolute;right:-8px;bottom:-72px;width:170px;height:auto;pointer-events:none}
.bubble{display:inline-block;position:relative;background:#fff;color:var(--r);font-family:"Bebas Neue",Arial,sans-serif;font-size:24px;letter-spacing:.3px;padding:6px 14px;border-radius:14px;margin-bottom:10px;line-height:1.1}
.bubble:after{content:"";position:absolute;right:-12px;top:12px;border:8px solid transparent;border-left:14px solid #fff}
.hero .h{font-size:54px}
.sub{color:var(--r);font-family:"Bebas Neue",Arial,sans-serif;font-size:26px;letter-spacing:.4px;margin-top:6px;line-height:1.05}
main{max-width:640px;margin:-52px auto 0;padding:0 16px 24px;position:relative;z-index:2}
.card{background:#fff;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.12);padding:20px;margin-bottom:16px}
.veh{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-weight:600}
.plate{font:700 15px Montserrat,Arial,sans-serif;letter-spacing:2px;background:#fff;border:2px solid var(--k);border-radius:6px;padding:2px 9px}
.today{background:var(--k);color:#fff;border-radius:16px;padding:18px 20px;margin-bottom:16px;box-shadow:0 10px 30px rgba(0,0,0,.18)}
.today .h{color:var(--y);font-size:34px}
.today p{margin:6px 0 0;font-size:15px}
.today button{margin-top:12px;background:var(--y);color:var(--k);border:0;border-radius:10px;padding:11px 14px;font-family:"Bebas Neue",Arial,sans-serif;font-size:21px;letter-spacing:.5px;cursor:pointer;width:100%}
.grp{border-radius:16px;overflow:hidden;background:#fff;box-shadow:0 10px 30px rgba(0,0,0,.12);margin-bottom:16px;border:3px solid var(--c)}
.grp-h{background:var(--c);color:#fff;padding:12px 16px}
.grp-h .h{font-size:30px}
.grp-h small{display:block;font-size:13px;font-weight:600;opacity:.95;margin-top:2px}
.grp.booked{--c:var(--g)}.it.lock{cursor:default}.it.lock .box{background:var(--g);border-color:var(--g);font-size:15px}.it.lock .box:after{content:""}.pill.lockp{background:#e8e8e3;color:var(--g)}
.grp.urgent{--c:var(--r)}.grp.rec{--c:var(--o)}.grp.std{--c:var(--k)}
.it{display:flex;gap:12px;align-items:flex-start;padding:14px 16px;border-top:1px solid var(--line);cursor:pointer;user-select:none}
.grp-h + .it{border-top:0}
.box{flex:none;width:30px;height:30px;border-radius:8px;border:3px solid var(--k);display:grid;place-items:center;font-weight:800;font-size:18px;color:#fff;margin-top:1px;background:#fff}
.it.on .box{background:var(--c);border-color:var(--c)}
.it.on .box:after{content:"✓"}
.it .t{flex:1;font-weight:700;font-size:16px;line-height:1.3}
.it .t small{display:block;font-weight:500;color:var(--muted);font-size:13px;margin-top:3px}
.it .p{font-weight:800;white-space:nowrap;font-size:16px}
.it:not(.on) .t,.it:not(.on) .p{opacity:.55}
.pill{display:inline-block;background:var(--y);color:var(--k);font:800 11px Montserrat,Arial,sans-serif;letter-spacing:.04em;padding:3px 8px;border-radius:99px;margin-top:6px;white-space:nowrap}
.pill.later{background:#ececec;color:var(--muted)}
.notes{font-size:13px;color:var(--muted);margin:0 16px 14px}
.sum{display:flex;align-items:center;justify-content:space-between;gap:12px}
.oval{position:relative;background:var(--y);border:4px solid var(--r);border-radius:50%;padding:12px 22px;transform:rotate(-4deg);text-align:center;min-width:140px}
.oval .h{font-size:40px;line-height:1}
.sec{font-size:38px;margin-bottom:4px}
.lead{color:var(--muted);font-size:14px;margin:0 0 6px}
label.f{display:block;font-weight:700;font-size:13px;text-transform:uppercase;letter-spacing:.05em;margin-top:16px}
input[type=text],input[type=tel],input[type=email]{width:100%;padding:12px;border:2px solid var(--line);border-radius:10px;font:16px Montserrat,Arial,sans-serif;margin-top:6px;background:#fff}
input:focus{outline:none;border-color:var(--k)}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.ack{display:flex;gap:10px;align-items:flex-start;margin-top:12px;font-size:14px;line-height:1.4;cursor:pointer}
.ack input{flex:none;width:22px;height:22px;margin-top:1px;accent-color:var(--r)}
details{margin-top:14px;border:2px solid var(--line);border-radius:10px;padding:10px 12px}
details summary{font-weight:700;cursor:pointer}
.terms{white-space:pre-wrap;font-size:12.5px;line-height:1.5;color:#333;max-height:46vh;overflow:auto;margin-top:8px}
.sigw{position:relative;margin-top:6px;border:2px dashed var(--g);border-radius:10px;background:#fff;height:170px;touch-action:none}
.sigw canvas{width:100%;height:100%;display:block;border-radius:10px}
.sigw span{position:absolute;inset:0;display:grid;place-items:center;color:#aaa;font-size:14px;pointer-events:none}
.clr{float:right;background:none;border:0;color:var(--r);font-weight:700;cursor:pointer;font-size:13px}
.btn{display:block;width:100%;border:0;border-radius:10px;padding:15px 16px;cursor:pointer;margin-top:16px;font-family:"Bebas Neue",Arial,sans-serif;font-size:26px;letter-spacing:.6px;line-height:1}
.btn.red{background:var(--r);color:#fff;box-shadow:0 6px 0 #9e1d14}
.btn.red:active{transform:translateY(3px);box-shadow:0 3px 0 #9e1d14}
.btn.out{background:#fff;color:var(--k);border:3px solid var(--k);font-size:22px}a.btn{text-decoration:none;text-align:center}
.btn:disabled{opacity:.45;cursor:default;box-shadow:none}
.err{color:var(--r);font-weight:700;margin-top:10px}
.bar{position:fixed;left:0;right:0;bottom:0;background:var(--k);color:#fff;z-index:20;box-shadow:0 -6px 20px rgba(0,0,0,.25)}
.bar-in{max-width:640px;margin:0 auto;padding:10px 16px;display:flex;align-items:center;gap:12px}
.bar b{font-family:"Bebas Neue",Arial,sans-serif;font-weight:400;font-size:30px;color:var(--y);letter-spacing:.5px}
.bar small{display:block;font-size:12px;color:#ccc}
.bar a{margin-left:auto;background:var(--r);color:#fff;text-decoration:none;font-family:"Bebas Neue",Arial,sans-serif;font-size:22px;letter-spacing:.5px;padding:10px 16px;border-radius:10px}
.done{text-align:center}.done .h{font-size:50px}.done .big{width:76px;height:76px;margin:6px auto 10px;border-radius:50%;background:var(--r);color:#fff;display:grid;place-items:center;font-size:44px;font-weight:800;line-height:1}
.done ul{text-align:left;padding-left:20px;margin:6px 0}
.trust{display:grid;gap:10px;margin:4px 0 16px}
.trust div{display:flex;gap:10px;align-items:flex-start;font-size:14px;font-weight:500}
.trust i{flex:none;width:22px;height:22px;border-radius:50%;background:var(--r);color:#fff;font-style:normal;font-weight:800;display:grid;place-items:center;font-size:13px;margin-top:1px}
.tag{font-family:"Bebas Neue",Arial,sans-serif;font-size:30px;text-align:center;margin:18px 0 6px}
footer{background:var(--k);color:#fff;padding:22px 16px 30px;font-size:14px}
footer .in{max-width:640px;margin:0 auto}
footer b{color:var(--y)}footer a{color:#fff}
@media(max-width:420px){.hero-in{padding-right:118px;min-height:170px}.frank{width:140px}.bubble{font-size:20px}.hero .h{font-size:46px}.oval .h{font-size:34px}.oval{min-width:116px;padding:10px 16px}.row2{grid-template-columns:1fr}}
</style></head><body>
<div class="top"><div class="top-in"><img src="/brand/midas-logo.png" alt="Midas"><div class="store">Ipswich<span class="dot">.</span><small>You can rely on us</small></div><a class="call" id="call" href="tel:0734442200">📞 Call us</a></div></div>
<div class="hero"><div class="hero-in"><div class="bubble" id="bub">Let’s be frank about…</div><h1 class="h" id="h1">Your car needs your OK<span class="dot">.</span></h1><div class="sub" id="hi">Loading…</div><img class="frank" src="/brand/frank.webp" alt="Frank from Midas"></div></div>
<main id="m"><div class="card">Loading…</div></main>
<div class="bar" id="bar" style="display:none"><div class="bar-in"><div><small>You're approving</small><b id="bt">$0.00</b></div><a href="#sign">Approve &amp; sign ↓</a></div></div>
<footer><div class="in"><b>Midas Ipswich</b><br><span id="ad"></span></div></footer>
<script>
const T=location.pathname.split("/").pop(),M=document.getElementById("m");
const $=(n)=>"$"+Number(n||0).toLocaleString("en-AU",{minimumFractionDigits:2,maximumFractionDigits:2});
const e=(s)=>String(s??"").replace(/[&<>"]/g,(c)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const tel=(p)=>"tel:"+String(p||"").replace(/[^0-9+]/g,"");
let D=null,on=new Set(),hasSig=false;
async function load(){const r=await fetch("/api/public/auth/"+T+(location.search.includes("staff=1")?"?staff=1":""));D=await r.json().catch(()=>({error:"Something went wrong – please call us."}));
 if(!r.ok){document.getElementById("hi").textContent="";M.innerHTML='<div class="card"><h2 class="h sec">Link not found<span class="dot">.</span></h2><p>'+e(D.error)+'</p></div>';return}
 const s=D.shop;document.getElementById("call").href=tel(s.phone);document.getElementById("ad").innerHTML=e(s.address||"")+(s.address?"<br>":"")+'<a href="'+tel(s.phone)+'">'+e(s.phone)+'</a>';
 on=new Set(D.units.filter(u=>u.priority!=="recommended").map(u=>u.key));render()}
const choice=()=>D.units.filter(u=>u.priority!=="booked");
function total(){return D.units.filter(u=>on.has(u.key)).reduce((a,u)=>a+u.total,0)}
function item(u){const sub="";
 if(u.priority==="booked")return '<div class="it on lock"><div class="box">🔒</div><div class="t">'+e(u.title)+(u.notes&&u.notes.length?'<small><i>'+u.notes.map(e).join("<br>")+'</i></small>':"")+'<br><span class="pill lockp">Price locked in</span></div><div class="p">'+$(u.total)+'</div></div>';
 const nt=u.notes&&u.notes.length?'<small><i>'+u.notes.map(e).join("<br>")+'</i></small>':"";
 const pill=u.today?'<span class="pill">⚡ Can do today'+(D.ready_by?" – ready by "+e(D.ready_by):"")+'</span>':(D.units.some(x=>x.today)?'<span class="pill later">Book in another day</span>':"");
 return '<div class="it'+(on.has(u.key)?" on":"")+'" data-k="'+e(u.key)+'" role="checkbox" aria-checked="'+on.has(u.key)+'" tabindex="0"><div class="box"></div><div class="t">'+e(u.title)+sub+nt+'<br>'+pill+'</div><div class="p">'+$(u.total)+'</div></div>'}
function group(cls,title,sub,list){if(!list.length)return"";return '<div class="grp '+cls+'"><div class="grp-h"><h2 class="h">'+title+'</h2>'+(sub?'<small>'+sub+'</small>':"")+'</div>'+list.map(item).join("")+'</div>'}
function render(){const s=D.shop,U=D.units;
 if(D.signed)return done();
 document.getElementById("hi").textContent="Hi "+D.first_name+" – here's what we found";
 const car=D.vehicle;
 let h='<div class="card"><div class="veh">'+(car?e(car):"Your vehicle")+(D.rego?' <span class="plate">'+e(D.rego)+'</span>':"")+'</div><p class="lead" style="margin-top:8px">Tick the work you\\'d like us to go ahead with, then sign at the bottom. Nothing extra gets done without your OK.</p>'+(D.note?'<p style="border-left:4px solid var(--r);padding:6px 12px;background:var(--soft);border-radius:0 8px 8px 0;margin:10px 0 0">'+e(D.note)+'</p>':"")+'</div>';
 h+=group("booked","🔒 Already booked in","This work is already booked – price locked in. Nothing to approve.",U.filter(u=>u.priority==="booked"));
 const td=U.filter(u=>u.today&&u.priority!=="booked");
 if(td.length)h+='<div class="today"><h2 class="h">⚡ What we can do today<span class="dot">.</span></h2><p>'+td.length+' of '+choice().length+' item'+(choice().length>1?"s":"")+' can be finished today'+(D.ready_by?', <b>ready by '+e(D.ready_by)+'</b>':"")+'.</p><button id="tday">Tick everything we can do today – '+$(td.reduce((a,u)=>a+u.total,0))+'</button></div>';
 h+=group("urgent","🔴 Needs urgent attention","Safety or reliability – we strongly recommend fixing these now.",U.filter(u=>u.priority==="urgent"));
 h+=group("rec","🟠 Recommended","Not urgent today, but worth doing soon.",U.filter(u=>u.priority==="recommended"));
 h+=group("std",U.some(u=>u.priority)?"Also found":"Your job","",U.filter(u=>!u.priority));
 if(D.notes.length)h+='<div class="card" style="font-size:14px;color:var(--muted)">'+D.notes.map(e).join("<br>")+'</div>';
 const t=total();
 h+='<div class="card" id="sign"><div class="sum"><div><h2 class="h sec">Approve &amp; sign<span class="dot">.</span></h2><p class="lead">'+(U.length?on.size+' of '+U.length+' item'+(U.length>1?"s":"")+' ticked':"Authority to commence works")+'</p></div>'+(U.length?'<div class="oval"><div class="h" id="ot">'+$(t)+'</div></div>':"")+'</div>'
  +'<div class="row2"><div><label class="f" for="g">First name</label><input type="text" id="g" value="'+e(D.given)+'" autocomplete="given-name"></div><div><label class="f" for="s">Last name</label><input type="text" id="s" value="'+e(D.surname)+'" autocomplete="family-name"></div></div>'
  +'<label class="f" for="p">Phone</label><input type="tel" id="p" value="'+e(D.phone)+'" autocomplete="tel">'
  +'<label class="f" for="em">Email <span style="font-weight:500;text-transform:none">(for your copy)</span></label><input type="email" id="em" value="'+e(D.email)+'" autocomplete="email">'
  +(D.terms?'<details><summary>Read the Authority to Commence Works &amp; Terms of Service</summary><div class="terms">'+e(D.terms)+'</div></details>':"")
  +'<label class="ack"><input type="checkbox" id="a1"><span>I accept the estimate for the items I have ticked, and understand I can ask for a written estimate before work starts.</span></label>'
  +'<label class="ack"><input type="checkbox" id="a2"><span>I understand more faults may be found during inspection or repair, and extra work or charges will only go ahead with my OK.</span></label>'
  +'<label class="ack"><input type="checkbox" id="a3"><span>I have read and agree to the Authority to Commence Works and Terms of Service.</span></label>'
  +'<label class="ack"><input type="checkbox" id="a4"><span>I am 18 years of age or older.</span></label>'
  +'<label class="f" style="margin-top:18px">Your signature <button class="clr" id="clr" type="button">Clear</button></label><div class="sigw"><canvas id="sig"></canvas><span id="sh">Sign here with your finger</span></div>'
  +'<label class="ack"><input type="checkbox" id="a5"><span>I agree to sign electronically instead of on paper, and know I can ask for a paper copy.</span></label>'
  +'<button class="btn red" id="go">'+(U.length?"Approve "+$(t)+" &amp; sign":"Agree &amp; sign")+'</button><div class="err" id="er"></div></div>'+trust();
 M.innerHTML=h;pad();bar()}
function bar(){const b=document.getElementById("bar");if(!D||D.signed||!D.units.length){b.style.display="none";return}b.style.display="block";const t=total();document.getElementById("bt").textContent=$(t);const ot=document.getElementById("ot");if(ot)ot.textContent=$(t);const go=document.getElementById("go");if(go)go.innerHTML="Approve "+$(t)+" &amp; sign"}
function toggle(k){const uu=D.units.find(x=>x.key===k);if(uu&&uu.priority==="booked")return;if(on.has(k))on.delete(k);else on.add(k);const el=M.querySelector('[data-k="'+CSS.escape(k)+'"]');if(el){el.classList.toggle("on",on.has(k));el.setAttribute("aria-checked",on.has(k))}const ls=M.querySelector("#sign .lead");if(ls)ls.textContent=on.size+" of "+D.units.length+" item"+(D.units.length>1?"s":"")+" ticked";bar()}
function pad(){const c=document.getElementById("sig");if(!c)return;const w=c.parentElement,r=window.devicePixelRatio||1;c.width=w.clientWidth*r;c.height=w.clientHeight*r;const x=c.getContext("2d");x.scale(r,r);x.lineWidth=2.6;x.lineCap="round";x.lineJoin="round";x.strokeStyle="#111";
 let dn=false,last=null;const pt=(ev)=>{const b=c.getBoundingClientRect();return{x:ev.clientX-b.left,y:ev.clientY-b.top}};
 c.onpointerdown=(ev)=>{dn=true;last=pt(ev);c.setPointerCapture(ev.pointerId);x.beginPath();x.arc(last.x,last.y,1.2,0,7);x.fillStyle="#111";x.fill()};
 c.onpointermove=(ev)=>{if(!dn)return;const p=pt(ev);x.beginPath();x.moveTo(last.x,last.y);x.lineTo(p.x,p.y);x.stroke();last=p;if(!hasSig){hasSig=true;document.getElementById("sh").style.display="none"}};
 c.onpointerup=c.onpointercancel=()=>{dn=false};
 document.getElementById("clr").onclick=()=>{x.clearRect(0,0,c.width,c.height);hasSig=false;document.getElementById("sh").style.display=""}}
function trust(){return '<div class="trust"><div><i>✓</i>No extra work without your OK – we always call first.</div><div><i>✓</i>Genuine quality parts, fitted by qualified technicians.</div><div><i>✓</i>Our work is guaranteed at Midas stores nationwide.</div></div><div class="tag">You can rely on us<span class="dot">.</span></div>'}
function done(){const sg=D.signed;document.getElementById("bub").textContent="Good as done!";document.getElementById("h1").innerHTML='Thanks – we\\'re on it<span class="dot">.</span>';document.getElementById("hi").textContent="Thanks "+D.first_name+" – you're all signed";
 let h='<div class="card done"><div class="big">✓</div><h2 class="h">Approved<span class="dot">.</span></h2>'+(sg.total!=null?'<p style="font-size:18px;margin:6px 0">'+$(sg.total)+' inc GST</p>':"")
  +(sg.approved&&sg.approved.length?'<p style="text-align:left;font-weight:700;margin:14px 0 0">Going ahead:</p><ul>'+sg.approved.map(x=>'<li>'+e(x)+'</li>').join("")+'</ul>':"")
  +(sg.declined&&sg.declined.length?'<p style="text-align:left;font-weight:700;margin:14px 0 0">Not this time:</p><ul>'+sg.declined.map(x=>'<li>'+e(x)+'</li>').join("")+'</ul><p class="lead" style="text-align:left">Changed your mind? Just give us a call.</p>':"")
  +(D.ready_by?'<p>We\\'ll aim to have it ready by <b>'+e(D.ready_by)+'</b> and we\\'ll call you when it is.</p>':'<p>We\\'ll call you when it\\'s ready.</p>')
  +'<a class="btn out" href="'+tel(D.shop.phone)+'">Questions? Call '+e(D.shop.phone)+'</a></div>'+trust();M.innerHTML=h;bar()}
async function sign(){const er=document.getElementById("er"),v=(i)=>(document.getElementById(i)||{}).value||"",c=(i)=>!!(document.getElementById(i)||{}).checked;er.textContent="";
 if(!v("g").trim()||!v("s").trim()||!v("p").trim()){er.textContent="Please fill in your name and phone.";return}
 if(!(c("a1")&&c("a2")&&c("a3"))){er.textContent="Please tick the three acknowledgements.";return}
 if(!c("a4")){er.textContent="Please confirm you are 18 or older.";return}
 if(!hasSig){er.textContent="Please sign in the box.";return}
 if(!c("a5")){er.textContent="Please agree to signing electronically.";return}
 const go=document.getElementById("go");go.disabled=true;go.textContent="Saving…";
 const body={keys:[...on],given:v("g").trim(),surname:v("s").trim(),phone:v("p").trim(),email:v("em").trim(),ack_quote:true,ack_faults:true,ack_terms:true,age_ack:true,esign_consent:true,signature:document.getElementById("sig").toDataURL("image/png")};
 const r=await fetch("/api/public/auth/"+T+"/sign",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)}).catch(()=>({ok:false,json:()=>({})}));const j=await Promise.resolve(r.json()).catch(()=>({}));
 if(!r.ok){er.textContent=j.error||"Something went wrong – please call us.";go.disabled=false;bar();return}
 window.scrollTo({top:0,behavior:"smooth"});load()}
M.addEventListener("click",(ev)=>{const it=ev.target.closest("[data-k]");if(it)return toggle(it.dataset.k);
 if(ev.target.closest("#tday")){D.units.forEach(u=>{if(u.today&&!on.has(u.key))toggle(u.key)});return}
 if(ev.target.closest("#go"))sign()});
M.addEventListener("keydown",(ev)=>{const it=ev.target.closest("[data-k]");if(it&&(ev.key===" "||ev.key==="Enter")){ev.preventDefault();toggle(it.dataset.k)}});
window.addEventListener("resize",()=>{if(!hasSig)pad()});
load();
</script></body></html>`;

// ---------------------------------------------------------------- daily phone report (emailed each evening)
// Costs are estimates in US$ from Twilio/Anthropic list prices – override any with Railway variables.
const COST = {
  relay: num(env.COST_RELAY_MIN) ?? 0.07,        // ConversationRelay per minute
  inbound: num(env.COST_INBOUND_MIN) ?? 0.01,    // receiving on the AU local number per minute
  outbound: num(env.COST_OUT_MIN) ?? 0.0252,     // putting calls through to the shop landline per minute
  sms: num(env.COST_SMS_SEG) ?? 0.0515,          // per SMS segment
  voicemail: num(env.COST_VM_MIN) ?? 0.0525,     // recording + transcription per minute
  aiReply: num(env.COST_AI_REPLY) ?? 0.007,      // Claude per AI reply (approx.)
  fixedDay: num(env.COST_FIXED_DAY) ?? 0.54,     // numbers + server per day
  audPerUsd: num(env.AUD_PER_USD) ?? 1.52,
};
const secsOf = (c) => (c.ended_at && c.started_at ? Math.max(0, (Date.parse(c.ended_at) - Date.parse(c.started_at)) / 1000) : 0);
const hhmm = (ts) => { const d = new Date(Date.parse(ts) + 10 * 3600e3); return time12(`${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`); };
const aud = (usd) => "$" + (usd * COST.audPerUsd).toFixed(2);

async function buildPhoneReport(day) {
  const s = await getSettings(true);
  const t0 = new Date(day + "T00:00:00+10:00").toISOString(), t1 = new Date(addDays(day, 1) + "T00:00:00+10:00").toISOString();
  const w0 = new Date(addDays(day, -7) + "T00:00:00+10:00").toISOString();
  const [callsR, prevR, smsR, vmR] = await Promise.all([
    db.from("voice_calls").select("*").gte("started_at", t0).lt("started_at", t1).order("started_at"),
    db.from("voice_calls").select("outcome,started_at").gte("started_at", w0).lt("started_at", t0),
    db.from("sms_log").select("direction,channel,segments,status,template_label,phone,message,created_at").gte("created_at", t0).lt("created_at", t1),
    db.from("phone_messages").select("caller_name,caller_phone,transcript,regarding,created_at,call_sid,source").gte("created_at", t0).lt("created_at", t1).order("created_at"),
  ]);
  const calls = callsR.data || [], prev = prevR.data || [], sms = smsR.data || [], msgs = vmR.data || [];
  const n = (o) => calls.filter((c) => (Array.isArray(o) ? o.includes(c.outcome) : c.outcome === o)).length;
  const bookingsN = n("booking"), quotesN = n("quote");
  const answered = n(["answered", "transfer", "transferred"]), missed = n("transfer_missed");
  const wlThrough = n("whitelist_put_through"), wlVm = n("whitelist_voicemail");
  const vmN = n(["voicemail", "message", "blacklist_message"]) + wlVm, hung = n(["hangup", "abusive", "error"]);
  const aiCalls = calls.filter((c) => !/^(whitelist|blacklist)/.test(String(c.outcome || "")));
  const aiSecs = aiCalls.reduce((a, c) => a + secsOf(c), 0);
  const avgSecs = aiCalls.length ? aiSecs / aiCalls.length : 0;
  // after hours + busiest hour
  const dh = dayHours(s, day), wd = new Date(day + "T00:00:00Z").getUTCDay();
  const hourOf = (ts) => new Date(Date.parse(ts) + 10 * 3600e3).getUTCHours();
  const tOf = (ts) => { const d = new Date(Date.parse(ts) + 10 * 3600e3); return `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`; };
  const afterHrs = calls.filter((c) => wd === 0 || tOf(c.started_at) < dh.start || tOf(c.started_at) > dh.end).length;
  const byHour = {}; calls.forEach((c) => { const h = hourOf(c.started_at); byHour[h] = (byHour[h] || 0) + 1; });
  const busiest = Object.entries(byHour).sort((a, b) => b[1] - a[1])[0];
  const hourName = (h) => `${((h + 11) % 12) + 1}${h < 12 ? "am" : "pm"}`;
  // last 7 days for comparison
  const prevDays = new Set(prev.map((c) => isoDay(new Date(Date.parse(c.started_at) + 10 * 3600e3)))).size || 1;
  const pAvg = (f) => Math.round((prev.filter(f).length / prevDays) * 10) / 10;
  const vs = (now, avg) => (avg ? ` (7-day avg ${avg})` : "");
  // bookings + quotes made by the AI today
  const bIds = calls.map((c) => c.booking_id).filter(Boolean), qIds = calls.map((c) => c.quote_id).filter(Boolean);
  const [{ data: bks }, { data: qts }] = await Promise.all([
    bIds.length ? db.from("bookings").select("id,booking_ref,customer_name,phone,rego,vehicle,make,model,service,booking_date,time_in,state,status").in("id", bIds) : { data: [] },
    qIds.length ? db.from("quotes").select("id,name,phone,vehicle,rego,notes").in("id", qIds) : { data: [] },
  ]);
  // call-backs: voicemails, missed transfers, and people who hung up part-way (not whitelisted suppliers)
  const cb = [];
  for (const m of msgs) cb.push(`**${hhmm(m.created_at)}** ${m.caller_name && m.caller_name !== "Unknown caller" ? m.caller_name + " · " : ""}${m.caller_phone || "withheld"} – voicemail: "${String(m.transcript || m.regarding || "").replace(/\s+/g, " ").slice(0, 160)}"`);
  const vmSids = new Set(msgs.map((m) => m.call_sid));
  for (const c of calls) {
    if (vmSids.has(c.call_sid)) continue;
    const said = Array.isArray(c.transcript) ? c.transcript.filter((m) => m.role === "user").length : 0;
    if (c.outcome === "transfer_missed" || (c.outcome === "hangup" && said >= 1)) {
      const cust = await lookupCustomer(c.from_number).catch(() => null);
      const nm = cust ? [cust.given, cust.surname].filter(Boolean).join(" ") : "";
      const lastSaid = Array.isArray(c.transcript) ? (c.transcript.filter((m) => m.role === "user").pop()?.text || "") : "";
      cb.push(`**${hhmm(c.started_at)}** ${nm ? nm + " · " : ""}${normPhone(c.from_number) || "withheld"} – ${c.outcome === "transfer_missed" ? "put through but nobody answered" : "hung up before finishing"}${lastSaid ? `: "${lastSaid.replace(/\s+/g, " ").slice(0, 120)}"` : ""}`);
    }
  }
  // what callers said before being put through (shows what the AI could handle in future)
  const throughSaid = calls.filter((c) => ["answered", "transfer", "transferred", "transfer_missed"].includes(c.outcome)).map((c) => {
    const said = Array.isArray(c.transcript) ? c.transcript.filter((m) => m.role === "user").map((m) => m.text).join(" ").replace(/\s+/g, " ").trim() : "";
    return `**${hhmm(c.started_at)}** ${said ? `"${said.slice(0, 140)}${said.length > 140 ? "…" : ""}"` : "*asked for the team straight away*"}${c.transfer_seconds ? ` · talked ${Math.ceil(c.transfer_seconds / 60)} min` : ""}`;
  });
  // SMS / email
  const smsOut = sms.filter((x) => x.channel !== "email" && x.direction === "outbound"), smsIn = sms.filter((x) => x.channel !== "email" && x.direction === "inbound");
  const segs = smsOut.reduce((a, x) => a + (Number(x.segments) || 1), 0);
  const emails = sms.filter((x) => x.channel === "email" && x.direction === "outbound").length;
  const byTpl = {}; smsOut.forEach((x) => { const k = x.template_label || "Manual"; byTpl[k] = (byTpl[k] || 0) + 1; });
  // cost estimate (US$)
  const relayMin = calls.reduce((a, c) => a + Math.max(1, Math.ceil(secsOf(c) / 60)), 0);
  const xferMin = calls.reduce((a, c) => a + Math.ceil((Number(c.transfer_seconds) || 0) / 60), 0);
  const xferKnown = calls.some((c) => c.transfer_seconds != null);
  const xferGuess = xferKnown ? xferMin : (answered + wlThrough) * 3;
  const replies = calls.reduce((a, c) => a + (Array.isArray(c.transcript) ? c.transcript.filter((m) => m.role === "assistant").length : 0), 0);
  const cost = {
    relay: relayMin * COST.relay, inbound: (relayMin + xferGuess) * COST.inbound, ai: replies * COST.aiReply,
    xfer: xferGuess * COST.outbound, vm: msgs.length * 1 * COST.voicemail, sms: segs * COST.sms, fixed: COST.fixedDay,
  };
  const total = Object.values(cost).reduce((a, b) => a + b, 0);
  const cut = calls.length ? Math.round((bookingsN / calls.length) * 1000) / 10 : 0;

  const md = [
    "## Today at a glance",
    `- **${calls.length} calls** to the phone AI${vs(calls.length, pAvg(() => true))}${afterHrs ? ` · ${afterHrs} outside opening hours` : ""}${busiest ? ` · busiest hour ${hourName(+busiest[0])} (${busiest[1]} calls)` : ""}`,
    `- **${bookingsN} booking${bookingsN === 1 ? "" : "s"}** made by the AI${vs(bookingsN, pAvg((c) => c.outcome === "booking"))} · cut-through to booking **${cut}%**`,
    `- **${quotesN} estimate request${quotesN === 1 ? "" : "s"}**`,
    `- **${answered + wlThrough} put through** to the team${wlThrough ? ` (incl. ${wlThrough} whitelisted supplier${wlThrough === 1 ? "" : "s"})` : ""}${missed ? ` · **${missed} not answered**` : ""}`,
    `- **${vmN} voicemail${vmN === 1 ? "" : "s"}** · ${hung} hung up or cut off`,
    `- Average time with the AI: ${Math.floor(avgSecs / 60)}m ${pad(Math.round(avgSecs % 60))}s`,
    "",
    "## Bookings made by the AI",
    ...((bks || []).length ? bks.map((b) => `- **#${b.booking_ref || "–"} ${b.customer_name || ""}** ${b.phone || ""} · ${[b.rego, b.vehicle || [b.make, b.model].filter(Boolean).join(" ")].filter(Boolean).join(" ") || "no rego yet"} · ${b.service || ""} · **${longDate(b.booking_date)} ${time12(b.time_in || "")}**${b.state === "inbox" ? " · *waiting to be accepted*" : ""}`) : ["- None today"]),
    "",
    "## Estimate requests",
    ...((qts || []).length ? qts.map((q) => `- **${q.name || ""}** ${q.phone || ""} · ${[q.rego, q.vehicle].filter(Boolean).join(" ")} – ${String(q.notes || "").replace(/^PHONE AI (QUOTE|ESTIMATE) REQUEST:\s*/, "").slice(0, 160)}`) : ["- None today"]),
    "",
    "## Why callers asked for the team",
    ...(throughSaid.length ? throughSaid.map((x) => "- " + x) : ["- No calls put through"]),
    "",
    "## Call backs needed",
    ...(cb.length ? cb.map((x) => "- " + x) : ["- Nobody waiting on a call back 👍"]),
    "",
    "## Texts & emails",
    `- **${smsOut.length} text${smsOut.length === 1 ? "" : "s"} sent** (${segs} SMS segment${segs === 1 ? "" : "s"})${Object.keys(byTpl).length ? " – " + Object.entries(byTpl).map(([k, v]) => `${k} ${v}`).join(", ") : ""}`,
    `- **${smsIn.length} repl${smsIn.length === 1 ? "y" : "ies"}** received${smsIn.length ? ": " + smsIn.slice(0, 5).map((x) => `${x.phone} "${String(x.message || "").slice(0, 60)}"`).join("; ") : ""}`,
    `- ${emails} emails sent by the system`,
    "",
    "## Running cost today (estimate)",
    `- Phone AI voice (${relayMin} min): **${aud(cost.relay + cost.inbound - xferGuess * COST.inbound)}**`,
    `- AI thinking (${replies} replies): **${aud(cost.ai)}**`,
    `- Putting calls through (${xferGuess} min${xferKnown ? "" : ", estimated"}): **${aud(cost.xfer + xferGuess * COST.inbound)}**`,
    `- Voicemail: **${aud(cost.vm)}** · Texts: **${aud(cost.sms)}** · Numbers + server: **${aud(cost.fixed)}**`,
    `- **Total ≈ ${aud(total)}** · about **${aud((cost.relay + cost.ai) / Math.max(1, aiCalls.length))} per AI call**`,
  ].join("\n");
  const subject = `Daily phone report – ${longDate(day)}: ${calls.length} calls, ${bookingsN} booking${bookingsN === 1 ? "" : "s"}${cb.length ? `, ${cb.length} call back${cb.length === 1 ? "" : "s"}` : ""}`;
  return { subject, html: reportEmail("Daily phone report", longDate(day), md), md };
}

async function sendPhoneReport(day, to) {
  const s = await getSettings(true);
  const rcpt = to || s.phone_report_to || s.report_email_to || s.notify_email;
  if (!rcpt) return { ok: false, error: "No email address set for the phone report" };
  const r = await buildPhoneReport(day);
  const sent = await sendEmail({ to: rcpt, subject: r.subject, html: r.html });
  return { ok: sent.ok, to: rcpt, subject: r.subject };
}

// once a day after the set time (default 6pm). Stored in settings so restarts don't double-send.
async function phoneReportJob() {
  const s = await getSettings(true);
  if (s.phone_report_enabled === false) return;
  const day = todayISO(), now = nowHHMM();
  if (now < (s.phone_report_time || "18:00") || now > "23:30" || s.phone_report_last === day) return;
  const { data: claimed } = await db.from("settings").update({ phone_report_last: day }).eq("id", s.id).or(`phone_report_last.is.null,phone_report_last.neq.${day}`).select("id");
  if (!claimed?.length) return;
  const r = await sendPhoneReport(day);
  console.log("Daily phone report", day, r.ok ? "sent to " + r.to : r.error || "failed");
}

// send / preview on demand from the app (Settings → "Send today's phone report now")
app.post("/api/reports/phone-daily", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const day = /^\d{4}-\d{2}-\d{2}$/.test(req.body?.day || "") ? req.body.day : todayISO();
  if (req.body?.preview) { const r = await buildPhoneReport(day); return reply.send({ ok: true, subject: r.subject, html: r.html }); }
  reply.send(await sendPhoneReport(day, req.body?.to));
});

// ---------------------------------------------------------------- weekly end-of-day report (Monday morning, last Mon–Sun)
const dow = (iso) => new Date(iso + "T00:00:00Z").getUTCDay();
const money = (n) => n == null ? "–" : "$" + Math.round(Number(n)).toLocaleString("en-AU");
async function buildWeeklyEod(weekEnd) {
  const start = addDays(weekEnd, -6), prevStart = addDays(start, -7);
  const [{ data: days }, { data: prev }, { data: invs }] = await Promise.all([
    db.from("eod_days").select("*").gte("register_date", start).lte("register_date", weekEnd).order("register_date"),
    db.from("eod_days").select("invoices,sales_inc,revenue,gross_profit").gte("register_date", prevStart).lt("register_date", start),
    db.from("eod_invoices").select("invoice_no,register_date,customer,salesperson,revenue,gross_profit,gpm").gte("register_date", start).lte("register_date", weekEnd),
  ]);
  const sum = (rows, k) => (rows || []).reduce((a, x) => a + (Number(x[k]) || 0), 0);
  const tot = { invoices: sum(days, "invoices"), sales: sum(days, "sales_inc"), revenue: sum(days, "revenue"), gp: sum(days, "gross_profit") };
  const ptot = { invoices: sum(prev, "invoices"), revenue: sum(prev, "revenue"), gp: sum(prev, "gross_profit") };
  const gpm = tot.revenue ? Math.round((tot.gp / tot.revenue) * 1000) / 10 : null;
  const have = new Set((days || []).map((x) => x.register_date));
  const missing = [];
  for (let i = 0; i < 6; i++) { const d = addDays(start, i); if (!have.has(d)) missing.push(longDate(d).replace(/\s+\d{4}$/, "")); }
  const pct = (a, b) => (b ? `${a >= b ? "▲" : "▼"} ${Math.abs(Math.round(((a - b) / b) * 100))}% on last week` : "");
  const rows = (days || []).map((x) => `| ${longDate(x.register_date).replace(/\s+\d{4}$/, "")} | ${x.invoices ?? "–"} | ${money(x.revenue ?? x.sales_inc)} | ${money(x.gross_profit)} | ${x.gpm ?? "–"}% | ${money(x.banking?.total)} |`);
  let notes = [];
  if ((days || []).length) {
    try {
      const sp = {};
      (invs || []).forEach((v) => { const k = v.salesperson || "?"; sp[k] = sp[k] || { invoices: 0, revenue: 0, gp: 0 }; sp[k].invoices++; sp[k].revenue += +v.revenue || 0; sp[k].gp += +v.gross_profit || 0; });
      const low = (invs || []).filter((v) => v.gpm != null && +v.revenue > 50).sort((a, b) => a.gpm - b.gpm).slice(0, 5);
      const res = await claude.messages.create({ model: REPORT_MODEL, max_tokens: 800, messages: [{ role: "user", content:
        `Weekly figures for Midas Ipswich (automotive workshop), week ${start} to ${weekEnd}.\nDays: ${JSON.stringify(days.map((x) => ({ d: x.register_date, inv: x.invoices, rev: x.revenue, gp: x.gross_profit, gpm: x.gpm })))}\nThis week totals: ${JSON.stringify(tot)}; last week: ${JSON.stringify(ptot)}\nBy salesperson: ${JSON.stringify(sp)}\nLowest-margin invoices: ${JSON.stringify(low)}\nMissing daily reports: ${missing.join(", ") || "none"}\n\nWrite 3-5 short, plain bullet points an owner would care about (trend vs last week, average invoice, margin, standout or low-margin jobs, salesperson split). Return ONLY the bullets, each starting with "- ".` }] });
      notes = res.content.filter((c) => c.type === "text").map((c) => c.text).join("").split("\n").map((l) => l.trim()).filter((l) => l.startsWith("- "));
    } catch (e) { console.warn("weekly eod notes", e.message); }
  }
  const md = [
    `## Week ${longDate(start).replace(/\s+\d{4}$/, "")} – ${longDate(weekEnd)}`,
    days?.length ? `- **Revenue:** ${money(tot.revenue)} ex GST ${pct(tot.revenue, ptot.revenue)}` : "- No end-of-day reports were received for this week.",
    days?.length ? `- **Invoices:** ${tot.invoices} · average ${tot.invoices ? money(tot.revenue / tot.invoices) : "–"} ${pct(tot.invoices, ptot.invoices)}` : null,
    days?.length ? `- **Gross profit:** ${money(tot.gp)} (${gpm ?? "–"}%)` : null,
    missing.length ? `- **No report received for:** ${missing.join(", ")} – email them in if the shop was open.` : null,
    rows.length ? "\n### By day\n| Day | Invoices | Revenue | GP | GP% | Banked |\n|---|---|---|---|---|---|\n" + rows.join("\n") : null,
    notes.length ? "\n### Worth knowing\n" + notes.join("\n") : null,
  ].filter(Boolean).join("\n");
  return { md, start, weekEnd, subject: `Weekly end of day – week ending ${longDate(weekEnd)}`, html: reportEmail("Weekly end of day", `Week ending ${longDate(weekEnd)}`, md) };
}
async function sendWeeklyEod(weekEnd, to) {
  const s = await getSettings();
  const r = await buildWeeklyEod(weekEnd);
  const dest = to || s.report_email_to || s.notify_email;
  if (!dest) return { ok: false, error: "No report email set" };
  await sendEmail({ to: dest, subject: r.subject, html: r.html });
  await db.from("ho_reports").insert({ uploaded_by: "Weekly report", file_name: null, period: "Weekly EOD " + weekEnd, extracted: { report_type: "weekly_eod", week_end: weekEnd }, analysis: r.md, model: REPORT_MODEL, source: "weekly" });
  return { ok: true, to: dest };
}
let weeklyEodDone = null;
async function weeklyEodJob() {
  const today = todayISO();
  if (dow(today) !== 1 || nowHHMM() < "07:30" || weeklyEodDone === today) return;
  weeklyEodDone = today;
  const weekEnd = addDays(today, -1);
  const { data: ex } = await db.from("ho_reports").select("id").eq("period", "Weekly EOD " + weekEnd).limit(1);
  if (ex?.length) return;
  const r = await sendWeeklyEod(weekEnd);
  console.log("Weekly EOD report", weekEnd, r.ok ? "sent to " + r.to : r.error);
}
// on demand: { week_end: "YYYY-MM-DD" (a Sunday, default last Sunday), preview, to }
app.post("/api/reports/eod-weekly", async (req, reply) => {
  const who = await requireAdmin(req, reply); if (!who) return;
  const t = todayISO();
  const we = /^\d{4}-\d{2}-\d{2}$/.test(req.body?.week_end || "") ? req.body.week_end : addDays(t, -(dow(t) || 7));
  if (req.body?.preview) { const r = await buildWeeklyEod(we); return reply.send({ ok: true, subject: r.subject, html: r.html }); }
  reply.send(await sendWeeklyEod(we, req.body?.to));
});

// ---------------------------------------------------------------- background jobs
let lastReminderDay = null;
async function reminderJob() {
  const s = await getSettings();
  const today = todayISO();
  const now = nowHHMM();
  if (lastReminderDay === today || now < (s.sms_reminder_time || "18:00") || now > "20:00") return;
  lastReminderDay = today;
  const t = await getTemplate("REMIND");
  if (!t?.enabled) return;
  const { data } = await db.from("bookings").select("*")
    .eq("shop_id", SHOP_ID).eq("booking_date", addDays(today, 1)).eq("state", "booking").is("reminder_sent_at", null);
  for (const b of data || []) {
    if (["cancelled", "no_show"].includes(b.status)) continue;
    await sendBookingTemplate(b, "REMIND", "Auto reminder");
  }
  console.log(`Reminders sent for ${addDays(today, 1)}: ${(data || []).length}`);
}

let lastHolidaySync = 0;
async function holidayJob() {
  if (Date.now() - lastHolidaySync < 24 * 3600e3) return;
  lastHolidaySync = Date.now();
  const s = await getSettings();
  const st = String(s.holiday_state || "QLD").toUpperCase();
  const y = Number(todayISO().slice(0, 4));
  for (const year of [y, y + 1]) {
    try {
      const r = await fetch(`https://date.nager.at/api/v3/PublicHolidays/${year}/AU`);
      if (!r.ok) { console.error("holidays fetch", year, r.status); continue; }
      const list = (await r.json()).filter((h) => !h.counties || h.counties.includes(`AU-${st}`));
      console.log(`Holidays ${year} ${st}: ${list.length}`);
      for (const h of list) {
        const { data: ex } = await db.from("calendar_blocks").select("id")
          .eq("shop_id", SHOP_ID).eq("block_date", h.date).eq("source", "holiday").maybeSingle();
        if (!ex) {
          const { error } = await db.from("calendar_blocks").insert({
            shop_id: SHOP_ID, block_date: h.date, source: "holiday",
            reason: `${h.localName || h.name} (${st})`, active: true, created_by: "auto",
          });
          if (error && !/duplicate/i.test(error.message)) console.error("holiday insert", h.date, error.message);
        }
      }
    } catch (e) { console.error("holidays", e.message); }
  }
}

setInterval(() => {
  reminderJob().catch((e) => console.error("reminders", e));
  phoneReportJob().catch((e) => console.error("phone report", e));
  holidayJob().catch((e) => console.error("holidays", e));
  weeklyEodJob().catch((e) => console.error("weekly eod", e));
}, 60e3);
setTimeout(() => holidayJob().catch((e) => console.error("holidays", e)), 5e3);

app.listen({ port: PORT, host: "0.0.0.0" });

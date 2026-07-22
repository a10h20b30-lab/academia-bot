import "dotenv/config";
import TelegramBot from "node-telegram-bot-api";
import Anthropic from "@anthropic-ai/sdk";
import path from "path";
import { fileURLToPath } from "url";
import * as XLSX from "xlsx";
import { query, initDB } from "./db.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const ADMIN_ID = 6021652936;
const ADMIN_PHONE_LINK = "https://wa.me/972548028082?text=%D7%94%D7%99%D7%99%20%D7%90%D7%A0%D7%99%20%D7%9E%D7%A2%D7%95%D7%A0%D7%99%D7%99%D7%9F%20%D7%9C%D7%A7%D7%91%D7%9C%20%D7%A7%D7%95%D7%93%20%D7%9Ckozo"; // לינק לוואטסאפ עם הודעה מוכנה
let EMPLOYER_ACCESS_CODE = "KOZO8"; // קוד האישור הקבוע ללשכות/עיריות

// ── בדיקת API key ────────────────────────────────────────────────────────────

async function checkAnthropicKey() {
  try {
    await anthropic.messages.create({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 10,
      messages: [{ role: "user", content: "ping" }],
    });
  } catch (err) {
    try {
      await bot.sendMessage(ADMIN_ID, "⚠️ מפתח Anthropic לא תקין — שיחה חופשית לא תעבוד");
    } catch (_) {}
  }
}

function scheduleApiHealthCheck() {
  setTimeout(async () => {
    await checkAnthropicKey();
    scheduleApiHealthCheck();
  }, 60 * 60 * 1000);
}

// ── ארכיון ───────────────────────────────────────────────────────────────────

async function archiveCandidate(telegramId) {
  await query(
    `UPDATE candidates SET status='archived' WHERE telegram_id=$1`,
    [telegramId]
  );
}

// ── פרופיל מועמד ─────────────────────────────────────────────────────────────

function buildProfileMessage(cand) {
  const lines = [
    "הפרופיל שלך 📋",
    "",
    `שם: ${cand.full_name || ""}`,
    `מייל: ${cand.email || ""}`,
    `עיר: ${cand.city || ""}`,
    `תואר: ${cand.degree || ""} — ${cand.field_of_study || ""}`,
    `שפות: ${cand.languages || ""}`,
    ...(cand.is_intern === "כן ✅" && cand.internship_mentor ? [`🏛 התמחות: דוברות הכנסת — שנה | אצל: ${cand.internship_mentor}`] : []),
    `ניסיון: ${cand.experience || ""}`,
    `תחומים: ${cand.interests || ""}`,
    `מקום מועדף: ${cand.workplace_pref || ""}`,
    `היקף: ${cand.availability || ""}`,
    `צד פוליטי: ${cand.political_side || ""}`,
    `סטטוס: ${cand.status === "active" ? "פעיל ✅" : cand.status}`,
    "",
    "לעדכון פרטים — כתוב עדכן פרטים",
  ];
  return lines.join("\n");
}

// ── המלצות ───────────────────────────────────────────────────────────────────

async function getRecommendation(candidateId) {
  const res = await query(
    `SELECT * FROM recommendations WHERE candidate_id=$1`,
    [candidateId]
  );
  return res.rows[0] || null;
}

async function saveRecommendationText(candidateId, text, recommenderName) {
  await query(
    `INSERT INTO recommendations (candidate_id, text, recommender_name)
     VALUES ($1, $2, $3)
     ON CONFLICT (candidate_id) DO UPDATE SET text=$2, recommender_name=$3`,
    [candidateId, text, recommenderName]
  );
}

// ── השהייה ────────────────────────────────────────────────────────────────────

async function isPaused(telegramId) {
  const res = await query(
    `SELECT 1 FROM candidates WHERE telegram_id=$1 AND status IN ('paused', 'archived')`,
    [telegramId]
  );
  return res.rows.length > 0;
}

async function pauseCandidate(telegramId) {
  await query(
    `UPDATE candidates SET status='paused' WHERE telegram_id=$1 AND status='active'`,
    [telegramId]
  );
}

async function resumeCandidate(telegramId) {
  await query(
    `UPDATE candidates SET status='active' WHERE telegram_id=$1 AND status IN ('paused', 'archived')`,
    [telegramId]
  );
}

async function isEmployerPaused(telegramId) {
  const res = await query(
    `SELECT 1 FROM employers WHERE telegram_id=$1 AND status='paused'`,
    [telegramId]
  );
  return res.rows.length > 0;
}

async function pauseEmployer(telegramId) {
  await query(
    `UPDATE employers SET status='paused' WHERE telegram_id=$1`,
    [telegramId]
  );
}

async function resumeEmployer(telegramId) {
  await query(
    `UPDATE employers SET status='active' WHERE telegram_id=$1`,
    [telegramId]
  );
}

async function getCandidateRecord(telegramId) {
  const res = await query(
    `SELECT * FROM candidates WHERE telegram_id=$1 ORDER BY created_at DESC LIMIT 1`,
    [telegramId]
  );
  return res.rows[0] || null;
}

async function getEmployerRecord(telegramId) {
  const res = await query(
    `SELECT * FROM employers WHERE telegram_id=$1 ORDER BY created_at DESC LIMIT 1`,
    [telegramId]
  );
  return res.rows[0] || null;
}

async function updateCandidateRecord(telegramId, updates) {
  const ALLOWED_COLUMNS = [
    "full_name", "phone", "email", "city", "degree", "field_of_study",
    "languages", "is_intern", "internship_mentor",
    "experience", "interests", "workplace_pref", "timing", "availability",
    "cv", "motivation", "has_references", "references", "declaration", "status",
    "availability_status"
  ];
  const keys = Object.keys(updates).filter((k) => ALLOWED_COLUMNS.includes(k));
  if (keys.length === 0) return;

  const setClauses = keys.map((k, i) => `${k}=$${i + 2}`).join(", ");
  const values = keys.map((k) => updates[k]);
  await query(
    `UPDATE candidates SET ${setClauses}, updated_at=NOW()
     WHERE telegram_id=$1
     AND id=(SELECT id FROM candidates WHERE telegram_id=$1 ORDER BY created_at DESC LIMIT 1)`,
    [telegramId, ...values]
  );
}

// ── היסטוריית חיבורים ────────────────────────────────────────────────────────

async function hasBeenMatched(candidateId, employerId) {
  const res = await query(
    `SELECT 1 FROM matches WHERE candidate_id=$1 AND employer_id=$2`,
    [candidateId, employerId]
  );
  return res.rows.length > 0;
}

async function recordMatch(candidateId, employerId, candidateName, employerName) {
  await query(
    `INSERT INTO matches (candidate_id, employer_id, candidate_name, employer_name)
     VALUES ($1, $2, $3, $4)
     ON CONFLICT (candidate_id, employer_id) DO NOTHING`,
    [candidateId, employerId, candidateName, employerName]
  );
}

// ── טלפון ────────────────────────────────────────────────────────────────────

function normalizePhone(phone) { return phone.replace(/\D/g, ""); }
function isValidPhone(phone)   { return normalizePhone(phone).length >= 9; }
function isValidEmail(email)   { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email); }

async function isApproved(phone) {
  const res = await query(
    `SELECT 1 FROM approved_phones WHERE phone=$1`,
    [normalizePhone(phone)]
  );
  return res.rows.length > 0;
}

async function approvePhone(phone) {
  await query(
    `INSERT INTO approved_phones (phone) VALUES ($1) ON CONFLICT DO NOTHING`,
    [normalizePhone(phone)]
  );
}

// ── שמירת רשומות ──────────────────────────────────────────────────────────────

async function saveRecord(type, chatId, username, data) {
  if (type === "candidate") {
    await query(
      `INSERT INTO candidates (
        telegram_id, telegram_username,
        full_name, phone, email, city, degree, field_of_study,
        languages, is_intern, internship_mentor,
        experience, interests, workplace_pref, timing, availability,
        cv, motivation, has_references, "references",
        political_side, has_license, declaration,
        availability_status
      ) VALUES (
        $1, $2,
        $3, $4, $5, $6, $7, $8,
        $9, $10, $11,
        $12, $13, $14, $15, $16,
        $17, $18, $19, $20,
        $21, $22, $23,
        $24
      )`,
      [
        chatId, username || "",
        data.full_name || "", data.phone || "", data.email || "",
        data.city || "", data.degree || "", data.field_of_study || "",
        data.languages || "", data.is_intern || "", data.internship_mentor || "",
        data.experience || "", data.interests || "",
        data.workplace_pref || "", data.timing || "", data.availability || "",
        data.cv || "", data.motivation || "", data.has_references || "",
        data.references || "",
        data.political_side || "", data.has_license || "", data.declaration || "",
        data.availability_status || "",
      ]
    );
  } else {
    await query(
      `INSERT INTO employers (
        telegram_id, telegram_username,
        org_type, contact_name, phone, email,
        fields, timing, availability, experience_importance,
        notes, political_side, requires_license, english_required, irregular_hours, future_search, declaration
      ) VALUES (
        $1, $2,
        $3, $4, $5, $6,
        $7, $8, $9, $10,
        $11, $12, $13, $14, $15, $16, $17
      )`,
      [
        chatId, username || "",
        data.org_type || "", data.contact_name || "", data.phone || "",
        data.email || "", data.fields || "", data.timing || "",
        data.availability || "", data.experience_importance || "",
        data.notes || "", data.political_side || "", data.requires_license || "",
        data.english_required || "", data.irregular_hours || "", data.future_search || "", data.declaration || "",
      ]
    );
  }
  console.log(`נשמר: ${type} | ${username || chatId}`);
}

async function scheduleFollowUp(candidateId, employerId) {
  const scheduledFor = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['follow_up', candidateId, JSON.stringify({ employerId }), scheduledFor]
  );
}

async function scheduleCVFollowUp(candidateId, employerId, candidateName) {
  const scheduledFor = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['cv_followup', candidateId, JSON.stringify({ employerId, candidateName }), scheduledFor]
  );
}

async function scheduleRatingRequest(employerId, candidateId, candidateName) {
  const scheduledFor = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['rating_request', employerId, JSON.stringify({ candidateId, candidateName }), scheduledFor]
  );
}

async function scheduleCandidateNudge(candidateId) {
  const scheduledFor = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['candidate_nudge', candidateId, null, scheduledFor]
  );
}

async function scheduleEmployerListFollowUp(employerId) {
  const scheduledFor = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['employer_list_followup', employerId, null, scheduledFor]
  );
}

async function scheduleMonthlyCheckin(candidateId) {
  const scheduledFor = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['monthly_checkin', candidateId, null, scheduledFor]
  );
}

async function scheduleBimonthlyReminder(candidateId) {
  const scheduledFor = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);
  await query(
    `INSERT INTO scheduled_tasks (task_type, target_id, extra_data, scheduled_for) VALUES ($1, $2, $3, $4)`,
    ['bimonthly_reminder', candidateId, null, scheduledFor]
  );
}

async function runTask(task) {
  const extra = task.extra_data ? JSON.parse(task.extra_data) : {};
  switch (task.task_type) {
    case 'follow_up': {
      const { employerId } = extra;
      const res = await query(
        `SELECT * FROM matches WHERE candidate_id=$1 AND employer_id=$2 AND status='active'`,
        [task.target_id, employerId]
      );
      const match = res.rows[0];
      if (!match) break;
      await bot.sendMessage(
        ADMIN_ID,
        `📊 מעקב התאמה, שבוע עבר\n\n👤 מועמד: ${match.candidate_name}\n🏛 לשכה: ${match.employer_name}\n\nהאם ההתאמה עדיין בתהליך?`,
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "כן, בתהליך ✅", callback_data: `FOLLOWUP_YES_${task.target_id}_${employerId}` },
              { text: "לא, נגמר ❌",   callback_data: `FOLLOWUP_NO_${task.target_id}_${employerId}`  },
            ]],
          },
        }
      );
      break;
    }
    case 'cv_followup': {
      const { employerId, candidateName } = extra;
      const res = await query(
        `SELECT status FROM cv_requests WHERE employer_id=$1 AND candidate_id=$2`,
        [employerId, task.target_id]
      );
      const req = res.rows[0];
      if (!req || req.status === "connected" || req.status === "rejected") break;
      await bot.sendMessage(
        employerId,
        `מה קרה עם ${candidateName}?`,
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "✅ יצרתי קשר",  callback_data: `CVFOLLOWUP_CONTACTED_${task.target_id}_${employerId}` },
              { text: "⏳ עוד בתהליך", callback_data: `CVFOLLOWUP_INPROGRESS_${task.target_id}_${employerId}` },
              { text: "❌ לא מתאים",   callback_data: `CVFOLLOWUP_NOTSUITABLE_${task.target_id}_${employerId}` },
            ]],
          },
        }
      );
      break;
    }
    case 'rating_request': {
      const { candidateId, candidateName } = extra;
      await bot.sendMessage(
        task.target_id,
        `איך יצא החיבור עם ${candidateName}? ⭐`,
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "⭐⭐⭐⭐⭐", callback_data: `RATING_5_${task.target_id}_${candidateId}` },
              { text: "⭐⭐⭐",    callback_data: `RATING_3_${task.target_id}_${candidateId}` },
              { text: "⭐",       callback_data: `RATING_1_${task.target_id}_${candidateId}` },
            ]],
          },
        }
      );
      break;
    }
    case 'candidate_nudge': {
      const cand = await getCandidateRecord(task.target_id);
      if (!cand || cand.status !== "active") break;
      const res = await query(`SELECT 1 FROM cv_requests WHERE candidate_id=$1 LIMIT 1`, [task.target_id]);
      if (res.rows.length > 0) break;
      await bot.sendMessage(task.target_id, "הפרופיל שלך פעיל אצלנו, עדיין מחפשים עבורך. ברגע שיהיה התאמה — תשמע ממני.");
      break;
    }
    case 'employer_list_followup': {
      const emp = await getEmployerRecord(task.target_id);
      if (!emp || emp.status !== "active") break;
      const res = await query(`SELECT 1 FROM cv_requests WHERE employer_id=$1 LIMIT 1`, [task.target_id]);
      if (res.rows.length > 0) break;
      await bot.sendMessage(
        task.target_id,
        "עדיין מחפשים? הרשימה עדיין פעילה.",
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "כן, עוד מחפשים", callback_data: `EMPLIST_STILL_${task.target_id}` },
              { text: "לא, מצאנו",       callback_data: `EMPLIST_FOUND_${task.target_id}` },
            ]],
          },
        }
      );
      break;
    }
    case 'monthly_checkin': {
      const cand = await getCandidateRecord(task.target_id);
      if (!cand || cand.status !== "active") break;
      await bot.sendMessage(
        task.target_id,
        "חודש עבר מאז שנרשמת. מה הסטטוס?",
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "עדיין מחפש",   callback_data: `CHECKIN_STILL_${task.target_id}` },
              { text: "מצאתי עבודה", callback_data: `CHECKIN_FOUND_${task.target_id}` },
            ]],
          },
        }
      );
      break;
    }
    case 'bimonthly_reminder': {
      const cand = await getCandidateRecord(task.target_id);
      if (!cand || cand.status !== "active") break;
      await bot.sendMessage(
        task.target_id,
        "היי, עדיין מחפש הזדמנות? 👋\nהפרופיל שלך פעיל — רק רוצים לוודא שהמידע עדכני.",
        {
          reply_markup: {
            inline_keyboard: [
              [{ text: "כן, עדיין מחפש",  callback_data: `CHECKIN_STILL_${task.target_id}` }],
              [{ text: "עדכן פרטים",       callback_data: "EXISTING_UPDATE" }],
              [{ text: "מצאתי עבודה 🎉",  callback_data: `CHECKIN_FOUND_${task.target_id}` }],
            ],
          },
        }
      );
      break;
    }
  }
}

async function checkScheduledTasks() {
  try {
    const res = await query(
      `SELECT * FROM scheduled_tasks WHERE done=false AND scheduled_for <= NOW()`
    );
    for (const task of res.rows) {
      try {
        await runTask(task);
      } catch (err) {
        console.error(`runTask error (id=${task.id} type=${task.task_type}):`, err.message);
      }
      await query(`UPDATE scheduled_tasks SET done=true WHERE id=$1`, [task.id]);
    }
  } catch (err) {
    console.error("checkScheduledTasks error:", err.message);
  }
}


// ── חיבורים ──────────────────────────────────────────────────────────────────

function workplaceMatches(candidatePref, orgType) {
  if (!candidatePref || candidatePref === "שניהם") return true;
  return candidatePref === orgType;
}

// ── שליחת סיכום התאמות למגייס ──────────────────────────────────────────────────

async function sendMatchSummary(candidates, employer, notifyCandidates = true) {
  const employerId = employer.telegram_id;

  for (const candidate of candidates) {
    await recordMatch(candidate.telegram_id, employerId, candidate.full_name || "מועמד", employer.contact_name || "לשכה");
    if (notifyCandidates) {
      await bot.sendMessage(
        candidate.telegram_id,
        `העברתי את הפרטים שלך ללשכה/גוף חדש שנרשם ונראה לי מתאים. אם יתאים — ייצרו איתך קשר 🤝`,
        { reply_markup: { inline_keyboard: [[{ text: "הפסק לקבל הצעות 🔕", callback_data: `STOP_OFFERS_CANDIDATE_${candidate.telegram_id}` }]] } }
      );
    }
  }

  const keyboard = candidates.map((c) => {
    const score = calcMatchScore(c, employer);
    return [{
      text: `⭐ ${score}% התאמה — ${c.full_name || "מועמד"}`,
      callback_data: `CV_${c.telegram_id}_${employerId}`,
    }];
  });
  keyboard.push([{ text: "🔄 הצג עוד מועמדים", callback_data: `REFRESH_MATCHES_${employerId}` }]);
  keyboard.push([{ text: "הפסק לקבל הצעות 🔕", callback_data: `STOP_OFFERS_EMPLOYER_${employerId}` }]);

  await bot.sendMessage(
    employerId,
    `מצאתי ${candidates.length} מועמדים שנראים לי מתאימים 👋\n\nברגע שיירשמו עוד מתאימים — תקבלו עדכון נוסף 🤝`,
    { reply_markup: { inline_keyboard: keyboard } }
  );

  await bot.sendMessage(
    ADMIN_ID,
    `🔗 התאמות אוטומטיות\n\n` +
    `🏛 ${employer.org_type || "לשכה"}: ${employer.contact_name || ""}\n` +
    `👤 מועמדים: ${candidates.map((c) => c.full_name || "מועמד").join(", ")}`
  );

  await scheduleEmployerListFollowUp(employerId);
}

function politicalMatches(candidateSide, employerSide) {
  // מסנן רק אם שניהם ציינו העדפה מפורשת
  if (!candidateSide || candidateSide === "שניהם")    return true;
  if (!employerSide  || employerSide  === "לא רלוונטי") return true;
  return candidateSide === employerSide;
}

function availabilityMatches(candidateAvail, employerAvail) {
  if (!candidateAvail || !employerAvail) return true;
  if (candidateAvail === "פתוח לכל הצעה" || employerAvail === "פתוח לכל הצעה") return true;
  return candidateAvail === employerAvail;
}

async function findMatches(employer) {
  const res = await query(`SELECT * FROM candidates WHERE status='active'`);
  const candidates = res.rows;
  const fields = (employer.fields || "").split(", ");

  const rejectedRes = await query(
    `SELECT candidate_id FROM cv_requests WHERE employer_id=$1 AND status='rejected'`,
    [employer.telegram_id]
  );
  const rejectedIds = new Set(rejectedRes.rows.map((r) => r.candidate_id));

  const filtered = [];
  for (const c of candidates) {
    if (rejectedIds.has(c.telegram_id)) continue;
    if (c.availability_status === "⚪ לא מחפש כרגע") continue;
    if (await hasBeenMatched(c.telegram_id, employer.telegram_id)) continue;
    if (!workplaceMatches(c.workplace_pref, employer.org_type)) continue;
    if (!politicalMatches(c.political_side, employer.political_side)) continue;
    if (!availabilityMatches(c.availability, employer.availability)) continue;
    const interests = (c.interests || "").split(", ");
    if (fields.some((f) => interests.some((i) => i.trim() === f.trim()))) {
      filtered.push(c);
    }
  }
  return filtered;
}

// מציאת לשכות מתאימות למועמד חדש
async function findMatchingEmployers(candidate) {
  const res = await query(`SELECT * FROM employers WHERE status='active'`);
  const employers = res.rows;
  const interests = (candidate.interests || "").split(", ").map((i) => i.trim());
  const filtered = [];
  for (const e of employers) {
    if (candidate.availability_status === "⚪ לא מחפש כרגע") continue;
    if (await hasBeenMatched(candidate.telegram_id, e.telegram_id)) continue;
    if (!workplaceMatches(candidate.workplace_pref, e.org_type)) continue;
    if (!politicalMatches(candidate.political_side, e.political_side)) continue;
    if (!availabilityMatches(candidate.availability, e.availability)) continue;
    const fields = (e.fields || "").split(", ").map((f) => f.trim());
    if (fields.some((f) => interests.includes(f))) {
      filtered.push(e);
    }
  }
  return filtered;
}

// לשכות שהסכימו לחיפושים עתידיים (status='paused' + future_search='כן')
async function findFutureSearchEmployers(candidate) {
  const res = await query(`SELECT * FROM employers WHERE status='paused' AND future_search='כן'`);
  const employers = res.rows;
  const interests = (candidate.interests || "").split(", ").map((i) => i.trim());
  const filtered = [];
  for (const e of employers) {
    if (await hasBeenMatched(candidate.telegram_id, e.telegram_id)) continue;
    if (!workplaceMatches(candidate.workplace_pref, e.org_type)) continue;
    if (!politicalMatches(candidate.political_side, e.political_side)) continue;
    if (!availabilityMatches(candidate.availability, e.availability)) continue;
    const fields = (e.fields || "").split(", ").map((f) => f.trim());
    if (fields.some((f) => interests.includes(f))) {
      filtered.push(e);
    }
  }
  return filtered;
}

// ── דירוג התאמה ──────────────────────────────────────────────────────────────

function calcMatchScore(candidate, employer) {
  let matched = 0;
  const interests = (candidate.interests || "").split(", ").map((s) => s.trim());
  const fields    = (employer.fields    || "").split(", ").map((s) => s.trim());
  if (fields.some((f) => interests.includes(f)))                              matched++;
  if (workplaceMatches(candidate.workplace_pref, employer.org_type))          matched++;
  if (politicalMatches(candidate.political_side, employer.political_side))    matched++;
  if (availabilityMatches(candidate.availability, employer.availability))     matched++;
  return Math.round((matched / 4) * 100);
}

// ── Excel ─────────────────────────────────────────────────────────────────────

async function exportExcel() {
  try {
    const [candidatesRes, employersRes, matchesRes, accessRes, connectedRes, ratingsRes] = await Promise.all([
      query(`SELECT * FROM candidates ORDER BY created_at ASC`),
      query(`SELECT * FROM employers ORDER BY created_at ASC`),
      query(`SELECT * FROM matches ORDER BY matched_at ASC`),
      query(`SELECT * FROM access_requests ORDER BY timestamp ASC`),
      query(`SELECT * FROM cv_requests WHERE status='connected' ORDER BY updated_at DESC`),
      query(`SELECT * FROM ratings`),
    ]);
    const candidates = candidatesRes.rows;
    const employers  = employersRes.rows;
    const matchesHistory = matchesRes.rows;
    const accessRequests = accessRes.rows;
    const connected  = connectedRes.rows;
    const ratings    = ratingsRes.rows;

    const CANDIDATE_HEADERS = ["תאריך","טלגרם","שם מלא","נייד","מייל","עיר","תואר","תחום לימודים","שפות","עבר התמחות","ניסיון","תחומי עניין","מקום עבודה מועדף","זמינות","קורות חיים","מוטיבציה","הצהרה","סטטוס","הוצע ל"];
    const EMPLOYER_HEADERS  = ["תאריך","טלגרם","מטעם","שם ותפקיד","נייד","מייל","תחומים","היקף","תזמון","חשיבות ניסיון","הערות","הצהרה","סטטוס"];

    const fmtC = (r) => ({
      "תאריך": r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
      "טלגרם": r.telegram_username ? `@${r.telegram_username}` : String(r.telegram_id || ""),
      "שם מלא": r.full_name || "", "נייד": r.phone || "", "מייל": r.email || "",
      "עיר": r.city || "", "תואר": r.degree || "", "תחום לימודים": r.field_of_study || "",
      "שפות": r.languages || "", "עבר התמחות": r.is_intern || "",
      "ניסיון": r.experience || "", "תחומי עניין": r.interests || "",
      "מקום עבודה מועדף": r.workplace_pref || "",
      "מועד פנוי": r.timing || "", "זמינות": r.availability || "", "קורות חיים": r.cv || "",
      "מוטיבציה": r.motivation || "", "הצהרה": r.declaration || "",
      "סטטוס": r.status === "paused" ? "מושהה" : r.status === "archived" ? "ארכיון" : "פעיל",
      "הוצע ל": matchesHistory
        .filter((m) => m.candidate_id === r.telegram_id)
        .map((m) => m.employer_name + (m.status === "closed" ? " ✗" : " ✓"))
        .join(", ") || "—",
    });

    const fmtE = (r) => ({
      "תאריך": r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
      "טלגרם": r.telegram_username ? `@${r.telegram_username}` : String(r.telegram_id || ""),
      "מטעם": r.org_type || "",
      "שם ותפקיד": r.contact_name || "", "נייד": r.phone || "", "מייל": r.email || "",
      "תחומים": r.fields || "", "מועד": r.timing || "", "היקף": r.availability || "",
      "חשיבות ניסיון": r.experience_importance || "", "הערות": r.notes || "",
      "הצהרה": r.declaration || "",
      "סטטוס": r.status === "paused" ? "מושהה" : "פעיל",
    });

    const makeSheet = (title, headers, rows) => {
      const ws = XLSX.utils.aoa_to_sheet([[title]]);
      XLSX.utils.sheet_add_aoa(ws, [[]], { origin: "A2" });
      if (rows.length > 0) {
        XLSX.utils.sheet_add_json(ws, rows, { origin: "A3", skipHeader: false });
      } else {
        XLSX.utils.sheet_add_aoa(ws, [headers, ["אין נתונים עדיין"]], { origin: "A3" });
      }
      ws["!cols"] = headers.map(() => ({ wch: 22 }));
      ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];
      return ws;
    };

    // גיליון ארכיון — מועמדים עם status='archived'
    const archived = candidates.filter((c) => c.status === "archived");
    const activeCandidates = candidates.filter((c) => c.status !== "archived");

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, makeSheet("רשימת יועצים ודוברים — קוזו", CANDIDATE_HEADERS, activeCandidates.map(fmtC)), "יועצים ודוברים");
    XLSX.utils.book_append_sheet(wb, makeSheet("רשימת לשכות חברי כנסת — קוזו", EMPLOYER_HEADERS, employers.map(fmtE)), "חברי כנסת");

    // גיליון ארכיון
    XLSX.utils.book_append_sheet(wb, makeSheet("ארכיון — מצאו עבודה", CANDIDATE_HEADERS, archived.map(fmtC)), "ארכיון");

    // גיליון בקשות גישה
    const ACCESS_HEADERS = ["תאריך", "שם", "נייד", "תחום", "מחפש", "שמע עלינו", "סטטוס"];
    const fmtA = (r) => ({
      "תאריך": r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
      "שם": r.full_name || "",
      "נייד": r.phone || "",
      "תחום": r.role || "",
      "מחפש": r.job_search || "",
      "שמע עלינו": r.heard_from || "",
      "סטטוס": r.status === "approved" ? "✅ אושר" : r.status === "denied" ? "❌ נדחה" : "⏳ ממתין",
    });
    XLSX.utils.book_append_sheet(wb, makeSheet("בקשות הצטרפות — קוזו", ACCESS_HEADERS, accessRequests.map(fmtA)), "בקשות הצטרפות");

    // גיליון חיבורים
    const MATCHES_HEADERS = ["תאריך", "שם מועמד", "שם לשכה/גוף", "תחומים", "סטטוס"];
    const fmtM = (m) => {
      const cand = candidates.find((c) => c.telegram_id === m.candidate_id);
      return {
        "תאריך": m.matched_at ? new Date(m.matched_at).toLocaleString("he-IL") : "",
        "שם מועמד": m.candidate_name || "",
        "שם לשכה/גוף": m.employer_name || "",
        "תחומים": cand?.interests || "",
        "סטטוס": m.status === "closed" ? "סגור" : "פעיל",
      };
    };
    XLSX.utils.book_append_sheet(wb, makeSheet("התאמות — קוזו", MATCHES_HEADERS, matchesHistory.map(fmtM)), "התאמות");

    // גיליון חיבורים מוצלחים
    const CONNECTED_HEADERS = ["תאריך חיבור", "שם יועץ", "שם מגייס", "גוף", "דירוג", "מקור"];
    const fmtConn = (r) => {
      const cand = candidates.find((c) => c.telegram_id === r.candidate_id);
      const emp  = employers.find((e) => e.telegram_id === r.employer_id);
      const rating = ratings.find((rt) => rt.employer_id === r.employer_id && rt.candidate_id === r.candidate_id);
      return {
        "תאריך חיבור": r.updated_at ? new Date(r.updated_at).toLocaleString("he-IL") : "",
        "שם יועץ":  cand?.full_name   || `ID:${r.candidate_id}`,
        "שם מגייס": emp?.contact_name || `ID:${r.employer_id}`,
        "גוף":      emp?.org_type     || "",
        "דירוג":    rating ? "⭐".repeat(rating.stars) : "",
        "מקור":     cand?.job_source === "kozo" ? "דרך קוזו 🤝" : cand?.job_source === "other" ? "ממקום אחר" : "",
      };
    };
    XLSX.utils.book_append_sheet(wb, makeSheet("חיבורים מוצלחים — קוזו", CONNECTED_HEADERS, connected.map(fmtConn)), "חיבורים מוצלחים");

    const outPath = path.join(__dirname, "../טבלה נתונים.xlsx");
    XLSX.writeFile(wb, outPath);
    console.log("📊 Excel עודכן:", outPath);
  } catch (e) {
    console.error("exportExcel error:", e.message);
  }
}

function formatRecord(type, session) {
  const d = session.data;
  if (type === "candidate") {
    return `שם: ${d.full_name}\nטלפון: ${d.phone}\nמייל: ${d.email}\nעיר: ${d.city}\nתואר: ${d.degree}\nלמד: ${d.field_of_study}\nשפות: ${d.languages}\nעבר התמחות: ${d.is_intern}\nניסיון: ${d.experience}\nתחומים: ${d.interests}\nמקום עבודה מועדף: ${d.workplace_pref}\nמועד: ${d.timing}\nזמינות: ${d.availability}\nמוטיבציה: ${d.motivation}`;
  } else {
    return `מטעם: ${d.org_type}\nשם: ${d.contact_name}\nטלפון: ${d.phone}\nמייל: ${d.email}\nתחום: ${d.fields}\nצד פוליטי: ${d.political_side || "לא צוין"}\nמועד: ${d.timing}\nהיקף: ${d.availability}\nניסיון: ${d.experience_importance}\nדגשים: ${d.notes}`;
  }
}

// ── Claude chat ───────────────────────────────────────────────────────────────

const CHAT_SYSTEM = `אתה קוזו – בוט שמכיר את כל אנשי המקצוע בתחום הפוליטי ומחבר ביניהם.

כללים:
- דבר תמיד בעברית בלבד.
- היה חם, מקצועי, תמציתי.
- ענה רק על שאלות שקשורות ל: רישום, חיבורים, השהייה/חזרה, מצב הפרופיל, מה זה קוזו.
- אם שואלים שאלות שלא קשורות – הסבר בנימוס שאתה כאן רק לצורך החיבורים המקצועיים.
- אל תמציא מידע על חברי כנסת, לשכות ספציפיות, או מועמדים.
- אם המשתמש רוצה להירשם, להשהות, להחזר אותי לפעילות, או לעדכן פרטים – תאמר לו שיכתוב את המילה המתאימה:
  • "השהה אותי" – לעצור זמנית את החיפוש
  • "החזר אותי לפעילות" – לחזור למאגר הפעיל
  • "עדכן פרטים" – לעדכן ניסיון/זמינות/תחומים
  • "/start" – להירשם מחדש`;

const chatHistories = {}; // chatId → [{role, content}]

async function claudeChat(chatId, userMessage) {
  if (!chatHistories[chatId]) chatHistories[chatId] = [];
  chatHistories[chatId].push({ role: "user", content: userMessage });

  // שמור היסטוריה עד 20 הודעות אחרונות
  if (chatHistories[chatId].length > 20) {
    chatHistories[chatId] = chatHistories[chatId].slice(-20);
  }

  try {
    const response = await anthropic.messages.create({
      model: "claude-sonnet-4-5",
      max_tokens: 500,
      system: CHAT_SYSTEM,
      messages: chatHistories[chatId],
    });
    const reply = response.content[0].text;
    chatHistories[chatId].push({ role: "assistant", content: reply });
    return reply;
  } catch (e) {
    console.error("Claude error:", e.message);
    return "מצטער, יש תקלה זמנית. נסה שוב עוד רגע 🙏\nלפניה ישירה: wa.me/972548028082";
  }
}

// ── שלבי טופס ────────────────────────────────────────────────────────────────

const sessions = {};

function newSession(type, username) {
  return { type, step: 0, data: {}, multiSelect: [], username: username || "", verified: true };
}

const CANDIDATE_STEPS = [
  { key: "full_name",         question: "נתחיל, מה השם?",                                                                    type: "text"   },
  { key: "email",             question: "כתובת מייל?",                                                                          type: "email"  },
  { key: "city",              question: "עיר מגורים?",                                                                          type: "text"   },
  { key: "degree",            question: "מה התואר?",                                                                            type: "single", options: [["תואר ראשון", "תואר שני"], ["אין תואר"]] },
  { key: "field_of_study",    question: "מה תחום הלימודים?",                                                                    type: "text"   },
  { key: "languages",         question: "באילו שפות יש שליטה?\nאפשר לסמן כמה ולחץ סיום ✓",                                  type: "multi",  options: [["עברית", "אנגלית"], ["ערבית", "רוסית"], ["אחר", "סיום ✓"]] },
  { key: "is_intern",         question: "האם עברת התמחות בכנסת?",                                                               type: "single", options: [["כן ✅", "לא ❌"]] },
  { key: "internship_mentor", question: "שם ונייד הדובר/ת שאצלו/ה התמחית",                                                       type: "text",   conditional: "is_intern=כן ✅" },
  { key: "experience",        question: "ספר/י על הניסיון המקצועי שלך — תפקידים, מקומות עבודה, מה עשית", type: "text"   },
  { key: "interests",         question: "באילו תחומים יש התמחות או עניין?\nאפשר לסמן כמה ולחץ סיום ✓",                      type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "workplace_pref",    question: "איפה מעדיפים לעבוד?",                                                                type: "single", options: [["כנסת", "עירייה"], ["שניהם"]] },
  { key: "timing",            question: "מתי פנוי להתחיל?",                                                                 type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
  { key: "availability",      question: "מה היקף המשרה המבוקש?",                                                                type: "single", options: [["משרה מלאה", "משרה חלקית"], ["פרילנס", "פתוח לכל הצעה"]] },
  { key: "cv",                question: "קורות חיים 📎\nגם לא מושלמים, ניצור קשר אם יידרשו פרטים נוספים.",                  type: "file"   },
  { key: "motivation",        question: "מה מביא אותך לקוזו?\nכמה מילים מהלב",                                   type: "text"   },
  { key: "has_references",    question: "האם יש לך ממליצים שלשכות יוכלו לפנות אליהם?",                                         type: "single", options: [["כן ✅", "לא ❌"]] },
  { key: "references",        question: "ציין שם ונייד של הממליצים (אפשר כמה, מופרדים בשורות)",                         type: "text",   conditional: "has_references=כן ✅" },
  { key: "political_side",    question: "עם איזה צד פוליטי תרצה לעבוד?",                                                  type: "single", options: [["קואליציה", "אופוזיציה"], ["שניהם"]] },
  { key: "has_license",       question: "יש רישיון רכב?",                                                                  type: "single", options: [["כן", "לא"]] },
  { key: "availability_status", question: "מה מצב החיפוש שלך?", type: "single", options: [["🟢 מחפש באופן פעיל", "🟡 פתוח להצעות"], ["⚪ לא מחפש כרגע"]] },
  { key: "declaration",       question: "רק לידיעה. הפרטים ישמשו אותי להתאמות בלבד. אין בזה התחייבות מאף צד 🤝", type: "single", options: [["מאשר ✅"]] },
];

const EMPLOYER_STEPS = [
  { key: "org_type",             question: "אתם מטעם...?",                                                                       type: "single", options: [["כנסת", "עירייה"], ["משרד יח\"צ / סושיאל", "אחר"]] },
  { key: "contact_name",         question: "שם ותפקיד?",                                                                type: "text"   },
  { key: "email",                question: "כתובת מייל?",                                                                        type: "email"  },
  { key: "fields",               question: "מה תחום החיזוק המבוקש?\nאפשר לסמן כמה ולחץ סיום ✓",                              type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "timing",               question: "מתי נדרש מישהו?",                                                                    type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
  { key: "availability",         question: "מה היקף המשרה המבוקשת?",                                                            type: "single", options: [["משרה מלאה", "משרה חלקית"], ["פרילנס", "פתוח לכל הצעה"]] },
  { key: "experience_importance",question: "כמה חשוב ניסיון קודם בעבודה ציבורית/פרלמנטרית?",                                     type: "single", options: [["חובה מוחלטת", "יתרון משמעותי"], ["לא הכרחי"]] },
  { key: "notes",                question: "יש דגשים נוספים שחשוב שנדע?\nאפשר לכתוב בחופשיות, גם 'אין' זה תשובה",          type: "text"   },
  { key: "political_side",       question: "אתם מהקואליציה או האופוזיציה?",                                                 type: "single", options: [["קואליציה", "אופוזיציה"], ["לא רלוונטי"]] },
  { key: "requires_license",     question: "נדרש רישיון רכב?",                                                              type: "single", options: [["חובה", "יתרון"], ["לא נדרש"]] },
  { key: "english_required",     question: "נדרשת אנגלית?",                                                                 type: "single", options: [["ברמה גבוהה", "בסיסית"], ["לא נדרש"]] },
  { key: "irregular_hours",      question: "נדרשת זמינות לשעות לא שגרתיות?",                                               type: "single", options: [["כן", "לא"]] },
  { key: "future_search",        question: "האם תרצו להישאר במאגר לחיפושים עתידיים?",                                   type: "single", options: [["כן", "לא"]] },
  { key: "declaration",          question: "רק לידיעה. הפרטים ישמשו להתאמה מקצועית בלבד. אין בזה התחייבות מאף צד 🤝", type: "single", options: [["מאשר ✅"]] },
];

// שאלות עדכון למועמד שחוזר מהשהייה
const UPDATE_STEPS = [
  { key: "experience",   question: "מה הניסיון המעודכן שלך מאז הרישום האחרון?",                                               type: "text"   },
  { key: "interests",    question: "באילו תחומים מעוניינים כיום?\nאפשר לסמן כמה ולחץ סיום ✓",                             type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "timing",       question: "מתי פנוי להתחיל?",                                                                      type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
  { key: "availability", question: "מה היקף המשרה המבוקש?",                                                                    type: "single", options: [["משרה מלאה", "משרה חלקית"], ["פרילנס", "פתוח לכל הצעה"]] },
];

function getSteps(type) {
  if (type === "candidate") return CANDIDATE_STEPS;
  if (type === "employer")  return EMPLOYER_STEPS;
  if (type === "update")    return UPDATE_STEPS;
  return [];
}

function buildKeyboard(options) {
  return {
    reply_markup: {
      inline_keyboard: options.map((row) =>
        row.map((label) => ({ text: label, callback_data: label.slice(0, 64) }))
      ),
    },
  };
}

async function sendStep(chatId, session) {
  const steps = getSteps(session.type);
  let step = steps[session.step];

  // דלג על שאלות conditional שלא מתקיים בהן התנאי
  while (step && step.conditional) {
    const [condKey, condVal] = step.conditional.split("=");
    if (session.data[condKey] !== condVal) {
      session.step++;
      step = steps[session.step];
    } else {
      break;
    }
  }

  if (!step) {
    await finishSession(chatId, session);
    return;
  }
  if (step.type === "single" || step.type === "multi") {
    await bot.sendMessage(chatId, step.question, buildKeyboard(step.options));
  } else {
    await bot.sendMessage(chatId, step.question);
  }
}

async function finishSession(chatId, session) {
  try {
  if (session.type === "update") {
    // עדכון פרטים ומחזיר לפעילות
    await updateCandidateRecord(chatId, session.data);
    await resumeCandidate(chatId);
    await exportExcel();
    await bot.sendMessage(ADMIN_ID, `🔄 מועמד חזר לפעילות (ID: ${chatId})\n${JSON.stringify(session.data, null, 2)}`);

    // חיפוש מחדש — מגייסים שמתאימים לפרופיל המעודכן
    const updatedCandidate = await getCandidateRecord(chatId);
    const matchingEmployers = updatedCandidate ? await findMatchingEmployers(updatedCandidate) : [];
    for (const employer of matchingEmployers) {
      await sendMatchSummary([updatedCandidate], employer, false);
    }
    if (matchingEmployers.length > 0) {
      await bot.sendMessage(chatId, `מעודכן! 🤝\nהעברתי את הפרופיל המעודכן שלך ל-${matchingEmployers.length} גופים מתאימים.`);
    } else {
      await bot.sendMessage(chatId, "מעודכן! 🤝\nחזרת לרשימה. ברגע שתהיה התאמה, אחבר.");
    }
    delete sessions[chatId];
    return;
  }

  await saveRecord(session.type, chatId, session.username, session.data);

  if (session.type === "candidate") {
    await bot.sendMessage(ADMIN_ID, `📥 מועמד חדש נרשם!\n\n${formatRecord("candidate", session)}`);
    if (session.data.cv) {
      if (session.data.cv.startsWith("file_id:")) {
        await bot.sendDocument(ADMIN_ID, session.data.cv.replace("file_id:", ""), {}, { caption: "קורות חיים" });
      } else if (session.data.cv.startsWith("photo_id:")) {
        await bot.sendPhoto(ADMIN_ID, session.data.cv.replace("photo_id:", ""), { caption: "קורות חיים" });
      }
    }

    // שלח לאדמין טקסט מוכן לשליחה לדובר (אם יש)
    if (session.data.internship_mentor) {
      const mentorMsg =
        `📋 *בקשת המלצה*\n\n` +
        `המועמד ${session.data.full_name} ציין שהתמחה אצל:\n` +
        `👤 ${session.data.internship_mentor}\n\n` +
        `*טקסט מוכן לשליחה בוואטסאפ:*\n` +
        `שלום, אני קוזו.\n` +
        `${session.data.full_name} שהתמחה אצלך ציין אותך בפרופיל שלו/ה.\n` +
        `אם תרצה/י להמליץ עליו/ה, פתח/י את הבוט כאן:\n` +
        `t.me/academiaB_advisor_bot`;
      await bot.sendMessage(ADMIN_ID, mentorMsg, { parse_mode: "Markdown" });
    }

    // חפש לשכות קיימות שמתאימות — חיבור אוטומטי
    const newCandidate = await getCandidateRecord(chatId);
    const matchingEmployers = newCandidate ? await findMatchingEmployers(newCandidate) : [];
    for (const employer of matchingEmployers) {
      await sendMatchSummary([newCandidate], employer, false);
    }
    // לשכות מושהות שביקשו להישאר לחיפושים עתידיים
    const futureEmployers = newCandidate ? await findFutureSearchEmployers(newCandidate) : [];
    for (const employer of futureEmployers) {
      await sendMatchSummary([newCandidate], employer, false);
    }
    await scheduleCandidateNudge(chatId);
    await scheduleMonthlyCheckin(chatId);
    await scheduleBimonthlyReminder(chatId);

    // הודעת קבלת פנים
    await bot.sendMessage(
      chatId,
      "ברוך הבא לקוזו 🤝\n\nהפרופיל שלך נשמר. מה קורה עכשיו?\nקוזו עובד ברקע ומחפש גופים מתאימים.\nברגע שיהיה התאמה — תשמע ממני ישירות.\n\nאין צורך לעשות כלום — קוזו עושה את השאר."
    );
    if (matchingEmployers.length > 0) {
      await bot.sendMessage(
        chatId,
        `העברתי את הפרטים שלך ל-${matchingEmployers.length} גופים שנראים לי מתאימים.\nברגע שיירשמו עוד גופים מתאימים — תשמע ממני 🤝`
      );
    }
    // פרופיל מלא
    const savedCand = await getCandidateRecord(chatId);
    if (savedCand) {
      await bot.sendMessage(chatId, buildProfileMessage(savedCand), {
        reply_markup: { inline_keyboard: [[{ text: "✏️ עדכן פרטים", callback_data: "UPDATE_PROFILE" }]] },
      });
    }

    // הפניית חברים
    await bot.sendMessage(
      chatId,
      "יש לך עמית שגם מחפש? שתף איתו את קוזו 🤝",
      {
        reply_markup: {
          inline_keyboard: [
            [{ text: "💬 שתף בוואטסאפ", url: "https://wa.me/?text=%D7%94%D7%A6%D7%98%D7%A8%D7%A3%20%D7%9C%D7%A7%D7%95%D7%96%D7%95%20%E2%80%94%20%D7%94%D7%9E%D7%A7%D7%95%D7%9D%20%D7%A9%D7%9E%D7%97%D7%91%D7%A8%20%D7%90%D7%A0%D7%A9%D7%99%20%D7%9E%D7%A7%D7%A6%D7%95%D7%A2%20%D7%91%D7%A2%D7%95%D7%9C%D7%9D%20%D7%94%D7%A4%D7%95%D7%9C%D7%99%D7%98%D7%99%3A%20t.me%2Fkozo_ai_bot" }],
            [{ text: "📱 שתף בטלגרם",   url: "https://t.me/share/url?url=t.me%2Fkozo_ai_bot&text=%D7%94%D7%A6%D7%98%D7%A8%D7%A3%20%D7%9C%D7%A7%D7%95%D7%96%D7%95" }],
            [{ text: "👥 שתף בפייסבוק", url: "https://www.facebook.com/sharer/sharer.php?u=t.me%2Fkozo_ai_bot" }],
          ],
        },
      }
    );
  } else {
    await bot.sendMessage(ADMIN_ID, `📥 לשכה חדשה נרשמה!\n\n${formatRecord("employer", session)}`);

    // חיפוש התאמות מיידי — חיבור אוטומטי
    const newEmployer = await getEmployerRecord(chatId);
    const employerForMatch = newEmployer || { ...session.data, telegram_id: chatId };
    const matches = await findMatches(employerForMatch);
    if (matches.length > 0) {
      await sendMatchSummary(matches, employerForMatch, true);
    } else {
      await bot.sendMessage(
        chatId,
        "אין עדיין יועצים מתאימים במאגר. ברגע שיירשם מישהו מתאים — תקבלו הודעה."
      );
    }
  }
  await exportExcel();
  delete sessions[chatId];
  } catch (err) {
    console.error("finishSession error:", err);
    try { await bot.sendMessage(chatId, "משהו השתבש, נסה שוב 🙏"); } catch (_) {}
    try { await bot.sendMessage(ADMIN_ID, `❌ שגיאה ב-finishSession (ID: ${chatId})\n${err.message}`); } catch (_) {}
  }
}

// ── /start ────────────────────────────────────────────────────────────────────

bot.onText(/\/start/, async (msg) => {
  const chatId = msg.chat.id;
  chatHistories[chatId] = [];

  // זיהוי רישום כפול
  const existingCandidate = await getCandidateRecord(chatId);
  const existingEmployer  = await getEmployerRecord(chatId);
  if (existingCandidate || existingEmployer) {
    const name = existingCandidate?.full_name || existingEmployer?.contact_name || "";
    await bot.sendMessage(
      chatId,
      `${name ? `שלום ${name}!\n` : ""}כבר רשום אצלנו. מה תרצה לעשות?`,
      {
        reply_markup: {
          inline_keyboard: [
            [{ text: "עדכן פרטים",      callback_data: "EXISTING_UPDATE" }],
            [{ text: "השהה אותי",       callback_data: "EXISTING_PAUSE" }],
            [{ text: "מצאתי עבודה 🎉", callback_data: "EXISTING_FOUND_JOB" }],
          ],
        },
      }
    );
    return;
  }

  sessions[chatId] = { stage: "awaiting_type", username: msg.from.username || "" };
  await bot.sendMessage(
    chatId,
    "היי, אני קוזו 👋\nמכיר את כולם אבל מחבר רק את המתאימים ביותר.\n\nשנתחיל?",
    {
      reply_markup: {
        inline_keyboard: [
          [{ text: "מחפש תפקיד 👤", callback_data: "CANDIDATE" }],
          [{ text: "מחפשים מישהו לתפקיד 🔎", callback_data: "EMPLOYER" }],
        ],
      },
    }
  );
});

// ── פקודות slash ─────────────────────────────────────────────────────────────

bot.onText(/\/help/, async (msg) => {
  await bot.sendMessage(msg.chat.id,
    `שאלות נפוצות 👋\n\n` +
    `👤 ליועצים:\n` +
    `❓ איך אני יודע שהפרופיל שלי פעיל?\n← כתוב /profile ותקבל את כל הפרטים שלך\n\n` +
    `❓ למה לא קיבלתי התאמות עדיין?\n← קוזו ממשיך לעבוד. ברגע שתהיה לשכה מתאימה — תקבל הודעה\n\n` +
    `❓ איך אני מעדכן פרטים?\n← כתוב /update או לחץ על כפתור עדכן בפרופיל\n\n` +
    `❓ איך יוצאים מהמאגר זמנית?\n← כתוב /pause. לחזרה — כתוב /resume\n\n` +
    `❓ מה קורה כשמגייס מקבל את הפרטים שלי?\n← הם רואים את הפרופיל שלך. אם מתאים — ייצרו קשר ישירות\n\n` +
    `🏛 למגייסים:\n` +
    `❓ איך מקבלים קוד גישה?\n← wa.me/972548028082\n\n` +
    `❓ איך מקבלים קורות חיים של מועמד?\n← לחצו על כפתור הקורות חיים ליד שם המועמד\n\n` +
    `❓ למה אני לא רואה מועמדים?\n← קוזו יעדכן כשיירשם מישהו שמתאים לדרישות שלכם\n\n` +
    `❓ האם המועמדים יודעים שפניתי אליהם?\n← לא. דיסקרטי לחלוטין\n\n` +
    `🔒 כללי:\n` +
    `❓ האם המידע שלי דיסקרטי?\n← כן. אף משתמש לא רואה את האחרים\n\n` +
    `❓ כמה עולה השירות?\n← חינמי ליועצים. מגייסים מצטרפים בהזמנה בלבד\n\n` +
    `❓ יש בעיה או שאלה?\n← wa.me/972548028082`
  );
});

bot.onText(/\/profile/, async (msg) => {
  const chatId = msg.chat.id;
  const cand = await getCandidateRecord(chatId);
  if (cand) {
    await bot.sendMessage(chatId, buildProfileMessage(cand), {
      reply_markup: { inline_keyboard: [[{ text: "✏️ עדכן פרטים", callback_data: "UPDATE_PROFILE" }]] },
    });
  } else {
    await bot.sendMessage(chatId, "לא מצאתי פרופיל. אם עוד לא נרשמת — כתוב /start");
  }
});

bot.onText(/\/update/, async (msg) => {
  const chatId = msg.chat.id;
  sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
  await bot.sendMessage(chatId, "יאללה, נעדכן את הפרטים שלך");
  await sendStep(chatId, sessions[chatId]);
});

bot.onText(/\/pause/, async (msg) => {
  const chatId = msg.chat.id;
  await pauseCandidate(chatId);
  await exportExcel();
  await bot.sendMessage(chatId, "הבנתי.\nעצרתי. כשתרצו לחזור — כתבו /resume", { parse_mode: "Markdown" });
  await bot.sendMessage(ADMIN_ID, `⏸ מועמד השהה את עצמו (ID: ${chatId})`);
});

bot.onText(/\/resume/, async (msg) => {
  const chatId = msg.chat.id;
  if (await isEmployerPaused(chatId)) {
    await resumeEmployer(chatId);
    await bot.sendMessage(chatId, "שמח שחזרתם 🤝\nאחזיר אתכם לרשימה. ברגע שיהיה מישהו מתאים, אחבר.");
    return;
  }
  const cand = await getCandidateRecord(chatId);
  if (!cand || cand.status === "active") {
    await bot.sendMessage(chatId, "כבר ברשימה שלי");
    return;
  }
  sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
  await bot.sendMessage(chatId, "שמחים שחזרת 🤝 כמה עדכונים קצרים ואחזיר אותך לרשימה:");
  await sendStep(chatId, sessions[chatId]);
});

// ── הודעות טקסט ──────────────────────────────────────────────────────────────

bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text?.trim();
  if (!text || text.startsWith("/")) return;
  try {

  // פקודות אדמין
  if (chatId === ADMIN_ID) {
    if (text === "מנהל") {
      await bot.sendMessage(chatId,
        "כלי ניהול קוזו 🛠\n\n" +
        "📊 סטטוס — נתונים כלליים\n" +
        "📋 טבלה — קובץ Excel מלא\n" +
        "📢 הודעה — שלח הודעה לקהל\n" +
        "👤 יועצים — רשימת יועצים פעילים\n" +
        "🏛 מגייסים — רשימת מגייסים פעילים\n" +
        "🎯 התאמות — מגייסים שקיבלו רשימת מועמדים\n" +
        "📞 יצירת קשר — מגייסים שלחצו קורות חיים\n" +
        "🔗 חיבורים — מגייסים שאישרו יצירת קשר\n" +
        "🗑 מחק [ID] — מחיקת משתמש מהמערכת"
      );
      return;
    }
    if (text.startsWith("שנה קוד ")) {
      const newCode = text.replace("שנה קוד ", "").trim();
      if (newCode) {
        EMPLOYER_ACCESS_CODE = newCode;
        await query(`INSERT INTO settings (key, value) VALUES ('employer_code', $1) ON CONFLICT (key) DO UPDATE SET value=$1`, [newCode]);
        await bot.sendMessage(chatId, `קוד עודכן ל-${newCode} ✅`);
      } else {
        await bot.sendMessage(chatId, "פורמט: שנה קוד [קוד חדש]");
      }
      return;
    }
    if (text === "מחק") {
      await bot.sendMessage(chatId, "לציין את ה-ID של המשתמש: מחק [ID]\nאת ה-ID תמצא בפקודות יועצים/מגייסים");
      return;
    }
    if (text.startsWith("מחק ")) {
      const targetId = Number(text.replace("מחק ", "").trim());
      if (!targetId) {
        await bot.sendMessage(chatId, "ID לא תקין. פורמט: מחק [ID]");
        return;
      }
      await Promise.all([
        query(`DELETE FROM candidates   WHERE telegram_id=$1`, [targetId]),
        query(`DELETE FROM employers    WHERE telegram_id=$1`, [targetId]),
        query(`DELETE FROM cv_requests  WHERE employer_id=$1 OR candidate_id=$1`, [targetId]),
        query(`DELETE FROM matches      WHERE candidate_id=$1 OR employer_id=$1`, [targetId]),
        query(`DELETE FROM ratings      WHERE employer_id=$1 OR candidate_id=$1`, [targetId]),
        query(`DELETE FROM recommendations WHERE candidate_id=$1`, [targetId]),
        query(`DELETE FROM access_requests WHERE telegram_id=$1`, [targetId]),
      ]);
      await exportExcel();
      await bot.sendMessage(chatId, `✅ המשתמש ${targetId} נמחק מהמערכת`);
      return;
    }
    if (text === "טבלה") { await sendExcel(); return; }
    if (text === "סטטוס") { await sendStatus(); return; }
    if (text === "קוזו במספרים") {
      const [candsRes, empsRes] = await Promise.all([
        query(`SELECT interests FROM candidates WHERE status='active'`),
        query(`SELECT COUNT(*) FROM employers WHERE status='active'`),
      ]);
      const countField = (field) => candsRes.rows.filter((c) => (c.interests || "").includes(field)).length;
      await bot.sendMessage(
        chatId,
        `📊 כרגע בקוזו\n\n` +
        `👤 יועצים פרלמנטריים: ${countField("ייעוץ פרלמנטרי")}\n` +
        `📢 דוברים: ${countField("דוברות")}\n` +
        `📱 אנשי סושיאל: ${countField("סושיאל ורשתות חברתיות")}\n` +
        `🎯 יועצים פוליטיים: ${countField("יועץ פוליטי")}\n` +
        `🎬 עריכת וידאו: ${countField("עריכת וידאו")}\n` +
        `🏛 מגייסים פעילים: ${empsRes.rows[0].count}`
      );
      return;
    }
    if (text === "סטטיסטיקות") {
      const [candsRes, empsRes, matchesRes, connectedRes] = await Promise.all([
        query(`SELECT COUNT(*) FROM candidates WHERE status='active'`),
        query(`SELECT COUNT(*) FROM employers WHERE status='active'`),
        query(`SELECT COUNT(*) FROM matches`),
        query(`SELECT COUNT(*) FROM cv_requests WHERE status='connected'`),
      ]);
      await bot.sendMessage(
        chatId,
        `📈 קוזו במספרים\n\n` +
        `👥 אנשי מקצוע רשומים: ${candsRes.rows[0].count}\n` +
        `🏛 מגייסים פעילים: ${empsRes.rows[0].count}\n` +
        `🎯 התאמות שבוצעו: ${matchesRes.rows[0].count}\n` +
        `🔗 חיבורים מוצלחים: ${connectedRes.rows[0].count}`
      );
      return;
    }
    if (text === "יועצים") {
      const res = await query(`SELECT full_name, interests, status FROM candidates WHERE status='active' ORDER BY created_at DESC`);
      if (res.rows.length === 0) {
        await bot.sendMessage(chatId, "אין יועצים פעילים כרגע.");
      } else {
        const lines = res.rows.map((r, i) =>
          `${i + 1}. ${r.full_name || "—"} | ${r.interests || "—"}`
        ).join("\n");
        await bot.sendMessage(chatId, `👥 יועצים פעילים (${res.rows.length}):\n\n${lines}`);
      }
      return;
    }
    if (text === "מגייסים") {
      const res = await query(`SELECT contact_name, org_type, fields, status FROM employers WHERE status='active' ORDER BY created_at DESC`);
      if (res.rows.length === 0) {
        await bot.sendMessage(chatId, "אין מגייסים פעילים כרגע.");
      } else {
        const lines = res.rows.map((r, i) =>
          `${i + 1}. ${r.contact_name || "—"} | ${r.org_type || "—"} | ${r.fields || "—"}`
        ).join("\n");
        await bot.sendMessage(chatId, `🏛 מגייסים פעילים (${res.rows.length}):\n\n${lines}`);
      }
      return;
    }
    if (text === "התאמות") {
      const res = await query(
        `SELECT e.contact_name, e.org_type, m.candidate_name, m.matched_at
         FROM matches m
         JOIN employers e ON e.telegram_id = m.employer_id
         WHERE m.status='active'
         ORDER BY e.contact_name, m.matched_at DESC
         LIMIT 100`
      );
      if (res.rows.length === 0) {
        await bot.sendMessage(chatId, "אין התאמות פעילות כרגע.");
      } else {
        // קיבוץ לפי מגייס
        const grouped = {};
        for (const r of res.rows) {
          const key = `${r.contact_name || "—"}|${r.org_type || "—"}`;
          if (!grouped[key]) grouped[key] = { contact_name: r.contact_name, org_type: r.org_type, candidates: [] };
          grouped[key].candidates.push({ name: r.candidate_name, date: r.matched_at });
        }
        const blocks = Object.values(grouped).map((g) => {
          const candLines = g.candidates.map((c) => {
            const date = c.date ? new Date(c.date).toLocaleDateString("he-IL") : "—";
            return `   👤 ${c.name || "—"} — נשלח ${date}`;
          }).join("\n");
          return `🏛 ${g.contact_name || "—"} (${g.org_type || "—"})\n${candLines}`;
        });
        const total = res.rows.length;
        await bot.sendMessage(chatId, `🎯 התאמות פעילות (${total} יועצים, ${Object.keys(grouped).length} מגייסים)\n\n${blocks.join("\n\n")}`);
      }
      return;
    }
    if (text === "יצירת קשר") {
      const res = await query(
        `SELECT e.contact_name, e.org_type, c.full_name AS candidate_name, cr.requested_at, cr.status
         FROM cv_requests cr
         JOIN employers e ON e.telegram_id = cr.employer_id
         JOIN candidates c ON c.telegram_id = cr.candidate_id
         ORDER BY cr.requested_at DESC LIMIT 30`
      );
      if (res.rows.length === 0) {
        await bot.sendMessage(chatId, "אין יצירות קשר כרגע.");
      } else {
        const statusLabel = (s) => s === "connected" ? "חיבור מוצלח ✅" : s === "rejected" ? "לא מתאים ❌" : s === "in_progress" ? "בתהליך ⏳" : "פתוח";
        const lines = res.rows.map((r, i) => {
          const date = r.requested_at ? new Date(r.requested_at).toLocaleDateString("he-IL") : "—";
          return `${i + 1}. ${r.contact_name || "—"} → ${r.candidate_name || "—"} | ${statusLabel(r.status)} | ${date}`;
        }).join("\n");
        await bot.sendMessage(chatId, `📞 יצירות קשר (${res.rows.length}):\n\n${lines}`);
      }
      return;
    }
    if (text === "חיבורים") {
      const res = await query(
        `SELECT e.contact_name, e.org_type, c.full_name AS candidate_name, cr.updated_at
         FROM cv_requests cr
         JOIN employers e ON e.telegram_id = cr.employer_id
         JOIN candidates c ON c.telegram_id = cr.candidate_id
         WHERE cr.status='connected'
         ORDER BY cr.updated_at DESC LIMIT 30`
      );
      if (res.rows.length === 0) {
        await bot.sendMessage(chatId, "אין חיבורים מוצלחים כרגע.");
      } else {
        const lines = res.rows.map((r, i) => {
          const date = r.updated_at ? new Date(r.updated_at).toLocaleDateString("he-IL") : "—";
          return `${i + 1}. ${r.contact_name || "—"} (${r.org_type || "—"}) → ${r.candidate_name || "—"} | ${date}`;
        }).join("\n");
        await bot.sendMessage(chatId, `🔗 חיבורים מוצלחים (${res.rows.length}):\n\n${lines}`);
      }
      return;
    }
    if (text === "הודעה") {
      sessions[chatId] = { stage: "broadcast", step: "audience" };
      await bot.sendMessage(chatId, "לאיזה קהל?", {
        reply_markup: { inline_keyboard: [
          [{ text: "יועצים", callback_data: "BROADCAST_AUDIENCE_candidates" }],
          [{ text: "מגייסים", callback_data: "BROADCAST_AUDIENCE_employers" }],
          [{ text: "כולם", callback_data: "BROADCAST_AUDIENCE_all" }],
        ]},
      });
      return;
    }
    const adminSession = sessions[chatId];
    if (adminSession?.stage === "broadcast" && adminSession?.step === "text") {
      adminSession.message = text;
      adminSession.step = "whatsapp";
      await bot.sendMessage(chatId, "להוסיף לינק וואטסאפ?", {
        reply_markup: { inline_keyboard: [[
          { text: "כן", callback_data: "BROADCAST_WHATSAPP_yes" },
          { text: "לא", callback_data: "BROADCAST_WHATSAPP_no" },
        ]]},
      });
      return;
    }
  }

  const session = sessions[chatId];

  // ── מצב השהייה ──
  if (!session || session.stage === "free_chat") {
    // בדיקת מילות מפתח לפני Claude
    const lower = text.toLowerCase();

    if (lower.includes("מצאתי עבודה")) {
      const cand = await getCandidateRecord(chatId);
      sessions[chatId] = { stage: "found_job_source", candidateName: cand?.full_name || "" };
      await bot.sendMessage(
        chatId,
        "מצוין! 🎉 איך מצאת?",
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "דרך קוזו 🤝",  callback_data: `JOB_SOURCE_kozo_${chatId}` },
              { text: "ממקום אחר",     callback_data: `JOB_SOURCE_other_${chatId}` },
            ]],
          },
        }
      );
      return;
    }

    if (lower.includes("השהה אותי")) {
      await pauseCandidate(chatId);
      await exportExcel();
      await bot.sendMessage(
        chatId,
        "הבנתי.\nעצרתי. כשתרצו לחזור — כתבו *החזר אותי לפעילות*",
        { parse_mode: "Markdown" }
      );
      await bot.sendMessage(ADMIN_ID, `⏸ מועמד השהה את עצמו (ID: ${chatId})`);
      sessions[chatId] = { stage: "free_chat" };
      return;
    }

    if (lower.includes("החזר אותי לפעילות")) {
      if (await isEmployerPaused(chatId)) {
        await resumeEmployer(chatId);
        await bot.sendMessage(chatId, "שמח שחזרתם 🤝\nאחזיר אתכם לרשימה. ברגע שיהיה מישהו מתאים, אחבר.");
        return;
      }
      const cand = await getCandidateRecord(chatId);
      if (!cand || cand.status === "active") {
        await bot.sendMessage(chatId, "כבר ברשימה שלי");
        return;
      }
      sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
      await bot.sendMessage(chatId, "שמחים שחזרת 🤝 כמה עדכונים קצרים ואחזיר אותך לרשימה:");
      await sendStep(chatId, sessions[chatId]);
      return;
    }

    if (lower.includes("הפרופיל שלי")) {
      const cand = await getCandidateRecord(chatId);
      if (cand) {
        await bot.sendMessage(chatId, buildProfileMessage(cand), {
          reply_markup: { inline_keyboard: [[{ text: "✏️ עדכן פרטים", callback_data: "UPDATE_PROFILE" }]] },
        });
      } else {
        await bot.sendMessage(chatId, "לא מצאתי פרופיל. אם עוד לא נרשמת — כתוב /start");
      }
      return;
    }

    if (lower.includes("עדכן פרטים")) {
      sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
      await bot.sendMessage(chatId, "יאללה, נעדכן את הפרטים שלך");
      await sendStep(chatId, sessions[chatId]);
      return;
    }

    // שיחה חופשית עם Claude
    if (!session || session.stage === "free_chat") {
      const reply = await claudeChat(chatId, text);
      await bot.sendMessage(chatId, reply);
      return;
    }
  }

  // ── שאלון בקשת גישה ──
  if (session && session.stage === "access_questionnaire") {
    const ACCESS_QUESTIONS = [
      { key: "full_name",   q: "מה שמך המלא?" },
      { key: "role",        q: "מה תחום העיסוק שלך?" },
      { key: "job_search",  q: "איזה תפקיד מחפשים?" },
      { key: "heard_from",  q: "איך שמעת על קוזו?" },
    ];

    const stepIndex = session.step;
    if (stepIndex < ACCESS_QUESTIONS.length) {
      session.answers[ACCESS_QUESTIONS[stepIndex].key] = text;
      session.step++;

      if (session.step < ACCESS_QUESTIONS.length) {
        await bot.sendMessage(chatId, ACCESS_QUESTIONS[session.step].q);
      } else {
        // סיום שאלון — שלח לאדמין
        const a = session.answers;
        await bot.sendMessage(
          ADMIN_ID,
          `📨 בקשת הצטרפות חדשה\n\n` +
          `👤 שם: ${a.full_name}\n` +
          `📱 נייד: ${session.phone}\n` +
          `💼 תחום: ${a.role}\n` +
          `🔍 מחפש: ${a.job_search}\n` +
          `📣 שמע עלינו: ${a.heard_from}\n` +
          `🆔 טלגרם: ${chatId}`,
          {
            reply_markup: {
              inline_keyboard: [[
                { text: "אשר ✅", callback_data: `APPROVE_ACCESS_${chatId}_${session.phone}` },
                { text: "דחה ❌", callback_data: `DENY_ACCESS_${chatId}` },
              ]],
            },
          }
        );
        // שמור בקשת גישה ב-DB
        await query(
          `INSERT INTO access_requests (telegram_id, telegram_username, phone, full_name, role, job_search, heard_from)
           VALUES ($1, $2, $3, $4, $5, $6, $7)`,
          [chatId, session.username || "", session.phone, a.full_name, a.role, a.job_search, a.heard_from]
        );

        await bot.sendMessage(chatId, "שמרתי 🙏\nאחזור אליך בהקדם.");
        delete sessions[chatId];
      }
    }
    return;
  }

  // ── כתיבת המלצה ──
  if (session && session.stage === "writing_recommendation") {
    const rec = text;
    await saveRecommendationText(session.candidateId, rec, "ממליץ");
    await bot.sendMessage(chatId, "תודה רבה! 🙏\nשמרתי את ההמלצה. היא תצורף לפרופיל ותגיע למקום הנכון 🤝");
    await bot.sendMessage(ADMIN_ID, `⭐ התקבלה המלצה על מועמד ID: ${session.candidateId}\n\n"${rec}"`);
    delete sessions[chatId];
    return;
  }

  // ── קוד אישור לשכה/עירייה ──
  if (session && session.stage === "awaiting_employer_code") {
    if (text.trim() === EMPLOYER_ACCESS_CODE) {
      sessions[chatId] = { ...newSession("employer", session.username), data: {} };
      await bot.sendMessage(chatId, "קוד אומת ✅\n\nהיי, אני קוזו 🤝\nבואו נכניס אתכם למאגר ונתחיל לחבר:");
      await sendStep(chatId, sessions[chatId]);
    } else {
      await bot.sendMessage(chatId, "הקוד לא מוכר לי 🙏 לקבלת קוד תקין, פנו אלינו ישירות.\nלפניה ישירה: wa.me/972548028082");
    }
    return;
  }

  // ── אימות נייד (יועצים בלבד) ──
  if (session && session.stage === "awaiting_phone") {
    if (!isValidPhone(text)) {
      await bot.sendMessage(chatId, "הנייד לא עובד לי 🙏 אפשר לנסות שוב?\nלפניה ישירה: wa.me/972548028082");
      return;
    }
    const phone = normalizePhone(text);

    sessions[chatId] = { ...newSession("candidate", session.username), phone, data: { phone } };
    await bot.sendMessage(chatId, "מעולה 🤝 בואו נכניס אותך למאגר:");
    await sendStep(chatId, sessions[chatId]);
    return;
  }

  if (session.stage === "awaiting_type") {
    await bot.sendMessage(chatId, "לא הבנתי 🙏 אפשר להתחיל עם /start\nלעזרה ישירה: wa.me/972548028082");
    return;
  }
  if (!session.verified && session.stage !== "updating") return;

  const steps = getSteps(session.type);
  const step = steps[session.step];
  if (!step) return;

  if (step.type === "text") {
    session.data[step.key] = text;
    session.step++;
    await sendStep(chatId, session);
  } else if (step.type === "email") {
    if (!isValidEmail(text)) {
      await bot.sendMessage(chatId, "המייל לא נראה תקין 🙏\nלדוגמה: name@gmail.com\nלפניה ישירה: wa.me/972548028082");
      return;
    }
    session.data[step.key] = text;
    session.step++;
    await sendStep(chatId, session);
  } else if (step.type === "file") {
    await bot.sendMessage(chatId, "שלח את הקובץ כצירוף, לא כתמונה 📎");
  } else {
    await bot.sendMessage(chatId, "לחץ על אחת מהאפשרויות למעלה 👆");
  }
  } catch (err) {
    console.error("message handler error:", err);
    try { await bot.sendMessage(chatId, "משהו השתבש, נסה שוב 🙏"); } catch (_) {}
    try { await bot.sendMessage(ADMIN_ID, `❌ שגיאה ב-message handler (ID: ${chatId}, text: ${text})\n${err.message}`); } catch (_) {}
  }
});

// ── קבצים ────────────────────────────────────────────────────────────────────

bot.on("document", async (msg) => {
  const chatId = msg.chat.id;
  const session = sessions[chatId];
  if (!session || (!session.verified && session.stage !== "updating")) return;
  const step = getSteps(session.type)[session.step];
  if (!step || step.type !== "file") return;
  session.data[step.key] = `file_id:${msg.document.file_id}`;
  session.step++;
  await sendStep(chatId, session);
});

bot.on("photo", async (msg) => {
  const chatId = msg.chat.id;
  const session = sessions[chatId];
  if (!session || (!session.verified && session.stage !== "updating")) return;
  const step = getSteps(session.type)[session.step];
  if (!step || step.type !== "file") return;
  const photo = msg.photo[msg.photo.length - 1];
  session.data[step.key] = `photo_id:${photo.file_id}`;
  session.step++;
  await sendStep(chatId, session);
});

// ── לחיצות כפתור ─────────────────────────────────────────────────────────────

bot.on("callback_query", async (cbQuery) => {
  const chatId = cbQuery.message.chat.id;
  const data   = cbQuery.data;
  try { await bot.answerCallbackQuery(cbQuery.id); } catch (_) {}
  try {

  // בקשת הצטרפות — התחל שאלון
  if (data === "REQUEST_ACCESS") {
    sessions[chatId] = {
      stage: "access_questionnaire",
      step: 0,
      phone: sessions[chatId]?.phone || "לא ידוע",
      username: sessions[chatId]?.username || "",
      answers: {}
    };
    await bot.sendMessage(chatId, "שאלה-שתיים ואחזור אליך:\n\nמה שמך המלא?");
    return;
  }

  // אישור גישה
  if (data.startsWith("APPROVE_ACCESS_")) {
    const parts = data.split("_");
    const targetChatId = parts[2];
    const phone = parts[3];
    await approvePhone(phone);
    // עדכן סטטוס בקשה
    await query(
      `UPDATE access_requests SET status='approved', approved_at=NOW() WHERE telegram_id=$1`,
      [Number(targetChatId)]
    );
    await bot.sendMessage(Number(targetChatId), "אושר! 🎉\nשלח /start ונתחיל 🤝");
    await bot.sendMessage(ADMIN_ID, `✅ אושר! נייד ${phone} נוסף לרשימה.`);
    return;
  }

  // דחיית גישה
  if (data.startsWith("DENY_ACCESS_")) {
    const targetChatId = data.split("_")[2];
    // עדכן סטטוס בקשה
    await query(
      `UPDATE access_requests SET status='denied', denied_at=NOW() WHERE telegram_id=$1`,
      [Number(targetChatId)]
    );
    await bot.sendMessage(Number(targetChatId), "מצטערים, הפעם לא הצלחנו לאשר 🙏\nלפניה ישירה: wa.me/972548028082");
    await bot.sendMessage(ADMIN_ID, "❌ הבקשה נדחתה.");
    return;
  }

  // שליחת קורות חיים לפי בקשה
  if (data.startsWith("CV_")) {
    const parts = data.split("_");
    const candidateId = Number(parts[1]);
    const employerId  = Number(parts[2]);
    const candidate   = await getCandidateRecord(candidateId);
    if (!candidate || !candidate.cv) {
      await bot.sendMessage(chatId, "לא מצאתי קורות חיים 🙏");
      return;
    }
    const internshipLine = candidate.is_intern === "כן ✅" && candidate.internship_mentor
      ? `\n🏛 התמחות: דוברות הכנסת — שנה | אצל: ${candidate.internship_mentor}`
      : "";
    const cvCaption = `קורות חיים — ${candidate.full_name || "מועמד"}${internshipLine}`;
    if (candidate.cv.startsWith("file_id:")) {
      await bot.sendDocument(chatId, candidate.cv.replace("file_id:", ""), {}, { caption: cvCaption });
    } else if (candidate.cv.startsWith("photo_id:")) {
      await bot.sendPhoto(chatId, candidate.cv.replace("photo_id:", ""), { caption: cvCaption });
    }
    if (!candidate.phone) {
      await bot.sendMessage(chatId, `ליצירת קשר ישירה עם ${candidate.full_name || "המועמד"} — פנו לקוזו: wa.me/972548028082`);
    }
    await query(
      `INSERT INTO cv_requests (employer_id, candidate_id)
       VALUES ($1, $2)
       ON CONFLICT (employer_id, candidate_id) DO UPDATE SET requested_at=NOW(), status='matched', updated_at=NOW()`,
      [employerId, candidateId]
    );
    await scheduleCVFollowUp(candidateId, employerId, candidate.full_name || "מועמד");
    return;
  }

  // מעקב קורות חיים
  if (data.startsWith("CVFOLLOWUP_")) {
    const parts = data.split("_");
    const action      = parts[1];
    const candidateId = Number(parts[2]);
    const employerId  = Number(parts[3]);
    const candidate   = await getCandidateRecord(candidateId);
    const employer    = await getEmployerRecord(employerId);
    const candidateName = candidate?.full_name || "מועמד";
    const employerName  = employer?.contact_name || "לשכה";

    if (action === "INPROGRESS") {
      await query(
        `UPDATE cv_requests SET status='in_progress', updated_at=NOW() WHERE employer_id=$1 AND candidate_id=$2`,
        [employerId, candidateId]
      );
      await bot.sendMessage(chatId, "תודה! אחזור אליך בעוד שבוע 🤝");
      await bot.sendMessage(ADMIN_ID, `עדכון על ${candidateName} ↔ ${employerName}: עוד בתהליך ⏳`);
      await scheduleCVFollowUp(candidateId, employerId, candidateName);
    } else {
      const status = action === "CONTACTED" ? "connected" : "rejected";
      const label  = action === "CONTACTED" ? "חיבור מוצלח ✅" : "לא מתאים ❌";
      await query(
        `UPDATE cv_requests SET status=$3, updated_at=NOW() WHERE employer_id=$1 AND candidate_id=$2`,
        [employerId, candidateId, status]
      );
      await bot.sendMessage(chatId, "תודה על העדכון 🙏");
      await bot.sendMessage(ADMIN_ID, `עדכון על ${candidateName} ↔ ${employerName}: ${label}`);
      if (action === "CONTACTED") {
        await scheduleRatingRequest(employerId, candidateId, candidateName);
      }
    }
    return;
  }

  if (data === "UPDATE_PROFILE") {
    sessions[chatId] = { ...newSession("update", cbQuery.from?.username || ""), stage: "updating" };
    await bot.sendMessage(chatId, "יאללה, נעדכן את הפרטים שלך");
    await sendStep(chatId, sessions[chatId]);
    return;
  }

  // רישום כפול — בחירת פעולה
  if (data === "EXISTING_UPDATE") {
    sessions[chatId] = { ...newSession("update", ""), stage: "updating" };
    await bot.sendMessage(chatId, "יאללה, נעדכן את הפרטים שלך");
    await sendStep(chatId, sessions[chatId]);
    return;
  }
  if (data === "EXISTING_PAUSE") {
    await pauseCandidate(chatId);
    await pauseEmployer(chatId);
    await bot.sendMessage(chatId, "הבנתי. עצרתי. כשתרצו לחזור — כתבו *החזר אותי לפעילות*", { parse_mode: "Markdown" });
    await bot.sendMessage(ADMIN_ID, `⏸ משתמש השהה את עצמו (ID: ${chatId})`);
    return;
  }
  if (data === "EXISTING_FOUND_JOB") {
    await archiveCandidate(chatId);
    await exportExcel();
    await bot.sendMessage(chatId, "כיף לשמוע! 🎉 אם יום אחד תרצו לחזור — /start תמיד פתוח");
    await bot.sendMessage(ADMIN_ID, `📦 מועמד הועבר לארכיון (ID: ${chatId}), מצא עבודה`);
    return;
  }

  // מעקב אחרי מגייס שלא פתח קורות חיים
  if (data.startsWith("EMPLIST_STILL_")) {
    await bot.sendMessage(chatId, "מצוין, ממשיכים לחפש עבורכם 🤝");
    return;
  }
  if (data.startsWith("EMPLIST_FOUND_")) {
    const employerId = Number(data.replace("EMPLIST_FOUND_", ""));
    await pauseEmployer(employerId);
    await exportExcel();
    await bot.sendMessage(chatId, "מעולה! 🎉 נסמן אתכם כלא פעיל. בהצלחה!");
    await bot.sendMessage(ADMIN_ID, `✅ לשכה מצאה מועמד, הועברה לסטטוס לא פעיל (ID: ${employerId})`);
    return;
  }

  // check-in חודשי ליועץ
  if (data.startsWith("CHECKIN_STILL_")) {
    const candidateId = Number(data.replace("CHECKIN_STILL_", ""));
    await bot.sendMessage(chatId, "ממשיכים לחפש עבורך 🤝");
    await scheduleMonthlyCheckin(candidateId);
    return;
  }
  if (data.startsWith("CHECKIN_FOUND_")) {
    const candidateId = Number(data.replace("CHECKIN_FOUND_", ""));
    await archiveCandidate(candidateId);
    await exportExcel();
    await bot.sendMessage(chatId, "מעולה! 🎉 כיף לשמוע. אם יום אחד תרצו לחזור — /start תמיד פתוח");
    await bot.sendMessage(ADMIN_ID, `📦 מועמד הועבר לארכיון (ID: ${candidateId}), מצא עבודה (check-in חודשי)`);
    return;
  }

  // מקור עבודה
  if (data.startsWith("JOB_SOURCE_")) {
    const parts = data.split("_");
    const source = parts[2];
    const candidateId = Number(parts[3]);
    const cand = await getCandidateRecord(candidateId);
    const name = cand?.full_name || `ID:${candidateId}`;
    await query(`UPDATE candidates SET job_source=$1 WHERE telegram_id=$2`, [source, candidateId]);
    await archiveCandidate(candidateId);
    await exportExcel();
    if (source === "kozo") {
      await bot.sendMessage(chatId, "מעולה! 🎉 שמחים שקוזו עזר — בהצלחה בתפקיד!");
      await bot.sendMessage(ADMIN_ID, `🏆 חיבור מוצלח דרך קוזו!\n\n${name} מצא עבודה דרך קוזו ✅`);
    } else {
      await bot.sendMessage(chatId, "תודה על העדכון 🙏 בהצלחה בתפקיד החדש!");
      await bot.sendMessage(ADMIN_ID, `📦 יועץ מצא עבודה ממקום אחר: ${name}`);
    }
    delete sessions[chatId];
    return;
  }

  // שידור כללי — אדמין בלבד
  if (chatId === ADMIN_ID && data.startsWith("BROADCAST_")) {
    const bs = sessions[chatId];
    if (!bs || bs.stage !== "broadcast") return;

    if (data.startsWith("BROADCAST_AUDIENCE_")) {
      bs.audience = data.replace("BROADCAST_AUDIENCE_", "");
      bs.fieldSelect = [];
      bs.step = "field";
      await bot.sendMessage(chatId, "רק תחום מסוים?\nאפשר לסמן כמה ולחץ סיום ✓", {
        reply_markup: { inline_keyboard: [
          [{ text: "ייעוץ פרלמנטרי", callback_data: "BROADCAST_FIELD_ייעוץ פרלמנטרי" }],
          [{ text: "דוברות",          callback_data: "BROADCAST_FIELD_דוברות" }],
          [{ text: "סושיאל",          callback_data: "BROADCAST_FIELD_סושיאל ורשתות חברתיות" }],
          [{ text: "יועץ פוליטי",    callback_data: "BROADCAST_FIELD_יועץ פוליטי" }],
          [{ text: "עריכת וידאו",     callback_data: "BROADCAST_FIELD_עריכת וידאו" }],
          [{ text: "כולם",  callback_data: "BROADCAST_FIELD_ALL" },
           { text: "סיום ✓", callback_data: "BROADCAST_FIELD_DONE" }],
        ]},
      });
      return;
    }

    if (data.startsWith("BROADCAST_FIELD_")) {
      if (data === "BROADCAST_FIELD_ALL") {
        bs.fields = [];
        bs.step = "political";
      } else if (data === "BROADCAST_FIELD_DONE") {
        if (!bs.fieldSelect || bs.fieldSelect.length === 0) {
          await bot.sendMessage(chatId, "בחר לפחות תחום אחד, או לחץ 'כולם'");
          return;
        }
        bs.fields = [...bs.fieldSelect];
        bs.step = "political";
      } else {
        if (!bs.fieldSelect) bs.fieldSelect = [];
        const field = data.replace("BROADCAST_FIELD_", "");
        if (bs.fieldSelect.includes(field)) {
          bs.fieldSelect = bs.fieldSelect.filter((f) => f !== field);
          await bot.sendMessage(chatId, `➖ ${field} הוסר\nהמשך לבחור או לחץ סיום ✓`);
        } else {
          bs.fieldSelect.push(field);
          await bot.sendMessage(chatId, `➕ ${field}\nהמשך לבחור או לחץ סיום ✓`);
        }
        return;
      }
      await bot.sendMessage(chatId, "רק צד פוליטי?", {
        reply_markup: { inline_keyboard: [
          [{ text: "קואליציה",  callback_data: "BROADCAST_POLITICAL_קואליציה" }],
          [{ text: "אופוזיציה", callback_data: "BROADCAST_POLITICAL_אופוזיציה" }],
          [{ text: "כולם",      callback_data: "BROADCAST_POLITICAL_all" }],
        ]},
      });
      return;
    }

    if (data.startsWith("BROADCAST_POLITICAL_")) {
      bs.political = data.replace("BROADCAST_POLITICAL_", "");
      bs.step = "text";
      await bot.sendMessage(chatId, "מה ההודעה?");
      return;
    }

    if (data.startsWith("BROADCAST_WHATSAPP_")) {
      bs.addWhatsapp = data.replace("BROADCAST_WHATSAPP_", "") === "yes";
      bs.step = "confirm";
      const ids = await getBroadcastRecipients(bs);
      await bot.sendMessage(chatId, `תשלח ל-${ids.size} אנשים. מאשר?`, {
        reply_markup: { inline_keyboard: [[
          { text: "שלח ✅", callback_data: "BROADCAST_SEND" },
          { text: "בטל ❌",  callback_data: "BROADCAST_CANCEL" },
        ]]},
      });
      return;
    }

    if (data === "BROADCAST_SEND") {
      const sent = await sendBroadcast(bs);
      delete sessions[chatId];
      await bot.sendMessage(chatId, `ההודעה נשלחה ל-${sent} אנשים ✅`);
      return;
    }

    if (data === "BROADCAST_CANCEL") {
      delete sessions[chatId];
      await bot.sendMessage(chatId, "בוטל.");
      return;
    }

    return;
  }

  // ביקורת חיבור
  if (data.startsWith("RATING_")) {
    const parts = data.split("_");
    const stars      = Number(parts[1]);
    const employerId = Number(parts[2]);
    const candidateId = Number(parts[3]);
    const candidate  = await getCandidateRecord(candidateId);
    const employer   = await getEmployerRecord(employerId);
    await query(
      `INSERT INTO ratings (employer_id, candidate_id, stars)
       VALUES ($1, $2, $3)
       ON CONFLICT (employer_id, candidate_id) DO UPDATE SET stars=$3, created_at=NOW()`,
      [employerId, candidateId, stars]
    );
    await bot.sendMessage(chatId, "תודה על הפידבק! 🙏");
    await bot.sendMessage(
      ADMIN_ID,
      `⭐ דירוג חיבור\n\n${employer?.contact_name || "לשכה"} דירג את ${candidate?.full_name || "מועמד"}: ${"⭐".repeat(stars)}`
    );
    return;
  }

  // רענון התאמות — מגייס מבקש עוד מועמדים
  if (data.startsWith("REFRESH_MATCHES_")) {
    const employerId = Number(data.replace("REFRESH_MATCHES_", ""));
    const employer   = await getEmployerRecord(employerId);
    if (!employer) { await bot.sendMessage(chatId, "לא מצאתי את הפרופיל שלכם 🙏"); return; }
    const matches = await findMatches(employer);
    if (matches.length > 0) {
      await sendMatchSummary(matches, employer, true);
    } else {
      await bot.sendMessage(chatId, "אין עדיין מועמדים נוספים מתאימים. ברגע שיירשם מישהו — תקבלו עדכון 🤝");
    }
    return;
  }

  // הפסקת הצעות — לשכה/עירייה
  if (data.startsWith("STOP_OFFERS_EMPLOYER_")) {
    const employerId = Number(data.replace("STOP_OFFERS_EMPLOYER_", ""));
    await pauseEmployer(employerId);
    await bot.sendMessage(chatId, "הבנתי 🙏 הורדתי אתכם מהרשימה.\nכשתהיו מוכנים לחזור, כתבו *החזר אותי לפעילות*", { parse_mode: "Markdown" });
    await bot.sendMessage(ADMIN_ID, `⏸ לשכה/עירייה הפסיקה לקבל הצעות (ID: ${employerId})`);
    return;
  }

  // הפסקת הצעות — מועמד
  if (data.startsWith("STOP_OFFERS_CANDIDATE_")) {
    const candidateId = Number(data.replace("STOP_OFFERS_CANDIDATE_", ""));
    await pauseCandidate(candidateId);
    await exportExcel();
    await bot.sendMessage(chatId, "הבנתי 🙏 הורדתי אותך מהרשימה.\nכשתהיו מוכנים לחזור, כתבו *החזר אותי לפעילות*", { parse_mode: "Markdown" });
    await bot.sendMessage(ADMIN_ID, `⏸ מועמד הפסיק לקבל הצעות (ID: ${candidateId})`);
    return;
  }

  // המלצה — כן
  if (data.startsWith("REC_YES_")) {
    const candidateId = Number(data.split("_")[2]);
    sessions[chatId] = { stage: "writing_recommendation", candidateId };
    await bot.sendMessage(chatId, "כתוב/י בחופשיות. ממליצים טובים עושים את ההבדל:");
    return;
  }

  if (data.startsWith("REC_NO_")) {
    await bot.sendMessage(chatId, "מובן, תודה על הזמן 🙏");
    delete sessions[chatId];
    return;
  }

  // follow-up אחרי שבוע
  if (data.startsWith("FOLLOWUP_YES_")) {
    const parts = data.split("_");
    const candidateId = Number(parts[2]);
    const employerId  = Number(parts[3]);
    await bot.sendMessage(ADMIN_ID, "✅ נרשם. ההתאמה עדיין פעילה. נבדוק שוב בשבוע הבא.");
    // שלח follow-up נוסף בעוד שבוע
    await scheduleFollowUp(candidateId, employerId);
    return;
  }

  if (data.startsWith("FOLLOWUP_NO_")) {
    const parts = data.split("_");
    const candidateId = Number(parts[2]);
    const employerId  = Number(parts[3]);
    // עדכן סטטוס בהיסטוריה
    await query(
      `UPDATE matches SET status='closed', closed_at=NOW() WHERE candidate_id=$1 AND employer_id=$2`,
      [candidateId, employerId]
    );
    await exportExcel();
    await bot.sendMessage(ADMIN_ID, "❌ נרשם. ההתאמה נסגרה. המועמד לא יוצע לאותה לשכה שוב.");
    return;
  }

  // בחירת סוג פונה
  const session = sessions[chatId];
  if (!session) return;

  if (session.stage === "awaiting_type") {
    if (data === "CANDIDATE") {
      // יועץ — צריך לאמת נייד קודם
      sessions[chatId] = { stage: "awaiting_phone", username: session.username, pendingType: "candidate" };
      await bot.sendMessage(chatId, "שלח לי את הנייד שלך ונתחיל:");
    } else if (data === "EMPLOYER") {
      // לשכה/עירייה — דורשת קוד אישור לפני שמתחילים
      sessions[chatId] = { stage: "awaiting_employer_code", username: session.username };
      await bot.sendMessage(
        chatId,
        "כדי שנוכל לחבר אתכם לנכונים, נצטרך קוד אישור קצר.\n\n📞 לקבלת הקוד:",
        {
          reply_markup: {
            inline_keyboard: [[
              { text: "💬 צור קשר לקבלת קוד", url: ADMIN_PHONE_LINK },
            ]],
          },
        }
      );
      await bot.sendMessage(chatId, "יש לכם קוד? שלחו אותו כאן:");
    }
    return;
  }

  if (session.stage === "awaiting_employer_code") return;

  if (!session.verified && session.stage !== "updating") return;

  const steps = getSteps(session.type);
  const step  = steps[session.step];
  if (!step) return;

  if (step.type === "single") {
    session.data[step.key] = data;
    await bot.sendMessage(chatId, `✅ ${data}`);
    session.step++;
    await sendStep(chatId, session);
  } else if (step.type === "multi") {
    if (data === "סיום ✓") {
      if (session.multiSelect.length === 0) {
        await bot.sendMessage(chatId, "בחר לפחות אפשרות אחת לפני שמסיימים");
        return;
      }
      session.data[step.key] = session.multiSelect.join(", ");
      session.multiSelect = [];
      await bot.sendMessage(chatId, `✅ נבחר: ${session.data[step.key]}`);
      session.step++;
      await sendStep(chatId, session);
    } else {
      if (!session.multiSelect.includes(data)) {
        session.multiSelect.push(data);
        await bot.sendMessage(chatId, `➕ ${data}\nהמשך לבחור או לחץ סיום ✓`);
      } else {
        session.multiSelect = session.multiSelect.filter((i) => i !== data);
        await bot.sendMessage(chatId, `➖ ${data} הוסר`);
      }
    }
  }

  } catch (err) {
    console.error("callback_query error:", err);
    try { await bot.sendMessage(chatId, "משהו השתבש, נסה שוב 🙏"); } catch (_) {}
    try { await bot.sendMessage(ADMIN_ID, `❌ שגיאה ב-callback_query (ID: ${chatId}, data: ${data})\n${err.message}`); } catch (_) {}
  }
});

// ── שידור כללי (אדמין) ───────────────────────────────────────────────────────

function broadcastFieldMatches(row, fields, isCandidate) {
  if (!fields || fields.length === 0) return true;
  const text = isCandidate ? (row.interests || "") : (row.fields || "");
  const rowFields = text.split(", ").map((s) => s.trim());
  return fields.some((f) => rowFields.includes(f));
}

function broadcastPoliticalMatches(row, political) {
  if (!political || political === "all") return true;
  const side = row.political_side || "";
  if (!side || side === "שניהם" || side === "לא רלוונטי") return true;
  return side === political;
}

async function getBroadcastRecipients(bs) {
  const audience  = bs.audience  || "all";
  const fields    = bs.fields    || [];
  const political = bs.political || "all";
  const ids = new Set();
  if (audience === "candidates" || audience === "all") {
    const res = await query(`SELECT * FROM candidates WHERE status='active'`);
    console.log(`[broadcast] candidates active: ${res.rows.length}, fields: ${JSON.stringify(fields)}, political: ${political}`);
    for (const c of res.rows) {
      if (!broadcastFieldMatches(c, fields, true)) continue;
      if (!broadcastPoliticalMatches(c, political)) continue;
      ids.add(c.telegram_id);
    }
  }
  if (audience === "employers" || audience === "all") {
    const res = await query(`SELECT * FROM employers WHERE status='active'`);
    console.log(`[broadcast] employers active: ${res.rows.length}`);
    for (const e of res.rows) {
      if (!broadcastFieldMatches(e, fields, false)) continue;
      if (!broadcastPoliticalMatches(e, political)) continue;
      ids.add(e.telegram_id);
    }
  }
  console.log(`[broadcast] total recipients: ${ids.size}`);
  return ids;
}

async function sendBroadcast(bs) {
  const ids = await getBroadcastRecipients(bs);
  const finalMessage = bs.addWhatsapp
    ? `🔥 המלצה חמה מקוזו\n\n${bs.message}\n\nלפרטים: wa.me/972548028082`
    : `🔥 המלצה חמה מקוזו\n\n${bs.message}`;
  let sent = 0;
  for (const id of ids) {
    try {
      await bot.sendMessage(id, finalMessage);
      sent++;
    } catch (e) {
      console.error(`Broadcast failed for ${id}:`, e.message);
    }
  }
  return sent;
}

// ── דוחות תקופתיים ───────────────────────────────────────────────────────────

async function sendMonthlyReport() {
  const oneMonthAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  const [noMatchRes, activeEmpRes, successRes] = await Promise.all([
    query(
      `SELECT COUNT(DISTINCT c.telegram_id) FROM candidates c
       WHERE c.status='active' AND c.created_at < $1
       AND NOT EXISTS (SELECT 1 FROM matches m WHERE m.candidate_id = c.telegram_id)`,
      [oneMonthAgo]
    ),
    query(`SELECT COUNT(DISTINCT telegram_id) FROM employers WHERE status='active'`),
    query(
      `SELECT COUNT(*) FROM cv_requests WHERE status='connected' AND updated_at >= $1`,
      [oneMonthAgo]
    ),
  ]);
  await bot.sendMessage(
    ADMIN_ID,
    `📅 סיכום חודשי\n\n` +
    `יועצים ללא התאמה מעל חודש: ${noMatchRes.rows[0].count}\n` +
    `מגייסים פעילים: ${activeEmpRes.rows[0].count}\n` +
    `חיבורים מוצלחים החודש: ${successRes.rows[0].count}`
  );
}

function scheduleMonthlyReport() {
  const now  = new Date();
  const next = new Date(now.getFullYear(), now.getMonth() + 1, 1, 9, 0, 0, 0);
  setTimeout(async () => {
    try { await sendMonthlyReport(); } catch (e) { console.error("monthly report error:", e.message); }
    scheduleMonthlyReport();
  }, next - now);
}

async function sendWeeklySummary() {
  const oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const [newCandsRes, newEmpsRes, matchesRes, cvsRes] = await Promise.all([
    query(`SELECT COUNT(*) FROM candidates WHERE created_at >= $1`, [oneWeekAgo]),
    query(`SELECT COUNT(*) FROM employers WHERE created_at >= $1`, [oneWeekAgo]),
    query(`SELECT COUNT(*) FROM matches WHERE matched_at >= $1`, [oneWeekAgo]),
    query(`SELECT COUNT(*) FROM cv_requests WHERE requested_at >= $1`, [oneWeekAgo]),
  ]);
  await bot.sendMessage(
    ADMIN_ID,
    `📊 סיכום שבוע\n\n` +
    `יועצים חדשים שנרשמו: ${newCandsRes.rows[0].count}\n` +
    `מגייסים חדשים שנרשמו: ${newEmpsRes.rows[0].count}\n` +
    `התאמות שבוצעו: ${matchesRes.rows[0].count}\n` +
    `קורות חיים שנפתחו: ${cvsRes.rows[0].count}`
  );
}

async function sendWeeklyDigest() {
  const oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const [newCandsRes, newEmpsRes, matchesRes, connectedRes, activeCandsRes] = await Promise.all([
    query(`SELECT COUNT(*) FROM candidates WHERE created_at >= $1`, [oneWeekAgo]),
    query(`SELECT COUNT(*) FROM employers WHERE created_at >= $1`, [oneWeekAgo]),
    query(`SELECT COUNT(*) FROM matches WHERE matched_at >= $1`, [oneWeekAgo]),
    query(`SELECT COUNT(*) FROM cv_requests WHERE status='connected' AND updated_at >= $1`, [oneWeekAgo]),
    query(`SELECT telegram_id FROM candidates WHERE status='active'`),
  ]);
  const msg =
    `🔥 השבוע בקוזו\n\n` +
    `- נוספו ${newCandsRes.rows[0].count} יועצים חדשים\n` +
    `- נוספו ${newEmpsRes.rows[0].count} מגייסים חדשים\n` +
    `- בוצעו ${matchesRes.rows[0].count} התאמות\n` +
    `- ${connectedRes.rows[0].count} חיבורים מוצלחים`;
  for (const row of activeCandsRes.rows) {
    try { await bot.sendMessage(row.telegram_id, msg); } catch (_) {}
  }
}

function scheduleWeeklySummary() {
  const now = new Date();
  const day = now.getDay(); // 0 = Sunday
  const daysUntil = day === 0 && now.getHours() < 9 ? 0 : (day === 0 ? 7 : 7 - day);
  const next = new Date(now);
  next.setDate(now.getDate() + daysUntil);
  next.setHours(9, 0, 0, 0);
  setTimeout(async () => {
    try { await sendWeeklySummary(); } catch (e) { console.error("weekly summary error:", e.message); }
    try { await sendWeeklyDigest(); } catch (e) { console.error("weekly digest error:", e.message); }
    scheduleWeeklySummary();
  }, next - now);
}

// ── פקודות אדמין בטקסט (טבלה / סטטוס) ──────────────────────────────────────

async function sendStatus() {
  const oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const [activeCRes, pausedCRes, activeERes, pausedERes, archivedRes, matchesTotalRes, matchesWeekRes] = await Promise.all([
    query(`SELECT COUNT(DISTINCT telegram_id) FROM candidates WHERE status='active'`),
    query(`SELECT COUNT(DISTINCT telegram_id) FROM candidates WHERE status='paused'`),
    query(`SELECT COUNT(DISTINCT telegram_id) FROM employers  WHERE status='active'`),
    query(`SELECT COUNT(DISTINCT telegram_id) FROM employers  WHERE status='paused'`),
    query(`SELECT COUNT(DISTINCT telegram_id) FROM candidates WHERE status='archived'`),
    query(`SELECT COUNT(*) FROM matches`),
    query(`SELECT COUNT(*) FROM matches WHERE matched_at >= $1`, [oneWeekAgo]),
  ]);

  const activeCands   = Number(activeCRes.rows[0].count);
  const pausedCands   = Number(pausedCRes.rows[0].count);
  const activeEmps    = Number(activeERes.rows[0].count);
  const pausedEmps    = Number(pausedERes.rows[0].count);
  const archivedCount = Number(archivedRes.rows[0].count);
  const matchesTotal  = Number(matchesTotalRes.rows[0].count);
  const matchesWeek   = Number(matchesWeekRes.rows[0].count);

  const msg_text =
    `📊 *סטטוס קוזו*\n\n` +
    `👤 *יועצים*\n` +
    `פעילים במאגר: ${activeCands}\n` +
    `מושהים: ${pausedCands}\n\n` +
    `🏛 *מגייסים*\n` +
    `פעילים: ${activeEmps}\n` +
    `מושהים: ${pausedEmps}\n\n` +
    `🔗 *התאמות*\n` +
    `סה"כ: ${matchesTotal}\n` +
    `השבוע: ${matchesWeek}\n\n` +
    `🎉 *מצאו עבודה*: ${archivedCount}`;

  await bot.sendMessage(ADMIN_ID, msg_text, { parse_mode: "Markdown" });
}

async function sendExcel() {
  await exportExcel();
  const outPath = path.join(__dirname, "../טבלה נתונים.xlsx");
  await bot.sendDocument(ADMIN_ID, outPath, {}, { caption: "📊 טבלת נתונים מעודכנת" });
}

// ── גיבוי יומי ───────────────────────────────────────────────────────────────

async function sendDailyBackup() {
  const [candidatesRes, employersRes, cvRes, matchesRes, accessRes] = await Promise.all([
    query(`SELECT * FROM candidates ORDER BY created_at ASC`),
    query(`SELECT * FROM employers ORDER BY created_at ASC`),
    query(`SELECT * FROM cv_requests ORDER BY requested_at ASC`),
    query(`SELECT * FROM matches ORDER BY matched_at ASC`),
    query(`SELECT * FROM access_requests ORDER BY timestamp ASC`),
  ]);
  const backup = {
    exported_at: new Date().toISOString(),
    candidates:     candidatesRes.rows,
    employers:      employersRes.rows,
    cv_requests:    cvRes.rows,
    matches:        matchesRes.rows,
    access_requests: accessRes.rows,
  };
  const dateStr = new Date().toLocaleDateString("he-IL");
  const outPath = path.join(__dirname, "../backup.json");
  const { writeFileSync } = await import("fs");
  writeFileSync(outPath, JSON.stringify(backup, null, 2), "utf8");
  await bot.sendDocument(ADMIN_ID, outPath, {}, { caption: `💾 גיבוי יומי — ${dateStr}` });
}

function scheduleDailyBackup() {
  const now  = new Date();
  const next = new Date(now);
  next.setHours(6, 0, 0, 0);
  if (next <= now) next.setDate(next.getDate() + 1);
  setTimeout(async () => {
    try { await sendDailyBackup(); } catch (e) { console.error("daily backup error:", e.message); }
    try { await deactivateInactiveEmployers(); } catch (e) { console.error("deactivate employers error:", e.message); }
    scheduleDailyBackup();
  }, next - now);
}

// ── ביטול מגייסים לא פעילים ──────────────────────────────────────────────────

async function deactivateInactiveEmployers() {
  const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  const res = await query(
    `SELECT e.telegram_id, e.contact_name
     FROM employers e
     WHERE e.status='active' AND e.created_at < $1
     AND NOT EXISTS (
       SELECT 1 FROM cv_requests cr WHERE cr.employer_id = e.telegram_id
     )`,
    [thirtyDaysAgo]
  );
  if (res.rows.length === 0) return;
  for (const emp of res.rows) {
    await query(`UPDATE employers SET status='inactive' WHERE telegram_id=$1`, [emp.telegram_id]);
  }
  const names = res.rows.map((e) => e.contact_name || `ID:${e.telegram_id}`).join(", ");
  await bot.sendMessage(ADMIN_ID, `⏸ הוצאו מהמאגר (לא פעילים 30 יום): ${names}`);
}

// ── אתחול ────────────────────────────────────────────────────────────────────

(async () => {
  try {
    await initDB();
    console.log("✅ Database connected");
    const codeRes = await query(`SELECT value FROM settings WHERE key='employer_code'`);
    if (codeRes.rows.length > 0) EMPLOYER_ACCESS_CODE = codeRes.rows[0].value;
    console.log("🔑 קוד מגייסים:", EMPLOYER_ACCESS_CODE);
  } catch (err) {
    console.error("❌ Database connection failed:", err.message);
    console.log("⚠️ Bot will continue without database");
  }
  await bot.setMyCommands([
    { command: "start",   description: "כניסה / הרשמה" },
    { command: "help",    description: "עזרה ושאלות נפוצות" },
    { command: "profile", description: "הפרופיל שלי" },
    { command: "update",  description: "עדכן פרטים" },
    { command: "pause",   description: "השהה אותי" },
    { command: "resume",  description: "החזר אותי לפעילות" },
  ]);
  scheduleMonthlyReport();
  scheduleWeeklySummary();
  scheduleDailyBackup();
  await checkAnthropicKey();
  scheduleApiHealthCheck();
  await checkScheduledTasks();
  setInterval(checkScheduledTasks, 60 * 1000);
  console.log("🟢 קוזו bot פועל בטלגרם...");
})();

// ── שגיאות גלובליות ───────────────────────────────────────────────────────────

process.on("uncaughtException", async (err) => {
  console.error("uncaughtException:", err);
  try { await bot.sendMessage(ADMIN_ID, `⚠️ שגיאה כללית (uncaughtException):\n${err.message}`); } catch (_) {}
});

process.on("unhandledRejection", async (reason) => {
  console.error("unhandledRejection:", reason);
  try { await bot.sendMessage(ADMIN_ID, `⚠️ שגיאה כללית (unhandledRejection):\n${reason?.message || reason}`); } catch (_) {}
});

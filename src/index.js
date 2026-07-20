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
const EMPLOYER_ACCESS_CODE = "KOZO8"; // קוד האישור הקבוע ללשכות/עיריות

// ── ארכיון ───────────────────────────────────────────────────────────────────

async function archiveCandidate(telegramId) {
  await query(
    `UPDATE candidates SET status='archived' WHERE telegram_id=$1`,
    [telegramId]
  );
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
    `SELECT 1 FROM candidates WHERE telegram_id=$1 AND status='paused'`,
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
    `UPDATE candidates SET status='active' WHERE telegram_id=$1 AND status='paused'`,
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
    "languages", "is_intern", "internship_mentor", "internship_phone",
    "experience", "interests", "workplace_pref", "timing", "availability",
    "cv", "motivation", "has_references", "references", "declaration", "status"
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
        languages, is_intern, internship_mentor, internship_phone,
        experience, interests, workplace_pref, timing, availability,
        cv, motivation, has_references, "references", declaration
      ) VALUES (
        $1, $2,
        $3, $4, $5, $6, $7, $8,
        $9, $10, $11, $12,
        $13, $14, $15, $16, $17,
        $18, $19, $20, $21, $22
      )`,
      [
        chatId, username || "",
        data.full_name || "", data.phone || "", data.email || "",
        data.city || "", data.degree || "", data.field_of_study || "",
        data.languages || "", data.is_intern || "", data.internship_mentor || "",
        data.internship_phone || "", data.experience || "", data.interests || "",
        data.workplace_pref || "", data.timing || "", data.availability || "",
        data.cv || "", data.motivation || "", data.has_references || "",
        data.references || "", data.declaration || "",
      ]
    );
  } else {
    await query(
      `INSERT INTO employers (
        telegram_id, telegram_username,
        org_type, contact_name, phone, email,
        fields, timing, availability, experience_importance,
        notes, declaration
      ) VALUES (
        $1, $2,
        $3, $4, $5, $6,
        $7, $8, $9, $10,
        $11, $12
      )`,
      [
        chatId, username || "",
        data.org_type || "", data.contact_name || "", data.phone || "",
        data.email || "", data.fields || "", data.timing || "",
        data.availability || "", data.experience_importance || "",
        data.notes || "", data.declaration || "",
      ]
    );
  }
  console.log(`נשמר: ${type} | ${username || chatId}`);
}

function scheduleFollowUp(candidateId, employerId) {
  // שולח follow-up אחרי 7 ימים
  const delay = 7 * 24 * 60 * 60 * 1000; // 7 ימים במילישניות
  setTimeout(async () => {
    const res = await query(
      `SELECT * FROM matches WHERE candidate_id=$1 AND employer_id=$2 AND status='active'`,
      [candidateId, employerId]
    );
    const match = res.rows[0];
    if (!match) return; // כבר טופל

    await bot.sendMessage(
      ADMIN_ID,
      `📊 מעקב חיבור, שבוע עבר

` +
      `👤 מועמד: ${match.candidate_name}
` +
      `🏛 לשכה: ${match.employer_name}

` +
      `האם החיבור עדיין בתהליך?`,
      {
        reply_markup: {
          inline_keyboard: [[
            { text: "כן, בתהליך ✅", callback_data: `FOLLOWUP_YES_${candidateId}_${employerId}` },
            { text: "לא, נגמר ❌",   callback_data: `FOLLOWUP_NO_${candidateId}_${employerId}`  },
          ]],
        },
      }
    );
  }, delay);
}

// ── חיבורים ──────────────────────────────────────────────────────────────────

function workplaceMatches(candidatePref, orgType) {
  if (!candidatePref || candidatePref === "שניהם") return true;
  return candidatePref === orgType;
}

// ── חיבור אוטומטי ──────────────────────────────────────────────────────────────

async function autoConnect(candidate, employer, skipCandidateNotification = false) {
  const candidateId = candidate.telegram_id;
  const employerId  = employer.telegram_id;
  const cd = candidate;
  const ed = employer;

  // שלח ללשכה/עירייה את פרטי המועמד
  const degreeStr = [cd.degree, cd.field_of_study].filter(Boolean).join(", ");
  const rec = await getRecommendation(candidateId);
  await bot.sendMessage(
    employerId,
    `יש מישהו שנראה לי מדויק בשבילכם 👋\n\n` +
    `שם: ${cd.full_name || ""}\n` +
    `נייד: ${cd.phone || ""}\n` +
    `מייל: ${cd.email || ""}\n` +
    (degreeStr ? `תואר: ${degreeStr}\n` : "") +
    `ניסיון: ${cd.experience || ""}\n` +
    (rec ? `\n⭐ המלצה: "${rec.text}"\n` : "") +
    (cd.references ? `\n📋 ממליצים: ${cd.references}\n` : "") +
    `\nתעדכנו אותי איך יצא 🤝`,
    {
      reply_markup: {
        inline_keyboard: [[
          { text: "הפסק לקבל הצעות מתאימות מסוג זה 🔕", callback_data: `STOP_OFFERS_EMPLOYER_${employerId}` },
        ]],
      },
    }
  );

  if (cd.cv) {
    if (cd.cv.startsWith("file_id:")) {
      await bot.sendDocument(employerId, cd.cv.replace("file_id:", ""), {}, { caption: "קורות חיים" });
    } else if (cd.cv.startsWith("photo_id:")) {
      await bot.sendPhoto(employerId, cd.cv.replace("photo_id:", ""), { caption: "קורות חיים" });
    }
  }

  // שלח למועמד הודעת עדכון — רק כשמגייס חדש נרשם ומוצא אותו
  if (!skipCandidateNotification) {
    await bot.sendMessage(
      candidateId,
      `העברתי את הפרטים שלך ללשכה/גוף חדש שנרשם ונראה לי מתאים. אם יתאים — ייצרו איתך קשר 🤝`,
      {
        reply_markup: {
          inline_keyboard: [[
            { text: "הפסק לקבל הצעות 🔕", callback_data: `STOP_OFFERS_CANDIDATE_${candidateId}` },
          ]],
        },
      }
    );
  }

  // עדכון אדמין — לצפייה בלבד
  await bot.sendMessage(
    ADMIN_ID,
    `🔗 חיבור אוטומטי בוצע\n\n` +
    `👤 מועמד: ${cd.full_name || ""}\n` +
    `🏛 ${ed.org_type || "לשכה"}: ${ed.contact_name || ""}\n` +
    `תחומים: ${cd.interests || ""}`
  );

  await recordMatch(candidateId, employerId, cd.full_name || "מועמד", ed.contact_name || "לשכה");
  scheduleFollowUp(candidateId, employerId);
}

async function findMatches(employer) {
  const res = await query(`SELECT * FROM candidates WHERE status='active'`);
  const candidates = res.rows;
  const fields = (employer.fields || "").split(", ");
  const filtered = [];
  for (const c of candidates) {
    if (await hasBeenMatched(c.telegram_id, employer.telegram_id)) continue;
    if (!workplaceMatches(c.workplace_pref, employer.org_type)) continue;
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
    if (await hasBeenMatched(candidate.telegram_id, e.telegram_id)) continue;
    if (!workplaceMatches(candidate.workplace_pref, e.org_type)) continue;
    const fields = (e.fields || "").split(", ").map((f) => f.trim());
    if (fields.some((f) => interests.includes(f))) {
      filtered.push(e);
    }
  }
  return filtered;
}

// ── Excel ─────────────────────────────────────────────────────────────────────

async function exportExcel() {
  try {
    const [candidatesRes, employersRes, matchesRes, accessRes] = await Promise.all([
      query(`SELECT * FROM candidates ORDER BY created_at ASC`),
      query(`SELECT * FROM employers ORDER BY created_at ASC`),
      query(`SELECT * FROM matches ORDER BY matched_at ASC`),
      query(`SELECT * FROM access_requests ORDER BY timestamp ASC`),
    ]);
    const candidates = candidatesRes.rows;
    const employers  = employersRes.rows;
    const matchesHistory = matchesRes.rows;
    const accessRequests = accessRes.rows;

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
    XLSX.utils.book_append_sheet(wb, makeSheet("חיבורים — קוזו", MATCHES_HEADERS, matchesHistory.map(fmtM)), "חיבורים");

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
    return `מטעם: ${d.org_type}\nאיש קשר: ${d.contact_name}\nטלפון: ${d.phone}\nמייל: ${d.email}\nתחום: ${d.fields}\nמועד: ${d.timing}\nהיקף: ${d.availability}\nניסיון: ${d.experience_importance}\nדגשים: ${d.notes}`;
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
      model: "claude-sonnet-4-20250514",
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
  { key: "internship_mentor", question: "אצל מי התמחית? (שם הדובר/ת וועדה)",                                                   type: "text",   conditional: "is_intern=כן ✅" },
  { key: "internship_phone",  question: "מה מספר הנייד שלו/ה?",                                                                 type: "text",   conditional: "is_intern=כן ✅" },
  { key: "experience",        question: "נשמח לשמוע על הדרך שלך עד כה",                                    type: "text"   },
  { key: "interests",         question: "באילו תחומים יש התמחות או עניין?\nאפשר לסמן כמה ולחץ סיום ✓",                      type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "workplace_pref",    question: "איפה מעדיפים לעבוד?",                                                                type: "single", options: [["כנסת", "עירייה"], ["שניהם"]] },
  { key: "timing",            question: "מתי פנוי להתחיל?",                                                                 type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
  { key: "availability",      question: "מה היקף המשרה המבוקש?",                                                                type: "single", options: [["משרה מלאה", "משרה חלקית"], ["פרילנס", "פתוח לכל הצעה"]] },
  { key: "cv",                question: "קורות חיים 📎\nגם לא מושלמים, ניצור קשר אם יידרשו פרטים נוספים.",                  type: "file"   },
  { key: "motivation",        question: "מה מביא אותך לקוזו?\nכמה מילים מהלב",                                   type: "text"   },
  { key: "has_references",    question: "האם יש לך ממליצים שלשכות יוכלו לפנות אליהם?",                                         type: "single", options: [["כן ✅", "לא ❌"]] },
  { key: "references",        question: "ציין שם ונייד של הממליצים (אפשר כמה, מופרדים בשורות)",                         type: "text",   conditional: "has_references=כן ✅" },
  { key: "declaration",       question: "רק לידיעה. הפרטים ישמשו אותי לחיבורים בלבד. אין בזה התחייבות מאף צד 🤝", type: "single", options: [["מאשר ✅"]] },
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
  { key: "declaration",          question: "רק לידיעה. הפרטים ישמשו לחיבור מקצועי בלבד. אין בזה התחייבות מאף צד 🤝", type: "single", options: [["מאשר ✅"]] },
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
  if (session.type === "update") {
    // עדכון פרטים ומחזיר לפעילות
    await updateCandidateRecord(chatId, session.data);
    await resumeCandidate(chatId);
    await exportExcel();
    await bot.sendMessage(chatId, "מעודכן! 🤝\nחזרת לרשימה. ברגע שתהיה התאמה, אחבר.");
    await bot.sendMessage(ADMIN_ID, `🔄 מועמד חזר לפעילות (ID: ${chatId})\n${JSON.stringify(session.data, null, 2)}`);
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
    if (session.data.internship_mentor && session.data.internship_phone) {
      const mentorMsg =
        `📋 *בקשת המלצה*\n\n` +
        `המועמד ${session.data.full_name} ציין שהתמחה אצל:\n` +
        `👤 ${session.data.internship_mentor}\n` +
        `📱 ${session.data.internship_phone}\n\n` +
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
      await autoConnect(newCandidate, employer, true);
    }
    if (matchingEmployers.length > 0) {
      await bot.sendMessage(
        chatId,
        `👋 קוזו עובד בשבילך!\n\nהעברתי את הפרטים שלך ל-${matchingEmployers.length} לשכות/גופים שנראים לי מתאימים.\nאם יתאים — ייצרו איתך קשר 🤝`
      );
    } else {
      await bot.sendMessage(
        chatId,
        `👋 קוזו כאן!\nהפרופיל שלך נשמר במאגר. ברגע שתהיה התאמה, תשמעו ממני 🤝`
      );
    }
  } else {
    await bot.sendMessage(
      chatId,
      "נרשם 🤝\nיש לי אנשים שנראים לי מתאימים — אחבר ברגע שיהיה נכון\n\nקוזו"
    );
    await bot.sendMessage(ADMIN_ID, `📥 לשכה חדשה נרשמה!\n\n${formatRecord("employer", session)}`);

    // חיפוש התאמות מיידי — חיבור אוטומטי
    const newEmployer = await getEmployerRecord(chatId);
    const employerForMatch = newEmployer || { ...session.data, telegram_id: chatId };
    const matches = await findMatches(employerForMatch);
    for (const match of matches) {
      await autoConnect(match, employerForMatch);
    }
    if (matches.length > 0) {
      await bot.sendMessage(chatId, `👋 יש לי ${matches.length} אנשים שנראים לי מדויקים בשבילכם. שלחתי להם את הפרטים שלכם 🤝`);
    }
  }
  await exportExcel();
  delete sessions[chatId];
}

// ── /start ────────────────────────────────────────────────────────────────────

bot.onText(/\/start/, async (msg) => {
  const chatId = msg.chat.id;
  chatHistories[chatId] = [];

  // בדוק אם יש מועמד שמחכה להמלצה מהמשתמש הזה
  const recRes = await query(
    `SELECT * FROM candidates WHERE internship_phone IS NOT NULL AND internship_phone != '' AND status != 'archived'`
  );
  const allCandidates = recRes.rows;
  const waitingForRec = allCandidates.find(
    (c) => c.internship_phone &&
           normalizePhone(c.internship_phone) === String(chatId) &&
           true // recommendation check is done async below
  );

  // additional check: no recommendation yet
  let recCandidate = null;
  for (const c of allCandidates) {
    if (normalizePhone(c.internship_phone) === String(chatId)) {
      const existingRec = await getRecommendation(c.telegram_id);
      if (!existingRec) {
        recCandidate = c;
        break;
      }
    }
  }

  if (recCandidate) {
    sessions[chatId] = { stage: "awaiting_recommendation", candidateId: recCandidate.telegram_id, candidateName: recCandidate.full_name };
    await bot.sendMessage(
      chatId,
      `שלום! 👋\n${recCandidate.full_name} הזכיר אותך כמי שהשפיע עליו/ה.\nכמה מילים ממך יכולות לעשות הבדל. תרצה/י להמליץ?`,
      {
        reply_markup: {
          inline_keyboard: [[
            { text: "כן, אשמח להמליץ ✅", callback_data: `REC_YES_${recCandidate.telegram_id}` },
            { text: "לא תודה ❌",          callback_data: `REC_NO_${recCandidate.telegram_id}`  },
          ]],
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
          [{ text: "👤 מחפש הזדמנות מקצועית", callback_data: "CANDIDATE" }],
          [{ text: "🔎 מחפשים איש מקצוע", callback_data: "EMPLOYER" }],
        ],
      },
    }
  );
});

// ── הודעות טקסט ──────────────────────────────────────────────────────────────

bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text?.trim();
  if (!text || text.startsWith("/")) return;

  // פקודות אדמין
  if (chatId === ADMIN_ID) {
    if (text === "טבלה") { await sendExcel(); return; }
    if (text === "סטטוס") { await sendStatus(); return; }
  }

  const session = sessions[chatId];

  // ── מצב השהייה ──
  if (!session || session.stage === "free_chat") {
    // בדיקת מילות מפתח לפני Claude
    const lower = text.toLowerCase();

    if (lower.includes("מצאתי עבודה")) {
      await archiveCandidate(chatId);
      await exportExcel();
      await bot.sendMessage(chatId, "כיף לשמוע! 🎉 אם יום אחד תרצו לחזור — /start תמיד פתוח");
      await bot.sendMessage(ADMIN_ID, `📦 מועמד הועבר לארכיון (ID: ${chatId}), מצא עבודה`);
      delete sessions[chatId];
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
      if (!(await isPaused(chatId))) {
        await bot.sendMessage(chatId, "כבר ברשימה שלי");
        return;
      }
      sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
      await bot.sendMessage(chatId, "כיף שחזרת 🤝\nרק כמה עדכונים קצרים ואחזיר אותך לרשימה:");
      await sendStep(chatId, sessions[chatId]);
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
    await bot.sendMessage(ADMIN_ID, "✅ נרשם. החיבור עדיין פעיל. נבדוק שוב בשבוע הבא.");
    // שלח follow-up נוסף בעוד שבוע
    scheduleFollowUp(candidateId, employerId);
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
    await bot.sendMessage(ADMIN_ID, "❌ נרשם. החיבור נסגר. המועמד לא יוצע לאותה לשכה שוב.");
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
});

// ── פקודות אדמין בטקסט (טבלה / סטטוס) ──────────────────────────────────────

async function sendStatus() {
  const [candidatesRes, employersRes, matchesRes, archivedRes] = await Promise.all([
    query(`SELECT telegram_id, status FROM candidates`),
    query(`SELECT telegram_id, status FROM employers`),
    query(`SELECT status, matched_at FROM matches`),
    query(`SELECT COUNT(*) FROM candidates WHERE status='archived'`),
  ]);

  const candidates = candidatesRes.rows;
  const employers  = employersRes.rows;
  const matches    = matchesRes.rows;
  const archivedCount = Number(archivedRes.rows[0].count);

  const uniqueCandIds = [...new Set(candidates.map((c) => c.telegram_id))];
  const uniqueEmpIds  = [...new Set(employers.map((e) => e.telegram_id))];

  const activeCands  = uniqueCandIds.filter((id) => {
    const recs = candidates.filter((c) => c.telegram_id === id);
    return recs.some((r) => r.status === "active");
  }).length;
  const pausedCands  = uniqueCandIds.filter((id) => {
    const recs = candidates.filter((c) => c.telegram_id === id);
    return recs.every((r) => r.status === "paused");
  }).length;

  const activeEmps   = uniqueEmpIds.filter((id) => {
    const recs = employers.filter((e) => e.telegram_id === id);
    return recs.some((r) => r.status === "active");
  }).length;
  const pausedEmps   = uniqueEmpIds.filter((id) => {
    const recs = employers.filter((e) => e.telegram_id === id);
    return recs.every((r) => r.status === "paused");
  }).length;

  const oneWeekAgo  = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const matchesWeek = matches.filter((m) => new Date(m.matched_at) >= oneWeekAgo).length;

  const msg_text =
    `📊 *סטטוס קוזו*\n\n` +
    `👤 *יועצים*\n` +
    `פעילים במאגר: ${activeCands}\n` +
    `מושהים: ${pausedCands}\n\n` +
    `🏛 *מגייסים*\n` +
    `פעילים: ${activeEmps}\n` +
    `מושהים: ${pausedEmps}\n\n` +
    `🔗 *חיבורים*\n` +
    `סה"כ: ${matches.length}\n` +
    `השבוע: ${matchesWeek}\n\n` +
    `🎉 *מצאו עבודה*: ${archivedCount}`;

  await bot.sendMessage(ADMIN_ID, msg_text, { parse_mode: "Markdown" });
}

async function sendExcel() {
  await exportExcel();
  const outPath = path.join(__dirname, "../טבלה נתונים.xlsx");
  await bot.sendDocument(ADMIN_ID, outPath, {}, { caption: "📊 טבלת נתונים מעודכנת" });
}

// ── אתחול ────────────────────────────────────────────────────────────────────

(async () => {
  await initDB();
  console.log("🟢 קוזו bot פועל בטלגרם...");
})();

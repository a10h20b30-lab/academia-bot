import "dotenv/config";
import TelegramBot from "node-telegram-bot-api";
import Anthropic from "@anthropic-ai/sdk";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import * as XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const ADMIN_ID = 6021652936;
const DATA_DIR = path.join(__dirname, "../data");
const CANDIDATES_FILE      = path.join(DATA_DIR, "candidates.json");
const EMPLOYERS_FILE       = path.join(DATA_DIR, "employers.json");
const APPROVED_FILE        = path.join(DATA_DIR, "approved_phones.json");
const PENDING_MATCHES_FILE = path.join(DATA_DIR, "pending_matches.json");
const PAUSED_FILE          = path.join(DATA_DIR, "paused_candidates.json");
const MATCHES_HISTORY_FILE   = path.join(DATA_DIR, "matches_history.json");
const RECOMMENDATIONS_FILE   = path.join(DATA_DIR, "recommendations.json");
const ARCHIVE_FILE           = path.join(DATA_DIR, "archive.json");
const ACCESS_REQUESTS_FILE   = path.join(DATA_DIR, "access_requests.json");

// ── קבצים ────────────────────────────────────────────────────────────────────

function ensureDataFiles() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(CANDIDATES_FILE))      fs.writeFileSync(CANDIDATES_FILE,      "[]");
  if (!fs.existsSync(EMPLOYERS_FILE))       fs.writeFileSync(EMPLOYERS_FILE,       "[]");
  if (!fs.existsSync(APPROVED_FILE))        fs.writeFileSync(APPROVED_FILE,        "[]");
  if (!fs.existsSync(PENDING_MATCHES_FILE)) fs.writeFileSync(PENDING_MATCHES_FILE, "[]");
  if (!fs.existsSync(PAUSED_FILE))          fs.writeFileSync(PAUSED_FILE,          "[]");
  if (!fs.existsSync(MATCHES_HISTORY_FILE))   fs.writeFileSync(MATCHES_HISTORY_FILE,   "[]");
  if (!fs.existsSync(RECOMMENDATIONS_FILE))  fs.writeFileSync(RECOMMENDATIONS_FILE,  "[]");
  if (!fs.existsSync(ARCHIVE_FILE))          fs.writeFileSync(ARCHIVE_FILE,          "[]");
  if (!fs.existsSync(ACCESS_REQUESTS_FILE))  fs.writeFileSync(ACCESS_REQUESTS_FILE,  "[]");
}

function readJSON(file) {
  try { return JSON.parse(fs.readFileSync(file, "utf-8")); }
  catch { return []; }
}

function writeJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}



// ── ארכיון ───────────────────────────────────────────────────────────────────

function archiveCandidate(telegramId) {
  const candidates = readJSON(CANDIDATES_FILE);
  const candidate = [...candidates].reverse().find((c) => c.telegram_id === telegramId);
  if (!candidate) return;

  const archive = readJSON(ARCHIVE_FILE);
  archive.push({ ...candidate, archived_at: new Date().toISOString() });
  writeJSON(ARCHIVE_FILE, archive);

  // הסר מהמאגר הפעיל והשהייה
  writeJSON(CANDIDATES_FILE, candidates.filter((c) => c.telegram_id !== telegramId));
  writeJSON(PAUSED_FILE, readJSON(PAUSED_FILE).filter((p) => p.telegram_id !== telegramId));
}

// ── המלצות ───────────────────────────────────────────────────────────────────

function loadRecommendations() { return readJSON(RECOMMENDATIONS_FILE); }
function saveRecommendations(list) { writeJSON(RECOMMENDATIONS_FILE, list); }

function getRecommendation(candidateId) {
  return loadRecommendations().find((r) => r.candidateId === candidateId) || null;
}

function saveRecommendationText(candidateId, text, recommenderName) {
  const list = loadRecommendations().filter((r) => r.candidateId !== candidateId);
  list.push({ candidateId, text, recommenderName, createdAt: new Date().toISOString() });
  saveRecommendations(list);
}

// ── השהייה ────────────────────────────────────────────────────────────────────

function isPaused(telegramId) {
  return readJSON(PAUSED_FILE).some((p) => p.telegram_id === telegramId);
}

function pauseCandidate(telegramId) {
  const list = readJSON(PAUSED_FILE).filter((p) => p.telegram_id !== telegramId);
  list.push({ telegram_id: telegramId, paused_at: new Date().toISOString() });
  writeJSON(PAUSED_FILE, list);
}

function resumeCandidate(telegramId) {
  writeJSON(PAUSED_FILE, readJSON(PAUSED_FILE).filter((p) => p.telegram_id !== telegramId));
}

function getCandidateRecord(telegramId) {
  const all = readJSON(CANDIDATES_FILE);
  // מחזיר את הרשומה האחרונה של המועמד
  return [...all].reverse().find((c) => c.telegram_id === telegramId) || null;
}

function updateCandidateRecord(telegramId, updates) {
  const all = readJSON(CANDIDATES_FILE);
  // עדכון הרשומה האחרונה
  let updated = false;
  for (let i = all.length - 1; i >= 0; i--) {
    if (all[i].telegram_id === telegramId) {
      all[i].data = { ...all[i].data, ...updates };
      all[i].updated_at = new Date().toISOString();
      updated = true;
      break;
    }
  }
  if (updated) writeJSON(CANDIDATES_FILE, all);
}

// ── Pending matches ────────────────────────────────────────────────────────────

function loadPendingMatches() { return readJSON(PENDING_MATCHES_FILE); }

function savePendingMatches(list) { writeJSON(PENDING_MATCHES_FILE, list); }

function addPendingMatch(candidateId, employerId) {
  const list = loadPendingMatches().filter(
    (m) => !(m.candidateId === candidateId && m.employerId === employerId)
  );
  list.push({ candidateId, employerId, timestamp: new Date().toISOString() });
  savePendingMatches(list);
}

function removePendingMatch(candidateId, employerId) {
  savePendingMatches(
    loadPendingMatches().filter(
      (m) => !(m.candidateId === candidateId && m.employerId === employerId)
    )
  );
}

// ── טלפון ────────────────────────────────────────────────────────────────────

function normalizePhone(phone) { return phone.replace(/\D/g, ""); }
function isValidPhone(phone)   { return normalizePhone(phone).length >= 9; }
function isValidEmail(email)   { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email); }

function isApproved(phone) {
  return readJSON(APPROVED_FILE).includes(normalizePhone(phone));
}

function approvePhone(phone) {
  const list = readJSON(APPROVED_FILE);
  const normalized = normalizePhone(phone);
  if (!list.includes(normalized)) {
    list.push(normalized);
    writeJSON(APPROVED_FILE, list);
  }
}

// ── שמירת רשומות ──────────────────────────────────────────────────────────────

function saveRecord(type, chatId, username, data) {
  const file = type === "candidate" ? CANDIDATES_FILE : EMPLOYERS_FILE;
  const records = readJSON(file);
  records.push({
    timestamp: new Date().toISOString(),
    type,
    telegram_id: chatId,
    telegram_username: username || "",
    data,
  });
  writeJSON(file, records);
  console.log(`נשמר: ${type} | ${username || chatId}`);
}


// ── היסטוריית חיבורים ────────────────────────────────────────────────────────

function loadMatchesHistory() { return readJSON(MATCHES_HISTORY_FILE); }

function saveMatchesHistory(list) { writeJSON(MATCHES_HISTORY_FILE, list); }

function hasBeenMatched(candidateId, employerId) {
  return loadMatchesHistory().some(
    (m) => m.candidateId === candidateId && m.employerId === employerId
  );
}

function recordMatch(candidateId, employerId, candidateName, employerName) {
  const history = loadMatchesHistory();
  if (!hasBeenMatched(candidateId, employerId)) {
    history.push({
      candidateId,
      employerId,
      candidateName,
      employerName,
      matchedAt: new Date().toISOString(),
      status: "active",
      followUpSent: false,
    });
    saveMatchesHistory(history);
  }
}

function scheduleFollowUp(candidateId, employerId) {
  // שולח follow-up אחרי 7 ימים
  const delay = 7 * 24 * 60 * 60 * 1000; // 7 ימים במילישניות
  setTimeout(async () => {
    const history = loadMatchesHistory();
    const match = history.find(
      (m) => m.candidateId === candidateId && m.employerId === employerId && m.status === "active"
    );
    if (!match) return; // כבר טופל

    await bot.sendMessage(
      ADMIN_ID,
      `📊 מעקב חיבור — שבוע עבר

` +
      `👤 מועמד: ${match.candidateName}
` +
      `🏛 לשכה: ${match.employerName}

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

function findMatches(employer) {
  const candidates = readJSON(CANDIDATES_FILE);
  const fields = (employer.data.fields || "").split(", ");
  const paused = readJSON(PAUSED_FILE).map((p) => p.telegram_id);
  return candidates.filter((c) => {
    if (paused.includes(c.telegram_id)) return false;
    if (hasBeenMatched(c.telegram_id, employer.telegram_id)) return false;
    const interests = (c.data.interests || "").split(", ");
    return fields.some((f) => interests.some((i) => i.trim() === f.trim()));
  });
}

// מציאת לשכות מתאימות למועמד חדש
function findMatchingEmployers(candidate) {
  const employers = readJSON(EMPLOYERS_FILE);
  const interests = (candidate.data.interests || "").split(", ").map((i) => i.trim());
  return employers.filter((e) => {
    if (hasBeenMatched(candidate.telegram_id, e.telegram_id)) return false;
    const fields = (e.data.fields || "").split(", ").map((f) => f.trim());
    return fields.some((f) => interests.includes(f));
  });
}

// ── Excel ─────────────────────────────────────────────────────────────────────

function exportExcel() {
  try {
    const candidates = readJSON(CANDIDATES_FILE);
    const employers  = readJSON(EMPLOYERS_FILE);

    const CANDIDATE_HEADERS = ["תאריך","טלגרם","שם מלא","נייד","מייל","עיר","תואר","תחום לימודים","שנת סיום","שפות","ניסיון","תחומי עניין","זמינות","קורות חיים","מוטיבציה","הצהרה","סטטוס"];
    const EMPLOYER_HEADERS  = ["תאריך","טלגרם","שם ותפקיד","נייד","מייל","תחומים","היקף","תזמון","חשיבות ניסיון","הערות","הצהרה"];

    const paused = readJSON(PAUSED_FILE).map((p) => p.telegram_id);

    const fmtC = (r) => {
      const d = r.data || {};
      return {
        "תאריך": r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
        "טלגרם": r.telegram_username ? `@${r.telegram_username}` : String(r.telegram_id || ""),
        "שם מלא": d.full_name || "", "נייד": d.phone || "", "מייל": d.email || "",
        "עיר": d.city || "", "תואר": d.degree || "", "תחום לימודים": d.field_of_study || "",
        "שנת סיום": d.graduation_year || "", "שפות": d.languages || "",
        "ניסיון": d.experience || "", "תחומי עניין": d.interests || "",
        "מועד פנוי": d.timing || "", "זמינות": d.availability || "", "קורות חיים": d.cv || "",
        "מוטיבציה": d.motivation || "", "הצהרה": d.declaration || "",
        "סטטוס": paused.includes(r.telegram_id) ? "מושהה" : "פעיל",
        "הוצע ל": loadMatchesHistory()
          .filter((m) => m.candidateId === r.telegram_id)
          .map((m) => m.employerName + (m.status === "closed" ? " ✗" : " ✓"))
          .join(", ") || "—",
      };
    };
    const fmtE = (r) => {
      const d = r.data || {};
      return {
        "תאריך": r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
        "טלגרם": r.telegram_username ? `@${r.telegram_username}` : String(r.telegram_id || ""),
        "שם ותפקיד": d.contact_name || "", "נייד": d.phone || "", "מייל": d.email || "",
        "תחומים": d.fields || "", "מועד": d.timing || "", "היקף": d.availability || "",
        "חשיבות ניסיון": d.experience_importance || "", "הערות": d.notes || "",
        "הצהרה": d.declaration || "",
      };
    };

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

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, makeSheet("רשימת יועצים ודוברים — אקדמיה B", CANDIDATE_HEADERS, candidates.map(fmtC)), "יועצים ודוברים");
    XLSX.utils.book_append_sheet(wb, makeSheet("רשימת לשכות חברי כנסת — אקדמיה B", EMPLOYER_HEADERS, employers.map(fmtE)), "חברי כנסת");

    // גיליון ארכיון
    const archived = readJSON(ARCHIVE_FILE);
    XLSX.utils.book_append_sheet(wb, makeSheet("ארכיון — מצאו עבודה", CANDIDATE_HEADERS, archived.map(fmtC)), "ארכיון");

    // גיליון בקשות גישה
    const accessRequests = readJSON(ACCESS_REQUESTS_FILE);
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
    XLSX.utils.book_append_sheet(wb, makeSheet("בקשות הצטרפות — אקדמיה B", ACCESS_HEADERS, accessRequests.map(fmtA)), "בקשות הצטרפות");

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
    return `שם: ${d.full_name}\nטלפון: ${d.phone}\nמייל: ${d.email}\nעיר: ${d.city}\nתואר: ${d.degree}\nלמד: ${d.field_of_study}\nשנה: ${d.graduation_year}\nשפות: ${d.languages}\nניסיון: ${d.experience}\nתחומים: ${d.interests}\nמועד: ${d.timing}\nזמינות: ${d.availability}\nמוטיבציה: ${d.motivation}`;
  } else {
    return `איש קשר: ${d.contact_name}\nטלפון: ${d.phone}\nמייל: ${d.email}\nתחום: ${d.fields}\nמועד: ${d.timing}\nהיקף: ${d.availability}\nניסיון: ${d.experience_importance}\nדגשים: ${d.notes}`;
  }
}

// ── Claude chat ───────────────────────────────────────────────────────────────

const CHAT_SYSTEM = `אתה קוזו – נציג דיגיטלי של אקדמיה B, מערכת חיבור בין לשכות כנסת ליועצים פרלמנטריים ודוברים.

כללים:
- דבר תמיד בעברית בלבד.
- היה חם, מקצועי, תמציתי.
- ענה רק על שאלות שקשורות ל: רישום, חיבורים, השהייה/חזרה, מצב הפרופיל, מה זה אקדמיה B.
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
    return "מצטער, יש תקלה זמנית. נסה שוב עוד רגע 🙏";
  }
}

// ── שלבי טופס ────────────────────────────────────────────────────────────────

const sessions = {};

function newSession(type, username) {
  return { type, step: 0, data: {}, multiSelect: [], username: username || "", verified: true };
}

const CANDIDATE_STEPS = [
  { key: "full_name",         question: "נתחיל 🙂 מה השם?",                                                                    type: "text"   },
  { key: "email",             question: "כתובת מייל?",                                                                          type: "email"  },
  { key: "city",              question: "עיר מגורים?",                                                                          type: "text"   },
  { key: "degree",            question: "מה התואר?",                                                                            type: "single", options: [["תואר ראשון", "תואר שני"], ["אין תואר"]] },
  { key: "field_of_study",    question: "מה תחום הלימודים?",                                                                    type: "text"   },
  { key: "languages",         question: "באילו שפות יש שליטה?\nאפשר לסמן כמה ולחץ סיום ✓",                                  type: "multi",  options: [["עברית", "אנגלית"], ["ערבית", "רוסית"], ["אחר", "סיום ✓"]] },
  { key: "internship_mentor", question: "אצל מי התמחית? (שם הדובר/ת וועדה)",                                                   type: "text"   },
  { key: "internship_phone",  question: "מה מספר הנייד שלו/ה?",                                                                 type: "text"   },
  { key: "experience",        question: "ספר לנו על הניסיון המקצועי שלך – מה הדרך עד כה?",                                    type: "text"   },
  { key: "interests",         question: "באילו תחומים יש התמחות או עניין?\nאפשר לסמן כמה ולחץ סיום ✓",                      type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "timing",            question: "מתי אתה פנוי להתחיל?",                                                                 type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
  { key: "availability",      question: "מה היקף המשרה המבוקש?",                                                                type: "single", options: [["משרה מלאה", "משרה חלקית"], ["פרילנס", "פתוח לכל הצעה"]] },
  { key: "cv",                question: "קורות חיים 📎\nגם לא מושלמים – ניצור קשר אם יידרשו פרטים נוספים.",                  type: "file"   },
  { key: "motivation",        question: "למה חשוב להיות חלק מאקדמיה B?\nכמה מילים מהלב 🙂",                                   type: "text"   },
  { key: "has_references",    question: "האם יש לך ממליצים שלשכות יוכלו לפנות אליהם?",                                         type: "single", options: [["כן ✅", "לא ❌"]] },
  { key: "references",        question: "מצוין! ציין שם ונייד של הממליצים (אפשר כמה, מופרדים בשורות)",                         type: "text",   conditional: "has_references=כן ✅" },
  { key: "declaration",       question: "לידיעה –\nהטופס משמש כמאגר לצורך בחינת חיבורים אפשריים.\nאין בהגשת הפרטים משום התחייבות.", type: "single", options: [["מאשר ✅"]] },
];

const EMPLOYER_STEPS = [
  { key: "contact_name",         question: "נתחיל 🙂 שם ותפקיד בלשכה?",                                                         type: "text"   },
  { key: "email",                question: "כתובת מייל?",                                                                        type: "email"  },
  { key: "fields",               question: "מה תחום החיזוק המבוקש?\nאפשר לסמן כמה ולחץ סיום ✓",                              type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "timing",               question: "מתי נדרש מישהו?",                                                                    type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
  { key: "availability",         question: "מה היקף המשרה המבוקשת?",                                                            type: "single", options: [["משרה מלאה", "משרה חלקית"], ["פרילנס", "פתוח לכל הצעה"]] },
  { key: "experience_importance",question: "כמה חשוב ניסיון קודם בעבודה פרלמנטרית?",                                            type: "single", options: [["חובה מוחלטת", "יתרון משמעותי"], ["לא הכרחי"]] },
  { key: "notes",                question: "יש דגשים נוספים שחשוב שנדע?\nאפשר לכתוב בחופשיות, גם 'אין' זה תשובה 😊",          type: "text"   },
  { key: "declaration",          question: "לידיעה –\nהפנייה מיועדת לצורכי היכרות וחיבור מקצועי בלבד,\nואינה מהווה התחייבות מכל סוג.", type: "single", options: [["מאשר ✅"]] },
];

// שאלות עדכון למועמד שחוזר מהשהייה
const UPDATE_STEPS = [
  { key: "experience",   question: "מה הניסיון המעודכן שלך מאז הרישום האחרון?",                                               type: "text"   },
  { key: "interests",    question: "באילו תחומים אתה מעוניין כיום?\nאפשר לסמן כמה ולחץ סיום ✓",                             type: "multi",  options: [["ייעוץ פרלמנטרי", "דוברות"], ["סושיאל ורשתות חברתיות", "יועץ פוליטי"], ["עריכת וידאו", "סיום ✓"]] },
  { key: "timing",       question: "מתי אתה פנוי להתחיל?",                                                                      type: "single", options: [["מיידי", "בחודש הקרוב"], ["גמיש / פתוח"]] },
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
    updateCandidateRecord(chatId, session.data);
    resumeCandidate(chatId);
    exportExcel();
    await bot.sendMessage(chatId, "הפרטים עודכנו והפרופיל שלך חזר לפעיל! 🎉\nנעדכן אותך אם תימצא התאמה חדשה.");
    await bot.sendMessage(ADMIN_ID, `🔄 מועמד חזר לפעילות (ID: ${chatId})\n${JSON.stringify(session.data, null, 2)}`);
    delete sessions[chatId];
    return;
  }

  saveRecord(session.type, chatId, session.username, session.data);

  if (session.type === "candidate") {
    await bot.sendMessage(
      chatId,
      "תודה רבה! 🙏\nהפרטים נקלטו במערכת אקדמיה B.\nאם תימצא התאמה רלוונטית – ניצור קשר 🙂\n\nבכל עת אפשר לכתוב:\n• *השהה אותי* – להפסיק זמנית לחפש\n• *עדכן פרטים* – לרענן את הפרופיל",
      { parse_mode: "Markdown" }
    );
    await bot.sendMessage(ADMIN_ID, `📥 מועמד חדש נרשם!\n\n${formatRecord("candidate", session)}`);

    // שלח לאדמין טקסט מוכן לשליחה לדובר (אם יש)
    if (session.data.internship_mentor && session.data.internship_phone) {
      const mentorMsg =
        `📋 *בקשת המלצה*\n\n` +
        `המועמד ${session.data.full_name} ציין שהתמחה אצל:\n` +
        `👤 ${session.data.internship_mentor}\n` +
        `📱 ${session.data.internship_phone}\n\n` +
        `*טקסט מוכן לשליחה בוואטסאפ:*\n` +
        `━━━━━━━━━━━━━━━━━━\n` +
        `שלום, אני מאקדמיה B.\n` +
        `${session.data.full_name} שהתמחה אצלך ציין אותך בפרופיל שלו/ה.\n` +
        `אם תרצה/י להמליץ עליו/ה — פתח/י את הבוט כאן:\n` +
        `t.me/academiaB_advisor_bot\n` +
        `━━━━━━━━━━━━━━━━━━`;
      await bot.sendMessage(ADMIN_ID, mentorMsg, { parse_mode: "Markdown" });
    }

    // חפש לשכות קיימות שמתאימות
    const newCandidate = readJSON(CANDIDATES_FILE).find((c) => c.telegram_id === chatId);
    if (newCandidate) {
      const matchingEmployers = findMatchingEmployers(newCandidate);
      for (const employer of matchingEmployers) {
        const matchSummary =
          `🔗 התאמה פוטנציאלית למועמד חדש!\n\n` +
          `👤 מועמד: ${newCandidate.data.full_name}\nתחומים: ${newCandidate.data.interests}\nזמינות: ${newCandidate.data.availability}\n\n` +
          `🏛 לשכה: ${employer.data.contact_name}\nתחום: ${employer.data.fields}`;
        await bot.sendMessage(ADMIN_ID, matchSummary, {
          reply_markup: {
            inline_keyboard: [[
              { text: "אשר חיבור ✅", callback_data: `APPROVE_${chatId}_${employer.telegram_id}` },
              { text: "דחה ❌",       callback_data: `REJECT_${chatId}_${employer.telegram_id}`  },
            ]],
          },
        });

      }
      if (matchingEmployers.length > 0) {
        await bot.sendMessage(
          chatId,
          `👋 היי! נמצאו ${matchingEmployers.length} לשכות שעשויות להתאים לפרופיל שלך.\nאם תאושר התאמה — נחזור אליך 🙂`
        );
      }
    }
  } else {
    await bot.sendMessage(
      chatId,
      "תודה! 🙏\nהפנייה התקבלה במערכת אקדמיה B.\nאם תימצא התאמה רלוונטית – ניצור קשר בהתאם 🙂\n\nאקדמיה B"
    );
    await bot.sendMessage(ADMIN_ID, `📥 לשכה חדשה נרשמה!\n\n${formatRecord("employer", session)}`);

    // חיפוש התאמות מיידי
    const matches = findMatches({ data: session.data });
    for (const match of matches) {
      const matchSummary =
        `🔗 התאמה פוטנציאלית!\n\n` +
        `👤 מועמד: ${match.data.full_name}\nטלפון: ${match.data.phone}\nתחומים: ${match.data.interests}\nמועד: ${match.data.timing}\nזמינות: ${match.data.availability}\n\n` +
        `🏛 לשכה: ${session.data.contact_name}\nטלפון: ${session.data.phone}\nתחום: ${session.data.fields}`;
      await bot.sendMessage(ADMIN_ID, matchSummary, {
        reply_markup: {
          inline_keyboard: [[
            { text: "אשר חיבור ✅", callback_data: `APPROVE_${match.telegram_id}_${chatId}` },
            { text: "דחה ❌",       callback_data: `REJECT_${match.telegram_id}_${chatId}`  },
          ]],
        },
      });

    }
    // התראה מיידית לכל המועמדים — הודעה אחת לכל אחד
    const notifiedCandidates = new Set();
    for (const match of matches) {
      if (!notifiedCandidates.has(match.telegram_id)) {
        await bot.sendMessage(
          match.telegram_id,
          `👋 היי! לשכה חדשה נרשמה למאגר שעשויה להתאים לפרופיל שלך.\nאם תאושר התאמה — נחזור אליך 🙂`
        );
        notifiedCandidates.add(match.telegram_id);
      }
    }
  }
  exportExcel();
  delete sessions[chatId];
}

// ── /start ────────────────────────────────────────────────────────────────────

bot.onText(/\/start/, async (msg) => {
  const chatId = msg.chat.id;
  chatHistories[chatId] = [];

  // בדוק אם יש מועמד שמחכה להמלצה מהמשתמש הזה
  const candidates = readJSON(CANDIDATES_FILE);
  const waitingForRec = candidates.find(
    (c) => c.data.internship_phone &&
           normalizePhone(c.data.internship_phone) === String(chatId) &&
           !getRecommendation(c.telegram_id)
  );

  if (waitingForRec) {
    sessions[chatId] = { stage: "awaiting_recommendation", candidateId: waitingForRec.telegram_id, candidateName: waitingForRec.data.full_name };
    await bot.sendMessage(
      chatId,
      `שלום! 👋\n${waitingForRec.data.full_name} ציין אותך כמי שהדריך/ה אותו/ה.\nהאם תרצה/י לכתוב המלצה עליו/ה?`,
      {
        reply_markup: {
          inline_keyboard: [[
            { text: "כן, אשמח להמליץ ✅", callback_data: `REC_YES_${waitingForRec.telegram_id}` },
            { text: "לא תודה ❌",          callback_data: `REC_NO_${waitingForRec.telegram_id}`  },
          ]],
        },
      }
    );
    return;
  }

  sessions[chatId] = { stage: "awaiting_type", username: msg.from.username || "" };
  await bot.sendMessage(
    chatId,
    "שלום וברוכים הבאים לאקדמיה B 👋\n\nאנחנו מחברים בין לשכות כנסת ליועצים פרלמנטריים ודוברים מקצועיים.\n\nמי פונה אלינו היום?",
    {
      reply_markup: {
        inline_keyboard: [
          [{ text: "👤 יועץ / דובר – מחפש הזדמנות מקצועית", callback_data: "CANDIDATE" }],
          [{ text: "🏛 לשכת כנסת – מחפשת איש מקצוע",        callback_data: "EMPLOYER"  }],
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
      archiveCandidate(chatId);
      exportExcel();
      await bot.sendMessage(chatId, "מעולה, בהצלחה! 🎉\nהפרופיל שלך הועבר לארכיון.\nאם תרצה לחזור בעתיד — פשוט שלח /start");
      await bot.sendMessage(ADMIN_ID, `📦 מועמד הועבר לארכיון (ID: ${chatId}) — מצא עבודה`);
      delete sessions[chatId];
      return;
    }

    if (lower.includes("השהה אותי")) {
      pauseCandidate(chatId);
      exportExcel();
      await bot.sendMessage(
        chatId,
        "הפרופיל שלך הושהה 🙏\nלא תקבל התראות על חיבורים חדשים עד שתחזור.\n\nכשתרצה לחזור – פשוט כתוב *החזר אותי לפעילות*",
        { parse_mode: "Markdown" }
      );
      await bot.sendMessage(ADMIN_ID, `⏸ מועמד השהה את עצמו (ID: ${chatId})`);
      sessions[chatId] = { stage: "free_chat" };
      return;
    }

    if (lower.includes("החזר אותי לפעילות") || lower.includes("החזר אותי לפעילות")) {
      if (!isPaused(chatId)) {
        await bot.sendMessage(chatId, "הפרופיל שלך כבר פעיל 🙂");
        return;
      }
      sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
      await bot.sendMessage(chatId, "שמחים שחזרת! 🎉\nכמה שאלות קצרות לעדכון הפרופיל שלך:");
      await sendStep(chatId, sessions[chatId]);
      return;
    }

    if (lower.includes("עדכן פרטים")) {
      sessions[chatId] = { ...newSession("update", msg.from?.username || ""), stage: "updating" };
      await bot.sendMessage(chatId, "בוא נעדכן את הפרופיל שלך 🙂");
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
      { key: "job_search",  q: "באיזה תפקיד אתה מחפש?" },
      { key: "heard_from",  q: "איך שמעת על אקדמיה B?" },
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
        // שמור בקשת גישה
        const requests = readJSON(ACCESS_REQUESTS_FILE);
        requests.push({
          timestamp: new Date().toISOString(),
          telegram_id: chatId,
          telegram_username: session.username || "",
          phone: session.phone,
          full_name: a.full_name,
          role: a.role,
          job_search: a.job_search,
          heard_from: a.heard_from,
          status: "pending"
        });
        writeJSON(ACCESS_REQUESTS_FILE, requests);

        await bot.sendMessage(chatId, "תודה! 🙏\nהבקשה שלך נשלחה לצוות אקדמיה B.\nניצור איתך קשר בהקדם.");
        delete sessions[chatId];
      }
    }
    return;
  }

  // ── כתיבת המלצה ──
  if (session && session.stage === "writing_recommendation") {
    const rec = text;
    saveRecommendationText(session.candidateId, rec, "ממליץ");
    await bot.sendMessage(chatId, "תודה רבה! 🙏\nההמלצה נשמרה ותועבר ללשכות הרלוונטיות בעת חיבור.");
    await bot.sendMessage(ADMIN_ID, `⭐ התקבלה המלצה על מועמד ID: ${session.candidateId}\n\n"${rec}"`);
    delete sessions[chatId];
    return;
  }

  // ── אימות נייד (יועצים בלבד) ──
  if (session && session.stage === "awaiting_phone") {
    if (!isValidPhone(text)) {
      await bot.sendMessage(chatId, "הנייד לא נראה תקין 🙏 אנא הכנס מספר תקין.");
      return;
    }
    const phone = normalizePhone(text);

    if (isApproved(phone)) {
      sessions[chatId] = { ...newSession("candidate", session.username), phone, data: { phone } };
      await bot.sendMessage(chatId, "היי! אני קוזו 👋 העוזר של אקדמיה B.\n\nבואו נתחיל 🙂");
      await sendStep(chatId, sessions[chatId]);
    } else {
      sessions[chatId].phone = phone;
      await bot.sendMessage(
        chatId,
        "לצערנו לא ניתן להמשיך כרגע 🙏",
        {
          reply_markup: {
            inline_keyboard: [[{ text: "שלח בקשת הצטרפות", callback_data: "REQUEST_ACCESS" }]],
          },
        }
      );
    }
    return;
  }

  if (session.stage === "awaiting_type") return;
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
      await bot.sendMessage(chatId, "כתובת המייל לא נראית תקינה 🙏\nלדוגמה: name@gmail.com");
      return;
    }
    session.data[step.key] = text;
    session.step++;
    await sendStep(chatId, session);
  } else if (step.type === "file") {
    await bot.sendMessage(chatId, "יש לשלוח את הקובץ כקובץ מצורף 📎");
  } else {
    await bot.sendMessage(chatId, "יש לבחור מהאפשרויות למעלה 👆");
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

bot.on("callback_query", async (query) => {
  const chatId = query.message.chat.id;
  const data   = query.data;
  try { await bot.answerCallbackQuery(query.id); } catch (_) {}

  // בקשת הצטרפות — התחל שאלון
  if (data === "REQUEST_ACCESS") {
    sessions[chatId] = {
      stage: "access_questionnaire",
      step: 0,
      phone: sessions[chatId]?.phone || "לא ידוע",
      username: sessions[chatId]?.username || "",
      answers: {}
    };
    await bot.sendMessage(chatId, "נשמח להכיר! 🙂\nכמה שאלות קצרות לפני שנאשר את הגישה.\n\nמה שמך המלא?");
    return;
  }

  // אישור גישה
  if (data.startsWith("APPROVE_ACCESS_")) {
    const parts = data.split("_");
    const targetChatId = parts[2];
    const phone = parts[3];
    approvePhone(phone);
    // עדכן סטטוס בקשה
    const reqs = readJSON(ACCESS_REQUESTS_FILE);
    const req = reqs.find((r) => String(r.telegram_id) === String(targetChatId));
    if (req) { req.status = "approved"; req.approved_at = new Date().toISOString(); writeJSON(ACCESS_REQUESTS_FILE, reqs); }
    await bot.sendMessage(Number(targetChatId), "הבקשה אושרה! 🎉\nאפשר לשלוח /start ולהתחיל.");
    await bot.sendMessage(ADMIN_ID, `✅ אושר! נייד ${phone} נוסף לרשימה.`);
    return;
  }

  // דחיית גישה
  if (data.startsWith("DENY_ACCESS_")) {
    const targetChatId = data.split("_")[2];
    // עדכן סטטוס בקשה
    const reqsDeny = readJSON(ACCESS_REQUESTS_FILE);
    const reqDeny = reqsDeny.find((r) => String(r.telegram_id) === String(targetChatId));
    if (reqDeny) { reqDeny.status = "denied"; reqDeny.denied_at = new Date().toISOString(); writeJSON(ACCESS_REQUESTS_FILE, reqsDeny); }
    await bot.sendMessage(Number(targetChatId), "מצטערים, הבקשה לא אושרה הפעם 🙏");
    await bot.sendMessage(ADMIN_ID, "❌ הבקשה נדחתה.");
    return;
  }

  // אישור חיבור — שלב 1
  if (data.startsWith("APPROVE_") && !data.startsWith("APPROVE_ACCESS_")) {
    const parts = data.split("_");
    const candidateId = Number(parts[1]);
    const employerId  = Number(parts[2]);

    addPendingMatch(candidateId, employerId);

    await bot.sendMessage(
      candidateId,
      "היי, יש לשכה שמעוניינת בפרופיל שלך 🙂\nהאם להעביר את הפרטים?",
      {
        reply_markup: {
          inline_keyboard: [[
            { text: "כן, בטח ✅", callback_data: `CONSENT_YES_${candidateId}_${employerId}` },
            { text: "לא תודה ❌", callback_data: `CONSENT_NO_${candidateId}_${employerId}`  },
          ]],
        },
      }
    );
    await bot.sendMessage(ADMIN_ID, "⏳ בקשת הסכמה נשלחה למועמד.");
    return;
  }

  // תגובת המועמד לחיבור
  if (data.startsWith("CONSENT_YES_") || data.startsWith("CONSENT_NO_")) {
    const parts       = data.split("_");
    const approved    = parts[1] === "YES";
    const candidateId = Number(parts[2]);
    const employerId  = Number(parts[3]);

    removePendingMatch(candidateId, employerId);

    if (!approved) {
      await bot.sendMessage(employerId, "המועמד בחר שלא להעביר פרטים הפעם 🙏");
      await bot.sendMessage(chatId, "הבחירה נרשמה. תודה 🙏");
      return;
    }

    const candidates = readJSON(CANDIDATES_FILE);
    const candidate  = candidates.find((c) => c.telegram_id === candidateId);

    if (!candidate) {
      await bot.sendMessage(chatId, "לא נמצאו פרטי המועמד במערכת.");
      return;
    }

    const cd = candidate.data;
    await bot.sendMessage(
      employerId,
      `🎉 יש התאמה!\n\nנמצא מועמד מתאים דרך אקדמיה B.\n\n` +
      `שם: ${cd.full_name || ""}\n` +
      `נייד: ${cd.phone || ""}\n` +
      `מייל: ${cd.email || ""}\n` +
      `תואר: ${cd.degree || ""} — ${cd.field_of_study || ""}\n` +
      `ניסיון: ${cd.experience || ""}\n` +
      (getRecommendation(candidateId) ? `\n⭐ המלצה: "${getRecommendation(candidateId).text}"\n` : "") +
      (cd.references ? `\n📋 ממליצים: ${cd.references}\n` : "") +
      `\nבהצלחה! 🌟`
    );

    if (cd.cv) {
      if (cd.cv.startsWith("file_id:")) {
        await bot.sendDocument(employerId, cd.cv.replace("file_id:", ""), {}, { caption: "קורות חיים" });
      } else if (cd.cv.startsWith("photo_id:")) {
        await bot.sendPhoto(employerId, cd.cv.replace("photo_id:", ""), { caption: "קורות חיים" });
      }
    }

    // שלח ליועץ את פרטי הלשכה
    const employers = readJSON(EMPLOYERS_FILE);
    const employer = employers.find((e) => e.telegram_id === employerId);
    const ed = employer?.data || {};
    await bot.sendMessage(
      candidateId,
      `🎉 מעולה! הפרטים שלך הועברו ללשכה.\n\n` +
      `🏛 פרטי הלשכה:\n` +
      `איש קשר: ${ed.contact_name || ""}\n` +
      `נייד: ${ed.phone || ""}\n` +
      `מייל: ${ed.email || ""}\n\n` +
      `בהצלחה! 🌟`
    );
    await bot.sendMessage(chatId, "הפרטים הועברו. בהצלחה! 🌟");
    await bot.sendMessage(ADMIN_ID, "✅ החיבור אושר ופרטי המועמד נשלחו ללשכה.\n📅 תזכורת מעקב תישלח בעוד 7 ימים.");

    // שמור היסטוריה וקבע follow-up
    recordMatch(candidateId, employerId, cd.full_name || "מועמד", "לשכה");
    scheduleFollowUp(candidateId, employerId);
    return;
  }

  // דחיית חיבור
  if (data.startsWith("REJECT_")) {
    await bot.sendMessage(ADMIN_ID, "❌ ההתאמה נדחתה.");
    return;
  }

  // המלצה — כן
  if (data.startsWith("REC_YES_")) {
    const candidateId = Number(data.split("_")[2]);
    sessions[chatId] = { stage: "writing_recommendation", candidateId };
    await bot.sendMessage(chatId, "מצוין! 🙂\nכתוב/י את ההמלצה שלך בחופשיות — היא תועבר ללשכות בעת חיבור:");
    return;
  }

  if (data.startsWith("REC_NO_")) {
    await bot.sendMessage(chatId, "בסדר גמור, תודה! 🙏");
    delete sessions[chatId];
    return;
  }

  // follow-up אחרי שבוע
  if (data.startsWith("FOLLOWUP_YES_")) {
    const parts = data.split("_");
    const candidateId = Number(parts[2]);
    const employerId  = Number(parts[3]);
    await bot.sendMessage(ADMIN_ID, "✅ נרשם — החיבור עדיין פעיל. נבדוק שוב בשבוע הבא.");
    // שלח follow-up נוסף בעוד שבוע
    scheduleFollowUp(candidateId, employerId);
    return;
  }

  if (data.startsWith("FOLLOWUP_NO_")) {
    const parts = data.split("_");
    const candidateId = Number(parts[2]);
    const employerId  = Number(parts[3]);
    // עדכן סטטוס בהיסטוריה
    const history = loadMatchesHistory();
    const match = history.find(
      (m) => m.candidateId === candidateId && m.employerId === employerId
    );
    if (match) {
      match.status = "closed";
      match.closedAt = new Date().toISOString();
      saveMatchesHistory(history);
      exportExcel();
    }
    await bot.sendMessage(ADMIN_ID, "❌ נרשם — החיבור נסגר. המועמד לא יוצע לאותה לשכה שוב.");
    return;
  }

  // בחירת סוג פונה
  const session = sessions[chatId];
  if (!session) return;

  if (session.stage === "awaiting_type") {
    if (data === "CANDIDATE") {
      // יועץ — צריך לאמת נייד קודם
      sessions[chatId] = { stage: "awaiting_phone", username: session.username, pendingType: "candidate" };
      await bot.sendMessage(chatId, "מה מספר הנייד שלך?");
    } else if (data === "EMPLOYER") {
      // לשכה — ממשיכה ישר לשאלות
      sessions[chatId] = { ...newSession("employer", session.username), data: {} };
      await bot.sendMessage(chatId, "היי! אני קוזו 👋 העוזר של אקדמיה B.\n\nבואו נתחיל 🙂");
      await sendStep(chatId, sessions[chatId]);
    }
    return;
  }

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
        await bot.sendMessage(chatId, "יש לבחור לפחות אפשרות אחת לפני הסיום.");
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
  const candidates = readJSON(CANDIDATES_FILE);
  const employers  = readJSON(EMPLOYERS_FILE);
  const paused     = readJSON(PAUSED_FILE);
  const pending    = loadPendingMatches();
  const approved   = readJSON(APPROVED_FILE);

  const uniqueCandidateIds = [...new Set(candidates.map((c) => c.telegram_id))];
  const uniqueEmployerIds  = [...new Set(employers.map((e) => e.telegram_id))];
  const pausedIds          = paused.map((p) => p.telegram_id);
  const activeCount        = uniqueCandidateIds.filter((id) => !pausedIds.includes(id)).length;

  const lastCandidates = candidates.slice(-5).reverse().map((c) =>
    `• ${c.data.full_name || "ללא שם"} | ${c.data.interests || "—"} | ${pausedIds.includes(c.telegram_id) ? "⏸ מושהה" : "✅ פעיל"}`
  ).join("\n");

  const lastEmployers = employers.slice(-3).reverse().map((e) =>
    `• ${e.data.contact_name || "ללא שם"} | ${e.data.fields || "—"}`
  ).join("\n");

  const msg_text =
    `📊 *סטטוס אקדמיה B*\n` +
    `━━━━━━━━━━━━━━━━━━\n\n` +
    `👤 *מועמדים*\n` +
    `סה"כ רשומים: ${uniqueCandidateIds.length}\n` +
    `פעילים: ${activeCount}\n` +
    `מושהים: ${pausedIds.length}\n\n` +
    `🏛 *לשכות*\n` +
    `סה"כ רשומות: ${uniqueEmployerIds.length}\n\n` +
    `⏳ *ממתינים לאישור*: ${pending.length}\n` +
    `🔐 *מאושרי גישה*: ${approved.length}\n\n` +
    `━━━━━━━━━━━━━━━━━━\n` +
    `📋 *5 מועמדים אחרונים*:\n${lastCandidates || "אין"}\n\n` +
    `🏛 *3 לשכות אחרונות*:\n${lastEmployers || "אין"}`;

  await bot.sendMessage(ADMIN_ID, msg_text, { parse_mode: "Markdown" });
}

async function sendExcel() {
  exportExcel();
  const outPath = path.join(__dirname, "../טבלה נתונים.xlsx");
  await bot.sendDocument(ADMIN_ID, outPath, {}, { caption: "📊 טבלת נתונים מעודכנת" });
}



ensureDataFiles();
console.log("🟢 AcademiaB bot פועל בטלגרם...");

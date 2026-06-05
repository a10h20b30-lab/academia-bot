import * as XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, "data");

const candidates = JSON.parse(fs.readFileSync(path.join(DATA_DIR, "candidates.json"), "utf-8"));
const employers  = JSON.parse(fs.readFileSync(path.join(DATA_DIR, "employers.json"),  "utf-8"));

const CANDIDATE_HEADERS = ["תאריך","טלגרם","שם מלא","נייד","מייל","עיר","תואר","תחום לימודים","שנת סיום","שפות","ניסיון","תחומי עניין","זמינות","קורות חיים","מוטיבציה","הצהרה"];
const EMPLOYER_HEADERS  = ["תאריך","טלגרם","איש קשר","נייד","מייל","תחומים","היקף","תזמון","חשיבות ניסיון","הערות","הצהרה"];

function formatCandidate(r) {
  const d = r.data || {};
  return {
    "תאריך":          r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
    "טלגרם":          r.telegram_username ? `@${r.telegram_username}` : String(r.telegram_id || ""),
    "שם מלא":         d.full_name          || "",
    "נייד":           d.phone              || "",
    "מייל":           d.email              || "",
    "עיר":            d.city               || "",
    "תואר":           d.degree             || "",
    "תחום לימודים":   d.field_of_study     || "",
    "שנת סיום":       d.graduation_year    || "",
    "שפות":           d.languages          || "",
    "ניסיון":         d.experience         || "",
    "תחומי עניין":    d.interests          || "",
    "זמינות":         d.availability       || "",
    "קורות חיים":     d.cv                 || "",
    "מוטיבציה":       d.motivation         || "",
    "הצהרה":          d.declaration        || "",
  };
}

function formatEmployer(r) {
  const d = r.data || {};
  return {
    "תאריך":           r.timestamp ? new Date(r.timestamp).toLocaleString("he-IL") : "",
    "טלגרם":           r.telegram_username ? `@${r.telegram_username}` : String(r.telegram_id || ""),
    "איש קשר":         d.contact_name           || "",
    "נייד":            d.phone                  || "",
    "מייל":            d.email                  || "",
    "תחומים":          d.fields                 || "",
    "היקף":            d.scope                  || "",
    "תזמון":           d.timing                 || "",
    "חשיבות ניסיון":   d.experience_importance  || "",
    "הערות":           d.notes                  || "",
    "הצהרה":           d.declaration            || "",
  };
}

function makeSheet(title, headers, dataRows) {
  const ws = XLSX.utils.aoa_to_sheet([[title]]);
  XLSX.utils.sheet_add_aoa(ws, [[]], { origin: "A2" });

  if (dataRows.length > 0) {
    XLSX.utils.sheet_add_json(ws, dataRows, { origin: "A3", skipHeader: false });
  } else {
    XLSX.utils.sheet_add_aoa(ws, [headers], { origin: "A3" });
    XLSX.utils.sheet_add_aoa(ws, [["אין נתונים עדיין"]], { origin: "A4" });
  }

  ws["!cols"] = headers.map(() => ({ wch: 22 }));
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];
  return ws;
}

const wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb, makeSheet(
  "רשימת מחפשי עבודה — אקדמיה B",
  CANDIDATE_HEADERS,
  candidates.map(formatCandidate)
), "מחפשי עבודה");

XLSX.utils.book_append_sheet(wb, makeSheet(
  "רשימת לשכות חברי כנסת — אקדמיה B",
  EMPLOYER_HEADERS,
  employers.map(formatEmployer)
), "חברי כנסת");

const outPath = path.join(__dirname, "טבלה נתונים.xlsx");
XLSX.writeFile(wb, outPath);
console.log("✅ נוצר:", outPath);

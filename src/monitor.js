import { TelegramClient } from "teleproto";
import { StringSession } from "teleproto/sessions/index.js";
import { NewMessage } from "teleproto/events/index.js";
import { Api } from "teleproto";
import Tesseract from "tesseract.js";
import nodemailer from "nodemailer";
import { query } from "./db.js";

const GROUP_USERNAME = "marathonjobs";
const OUTREACH_DELAY_MS = 3 * 60 * 60 * 1000; // 3 שעות

const EMAIL_SUBJECT = "קוזו — פתרון חכם לגיוס אנשי מקצוע בתחום הפוליטי";
const EMAIL_BODY = `שלום,

ראינו שאתם מחפשים איש מקצוע בתחום הפוליטי.

קוזו הוא מאגר דיסקרטי של יועצים, דוברים ואנשי סושיאל — שמחבר אתכם ישירות למי שמתאים, בלי לפרסם בעשרות קבוצות.

מגדירים מה אתם צריכים, וקוזו מביא רק את הרלוונטיים.

לקבלת גישה למאגר — השיבו למייל זה או צרו קשר:

wa.me/972548028082

בברכה,
קוזו`;

async function initMonitorDB() {
  await query(`
    CREATE TABLE IF NOT EXISTS monitored_emails (
      id SERIAL PRIMARY KEY,
      email TEXT UNIQUE,
      message_text TEXT,
      found_at TIMESTAMPTZ DEFAULT NOW(),
      scheduled_for TIMESTAMPTZ,
      email_sent_at TIMESTAMPTZ
    )
  `);
}

function extractEmail(text) {
  const match = (text || "").match(/[\w.-]+@[\w.-]+\.\w+/);
  return match ? match[0].toLowerCase() : null;
}

async function ocrBuffer(buffer) {
  try {
    const { data: { text } } = await Tesseract.recognize(buffer, "heb+eng", { logger: () => {} });
    return text;
  } catch (err) {
    console.error("Monitor OCR error:", err.message);
    return "";
  }
}

function createTransporter() {
  return nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.GMAIL_USER,
      pass: process.env.GMAIL_APP_PASSWORD,
    },
  });
}

async function sendOutreachEmail(toEmail) {
  const transporter = createTransporter();
  await transporter.sendMail({
    from: `"קוזו" <${process.env.GMAIL_USER}>`,
    to: toEmail,
    subject: EMAIL_SUBJECT,
    text: EMAIL_BODY,
  });
}

async function processScheduledEmails() {
  let rows;
  try {
    const res = await query(
      `SELECT * FROM monitored_emails WHERE email_sent_at IS NULL AND scheduled_for <= NOW()`
    );
    rows = res.rows;
  } catch (err) {
    console.error("Monitor DB query error:", err.message);
    return;
  }

  for (const row of rows) {
    try {
      await sendOutreachEmail(row.email);
      await query(`UPDATE monitored_emails SET email_sent_at=NOW() WHERE id=$1`, [row.id]);
      console.log(`Monitor: email sent to ${row.email}`);
    } catch (err) {
      console.error(`Monitor: failed to send to ${row.email}:`, err.message);
    }
  }
}

async function handleMessage(client, message) {
  let text = message.message || "";

  if (message.photo || message.document) {
    try {
      const buffer = await client.downloadMedia(message, { outputFile: Buffer.alloc(0) });
      if (buffer) text += " " + (await ocrBuffer(buffer));
    } catch (err) {
      console.error("Monitor: media download error:", err.message);
    }
  }

  const email = extractEmail(text);
  if (!email) return;

  try {
    const scheduledFor = new Date(Date.now() + OUTREACH_DELAY_MS);
    await query(
      `INSERT INTO monitored_emails (email, message_text, scheduled_for)
       VALUES ($1, $2, $3)
       ON CONFLICT (email) DO NOTHING`,
      [email, text.substring(0, 2000), scheduledFor]
    );
    console.log(`Monitor: found email ${email}, scheduled for ${scheduledFor.toISOString()}`);
  } catch (err) {
    console.error("Monitor: DB insert error:", err.message);
  }
}

export async function startMonitor() {
  const apiId = parseInt(process.env.TELEGRAM_API_ID || "0");
  const apiHash = process.env.TELEGRAM_API_HASH || "";

  if (!apiId || !apiHash) {
    console.log("Monitor: TELEGRAM_API_ID/TELEGRAM_API_HASH not set — skipping group monitor");
    return;
  }

  await initMonitorDB();

  const sessionStr = process.env.TELEGRAM_SESSION || "";
  const client = new TelegramClient(new StringSession(sessionStr), apiId, apiHash, {
    connectionRetries: 5,
  });

  try {
    await client.start({
      phoneNumber: async () => {
        const phone = process.env.TELEGRAM_PHONE;
        if (!phone) throw new Error("TELEGRAM_PHONE not set");
        return phone;
      },
      password: async () => process.env.TELEGRAM_PASSWORD || "",
      phoneCode: async () => {
        throw new Error("Interactive code entry not supported on server — set TELEGRAM_SESSION in env");
      },
      onError: (err) => console.error("Monitor Telegram auth error:", err.message),
    });
  } catch (err) {
    console.error("Monitor: failed to start Telegram client:", err.message);
    return;
  }

  const savedSession = client.session.save();
  if (savedSession && savedSession !== sessionStr) {
    console.log("Monitor: new session string (save to TELEGRAM_SESSION):", savedSession);
  }

  // הצטרף לקבוצה אם לא שם
  try {
    const entity = await client.getInputEntity(GROUP_USERNAME);
    await client.invoke(new Api.channels.JoinChannel({ channel: entity }));
    console.log(`Monitor: joined @${GROUP_USERNAME}`);
  } catch (err) {
    if (!err.message?.includes("USER_ALREADY_PARTICIPANT")) {
      console.error("Monitor: join error:", err.message);
    }
  }

  client.addEventHandler(
    (event) => handleMessage(client, event.message).catch(console.error),
    new NewMessage({ chats: [GROUP_USERNAME] })
  );

  setInterval(processScheduledEmails, 60 * 1000);
  console.log(`Monitor: watching @${GROUP_USERNAME} for emails`);
}

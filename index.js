const express = require("express");
const line = require("@line/bot-sdk");
const { GoogleSpreadsheet } = require("google-spreadsheet");
const { JWT } = require("google-auth-library");
const Anthropic = require("@anthropic-ai/sdk");

const app = express();

// ── Config ──────────────────────────────────────────────────────────────────
const lineConfig = {
  channelSecret: process.env.LINE_CHANNEL_SECRET,
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
};
const client = new line.Client(lineConfig);
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Google Sheets Auth ───────────────────────────────────────────────────────
const serviceAccountAuth = new JWT({
  email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID, serviceAccountAuth);

// ── Webhook ──────────────────────────────────────────────────────────────────
app.post("/webhook", line.middleware(lineConfig), async (req, res) => {
  res.sendStatus(200);
  await Promise.all(req.body.events.map(handleEvent));
});

async function handleEvent(event) {
  if (event.type !== "message" || event.message.type !== "text") return;
  const userMsg = event.message.text.trim();
  const replyToken = event.replyToken;

  // Summary commands
  if (/summary|รายงาน|ดูรายจ่าย/i.test(userMsg)) {
    const summary = await getSummary();
    return client.replyMessage(replyToken, { type: "text", text: summary });
  }

  // Parse expense with Claude
  const parsed = await parseExpense(userMsg);
  if (!parsed) {
    return client.replyMessage(replyToken, {
      type: "text",
      text: "❓ I couldn't understand that. Try something like:\n\"coffee 85\" or \"taxi 120 baht\" or type \"summary\" to see your spending.",
    });
  }

  // Save to Google Sheets
  await saveExpense(parsed, userMsg);

  // Reply confirmation
  const emoji = categoryEmoji(parsed.category);
  const reply = `✅ Recorded!\n${emoji} ${parsed.category}\n💰 ฿${parsed.amount.toLocaleString()}\n📝 ${parsed.description}`;
  return client.replyMessage(replyToken, { type: "text", text: reply });
}

// ── Parse with Claude ────────────────────────────────────────────────────────
async function parseExpense(msg) {
  // Simple regex fallback first — handles "coffee 85", "85 coffee", "taxi 120 baht"
  const simple = parseSimple(msg);
  if (simple) return simple;

  // Try Claude for more complex messages
  try {
    const res = await anthropic.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 200,
      messages: [{
        role: "user",
        content: `Extract expense info from this message and return ONLY valid JSON, no markdown, no explanation.
Return: {"amount": number, "description": "short description", "category": "one of: Food, Transport, Shopping, Bills, Entertainment, Health, Travel, Other"}
If it's not an expense, return: {"error": "not an expense"}
Message: "${msg}"`,
      }],
    });

    const text = res.content[0].text.trim();
    const parsed = JSON.parse(text);
    if (parsed.error) return null;
    if (!parsed.amount || parsed.amount <= 0) return null;
    return parsed;
  } catch {
    return null;
  }
}

function parseSimple(msg) {
  const clean = msg.replace(/baht|บาท|฿/gi, "").trim();
  // Match "word 123" or "123 word"
  const m = clean.match(/^([a-zA-Zก-๙\s]+)\s+(\d+(?:\.\d+)?)$/) ||
            clean.match(/^(\d+(?:\.\d+)?)\s+([a-zA-Zก-๙\s]+)$/);
  if (!m) return null;

  const isFirstNum = /^\d/.test(clean);
  const amount = parseFloat(isFirstNum ? m[1] : m[2]);
  const desc = (isFirstNum ? m[2] : m[1]).trim();
  if (!amount || amount <= 0 || !desc) return null;

  return { amount, description: desc, category: guessCategory(desc) };
}

function guessCategory(desc) {
  const d = desc.toLowerCase();
  if (/coffee|cafe|food|lunch|dinner|breakfast|eat|rice|noodle|beer|drink|ข้าว|อาหาร|กาแฟ|ชา/.test(d)) return "Food";
  if (/taxi|grab|bus|bts|mrt|uber|fuel|gas|car|transport/.test(d)) return "Transport";
  if (/shop|mall|clothes|shirt|shoes|buy/.test(d)) return "Shopping";
  if (/electric|water|internet|phone|bill|rent/.test(d)) return "Bills";
  if (/movie|netflix|game|concert|entertainment/.test(d)) return "Entertainment";
  if (/doctor|hospital|medicine|health|pharmacy/.test(d)) return "Health";
  if (/hotel|flight|travel|trip/.test(d)) return "Travel";
  return "Other";
}

// ── Save to Sheet ────────────────────────────────────────────────────────────
async function saveExpense({ amount, category, description }, rawMsg) {
  await doc.loadInfo();
  const sheet = doc.sheetsByIndex[0];
  const now = new Date();
  const dateStr = now.toLocaleDateString("th-TH", {
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit",
  });
  await sheet.addRow({
    Date: dateStr,
    Amount: amount,
    Category: category,
    Description: description,
    "Raw Message": rawMsg,
  });
}

// ── Summary ──────────────────────────────────────────────────────────────────
async function getSummary() {
  try {
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    const rows = await sheet.getRows();

    if (rows.length === 0) return "📭 No expenses recorded yet!";

    const now = new Date();
    const thisMonth = now.getMonth();
    const thisYear = now.getFullYear();

    const totals = {};
    let grandTotal = 0;

    for (const row of rows) {
      const amount = parseFloat(row.get("Amount")) || 0;
      const category = row.get("Category") || "Other";
      // Filter current month (simple check by year in date string)
      totals[category] = (totals[category] || 0) + amount;
      grandTotal += amount;
    }

    let msg = `📊 *Expense Summary (All Time)*\n\n`;
    for (const [cat, total] of Object.entries(totals).sort((a, b) => b[1] - a[1])) {
      msg += `${categoryEmoji(cat)} ${cat}: ฿${total.toLocaleString()}\n`;
    }
    msg += `\n💳 Total: ฿${grandTotal.toLocaleString()}`;
    return msg;
  } catch (e) {
    return "❌ Couldn't load summary. Please try again.";
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function categoryEmoji(cat) {
  const map = {
    Food: "🍜", Transport: "🚗", Shopping: "🛍️",
    Bills: "💡", Entertainment: "🎬", Health: "💊",
    Travel: "✈️", Other: "📦",
  };
  return map[cat] || "📦";
}

// ── Start ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot running on port ${PORT}`));

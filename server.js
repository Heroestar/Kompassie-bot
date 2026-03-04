"use strict";

require("dotenv").config();

const express = require("express");
const axios = require("axios");
const nodemailer = require("nodemailer");

const app = express();
app.use(express.json({ limit: "25mb" }));

function must(name) {
  const v = process.env[name];
  if (!v) throw new Error("Mist env: " + name);
  return v;
}

const WA_TOKEN = must("WHATSAPP_TOKEN");
const PHONE_NUMBER_ID = must("PHONE_NUMBER_ID");
const VERIFY_TOKEN = must("VERIFY_TOKEN");

const SMTP_HOST = must("SMTP_HOST");
const SMTP_PORT = Number(must("SMTP_PORT"));
const SMTP_SECURE = String(process.env.SMTP_SECURE || "true") === "true";
const SMTP_USER = must("SMTP_USER");
const SMTP_PASS = must("SMTP_PASS");

const MAIL_TO_1 = must("MAIL_TO_1");
const MAIL_TO_2 = must("MAIL_TO_2");
const MAIL_TO_1_NAME = process.env.MAIL_TO_1_NAME || "Print";
const MAIL_TO_2_NAME = process.env.MAIL_TO_2_NAME || "Studio";

const PORT = Number(process.env.PORT || 3000);

const transporter = nodemailer.createTransport({
  host: SMTP_HOST,
  port: SMTP_PORT,
  secure: SMTP_SECURE,
  auth: { user: SMTP_USER, pass: SMTP_PASS },
});

const sessions = new Map();

function nowDateNL() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = d.getFullYear();
  return `${dd}-${mm}-${yy}`;
}

function cleanText(t) {
  return String(t || "").replace(/\r/g, "").trim();
}

function lower(t) {
  return cleanText(t).toLowerCase();
}

function isReset(t) {
  return lower(t) === "reset";
}

function isKlaar(t) {
  return lower(t) === "klaar";
}

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function resetSession(from) {
  sessions.set(from, {
    started: false,
    step: "collect",
    company: "",
    extraInfo: "",
    dept: "",
    photos: [],
    looseTexts: [],
    lastPhotoIndex: -1,
  });
}

function getSession(from) {
  if (!sessions.has(from)) resetSession(from);
  return sessions.get(from);
}

async function waSend(payload) {
  const url = `https://graph.facebook.com/v25.0/${PHONE_NUMBER_ID}/messages`;
  const res = await axios.post(url, payload, {
    headers: {
      Authorization: `Bearer ${WA_TOKEN}`,
      "Content-Type": "application/json",
    },
    validateStatus: () => true,
  });

  if (res.status < 200 || res.status >= 300) {
    console.log("WA SEND ERROR", res.status, JSON.stringify(res.data));
  }

  return res;
}

async function waSendText(to, bodyText) {
  return waSend({
    messaging_product: "whatsapp",
    to,
    type: "text",
    text: { body: bodyText },
  });
}

async function waSendButtons(to, bodyText, buttons) {
  return waSend({
    messaging_product: "whatsapp",
    to,
    type: "interactive",
    interactive: {
      type: "button",
      body: { text: bodyText },
      action: {
        buttons: buttons.map((b) => ({
          type: "reply",
          reply: { id: b.id, title: b.title },
        })),
      },
    },
  });
}

async function waGetMediaUrl(mediaId) {
  const url = `https://graph.facebook.com/v25.0/${mediaId}`;
  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${WA_TOKEN}` },
    validateStatus: () => true,
  });

  if (res.status < 200 || res.status >= 300) {
    throw new Error("Media meta fout: " + res.status);
  }

  if (!res.data || !res.data.url) {
    throw new Error("Media url ontbreekt");
  }

  return res.data.url;
}

async function waDownloadMedia(mediaId) {
  const mediaUrl = await waGetMediaUrl(mediaId);

  const res = await axios.get(mediaUrl, {
    headers: { Authorization: `Bearer ${WA_TOKEN}` },
    responseType: "arraybuffer",
    validateStatus: () => true,
  });

  if (res.status < 200 || res.status >= 300) {
    throw new Error("Media download fout: " + res.status);
  }

  const ctype = String(res.headers["content-type"] || "application/octet-stream");
  let ext = "bin";
  if (ctype.includes("jpeg")) ext = "jpg";
  else if (ctype.includes("png")) ext = "png";
  else if (ctype.includes("webp")) ext = "webp";

  return {
    buf: Buffer.from(res.data),
    ctype,
    ext,
  };
}

function buildMailHtml(session, from, modeName) {
  let html = "";
html = "";
html += `<b>WERKBON INTERN</b><br>`;
html += `Datum: ${escapeHtml(nowDateNL())}<br>`;
html += `Bedrijf of locatie: ${escapeHtml(session.company || "Onbekend")}<br>`;
html += `<br>`;

  session.photos.forEach((p, idx) => {
    html += `<b>FOTO ${idx + 1}</b><br>`;
    if (p.texts.length === 0) {
      html += `Geen tekst<br><br>`;
    } else {
      p.texts.forEach((t, i) => {
        html += `${i + 1}. ${escapeHtml(t)}<br>`;
      });
      html += `<br>`;
    }
  });

  if (session.looseTexts.length) {
    html += `<b>LOSSE TEKSTEN</b><br>`;
    session.looseTexts.forEach((t, i) => {
      html += `${i + 1}. ${escapeHtml(t)}<br>`;
    });
    html += `<br>`;
  }

  if (session.extraInfo) {
    html += `<b>EXTRA INFO</b><br>`;
    html += `${escapeHtml(session.extraInfo)}<br><br>`;
  }

  return html;
}

async function sendWorkMailComplete(session, from, modeName) {
  const to = modeName === MAIL_TO_1_NAME ? MAIL_TO_1 : MAIL_TO_2;

  const attachments = [];
  for (let i = 0; i < session.photos.length; i++) {
    const p = session.photos[i];
    if (!p.mediaId) continue;

    const media = await waDownloadMedia(p.mediaId);
    attachments.push({
      filename: `foto_${i + 1}.${media.ext}`,
      content: media.buf,
      contentType: media.ctype,
      contentDisposition: "attachment",
    });
  }

  const subject = `Werkbon intern ${session.company || "Onbekend"} ${modeName} ${nowDateNL()}`;

  await transporter.sendMail({
    from: SMTP_USER,
    to,
    subject,
    html: buildMailHtml(session, from, modeName),
    attachments,
  });
}

function getMsgFromWebhook(body) {
  const entry = body && body.entry && body.entry[0];
  const changes = entry && entry.changes && entry.changes[0];
  const value = changes && changes.value;
  const messages = value && value.messages;
  if (!messages || !messages.length) return null;
  return messages[0];
}

function getText(msg) {
  if (msg && msg.type === "text" && msg.text && msg.text.body) return msg.text.body;
  return "";
}

function getImage(msg) {
  if (msg && msg.type === "image" && msg.image && msg.image.id) {
    return {
      id: msg.image.id,
      caption: msg.image.caption || "",
    };
  }
  return null;
}

function getButtonId(msg) {
  if (msg && msg.type === "interactive" && msg.interactive && msg.interactive.button_reply) {
    return msg.interactive.button_reply.id || "";
  }
  return "";
}

async function sendWelcome(from) {
  await waSendText(
    from,
    "Welkom bij Kompassie!\nJe kan nu fotos en tekst insturen.\nTyp klaar als je alles hebt toegevoegd.\nTyp reset om opnieuw te beginnen."
  );
}

async function handleMessage(msg) {
  const from = msg.from;
  const session = getSession(from);

  const text = cleanText(getText(msg));
  const img = getImage(msg);
  const btnId = getButtonId(msg);

  if (text && isReset(text)) {
    resetSession(from);
    const s2 = getSession(from);
    s2.started = true;
    await sendWelcome(from);
    return;
  }

  if (!session.started) {
    session.started = true;
    await sendWelcome(from);
    return;
  }

  if (session.step === "collect") {
    if (img) {
      session.photos.push({
        mediaId: img.id,
        texts: [],
      });
      session.lastPhotoIndex = session.photos.length - 1;

      const cap = cleanText(img.caption);
      if (cap) session.photos[session.lastPhotoIndex].texts.push(cap);

      return;
    }

    if (text) {
      if (isKlaar(text)) {
        if (session.photos.length === 0 && session.looseTexts.length === 0) {
          await waSendText(from, "Ik heb nog niks ontvangen. Stuur fotos of tekst. Typ klaar als je wilt afronden.");
          return;
        }
        session.step = "ask_company";
        await waSendText(from, "Wat is het bedrijf of de locatie?");
        return;
      }

      if (session.lastPhotoIndex >= 0) {
        session.photos[session.lastPhotoIndex].texts.push(text);
      } else {
        session.looseTexts.push(text);
      }
      return;
    }

    return;
  }

  if (session.step === "ask_company") {
    if (!text) return;
    session.company = text;
    session.step = "ask_extra_buttons";
    await waSendButtons(from, "Meer informatie toevoegen?", [
      { id: "extra_ja", title: "Ja" },
      { id: "extra_nee", title: "Nee" },
    ]);
    return;
  }

  if (session.step === "ask_extra_buttons") {
    if (!btnId) return;

    if (btnId === "extra_ja") {
      session.step = "collect_extra_text";
      await waSendText(from, "Stuur de extra informatie als 1 bericht.");
      return;
    }

    if (btnId === "extra_nee") {
      session.extraInfo = "";
      session.step = "choose_dept_buttons";
      await waSendButtons(from, "Naar wie sturen?", [
        { id: "mode_studio", title: MAIL_TO_2_NAME },
        { id: "mode_print", title: MAIL_TO_1_NAME },
      ]);
      return;
    }

    return;
  }

  if (session.step === "collect_extra_text") {
    if (!text) return;
    session.extraInfo = text;
    session.step = "choose_dept_buttons";
    await waSendButtons(from, "Naar wie sturen?", [
      { id: "mode_studio", title: MAIL_TO_2_NAME },
      { id: "mode_print", title: MAIL_TO_1_NAME },
    ]);
    return;
  }

  if (session.step === "choose_dept_buttons") {
    if (!btnId) return;

    const modeName = btnId === "mode_print" ? MAIL_TO_1_NAME : btnId === "mode_studio" ? MAIL_TO_2_NAME : "";
    if (!modeName) return;

    try {
      await sendWorkMailComplete(session, from, modeName);
      await waSendText(from, "Verzonden.");
    } catch (e) {
      console.log("MAIL ERROR", e && e.message ? e.message : e);
      await waSendText(from, "Mail fout.");
    }

    resetSession(from);
    const s2 = getSession(from);
    s2.started = true;
    await sendWelcome(from);
    return;
  }
}

app.get("/api/health", (req, res) => {
  res.json({ ok: true });
});

app.get("/webhook", (req, res) => {
  const mode = req.query["hub.mode"];
  const token = req.query["hub.verify_token"];
  const challenge = req.query["hub.challenge"];

  if (mode === "subscribe" && token === VERIFY_TOKEN) {
    return res.status(200).send(String(challenge));
  }
  return res.sendStatus(403);
});

app.post("/webhook", async (req, res) => {
  try {
    const msg = getMsgFromWebhook(req.body);
    if (msg) await handleMessage(msg);
  } catch (e) {
    console.log("WEBHOOK ERROR", e && e.message ? e.message : e);
  }
  return res.sendStatus(200);
});

app.get("/", (req, res) => res.status(200).send("OK"));

app.listen(PORT, () => {
  console.log("Server draait op poort " + PORT);
  console.log("Webhook is jouw ngrok url plus /webhook");
});
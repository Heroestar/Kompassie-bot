/* eslint-disable no-console */
"use strict";

require("dotenv").config();

const express = require("express");
const axios = require("axios");
const sgMail = require("@sendgrid/mail");
const { google } = require("googleapis");

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

const MAIL_TO_1 = must("MAIL_TO_1");
const MAIL_TO_2 = must("MAIL_TO_2");
const MAIL_TO_1_NAME = process.env.MAIL_TO_1_NAME || "Print";
const MAIL_TO_2_NAME = process.env.MAIL_TO_2_NAME || "Studio";

const SENDGRID_API_KEY = must("SENDGRID_API_KEY");
const SENDGRID_FROM = must("SENDGRID_FROM");

const GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 = must("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64");
const DRIVE_PARENT_FOLDER_ID = must("DRIVE_PARENT_FOLDER_ID");

const COMPANIES = (() => {
  try {
    const raw = process.env.COMPANIES_JSON || "[]";
    const arr = JSON.parse(raw);
    if (Array.isArray(arr) && arr.length) return arr.map(String);
  } catch (e) {
    console.log("COMPANIES_JSON parse fout", e && e.message ? e.message : e);
  }
  return ["Baas_verkley"];
})();

const PORT = Number(process.env.PORT || 3000);

sgMail.setApiKey(SENDGRID_API_KEY);

const sessions = new Map();

function nowDateNL() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = d.getFullYear();
  return `${dd}-${mm}-${yy}`;
}

function nowISODate() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
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

    mode: "",
    step: "menu",

    userName: "",

    work: {
      step: "collect",
      company: "",
      extraInfo: "",
      dept: "",
      photos: [],
      looseTexts: [],
      lastPhotoIndex: -1,
    },

    admin: {
      step: "pick_company",
      company: "",
      number: "",
      category: "",
      tab: "",

      merk: "",
      model: "",
      uitvoering: "",
      kenteken: "",
      chassis: "",

      sheetId: "",
      driveRootId: "",
      driveCatId: "",
      driveObjId: "",
      driveFotosId: "",
      driveDocsId: "",

      uploadCount: 0,
      existsMode: false,
    },
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
    timeout: 25000,
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
        buttons: buttons.slice(0, 3).map((b) => ({
          type: "reply",
          reply: { id: b.id, title: b.title },
        })),
      },
    },
  });
}

async function waSendList(to, bodyText, buttonText, rows) {
  return waSend({
    messaging_product: "whatsapp",
    to,
    type: "interactive",
    interactive: {
      type: "list",
      body: { text: bodyText },
      action: {
        button: buttonText,
        sections: [
          {
            title: "Keuzes",
            rows: rows.map((r) => ({
              id: r.id,
              title: r.title,
              description: r.description || "",
            })),
          },
        ],
      },
    },
  });
}

async function waGetMediaUrl(mediaId) {
  const url = `https://graph.facebook.com/v25.0/${mediaId}`;
  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${WA_TOKEN}` },
    validateStatus: () => true,
    timeout: 25000,
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
    timeout: 45000,
  });

  if (res.status < 200 || res.status >= 300) {
    throw new Error("Media download fout: " + res.status);
  }

  const ctype = String(res.headers["content-type"] || "application/octet-stream");
  let ext = "bin";
  if (ctype.includes("jpeg")) ext = "jpg";
  else if (ctype.includes("png")) ext = "png";
  else if (ctype.includes("webp")) ext = "webp";
  else if (ctype.includes("pdf")) ext = "pdf";

  return {
    buf: Buffer.from(res.data),
    ctype,
    ext,
  };
}

function buildMailHtml(session) {
  let html = "";
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

async function sendWorkMailComplete(session, modeName) {
  const to = modeName === MAIL_TO_1_NAME ? MAIL_TO_1 : MAIL_TO_2;

  const attachments = [];
  for (let i = 0; i < session.photos.length; i++) {
    const p = session.photos[i];
    if (!p.mediaId) continue;

    const media = await waDownloadMedia(p.mediaId);
    attachments.push({
      filename: `foto_${i + 1}.${media.ext}`,
      type: media.ctype,
      disposition: "attachment",
      content: media.buf.toString("base64"),
    });
  }

  const subject = `Werkbon intern ${session.company || "Onbekend"} ${modeName} ${nowDateNL()}`;

  await sgMail.send({
    to,
    from: SENDGRID_FROM,
    subject,
    html: buildMailHtml(session),
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
      kind: "image",
      id: msg.image.id,
      caption: msg.image.caption || "",
    };
  }
  return null;
}

function getDocument(msg) {
  if (msg && msg.type === "document" && msg.document && msg.document.id) {
    return {
      kind: "document",
      id: msg.document.id,
      caption: msg.document.caption || "",
      filename: msg.document.filename || "",
    };
  }
  return null;
}

function getInteractive(msg) {
  if (!msg || msg.type !== "interactive" || !msg.interactive) return null;

  const it = msg.interactive;

  if (it.button_reply) {
    return { type: "button", id: it.button_reply.id || "", title: it.button_reply.title || "" };
  }

  if (it.list_reply) {
    return { type: "list", id: it.list_reply.id || "", title: it.list_reply.title || "" };
  }

  return null;
}

function getCategoryByNumber(n) {
  const num = Number(n);
  if (!Number.isFinite(num)) return { tab: "Overig", folder: "Overig" };

  if (num >= 1000 && num <= 1999) return { tab: "Autos", folder: "Autos" };
  if (num >= 2000 && num <= 2999) return { tab: "Pompen", folder: "Pompen" };
  if (num >= 3000 && num <= 3999) return { tab: "Kranen", folder: "Kranen" };
  if (num >= 4000 && num <= 4999) return { tab: "Keten", folder: "Keten" };
  if (num >= 5000 && num <= 5999) return { tab: "Containers", folder: "Containers" };
  if (num >= 9000 && num <= 9999) return { tab: "Overig", folder: "Overig" };

  return { tab: "Overig", folder: "Overig" };
}

function googleAuth() {
  const jsonStr = Buffer.from(GOOGLE_SERVICE_ACCOUNT_JSON_BASE64, "base64").toString("utf8");
  const creds = JSON.parse(jsonStr);
  return new google.auth.GoogleAuth({
    credentials: creds,
    scopes: [
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/spreadsheets",
    ],
  });
}

function driveClient() {
  const auth = googleAuth();
  return google.drive({ version: "v3", auth });
}

function sheetsClient() {
  const auth = googleAuth();
  return google.sheets({ version: "v4", auth });
}

async function driveFindFolder(drive, parentId, name) {
  const safeName = String(name).replace(/'/g, "\\'");
  const q = `mimeType='application/vnd.google-apps.folder' and name='${safeName}' and '${parentId}' in parents and trashed=false`;
  const r = await drive.files.list({ q, fields: "files(id,name)", pageSize: 5 });
  return (r.data.files || [])[0] || null;
}

async function driveGetOrCreateFolder(drive, parentId, name) {
  const found = await driveFindFolder(drive, parentId, name);
  if (found) return found;

  const r = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId],
    },
    fields: "id,name",
  });

  return r.data;
}

async function driveFindFirstFolderStartingWith(drive, parentId, prefix) {
  const safePrefix = String(prefix).replace(/'/g, "\\'");
  const q = `mimeType='application/vnd.google-apps.folder' and name contains '${safePrefix}' and '${parentId}' in parents and trashed=false`;
  const r = await drive.files.list({ q, fields: "files(id,name)", pageSize: 20 });
  const files = r.data.files || [];
  const pref = String(prefix);
  const exact = files.find((f) => String(f.name || "").startsWith(pref));
  return exact || files[0] || null;
}

async function driveUploadBuffer(drive, folderId, filename, buffer, contentType) {
  const r = await drive.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: { mimeType: contentType || "application/octet-stream", body: buffer },
    fields: "id,name,webViewLink",
  });
  return r.data;
}

async function sheetsFindSpreadsheetIdByName(drive, parentFolderId, name) {
  const safeName = String(name).replace(/'/g, "\\'");
  const q = `mimeType='application/vnd.google-apps.spreadsheet' and name='${safeName}' and '${parentFolderId}' in parents and trashed=false`;
  const r = await drive.files.list({ q, fields: "files(id,name)", pageSize: 5 });
  return (r.data.files || [])[0] || null;
}

async function sheetsCreateCompanySpreadsheet(company, companyFolderId) {
  const sheets = sheetsClient();
  const drive = driveClient();

  const title = `${company}_administratie`;

  const r = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title },
      sheets: [
        { properties: { title: "Autos" } },
        { properties: { title: "Pompen" } },
        { properties: { title: "Kranen" } },
        { properties: { title: "Keten" } },
        { properties: { title: "Containers" } },
        { properties: { title: "Overig" } },
      ],
    },
  });

  const spreadsheetId = r.data.spreadsheetId;

  await drive.files.update({
    fileId: spreadsheetId,
    addParents: companyFolderId,
    fields: "id,parents",
  });

  const header = [[
    "Nummer",
    "Merk",
    "Model",
    "Uitvoering",
    "Kenteken",
    "Chassisnummer",
    "Datum",
    "Medewerker",
    "Drive map link",
  ]];

  const tabs = ["Autos", "Pompen", "Kranen", "Keten", "Containers", "Overig"];

  for (const tab of tabs) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${tab}!A1:I1`,
      valueInputOption: "RAW",
      requestBody: { values: header },
    });
  }

  return spreadsheetId;
}

async function sheetsAppendRow(spreadsheetId, tab, rowValues) {
  const sheets = sheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${tab}!A:I`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [rowValues] },
  });
}

async function sheetsNumberExists(spreadsheetId, tab, numberStr) {
  const sheets = sheetsClient();
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${tab}!A2:A`,
  });

  const vals = r.data.values || [];
  const needle = String(numberStr).trim();
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i] && vals[i][0] ? String(vals[i][0]).trim() : "";
    if (v === needle) return true;
  }
  return false;
}

async function ensureAdminCompanyFolder(session) {
  const drive = driveClient();
  const company = session.admin.company;
  const rootName = `${company}_administratie`;
  const rootFolder = await driveGetOrCreateFolder(drive, DRIVE_PARENT_FOLDER_ID, rootName);
  session.admin.driveRootId = rootFolder.id;
  return { drive, rootFolder };
}

async function ensureAdminCategoryFolder(session) {
  const { drive, rootFolder } = await ensureAdminCompanyFolder(session);
  const catName = session.admin.category;
  const catFolder = await driveGetOrCreateFolder(drive, rootFolder.id, catName);
  session.admin.driveCatId = catFolder.id;
  return { drive, rootFolder, catFolder };
}

async function ensureCompanySheet(session) {
  if (session.admin.sheetId) return session.admin.sheetId;

  const { drive, rootFolder } = await ensureAdminCompanyFolder(session);
  const company = session.admin.company;
  const sheetName = `${company}_administratie`;

  const found = await sheetsFindSpreadsheetIdByName(drive, rootFolder.id, sheetName);
  if (found && found.id) {
    session.admin.sheetId = found.id;
    return found.id;
  }

  const createdId = await sheetsCreateCompanySpreadsheet(company, rootFolder.id);
  session.admin.sheetId = createdId;
  return createdId;
}

async function ensureAdminObjectFolders(session) {
  const a = session.admin;

  const { drive, catFolder } = await ensureAdminCategoryFolder(session);

  const parts = [a.number, a.merk, a.model];
  if (a.uitvoering) parts.push(a.uitvoering);
  const objName = parts.filter(Boolean).join(" ").trim();

  const objFolder = await driveGetOrCreateFolder(drive, catFolder.id, objName);
  a.driveObjId = objFolder.id;

  const fotosFolder = await driveGetOrCreateFolder(drive, objFolder.id, "Fotos");
  const docsFolder = await driveGetOrCreateFolder(drive, objFolder.id, "Documenten");
  a.driveFotosId = fotosFolder.id;
  a.driveDocsId = docsFolder.id;

  return { drive };
}

async function sendMenu(from) {
  await waSendButtons(from, "Wat wil je doen?", [
    { id: "menu_workbon", title: "Werkbon" },
    { id: "menu_admin", title: "Administratie" },
  ]);
}

async function sendWorkbonWelcome(from) {
  await waSendText(
    from,
    "Welkom bij Kompassie!\nJe kan nu fotos en tekst insturen.\nTyp klaar als je alles hebt toegevoegd.\nTyp reset om opnieuw te beginnen."
  );
}

async function ensureUserName(from, session) {
  if (session.userName) return true;
  session.step = "ask_name";
  await waSendText(from, "Wat is jouw naam?");
  return false;
}

async function handleAskName(from, session, text) {
  const name = cleanText(text);
  if (!name) {
    await waSendText(from, "Stuur jouw naam.");
    return;
  }
  session.userName = name;
  session.step = "menu";
  await sendMenu(from);
}

async function handleWorkbon(from, session, msg, text, img) {
  const work = session.work;
  const it = getInteractive(msg);

  if (work.step === "collect") {
    if (img) {
      work.photos.push({
        mediaId: img.id,
        texts: [],
      });
      work.lastPhotoIndex = work.photos.length - 1;

      const cap = cleanText(img.caption);
      if (cap) work.photos[work.lastPhotoIndex].texts.push(cap);

      return;
    }

    if (text) {
      if (isKlaar(text)) {
        if (work.photos.length === 0 && work.looseTexts.length === 0) {
          await waSendText(from, "Ik heb nog niks ontvangen. Stuur fotos of tekst. Typ klaar als je wilt afronden.");
          return;
        }
        work.step = "ask_company";
        await waSendText(from, "Wat is het bedrijf of de locatie?");
        return;
      }

      if (work.lastPhotoIndex >= 0) {
        work.photos[work.lastPhotoIndex].texts.push(text);
      } else {
        work.looseTexts.push(text);
      }
      return;
    }

    return;
  }

  if (work.step === "ask_company") {
    if (!text) return;
    work.company = text;
    work.step = "ask_extra_buttons";
    await waSendButtons(from, "Meer informatie toevoegen?", [
      { id: "extra_ja", title: "Ja" },
      { id: "extra_nee", title: "Nee" },
    ]);
    return;
  }

  if (work.step === "ask_extra_buttons") {
    if (!it || it.type !== "button") return;

    if (it.id === "extra_ja") {
      work.step = "collect_extra_text";
      await waSendText(from, "Stuur de extra informatie als 1 bericht.");
      return;
    }

    if (it.id === "extra_nee") {
      work.extraInfo = "";
      work.step = "choose_dept_buttons";
      await waSendButtons(from, "Naar wie sturen?", [
        { id: "mode_studio", title: MAIL_TO_2_NAME },
        { id: "mode_print", title: MAIL_TO_1_NAME },
      ]);
      return;
    }

    return;
  }

  if (work.step === "collect_extra_text") {
    if (!text) return;
    work.extraInfo = text;
    work.step = "choose_dept_buttons";
    await waSendButtons(from, "Naar wie sturen?", [
      { id: "mode_studio", title: MAIL_TO_2_NAME },
      { id: "mode_print", title: MAIL_TO_1_NAME },
    ]);
    return;
  }

  if (work.step === "choose_dept_buttons") {
    if (!it || it.type !== "button") return;

    const modeName =
      it.id === "mode_print" ? MAIL_TO_1_NAME :
      it.id === "mode_studio" ? MAIL_TO_2_NAME : "";

    if (!modeName) return;

    try {
      const payload = {
        company: work.company,
        extraInfo: work.extraInfo,
        photos: work.photos,
        looseTexts: work.looseTexts,
      };
      await sendWorkMailComplete(payload, modeName);
      await waSendText(from, "Verzonden.");
    } catch (e) {
      console.log("MAIL ERROR", e && e.message ? e.message : e);
      await waSendText(from, "Mail fout.");
    }

    resetSession(from);
    const s2 = getSession(from);
    s2.started = true;
    await sendMenu(from);
    return;
  }
}

async function adminAskCompany(from) {
  if (COMPANIES.length <= 3) {
    await waSendButtons(from, "Voor welk bedrijf is dit?", COMPANIES.map((c, i) => ({
      id: "cmp_" + i,
      title: c,
    })));
    return;
  }

  await waSendList(
    from,
    "Voor welk bedrijf is dit?",
    "Kies bedrijf",
    COMPANIES.slice(0, 10).map((c, i) => ({ id: "cmp_" + i, title: c }))
  );
}

async function adminAskUitvoeringPage1(from) {
  await waSendButtons(from, "Wat is de uitvoering?", [
    { id: "u_l1h1", title: "L1 H1" },
    { id: "u_l1h2", title: "L1 H2" },
    { id: "u_l2h1", title: "L2 H1" },
  ]);
}

async function adminAskUitvoeringPage2(from) {
  await waSendButtons(from, "Kies uitvoering", [
    { id: "u_l2h2", title: "L2 H2" },
    { id: "u_l3h2", title: "L3 H2" },
    { id: "u_l3h3", title: "L3 H3" },
  ]);
}

async function adminAskUitvoeringPage3(from) {
  await waSendButtons(from, "Andere uitvoering", [
    { id: "u_overig", title: "Overig" },
    { id: "u_terug", title: "Terug" },
    { id: "u_opnieuw", title: "Opnieuw" },
  ]);
}

function uitvoeringFromId(id) {
  if (id === "u_l1h1") return "L1 H1";
  if (id === "u_l1h2") return "L1 H2";
  if (id === "u_l2h1") return "L2 H1";
  if (id === "u_l2h2") return "L2 H2";
  if (id === "u_l3h2") return "L3 H2";
  if (id === "u_l3h3") return "L3 H3";
  return "";
}

async function handleAdmin(from, session, msg, text, img, doc) {
  const a = session.admin;
  const it = getInteractive(msg);

  if (a.step === "pick_company") {
    if (!it) return;

    const chosen = cleanText(it.title);
    if (!COMPANIES.includes(chosen)) {
      await waSendText(from, "Kies een bedrijf via de lijst of knoppen.");
      return;
    }

    a.company = chosen;
    a.step = "ask_number";
    await waSendText(from, "Stuur het objectnummer.");
    return;
  }

  if (a.step === "ask_number") {
    if (!text) return;
    const number = cleanText(text);

    if (!/^\d{3,6}$/.test(number)) {
      await waSendText(from, "Stuur alleen het nummer, bijvoorbeeld 1150.");
      return;
    }

    a.number = number;

    const cat = getCategoryByNumber(number);
    a.category = cat.folder;
    a.tab = cat.tab;

    const sheetId = await ensureCompanySheet(session);

    const exists = await sheetsNumberExists(sheetId, a.tab, a.number);

    if (exists) {
      a.existsMode = true;
      a.step = "number_exists";
      await waSendButtons(from, `Nummer ${a.number} bestaat al in de administratie. Wil je foto’s of documenten toevoegen?`, [
        { id: "admin_add", title: "Toevoegen" },
        { id: "admin_cancel", title: "Annuleren" },
      ]);
      return;
    }

    a.existsMode = false;
    a.step = "ask_merk";
    await waSendText(from, `Categorie ${a.category}. Wat is het merk?`);
    return;
  }

  if (a.step === "number_exists") {
    if (!it || it.type !== "button") return;

    if (it.id === "admin_cancel") {
      resetSession(from);
      const s2 = getSession(from);
      s2.started = true;
      await sendMenu(from);
      return;
    }

    if (it.id !== "admin_add") return;

    const { drive, catFolder } = await ensureAdminCategoryFolder(session);

    const found = await driveFindFirstFolderStartingWith(drive, catFolder.id, `${a.number} `);
    if (!found) {
      await waSendText(from, "Ik vond de map niet. Stuur reset en probeer opnieuw.");
      return;
    }

    a.driveObjId = found.id;

    const fotosFolder = await driveGetOrCreateFolder(drive, found.id, "Fotos");
    const docsFolder = await driveGetOrCreateFolder(drive, found.id, "Documenten");
    a.driveFotosId = fotosFolder.id;
    a.driveDocsId = docsFolder.id;

    a.step = "collect_files";
    a.uploadCount = 0;
    await waSendText(from, "Stuur foto’s of documenten. Typ klaar als alles is toegevoegd.");
    return;
  }

  if (a.step === "ask_merk") {
    if (!text) return;
    a.merk = cleanText(text);
    a.step = "ask_model";
    await waSendText(from, "Wat is het model?");
    return;
  }

  if (a.step === "ask_model") {
    if (!text) return;
    a.model = cleanText(text);

    if (a.tab === "Autos") {
      a.step = "ask_uitvoering_p1";
      await adminAskUitvoeringPage1(from);
      return;
    }

    a.uitvoering = "";
    a.step = "ask_kenteken";
    await waSendText(from, "Wat is het kenteken? Stuur leeg als dit niet geldt.");
    return;
  }

  if (a.step === "ask_uitvoering_p1") {
    if (!it || it.type !== "button") return;

    const v = uitvoeringFromId(it.id);
    if (v) {
      a.uitvoering = v;
      a.step = "ask_kenteken";
      await waSendText(from, "Wat is het kenteken?");
      return;
    }

    a.step = "ask_uitvoering_p2";
    await adminAskUitvoeringPage2(from);
    return;
  }

  if (a.step === "ask_uitvoering_p2") {
    if (!it || it.type !== "button") return;

    const v = uitvoeringFromId(it.id);
    if (v) {
      a.uitvoering = v;
      a.step = "ask_kenteken";
      await waSendText(from, "Wat is het kenteken?");
      return;
    }

    a.step = "ask_uitvoering_p3";
    await adminAskUitvoeringPage3(from);
    return;
  }

  if (a.step === "ask_uitvoering_p3") {
    if (!it || it.type !== "button") return;

    if (it.id === "u_overig") {
      a.step = "ask_uitvoering_text";
      await waSendText(from, "Typ de uitvoering, bijvoorbeeld L2 H2.");
      return;
    }

    if (it.id === "u_terug") {
      a.step = "ask_uitvoering_p2";
      await adminAskUitvoeringPage2(from);
      return;
    }

    a.step = "ask_uitvoering_p1";
    await adminAskUitvoeringPage1(from);
    return;
  }

  if (a.step === "ask_uitvoering_text") {
    if (!text) return;
    a.uitvoering = cleanText(text);
    a.step = "ask_kenteken";
    await waSendText(from, "Wat is het kenteken?");
    return;
  }

  if (a.step === "ask_kenteken") {
    a.kenteken = cleanText(text);
    a.step = "ask_chassis";
    await waSendText(from, "Wat is het chassisnummer? Stuur leeg als dit niet geldt.");
    return;
  }

  if (a.step === "ask_chassis") {
    a.chassis = cleanText(text);

    await ensureAdminObjectFolders(session);

    a.step = "collect_files";
    a.uploadCount = 0;
    await waSendText(from, "Stuur foto’s of documenten. Typ klaar als alles is toegevoegd.");
    return;
  }

  if (a.step === "collect_files") {
    if (text && isKlaar(text)) {
      const sheetId = await ensureCompanySheet(session);

      const driveLink = a.driveObjId ? `https://drive.google.com/drive/folders/${a.driveObjId}` : "";

      const row = [
        a.number,
        a.existsMode ? "" : (a.merk || ""),
        a.existsMode ? "" : (a.model || ""),
        a.existsMode ? "" : (a.uitvoering || ""),
        a.existsMode ? "" : (a.kenteken || ""),
        a.existsMode ? "" : (a.chassis || ""),
        nowISODate(),
        session.userName || "",
        driveLink,
      ];

      await sheetsAppendRow(sheetId, a.tab, row);

      await waSendText(from, "Opgeslagen.");
      resetSession(from);
      const s2 = getSession(from);
      s2.started = true;
      await sendMenu(from);
      return;
    }

    const media = img || doc;
    if (media) {
      const dl = await waDownloadMedia(media.id);

      const { drive } = await ensureAdminCategoryFolder(session);

      a.uploadCount += 1;

      if (media.kind === "document") {
        const baseName = cleanText(media.filename || `document_${a.uploadCount}.${dl.ext}`) || `document_${a.uploadCount}.${dl.ext}`;
        const filename = `${a.number}_${baseName}`;
        await driveUploadBuffer(drive, a.driveDocsId, filename, dl.buf, dl.ctype);
        return;
      }

      const filename = `${a.number}_${String(a.uploadCount).padStart(2, "0")}.${dl.ext || "jpg"}`;
      await driveUploadBuffer(drive, a.driveFotosId, filename, dl.buf, dl.ctype);
      return;
    }

    if (text) {
      await waSendText(from, "Ik verwacht hier foto’s of documenten. Typ klaar als je klaar bent.");
    }

    return;
  }
}

async function handleMessage(msg) {
  const from = msg.from;
  const session = getSession(from);

  const text = cleanText(getText(msg));
  const img = getImage(msg);
  const doc = getDocument(msg);
  const it = getInteractive(msg);

  if (text && isReset(text)) {
    resetSession(from);
    const s2 = getSession(from);
    s2.started = true;
    await sendMenu(from);
    return;
  }

  if (!session.started) {
    session.started = true;
    await sendMenu(from);
    return;
  }

  if (session.step === "ask_name") {
    await handleAskName(from, session, text);
    return;
  }

  if (!session.userName) {
    const ok = await ensureUserName(from, session);
    if (!ok) return;
  }

  if (session.step === "menu") {
    if (!it || it.type !== "button") {
      await sendMenu(from);
      return;
    }

    if (it.id === "menu_workbon") {
      session.mode = "workbon";
      session.step = "workbon";
      session.work.step = "collect";
      await sendWorkbonWelcome(from);
      return;
    }

    if (it.id === "menu_admin") {
      session.mode = "admin";
      session.step = "admin";

      session.admin = {
        step: "pick_company",
        company: "",
        number: "",
        category: "",
        tab: "",

        merk: "",
        model: "",
        uitvoering: "",
        kenteken: "",
        chassis: "",

        sheetId: "",
        driveRootId: "",
        driveCatId: "",
        driveObjId: "",
        driveFotosId: "",
        driveDocsId: "",

        uploadCount: 0,
        existsMode: false,
      };

      await adminAskCompany(from);
      return;
    }

    await sendMenu(from);
    return;
  }

  if (session.step === "workbon") {
    await handleWorkbon(from, session, msg, text, img);
    return;
  }

  if (session.step === "admin") {
    await handleAdmin(from, session, msg, text, img, doc);
    return;
  }

  session.step = "menu";
  await sendMenu(from);
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
});

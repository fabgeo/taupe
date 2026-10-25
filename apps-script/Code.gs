// Taupe : boîte à propositions de paires de mots.
// À coller dans Google Sheets > Extensions > Apps Script, puis à déployer en « Application Web ».

// ⚠️ Change ce mot de passe avant de déployer. Il n'est jamais visible dans le jeu.
const ADMIN_PASSWORD = "change-moi";

const SHEET_NAME = "Propositions";
const HEADERS = ["id", "date", "mot1", "mot2", "indice", "theme_propose", "auteur", "statut", "theme_final"];
const THEMES = ["classique", "border", "film", "geek"];
const MAX_PENDING = 300;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function rows_() {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  return values.slice(1).map((r, i) => ({
    row: i + 2, id: String(r[0]), date: r[1], a: String(r[2]), b: String(r[3]), hint: String(r[4]),
    theme: String(r[5]), author: String(r[6]), status: String(r[7]), finalTheme: String(r[8])
  }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const clean_ = (v, n) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, "").trim().slice(0, n);

// Lecture publique : uniquement les paires acceptées.
function doGet() {
  const accepted = rows_().filter(r => r.status === "acceptee" && THEMES.includes(r.finalTheme))
    .map(r => ({ a: r.a, b: r.b, hint: r.hint, theme: r.finalTheme }));
  return json_({ ok: true, accepted });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "Requête illisible." }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (body.action === "propose") {
      const a = clean_(body.a, 40), b = clean_(body.b, 40);
      if (!a || !b) return json_({ ok: false, error: "Il faut les deux mots." });
      if (a.toLowerCase() === b.toLowerCase()) return json_({ ok: false, error: "Les deux mots doivent être différents." });
      const all = rows_();
      if (all.filter(r => r.status === "attente").length >= MAX_PENDING) return json_({ ok: false, error: "La boîte à idées est pleine pour le moment." });
      const theme = THEMES.includes(body.theme) ? body.theme : "classique";
      sheet_().appendRow([Utilities.getUuid().slice(0, 8), new Date(), a, b, clean_(body.hint, 40), theme, clean_(body.author, 20), "attente", ""]);
      return json_({ ok: true });
    }

    // Tout le reste est réservé à l'admin.
    if (body.pw !== ADMIN_PASSWORD) return json_({ ok: false, error: "Mot de passe incorrect." });

    if (body.action === "list") {
      const all = rows_();
      return json_({ ok: true, pending: all.filter(r => r.status === "attente").map(({ row, ...r }) => r),
        accepted: all.filter(r => r.status === "acceptee").map(({ row, ...r }) => r) });
    }
    if (body.action === "review") {
      const r = rows_().find(x => x.id === String(body.id));
      if (!r) return json_({ ok: false, error: "Proposition introuvable." });
      const sh = sheet_();
      if (body.decision === "accept") {
        const theme = THEMES.includes(body.theme) ? body.theme : "classique";
        sh.getRange(r.row, 3, 1, 7).setValues([[clean_(body.a, 40) || r.a, clean_(body.b, 40) || r.b, clean_(body.hint, 40), r.theme, r.author, "acceptee", theme]]);
      } else if (body.decision === "reject") {
        sh.getRange(r.row, 8).setValue("refusee");
      } else if (body.decision === "remove") {
        sh.getRange(r.row, 8).setValue("retiree");
      }
      return json_({ ok: true });
    }
    return json_({ ok: false, error: "Action inconnue." });
  } finally {
    lock.releaseLock();
  }
}

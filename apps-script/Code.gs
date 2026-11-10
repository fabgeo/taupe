// Taupe : boîte à propositions de paires de mots.
// À coller dans Google Sheets > Extensions > Apps Script, puis à déployer en « Application Web ».

// Le mot de passe admin n'est PAS écrit ici : il est rangé dans les propriétés du script.
// Apps Script > ⚙️ Paramètres du projet > Propriétés du script > Ajouter : nom ADMIN_PASSWORD, valeur = ton mot de passe.
function adminPassword_() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "";
}

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

// Same word, ignoring case, accents, spaces and punctuation; a one-letter difference counts too for words of 5+ letters.
const norm_ = t => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
function close_(x, y) {
  x = norm_(x); y = norm_(y);
  if (!x || !y) return false;
  if (x === y) return true;
  if (Math.min(x.length, y.length) < 5 || Math.abs(x.length - y.length) > 1) return false;
  let i = 0, j = 0, diff = 0;
  while (i < x.length && j < y.length) {
    if (x[i] === y[j]) { i++; j++; continue; }
    if (++diff > 1) return false;
    if (x.length > y.length) i++; else if (y.length > x.length) j++; else { i++; j++; }
  }
  return diff + (x.length - i) + (y.length - j) <= 1;
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
      if (norm_(a) === norm_(b)) return json_({ ok: false, error: "Les deux mots doivent être différents." });
      const all = rows_();
      // Same word already proposed: refused. Almost the same (one letter): the player is asked to confirm.
      for (const r of all.filter(r => r.status === "attente" || r.status === "acceptee")) {
        for (const w of [a, b]) {
          const where = `${r.a} / ${r.b}${r.status === "acceptee" ? ", déjà accepté" : ", en attente de validation"}`;
          const same = [r.a, r.b].find(x => norm_(x) === norm_(w));
          if (same) return json_({ ok: false, error: `« ${w} » a déjà été proposé (${where}). Propose une autre paire.` });
          const near = [r.a, r.b].find(x => close_(w, x));
          if (near && !body.force) return json_({ ok: false, near: true, error: `« ${w} » ressemble beaucoup à « ${near} », déjà proposé (${where}). Si c'est bien un autre mot, clique sur « Envoyer quand même ».` });
        }
      }
      if (all.filter(r => r.status === "attente").length >= MAX_PENDING) return json_({ ok: false, error: "La boîte à idées est pleine pour le moment." });
      const theme = THEMES.includes(body.theme) ? body.theme : "classique";
      sheet_().appendRow([Utilities.getUuid().slice(0, 8), new Date(), a, b, clean_(body.hint, 40), theme, clean_(body.author, 20), "attente", ""]);
      return json_({ ok: true });
    }

    // Tout le reste est réservé à l'admin.
    const pw = adminPassword_();
    if (!pw) return json_({ ok: false, error: "Aucun mot de passe admin n'est configuré dans le script." });
    if (body.pw !== pw) return json_({ ok: false, error: "Mot de passe incorrect." });

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

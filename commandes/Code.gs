/**
 * Mi Housse Tech — commandes par e-mail et suivi de commande
 * ---------------------------------------------------------------
 * Ce programme se colle dans Google Sheets (Extensions > Apps Script).
 * Il :
 *   - reçoit les commandes envoyées par le site et les écrit dans l'onglet « Commandes » ;
 *   - envoie au client un e-mail avec le bouton « Je confirme ma commande » ;
 *   - vous envoie un e-mail à chaque nouvelle commande ;
 *   - répond à la page « Suivre ma commande » du site ;
 *   - envoie un e-mail au client quand vous changez le statut (Expédiée, Livrée…).
 *
 * Mise en place : voir commandes/LISEZ-MOI.txt
 */

// ===================== RÉGLAGES (modifiables) =====================
var CONFIG = {
  BOUTIQUE: 'Mi Housse Tech',
  SITE_URL: 'https://mihoussetech.com',      // adresse du site (sans / à la fin)
  TELEPHONE: '06 43 74 58 42',
  ADRESSE: 'Avenue Mohammed V – Centre commercial City Center, Kénitra',
  HORAIRES: 'Ouvert 7j/7 de 10h à 20h',
  // Livraison : laissez FRAIS à null tant que les tarifs ne sont pas fixés
  // (l'e-mail dira « frais confirmés par e-mail »). Exemple : { 'kenitra': 0, 'defaut': 35 }
  FRAIS: null,
  DELAI: 'Livraison à Kénitra. Nous vous contactons pour fixer le jour et les frais de livraison (selon le produit).',
  ANNULATION_HEURES: 48,          // commandes non confirmées annulées après N heures (0 = jamais)
  MAX_COMMANDES_PAR_HEURE: 5      // par numéro de téléphone, contre les abus
};

var STATUTS = ['En attente de confirmation', 'Confirmée', 'Préparée', 'Expédiée', 'Livrée', 'Annulée'];
var COLS = ['Date', 'Code', 'Statut', 'Prénom', 'Nom', 'Téléphone', 'E-mail', 'Réception', 'Ville', 'Adresse',
            'Articles', 'Sous-total (DH)', 'Livraison (DH)', 'Total (DH)', 'Note', 'Historique', 'Jeton', 'Langue'];
var C = {}; COLS.forEach(function (n, i) { C[n] = i + 1; });
var FEUILLE = 'Commandes';

// ===================== INSTALLATION (à lancer une fois) =====================
function installer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FEUILLE) || ss.insertSheet(FEUILLE);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#121418').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  var regle = SpreadsheetApp.newDataValidation().requireValueInList(STATUTS, true).setAllowInvalid(false).build();
  sh.getRange(2, C['Statut'], 2000, 1).setDataValidation(regle);
  sh.setColumnWidth(C['Articles'], 380); sh.setColumnWidth(C['Statut'], 190);
  sh.hideColumns(C['Historique']); sh.hideColumns(C['Jeton']); sh.hideColumns(C['Langue']);
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('quandStatutChange').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('annulerNonConfirmees').timeBased().everyHours(1).create();
  try { SpreadsheetApp.getUi().alert('Installation terminée. Déployez maintenant le programme (Déployer > Nouveau déploiement > Application Web).'); } catch (x) {}
}

// ===================== RÉCEPTION D'UNE COMMANDE (site → tableau) =====================
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents || '{}');
    if (d.site) return json({ ok: false, erreur: 'Requête refusée.' });            // champ piège anti-robot
    var c = nettoyer(d);
    if (c.erreur) return json({ ok: false, erreur: c.erreur });

    var cache = CacheService.getScriptCache(), cle = 'cmd_' + c.telNum, n = Number(cache.get(cle) || 0);
    if (n >= CONFIG.MAX_COMMANDES_PAR_HEURE) return json({ ok: false, erreur: 'Trop de commandes depuis ce numéro. Réessayez plus tard ou appelez-nous.' });
    cache.put(cle, String(n + 1), 3600);

    var articles = verifierPrix(c.articles);
    var sousTotal = articles.reduce(function (s, a) { return s + a.prix * a.q; }, 0);
    var frais = c.retrait ? 0 : fraisLivraison(c.ville);
    var total = sousTotal + (frais || 0);

    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
    var sh = feuille(), code, jeton = aleatoire(32), maintenant = new Date(), hist = [{ s: STATUTS[0], d: maintenant.toISOString() }];
    do { code = 'MHT-' + aleatoire(6); } while (trouver(sh, 'Code', code));
    sh.appendRow([maintenant, code, STATUTS[0], c.prenom, c.nom, "'" + c.tel, c.email, c.retrait ? 'Retrait en boutique' : 'Livraison',
      c.retrait ? '' : c.ville, c.retrait ? '' : c.adresse, texteArticles(articles), sousTotal, frais === null ? 'à confirmer' : frais, total,
      '', JSON.stringify(hist), jeton, c.langue]);
    } finally { lock.releaseLock(); }

    var cmd = { code: code, jeton: jeton, date: maintenant, prenom: c.prenom, nom: c.nom, tel: c.tel, email: c.email, retrait: c.retrait,
      ville: c.ville, adresse: c.adresse, articles: articles, sousTotal: sousTotal, frais: frais, total: total };
    envoyer(c.email, 'Mi Housse Tech · votre commande ' + code, emailClient(cmd));
    envoyer(Session.getEffectiveUser().getEmail(), 'Nouvelle commande ' + code + ' · ' + c.prenom + ' ' + c.nom + ' · ' + dh(total), emailBoutique(cmd));
    return json({ ok: true, code: code, jeton: jeton });
  } catch (err) {
    return json({ ok: false, erreur: 'Erreur technique. Réessayez ou appelez le ' + CONFIG.TELEPHONE + '.' });
  }
}

// ===================== CONFIRMATION (lien dans l'e-mail) et SUIVI (page du site) =====================
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'confirmer') return pageConfirmation(p.t);
  if (p.action === 'confirmer_json') return json(confirmerCommande(p.t));
  if (p.action === 'suivi') return json(suivi(p));
  return json({ ok: true, service: CONFIG.BOUTIQUE });
}

function confirmerCommande(jeton) {
  var sh = feuille(), r = jeton && trouver(sh, 'Jeton', String(jeton));
  if (!r) return { ok: false, titre: 'Lien invalide', texte: 'Cette commande est introuvable. Appelez-nous au ' + CONFIG.TELEPHONE + '.' };
  var statut = sh.getRange(r, C['Statut']).getValue(), code = sh.getRange(r, C['Code']).getValue();
  if (statut === STATUTS[0]) { changerStatut(sh, r, STATUTS[1], true); return { ok: true, code: code, titre: 'Commande confirmée', texte: 'Merci ! Votre commande ' + code + ' est confirmée. Nous la préparons et vous tenons informé par e-mail.' }; }
  if (statut === 'Annulée') return { ok: false, code: code, titre: 'Commande annulée', texte: 'La commande ' + code + ' a été annulée. Pour la reprendre, appelez-nous au ' + CONFIG.TELEPHONE + '.' };
  return { ok: true, code: code, titre: 'Déjà confirmée', texte: 'Votre commande ' + code + ' est déjà confirmée (statut : ' + statut + ').' };
}

function pageConfirmation(jeton) {
  var res = confirmerCommande(jeton), titre = res.titre, texte = res.texte;
  var lien = jeton ? CONFIG.SITE_URL + '/suivi.html?t=' + encodeURIComponent(jeton) : CONFIG.SITE_URL;
  var html = '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc(titre) + '</title></head>' +
    '<body style="margin:0;background:#f3efe8;font-family:Arial,sans-serif;color:#141414">' +
    '<div style="max-width:520px;margin:40px auto;padding:0 18px"><div style="background:#121418;color:#fff;padding:18px 22px;border-radius:18px 18px 0 0;font-weight:bold">' + esc(CONFIG.BOUTIQUE) + '</div>' +
    '<div style="background:#fff;padding:26px 22px;border-radius:0 0 18px 18px"><h1 style="font-family:Georgia,serif;font-weight:400;font-size:30px;margin:0 0 12px">' + esc(titre) + '</h1>' +
    '<p style="font-size:16px;line-height:1.55;color:#2c2925">' + esc(texte) + '</p>' +
    '<p style="margin:22px 0 0"><a href="' + esc(lien) + '" target="_top" style="display:inline-block;background:#ff6900;color:#fff;text-decoration:none;font-weight:bold;padding:13px 24px;border-radius:999px">Suivre ma commande</a></p></div></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(titre).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function suivi(p) {
  var sh = feuille(), r = null, cache = CacheService.getScriptCache();
  if (p.t) r = trouver(sh, 'Jeton', String(p.t));
  else if (p.code && p.tel) {
    var code = String(p.code).toUpperCase().replace(/\s/g, ''), cle = 'echec_' + code, n = Number(cache.get(cle) || 0);
    if (n >= 5) return { ok: false, erreur: 'Trop d’essais. Réessayez dans 15 minutes.' };
    r = trouver(sh, 'Code', code);
    if (r && chiffres(sh.getRange(r, C['Téléphone']).getValue()).slice(-9) !== chiffres(p.tel).slice(-9)) r = null;
    if (!r) { cache.put(cle, String(n + 1), 900); return { ok: false, erreur: 'Aucune commande trouvée. Vérifiez le numéro et le téléphone.' }; }
  }
  if (!r) return { ok: false, erreur: 'Aucune commande trouvée.' };
  var v = sh.getRange(r, 1, 1, COLS.length).getValues()[0], g = function (n) { return v[C[n] - 1]; };
  var tel = chiffres(g('Téléphone'));
  return { ok: true, code: g('Code'), statut: g('Statut'), date: new Date(g('Date')).toISOString(), prenom: g('Prénom'),
    tel: tel.length > 4 ? tel.slice(0, 2) + ' •• •• •• ' + tel.slice(-2) : '', reception: g('Réception'), ville: g('Ville'),
    articles: String(g('Articles')).split('\n').filter(String), sousTotal: g('Sous-total (DH)'), frais: g('Livraison (DH)'), total: g('Total (DH)'),
    note: g('Note'), historique: JSON.parse(g('Historique') || '[]'), statuts: STATUTS };
}

// ===================== CHANGEMENT DE STATUT (vous, dans le tableau) =====================
function quandStatutChange(e) {
  var rg = e && e.range; if (!rg) return;
  var sh = rg.getSheet(); if (sh.getName() !== FEUILLE || rg.getColumn() !== C['Statut'] || rg.getRow() < 2) return;
  for (var i = 0; i < rg.getNumRows(); i++) changerStatut(sh, rg.getRow() + i, sh.getRange(rg.getRow() + i, C['Statut']).getValue(), false);
}

function changerStatut(sh, r, statut, ecrire) {
  if (STATUTS.indexOf(statut) < 0) return;
  if (ecrire) sh.getRange(r, C['Statut']).setValue(statut);
  var hc = sh.getRange(r, C['Historique']), hist = JSON.parse(hc.getValue() || '[]');
  if (hist.length && hist[hist.length - 1].s === statut) return;
  hist.push({ s: statut, d: new Date().toISOString() }); hc.setValue(JSON.stringify(hist));
  var v = sh.getRange(r, 1, 1, COLS.length).getValues()[0], g = function (n) { return v[C[n] - 1]; };
  var msg = {
    'Confirmée': ['Commande ' + g('Code') + ' confirmée', 'Merci, votre commande est confirmée. Nous la préparons.'],
    'Préparée': ['Commande ' + g('Code') + ' prête', g('Réception') === 'Retrait en boutique' ? 'Votre commande est prête : vous pouvez venir la retirer au City Center.' : 'Votre commande est prête et va bientôt partir en livraison.'],
    'Expédiée': ['Commande ' + g('Code') + ' expédiée', 'Votre commande est en route. Le livreur vous appellera avant de passer. Paiement à la réception.'],
    'Livrée': ['Merci pour votre achat · ' + g('Code'), 'Votre commande est livrée. Merci de votre confiance ! Une question sur votre produit ? Notre SAV est en boutique.'],
    'Annulée': ['Commande ' + g('Code') + ' annulée', 'Votre commande a été annulée. Pour toute question, appelez-nous au ' + CONFIG.TELEPHONE + '.']
  }[statut];
  if (msg && g('E-mail')) envoyer(g('E-mail'), msg[0], emailStatut(g('Prénom'), msg[1], g('Code'), g('Jeton'), statut, g('Note')));
}

function annulerNonConfirmees() {
  if (!CONFIG.ANNULATION_HEURES) return;
  var sh = feuille(), last = sh.getLastRow(); if (last < 2) return;
  var v = sh.getRange(2, 1, last - 1, COLS.length).getValues(), limite = Date.now() - CONFIG.ANNULATION_HEURES * 3600000;
  v.forEach(function (row, i) { if (row[C['Statut'] - 1] === STATUTS[0] && new Date(row[C['Date'] - 1]).getTime() < limite) changerStatut(sh, i + 2, 'Annulée', true); });
}

// ===================== OUTILS =====================
function nettoyer(d) {
  var s = function (x, m) { return String(x == null ? '' : x).replace(/[<>]/g, '').trim().slice(0, m || 120); };
  var c = { prenom: s(d.prenom, 60), nom: s(d.nom, 60), tel: s(d.tel, 30), email: s(d.email, 120).toLowerCase(), retrait: d.mode === 'retrait',
    ville: s(d.ville, 60), adresse: s(d.adresse, 200), langue: d.langue === 'ar' ? 'ar' : 'fr' };
  c.telNum = chiffres(c.tel);
  if (!c.prenom || !c.nom) return { erreur: 'Indiquez votre prénom et votre nom.' };
  if (c.telNum.length < 9) return { erreur: 'Vérifiez votre numéro de téléphone.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c.email)) return { erreur: 'Vérifiez votre adresse e-mail.' };
  if (!c.retrait && (!c.ville || !c.adresse)) return { erreur: 'Indiquez votre ville et votre adresse.' };
  var a = Array.isArray(d.articles) ? d.articles.slice(0, 30) : [];
  c.articles = a.map(function (x) { return { nom: s(x.nom), couleur: s(x.couleur, 60), version: s(x.version, 60), q: Math.max(1, Math.min(9, parseInt(x.q, 10) || 1)), prix: Math.max(0, parseInt(x.prix, 10) || 0), photo: s(x.photo, 200) }; })
    .filter(function (x) { return x.nom; });
  if (!c.articles.length) return { erreur: 'Votre panier est vide.' };
  return c;
}

// Revérifie les prix avec le catalogue publié sur le site (on ne fait pas confiance au navigateur)
function verifierPrix(articles) {
  var cat = {};
  try {
    var cache = CacheService.getScriptCache(), mem = cache.get('catalogue');
    if (mem) cat = JSON.parse(mem);
    else {
      ['smartphones', 'tablettes', 'objets', 'televiseurs', 'ecrans', 'securite', 'purificateurs', 'electromenager', 'aspirateurs', 'mobilite'].forEach(function (r) {
        var res = UrlFetchApp.fetch(CONFIG.SITE_URL + '/data/' + r + '.json', { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) return;
        (JSON.parse(res.getContentText()).produits || []).forEach(function (p) {
          var e = { p: chiffres(p.prix) ? Number(chiffres(p.prix)) : 0, v: {} };
          (p.versions || []).forEach(function (v) { e.v[v.nom] = Number(chiffres(v.prix)) || 0; });
          cat[p.nom] = e;
        });
      });
      try { cache.put('catalogue', JSON.stringify(cat), 600); } catch (x) {}
    }
  } catch (x) { cat = {}; }
  return articles.map(function (a) {
    var e = cat[a.nom];
    if (e) { var vrai = (a.version && e.v[a.version]) || e.p; if (vrai) a.prix = vrai; }
    return a;
  });
}

function fraisLivraison(ville) {
  if (!CONFIG.FRAIS) return null;
  var k = String(ville || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]/g, '');
  return CONFIG.FRAIS.hasOwnProperty(k) ? CONFIG.FRAIS[k] : (CONFIG.FRAIS.defaut == null ? null : CONFIG.FRAIS.defaut);
}

function feuille() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(FEUILLE);
  if (!sh) { sh = ss.insertSheet(FEUILLE); sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function trouver(sh, col, val) {
  var last = sh.getLastRow(); if (last < 2) return null;
  var v = sh.getRange(2, C[col], last - 1, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === String(val)) return i + 2;
  return null;
}
function aleatoire(n) {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Math.random() + Date.now());
  while (s.length < n) { b = b.concat(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid())); for (var i = 0; i < b.length && s.length < n; i++) s += a.charAt((b[i] + 256) % a.length); }
  return s;
}
function chiffres(x) { return String(x == null ? '' : x).replace(/\D/g, ''); }
function dh(n) { return typeof n === 'number' ? String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' DH' : String(n); }
function esc(x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function texteArticles(a) { return a.map(function (x) { return x.q + ' × ' + x.nom + ([x.couleur, x.version].filter(String).length ? ' (' + [x.couleur, x.version].filter(String).join(', ') + ')' : '') + ' — ' + dh(x.prix * x.q); }).join('\n'); }
function versTexte(html) {
  return String(html)
    .replace(/<a [^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, function (m, href, t) { t = t.replace(/<[^>]+>/g, '').trim(); return t ? t + ' : ' + href : href; })
    .replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|tr|h1|h2|li)>/gi, '\n').replace(/<\/td>/gi, '  ')
    .replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}
function envoyer(a, objet, html) {
  if (!a) return;
  MailApp.sendEmail({ to: a, subject: objet, body: versTexte(html), htmlBody: html, name: CONFIG.BOUTIQUE, replyTo: Session.getEffectiveUser().getEmail() });
}

// ===================== E-MAILS =====================
function cadre(contenu) {
  return '<div style="background:#e9e5de;padding:18px 8px;font-family:Arial,Helvetica,sans-serif;color:#141414">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;background:#f3efe8;border-radius:14px;overflow:hidden">' +
    '<tr><td style="background:#121418;padding:20px 22px"><img src="' + CONFIG.SITE_URL + '/logo.png" width="38" height="38" alt="" style="border-radius:9px;vertical-align:middle"> ' +
    '<span style="color:#fff;font-weight:bold;font-size:18px;vertical-align:middle;margin-left:8px">' + esc(CONFIG.BOUTIQUE) + '</span></td></tr>' +
    '<tr><td style="padding:24px 22px">' + contenu + '</td></tr>' +
    '<tr><td style="background:#121418;color:#cfcac2;padding:18px 22px;font-size:13px;line-height:1.7"><b style="color:#fff">' + esc(CONFIG.BOUTIQUE) + '</b> · ' + esc(CONFIG.ADRESSE) +
    '<br>' + esc(CONFIG.HORAIRES) + ' · Tél. ' + esc(CONFIG.TELEPHONE) + '<br><a href="' + CONFIG.SITE_URL + '" style="color:#ff9a4d">' + CONFIG.SITE_URL.replace('https://', '') + '</a></td></tr></table></div>';
}
function bouton(lien, texte) { return '<p style="text-align:center;margin:20px 0"><a href="' + lien + '" style="display:inline-block;background:#ff6900;color:#ffffff;text-decoration:none;font-weight:bold;font-size:16px;padding:14px 28px;border-radius:999px">' + texte + '</a></p>'; }
function bloc(titre, html) { return '<div style="background:#fff;border-radius:12px;padding:16px 18px;margin:0 0 14px"><div style="font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:#7a746b;font-weight:bold;margin-bottom:10px">' + titre + '</div>' + html + '</div>'; }
function tableArticles(c) {
  var h = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:14px">';
  c.articles.forEach(function (a) {
    h += '<tr><td style="padding:8px 0;border-bottom:1px solid #f0ece5;width:64px">' + (a.photo ? '<img src="' + CONFIG.SITE_URL + '/' + esc(a.photo) + '" width="54" height="54" alt="" style="background:#f5f5f7;border-radius:8px;object-fit:contain">' : '') + '</td>' +
      '<td style="padding:8px 6px;border-bottom:1px solid #f0ece5">' + esc(a.nom) + '<br><span style="color:#7a746b;font-size:12.5px">' + esc([a.couleur, a.version, 'Qté ' + a.q].filter(String).join(' · ')) + '</span></td>' +
      '<td style="padding:8px 0;border-bottom:1px solid #f0ece5;text-align:right;font-weight:bold;white-space:nowrap">' + dh(a.prix * a.q) + '</td></tr>';
  });
  h += '<tr><td colspan="2" style="padding:10px 0 2px">Sous-total</td><td style="text-align:right;padding:10px 0 2px">' + dh(c.sousTotal) + '</td></tr>';
  h += '<tr><td colspan="2" style="padding:2px 0">' + (c.retrait ? 'Retrait en boutique' : 'Livraison') + '</td><td style="text-align:right;padding:2px 0">' + (c.retrait ? 'Gratuit' : (c.frais === null ? 'Confirmés par e-mail' : (c.frais ? dh(c.frais) : 'Gratuite'))) + '</td></tr>';
  h += '<tr><td colspan="2" style="padding:12px 0 0;border-top:1px solid #e2ddd4;font-weight:bold;font-size:16px">Total à payer ' + (c.retrait ? 'en boutique' : 'à la livraison') + '</td><td style="padding:12px 0 0;border-top:1px solid #e2ddd4;text-align:right;font-weight:bold;font-size:16px;white-space:nowrap">' + dh(c.total) + (c.frais === null && !c.retrait ? '<br><span style="font-size:12px;font-weight:normal;color:#7a746b">+ livraison</span>' : '') + '</td></tr></table>';
  return h;
}
function emailClient(c) {
  var confirmer = CONFIG.SITE_URL + '/confirmer.html?t=' + c.jeton;
  var suivre = CONFIG.SITE_URL + '/suivi.html?t=' + c.jeton;
  var liv = c.retrait ? 'Retrait en boutique au City Center, Kénitra' : esc(c.adresse) + ', ' + esc(c.ville);
  return cadre(
    '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:28px;margin:0 0 10px">Merci <span style="color:#ff6900;font-style:italic">' + esc(c.prenom) + '</span> !</h1>' +
    '<p style="font-size:15px;line-height:1.55;margin:0 0 14px">Nous avons bien reçu votre commande. Pour que nous la préparions, merci de la confirmer en cliquant sur le bouton ci-dessous.</p>' +
    '<p style="background:#fff;border:1px solid #e2ddd4;border-radius:10px;padding:10px 14px;font-size:14px;margin:0">Numéro de commande : <b style="font-size:17px">' + c.code + '</b></p>' +
    bouton(confirmer, 'Je confirme ma commande') +
    (CONFIG.ANNULATION_HEURES ? '<p style="text-align:center;font-size:12.5px;color:#7a746b;margin:-8px 0 20px">Sans confirmation sous ' + CONFIG.ANNULATION_HEURES + ' h, la commande est annulée automatiquement.</p>' : '') +
    bloc('Votre commande', tableArticles(c)) +
    bloc(c.retrait ? 'Retrait' : 'Livraison', '<div style="font-size:14px;line-height:1.8"><span style="color:#7a746b">Nom :</span> ' + esc(c.prenom + ' ' + c.nom) + '<br><span style="color:#7a746b">Téléphone :</span> ' + esc(c.tel) +
      '<br><span style="color:#7a746b">' + (c.retrait ? 'Lieu' : 'Adresse') + ' :</span> ' + liv + (c.retrait ? '' : '<br><span style="color:#7a746b">Délai :</span> ' + esc(CONFIG.DELAI)) + '<br><span style="color:#7a746b">Paiement :</span> en espèces ' + (c.retrait ? 'en boutique' : 'à la livraison') + '</div>') +
    '<p style="text-align:center;margin:18px 0 6px"><a href="' + suivre + '" style="color:#141414;font-weight:bold">Suivre ma commande</a></p>' +
    '<p style="font-size:12.5px;color:#7a746b;text-align:center;margin:14px 0 0">Vous n’avez pas passé cette commande ? Ignorez simplement cet e-mail.</p>');
}
function emailBoutique(c) {
  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return cadre('<h1 style="font-family:Georgia,serif;font-weight:400;font-size:26px;margin:0 0 10px">Nouvelle commande ' + c.code + '</h1>' +
    '<p style="font-size:15px;margin:0 0 14px">' + esc(c.prenom + ' ' + c.nom) + ' · ' + esc(c.tel) + ' · ' + esc(c.email) + '<br>' + (c.retrait ? 'Retrait en boutique' : 'Livraison : ' + esc(c.adresse) + ', ' + esc(c.ville)) + '</p>' +
    bloc('Articles', tableArticles(c)) + '<p style="font-size:14px;color:#7a746b">Le client doit confirmer avec le bouton de son e-mail. Statut actuel : « En attente de confirmation ».</p>' +
    bouton(url, 'Ouvrir le tableau des commandes'));
}
function emailStatut(prenom, texte, code, jeton, statut, note) {
  return cadre('<h1 style="font-family:Georgia,serif;font-weight:400;font-size:26px;margin:0 0 10px">' + esc(statut) + '</h1>' +
    '<p style="font-size:15px;line-height:1.55;margin:0 0 10px">Bonjour ' + esc(prenom) + ',</p><p style="font-size:15px;line-height:1.55;margin:0 0 10px">' + esc(texte) + '</p>' +
    (note ? '<p style="font-size:15px;line-height:1.55;margin:0 0 10px;background:#fff;border-radius:10px;padding:10px 14px">' + esc(note) + '</p>' : '') +
    '<p style="font-size:14px;color:#7a746b">Commande ' + esc(code) + '</p>' + bouton(CONFIG.SITE_URL + '/suivi.html?t=' + jeton, 'Suivre ma commande'));
}

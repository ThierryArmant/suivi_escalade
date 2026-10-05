/**
 * Bloc'Note EPS — stockage en ligne des espaces d'équipe
 * ------------------------------------------------------
 * À coller dans le Google Sheet « Bloc Note EPS » (Extensions > Apps Script).
 *
 * Ce script range, pour chaque équipe :
 *   - ses blocs et tracés   -> onglet « Blocs »
 *   - son topo des voies    -> onglet « Topo »
 *   - ses photos de murs    -> sous-dossiers « photos - <équipe> » à côté de ce Sheet
 * Les équipes et leurs mots de passe sont dans l'onglet « Equipes ». Une équipe se crée soit
 * en ajoutant une ligne dans cet onglet, soit depuis l'appli avec le code d'invitation
 * rangé dans l'onglet « Reglages » (à ne donner qu'aux collègues à qui vous ouvrez l'outil).
 * Un collègue sans code peut le demander depuis l'appli : la demande arrive par e-mail à
 * l'adresse « email_referent » de l'onglet « Reglages », et seules les adresses académiques
 * (@ac-….fr) sont acceptées.
 *
 * Aucune donnée d'élève ne doit être rangée ici.
 */

var ONGLETS = {
  Equipes: ['equipe', 'mot_de_passe', 'etablissement', 'cree_le', 'email', 'dernier_rappel'],
  Blocs:   ['equipe', 'id', 'modifie_le', 'donnees'],
  Topo:    ['equipe', 'modifie_le', 'donnees'],
  Photos:  ['equipe', 'nom', 'fichier_id', 'modifie_le'],
  Reglages: ['cle', 'valeur'],
  Demandes: ['date', 'nom', 'etablissement', 'email'],
  Partages: ['proprietaire', 'demandeur', 'etat', 'date'],
  Validations: ['equipe', 'classe', 'eleve', 'fiche_id', 'date', 'etat'],
  Classes: ['equipe', 'classe', 'codes', 'modifie_le'],
  Suivi: ['equipe', 'classe', 'eleve', 'donnees', 'cree_le', 'modifie_le']
};
var VERSION = 11;
var ELEVES_MAX_PAR_EQUIPE = 3000;      // garde-fou contre le remplissage abusif de la feuille
var SAISIES_MAX_PAR_ELEVE = 400;
var DUREE_SUIVI_JOURS = 183;          // les saisies d'une classe sont effacées 6 mois après la première   // garde-fou contre le remplissage abusif de la feuille
var DEMANDES_MAX_PAR_JOUR = 30;
var DELAI_RAPPEL_MINUTES = 10;
var TAILLE_MORCEAU = 45000;       // une case de Google Sheet accepte 50 000 caractères au maximum
var TAILLE_PHOTO_MAX = 3000000;   // environ 2 Mo par photo

/* =========================================================
   1. INSTALLATION (à lancer une seule fois depuis l'éditeur)
   ========================================================= */
function installer() {
  // Demande toutes les autorisations nécessaires (feuille, Drive, envoi d'e-mails).
  // Si l'une d'elles a été décochée ou n'a jamais été demandée, Google rouvre la fenêtre d'autorisation.
  if (ScriptApp.requireAllScopes) ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  MailApp.getRemainingDailyQuota();

  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(ONGLETS).forEach(function (nom) { ongletPret_(classeur, nom); });

  var equipes = classeur.getSheetByName('Equipes');
  if (equipes.getLastRow() < 2) {
    equipes.appendRow(['giono', 'a-changer', 'Collège Jean Giono', maintenant_()]);
  }
  codeInvitation_(); // crée le code d'invitation s'il n'existe pas encore
  emailReferent_();  // crée la ligne « email_referent » (adresse qui reçoit les demandes d'accès)
  dossierParent_(); // demande dès maintenant l'autorisation d'accéder à Drive
  return 'Installation terminée';
}

function ongletPret_(classeur, nom) {
  var feuille = classeur.getSheetByName(nom);
  if (!feuille) feuille = classeur.insertSheet(nom);
  // La ligne de titres est toujours remise à jour (de nouvelles colonnes peuvent s'ajouter avec les versions)
  var titres = ONGLETS[nom];
  feuille.getRange(1, 1, 1, titres.length).setValues([titres]).setFontWeight('bold');
  feuille.setFrozenRows(1);
  // Tout en texte brut : évite qu'une donnée soit prise pour une formule
  feuille.getRange(1, 1, feuille.getMaxRows(), feuille.getMaxColumns()).setNumberFormat('@');
  return feuille;
}

/* =========================================================
   2. POINTS D'ENTRÉE appelés par l'appli
   ========================================================= */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'photo') return repondre_(traiter_({ action: 'photo', equipe: p.equipe, nom: p.nom }));
  var n = 0;
  try { n = lignes_('Equipes').length; } catch (err) { /* feuille pas encore installée */ }
  return ContentService.createTextOutput("Bloc'Note EPS : stockage en ligne prêt (" + n + " équipe(s)).");
}

function doPost(e) {
  var demande;
  try {
    demande = JSON.parse(e.postData.contents);
  } catch (err) {
    return repondre_({ ok: false, erreur: 'Demande illisible.' });
  }
  return repondre_(traiter_(demande));
}

function repondre_(objet) {
  return ContentService.createTextOutput(JSON.stringify(objet)).setMimeType(ContentService.MimeType.JSON);
}

function traiter_(d) {
  try {
    d = d || {};
    var action = String(d.action || '');
    if (action === 'ping') return { ok: true, service: "Bloc'Note EPS", version: VERSION };
    if (action === 'creerEquipe') return creerEquipe_(d);
    if (action === 'motDePasseOublie') return motDePasseOublie_(d);
    if (action === 'demanderAcces') return demanderAcces_(d);
    if (action === 'photo') return lirePhoto_(cle_(d.equipe), String(d.nom || ''));
    if (action === 'ficheBloc') return ficheBloc_(cle_(d.equipe), String(d.id || ''));
    if (action === 'parcoursSecurite') return parcoursSecurite_(cle_(d.equipe), d.classe, d.eleve);
    if (action === 'validerFiche') {
      var verrouEleve = LockService.getScriptLock();
      verrouEleve.waitLock(25000);
      try { return validerFiche_(cle_(d.equipe), String(d.id || ''), d.classe, d.eleve, d.etat); }
      finally { verrouEleve.releaseLock(); }
    }

    // Toutes les autres actions demandent le mot de passe de l'équipe
    var equipe = verifier_(d.equipe, d.mdp);
    if (!equipe) return { ok: false, erreur: "Nom d'équipe ou mot de passe incorrect." };

    if (action === 'connexion') return { ok: true, equipe: equipe.equipe, etablissement: equipe.etablissement };
    if (action === 'lire') return lireTout_(equipe.equipe);
    if (action === 'annuaire') return annuaire_(equipe.equipe);
    if (action === 'lireBlocsEquipe') return lireBlocsEquipe_(equipe.equipe, cle_(d.cible));
    if (action === 'lireValidations') return lireValidations_(equipe.equipe);
    if (action === 'lireClasses') return lireClasses_(equipe.equipe);

    // Actions qui modifient : une seule à la fois, pour ne pas mélanger deux enregistrements
    var verrou = LockService.getScriptLock();
    verrou.waitLock(25000);
    try {
      if (action === 'enregistrerBlocs') return enregistrerBlocs_(equipe.equipe, d.blocs);
      if (action === 'supprimerBloc') return supprimerBloc_(equipe.equipe, String(d.id || ''));
      if (action === 'enregistrerClasse') return enregistrerClasse_(equipe.equipe, d.classe, d.codes);
      if (action === 'effacerValidations') return effacerValidations_(equipe.equipe, d.classe, d.genre);
      if (action === 'enregistrerTopo') return enregistrerTopo_(equipe.equipe, d.topo);
      if (action === 'enregistrerPhoto') return enregistrerPhoto_(equipe.equipe, String(d.nom || ''), String(d.dataUrl || ''));
      if (action === 'supprimerPhoto') return supprimerPhoto_(equipe.equipe, String(d.nom || ''));
      if (action === 'demanderPartage') return demanderPartage_(equipe.equipe, cle_(d.cible));
      if (action === 'repondrePartage') return repondrePartage_(equipe.equipe, cle_(d.demandeur), d.accepter === true);
      if (action === 'contacterEquipe') return contacterEquipe_(equipe.equipe, cle_(d.cible), String(d.message || ''));
    } finally {
      verrou.releaseLock();
    }
    return { ok: false, erreur: 'Action inconnue : ' + action };
  } catch (err) {
    return { ok: false, erreur: 'Erreur du stockage : ' + (err && err.message ? err.message : err) };
  }
}

/* =========================================================
   3. ÉQUIPES
   ========================================================= */
function cle_(texte) {
  return String(texte == null ? '' : texte).trim().toLowerCase();
}

function verifier_(equipe, mdp) {
  var nom = cle_(equipe);
  var motDePasse = String(mdp == null ? '' : mdp);
  if (!nom || !motDePasse) return null;
  var lignes = lignes_('Equipes');
  for (var i = 0; i < lignes.length; i++) {
    if (cle_(lignes[i][0]) === nom && String(lignes[i][1]) === motDePasse) {
      return { equipe: nom, etablissement: String(lignes[i][2] || '') };
    }
  }
  return null;
}

// Code d'invitation : demandé pour créer une équipe depuis l'appli. Il est lisible et modifiable
// dans l'onglet « Reglages ». Si la case est vidée, plus personne ne peut créer d'équipe depuis l'appli.
function codeInvitation_() {
  return reglage_('code_invitation', function () {
    var lettres = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', code = '';
    for (var j = 0; j < 8; j++) code += lettres.charAt(Math.floor(Math.random() * lettres.length));
    return code;
  });
}

function creerEquipe_(d) {
  var nom = cle_(d.equipe);
  var motDePasse = String(d.mdp == null ? '' : d.mdp);
  var etablissement = String(d.etablissement == null ? '' : d.etablissement).trim().substring(0, 80);
  if (!/^[a-z0-9][a-z0-9 _-]{2,29}$/.test(nom)) {
    return { ok: false, erreur: "Nom d'équipe invalide : 3 à 30 caractères, lettres sans accent, chiffres, espaces ou tirets." };
  }
  if (motDePasse.length < 6 || motDePasse.length > 60) {
    return { ok: false, erreur: 'Le mot de passe doit faire entre 6 et 60 caractères.' };
  }
  var email = emailPropre_(d.email);
  if (!emailAcademique_(email)) {
    return { ok: false, erreur: "L'adresse e-mail doit être une adresse académique (se terminant par @ac-….fr)." };
  }
  var verrou = LockService.getScriptLock();
  verrou.waitLock(25000);
  try {
    var attendu = codeInvitation_();
    if (!attendu) return { ok: false, erreur: "La création d'équipe depuis l'appli est fermée. Demandez au référent de créer votre équipe." };
    if (String(d.invitation == null ? '' : d.invitation).trim().toUpperCase() !== attendu.toUpperCase()) {
      return { ok: false, erreur: "Code d'invitation incorrect." };
    }
    var existe = chercher_('Equipes', function (l) { return cle_(l[0]) === nom; });
    if (existe > 0) return { ok: false, erreur: 'Une équipe porte déjà ce nom. Choisissez-en un autre.' };
    ecrireLigne_(feuille_('Equipes'), -1, [nom, motDePasse, etablissement, maintenant_(), email, '']);
    return { ok: true, equipe: nom, etablissement: etablissement };
  } finally {
    verrou.releaseLock();
  }
}

// Une adresse e-mail ne porte jamais d'accent : « sébastien.dupont@… » devient « sebastien.dupont@… »
// (sinon la messagerie refuse la réponse : « bad UTF-8 syntax »).
function emailPropre_(texte) {
  return String(texte == null ? '' : texte).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '').toLowerCase();
}

function emailValide_(email) {
  return email.length <= 120 && /^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/i.test(email);
}

// Adresse professionnelle de l'Éducation nationale : prenom.nom@ac-academie.fr
function emailAcademique_(email) {
  return emailValide_(email) && /@ac-[a-z0-9-]+\.[a-z]{2,4}$/i.test(email);
}

// Lit une valeur de l'onglet « Reglages » ; crée la ligne (avec la valeur par défaut) si elle manque
function reglage_(cle, parDefaut) {
  var lignes = lignes_('Reglages');
  for (var i = 0; i < lignes.length; i++) {
    if (cle_(lignes[i][0]) === cle) return String(lignes[i][1]).trim();
  }
  var valeur = typeof parDefaut === 'function' ? parDefaut() : parDefaut;
  ecrireLigne_(feuille_('Reglages'), -1, [cle, valeur]);
  return String(valeur);
}

// Adresse qui reçoit les demandes d'accès : case « email_referent » de l'onglet « Reglages »
function emailReferent_() {
  return reglage_('email_referent', function () {
    try { return Session.getEffectiveUser().getEmail() || ''; } catch (err) { return ''; }
  });
}

// Demande d'accès : un collègue sans code d'invitation écrit au référent depuis l'appli.
// Seules les adresses académiques sont acceptées ; le référent répond ensuite à cette adresse avec le code.
function demanderAcces_(d) {
  var nom = String(d.nom == null ? '' : d.nom).trim().replace(/\s+/g, ' ');
  var etablissement = String(d.etablissement == null ? '' : d.etablissement).trim().replace(/\s+/g, ' ');
  var email = emailPropre_(d.email);
  if (nom.length < 2 || nom.length > 80) return { ok: false, erreur: 'Indiquez votre prénom et votre nom.' };
  if (etablissement.length < 2 || etablissement.length > 120) return { ok: false, erreur: 'Indiquez votre établissement.' };
  if (!emailAcademique_(email)) {
    return { ok: false, erreur: "Seules les adresses académiques sont acceptées (se terminant par @ac-….fr)." };
  }
  var verrou = LockService.getScriptLock();
  verrou.waitLock(25000);
  try {
    var referent = emailReferent_();
    if (!emailValide_(referent)) return { ok: false, erreur: "Les demandes d'accès ne sont pas ouvertes pour le moment." };

    var unJour = Date.now() - 24 * 3600000;
    var recentes = lignes_('Demandes').filter(function (l) { var t = Date.parse(String(l[0])); return !isNaN(t) && t > unJour; });
    if (recentes.some(function (l) { return cle_(l[3]) === email; })) {
      return { ok: false, erreur: 'Votre demande a déjà été envoyée. Le référent vous répondra à votre adresse académique.' };
    }
    if (recentes.length >= DEMANDES_MAX_PAR_JOUR) {
      return { ok: false, erreur: "Trop de demandes aujourd'hui. Réessayez demain." };
    }

    var code = codeInvitation_();
    MailApp.sendEmail(referent, "Bloc'Note EPS : demande d'accès de " + nom + ' (' + etablissement + ')',
      'Bonjour,\n\n' +
      "Un collègue demande à créer son équipe dans l'appli Bloc'Note EPS.\n\n" +
      'Nom : ' + nom + '\n' +
      'Établissement : ' + etablissement + '\n' +
      'Adresse académique : ' + email + '\n\n' +
      'POUR ACCEPTER : cliquez sur « Répondre ». Votre réponse partira directement à ' + email + '.\n' +
      "Indiquez-lui le code d'invitation : " + (code || "(aucun code : la création d'équipe est fermée dans l'onglet Reglages)") + '\n\n' +
      'POUR REFUSER : ne répondez pas, rien ne sera créé.',
      { replyTo: email, name: "Bloc'Note EPS" });
    ecrireLigne_(feuille_('Demandes'), -1, [maintenant_(), nom, etablissement, email]);
    return { ok: true };
  } finally {
    verrou.releaseLock();
  }
}

function masquerEmail_(email) {
  var parties = email.split('@');
  return parties[0].charAt(0) + '***@' + parties[1];
}

// Mot de passe oublié : il est envoyé à l'adresse enregistrée pour l'équipe (colonne « email »),
// jamais affiché dans l'appli. Un seul envoi toutes les 10 minutes par équipe.
function motDePasseOublie_(d) {
  var nom = cle_(d.equipe);
  if (!nom) return { ok: false, erreur: "Indiquez le nom de l'équipe." };
  var verrou = LockService.getScriptLock();
  verrou.waitLock(25000);
  try {
    var position = chercher_('Equipes', function (l) { return cle_(l[0]) === nom; });
    if (position < 0) return { ok: false, erreur: 'Aucune équipe ne porte ce nom.' };
    var feuille = feuille_('Equipes');
    var ligne = feuille.getRange(position, 1, 1, ONGLETS.Equipes.length).getValues()[0];
    var email = String(ligne[4] || '').trim();
    if (!emailValide_(email)) {
      return { ok: false, erreur: "Aucune adresse e-mail n'est enregistrée pour cette équipe. Contactez le référent de l'outil." };
    }
    var dernier = Date.parse(String(ligne[5] || ''));
    if (!isNaN(dernier) && (Date.now() - dernier) < DELAI_RAPPEL_MINUTES * 60000) {
      return { ok: false, erreur: 'Un e-mail vient déjà d\'être envoyé. Vérifiez la boîte ' + masquerEmail_(email) + ' (et les indésirables), ou réessayez dans ' + DELAI_RAPPEL_MINUTES + ' minutes.' };
    }
    MailApp.sendEmail(email, "Bloc'Note EPS : mot de passe de l'équipe « " + nom + " »",
      'Bonjour,\n\n' +
      "Le mot de passe de votre équipe a été demandé depuis l'appli Bloc'Note EPS.\n\n" +
      "Nom de l'équipe : " + nom + '\n' +
      'Mot de passe : ' + String(ligne[1]) + '\n\n' +
      "Si vous n'êtes pas à l'origine de cette demande, vous pouvez ignorer ce message.");
    var plage = feuille.getRange(position, 6);
    plage.setNumberFormat('@');
    plage.setValues([[maintenant_()]]);
    return { ok: true, envoyeA: masquerEmail_(email) };
  } finally {
    verrou.releaseLock();
  }
}

/* =========================================================
   4. LECTURE COMPLÈTE D'UN ESPACE D'ÉQUIPE
   ========================================================= */
function lireTout_(equipe) {
  var blocs = [];
  lignes_('Blocs').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    var contenu = recoller_(l, 3);
    if (!contenu) return;
    try {
      var bloc = JSON.parse(contenu);
      blocs.push({ id: String(l[1]), modifie_le: String(l[2]), route: bloc.route, markers: bloc.markers || [] });
    } catch (err) { /* ligne abîmée : on l'ignore */ }
  });

  var topo = null;
  lignes_('Topo').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    try { topo = JSON.parse(recoller_(l, 2)); } catch (err) { topo = null; }
  });

  var photos = [];
  lignes_('Photos').forEach(function (l) {
    if (cle_(l[0]) === equipe) photos.push({ nom: String(l[1]), modifie_le: String(l[3]) });
  });

  return { ok: true, equipe: equipe, blocs: blocs, topo: topo, photos: photos };
}

/* =========================================================
   4 bis. LES ÉQUIPES : annuaire, titres des blocs, partage sur demande
   =========================================================
   Chaque équipe connectée voit la liste des équipes et les TITRES de leurs blocs.
   Pour consulter les blocs eux-mêmes, elle envoie une demande que l'autre équipe accepte ou refuse.
   Les adresses e-mail ne sont jamais communiquées à l'appli. */
function infosEquipes_() {
  var infos = {};
  lignes_('Equipes').forEach(function (l) {
    var nom = cle_(l[0]);
    if (nom) infos[nom] = { etablissement: String(l[2] || ''), email: String(l[4] || '').trim() };
  });
  return infos;
}

function blocsDe_(equipe, complet) {
  var blocs = [];
  lignes_('Blocs').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    try {
      var bloc = JSON.parse(recoller_(l, 3));
      if (complet) blocs.push({ id: String(l[1]), route: bloc.route, markers: bloc.markers || [] });
      else blocs.push({ name: String(bloc.route.name || ''), grade: String(bloc.route.grade || ''), theme: String(bloc.route.theme || '') });
    } catch (err) { /* ligne abîmée : on l'ignore */ }
  });
  return blocs;
}

function partage_(proprietaire, demandeur) {
  var position = chercher_('Partages', function (l) { return cle_(l[0]) === proprietaire && cle_(l[1]) === demandeur; });
  if (position < 0) return { position: -1, etat: 'aucun', date: '' };
  var ligne = feuille_('Partages').getRange(position, 1, 1, 4).getValues()[0];
  return { position: position, etat: String(ligne[2]), date: String(ligne[3]) };
}

function annuaire_(moi) {
  var infos = infosEquipes_();
  var partages = lignes_('Partages');
  var equipes = Object.keys(infos).sort().map(function (nom) {
    var acces = nom === moi ? 'moi' : 'aucun';
    partages.forEach(function (l) { if (cle_(l[0]) === nom && cle_(l[1]) === moi) acces = String(l[2]); });
    return { equipe: nom, etablissement: infos[nom].etablissement, blocs: blocsDe_(nom, false), acces: acces, contact: emailValide_(infos[nom].email) };
  });
  var recues = [], donnes = [];
  partages.forEach(function (l) {
    if (cle_(l[0]) !== moi) return;
    var demandeur = cle_(l[1]);
    var fiche = { demandeur: demandeur, etablissement: (infos[demandeur] || {}).etablissement || '', date: String(l[3]) };
    if (String(l[2]) === 'demande') recues.push(fiche);
    if (String(l[2]) === 'accorde') donnes.push(fiche);
  });
  return { ok: true, moi: moi, equipes: equipes, demandesRecues: recues, accesDonnes: donnes };
}

function envoyer_(destinataire, sujet, texte, repondreA) {
  if (!emailValide_(destinataire)) return false;
  try {
    var options = { name: "Bloc'Note EPS" };
    if (emailValide_(repondreA)) options.replyTo = repondreA;
    MailApp.sendEmail(destinataire, sujet, texte, options);
    return true;
  } catch (err) {
    return false;
  }
}

function demanderPartage_(moi, cible) {
  var infos = infosEquipes_();
  if (!infos[cible] || cible === moi) return { ok: false, erreur: 'Équipe introuvable.' };
  var actuel = partage_(cible, moi);
  if (actuel.etat === 'accorde') return { ok: true, acces: 'accorde' };
  var date = Date.parse(actuel.date);
  if (actuel.etat === 'demande' && !isNaN(date) && Date.now() - date < 24 * 3600000) {
    return { ok: false, erreur: 'Votre demande a déjà été envoyée à cette équipe. Elle doit maintenant y répondre.' };
  }
  ecrireLigne_(feuille_('Partages'), actuel.position, [cible, moi, 'demande', maintenant_()]);
  var prevenu = envoyer_(infos[cible].email, "Bloc'Note EPS : l'équipe « " + moi + " » souhaite consulter vos blocs",
    'Bonjour,\n\n' +
    "L'équipe « " + moi + " »" + (infos[moi].etablissement ? ' (' + infos[moi].etablissement + ')' : '') + " souhaite consulter les blocs de votre équipe « " + cible + " » dans l'appli Bloc'Note EPS.\n\n" +
    "POUR ACCEPTER OU REFUSER : ouvrez l'appli, connectez-vous à votre équipe, cliquez sur « Les équipes », puis regardez la rubrique « Demandes reçues ».\n\n" +
    "Accepter permet à cette équipe de voir vos blocs et de les recopier chez elle. Elle ne peut ni les modifier ni les supprimer chez vous.\n\n" +
    'Pour écrire à ce collègue, répondez simplement à ce message.',
    infos[moi].email);
  return { ok: true, acces: 'demande', prevenu: prevenu };
}

function repondrePartage_(moi, demandeur, accepter) {
  var actuel = partage_(moi, demandeur);
  if (actuel.position < 0) return { ok: false, erreur: 'Demande introuvable.' };
  var infos = infosEquipes_();
  ecrireLigne_(feuille_('Partages'), actuel.position, [moi, demandeur, accepter ? 'accorde' : 'refuse', maintenant_()]);
  if (infos[demandeur] && actuel.etat === 'demande') {
    envoyer_(infos[demandeur].email, "Bloc'Note EPS : réponse de l'équipe « " + moi + " »",
      'Bonjour,\n\n' +
      "L'équipe « " + moi + " » a " + (accepter ? 'accepté' : 'refusé') + " votre demande de consultation de ses blocs.\n\n" +
      (accepter ? "Ouvrez l'appli, cliquez sur « Les équipes », puis sur le titre d'un bloc de cette équipe pour l'afficher.\n\n" : '') +
      'Pour écrire à ce collègue, répondez simplement à ce message.',
      (infos[moi] || {}).email);
  }
  return { ok: true, acces: accepter ? 'accorde' : 'refuse' };
}

function lireBlocsEquipe_(moi, cible) {
  if (cible !== moi && partage_(cible, moi).etat !== 'accorde') {
    return { ok: false, erreur: "Cette équipe ne vous a pas (encore) autorisé à consulter ses blocs." };
  }
  return { ok: true, equipe: cible, blocs: blocsDe_(cible, true) };
}

// Message libre d'une équipe à une autre : il part à l'adresse du responsable, qui n'est jamais montrée
function contacterEquipe_(moi, cible, message) {
  var infos = infosEquipes_();
  if (!infos[cible] || cible === moi) return { ok: false, erreur: 'Équipe introuvable.' };
  var texte = message.trim();
  if (texte.length < 5 || texte.length > 1500) return { ok: false, erreur: 'Écrivez un message de 5 à 1500 caractères.' };
  if (!emailValide_(infos[moi].email)) {
    return { ok: false, erreur: "Votre équipe n'a pas d'adresse e-mail enregistrée : le collègue ne pourrait pas vous répondre. Demandez au référent de l'outil de l'ajouter." };
  }
  if (!emailValide_(infos[cible].email)) return { ok: false, erreur: "Cette équipe n'a pas d'adresse e-mail enregistrée." };
  var cache = CacheService.getScriptCache();
  var cleCache = 'contact|' + moi + '|' + cible;
  if (cache.get(cleCache)) return { ok: false, erreur: 'Vous venez déjà d\'écrire à cette équipe. Réessayez dans 10 minutes.' };
  var envoye = envoyer_(infos[cible].email, "Bloc'Note EPS : message de l'équipe « " + moi + " »",
    'Bonjour,\n\n' +
    "L'équipe « " + moi + " »" + (infos[moi].etablissement ? ' (' + infos[moi].etablissement + ')' : '') + " vous écrit depuis l'appli Bloc'Note EPS :\n\n" +
    '----------\n' + texte + '\n----------\n\n' +
    'Pour lui répondre, répondez simplement à ce message.',
    infos[moi].email);
  if (!envoye) return { ok: false, erreur: "Le message n'a pas pu être envoyé." };
  cache.put(cleCache, '1', 600);
  return { ok: true };
}

/* =========================================================
   5. BLOCS ET TRACÉS
   ========================================================= */
function enregistrerBlocs_(equipe, blocs) {
  if (!Array.isArray(blocs)) return { ok: false, erreur: 'Aucun bloc reçu.' };
  var feuille = feuille_('Blocs');
  var compte = 0;
  blocs.forEach(function (b) {
    if (!b || !b.id || !b.route) return;
    var id = String(b.id);
    var contenu = JSON.stringify({ route: b.route, markers: Array.isArray(b.markers) ? b.markers : [] });
    var ligne = [equipe, id, maintenant_()].concat(decouper_(contenu));
    var position = chercher_('Blocs', function (l) { return cle_(l[0]) === equipe && String(l[1]) === id; });
    ecrireLigne_(feuille, position, ligne);
    compte++;
  });
  return { ok: true, enregistres: compte };
}

function supprimerBloc_(equipe, id) {
  var position = chercher_('Blocs', function (l) { return cle_(l[0]) === equipe && String(l[1]) === id; });
  if (position > 0) feuille_('Blocs').deleteRow(position);
  return { ok: true, supprime: position > 0 };
}

/* =========================================================
   6. TOPO DES VOIES
   ========================================================= */
function enregistrerTopo_(equipe, topo) {
  if (!topo || !Array.isArray(topo.routes)) return { ok: false, erreur: 'Aucun topo reçu.' };
  var ligne = [equipe, maintenant_()].concat(decouper_(JSON.stringify(topo)));
  var position = chercher_('Topo', function (l) { return cle_(l[0]) === equipe; });
  ecrireLigne_(feuille_('Topo'), position, ligne);
  return { ok: true, voies: topo.routes.length };
}

/* =========================================================
   7. PHOTOS (rangées dans Drive, à côté de ce Sheet)
   ========================================================= */
function dossierParent_() {
  var fichier = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parents = fichier.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function dossierEquipe_(equipe) {
  var parent = dossierParent_();
  var nom = 'photos - ' + equipe;
  var existants = parent.getFoldersByName(nom);
  return existants.hasNext() ? existants.next() : parent.createFolder(nom);
}

function nomPhotoValide_(nom) {
  return nom.length > 0 && nom.length <= 150 && nom.indexOf('/') === -1 && nom.indexOf('\\') === -1;
}

function enregistrerPhoto_(equipe, nom, dataUrl) {
  if (!nomPhotoValide_(nom)) return { ok: false, erreur: 'Nom de photo invalide.' };
  var trouve = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(dataUrl);
  if (!trouve) return { ok: false, erreur: "Ce fichier n'est pas une photo." };
  if (trouve[2].length > TAILLE_PHOTO_MAX) return { ok: false, erreur: 'Photo trop lourde.' };

  var blob = Utilities.newBlob(Utilities.base64Decode(trouve[2]), trouve[1], nom);
  var position = chercher_('Photos', function (l) { return cle_(l[0]) === equipe && String(l[1]) === nom; });
  if (position > 0) {
    var ancienne = String(feuille_('Photos').getRange(position, 3).getValue());
    try { DriveApp.getFileById(ancienne).setTrashed(true); } catch (err) { /* déjà supprimée */ }
  }
  var fichier = dossierEquipe_(equipe).createFile(blob);
  ecrireLigne_(feuille_('Photos'), position, [equipe, nom, fichier.getId(), maintenant_()]);
  return { ok: true, nom: nom };
}

function supprimerPhoto_(equipe, nom) {
  var position = chercher_('Photos', function (l) { return cle_(l[0]) === equipe && String(l[1]) === nom; });
  if (position > 0) {
    var feuille = feuille_('Photos');
    try { DriveApp.getFileById(String(feuille.getRange(position, 3).getValue())).setTrashed(true); } catch (err) { /* déjà supprimée */ }
    feuille.deleteRow(position);
  }
  return { ok: true, supprime: position > 0 };
}

// Fiche d'un bloc pour l'élève : sans mot de passe, comme la photo du mur.
// Ne renvoie que les textes destinés aux élèves (jamais les notes de l'enseignant).
function ficheBloc_(equipe, id) {
  if (!equipe || !/^custom_[0-9A-Za-z_-]{1,40}$/.test(id)) return { ok: false, erreur: 'Bloc introuvable.' };
  var lignes = lignes_('Blocs');
  for (var i = 0; i < lignes.length; i++) {
    if (cle_(lignes[i][0]) === equipe && String(lignes[i][1]) === id) {
      try {
        var route = JSON.parse(recoller_(lignes[i], 3)).route || {};
        return {
          ok: true,
          nom: String(route.name || ''),
          consigne: String(route.consigne || ''),
          attendus: String(route.attendus || ''),
          competences: String(route.competences || ''),
          source: String(route.source || ''),
          video: String(route.video || ''),
          suivi: ficheSuivie_(route),
          suiviBloc: !ficheSuivie_(route) && route.type === 'bloc'
        };
      } catch (err) { break; }
    }
  }
  return { ok: false, erreur: 'Bloc introuvable.' };
}

/* =========================================================
   SUIVI DES FICHES DE SÉCURITÉ
   L'élève ne donne jamais son nom : seulement sa classe et un code
   (initiale du prénom, initiale du nom, numéro dans la liste), ex. « TA 12 ».
   ========================================================= */
// Seules les fiches du thème « Sécurité… » sont suivies
function ficheSuivie_(route) {
  return /^s[ée]curit[ée]/i.test(String(route && route.theme || '').trim());
}

function classePropre_(texte) {
  return String(texte == null ? '' : texte).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^0-9A-Za-z ]/g, '').replace(/\s+/g, ' ').trim().toUpperCase().slice(0, 10);
}

// « t a 12 », « ta12 », « T.A. 12 » → « TA 12 » ; tout le reste est refusé
function codeElevePropre_(texte) {
  var trouve = /^([A-Z])[ .\-]*([A-Z])[ .\-]*(\d{1,2})$/.exec(String(texte == null ? '' : texte).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase());
  return trouve ? trouve[1] + trouve[2] + ' ' + parseInt(trouve[3], 10) : '';
}

// Validation envoyée par le téléphone de l'élève : sans mot de passe, donc très encadrée
function validerFiche_(equipe, id, classe, eleve, etat) {
  classe = classePropre_(classe);
  eleve = codeElevePropre_(eleve);
  if (!equipe || !/^custom_[0-9A-Za-z_-]{1,40}$/.test(id)) return { ok: false, erreur: 'Fiche introuvable.' };
  if (!classe) return { ok: false, erreur: 'Indique ta classe (par exemple 4B).' };
  if (!eleve) return { ok: false, erreur: 'Indique ton code : initiale du prénom, initiale du nom, puis ton numéro dans la liste (par exemple TA 12).' };

  // La fiche doit exister dans cette équipe et faire partie du thème « Sécurité »
  // Un exercice de bloc se déclare « essayé » ou « réussi » ; une fiche de sécurité se déclare seulement « vue »
  var suivie = false, exercice = false;
  var blocs = lignes_('Blocs');
  for (var i = 0; i < blocs.length; i++) {
    if (cle_(blocs[i][0]) === equipe && String(blocs[i][1]) === id) {
      try {
        var route = JSON.parse(recoller_(blocs[i], 3)).route || {};
        suivie = ficheSuivie_(route);
        exercice = !suivie && route.type === 'bloc';
      } catch (err) { suivie = false; exercice = false; }
      break;
    }
  }
  if (!suivie && !exercice) return { ok: false, erreur: "Cette fiche n'est pas suivie." };
  etat = exercice ? (String(etat || '') === 'reussi' ? 'reussi' : 'essaye') : '';

  // Une ligne par élève : on ne relit que les trois premières colonnes pour le retrouver
  var feuille = feuille_('Suivi');
  var derniere = feuille.getLastRow();
  var cles = derniere < 2 ? [] : feuille.getRange(2, 1, derniere - 1, 3).getValues();
  var total = 0, ligne = -1;
  for (var j = 0; j < cles.length; j++) {
    if (cle_(cles[j][0]) !== equipe) continue;
    total++;
    if (String(cles[j][1]) === classe && String(cles[j][2]) === eleve) { ligne = j + 2; break; }
  }
  var quand = maintenant_();
  if (ligne === -1) {
    if (total >= ELEVES_MAX_PAR_EQUIPE) return { ok: false, erreur: 'Le suivi est plein : préviens ton professeur.' };
    var premier = {};
    premier[id] = [etat, quand];
    feuille.appendRow([equipe, classe, eleve, JSON.stringify(premier), quand, quand]);
    return { ok: true, etat: etat };
  }
  var plage = feuille.getRange(ligne, 4, 1, 3);
  var valeurs = plage.getValues()[0];
  var d = donneesSuivi_(valeurs[0]);
  var cree = String(valeurs[1]);
  // Saisies de plus de 6 mois : un nouveau cycle commence, on repart de zéro
  if (tropVieux_(cree)) { d = {}; cree = quand; }
  if (d[id]) {
    // Déjà notée : un bloc « essayé » peut devenir « réussi », jamais l'inverse
    if (!(etat === 'reussi' && String(d[id][0]) !== 'reussi')) return { ok: true, deja: true, etat: String(d[id][0] || '') };
  } else if (Object.keys(d).length >= SAISIES_MAX_PAR_ELEVE) {
    return { ok: false, erreur: 'Le suivi est plein : préviens ton professeur.' };
  }
  d[id] = [etat, quand];
  plage.setValues([[JSON.stringify(d), cree, quand]]);
  return { ok: true, etat: etat };
}

// Parcours d'un élève : toutes les fiches de sécurité de l'équipe, avec ce qu'il a déjà vu.
// Sans mot de passe (l'élève y accède de chez lui) ; ne renvoie que ce qu'un QR code de fiche montre déjà.
function parcoursSecurite_(equipe, classe, eleve) {
  classe = classePropre_(classe);
  eleve = codeElevePropre_(eleve);
  if (!equipe || !classe || !eleve) return { ok: false, erreur: 'Indique ta classe et ton code.' };
  var vues = {};
  lignes_('Validations').forEach(function (l) {
    if (cle_(l[0]) === equipe && String(l[1]) === classe && String(l[2]) === eleve) vues[String(l[3])] = true;
  });
  lignes_('Suivi').forEach(function (l) {
    if (cle_(l[0]) === equipe && String(l[1]) === classe && String(l[2]) === eleve) Object.keys(donneesSuivi_(l[3])).forEach(function (id) { vues[id] = true; });
  });
  var fiches = [];
  lignes_('Blocs').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    try {
      var bloc = JSON.parse(recoller_(l, 3));
      if (!ficheSuivie_(bloc.route)) return;
      fiches.push({
        id: String(l[1]),
        nom: String(bloc.route.name || ''),
        theme: String(bloc.route.theme || ''),
        niveau: String(bloc.route.grade || ''),
        gabarit: String(bloc.route.height || ''),
        couleur: String(bloc.route.color || ''),
        image: String(bloc.route.imageName || ''),
        markers: bloc.markers || [],
        vue: vues[String(l[1])] === true
      });
    } catch (err) { /* ligne abîmée : on l'ignore */ }
  });
  return { ok: true, fiches: fiches };
}

// Les saisies sont rangées par élève (une ligne par élève, onglet « Suivi ») : chaque validation ne relit que
// les trois premières colonnes, ce qui reste rapide même avec beaucoup d'établissements.
// L'ancien onglet « Validations » (une ligne par saisie) est encore lu, pour ne rien perdre.
function donneesSuivi_(cellule) {
  try { var d = JSON.parse(String(cellule) || '{}'); return (d && typeof d === 'object') ? d : {}; } catch (err) { return {}; }
}

function tropVieux_(date) {
  var t = new Date(String(date)).getTime();
  return !isNaN(t) && (Date.now() - t) > DUREE_SUIVI_JOURS * 86400000;
}

// Ménage : les saisies d'une classe sont effacées 6 mois après sa première saisie. Renvoie le nombre de lignes retirées.
function menageSuivi_(equipe) {
  var feuille = feuille_('Suivi');
  var lignes = lignes_('Suivi');
  var debut = {};
  lignes.forEach(function (l) {
    if (equipe && cle_(l[0]) !== equipe) return;
    var cle = cle_(l[0]) + '|' + String(l[1]);
    var t = String(l[4]);
    if (!debut[cle] || t < debut[cle]) debut[cle] = t;
  });
  var retire = 0;
  for (var i = lignes.length - 1; i >= 0; i--) {
    if (equipe && cle_(lignes[i][0]) !== equipe) continue;
    if (tropVieux_(debut[cle_(lignes[i][0]) + '|' + String(lignes[i][1])])) { feuille.deleteRow(i + 2); retire++; }
  }
  var anciennes = lignes_('Validations');
  var ancienne = feuille_('Validations');
  for (var j = anciennes.length - 1; j >= 0; j--) {
    if (equipe && cle_(anciennes[j][0]) !== equipe) continue;
    if (tropVieux_(anciennes[j][4])) { ancienne.deleteRow(j + 2); retire++; }
  }
  return retire;
}

function lireValidations_(equipe) {
  var verrou = LockService.getScriptLock();
  verrou.waitLock(25000);
  try { menageSuivi_(equipe); } finally { verrou.releaseLock(); }
  var liste = [];
  var debut = {};
  lignes_('Validations').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    liste.push({ classe: String(l[1]), eleve: String(l[2]), id: String(l[3]), date: String(l[4]), etat: String(l[5] || '') });
    if (!debut[String(l[1])] || String(l[4]) < debut[String(l[1])]) debut[String(l[1])] = String(l[4]);
  });
  lignes_('Suivi').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    var d = donneesSuivi_(l[3]);
    Object.keys(d).forEach(function (id) {
      liste.push({ classe: String(l[1]), eleve: String(l[2]), id: id, date: String(d[id][1] || ''), etat: String(d[id][0] || '') });
    });
    if (!debut[String(l[1])] || String(l[4]) < debut[String(l[1])]) debut[String(l[1])] = String(l[4]);
  });
  // Date à laquelle les saisies de chaque classe seront effacées
  var echeances = {};
  Object.keys(debut).forEach(function (classe) {
    var t = new Date(debut[classe]).getTime();
    if (!isNaN(t)) echeances[classe] = new Date(t + DUREE_SUIVI_JOURS * 86400000).toISOString();
  });
  return { ok: true, validations: liste, echeances: echeances };
}

// Classes de l'équipe : seulement les codes des élèves (« TA 12 »), jamais les noms.
// Sert à retrouver les mêmes classes sur tous les appareils du professeur.
function lireClasses_(equipe) {
  var classes = {};
  lignes_('Classes').forEach(function (l) {
    if (cle_(l[0]) !== equipe) return;
    try { classes[String(l[1])] = JSON.parse(String(l[2]) || '[]'); } catch (err) { /* ligne abîmée : on l'ignore */ }
  });
  return { ok: true, classes: classes };
}

function enregistrerClasse_(equipe, classe, codes) {
  classe = classePropre_(classe);
  if (!classe) return { ok: false, erreur: 'Classe manquante.' };
  var propres = [];
  (Array.isArray(codes) ? codes : []).slice(0, 60).forEach(function (c) {
    c = codeElevePropre_(c);
    if (c && propres.indexOf(c) === -1) propres.push(c);
  });
  var ligne = chercher_('Classes', function (l) { return cle_(l[0]) === equipe && String(l[1]) === classe; });
  var feuille = feuille_('Classes');
  if (!propres.length) { if (ligne > 0) feuille.deleteRow(ligne); return { ok: true, efface: true }; }
  var valeurs = [equipe, classe, JSON.stringify(propres), maintenant_()];
  if (ligne > 0) feuille.getRange(ligne, 1, 1, 4).setValues([valeurs]); else feuille.appendRow(valeurs);
  return { ok: true };
}

// Efface le suivi d'une classe (ou de toute l'équipe si aucune classe n'est donnée)
function effacerValidations_(equipe, classe, genre) {
  classe = classePropre_(classe);
  var efface = 0;
  // genre « blocs » : seulement les exercices (essayé / réussi) ; « secu » : seulement les fiches de sécurité
  var feuille = feuille_('Validations');
  var lignes = lignes_('Validations');
  for (var i = lignes.length - 1; i >= 0; i--) {
    var estBloc = String(lignes[i][5] || '') !== '';
    if (genre === 'blocs' && !estBloc) continue;
    if (genre === 'secu' && estBloc) continue;
    if (cle_(lignes[i][0]) === equipe && (!classe || String(lignes[i][1]) === classe)) {
      feuille.deleteRow(i + 2);   // +2 : la ligne de titres, et les lignes comptées à partir de 1
      efface++;
    }
  }
  var suivi = feuille_('Suivi');
  var eleves = lignes_('Suivi');
  for (var j = eleves.length - 1; j >= 0; j--) {
    if (cle_(eleves[j][0]) !== equipe || (classe && String(eleves[j][1]) !== classe)) continue;
    var d = donneesSuivi_(eleves[j][3]);
    var garde = {};
    Object.keys(d).forEach(function (id) {
      var bloc = String(d[id][0] || '') !== '';
      if ((genre === 'blocs' && !bloc) || (genre === 'secu' && bloc)) garde[id] = d[id]; else efface++;
    });
    if (Object.keys(garde).length) suivi.getRange(j + 2, 4).setValue(JSON.stringify(garde)); else suivi.deleteRow(j + 2);
  }
  return { ok: true, efface: efface };
}

/* =========================================================
   TABLEAU DE BORD (pour le propriétaire du Sheet seulement)
   Menu « Bloc'Note EPS » > « Tableau de bord » : un onglet « Bord » résume la place prise par chaque équipe.
   ========================================================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Bloc'Note EPS")
    .addItem('Tableau de bord', 'tableauDeBord')
    .addItem('Ménage des saisies de plus de 6 mois', 'menageComplet')
    .addToUi();
}

function menageComplet() {
  var n = menageSuivi_('');
  SpreadsheetApp.getUi().alert(n + ' ligne(s) de saisies de plus de 6 mois effacée(s).');
}

function tableauDeBord() {
  var bord = {};
  var une = function (equipe) { return bord[equipe] = bord[equipe] || { blocs: 0, photos: 0, classes: 0, eleves: 0, saisies: 0, derniere: '' }; };
  lignes_('Equipes').forEach(function (l) { if (cle_(l[0])) une(cle_(l[0])); });
  lignes_('Blocs').forEach(function (l) { une(cle_(l[0])).blocs++; });
  lignes_('Photos').forEach(function (l) { une(cle_(l[0])).photos++; });
  lignes_('Classes').forEach(function (l) { une(cle_(l[0])).classes++; });
  lignes_('Suivi').forEach(function (l) {
    var e = une(cle_(l[0]));
    e.eleves++;
    e.saisies += Object.keys(donneesSuivi_(l[3])).length;
    if (String(l[5]) > e.derniere) e.derniere = String(l[5]);
  });
  lignes_('Validations').forEach(function (l) { var e = une(cle_(l[0])); e.saisies++; if (String(l[4]) > e.derniere) e.derniere = String(l[4]); });
  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = classeur.getSheetByName('Bord') || classeur.insertSheet('Bord');
  feuille.clear();
  var lignes = [['equipe', 'blocs', 'photos', 'classes', 'eleves_suivis', 'saisies', 'derniere_saisie']];
  Object.keys(bord).sort(function (a, b) { return bord[b].saisies - bord[a].saisies; }).forEach(function (k) {
    var e = bord[k];
    lignes.push([k, e.blocs, e.photos, e.classes, e.eleves, e.saisies, e.derniere.slice(0, 10)]);
  });
  feuille.getRange(1, 1, lignes.length, 7).setValues(lignes);
  feuille.getRange(1, 1, 1, 7).setFontWeight('bold');
  feuille.setFrozenRows(1);
  var cases = 0;
  classeur.getSheets().forEach(function (f) { cases += f.getMaxRows() * f.getMaxColumns(); });
  feuille.getRange(lignes.length + 2, 1, 1, 2).setValues([['Cases utilisées dans le classeur (maximum 10 000 000)', cases]]);
  classeur.setActiveSheet(feuille);
}

// Lecture d'une photo : sans mot de passe, pour que le téléphone d'un élève puisse afficher le mur
function lirePhoto_(equipe, nom) {
  if (!equipe || !nomPhotoValide_(nom)) return { ok: false, erreur: 'Photo introuvable.' };
  var lignes = lignes_('Photos');
  for (var i = 0; i < lignes.length; i++) {
    if (cle_(lignes[i][0]) === equipe && String(lignes[i][1]) === nom) {
      var blob = DriveApp.getFileById(String(lignes[i][2])).getBlob();
      return {
        ok: true,
        nom: nom,
        modifie_le: String(lignes[i][3]),
        dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes())
      };
    }
  }
  return { ok: false, erreur: 'Photo introuvable.' };
}

/* =========================================================
   8. OUTILS
   ========================================================= */
function maintenant_() {
  return new Date().toISOString();
}

function feuille_(nom) {
  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  return classeur.getSheetByName(nom) || ongletPret_(classeur, nom);
}

// Toutes les lignes d'un onglet, sans la ligne de titres
function lignes_(nom) {
  var feuille = feuille_(nom);
  var derniere = feuille.getLastRow();
  if (derniere < 2) return [];
  return feuille.getRange(2, 1, derniere - 1, Math.max(feuille.getLastColumn(), ONGLETS[nom].length)).getValues();
}

// Numéro de la ligne (dans le Sheet) qui correspond au test, ou -1
function chercher_(nom, test) {
  var lignes = lignes_(nom);
  for (var i = 0; i < lignes.length; i++) {
    if (test(lignes[i])) return i + 2;
  }
  return -1;
}

// Remplace la ligne trouvée, ou en ajoute une nouvelle à la fin
function ecrireLigne_(feuille, position, valeurs) {
  var ligne = position > 0 ? position : feuille.getLastRow() + 1;
  var largeur = Math.max(valeurs.length, feuille.getLastColumn(), 1);
  var complete = valeurs.slice();
  while (complete.length < largeur) complete.push('');
  if (feuille.getMaxColumns() < largeur) feuille.insertColumnsAfter(feuille.getMaxColumns(), largeur - feuille.getMaxColumns());
  var plage = feuille.getRange(ligne, 1, 1, largeur);
  plage.setNumberFormat('@');
  plage.setValues([complete]);
}

// Un long texte est réparti sur plusieurs cases, puis recollé à la lecture
function decouper_(texte) {
  var morceaux = [];
  for (var i = 0; i < texte.length; i += TAILLE_MORCEAU) morceaux.push(texte.substring(i, i + TAILLE_MORCEAU));
  return morceaux.length ? morceaux : [''];
}

function recoller_(ligne, depuis) {
  var texte = '';
  for (var i = depuis; i < ligne.length; i++) texte += String(ligne[i] == null ? '' : ligne[i]);
  return texte;
}

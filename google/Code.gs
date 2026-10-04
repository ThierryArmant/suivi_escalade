/**
 * Bloc'Note EPS — stockage en ligne des espaces d'équipe
 * ------------------------------------------------------
 * À coller dans le Google Sheet « Bloc Note EPS » (Extensions > Apps Script).
 *
 * Ce script range, pour chaque équipe :
 *   - ses blocs et tracés   -> onglet « Blocs »
 *   - son topo des voies    -> onglet « Topo »
 *   - ses photos de murs    -> sous-dossiers « photos - <équipe> » à côté de ce Sheet
 * Les équipes et leurs mots de passe sont dans l'onglet « Equipes » : c'est vous qui les créez,
 * en ajoutant une ligne par équipe.
 *
 * Aucune donnée d'élève ne doit être rangée ici.
 */

var ONGLETS = {
  Equipes: ['equipe', 'mot_de_passe', 'etablissement', 'cree_le'],
  Blocs:   ['equipe', 'id', 'modifie_le', 'donnees'],
  Topo:    ['equipe', 'modifie_le', 'donnees'],
  Photos:  ['equipe', 'nom', 'fichier_id', 'modifie_le']
};
var TAILLE_MORCEAU = 45000;       // une case de Google Sheet accepte 50 000 caractères au maximum
var TAILLE_PHOTO_MAX = 3000000;   // environ 2 Mo par photo

/* =========================================================
   1. INSTALLATION (à lancer une seule fois depuis l'éditeur)
   ========================================================= */
function installer() {
  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(ONGLETS).forEach(function (nom) { ongletPret_(classeur, nom); });

  var equipes = classeur.getSheetByName('Equipes');
  if (equipes.getLastRow() < 2) {
    equipes.appendRow(['giono', 'a-changer', 'Collège Jean Giono', maintenant_()]);
  }
  dossierParent_(); // demande dès maintenant l'autorisation d'accéder à Drive
  return 'Installation terminée';
}

function ongletPret_(classeur, nom) {
  var feuille = classeur.getSheetByName(nom);
  if (!feuille) feuille = classeur.insertSheet(nom);
  if (feuille.getLastRow() === 0) {
    var titres = ONGLETS[nom];
    feuille.getRange(1, 1, 1, titres.length).setValues([titres]).setFontWeight('bold');
    feuille.setFrozenRows(1);
  }
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
    if (action === 'ping') return { ok: true, service: "Bloc'Note EPS" };
    if (action === 'photo') return lirePhoto_(cle_(d.equipe), String(d.nom || ''));

    // Toutes les autres actions demandent le mot de passe de l'équipe
    var equipe = verifier_(d.equipe, d.mdp);
    if (!equipe) return { ok: false, erreur: "Nom d'équipe ou mot de passe incorrect." };

    if (action === 'connexion') return { ok: true, equipe: equipe.equipe, etablissement: equipe.etablissement };
    if (action === 'lire') return lireTout_(equipe.equipe);

    // Actions qui modifient : une seule à la fois, pour ne pas mélanger deux enregistrements
    var verrou = LockService.getScriptLock();
    verrou.waitLock(25000);
    try {
      if (action === 'enregistrerBlocs') return enregistrerBlocs_(equipe.equipe, d.blocs);
      if (action === 'supprimerBloc') return supprimerBloc_(equipe.equipe, String(d.id || ''));
      if (action === 'enregistrerTopo') return enregistrerTopo_(equipe.equipe, d.topo);
      if (action === 'enregistrerPhoto') return enregistrerPhoto_(equipe.equipe, String(d.nom || ''), String(d.dataUrl || ''));
      if (action === 'supprimerPhoto') return supprimerPhoto_(equipe.equipe, String(d.nom || ''));
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

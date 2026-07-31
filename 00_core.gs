// Café&Co — noyau commun

const CAFCO_MOIS = [
  'JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
  'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE'
];

const CAFCO_SLOTS = [
  { beneCol: 3, statutCol: 4, ticketCol: 5, periode: 'MATIN', nom: 'Matin 1' },
  { beneCol: 6, statutCol: 7, ticketCol: 8, periode: 'MATIN', nom: 'Matin 2' },
  { beneCol: 9, statutCol: 10, ticketCol: 11, periode: 'MATIN', nom: 'Matin 3' },
  { beneCol: 12, statutCol: 13, ticketCol: 14, periode: 'MATIN', nom: 'Matin 4' },
  { beneCol: 15, statutCol: 16, ticketCol: 17, periode: 'APRES_MIDI', nom: 'Après-midi 1' },
  { beneCol: 18, statutCol: 19, ticketCol: 20, periode: 'APRES_MIDI', nom: 'Après-midi 2' },
  { beneCol: 21, statutCol: 22, ticketCol: 23, periode: 'APRES_MIDI', nom: 'Après-midi 3' },
  { beneCol: 24, statutCol: 25, ticketCol: 26, periode: 'APRES_MIDI', nom: 'Après-midi 4' }
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const planning = ui.createMenu('Planning')
    .addItem('Créer ou réinitialiser un mois', 'genererMoisDepuisModele')
    .addItem('Copier la semaine type', 'appliquerSemaineTypeUneSemaine');

  const controle = ui.createMenu('Contrôle')
    .addItem('Vérifier le planning', 'verifierControleQualite')
    .addItem('Forcer le recalcul des tickets', 'recalculerTicketsMois');

  const depannage = ui.createMenu('Dépannage')
    .addItem('Réparer et recalculer les tickets', 'reparerFormulesTickets')
    .addItem('Protéger les colonnes Ticket', 'protegerTicketsMoisExistant')
    .addItem('Auditer la structure du classeur', 'auditerStructureClasseur');

  ui.createMenu('Café&Co')
    .addSubMenu(planning)
    .addSubMenu(controle)
    .addSubMenu(depannage)
    .addToUi();
}

function lireParametres_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('_PARAMETRES');
  if (!sh) throw new Error('Feuille _PARAMETRES introuvable.');

  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const cle = String(values[i][0] || '').trim();
    if (cle) map[cle] = values[i][1];
  }

  const postes = splitParam_(map['Postes']);
  const joursOuvres = splitParam_(map['Jours ouvrés']);
  if (!postes.length) throw new Error('Aucun poste défini dans _PARAMETRES.');
  if (!joursOuvres.length) throw new Error('Aucun jour ouvré défini dans _PARAMETRES.');

  return {
    anneeActive: Number(map['Année active']) || new Date().getFullYear(),
    modeleMois: String(map['Modèle mois'] || '_MODELE_MOIS').trim(),
    postes: postes,
    joursOuvres: joursOuvres,
    benevolesMaxMatin: Number(map['Bénévoles max matin']) || 4,
    benevolesMaxApresMidi: Number(map['Bénévoles max après-midi']) || 4
  };
}

function splitParam_(valeur) {
  if (!valeur) return [];
  return String(valeur).split(',').map(function(v) { return v.trim(); }).filter(Boolean);
}

function estFeuilleMois_(nom) {
  return CAFCO_MOIS.indexOf(String(nom || '').toUpperCase()) !== -1;
}

function normaliserOui_(valeur) {
  return String(valeur || '').trim().toLowerCase() === 'oui';
}

function cleJour_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getSemaineCle_(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
  const day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  return 'Semaine du ' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function journaliser_(action, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_CHANGELOG');
  if (!sheet) return;
  sheet.appendRow([new Date(), action, detail]);
}

function forcerRecalculSyntheses_() {
  // Google Sheets recalcule les formules automatiquement. Réécrire chaque formule
  // une par une est lent et risqué. Un flush suffit à pousser les écritures en attente.
  SpreadsheetApp.flush();
}

function nomFeuilleTicketsCalcul_(nomMois) {
  return '_TICKETS_' + String(nomMois).toUpperCase();
}

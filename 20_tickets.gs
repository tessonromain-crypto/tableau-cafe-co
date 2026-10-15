// Café&Co — calcul sécurisé et réactif des tickets

function calculerTickets_(planning, infosBenevoles, schema) {
  schema = schema || { dateCol: 1, totalCols: 26, slots: CAFCO_SLOTS };
  const resultat = planning.map(function(l) { return l.slice(); });
  const compteursSemaine = {};
  const creneauxDejaComptes = {};
  let oui = 0, non = 0, vides = 0, doublons = 0;

  for (let r = 0; r < resultat.length; r++) {
    const date = resultat[r][schema.dateCol - 1];
    if (!(date instanceof Date)) {
      schema.slots.forEach(function(slot) {
        resultat[r][slot.ticketCol - 1] = '';
        vides++;
      });
      continue;
    }

    const semaine = getSemaineCle_(date);
    const jour = cleJour_(date);

    schema.slots.forEach(function(slot) {
      const benevole = String(resultat[r][slot.beneCol - 1] || '').trim();
      const presence = String(resultat[r][slot.statutCol - 1] || '').trim();
      let ticket = '';

      if (!benevole || !presence) {
        ticket = '';
      } else if (presence !== 'Présent') {
        ticket = 'Non';
      } else {
        const infos = infosBenevoles[benevole];
        const eligible = estEligibleTicketADate_(infos, date);

        if (!eligible) {
          ticket = 'Non';
        } else {
          const cleCreneau = jour + '|' + slot.periode + '|' + benevole;
          const cleSemaine = semaine + '|' + benevole;
          if (creneauxDejaComptes[cleCreneau]) {
            ticket = 'Non';
            doublons++;
          } else if ((compteursSemaine[cleSemaine] || 0) >= 3) {
            ticket = 'Non';
          } else {
            ticket = 'Oui';
            creneauxDejaComptes[cleCreneau] = true;
            compteursSemaine[cleSemaine] = (compteursSemaine[cleSemaine] || 0) + 1;
          }
        }
      }

      resultat[r][slot.ticketCol - 1] = ticket;
      if (ticket === 'Oui') oui++;
      else if (ticket === 'Non') non++;
      else vides++;
    });
  }

  return { planning: resultat, stats: { oui: oui, non: non, vides: vides, doublons: doublons } };
}

function lireInfosBenevoles_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('BENEVOLES');
  if (!sh) throw new Error('Feuille BENEVOLES introuvable.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const data = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  const infos = {};
  data.forEach(function(l) {
    const nom = String(l[0] || '').trim();
    if (!nom) return;
    infos[nom] = {
      statutManuel: l[1],
      dateEntree: l[2],
      dateTitularisation: l[3],
      dateSortie: l[4],
      referent: l[5],
      souhaiteTicket: l[6],
      statutAutomatique: l[7],
      disponiblePlanning: l[8]
    };
  });
  return infos;
}

function normaliserDateTicket_(valeur) {
  if (!(valeur instanceof Date) || isNaN(valeur.getTime())) return null;
  return new Date(valeur.getFullYear(), valeur.getMonth(), valeur.getDate(), 12, 0, 0);
}

// Éligibilité historique : on juge la personne à la date du créneau,
// pas avec son seul statut actuel.
function estEligibleTicketADate_(infos, date) {
  if (!infos || !(date instanceof Date) || !normaliserOui_(infos.souhaiteTicket)) return false;

  const jour = normaliserDateTicket_(date);
  const entree = normaliserDateTicket_(infos.dateEntree);
  const titularisation = normaliserDateTicket_(infos.dateTitularisation);
  const sortie = normaliserDateTicket_(infos.dateSortie);
  if (!jour) return false;

  // La titularisation datée est la source prioritaire.
  if (titularisation) {
    if (jour.getTime() < titularisation.getTime()) return false;
  } else {
    // Compatibilité avec les anciennes fiches sans date exploitable :
    // on conserve la logique du statut actuel, sans autoriser une date avant l'entrée.
    if (infos.statutAutomatique !== 'Bénévole' && infos.statutAutomatique !== 'Référent') return false;
    if (entree && jour.getTime() < entree.getTime()) return false;
  }

  // La date de sortie est considérée comme le dernier jour encore éligible.
  if (sortie && jour.getTime() > sortie.getTime()) return false;

  return true;
}

function ecrireTicketsCalcules_(sheet, calcul, schema, rowOffset) {
  rowOffset = rowOffset || 2;
  schema.slots.forEach(function(slot) {
    const valeurs = calcul.planning.map(function(l) { return [l[slot.ticketCol - 1]]; });
    if (valeurs.length) sheet.getRange(rowOffset, slot.ticketCol, valeurs.length, 1).setValues(valeurs);
  });
}

function recalculerTicketsFeuille_(sheet, infosBenevoles) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { oui: 0, non: 0, vides: 0, doublons: 0 };

  const schema = schemaPlanning_(sheet);
  const planning = sheet.getRange(2, 1, lastRow - 1, schema.totalCols).getValues();
  const calcul = calculerTickets_(planning, infosBenevoles || lireInfosBenevoles_(), schema);
  ecrireTicketsCalcules_(sheet, calcul, schema, 2);
  SpreadsheetApp.flush();
  return calcul.stats;
}

// Recalcule uniquement les semaines demandées, sans parcourir tout le mois pour le calcul.
function recalculerTicketsSemaines_(sheet, semaines, infosBenevoles) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const cibles = {};
  (semaines || []).forEach(function(semaine) { if (semaine) cibles[semaine] = true; });
  if (!Object.keys(cibles).length) return null;

  const schema = schemaPlanning_(sheet);
  const infos = infosBenevoles || lireInfosBenevoles_();
  const dates = sheet.getRange(2, schema.dateCol, lastRow - 1, 1).getValues();
  const groupes = {};

  dates.forEach(function(ligne, index) {
    const date = ligne[0];
    if (!(date instanceof Date)) return;
    const semaine = getSemaineCle_(date);
    if (!cibles[semaine]) return;
    if (!groupes[semaine]) groupes[semaine] = { debut: index + 2, fin: index + 2 };
    groupes[semaine].fin = index + 2;
  });

  Object.keys(groupes).forEach(function(semaine) {
    const g = groupes[semaine];
    const nbLignes = g.fin - g.debut + 1;
    const planning = sheet.getRange(g.debut, 1, nbLignes, schema.totalCols).getValues();
    const calcul = calculerTickets_(planning, infos, schema);
    ecrireTicketsCalcules_(sheet, calcul, schema, g.debut);
  });

  SpreadsheetApp.flush();
  return true;
}

// Recalcul ultra-ciblé : uniquement la ou les personnes concernées sur la semaine modifiée.
function recalculerTicketsPersonnesSemaine_(sheet, ligneEditee, noms, ticketColAForcerVide) {
  if (!sheet || !estFeuilleMois_(sheet.getName()) || ligneEditee < 2) return;

  const schema = schemaPlanning_(sheet);
  const dateEditee = sheet.getRange(ligneEditee, schema.dateCol).getValue();
  if (!(dateEditee instanceof Date)) return;

  const nomsCibles = {};
  (noms || []).forEach(function(nom) {
    nom = String(nom || '').trim();
    if (nom) nomsCibles[nom] = true;
  });

  if (ticketColAForcerVide) sheet.getRange(ligneEditee, ticketColAForcerVide).clearContent();

  const listeNoms = Object.keys(nomsCibles);
  if (!listeNoms.length) return;

  const semaineCible = getSemaineCle_(dateEditee);
  const lastRow = sheet.getLastRow();
  const dates = sheet.getRange(2, schema.dateCol, lastRow - 1, 1).getValues();
  let premiereLigneSemaine = null;
  let derniereLigneSemaine = null;

  dates.forEach(function(ligne, index) {
    const date = ligne[0];
    if (!(date instanceof Date) || getSemaineCle_(date) !== semaineCible) return;
    const numeroLigne = index + 2;
    if (premiereLigneSemaine === null) premiereLigneSemaine = numeroLigne;
    derniereLigneSemaine = numeroLigne;
  });
  if (premiereLigneSemaine === null) return;

  const nbLignes = derniereLigneSemaine - premiereLigneSemaine + 1;
  const bloc = sheet.getRange(premiereLigneSemaine, 1, nbLignes, schema.totalCols).getValues();
  const infosBenevoles = lireInfosBenevoles_();
  const compteurs = {};
  const creneauxDejaComptes = {};
  const misesAJour = [];
  listeNoms.forEach(function(nom) { compteurs[nom] = 0; });

  for (let r = 0; r < bloc.length; r++) {
    const date = bloc[r][schema.dateCol - 1];
    if (!(date instanceof Date)) continue;
    const jour = cleJour_(date);

    schema.slots.forEach(function(slot) {
      const benevole = String(bloc[r][slot.beneCol - 1] || '').trim();
      if (!nomsCibles[benevole]) return;

      const presence = String(bloc[r][slot.statutCol - 1] || '').trim();
      let ticket = '';
      if (!presence) {
        ticket = '';
      } else if (presence !== 'Présent') {
        ticket = 'Non';
      } else {
        const infos = infosBenevoles[benevole];
        const eligible = estEligibleTicketADate_(infos, date);
        if (!eligible) {
          ticket = 'Non';
        } else {
          const cleCreneau = benevole + '|' + jour + '|' + slot.periode;
          if (creneauxDejaComptes[cleCreneau]) {
            ticket = 'Non';
          } else if ((compteurs[benevole] || 0) >= 3) {
            ticket = 'Non';
          } else {
            ticket = 'Oui';
            creneauxDejaComptes[cleCreneau] = true;
            compteurs[benevole] = (compteurs[benevole] || 0) + 1;
          }
        }
      }

      misesAJour.push({ row: premiereLigneSemaine + r, col: slot.ticketCol, value: ticket });
    });
  }

  misesAJour.forEach(function(item) {
    sheet.getRange(item.row, item.col).setValue(item.value);
  });
}

function recalculerTicketsMois() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt('Recalculer les tickets', 'Indique le mois, par exemple : JUILLET', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;

  const nomMois = rep.getResponseText().trim().toUpperCase();
  const sheet = ss.getSheetByName(nomMois);
  if (!sheet || !estFeuilleMois_(nomMois)) {
    ui.alert('Feuille mensuelle introuvable : ' + nomMois);
    return;
  }

  const stats = recalculerTicketsFeuille_(sheet);
  protegerFormulesTickets_(sheet);
  journaliser_('Recalcul tickets', nomMois + ' : ' + stats.oui + ' Oui, ' + stats.non + ' Non');
  ui.alert('Tickets recalculés pour ' + nomMois + '.');
}

function corrigerTicketsMaxSemaine() {
  recalculerTicketsMois();
}

// Nom conservé pour compatibilité avec les anciennes versions du menu.
// Cette fonction ne répare plus de formules : elle recalcule les valeurs Ticket de tous les mois.
function reparerFormulesTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const choix = ui.alert(
    'Recalculer les tickets de tous les mois',
    'Cette action recalculera uniquement les valeurs des colonnes Ticket. Les noms et présences ne seront pas modifiés. Continuer ?',
    ui.ButtonSet.YES_NO
  );
  if (choix !== ui.Button.YES) return;

  let feuilles = 0;
  CAFCO_MOIS.forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;
    recalculerTicketsFeuille_(sh);
    protegerFormulesTickets_(sh);
    feuilles++;
  });

  journaliser_('Recalcul tickets tous mois', feuilles + ' feuille(s)');
  ui.alert('Recalcul terminé.\n\nMois traités : ' + feuilles);
}

// Le nom historique est conservé pour éviter de casser les appels existants.
function protegerFormulesTickets_(sheet) {
  if (!sheet) throw new Error('Feuille à protéger introuvable.');
  const prefix = 'CAFCO_TICKETS_';
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const protectionsTickets = {};

  protections.forEach(function(p) {
    const description = p.getDescription() || '';
    if (description.indexOf(prefix) !== 0) return;
    if (!protectionsTickets[description]) protectionsTickets[description] = [];
    protectionsTickets[description].push(p);
  });

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const schema = schemaPlanning_(sheet);
  schema.slots.forEach(function(slot) {
    const range = sheet.getRange(2, slot.ticketCol, lastRow - 1, 1);
    const description = prefix + sheet.getName() + '_COL_' + slot.ticketCol;
    const existantes = protectionsTickets[description] || [];
    let protection = existantes.shift();

    if (!protection) protection = range.protect().setDescription(description);
    else if (protection.getRange().getA1Notation() !== range.getA1Notation()) protection.setRange(range);
    if (!protection.isWarningOnly()) protection.setWarningOnly(true);

    existantes.forEach(function(doublon) { if (doublon.canEdit()) doublon.remove(); });
    delete protectionsTickets[description];
  });

  Object.keys(protectionsTickets).forEach(function(description) {
    protectionsTickets[description].forEach(function(p) { if (p.canEdit()) p.remove(); });
  });
}

function protegerTicketsMoisExistant() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt('Protéger les tickets', 'Indique le mois, par exemple : JUILLET', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const nom = rep.getResponseText().trim().toUpperCase();
  const sh = ss.getSheetByName(nom);
  if (!sh || !estFeuilleMois_(nom)) {
    ui.alert('Feuille mensuelle introuvable : ' + nom);
    return;
  }
  protegerFormulesTickets_(sh);
  journaliser_('Protection tickets', nom + ' : colonnes Ticket protégées');
  ui.alert('Colonnes Ticket protégées contre les modifications accidentelles.');
}

// Recalcule uniquement les tickets des personnes demandées sur les mois existants.
// Chaque feuille n'est lue qu'une fois et seules les cellules réellement différentes sont écrites.
function recalculerTicketsPersonnesHistorique_(noms) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infos = lireInfosBenevoles_();
  const cibles = {};
  (noms || []).forEach(function(nom) {
    nom = String(nom || '').trim();
    if (nom) cibles[nom] = true;
  });
  const liste = Object.keys(cibles);
  if (!liste.length) return { personnes: 0, cellules: 0, mois: 0 };

  let cellules = 0;
  let mois = 0;

  CAFCO_MOIS.forEach(function(nomMois) {
    const sheet = ss.getSheetByName(nomMois);
    if (!sheet || sheet.getLastRow() < 2) return;

    const schema = schemaPlanning_(sheet);
    const nbLignes = sheet.getLastRow() - 1;
    const planning = sheet.getRange(2, 1, nbLignes, schema.totalCols).getValues();
    const calcul = calculerTickets_(planning, infos, schema);
    const adresses = { Oui: [], Non: [], Vide: [] };

    for (let r = 0; r < planning.length; r++) {
      schema.slots.forEach(function(slot) {
        const benevole = String(planning[r][slot.beneCol - 1] || '').trim();
        if (!cibles[benevole]) return;

        const actuel = String(planning[r][slot.ticketCol - 1] || '').trim();
        const attendu = String(calcul.planning[r][slot.ticketCol - 1] || '').trim();
        if (actuel === attendu) return;

        const adresse = colonneA1_(slot.ticketCol) + (r + 2);
        if (attendu === 'Oui') adresses.Oui.push(adresse);
        else if (attendu === 'Non') adresses.Non.push(adresse);
        else adresses.Vide.push(adresse);
      });
    }

    ecrireRangeListParLots_(sheet, adresses.Oui, 'Oui');
    ecrireRangeListParLots_(sheet, adresses.Non, 'Non');
    ecrireRangeListParLots_(sheet, adresses.Vide, '');
    const nbModifs = adresses.Oui.length + adresses.Non.length + adresses.Vide.length;
    if (nbModifs) {
      cellules += nbModifs;
      mois++;
    }
  });

  SpreadsheetApp.flush();
  return { personnes: liste.length, cellules: cellules, mois: mois };
}

function ecrireRangeListParLots_(sheet, adresses, valeur) {
  const tailleLot = 100;
  for (let i = 0; i < adresses.length; i += tailleLot) {
    sheet.getRangeList(adresses.slice(i, i + tailleLot)).setValue(valeur);
  }
}

function gererEditionBenevolesTickets_(e) {
  const range = e.range;
  if (range.getRow() < 2) return;

  // Colonnes pouvant modifier l'éligibilité historique :
  // B statut manuel, C entrée, D titularisation, E sortie,
  // F référent, G souhaite ticket, H statut automatique.
  const colDebut = range.getColumn();
  const colFin = range.getLastColumn();
  const toucheEligibilite = !(colFin < 2 || colDebut > 8);
  if (!toucheEligibilite) return;

  const sheet = range.getSheet();
  const ligneDebut = range.getRow();
  const ligneFin = range.getLastRow();
  const noms = sheet.getRange(ligneDebut, 1, ligneFin - ligneDebut + 1, 1).getDisplayValues()
    .map(function(ligne) { return String(ligne[0] || '').trim(); })
    .filter(Boolean);
  if (!noms.length) return;

  const resultat = recalculerTicketsPersonnesHistorique_(noms);
  journaliser_(
    'Recalcul ciblé après BENEVOLES',
    noms.join(', ') + ' : ' + resultat.cellules + ' ticket(s) modifié(s) sur ' + resultat.mois + ' mois'
  );
}

function recalculerTicketsBenevoleManuel() {
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt(
    'Recalculer un bénévole',
    'Indique exactement le nom du bénévole tel qu’il apparaît dans BENEVOLES.',
    ui.ButtonSet.OK_CANCEL
  );
  if (rep.getSelectedButton() !== ui.Button.OK) return;

  const nom = rep.getResponseText().trim();
  if (!nom) return;

  const resultat = recalculerTicketsPersonnesHistorique_([nom]);
  journaliser_('Recalcul bénévole manuel', nom + ' : ' + resultat.cellules + ' ticket(s) modifié(s)');
  ui.alert(nom + ' : ' + resultat.cellules + ' ticket(s) mis à jour sur ' + resultat.mois + ' mois.');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() === 'BENEVOLES') {
    gererEditionBenevolesTickets_(e);
    return;
  }

  if (!estFeuilleMois_(sheet.getName()) || range.getRow() < 2) return;

  const schema = schemaPlanning_(sheet);
  const colDebut = range.getColumn();
  const colFin = range.getLastColumn();
  const ligneDebut = range.getRow();
  const ligneFin = range.getLastRow();

  const slotsTouches = schema.slots.filter(function(slot) {
    return (slot.beneCol >= colDebut && slot.beneCol <= colFin) ||
           (slot.statutCol >= colDebut && slot.statutCol <= colFin);
  });
  if (!slotsTouches.length) return;

  // Cas courant : une seule cellule modifiée. On recalcule uniquement la personne et sa semaine.
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    const col = range.getColumn();
    const slot = slotsTouches[0];
    const noms = [];
    const nomActuel = String(sheet.getRange(ligneDebut, slot.beneCol).getValue() || '').trim();
    if (nomActuel) noms.push(nomActuel);

    let ticketAForcerVide = null;
    if (col === slot.beneCol) {
      const ancienNom = String(e.oldValue || '').trim();
      if (ancienNom && noms.indexOf(ancienNom) === -1) noms.push(ancienNom);
      ticketAForcerVide = slot.ticketCol;
    }

    recalculerTicketsPersonnesSemaine_(sheet, ligneDebut, noms, ticketAForcerVide);
    return;
  }

  // Collage ou modification multiple : on limite le recalcul aux semaines réellement touchées.
  const semaines = {};
  for (let row = ligneDebut; row <= ligneFin; row++) {
    const date = sheet.getRange(row, schema.dateCol).getValue();
    if (date instanceof Date) semaines[getSemaineCle_(date)] = true;
  }
  recalculerTicketsSemaines_(sheet, Object.keys(semaines));
}

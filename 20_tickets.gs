// Café&Co — calcul sécurisé et réactif des tickets

function calculerTickets_(planning, infosBenevoles, schema) {
  schema = schema || {
    dateCol: 1,
    totalCols: 26,
    slots: CAFCO_SLOTS
  };

  const resultat = planning.map(function(l) { return l.slice(); });
  const compteursSemaine = {};
  const creneauxDejaComptes = {};
  let oui = 0;
  let non = 0;
  let vides = 0;
  let doublons = 0;

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
        const eligible = infos && normaliserOui_(infos.souhaiteTicket) &&
          (infos.statutAutomatique === 'Bénévole' || infos.statutAutomatique === 'Référent');

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
    infos[nom] = { souhaiteTicket: l[6], statutAutomatique: l[7] };
  });
  return infos;
}

function ecrireTicketsCalcules_(sheet, calcul, schema) {
  schema.slots.forEach(function(slot) {
    const valeurs = calcul.planning.map(function(l) {
      return [l[slot.ticketCol - 1]];
    });
    sheet.getRange(2, slot.ticketCol, valeurs.length, 1).setValues(valeurs);
  });
}

function recalculerTicketsFeuille_(sheet, infosBenevoles) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { oui: 0, non: 0, vides: 0, doublons: 0 };

  const schema = schemaPlanning_(sheet);
  const planning = sheet.getRange(2, 1, lastRow - 1, schema.totalCols).getValues();
  const calcul = calculerTickets_(planning, infosBenevoles || lireInfosBenevoles_(), schema);
  ecrireTicketsCalcules_(sheet, calcul, schema);
  SpreadsheetApp.flush();
  return calcul.stats;
}

function recalculerTicketsSemaines_(sheet, semaines, infosBenevoles) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { oui: 0, non: 0, vides: 0, doublons: 0 };

  const semainesCibles = {};
  (semaines || []).forEach(function(semaine) {
    if (semaine) semainesCibles[semaine] = true;
  });
  if (!Object.keys(semainesCibles).length) return null;

  const schema = schemaPlanning_(sheet);
  const planning = sheet.getRange(2, 1, lastRow - 1, schema.totalCols).getValues();
  const calcul = calculerTickets_(planning, infosBenevoles || lireInfosBenevoles_(), schema);

  const lignesCibles = [];
  planning.forEach(function(ligne, index) {
    const date = ligne[schema.dateCol - 1];
    if (date instanceof Date && semainesCibles[getSemaineCle_(date)]) {
      lignesCibles.push(index);
    }
  });
  if (!lignesCibles.length) return calcul.stats;

  const groupes = [];
  let debut = lignesCibles[0];
  let precedent = debut;
  for (let i = 1; i < lignesCibles.length; i++) {
    const courant = lignesCibles[i];
    if (courant === precedent + 1) {
      precedent = courant;
      continue;
    }
    groupes.push({ debut: debut, fin: precedent });
    debut = courant;
    precedent = courant;
  }
  groupes.push({ debut: debut, fin: precedent });

  schema.slots.forEach(function(slot) {
    groupes.forEach(function(groupe) {
      const valeurs = [];
      for (let index = groupe.debut; index <= groupe.fin; index++) {
        valeurs.push([calcul.planning[index][slot.ticketCol - 1]]);
      }
      sheet.getRange(groupe.debut + 2, slot.ticketCol, valeurs.length, 1).setValues(valeurs);
    });
  });

  SpreadsheetApp.flush();
  return calcul.stats;
}

// Recalcul ultra-ciblé : uniquement la ou les personnes concernées,
// sur la semaine de la ligne modifiée. Aucun recalcul du mois entier.
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

  // Si un nom vient d'être supprimé/remplacé, son ancien Ticket doit disparaître immédiatement.
  if (ticketColAForcerVide) {
    sheet.getRange(ligneEditee, ticketColAForcerVide).clearContent();
  }

  const listeNoms = Object.keys(nomsCibles);
  if (!listeNoms.length) return;

  const semaineCible = getSemaineCle_(dateEditee);
  const lastRow = sheet.getLastRow();

  // Une seule lecture de la colonne Date pour trouver les lignes de la semaine.
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

  // Une seule lecture du petit bloc de la semaine (en général ~20 lignes).
  const nbLignes = derniereLigneSemaine - premiereLigneSemaine + 1;
  const bloc = sheet.getRange(
    premiereLigneSemaine,
    1,
    nbLignes,
    schema.totalCols
  ).getValues();

  const infosBenevoles = lireInfosBenevoles_();
  const compteurs = {};
  const creneauxDejaComptes = {};
  const misesAJour = [];

  listeNoms.forEach(function(nom) {
    compteurs[nom] = 0;
  });

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
        const eligible = infos && normaliserOui_(infos.souhaiteTicket) &&
          (infos.statutAutomatique === 'Bénévole' || infos.statutAutomatique === 'Référent');

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

      misesAJour.push({
        row: premiereLigneSemaine + r,
        col: slot.ticketCol,
        value: ticket
      });
    });
  }

  // Seulement les cellules Ticket des personnes concernées sont écrites.
  misesAJour.forEach(function(item) {
    sheet.getRange(item.row, item.col).setValue(item.value);
  });
}

// Compatibilité avec le reste du projet : les Ticket sont maintenant des valeurs
// calculées par le script, et non des formules reliées à une feuille cachée.
function colA1_(col) {
  let s = '';
  while (col > 0) {
    const r = (col - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

function formuleTicketNative_(row, slotIndex, schema, lastRow) {
  const slot = schema.slots[slotIndex];
  const dateCell = colA1_(schema.dateCol) + row;
  const beneCell = colA1_(slot.beneCol) + row;
  const statutCell = colA1_(slot.statutCol) + row;
  const dateRange = '$' + colA1_(schema.dateCol) + '$2:$' + colA1_(schema.dateCol) + '$' + lastRow;
  const debutSemaine = '(' + dateCell + '-WEEKDAY(' + dateCell + ',2)+1)';

  // Nombre de tickets Oui déjà attribués avant cette cellule dans la même semaine
  // pour la même personne. On ne regarde que les cellules antérieures afin d'éviter
  // toute référence circulaire.
  const compteOui = [];
  schema.slots.forEach(function(s, i) {
    const beneCol = '$' + colA1_(s.beneCol);
    const ticketCol = '$' + colA1_(s.ticketCol);
    let fin = row - 1;
    if (i < slotIndex) fin = row;
    if (fin < 2) return;
    compteOui.push(
      'COUNTIFS(' + dateRange + ',">="&' + debutSemaine + ',' +
      dateRange + ',"<="&(' + debutSemaine + '+6),' +
      beneCol + '$2:' + beneCol + '$' + fin + ',' + beneCell + ',' +
      ticketCol + '$2:' + ticketCol + '$' + fin + ',"Oui")'
    );
  });
  const nbOuiAvant = compteOui.length ? '(' + compteOui.join('+') + ')' : '0';

  // Doublon = même bénévole, même date et même demi-journée déjà rencontré
  // avant cette cellule avec le statut Présent.
  const doublons = [];
  schema.slots.forEach(function(s, i) {
    if (s.periode !== slot.periode) return;
    let fin = row - 1;
    if (i < slotIndex) fin = row;
    if (fin < 2) return;
    const beneCol = '$' + colA1_(s.beneCol);
    const statutCol = '$' + colA1_(s.statutCol);
    doublons.push(
      'COUNTIFS(' + dateRange + ',' + dateCell + ',' +
      beneCol + '$2:' + beneCol + '$' + fin + ',' + beneCell + ',' +
      statutCol + '$2:' + statutCol + '$' + fin + ',"Présent")'
    );
  });
  const nbDoublons = doublons.length ? '(' + doublons.join('+') + ')' : '0';

  return '=IF(OR(' + dateCell + '="",' + beneCell + '="",' + statutCell + '=""),"",'+
    'IF(' + statutCell + '<>"Présent","Non",'+
    'IF(IFERROR(VLOOKUP(' + beneCell + ',BENEVOLES!$A:$I,7,FALSE),"")<>"Oui","Non",'+
    'IF(AND(IFERROR(VLOOKUP(' + beneCell + ',BENEVOLES!$A:$I,8,FALSE),"")<>"Bénévole",'+
    'IFERROR(VLOOKUP(' + beneCell + ',BENEVOLES!$A:$I,8,FALSE),"")<>"Référent"),"Non",'+
    'IF(' + nbDoublons + '>0,"Non",IF(' + nbOuiAvant + '>=3,"Non","Oui"))))))';
}

// Installe les formules natives Google Sheets. Ensuite les Tickets se recalculent
// automatiquement sans onEdit Apps Script.
function installerFormulesTicketsPourFeuille_(sheet) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return 0;
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const schema = schemaPlanning_(sheet);
  let cellulesPreparees = 0;

  schema.slots.forEach(function(slot, slotIndex) {
    const formules = [];
    for (let row = 2; row <= lastRow; row++) {
      formules.push([formuleTicketNative_(row, slotIndex, schema, lastRow)]);
    }
    sheet.getRange(2, slot.ticketCol, formules.length, 1).setFormulas(formules);
    cellulesPreparees += formules.length;
  });
  return cellulesPreparees;
}

function recalculerTicketsMois() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt('Réinstaller les formules Ticket', 'Indique le mois, par exemple : JUILLET', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;

  const nomMois = rep.getResponseText().trim().toUpperCase();
  const sheet = ss.getSheetByName(nomMois);
  if (!sheet || !estFeuilleMois_(nomMois)) {
    ui.alert('Feuille mensuelle introuvable : ' + nomMois);
    return;
  }

  const nb = installerFormulesTicketsPourFeuille_(sheet);
  protegerFormulesTickets_(sheet);
  SpreadsheetApp.flush();
  journaliser_('Formules tickets', nomMois + ' : ' + nb + ' formule(s) installée(s)');
  ui.alert('Formules Ticket réinstallées.\\n\\nCellules préparées : ' + nb);
}

function corrigerTicketsMaxSemaine() {
  recalculerTicketsMois();
}

function reparerFormulesTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const choix = ui.alert(
    'Réparer les formules Ticket de tous les mois',
    'Cette action réinstallera uniquement les formules des colonnes Ticket et leurs protections. Les noms, présences et autres données ne seront pas modifiés. Continuer ?',
    ui.ButtonSet.YES_NO
  );
  if (choix !== ui.Button.YES) return;

  let feuilles = 0;
  let cellules = 0;
  CAFCO_MOIS.forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;
    cellules += installerFormulesTicketsPourFeuille_(sh);
    protegerFormulesTickets_(sh);
    feuilles++;
  });

  SpreadsheetApp.flush();
  journaliser_('Réparation formules tickets', feuilles + ' feuille(s), ' + cellules + ' formule(s)');
  ui.alert('Réparation terminée.\\n\\nMois traités : ' + feuilles + '\\nFormules installées : ' + cellules);
}

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

    if (!protection) {
      protection = range.protect().setDescription(description);
    } else if (protection.getRange().getA1Notation() !== range.getA1Notation()) {
      protection.setRange(range);
    }
    if (!protection.isWarningOnly()) protection.setWarningOnly(true);

    existantes.forEach(function(doublon) {
      if (doublon.canEdit()) doublon.remove();
    });
    delete protectionsTickets[description];
  });

  Object.keys(protectionsTickets).forEach(function(description) {
    protectionsTickets[description].forEach(function(p) {
      if (p.canEdit()) p.remove();
    });
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

function onEdit(e) {
  // Les Tickets sont désormais calculés par des formules natives Google Sheets.
  // Aucun recalcul Apps Script n'est nécessaire lors de la saisie quotidienne.
  return;
}

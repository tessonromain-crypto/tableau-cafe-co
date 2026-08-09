// Café&Co — calcul sécurisé des tickets

function obtenirFeuilleTicketsCalcul_(nomMois) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nom = nomFeuilleTicketsCalcul_(nomMois);
  let sh = ss.getSheetByName(nom);
  if (!sh) {
    sh = ss.insertSheet(nom);
    sh.hideSheet();
  }
  if (!sh.isSheetHidden()) sh.hideSheet();
  return sh;
}

function calculerTickets_(planning, infosBenevoles) {
  const resultat = planning.map(function(l) { return l.slice(); });
  const compteursSemaine = {};
  const creneauxDejaComptes = {};
  let oui = 0;
  let non = 0;
  let vides = 0;
  let doublons = 0;

  for (let r = 0; r < resultat.length; r++) {
    const date = resultat[r][0];
    if (!(date instanceof Date)) {
      CAFCO_SLOTS.forEach(function(slot) {
        resultat[r][slot.ticketCol - 1] = '';
        vides++;
      });
      continue;
    }

    const semaine = getSemaineCle_(date);
    const jour = cleJour_(date);

    CAFCO_SLOTS.forEach(function(slot) {
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

function recalculerTicketsFeuille_(sheet, infosBenevoles) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { oui: 0, non: 0, vides: 0, doublons: 0 };

  const planning = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
  const calcul = calculerTickets_(planning, infosBenevoles || lireInfosBenevoles_());
  const ledger = obtenirFeuilleTicketsCalcul_(sheet.getName());

  if (ledger.getMaxRows() < lastRow) ledger.insertRowsAfter(ledger.getMaxRows(), lastRow - ledger.getMaxRows());
  if (ledger.getMaxColumns() < 26) ledger.insertColumnsAfter(ledger.getMaxColumns(), 26 - ledger.getMaxColumns());
  ledger.clearContents();

  CAFCO_SLOTS.forEach(function(slot) {
    const valeurs = calcul.planning.map(function(l) { return [l[slot.ticketCol - 1]]; });
    ledger.getRange(2, slot.ticketCol, valeurs.length, 1).setValues(valeurs);
  });

  SpreadsheetApp.flush();
  return calcul.stats;
}

function installerFormulesTicketsPourFeuille_(sheet) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return 0;
  const lastRow = Math.max(sheet.getLastRow(), 2);
  obtenirFeuilleTicketsCalcul_(sheet.getName());
  let nbFormules = 0;

  CAFCO_SLOTS.forEach(function(slot) {
    const formules = [];
    const nomLedger = nomFeuilleTicketsCalcul_(sheet.getName()).replace(/'/g, "''");
    for (let r = 2; r <= lastRow; r++) {
      formules.push(["=IFERROR('" + nomLedger + "'!" + sheet.getRange(r, slot.ticketCol).getA1Notation() + ';"")']);
    }
    sheet.getRange(2, slot.ticketCol, formules.length, 1).setFormulas(formules);
    nbFormules += formules.length;
  });
  return nbFormules;
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

  installerFormulesTicketsPourFeuille_(sheet);
  const stats = recalculerTicketsFeuille_(sheet);
  executerControleQualiteGuide_(false);
  journaliser_('Recalcul tickets', nomMois + ' : ' + stats.oui + ' Oui, ' + stats.non + ' Non, ' + stats.doublons + ' doublon(s) neutralisé(s)');
  ui.alert('Tickets recalculés.\n\nOui : ' + stats.oui + '\nNon : ' + stats.non + '\nVides : ' + stats.vides + '\nDoublons neutralisés : ' + stats.doublons);
}

function corrigerTicketsMaxSemaine() {
  // Compatibilité avec d'anciens déclencheurs éventuels.
  // Cette fonction ne doit plus apparaître dans le menu.
  recalculerTicketsMois();
}

function reparerFormulesTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const choix = ui.alert(
    'Réparer les tickets de tous les mois',
    'Cette action corrigera les en-têtes Ticket, réinstallera les formules, recalculera les tickets et remettra les protections sur tous les mois existants. Continuer ?',
    ui.ButtonSet.YES_NO
  );
  if (choix !== ui.Button.YES) return;

  let feuilles = 0;
  let formules = 0;
  let entetesCorriges = 0;
  const infosBenevoles = lireInfosBenevoles_();

  CAFCO_MOIS.forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;

    if (sh.getMaxColumns() < 26) {
      sh.insertColumnsAfter(sh.getMaxColumns(), 26 - sh.getMaxColumns());
    }

    CAFCO_SLOTS.forEach(function(slot) {
      const celluleEntete = sh.getRange(1, slot.ticketCol);
      if (celluleEntete.getDisplayValue() !== 'Ticket') {
        celluleEntete.setValue('Ticket');
        entetesCorriges++;
      }
    });

    formules += installerFormulesTicketsPourFeuille_(sh);
    recalculerTicketsFeuille_(sh, infosBenevoles);
    protegerFormulesTickets_(sh);
    feuilles++;
  });

  executerControleQualiteGuide_(false);
  SpreadsheetApp.flush();
  journaliser_(
    'Réparation formules Ticket',
    feuilles + ' feuille(s), ' + formules + ' formule(s), ' + entetesCorriges + ' en-tête(s) corrigé(s)'
  );
  ui.alert(
    'Réparation terminée.\n\n' +
    'Mois traités : ' + feuilles + '\n' +
    'Formules réinstallées : ' + formules + '\n' +
    'En-têtes corrigés : ' + entetesCorriges
  );
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
  CAFCO_SLOTS.forEach(function(slot) {
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
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const nom = sh.getName();

  if (estFeuilleMois_(nom)) {
    const col = e.range.getColumn();
    const concernePlanning = CAFCO_SLOTS.some(function(slot) {
      return col === slot.beneCol || col === slot.statutCol;
    });
    if (concernePlanning) {
      recalculerTicketsFeuille_(sh);
      executerControleQualiteGuide_(false);
    }
    return;
  }

  if (nom === 'BENEVOLES' && e.range.getRow() >= 2 && e.range.getColumn() <= 8) {
    CAFCO_MOIS.forEach(function(mois) {
      const moisSheet = e.source.getSheetByName(mois);
      if (moisSheet) recalculerTicketsFeuille_(moisSheet);
    });
    executerControleQualiteGuide_(false);
  }
}

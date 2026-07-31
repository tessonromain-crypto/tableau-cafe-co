// Café&Co — contrôle qualité et maintenance

function verifierControleQualite() {
  verifierControleQualite_(true);
}

function verifierControleQualite_(afficherAlerte) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controle = ss.getSheetByName('CONTROLE_QUALITE');
  if (!controle) {
    if (afficherAlerte) SpreadsheetApp.getUi().alert('Feuille CONTROLE_QUALITE introuvable.');
    return [];
  }

  const resultats = [];
  const ticketsParSemaine = {};
  const creneaux = {};

  CAFCO_MOIS.forEach(function(nomMois) {
    const sh = ss.getSheetByName(nomMois);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 26).getValues();

    data.forEach(function(ligne) {
      const date = ligne[0];
      const poste = String(ligne[1] || '').trim();
      if (!(date instanceof Date) || !poste) return;

      CAFCO_SLOTS.forEach(function(slot) {
        const benevole = String(ligne[slot.beneCol - 1] || '').trim();
        const statut = String(ligne[slot.statutCol - 1] || '').trim();
        const ticket = String(ligne[slot.ticketCol - 1] || '').trim();
        if (!benevole) return;

        if (!statut) {
          resultats.push([date, poste, slot.nom, benevole, 'Présence à valider']);
        }

        if (statut === 'Présent') {
          const cleCreneau = cleJour_(date) + '|' + slot.periode + '|' + benevole;
          if (!creneaux[cleCreneau]) creneaux[cleCreneau] = [];
          creneaux[cleCreneau].push(poste + ' / ' + slot.nom);
        }

        if (ticket === 'Oui') {
          const cleSemaine = getSemaineCle_(date) + '|' + benevole;
          if (!ticketsParSemaine[cleSemaine]) {
            ticketsParSemaine[cleSemaine] = { benevole: benevole, semaine: getSemaineCle_(date), dates: [], count: 0 };
          }
          ticketsParSemaine[cleSemaine].count++;
          ticketsParSemaine[cleSemaine].dates.push(date);
        }
      });
    });
  });

  Object.keys(creneaux).forEach(function(cle) {
    const affectations = creneaux[cle];
    if (affectations.length <= 1) return;
    const parties = cle.split('|');
    resultats.push([
      new Date(parties[0] + 'T12:00:00'),
      'Plusieurs postes',
      parties[1],
      parties.slice(2).join('|'),
      'Doublon même demi-journée : ' + affectations.length + ' affectations'
    ]);
  });

  Object.keys(ticketsParSemaine).forEach(function(cle) {
    const item = ticketsParSemaine[cle];
    if (item.count > 3) {
      resultats.push([item.dates[0], 'Tous postes', item.semaine, item.benevole, 'Dépassement 3 tickets/semaine : ' + item.count]);
    }
  });

  const rowsToClear = Math.max(controle.getMaxRows() - 1, 1);
  controle.getRange(2, 1, rowsToClear, 5).clearContent();
  if (resultats.length) {
    if (controle.getMaxRows() < resultats.length + 1) {
      controle.insertRowsAfter(controle.getMaxRows(), resultats.length + 1 - controle.getMaxRows());
    }
    controle.getRange(2, 1, resultats.length, 5).setValues(resultats);
  }

  journaliser_('Contrôle qualité', resultats.length + ' anomalie(s) détectée(s)');
  if (afficherAlerte) SpreadsheetApp.getUi().alert('Contrôle qualité terminé : ' + resultats.length + ' anomalie(s).');
  return resultats;
}

function sauvegarderFormulesReference() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const choix = ui.alert('Sauvegarder les formules', 'La sauvegarde actuelle sera remplacée. Continuer ?', ui.ButtonSet.YES_NO);
  if (choix !== ui.Button.YES) return;

  let backup = ss.getSheetByName('_FORMULES_BACKUP');
  if (!backup) backup = ss.insertSheet('_FORMULES_BACKUP');
  backup.clearContents();

  const exclues = ['_FORMULES_BACKUP', '_CHANGELOG', '_DOC_SCRIPT', '_DOC_FORMULES', '_PARAMETRES'];
  const sauvegarde = [['Feuille', 'Cellule', 'Formule']];

  ss.getSheets().forEach(function(sh) {
    const nom = sh.getName();
    if (exclues.indexOf(nom) !== -1 || nom.indexOf('_TICKETS_') === 0 || nom.indexOf('_BACKUP_') === 0) return;
    const range = sh.getDataRange();
    const formulas = range.getFormulas();
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c]) sauvegarde.push([nom, range.getCell(r + 1, c + 1).getA1Notation(), formulas[r][c]]);
      }
    }
  });

  if (backup.getMaxRows() < sauvegarde.length) backup.insertRowsAfter(backup.getMaxRows(), sauvegarde.length - backup.getMaxRows());
  backup.getRange(1, 1, sauvegarde.length, 3).setValues(sauvegarde);
  journaliser_('Sauvegarde formules', (sauvegarde.length - 1) + ' formule(s) sauvegardée(s)');
  ui.alert((sauvegarde.length - 1) + ' formule(s) sauvegardée(s).');
}

function reparerToutesLesFormules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const backup = ss.getSheetByName('_FORMULES_BACKUP');
  if (!backup || backup.getLastRow() < 2) {
    ui.alert('Aucune sauvegarde de formules disponible.');
    return;
  }

  const choix = ui.alert('Réparer toutes les formules', 'Les formules sauvegardées seront restaurées. Continuer ?', ui.ButtonSet.YES_NO);
  if (choix !== ui.Button.YES) return;

  const data = backup.getRange(2, 1, backup.getLastRow() - 1, 3).getValues();
  const parFeuille = {};
  data.forEach(function(l) {
    const nom = String(l[0] || '');
    if (!nom || !l[1] || !l[2]) return;
    if (!parFeuille[nom]) parFeuille[nom] = [];
    parFeuille[nom].push({ cellule: l[1], formule: l[2] });
  });

  let restaurees = 0;
  let erreurs = 0;
  Object.keys(parFeuille).forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (!sh) {
      erreurs += parFeuille[nom].length;
      return;
    }
    parFeuille[nom].forEach(function(item) {
      try {
        sh.getRange(item.cellule).setFormula(item.formule);
        restaurees++;
      } catch (err) {
        erreurs++;
      }
    });
  });

  // Les colonnes Ticket utilisent désormais un calcul sécurisé séparé.
  CAFCO_MOIS.forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (sh) {
      installerFormulesTicketsPourFeuille_(sh);
      recalculerTicketsFeuille_(sh);
    }
  });

  SpreadsheetApp.flush();
  journaliser_('Réparation formules', restaurees + ' restaurée(s), ' + erreurs + ' erreur(s)');
  ui.alert('Réparation terminée.\n\nFormules restaurées : ' + restaurees + '\nErreurs : ' + erreurs);
}

function auditerStructureClasseur() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const anomalies = [];
  ['_PARAMETRES', '_MODELE_MOIS', 'BENEVOLES', 'SEMAINE_TYPE'].forEach(function(nom) {
    if (!ss.getSheetByName(nom)) anomalies.push('Feuille obligatoire absente : ' + nom);
  });

  const modele = ss.getSheetByName('_MODELE_MOIS');
  if (modele) {
    CAFCO_SLOTS.forEach(function(slot) {
      const entete = modele.getRange(1, slot.ticketCol).getDisplayValue();
      if (entete !== 'Ticket') anomalies.push('_MODELE_MOIS : en-tête incorrect en ' + modele.getRange(1, slot.ticketCol).getA1Notation());
    });
  }

  CAFCO_MOIS.forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;
    if (sh.getMaxColumns() < 26) anomalies.push(nom + ' : moins de 26 colonnes');
    CAFCO_SLOTS.forEach(function(slot) {
      if (sh.getRange(1, slot.ticketCol).getDisplayValue() !== 'Ticket') anomalies.push(nom + ' : en-tête Ticket incorrect colonne ' + slot.ticketCol);
    });
  });

  ui.alert(anomalies.length ? 'Audit : ' + anomalies.length + ' anomalie(s)\n\n' + anomalies.join('\n') : 'Audit structure : aucun problème détecté.');
  return anomalies;
}

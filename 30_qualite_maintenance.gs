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

function estFeuilleExclueSauvegarde_(nom) {
  const exactes = ['_CHANGELOG', '_DOC_SCRIPT', '_DOC_FORMULES', '_PARAMETRES'];
  return exactes.indexOf(nom) !== -1 ||
    nom.indexOf('_TICKETS_') === 0 ||
    nom.indexOf('_BACKUP_') === 0 ||
    nom.indexOf('_FORMULES_BACKUP') === 0;
}

function contientErreurFormule_(valeurAffichee) {
  const texte = String(valeurAffichee || '').toUpperCase();
  const marqueurs = [
    '#REF!', '#N/A', '#VALUE!', '#VALEUR!', '#DIV/0!', '#NAME?', '#NOM?',
    '#NUM!', '#NOMBRE!', '#NULL!', '#ERREUR!', '#ERROR!'
  ];
  return marqueurs.some(function(marqueur) { return texte.indexOf(marqueur) !== -1; });
}

function sauvegarderFormulesReference() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sauvegarde = [['Feuille', 'Cellule', 'Formule']];
  const erreurs = [];

  ss.getSheets().forEach(function(sh) {
    const nom = sh.getName();
    if (estFeuilleExclueSauvegarde_(nom)) return;

    const range = sh.getDataRange();
    const formulas = range.getFormulas();
    const displays = range.getDisplayValues();

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formule = formulas[r][c];
        if (!formule) continue;
        const cellule = range.getCell(r + 1, c + 1).getA1Notation();
        if (contientErreurFormule_(displays[r][c])) {
          erreurs.push(nom + '!' + cellule + ' = ' + displays[r][c]);
        }
        sauvegarde.push([nom, cellule, formule]);
      }
    }
  });

  if (erreurs.length) {
    const apercu = erreurs.slice(0, 15).join('\n');
    const suite = erreurs.length > 15 ? '\n… et ' + (erreurs.length - 15) + ' autre(s).' : '';
    journaliser_('Sauvegarde formules refusée', erreurs.length + ' formule(s) en erreur');
    throw new Error(
      'Sauvegarde refusée : le classeur contient ' + erreurs.length +
      ' formule(s) en erreur. Corrige-les avant de créer une sauvegarde.\n\n' + apercu + suite
    );
  }

  const base = '_FORMULES_BACKUP_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  let nomBackup = base;
  let suffixe = 2;
  while (ss.getSheetByName(nomBackup)) {
    nomBackup = base + '_' + suffixe;
    suffixe++;
  }

  const backup = ss.insertSheet(nomBackup);
  if (backup.getMaxRows() < sauvegarde.length) {
    backup.insertRowsAfter(backup.getMaxRows(), sauvegarde.length - backup.getMaxRows());
  }
  backup.getRange(1, 1, sauvegarde.length, 3).setValues(sauvegarde);
  backup.hideSheet();

  journaliser_('Sauvegarde formules', nomBackup + ' : ' + (sauvegarde.length - 1) + ' formule(s)');
  Logger.log('Sauvegarde créée : ' + nomBackup + ' — ' + (sauvegarde.length - 1) + ' formule(s).');
  return { feuille: nomBackup, formules: sauvegarde.length - 1 };
}

function reparerToutesLesFormules() {
  journaliser_('Réparation globale bloquée', 'Fonction désactivée pour éviter les #REF!');
  SpreadsheetApp.getUi().alert(
    'Fonction désactivée',
    'La restauration globale des formules est désactivée pour sécurité. Les tests ont montré qu’elle pouvait générer des erreurs et des #REF!.\n\nUtilise uniquement « Réparer les formules Ticket de tous les mois » pour les colonnes Ticket.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
    if (sh.getMaxColumns() < 26) {
      anomalies.push(nom + ' : moins de 26 colonnes');
      return;
    }
    CAFCO_SLOTS.forEach(function(slot) {
      if (sh.getRange(1, slot.ticketCol).getDisplayValue() !== 'Ticket') {
        anomalies.push(nom + ' : en-tête Ticket incorrect colonne ' + slot.ticketCol);
      }
    });
  });

  journaliser_('Audit structure', anomalies.length + ' anomalie(s)');
  ui.alert(anomalies.length ? 'Audit : ' + anomalies.length + ' anomalie(s)\n\n' + anomalies.join('\n') : 'Audit structure : aucun problème détecté.');
  return anomalies;
}

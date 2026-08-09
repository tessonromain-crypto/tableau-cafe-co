// Café&Co — contrôle qualité et maintenance

function lancerControleQualiteGuide() {
  executerControleQualiteGuide_(true);
}

function executerControleQualiteGuide_(afficherAlerte) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controle = ss.getSheetByName('CONTROLE_QUALITE');
  if (!controle) {
    if (afficherAlerte) SpreadsheetApp.getUi().alert('Feuille CONTROLE_QUALITE introuvable.');
    return [];
  }

  const resultats = [];
  const ticketsParSemaine = {};
  const creneaux = {};
  const benevolesConnus = chargerNomsBenevoles_(ss);

  CAFCO_MOIS.forEach(function(nomMois) {
    const sh = ss.getSheetByName(nomMois);
    if (!sh || sh.getLastRow() < 2) return;
    const nbLignes = sh.getLastRow() - 1;
    const plage = sh.getRange(2, 1, nbLignes, 26);
    const data = plage.getValues();
    const formules = plage.getFormulas();

    data.forEach(function(ligne, indexLigne) {
      const date = ligne[0];
      const poste = String(ligne[1] || '').trim();
      if (!(date instanceof Date) || !poste) return;
      const numeroLigne = indexLigne + 2;

      CAFCO_SLOTS.forEach(function(slot) {
        const benevole = String(ligne[slot.beneCol - 1] || '').trim();
        const statut = String(ligne[slot.statutCol - 1] || '').trim();
        const ticket = String(ligne[slot.ticketCol - 1] || '').trim();
        const celluleBenevole = colonneA1_(slot.beneCol) + numeroLigne;
        const celluleStatut = colonneA1_(slot.statutCol) + numeroLigne;
        const celluleTicket = colonneA1_(slot.ticketCol) + numeroLigne;
        const formuleTicket = formules[indexLigne][slot.ticketCol - 1];

        if (!formuleTicket) {
          ajouterControleGuide_(resultats, 'Erreur', sh, celluleTicket, date, poste, slot.nom, benevole,
            'Formule Ticket absente ou remplacée',
            'Utiliser Maintenance et tests > Réparer les formules Ticket de tous les mois.');
        }
        if (!benevole) return;

        if (!benevolesConnus[benevole]) {
          ajouterControleGuide_(resultats, 'Erreur', sh, celluleBenevole, date, poste, slot.nom, benevole,
            'Bénévole absent de la feuille BENEVOLES',
            'Ajouter cette personne dans BENEVOLES ou corriger le nom sélectionné.');
        }

        if (!statut) {
          ajouterControleGuide_(resultats, 'Avertissement', sh, celluleStatut, date, poste, slot.nom, benevole,
            'Présence à valider',
            'Choisir Présent, Absent ou Retard dans cette cellule.');
        }

        if (ticket === 'Oui' && statut !== 'Présent') {
          ajouterControleGuide_(resultats, 'Erreur', sh, celluleTicket, date, poste, slot.nom, benevole,
            'Ticket accordé sans statut Présent',
            'Vérifier le statut puis recalculer les tickets.');
        }

        if (statut === 'Présent') {
          const cleCreneau = cleJour_(date) + '|' + slot.periode + '|' + benevole;
          if (!creneaux[cleCreneau]) creneaux[cleCreneau] = [];
          creneaux[cleCreneau].push({
            feuille: sh,
            cellule: celluleStatut,
            date: date,
            poste: poste,
            creneau: slot.nom
          });
        }

        if (ticket === 'Oui') {
          const cleSemaine = getSemaineCle_(date) + '|' + benevole;
          if (!ticketsParSemaine[cleSemaine]) {
            ticketsParSemaine[cleSemaine] = { benevole: benevole, semaine: getSemaineCle_(date), tickets: [], count: 0 };
          }
          ticketsParSemaine[cleSemaine].count++;
          ticketsParSemaine[cleSemaine].tickets.push({
            feuille: sh,
            cellule: celluleTicket,
            date: date,
            poste: poste,
            creneau: slot.nom
          });
        }
      });
    });
  });

  Object.keys(creneaux).forEach(function(cle) {
    const affectations = creneaux[cle];
    if (affectations.length <= 1) return;
    const parties = cle.split('|');
    const premiere = affectations[0];
    ajouterControleGuide_(resultats, 'Erreur', premiere.feuille, premiere.cellule, premiere.date,
      'Plusieurs postes', parties[1], parties.slice(2).join('|'),
      'Doublon sur la même demi-journée : ' + affectations.length + ' affectations',
      'Conserver une seule affectation pour cette personne et cette demi-journée.');
  });

  Object.keys(ticketsParSemaine).forEach(function(cle) {
    const item = ticketsParSemaine[cle];
    if (item.count > 3) {
      const premier = item.tickets[0];
      ajouterControleGuide_(resultats, 'Erreur', premier.feuille, premier.cellule, premier.date,
        'Tous postes', item.semaine, item.benevole,
        'Dépassement de 3 tickets par semaine : ' + item.count,
        'Vérifier les présences, puis utiliser Contrôle > Recalculer les tickets.');
    }
  });

  const entetes = ['Niveau', 'Feuille', 'Cellule', 'Date', 'Poste', 'Créneau', 'Bénévole', 'Problème', 'Correction conseillée'];
  const initialiserMiseEnPage = controle.getRange(1, 1).getDisplayValue() !== 'Niveau';
  controle.getRange(1, 1, 1, entetes.length).setValues([entetes]);
  const rowsToClear = Math.max(controle.getMaxRows() - 1, 1);
  controle.getRange(2, 1, rowsToClear, entetes.length).clearContent();
  if (resultats.length) {
    if (controle.getMaxRows() < resultats.length + 1) {
      controle.insertRowsAfter(controle.getMaxRows(), resultats.length + 1 - controle.getMaxRows());
    }
    const valeurs = resultats.map(function(item) {
      return [item.niveau, item.feuille, item.cellule, item.date, item.poste, item.creneau, item.benevole, item.probleme, item.correction];
    });
    controle.getRange(2, 1, valeurs.length, entetes.length).setValues(valeurs);
    const liens = resultats.map(function(item) {
      const url = ss.getUrl() + '#gid=' + item.gid + '&range=' + encodeURIComponent(item.cellule);
      return [SpreadsheetApp.newRichTextValue().setText(item.cellule).setLinkUrl(url).build()];
    });
    controle.getRange(2, 3, liens.length, 1).setRichTextValues(liens);
  }

  mettreEnFormeControleGuide_(controle, resultats.length, initialiserMiseEnPage);

  const bilan = compterNiveauxControle_(resultats);
  const resume = bilan.erreurs + ' erreur(s), ' + bilan.avertissements + ' avertissement(s), ' + bilan.informations + ' information(s)';
  if (afficherAlerte) journaliser_('Contrôle qualité guidé', resume);
  if (afficherAlerte) SpreadsheetApp.getUi().alert('Contrôle qualité guidé terminé :\n\n' + resume + '\n\nConsulte la feuille CONTROLE_QUALITE pour les corrections conseillées.');
  return resultats;
}

function chargerNomsBenevoles_(ss) {
  const resultat = {};
  const feuille = ss.getSheetByName('BENEVOLES');
  if (!feuille || feuille.getLastRow() < 2) return resultat;
  feuille.getRange(2, 1, feuille.getLastRow() - 1, 1).getDisplayValues().forEach(function(ligne) {
    const nom = String(ligne[0] || '').trim();
    if (nom) resultat[nom] = true;
  });
  return resultat;
}

function ajouterControleGuide_(resultats, niveau, feuille, cellule, date, poste, creneau, benevole, probleme, correction) {
  resultats.push({
    niveau: niveau,
    feuille: feuille.getName(),
    gid: feuille.getSheetId(),
    cellule: cellule,
    date: date,
    poste: poste,
    creneau: creneau,
    benevole: benevole,
    probleme: probleme,
    correction: correction
  });
}

function colonneA1_(numeroColonne) {
  let valeur = numeroColonne;
  let resultat = '';
  while (valeur > 0) {
    const reste = (valeur - 1) % 26;
    resultat = String.fromCharCode(65 + reste) + resultat;
    valeur = Math.floor((valeur - 1) / 26);
  }
  return resultat;
}

function compterNiveauxControle_(resultats) {
  return resultats.reduce(function(total, item) {
    if (item.niveau === 'Erreur') total.erreurs++;
    else if (item.niveau === 'Avertissement') total.avertissements++;
    else total.informations++;
    return total;
  }, { erreurs: 0, avertissements: 0, informations: 0 });
}

function mettreEnFormeControleGuide_(controle, nbResultats, initialiserMiseEnPage) {
  const largeur = 9;
  controle.getRange(1, 1, 1, largeur)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setWrap(true);
  if (initialiserMiseEnPage) {
    controle.setFrozenRows(1);
    controle.setColumnWidth(1, 115);
    controle.setColumnWidth(2, 120);
    controle.setColumnWidth(3, 80);
    controle.setColumnWidth(4, 105);
    controle.setColumnWidth(5, 130);
    controle.setColumnWidth(6, 140);
    controle.setColumnWidth(7, 180);
    controle.setColumnWidth(8, 260);
    controle.setColumnWidth(9, 360);
  }
  const nbAnciennesLignes = Math.max(controle.getMaxRows() - 1, 1);
  controle.getRange(2, 1, nbAnciennesLignes, 1)
    .setBackground(null)
    .setFontWeight('normal');
  if (!nbResultats) return;
  const plage = controle.getRange(2, 1, nbResultats, largeur);
  plage.setVerticalAlignment('top').setWrap(true);
  controle.getRange(2, 4, nbResultats, 1).setNumberFormat('dd/mm/yyyy');
  const couleurs = controle.getRange(2, 1, nbResultats, 1).getDisplayValues().map(function(ligne) {
    if (ligne[0] === 'Erreur') return ['#f4cccc'];
    if (ligne[0] === 'Avertissement') return ['#fce5cd'];
    return ['#d9ead3'];
  });
  controle.getRange(2, 1, nbResultats, 1).setBackgrounds(couleurs).setFontWeight('bold');
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

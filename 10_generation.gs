// Café&Co — génération des mois et application de la semaine type

function genererMoisDepuisModele() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const params = lireParametres_();

  const anneeRep = ui.prompt('Année', 'Exemple : ' + params.anneeActive, ui.ButtonSet.OK_CANCEL);
  if (anneeRep.getSelectedButton() !== ui.Button.OK) return;

  const moisRep = ui.prompt('Mois', 'Numéro du mois : 1 à 12', ui.ButtonSet.OK_CANCEL);
  if (moisRep.getSelectedButton() !== ui.Button.OK) return;

  const annee = Number(anneeRep.getResponseText());
  const mois = Number(moisRep.getResponseText());
  if (!Number.isInteger(annee) || annee < 2000 || annee > 2100 || !Number.isInteger(mois) || mois < 1 || mois > 12) {
    ui.alert('Année ou mois invalide.');
    return;
  }

  const nomFeuille = CAFCO_MOIS[mois - 1];
  const modele = ss.getSheetByName(params.modeleMois);
  if (!modele) {
    ui.alert('Feuille modèle introuvable : ' + params.modeleMois);
    return;
  }

  let cible = ss.getSheetByName(nomFeuille);
  if (cible) {
    const choix = ui.alert(
      nomFeuille + ' existe déjà',
      'La feuille sera sauvegardée puis régénérée sans changer son identité. Continuer ?',
      ui.ButtonSet.YES_NO
    );
    if (choix !== ui.Button.YES) return;

    const horodatage = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const sauvegarde = cible.copyTo(ss).setName('_BACKUP_' + nomFeuille + '_' + horodatage);
    sauvegarde.hideSheet();
    reinitialiserFeuilleDepuisModele_(cible, modele);
  } else {
    cible = modele.copyTo(ss).setName(nomFeuille);
  }

  const lignes = construireLignesMois_(annee, mois, params);
  ajusterNombreLignes_(cible, lignes.length, params.postes.length);
  nettoyerMoisV2_(cible);
  ecrireDatesEtPostesV2_(cible, lignes);
  installerFormulesTicketsPourFeuille_(cible);
  recalculerTicketsFeuille_(cible);
  protegerFormulesTickets_(cible);
  SpreadsheetApp.flush();

  journaliser_('Génération mois', nomFeuille + ' ' + annee + ' : ' + lignes.length + ' ligne(s)');
  ui.alert(nomFeuille + ' ' + annee + ' généré avec ' + lignes.length + ' ligne(s).');
}

function reinitialiserFeuilleDepuisModele_(cible, modele) {
  const lignesModele = modele.getDataRange().getNumRows();
  const colsModele = modele.getDataRange().getNumColumns();

  if (cible.getMaxColumns() < colsModele) {
    cible.insertColumnsAfter(cible.getMaxColumns(), colsModele - cible.getMaxColumns());
  }
  if (cible.getMaxRows() < lignesModele) {
    cible.insertRowsAfter(cible.getMaxRows(), lignesModele - cible.getMaxRows());
  }

  cible.clear();
  modele.getRange(1, 1, lignesModele, colsModele).copyTo(
    cible.getRange(1, 1, lignesModele, colsModele),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );
  cible.setFrozenRows(modele.getFrozenRows());
  cible.setFrozenColumns(modele.getFrozenColumns());
}

function construireLignesMois_(annee, mois, params) {
  const joursNoms = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  const dernierJour = new Date(annee, mois, 0).getDate();
  const lignes = [];

  for (let jour = 1; jour <= dernierJour; jour++) {
    const date = new Date(annee, mois - 1, jour, 12, 0, 0);
    if (params.joursOuvres.indexOf(joursNoms[date.getDay()]) === -1) continue;
    params.postes.forEach(function(poste) { lignes.push([date, poste]); });
  }
  return lignes;
}

function ajusterNombreLignes_(sheet, nbLignesDonnees, nbPostes) {
  if (!nbPostes || nbPostes < 1) throw new Error('Le nombre de postes est invalide.');
  const necessaires = nbLignesDonnees + 1;

  if (sheet.getMaxRows() < necessaires) {
    const aAjouter = necessaires - sheet.getMaxRows();
    sheet.insertRowsAfter(sheet.getMaxRows(), aAjouter);

    const sourceStart = Math.max(2, sheet.getLastRow() - nbPostes + 1);
    const bloc = sheet.getRange(sourceStart, 1, Math.min(nbPostes, sheet.getMaxRows() - sourceStart + 1), 26);
    let destinationRow = sourceStart + bloc.getNumRows();
    while (destinationRow <= necessaires) {
      const hauteur = Math.min(bloc.getNumRows(), necessaires - destinationRow + 1);
      bloc.offset(0, 0, hauteur, 26).copyTo(
        sheet.getRange(destinationRow, 1, hauteur, 26),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );
      destinationRow += hauteur;
    }
  }

  if (sheet.getMaxRows() > necessaires) {
    sheet.deleteRows(necessaires + 1, sheet.getMaxRows() - necessaires);
  }
}

function nettoyerMoisV2_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 1) return;
  [3,4,6,7,9,10,12,13,15,16,18,19,21,22,24,25].forEach(function(col) {
    sheet.getRange(2, col, maxRows - 1, 1).clearContent();
  });
}

function ecrireDatesEtPostesV2_(sheet, lignes) {
  if (lignes.length) sheet.getRange(2, 1, lignes.length, 2).setValues(lignes);
}

function appliquerSemaineTypeUneSemaine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const semaineType = ss.getSheetByName('SEMAINE_TYPE');
  if (!semaineType) {
    ui.alert('La feuille SEMAINE_TYPE est introuvable.');
    return;
  }

  const repMois = ui.prompt('Mois cible', 'Indique le nom exact du mois, par exemple : OCTOBRE', ui.ButtonSet.OK_CANCEL);
  if (repMois.getSelectedButton() !== ui.Button.OK) return;
  const nomMois = repMois.getResponseText().trim().toUpperCase();
  const feuille = ss.getSheetByName(nomMois);
  if (!feuille || !estFeuilleMois_(nomMois)) {
    ui.alert('Feuille mensuelle introuvable : ' + nomMois);
    return;
  }

  const repLundi = ui.prompt('Semaine cible', 'Lundi au format AAAA-MM-JJ', ui.ButtonSet.OK_CANCEL);
  if (repLundi.getSelectedButton() !== ui.Button.OK) return;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(repLundi.getResponseText().trim());
  if (!m) {
    ui.alert('Date invalide. Utilise AAAA-MM-JJ.');
    return;
  }
  const lundi = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  if (isNaN(lundi.getTime()) || lundi.getDay() !== 1) {
    ui.alert('La date doit correspondre à un lundi.');
    return;
  }
  const fin = new Date(lundi); fin.setDate(fin.getDate() + 4);

  const choixMode = ui.alert('Périmètre', 'Oui = tous les postes. Non = un seul poste.', ui.ButtonSet.YES_NO_CANCEL);
  if (choixMode === ui.Button.CANCEL) return;
  let posteChoisi = null;
  if (choixMode === ui.Button.NO) {
    const params = lireParametres_();
    const repPoste = ui.prompt('Poste', 'Choisir parmi : ' + params.postes.join(', '), ui.ButtonSet.OK_CANCEL);
    if (repPoste.getSelectedButton() !== ui.Button.OK) return;
    posteChoisi = repPoste.getResponseText().trim();
    if (params.postes.indexOf(posteChoisi) === -1) {
      ui.alert('Poste inconnu : ' + posteChoisi);
      return;
    }
  }

  const choixEcr = ui.alert('Écraser les noms existants ?', 'Oui = remplacer. Non = remplir seulement les cellules vides.', ui.ButtonSet.YES_NO_CANCEL);
  if (choixEcr === ui.Button.CANCEL) return;
  const ecraser = choixEcr === ui.Button.YES;

  const st = semaineType.getRange(2, 1, Math.max(semaineType.getLastRow() - 1, 1), 10).getValues();
  const modele = {};
  st.forEach(function(l) {
    const jour = String(l[0] || '').trim();
    const poste = String(l[1] || '').trim();
    if (jour && poste) modele[jour + '|' + poste] = l.slice(2, 10);
  });
  const jours = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  const data = feuille.getRange(2, 1, Math.max(feuille.getLastRow() - 1, 1), 26).getValues();
  let modifies = 0;

  data.forEach(function(ligne) {
    const date = ligne[0];
    const poste = String(ligne[1] || '').trim();
    if (!(date instanceof Date) || !poste || date < lundi || date > fin || (posteChoisi && poste !== posteChoisi)) return;
    const noms = modele[jours[date.getDay()] + '|' + poste];
    if (!noms) return;

    CAFCO_SLOTS.forEach(function(slot, i) {
      const beneIdx = slot.beneCol - 1;
      const statutIdx = slot.statutCol - 1;
      const nouveau = noms[i] || '';
      if (ecraser || (!ligne[beneIdx] && nouveau)) {
        if (ligne[beneIdx] !== nouveau) modifies++;
        ligne[beneIdx] = nouveau;
        ligne[statutIdx] = '';
      }
    });
  });

  CAFCO_SLOTS.forEach(function(slot) {
    feuille.getRange(2, slot.beneCol, data.length, 1).setValues(data.map(function(l) { return [l[slot.beneCol - 1]]; }));
    feuille.getRange(2, slot.statutCol, data.length, 1).setValues(data.map(function(l) { return [l[slot.statutCol - 1]]; }));
  });

  recalculerTicketsFeuille_(feuille);
  verifierControleQualite_(false);
  journaliser_('Application SEMAINE_TYPE', nomMois + ' — semaine du ' + cleJour_(lundi) + ' — ' + modifies + ' affectation(s) modifiée(s)');
  ui.alert('Semaine type appliquée. Affectations modifiées : ' + modifies + '.');
}

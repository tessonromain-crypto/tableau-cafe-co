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
function installerFormulesTicketsPourFeuille_(sheet) {
  if (!sheet || !estFeuilleMois_(sheet.getName())) return 0;
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const schema = schemaPlanning_(sheet);
  let cellulesPreparees = 0;

  schema.slots.forEach(function(slot) {
    const range = sheet.getRange(2, slot.ticketCol, lastRow - 1, 1);
    range.clearContent();
    cellulesPreparees += lastRow - 1;
  });
  return cellulesPreparees;
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
  executerControleQualiteGuide_(false);
  journaliser_('Recalcul tickets', nomMois + ' : ' + stats.oui + ' Oui, ' + stats.non + ' Non, ' + stats.doublons + ' doublon(s) neutralisé(s)');
  ui.alert('Tickets recalculés.\n\nOui : ' + stats.oui + '\nNon : ' + stats.non + '\nVides : ' + stats.vides + '\nDoublons neutralisés : ' + stats.doublons);
}

function corrigerTicketsMaxSemaine() {
  recalculerTicketsMois();
}

function reparerFormulesTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const choix = ui.alert(
    'Réparer les tickets de tous les mois',
    'Cette action recalculera les tickets et remettra les protections sur tous les mois existants. Continuer ?',
    ui.ButtonSet.YES_NO
  );
  if (choix !== ui.Button.YES) return;

  let feuilles = 0;
  let entetesCorriges = 0;
  const infosBenevoles = lireInfosBenevoles_();

  CAFCO_MOIS.forEach(function(nom) {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;
    const schema = schemaPlanning_(sh);

    if (sh.getMaxColumns() < schema.totalCols) {
      sh.insertColumnsAfter(sh.getMaxColumns(), schema.totalCols - sh.getMaxColumns());
    }

    schema.slots.forEach(function(slot) {
      const celluleEntete = sh.getRange(1, slot.ticketCol);
      if (celluleEntete.getDisplayValue() !== 'Ticket') {
        celluleEntete.setValue('Ticket');
        entetesCorriges++;
      }
    });

    recalculerTicketsFeuille_(sh, infosBenevoles);
    protegerFormulesTickets_(sh);
    feuilles++;
  });

  executerControleQualiteGuide_(false);
  SpreadsheetApp.flush();
  journaliser_(
    'Réparation tickets',
    feuilles + ' feuille(s), ' + entetesCorriges + ' en-tête(s) corrigé(s)'
  );
  ui.alert(
    'Réparation terminée.\n\n' +
    'Mois traités : ' + feuilles + '\n' +
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
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const nom = sh.getName();

  if (estFeuilleMois_(nom)) {
    const schema = schemaPlanning_(sh);

    // Chemin rapide : une seule cellule modifiée (cas courant au quotidien).
    if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && e.range.getRow() >= 2) {
      const col = e.range.getColumn();
      const slot = schema.slots.find(function(item) {
        return item.beneCol === col || item.statutCol === col;
      });
      if (!slot) return;

      const row = e.range.getRow();
      const noms = [];
      let ticketAForcerVide = null;

      if (col === slot.beneCol) {
        // Nouveau nom + ancien nom : nécessaire lors d'un remplacement ou d'une suppression.
        const nouveauNom = String(e.value || '').trim();
        const ancienNom = String(e.oldValue || '').trim();
        if (nouveauNom) noms.push(nouveauNom);
        if (ancienNom && ancienNom !== nouveauNom) noms.push(ancienNom);
        ticketAForcerVide = slot.ticketCol;
      } else {
        // Modification du statut : le nom est dans la cellule bénévole de la même ligne.
        const benevole = String(sh.getRange(row, slot.beneCol).getValue() || '').trim();
        if (benevole) noms.push(benevole);
      }

      recalculerTicketsPersonnesSemaine_(sh, row, noms, ticketAForcerVide);
      return;
    }

    // Collages/plages multiples : filet de sécurité, recalcul des semaines touchées.
    const premiereCol = e.range.getColumn();
    const derniereCol = e.range.getLastColumn();
    const concernePlanning = schema.slots.some(function(slot) {
      return (slot.beneCol >= premiereCol && slot.beneCol <= derniereCol) ||
        (slot.statutCol >= premiereCol && slot.statutCol <= derniereCol);
    });
    if (!concernePlanning || e.range.getLastRow() < 2) return;

    const premiereLigne = Math.max(2, e.range.getRow());
    const derniereLigne = e.range.getLastRow();
    const dates = sh.getRange(
      premiereLigne,
      schema.dateCol,
      derniereLigne - premiereLigne + 1,
      1
    ).getValues();

    const semaines = {};
    dates.forEach(function(ligne) {
      const date = ligne[0];
      if (date instanceof Date) semaines[getSemaineCle_(date)] = true;
    });

    const clesSemaines = Object.keys(semaines);
    if (clesSemaines.length) {
      recalculerTicketsSemaines_(sh, clesSemaines);
    } else {
      recalculerTicketsFeuille_(sh);
    }
    return;
  }

  if (nom === 'BENEVOLES' && e.range.getRow() >= 2 && e.range.getColumn() <= 8) {
    const infosBenevoles = lireInfosBenevoles_();
    CAFCO_MOIS.forEach(function(mois) {
      const moisSheet = e.source.getSheetByName(mois);
      if (moisSheet) recalculerTicketsFeuille_(moisSheet, infosBenevoles);
    });
    executerControleQualiteGuide_(false);
  }
}

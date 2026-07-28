// CAFÉ&CO — logique Tickets V2
// Branche de test : ne remplace pas encore le code de production.
// Objectifs :
// 1) un même bénévole présent plusieurs fois sur le même créneau (matin/après-midi)
//    ne consomme qu'un seul ticket ;
// 2) maximum 3 tickets par semaine et par bénévole ;
// 3) aucune écriture dans les cellules Ticket pendant les tests, afin de ne pas
//    remplacer les formules existantes.

function calculerTicketsV2_(planning, infosBenevoles) {
  const slots = [
    { benevoleIndex: 2, statutIndex: 3, ticketIndex: 4, periode: 'MATIN' },
    { benevoleIndex: 5, statutIndex: 6, ticketIndex: 7, periode: 'MATIN' },
    { benevoleIndex: 8, statutIndex: 9, ticketIndex: 10, periode: 'MATIN' },
    { benevoleIndex: 11, statutIndex: 12, ticketIndex: 13, periode: 'MATIN' },
    { benevoleIndex: 14, statutIndex: 15, ticketIndex: 16, periode: 'APRES_MIDI' },
    { benevoleIndex: 17, statutIndex: 18, ticketIndex: 19, periode: 'APRES_MIDI' },
    { benevoleIndex: 20, statutIndex: 21, ticketIndex: 22, periode: 'APRES_MIDI' },
    { benevoleIndex: 23, statutIndex: 24, ticketIndex: 25, periode: 'APRES_MIDI' }
  ];

  const compteursSemaine = {};
  const creneauxDejaComptes = {};
  const resultats = planning.map(function(ligne) { return ligne.slice(); });

  let ticketsOui = 0;
  let ticketsNon = 0;
  let ticketsVides = 0;
  let doublonsCreneau = 0;

  for (let r = 0; r < resultats.length; r++) {
    const date = resultats[r][0];

    if (!(date instanceof Date)) {
      slots.forEach(function(slot) {
        resultats[r][slot.ticketIndex] = '';
        ticketsVides++;
      });
      continue;
    }

    const semaine = getSemaineCle_(date);
    const jour = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    slots.forEach(function(slot) {
      const benevole = String(resultats[r][slot.benevoleIndex] || '').trim();
      const presence = String(resultats[r][slot.statutIndex] || '').trim();
      let resultat = '';

      if (!benevole || !presence) {
        resultat = '';
      } else if (presence !== 'Présent') {
        resultat = 'Non';
      } else {
        const infos = infosBenevoles[benevole];
        const souhaiteTicket = infos && String(infos.souhaiteTicket || '').toLowerCase() === 'oui';
        const statutEligible = infos && (
          infos.statutAutomatique === 'Bénévole' ||
          infos.statutAutomatique === 'Référent'
        );

        if (!souhaiteTicket || !statutEligible) {
          resultat = 'Non';
        } else {
          // Une personne = un seul ticket possible par demi-journée,
          // même si elle apparaît sur plusieurs postes.
          const cleCreneau = jour + '|' + slot.periode + '|' + benevole;

          if (creneauxDejaComptes[cleCreneau]) {
            resultat = 'Non';
            doublonsCreneau++;
          } else {
            const cleSemaine = semaine + '|' + benevole;
            const nombreDejaAttribue = compteursSemaine[cleSemaine] || 0;

            if (nombreDejaAttribue < 3) {
              resultat = 'Oui';
              compteursSemaine[cleSemaine] = nombreDejaAttribue + 1;
              creneauxDejaComptes[cleCreneau] = true;
            } else {
              resultat = 'Non';
            }
          }
        }
      }

      resultats[r][slot.ticketIndex] = resultat;

      if (resultat === 'Oui') ticketsOui++;
      else if (resultat === 'Non') ticketsNon++;
      else ticketsVides++;
    });
  }

  return {
    planning: resultats,
    stats: {
      ticketsOui: ticketsOui,
      ticketsNon: ticketsNon,
      ticketsVides: ticketsVides,
      doublonsCreneau: doublonsCreneau
    }
  };
}

// Analyse le mois sans modifier aucune cellule du classeur.
function analyserTicketsMoisV2SansEcriture() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const reponse = ui.prompt(
    'Tester Tickets V2',
    'Indique le mois, par exemple : JUIN',
    ui.ButtonSet.OK_CANCEL
  );
  if (reponse.getSelectedButton() !== ui.Button.OK) return;

  const nomMois = reponse.getResponseText().trim().toUpperCase();
  const sheet = ss.getSheetByName(nomMois);
  const benevolesSheet = ss.getSheetByName('BENEVOLES');

  if (!sheet || !benevolesSheet) {
    ui.alert('Feuille mensuelle ou BENEVOLES introuvable.');
    return;
  }

  const derniereLigne = sheet.getLastRow();
  if (derniereLigne < 2) {
    ui.alert('Aucune donnée à analyser dans ' + nomMois + '.');
    return;
  }

  const planning = sheet.getRange(2, 1, derniereLigne - 1, 26).getValues();
  const benevolesData = benevolesSheet
    .getRange(2, 1, Math.max(benevolesSheet.getLastRow() - 1, 1), 9)
    .getValues();

  const infosBenevoles = {};
  benevolesData.forEach(function(ligne) {
    const nom = String(ligne[0] || '').trim();
    if (!nom) return;
    infosBenevoles[nom] = {
      souhaiteTicket: ligne[6],
      statutAutomatique: ligne[7]
    };
  });

  const calcul = calculerTicketsV2_(planning, infosBenevoles);

  Logger.log(JSON.stringify(calcul.stats));
  ui.alert(
    'Analyse V2 terminée — aucune cellule modifiée.\n\n' +
    'Tickets Oui : ' + calcul.stats.ticketsOui + '\n' +
    'Tickets Non : ' + calcul.stats.ticketsNon + '\n' +
    'Vides : ' + calcul.stats.ticketsVides + '\n' +
    'Doublons de créneau neutralisés : ' + calcul.stats.doublonsCreneau
  );
}

// Tests unitaires simples, sans accès au classeur.
// À exécuter depuis Apps Script : testerLogiqueTicketsV2
function testerLogiqueTicketsV2() {
  const infos = {
    test: { souhaiteTicket: 'Oui', statutAutomatique: 'Bénévole' }
  };

  function ligne(date, nomMatin) {
    const l = new Array(26).fill('');
    l[0] = date;
    l[1] = 'Café';
    l[2] = nomMatin;
    l[3] = nomMatin ? 'Présent' : '';
    return l;
  }

  // Cas 1 : quatre postes le même matin = un seul Oui.
  const date1 = new Date(2026, 5, 1, 12, 0, 0);
  const quatrePostes = [
    ligne(date1, 'test'),
    ligne(date1, 'test'),
    ligne(date1, 'test'),
    ligne(date1, 'test')
  ];
  const r1 = calculerTicketsV2_(quatrePostes, infos);
  const ouiCas1 = r1.planning.filter(function(l) { return l[4] === 'Oui'; }).length;
  if (ouiCas1 !== 1) {
    throw new Error('TEST 1 ÉCHEC : attendu 1 ticket, obtenu ' + ouiCas1);
  }

  // Cas 2 : quatre jours différents dans la même semaine = seulement trois Oui.
  const quatreJours = [
    ligne(new Date(2026, 5, 1, 12), 'test'),
    ligne(new Date(2026, 5, 2, 12), 'test'),
    ligne(new Date(2026, 5, 3, 12), 'test'),
    ligne(new Date(2026, 5, 4, 12), 'test')
  ];
  const r2 = calculerTicketsV2_(quatreJours, infos);
  const ouiCas2 = r2.planning.filter(function(l) { return l[4] === 'Oui'; }).length;
  if (ouiCas2 !== 3) {
    throw new Error('TEST 2 ÉCHEC : attendu 3 tickets, obtenu ' + ouiCas2);
  }

  // Cas 3 : absent = Non.
  const absent = ligne(date1, 'test');
  absent[3] = 'Absent';
  const r3 = calculerTicketsV2_([absent], infos);
  if (r3.planning[0][4] !== 'Non') {
    throw new Error('TEST 3 ÉCHEC : un absent doit avoir Non.');
  }

  Logger.log('Tous les tests Tickets V2 sont OK.');
  return 'OK';
}

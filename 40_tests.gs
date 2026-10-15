// Café&Co — tests de non-régression

function testerLogiqueTickets() {
  const infos = {
    test: { souhaiteTicket: 'Oui', statutAutomatique: 'Bénévole' },
    sansTicket: { souhaiteTicket: 'Non', statutAutomatique: 'Bénévole' },
    formation: { souhaiteTicket: 'Oui', statutAutomatique: 'Formation' }
  };

  function schemaHistorique_() {
    return { avecJour: false, dateCol: 1, posteCol: 2, totalCols: 26, slots: CAFCO_SLOTS };
  }

  function schemaAvecJour_() {
    return {
      avecJour: true,
      jourCol: 1,
      dateCol: 2,
      posteCol: 3,
      totalCols: 27,
      slots: CAFCO_SLOTS.map(function(slot) {
        return {
          beneCol: slot.beneCol + 1,
          statutCol: slot.statutCol + 1,
          ticketCol: slot.ticketCol + 1,
          periode: slot.periode,
          nom: slot.nom
        };
      })
    };
  }

  function ligne(date, nom, presence, slotIndex, schema) {
    schema = schema || schemaHistorique_();
    const l = new Array(schema.totalCols).fill('');
    if (schema.avecJour) l[0] = 'Lundi';
    l[schema.dateCol - 1] = date;
    l[schema.posteCol - 1] = 'Café';
    const slot = schema.slots[slotIndex || 0];
    l[slot.beneCol - 1] = nom || '';
    l[slot.statutCol - 1] = presence || '';
    return l;
  }

  function assert_(condition, message) {
    if (!condition) throw new Error(message);
  }

  const ancien = schemaHistorique_();

  // 1. Même personne sur quatre postes le même matin = un seul ticket Oui.
  const d1 = new Date(2026, 5, 1, 12, 0, 0);
  const quatrePostes = [
    ligne(d1, 'test', 'Présent', 0, ancien),
    ligne(d1, 'test', 'Présent', 0, ancien),
    ligne(d1, 'test', 'Présent', 0, ancien),
    ligne(d1, 'test', 'Présent', 0, ancien)
  ];
  const r1 = calculerTickets_(quatrePostes, infos, ancien);
  assert_(r1.planning.filter(function(l) { return l[4] === 'Oui'; }).length === 1,
    'TEST 1 : quatre postes le même matin doivent produire un seul ticket.');

  // 2. Quatre demi-journées différentes la même semaine = trois tickets maximum.
  const quatreJours = [
    ligne(new Date(2026, 5, 1, 12), 'test', 'Présent', 0, ancien),
    ligne(new Date(2026, 5, 2, 12), 'test', 'Présent', 0, ancien),
    ligne(new Date(2026, 5, 3, 12), 'test', 'Présent', 0, ancien),
    ligne(new Date(2026, 5, 4, 12), 'test', 'Présent', 0, ancien)
  ];
  const r2 = calculerTickets_(quatreJours, infos, ancien);
  assert_(r2.planning.filter(function(l) { return l[4] === 'Oui'; }).length === 3,
    'TEST 2 : maximum trois tickets par semaine.');

  // 3. Matin + après-midi le même jour = deux créneaux distincts.
  const matinAprem = [
    ligne(d1, 'test', 'Présent', 0, ancien),
    ligne(d1, 'test', 'Présent', 4, ancien)
  ];
  const r3 = calculerTickets_(matinAprem, infos, ancien);
  assert_(r3.planning[0][4] === 'Oui' && r3.planning[1][16] === 'Oui',
    'TEST 3 : matin et après-midi doivent être deux créneaux distincts.');

  // 4. Absent ou retard = Non.
  const r4 = calculerTickets_([
    ligne(d1, 'test', 'Absent', 0, ancien),
    ligne(new Date(2026, 5, 2, 12), 'test', 'Retard', 0, ancien)
  ], infos, ancien);
  assert_(r4.planning[0][4] === 'Non' && r4.planning[1][4] === 'Non',
    'TEST 4 : absent/retard doit donner Non.');

  // 5. Pas de souhait ticket = Non.
  const r5 = calculerTickets_([ligne(d1, 'sansTicket', 'Présent', 0, ancien)], infos, ancien);
  assert_(r5.planning[0][4] === 'Non', 'TEST 5 : bénévole ne souhaitant pas de ticket = Non.');

  // 6. Formation = Non.
  const r6 = calculerTickets_([ligne(d1, 'formation', 'Présent', 0, ancien)], infos, ancien);
  assert_(r6.planning[0][4] === 'Non', 'TEST 6 : statut Formation = Non.');

  // 7. Nom ou présence vide = ticket vide.
  const r7 = calculerTickets_([
    ligne(d1, '', '', 0, ancien),
    ligne(new Date(2026, 5, 2, 12), 'test', '', 0, ancien)
  ], infos, ancien);
  assert_(r7.planning[0][4] === '' && r7.planning[1][4] === '',
    'TEST 7 : données incomplètes = ticket vide.');

  // 8. Nouveau format avec colonne Jour : la logique Ticket reste identique.
  const nouveau = schemaAvecJour_();
  const r8 = calculerTickets_([
    ligne(d1, 'test', 'Présent', 0, nouveau),
    ligne(d1, 'test', 'Présent', 4, nouveau)
  ], infos, nouveau);
  assert_(r8.planning[0][nouveau.slots[0].ticketCol - 1] === 'Oui' &&
          r8.planning[1][nouveau.slots[4].ticketCol - 1] === 'Oui',
    'TEST 8 : le format Jour/Date/Poste doit conserver la logique Ticket.');

  // 9. Une date comprise dans une période de fermeture doit être exclue.
  const fermetures = [{
    debut: new Date(2027, 1, 15, 12, 0, 0),
    fin: new Date(2027, 1, 19, 12, 0, 0)
  }];
  assert_(estDateDansFermeture_(new Date(2027, 1, 17, 12), fermetures) === true,
    'TEST 9 : une date située dans une fermeture doit être exclue.');
  assert_(estDateDansFermeture_(new Date(2027, 1, 22, 12), fermetures) === false,
    'TEST 9 : une date hors fermeture ne doit pas être exclue.');

  // 10. Avant titularisation = Non ; à partir de la titularisation = Oui.
  const infosHistorique = {
    historique: {
      dateEntree: new Date(2026, 5, 1, 12),
      dateTitularisation: new Date(2026, 8, 1, 12),
      dateSortie: '',
      souhaiteTicket: 'Oui',
      statutAutomatique: 'Référent'
    }
  };
  const r10 = calculerTickets_([
    ligne(new Date(2026, 7, 31, 12), 'historique', 'Présent', 0, ancien),
    ligne(new Date(2026, 8, 1, 12), 'historique', 'Présent', 0, ancien)
  ], infosHistorique, ancien);
  assert_(r10.planning[0][4] === 'Non' && r10.planning[1][4] === 'Oui',
    'TEST 10 : la titularisation doit prendre effet à sa date, sans rétroactivité.');

  // 11. Un statut actuel Bénévole ne rend pas éligible avant la date de titularisation.
  const infosTitFuture = {
    futur: {
      dateEntree: new Date(2026, 4, 1, 12),
      dateTitularisation: new Date(2026, 9, 1, 12),
      dateSortie: '',
      souhaiteTicket: 'Oui',
      statutAutomatique: 'Bénévole'
    }
  };
  const r11 = calculerTickets_([
    ligne(new Date(2026, 8, 15, 12), 'futur', 'Présent', 0, ancien),
    ligne(new Date(2026, 9, 1, 12), 'futur', 'Présent', 0, ancien)
  ], infosTitFuture, ancien);
  assert_(r11.planning[0][4] === 'Non' && r11.planning[1][4] === 'Oui',
    'TEST 11 : le statut actuel ne doit pas antidater l’éligibilité.');

  // 12. Une sortie ultérieure ne supprime pas les tickets acquis avant la sortie.
  const infosSorti = {
    sorti: {
      dateEntree: new Date(2026, 4, 1, 12),
      dateTitularisation: new Date(2026, 5, 1, 12),
      dateSortie: new Date(2026, 9, 9, 12),
      souhaiteTicket: 'Oui',
      statutAutomatique: 'Sorti'
    }
  };
  const r12 = calculerTickets_([
    ligne(new Date(2026, 8, 20, 12), 'sorti', 'Présent', 0, ancien),
    ligne(new Date(2026, 9, 9, 12), 'sorti', 'Présent', 0, ancien),
    ligne(new Date(2026, 9, 10, 12), 'sorti', 'Présent', 0, ancien)
  ], infosSorti, ancien);
  assert_(r12.planning[0][4] === 'Oui' && r12.planning[1][4] === 'Oui' && r12.planning[2][4] === 'Non',
    'TEST 12 : les tickets avant/au jour de sortie restent valides, après sortie = Non.');

  // 13. Cas Malik : cinq présences réparties sur plusieurs semaines en septembre = cinq Oui.
  const infosMalik = {
    Malik: {
      dateEntree: new Date(2026, 5, 8, 12),
      dateTitularisation: new Date(2026, 8, 1, 12),
      dateSortie: '',
      souhaiteTicket: 'Oui',
      statutAutomatique: 'Référent'
    }
  };
  const r13 = calculerTickets_([
    ligne(new Date(2026, 8, 2, 12), 'Malik', 'Présent', 0, ancien),
    ligne(new Date(2026, 8, 4, 12), 'Malik', 'Présent', 0, ancien),
    ligne(new Date(2026, 8, 16, 12), 'Malik', 'Présent', 0, ancien),
    ligne(new Date(2026, 8, 23, 12), 'Malik', 'Présent', 0, ancien),
    ligne(new Date(2026, 8, 30, 12), 'Malik', 'Présent', 0, ancien)
  ], infosMalik, ancien);
  assert_(r13.planning.filter(function(l) { return l[4] === 'Oui'; }).length === 5,
    'TEST 13 : Malik doit obtenir cinq tickets en septembre.');

  Logger.log('Tous les tests Tickets, historique et fermeture sont OK.');
  SpreadsheetApp.getUi().alert('Tous les tests Tickets, historique et fermeture sont OK.');
  return 'OK';
}

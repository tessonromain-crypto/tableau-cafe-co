// Café&Co — tests de non-régression

function testerLogiqueTickets() {
  const infos = {
    test: { souhaiteTicket: 'Oui', statutAutomatique: 'Bénévole' },
    sansTicket: { souhaiteTicket: 'Non', statutAutomatique: 'Bénévole' },
    formation: { souhaiteTicket: 'Oui', statutAutomatique: 'Formation' }
  };

  function ligne(date, nom, presence, slotIndex) {
    const l = new Array(26).fill('');
    l[0] = date;
    l[1] = 'Café';
    const slot = CAFCO_SLOTS[slotIndex || 0];
    l[slot.beneCol - 1] = nom || '';
    l[slot.statutCol - 1] = presence || '';
    return l;
  }

  function assert_(condition, message) {
    if (!condition) throw new Error(message);
  }

  // 1. Même personne sur quatre postes le même matin = un seul ticket Oui.
  const d1 = new Date(2026, 5, 1, 12, 0, 0);
  const quatrePostes = [
    ligne(d1, 'test', 'Présent', 0),
    ligne(d1, 'test', 'Présent', 0),
    ligne(d1, 'test', 'Présent', 0),
    ligne(d1, 'test', 'Présent', 0)
  ];
  const r1 = calculerTickets_(quatrePostes, infos);
  assert_(r1.planning.filter(function(l) { return l[4] === 'Oui'; }).length === 1,
    'TEST 1 : quatre postes le même matin doivent produire un seul ticket.');

  // 2. Quatre demi-journées différentes la même semaine = trois tickets maximum.
  const quatreJours = [
    ligne(new Date(2026, 5, 1, 12), 'test', 'Présent', 0),
    ligne(new Date(2026, 5, 2, 12), 'test', 'Présent', 0),
    ligne(new Date(2026, 5, 3, 12), 'test', 'Présent', 0),
    ligne(new Date(2026, 5, 4, 12), 'test', 'Présent', 0)
  ];
  const r2 = calculerTickets_(quatreJours, infos);
  assert_(r2.planning.filter(function(l) { return l[4] === 'Oui'; }).length === 3,
    'TEST 2 : maximum trois tickets par semaine.');

  // 3. Matin + après-midi le même jour = deux créneaux distincts.
  const matinAprem = [
    ligne(d1, 'test', 'Présent', 0),
    ligne(d1, 'test', 'Présent', 4)
  ];
  const r3 = calculerTickets_(matinAprem, infos);
  assert_(r3.planning[0][4] === 'Oui' && r3.planning[1][16] === 'Oui',
    'TEST 3 : matin et après-midi doivent être deux créneaux distincts.');

  // 4. Absent ou retard = Non.
  const r4 = calculerTickets_([
    ligne(d1, 'test', 'Absent', 0),
    ligne(new Date(2026, 5, 2, 12), 'test', 'Retard', 0)
  ], infos);
  assert_(r4.planning[0][4] === 'Non' && r4.planning[1][4] === 'Non',
    'TEST 4 : absent/retard doit donner Non.');

  // 5. Pas de souhait ticket = Non.
  const r5 = calculerTickets_([ligne(d1, 'sansTicket', 'Présent', 0)], infos);
  assert_(r5.planning[0][4] === 'Non', 'TEST 5 : bénévole ne souhaitant pas de ticket = Non.');

  // 6. Formation = Non.
  const r6 = calculerTickets_([ligne(d1, 'formation', 'Présent', 0)], infos);
  assert_(r6.planning[0][4] === 'Non', 'TEST 6 : statut Formation = Non.');

  // 7. Nom ou présence vide = ticket vide.
  const r7 = calculerTickets_([
    ligne(d1, '', '', 0),
    ligne(new Date(2026, 5, 2, 12), 'test', '', 0)
  ], infos);
  assert_(r7.planning[0][4] === '' && r7.planning[1][4] === '',
    'TEST 7 : données incomplètes = ticket vide.');

  Logger.log('Tous les tests Tickets sont OK.');
  SpreadsheetApp.getUi().alert('Tous les tests Tickets sont OK.');
  return 'OK';
}

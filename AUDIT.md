# Audit du code Café&Co

## Problèmes critiques corrigés

1. **Destruction des formules Ticket**
   - L’ancien `recalculerTicketsMois()` écrivait directement `Oui/Non` dans les colonnes Ticket avec `setValues()`.
   - Le nouveau système calcule dans des feuilles cachées `_TICKETS_<MOIS>` et les cellules Ticket restent des formules.

2. **Même bénévole compté plusieurs fois sur la même demi-journée**
   - L’ancien compteur incrémentait à chaque apparition.
   - Le nouveau calcul utilise une clé `date + période + bénévole` : une seule attribution par matin/après-midi.

3. **Plafond de 3 tickets/semaine**
   - Le plafond est appliqué après dédoublonnage, donc plusieurs postes le même matin ne consomment plus plusieurs tickets.

4. **Fonction de correction destructive**
   - L’ancien `corrigerTicketsMaxSemaine()` remplaçait des formules par la valeur `Non`.
   - Il devient un alias vers le recalcul sécurisé.

5. **Forçage des synthèses très coûteux**
   - L’ancien code relisait puis réécrivait chaque formule cellule par cellule.
   - Remplacé par `SpreadsheetApp.flush()`.

6. **Génération mensuelle risquée**
   - L’ancien code supprimait une feuille existante avant de reconstruire la nouvelle.
   - Le nouveau code crée une sauvegarde cachée puis régénère la feuille existante sans changer son identité.

7. **Application SEMAINE_TYPE lente**
   - L’ancien code utilisait de nombreux `getRange()/setValue()` dans des boucles.
   - Le nouveau code travaille en mémoire et écrit les colonnes par lots.

8. **Protection Ticket fragile**
   - L’ancien code manipulait la liste d’éditeurs et pouvait retirer des droits inutilement.
   - La nouvelle protection est `warningOnly`, suffisante contre les écrasements accidentels sans bloquer le propriétaire.

## Améliorations ajoutées

- `onEdit(e)` recalcule automatiquement les tickets après modification d’un nom ou d’une présence.
- Modification de `BENEVOLES` : recalcul des mois existants.
- Contrôle qualité : présence non validée, doublon sur la même demi-journée, dépassement de 3 tickets/semaine.
- Audit de structure du classeur.
- 7 tests de non-régression sur la logique Tickets.
- Code séparé en modules lisibles.

## Limite de validation

Le code a été audité statiquement et restructuré sur GitHub. Il reste à exécuter les tests dans le projet Google Apps Script lié au classeur, car GitHub ne peut pas exécuter directement `SpreadsheetApp`, `Utilities`, les protections ou les triggers du classeur.

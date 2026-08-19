# Audit préparatoire à la modularisation de EPC :8800

> Audit local, sans modification fonctionnelle, réalisé le 19 août 2026 sur la branche
> `refonte-modulaire-8800`. La branche `stable-simple` a été consultée uniquement par
> lecture Git (`git diff`/`git log`) et n'a été ni checkoutée, ni modifiée, ni fusionnée.

## 1. État Git et commit de départ

- Dépôt : `https://github.com/Ba-Jules/epc-diagnostic-tool.git`
- Remote `origin` fetch/push : URL ci-dessus, correcte.
- Branche de production : `master` (base déclarée de `:8800`).
- Branche simple séparée : `origin/stable-simple` (base déclarée de `:8810`).
- Synchronisation exécutée : `git fetch origin --prune`, `git checkout master`, puis
  `git pull --ff-only origin master`.
- Commit de départ exact : `909236cbf26843b28b46ef01ceaf65715d0be67f` —
  « Corrige la reprise participant qui renvoyait toujours au domaine 1 sur V2 ».
- Au moment de la création de la branche, `master` et `origin/master` étaient identiques
  (`git rev-list --left-right --count master...origin/master` = `0 0`).
- Branche de chantier locale créée depuis ce commit : `refonte-modulaire-8800`.
- Aucun push, merge ou cherry-pick effectué.
- Baseline : `20/20` tests réussis avec `python -m unittest -v tests.py`.
- Seul ce fichier d'audit est ajouté au working tree ; il ne doit pas être commité à ce stade.

Branches distantes observées : `origin/master`, `origin/stable-simple`, et
`origin/HEAD -> origin/master`. Aucune autre branche distante pertinente n'est visible.

## 2. Architecture générale

L'application est un monolithe compact, sans framework applicatif :

- `app.py` (~163 Ko) : serveur HTTP `ThreadingHTTPServer`, routage API manuel dans
  `Handler`, accès SQLite en SQL brut, migrations au démarrage, moteur de calcul,
  authentification, IA, génération XLSX/DOCX/ZIP et préparation des données de rapport.
- `static/app.js` (~167 Ko) : SPA en JavaScript sans framework, navigation, formulaires,
  appels `fetch`, graphiques SVG/HTML, écrans campagne/groupe/participant/analyse/rapport.
- `static/index.html` : coquille minimale chargeant la SPA.
- `static/style.css` : styles écran/impression.
- `tests.py` : 20 tests unitaires/intégration SQLite du moteur et des invariants V2.
- `Dockerfile` : image `python:3.12-slim`, port interne 8000, volume `/app/data`.
- `docker-compose.yml` : service unique `epc-diagnostic`, volume nommé `epc-data`,
  mapping `8000:8000` (le mapping VPS vers `:8800` est donc hors dépôt).
- `requirements.txt` : `xlsxwriter`, `python-docx`, `Pillow`, tous optionnels pour le
  cœur ; le backend repose sinon sur la bibliothèque standard et SQLite.
- Documents/classeurs racine et `archive/` : références méthodologiques historiques,
  non consommées dynamiquement à l'exécution.
- Base locale : `data/workshops.sqlite3`, ou `$DATA_DIR/workshops.sqlite3`.

Il n'existe actuellement ni couche repository/service, ni ORM, ni migrations versionnées,
ni découpage backend/frontend par domaine. Le `Handler` mélange autorisation, validation,
SQL, orchestration et sérialisation. Le frontend concentre également toutes les vues dans
un fichier global.

## 3. Schéma de données

### 3.1 Tables principales (20)

| Table | Rôle et relations principales |
|---|---|
| `users` | Comptes pilotes/admins ; email unique, PBKDF2, rôle. |
| `auth_tokens` | Sessions de connexion ; hash du token -> `users.id`, expiration 14 jours. |
| `templates` | Questionnaire/version ; paramètres JSON de score, échelle, consensus, graduation et priorité ; `owner_user_id` ajouté par migration. |
| `domains` | Domaines ordonnés d'un `template`. |
| `indicators` | Questions ordonnées d'un domaine ; type, obligatoire, actif, configuration JSON. |
| `campaigns` | Campagne possédée par un utilisateur, épinglée sur `template_id` + `template_version`. |
| `sessions` | Mission autonome ou groupe d'une campagne ; questionnaire, propriétaire, campagne, code/couleur/relais, objectif. |
| `participants` | Participant d'une session ; identifiant anonyme, nom facultatif, statut et timestamps. |
| `responses` | Réponse individuelle unique par `(participant_id, indicator_id)`, avec copie de `session_id`. |
| `priorities` | Indicateur retenu comme priorité dans une session. |
| `analysis_notes` | Ancienne couche de notes qualitatives par indicateur. |
| `recommendations` | Ancienne couche de recommandations par indicateur. |
| `priority_analyses` | Constat unique par `(session, priorité)`. |
| `analysis_entries` | Causes, conséquences et leviers ; hiérarchie facultative via `parent_id`. |
| `workshop_recommendations` | Recommandations V2 reliables à priorité/cause/levier. |
| `training_topics` | Besoins de formation, reliables à priorité/recommandation. |
| `session_report_meta` | Animateur, audience, contexte et conclusion d'une session. |
| `ai_config` | Configuration IA globale singleton, y compris clé API en clair dans SQLite. |
| `ai_suggestions` | Journal minimal des suggestions et de leur statut. |
| `report_ai_blocks` | Blocs rédactionnels IA explicitement retenus dans le rapport. |

### 3.2 Relations métier

```text
user 1 ── n campaign 1 ── n session/groupe 1 ── n participant 1 ── n response
  │             │                 │                                      │
  │             └── template/version partagé par tous les groupes        └── indicator
  ├── n session autonome
  └── n template personnel

template 1 ── n domain 1 ── n indicator
session 1 ── n priority 1 ── 0..1 priority_analysis
                        └── n analysis_entry ── n workshop_recommendation
session 1 ── n workshop_recommendation ── n training_topic
session 1 ── 0..1 session_report_meta
session 1 ── n ai_suggestion / report_ai_block
```

`sessions` est le pivot réel : une campagne n'a pas de réponses propres ; ses groupes
sont des sessions avec `campaign_id`. Un questionnaire de campagne est fixé à la création
de la campagne, puis recopié dans chaque session/groupe. La consolidation agrège les
réponses individuelles de sessions sélectionnées, jamais des agrégats de groupe.

Points de dette de schéma : colonnes ajoutées par `ALTER TABLE` au démarrage sans table de
version de migration ; contraintes FK absentes sur plusieurs colonnes ajoutées
(`owner_user_id`, `campaign_id`) ; coexistence de deux modèles qualitatifs
(`analysis_notes`/`recommendations` et V2) ; JSON sans validation de schéma ; pas de profil
participant générique ; `template_version` est redondant par rapport à l'identifiant de
template et n'est pas utilisé pour reconstituer une version.

## 4. Routes/API importantes

Environ 55 formes de routes métier sont implémentées manuellement, réparties entre
`GET`, `POST`, `PUT`, `DELETE` (le nombre dépend de la manière de compter les routes
paramétrées et variantes d'export).

### Authentification et public

- `GET /api/auth/setup-status`, `GET /api/auth/me`
- `POST /api/auth/setup`, `/api/auth/login`, `/api/auth/logout`
- Public participant : `POST /api/sessions/{id}/participants`, `/responses`, `/complete`
- Reprise : `GET /api/participant?session=...&participant=...`
- Relais public : `GET /api/relay/{token}`

### Questionnaires

- `GET/POST /api/templates`, `GET/PUT/DELETE /api/templates/{id}`
- `POST /api/templates/{id}/clone`
- `POST /api/templates/{id}/domains`, `PUT/DELETE /api/domains/{id}`
- `POST /api/domains/{id}/indicators`, `PUT/DELETE /api/indicators/{id}`
- Import : `POST /api/templates/import/preview`, `/import/confirm`
- Matrices : `GET /api/templates/matrix.xlsx`, `/api/templates/{id}/matrix.xlsx`

### Missions, campagnes, groupes et relais

- `GET/POST /api/sessions`, `PUT/DELETE /api/sessions/{id}`
- `POST /api/sessions/{id}/status`
- `GET/POST /api/campaigns`, `GET/PUT/DELETE /api/campaigns/{id}`
- `GET/POST /api/campaigns/{id}/groups`
- `DELETE /api/campaigns/{id}/groups/{sessionId}`
- `POST /api/campaigns/{id}/groups/{sessionId}/regenerate-relay`
- `POST /api/campaigns/{id}/consolidate`
- `GET /api/campaigns/{id}/deletion-summary`, `/kits.zip`

### Collecte, calcul et restitution

- `POST /api/sessions/{id}/participants`, `/responses`, `/complete`
- `PUT /api/participants/{id}` (nom affiché)
- `GET /api/sessions/{id}/analysis`, `/workshop-data`, `/qualitative-data`, `/report-data`
- Priorités/qualitatif : sous-routes `priorities`, `priority-analyses`, `analysis-entries`,
  `recommendations-v2`, `training-topics`, `report-meta`, plus anciennes
  `analysis-notes` et `recommendations`.
- Exports : `/export.json`, `/responses.csv`, `/report.xlsx`, `/report.docx` ; PDF via
  impression navigateur.
- IA : configuration/test, diagnostic, préparation de priorité, entrées, recommandations,
  formations, plan et blocs/sections de rapport.

## 5. Comptes, authentification et cloisonnement

- Premier démarrage : création d'un premier compte `admin` via `/api/auth/setup`.
- Mot de passe : PBKDF2-HMAC-SHA256, sel aléatoire, 200 000 itérations.
- Cookie : token aléatoire stocké hashé, `HttpOnly; SameSite=Lax`, durée 14 jours.
  Pas de `Secure`, choix documenté car le VPS sert actuellement en HTTP.
- `is_public_api()` limite l'accès sans compte aux opérations participant et relais.
- `require_auth()` résout l'utilisateur puis vérifie l'ownership des IDs présents dans
  l'URL pour sessions, templates et campagnes. Un admin contourne l'ownership.
- Les listes sont filtrées par propriétaire pour les non-admins ; les questionnaires
  communs ont `owner_user_id IS NULL`.
- Les groupes héritent de l'utilisateur créateur et sont liés à une campagne possédée.
- Les tests couvrent refus inter-pilotes, suppression isolée et campagnes/groupes
  homonymes.

Ressources publiques : le participant connaît `session_id` et `participant_id`; le relais
connaît un token secret dont seul le SHA-256 est stocké. La route relais n'expose que nom
de campagne/groupe/relais, objectif, commencés, validés, lien participant. Régénérer le
token invalide l'ancien. Les participants n'accèdent pas aux analyses ni aux autres
réponses.

Risques actuels : IDs participant/session font office de capability sans jeton participant
dédié ; absence de CSRF explicite au-delà de SameSite ; cookies non sécurisés sur HTTP ;
clé IA globale en clair ; autorisation dispersée et dépendante du parsing de chemin ; les
mutations SQL doivent rester systématiquement précédées du garde générique.

## 6. Modèle métier actuel : configurable ou rigide

| Élément | État | Détails/dépendances |
|---|---|---|
| EPC/SENEVAL | Codé en dur + semé en DB | `EPC_DOMAINS` contient 7 domaines × 10 indicateurs ; `seed_epc()` crée le modèle nommé exactement `EPC / SENEVAL`. Nom utilisé aussi pour protection/sémantique UI. |
| Domaines | Configurables | CRUD, ordre, activation, description ; mais couleurs UI limitées à 7 variables et textes de rapport EPC. |
| Indicateurs | Partiellement configurables | CRUD/import, type/config JSON ; collecte et calcul n'acceptent réellement que des valeurs numériques. |
| Échelle | Partiellement configurable | min/max/libellés stockés en JSON et importables ; frontend et distributions supposent une échelle entière contiguë. |
| Capacité | Méthode quasi codée en dur | `scoring_json` annonce une méthode, mais `analysis_for()` applique toujours `mean / high * output_max` sans dispatcher sur `scoring.capacity`. Le minimum de l'échelle n'entre pas dans la standardisation. |
| Consensus | Partiellement paramétré | facteur lu depuis JSON, mais méthode écart-type d'échantillon et normalisation par amplitude codées dans `analysis_for()`. |
| Graduation | Configurable en données, rigide en moteur | bandes JSON issues de `GRADING`; `grade()` arrondit puis cherche une plage. Pas de stratégie alternative. |
| Priorités | Partiellement configurable | `maxPerDomain` en JSON ; workflow, sémantique par indicateur et écrans codés. |
| Questionnaire canonique | Identifié par nom, pas par flag | Aucun `is_canonical` sur `master`; toute ligne nommée `EPC / SENEVAL` est protégée/sémantique. |
| Restitution | EPC-spécifique | Titres, méthodologie, niveaux, capacité/consensus, sections et prompts IA mentionnent EPC/SENEVAL. |

Le nom « moteur universel » du README est donc anticipatoire : la structure domaine/
indicateur est configurable, mais la sémantique de calcul et de restitution demeure EPC.

## 7. Questionnaires et versionnement

### Fonctionnement de référence sur `master`

`seed_epc()` crée la V1 avec échelle 1–5, capacité standardisée 0–100, consensus par
écart-type, graduation KOICA et maximum de 3 priorités par domaine. À chaque ouverture de
base par `Handler.db()`, `init_db()` exécute `ensure_reference_questionnaire_version()`.

Cette fonction :

1. cherche la version au numéro le plus élevé dont le nom est `EPC / SENEVAL` ;
2. compare seulement la liste des codes de domaines et le nombre d'indicateurs par domaine
   à `EPC_DOMAINS` ;
3. si écart, crée une nouvelle ligne `templates` avec `MAX(version)+1` ;
4. recrée les 7×10 domaines/indicateurs ;
5. ne modifie jamais les anciennes lignes, réponses ou sessions.

Une session/campagne stocke `template_id` et `template_version`. Une campagne choisit un
template précis à sa création. Tous ses groupes recopient exactement ce couple ; la route
de consolidation refuse des groupes aux couples différents. Les anciennes campagnes
restent donc épinglées à leur version historique.

### Modification et suppression

- Modifier un template déjà utilisé par une session appelle `clone_template()` avant la
  modification : nouvelle version/ID ; l'ancien reste intact.
- Ajouter/modifier directement domaines ou indicateurs ne déclenche toutefois pas ce fork
  dans ces routes : la protection est principalement au niveau PUT du template et via les
  refus de suppression en présence de réponses. C'est un point à verrouiller avant refactor.
- Supprimer un indicateur avec réponses est refusé ; supprimer un domaine contenant des
  réponses est refusé ; un template utilisé est refusé ou archivé avec `force=1`.
- Tout template nommé `EPC / SENEVAL` est non supprimable.
- Il n'y a pas de colonne `is_canonical` sur `master`.

### Différences observées avec `stable-simple` (lecture uniquement)

`stable-simple` a introduit `is_canonical`, sépare initialisation de schéma et migrations,
forke les questionnaires des missions ayant commencé, et met à jour le canonique par ID.
Elle contient aussi des colonnes de profil participant et des analyses filtrées/
désagrégées. Elle ne contient plus campagnes/groupes/relais. Ces idées signalent des
problèmes déjà rencontrés, mais son architecture ne doit pas être copiée ni cherry-pickée :
la V2 `master` doit conserver ses invariants multi-groupes et recevoir un modèle générique
conçu explicitement.

## 8. Campagnes, groupes, relais et participants

### Campagne et groupes

- `campaigns` porte propriétaire, libellé/période, statut et questionnaire figé.
- Un groupe est une `session` : `campaign_id`, `group_code`, `group_color`, `relay_name`,
  `relay_token_hash`, objectif facultatif.
- `generate_group_code()` génère un code globalement unique ; le code n'est jamais utilisé
  comme clé de sélection.
- L'objectif `expected_participants` est indicatif, nullable et jamais un plafond.
- Progression = validés/objectif ; commencés et validés sont comptés séparément.
- Kits ZIP : un HTML par groupe ; leur génération régénère les tokens relais et invalide
  donc les anciens liens (effet important à préserver/documenter dans l'UX).
- Consolidation : sélection d'au moins un groupe de la même campagne et du même template ;
  recalcul direct via `analysis_for(session_ids)`.

### Participants

- Données : `id`, `session_id`, `anonymous_id`, `status`, `started_at`, `completed_at`,
  `display_name` facultatif.
- Aucun sexe, âge, profil, organisation structurée, champ personnalisé ou métadonnée JSON
  sur `master`.
- L'anonymat est un choix UI : absence de `display_name`; il n'existe pas de booléen
  explicite permettant de distinguer anonyme volontaire et nom non saisi.
- Statuts utilisés : `in_progress` puis `completed`.
- Reprise : stockage côté navigateur des IDs ; `GET /api/participant` renvoie participant,
  session, template et réponses, puis le frontend retrouve le premier domaine incomplet
  (correctif du commit de départ).
- Les réponses sont upsertées individuellement et restent rattachées à session/groupe.
- Un participant validé peut exporter sa propre copie depuis le frontend.

La cible « profil composable » est absente. Elle nécessite définitions de champs versionnées,
valeurs participant typées, consentement/anonymat explicite et politique de restitution.

## 9. Calculs

Le calcul unique est `analysis_for(db, session_ids)` ; `analysis()` est un wrapper pour une
session. Il faut conserver ce chemin unique.

1. Charge le template de la première session.
2. Pour chaque indicateur actif, sélectionne uniquement les réponses dont le participant
   est `completed`.
3. Convertit uniquement les valeurs JSON numériques.
4. Moyenne brute par indicateur.
5. Capacité indicateur = `moyenne / max_échelle × max_sortie`.
6. Écart-type échantillonnal (`n-1`) si N > 1.
7. Consensus = `max(0, max_sortie - facteur × écart-type / amplitude × max_sortie)`.
8. Pour un domaine, calcule d'abord la moyenne de chaque participant sur ses indicateurs,
   puis capacité/consensus entre participants.
9. Global standardisé = moyenne non pondérée des scores des domaines.
10. Global gradué = moyenne non pondérée des scores déjà gradués des domaines, et non
    graduation du score global.

Cas particuliers :

- N=1 : capacité calculée, consensus `null`, note `single_respondent`.
- Incomplets : exclus de tous les scores, mais inclus dans `participantCount` (commencés).
- Données manquantes : comptées par rapport à tous les participants commencés, ce qui peut
  inclure des incomplets alors que le score les exclut.
- Consolidation : pool de réponses individuelles complètes, jamais moyenne des groupes.
- Le moteur suppose implicitement que tous les `session_ids` utilisent le même template ;
  la route campagne le vérifie, mais la fonction elle-même non.
- Les règles JSON nomment les méthodes mais ne les sélectionnent pas réellement.

Tests existants essentiels : scores/réponses brutes, N=1, exclusion incomplets,
consolidations A/B, scénarios multi-groupes, homonymes et isolation massive.

## 10. Analyses, graphiques et exports

- Diagnostic : global/domaines/indicateurs, capacité, consensus, gradués, distributions.
- Frontend : barres, radar, grille positionnement capacité-consensus et cohorte/graduation.
- Priorités humaines par indicateur, limitées selon le template.
- Chaîne qualitative : constat -> causes/conséquences/leviers -> recommandations -> thèmes
  de formation -> plan d'action.
- Synthèse/rapport : contexte, méthodologie, participation, diagnostic, forces/faiblesses,
  priorités, qualitatif, recommandations, formations, plan, conclusion, annexes.
- Comparaison disponible au niveau groupes dans les écrans de campagne ; consolidation de
  groupes sélectionnés. Pas de désagrégation par attribut participant sur `master`.
- Exports : JSON complet, réponses CSV, rapport Excel, Word avec graphiques Pillow,
  impression/PDF navigateur, matrice questionnaire XLSX, kits relais ZIP.
- IA facultative multi-fournisseur : ne recalcule jamais les scores ; suggestions non
  sauvegardées avant validation humaine ; blocs retenus distincts des résultats EPC.

Couplages forts : fonctions de graphiques/export connaissent capacité/consensus/graduation ;
méthodologie et libellés sont EPC ; prompts IA imposent la lecture EPC ; données du rapport
sont assemblées directement depuis le schéma SQL courant.

## 11. Rigidités empêchant un moteur générique

1. `EPC_DOMAINS`, `GRADING`, échelle et seed sont dans `app.py`.
2. Identité canonique déduite du nom `EPC / SENEVAL`.
3. Une seule famille de score exécutée malgré les configurations JSON.
4. Réponses supposées numériques dans l'analyse ; `response_type` n'a pas de moteur typé.
5. Frontend, aide, titres, niveaux, rapports, noms de fichiers et prompts IA parlent EPC.
6. Couleurs et plusieurs visualisations supposent sept domaines et capacité/consensus.
7. Profil participant inexistant et aucun axe générique de segmentation.
8. Anonymat implicite par nom vide, non modélisé comme politique de campagne.
9. Campagne/groupe/relais sont structurés, mais options de collecte non modélisées comme
   configuration/version (mode simple, groupes, relais, anonymat, objectif).
10. Routage, SQL, ownership et métier sont imbriqués dans `Handler`.
11. Le schéma JSON n'est ni versionné ni validé ; les algorithmes ne sont pas enregistrés
    par identifiant stable.
12. Versionnement questionnaire incomplet au niveau des sous-ressources.
13. Deux modèles qualitatifs coexistent.
14. Les analyses/export accèdent au résultat EPC concret, sans contrat de résultat neutre.

La cible doit faire de « EPC/SENEVAL » un paquet de modèle fourni : questionnaire par
défaut + règles de calcul + graduations + visualisations compatibles + textes de méthode,
et non une condition implicite disséminée dans l'application.

## 12. Invariants à préserver absolument

- Historique : aucune modification rétroactive des domaines/indicateurs/réponses d'une
  campagne ayant commencé ; campagne toujours épinglée à une révision immuable.
- Tous les groupes d'une campagne partagent exactement son questionnaire.
- Consolidation recalculée depuis les réponses individuelles, jamais moyenne de moyennes.
- Participants incomplets exclus des scores mais visibles dans les compteurs.
- N=1 : consensus non calculable, capacité conservée, signal explicite.
- Formules EPC, arrondi de graduation et moyenne globale non pondérée inchangés tant que le
  modèle sélectionné est EPC.
- Isolation pilote A/B et isolation entre campagnes/groupes homonymes.
- Sélection par IDs, jamais par nom ou code affiché.
- Tokens relais stockés hashés, accès minimal et régénération invalidante.
- Objectif nullable, indicatif et sans plafond.
- Anonymat participant et impossibilité de consulter les réponses des autres.
- Chaîne qualitative et validation humaine des priorités/causes/recommandations/IA.
- Exports existants et rapport terrain, y compris mission réelle et historique.
- Suppressions en cascade strictement limitées à la campagne/session cible ; questionnaire
  partagé et données des autres pilotes préservés.
- `stable-simple` et le service `:8810` restent totalement indépendants.

## 13. Cible modulaire proposée

Séparer progressivement cinq concepts sans réécriture :

1. **Définition de modèle** : identité stable, révision immuable, domaines/questions,
   types de réponse, échelle, stratégies de calcul, graduations, visualisations et textes.
2. **Définition de profil** : champs typés/versionnés, options, cardinalité, requis,
   sensibilité et autorisation comme axe analytique.
3. **Configuration de collecte** : mode simple/groupes/relais, anonymat, objectif, profil,
   questionnaire/modèle épinglé.
4. **Moteur d'analyse** : stratégies enregistrées par identifiant, entrée normalisée,
   résultat neutre, filtres/axes génériques, invariants EPC couverts par tests de référence.
5. **Restitution** : composants sélectionnés selon capacités du modèle, sans supposer EPC.

Les données catégorielles du profil doivent produire automatiquement des dimensions
filtrables/comparables, mais uniquement si le champ l'autorise et si les seuils de
confidentialité sont respectés. Texte libre et nombres ne doivent pas devenir naïvement
des axes catégoriels.

## 14. Stratégie incrémentale de refactoring

### Lot 0 — Baseline, caractérisation et garde-fous

- Objectif : figer le comportement de `master` avant extraction.
- Fichiers : `tests.py`, futurs fixtures/snapshots ; pas de changement métier.
- Tables/routes : toutes en lecture ; priorité à campagnes, templates, collecte, calculs,
  exports et ownership.
- Ajouter : golden tests EPC sur petit jeu connu, contrats JSON des routes, empreintes des
  exports, tests directs d'autorisation pour chaque verbe, test questionnaire partagé.
- Risque : faible ; dépendance : aucune.
- Rollback : retrait des seuls tests/outils.

### Lot 1 — Extraire les couches techniques sans changer les contrats

- Objectif : sortir de `app.py` connexion/migrations, repositories, auth/ownership et
  services campagne/questionnaire, en conservant routes et SQL équivalents.
- Fichiers : scinder vers `epc/db.py`, `epc/auth.py`, `epc/repositories/*`,
  `epc/services/*`, laisser `app.py` comme composition/compatibilité.
- Tables/routes : toutes, sans changement de forme ni migration fonctionnelle.
- Risque : moyen/élevé à cause des gardes dispersés.
- Tests : suite complète + matrice accès public/pilote/admin + démarrage sur copie d'une DB.
- Rollback : commits d'extraction petits et indépendants.

### Lot 2 — Révisions immuables et catalogue de modèles

- Objectif : identité canonique stable, séparation modèle/révision, EPC comme modèle fourni,
  questionnaires vierges et futurs CAD/OCDE.
- Tables : introduire prudemment `model_definitions`/`model_revisions` ou enrichir
  `templates` avec `model_key`, `revision`, `is_builtin`, `is_canonical`; migration
  additive et mapping des IDs existants.
- Routes : façade `/api/templates` conservée, nouvelles routes catalogue internes.
- Fichiers : service questionnaire/versionnement, paquet `models/epc_seneval.*`.
- Risque : critique pour historique et FK.
- Tests : aucune mutation d'une révision utilisée, idempotence migration, campagnes
  historiques inchangées, 7×70 exact, comparaison avant/après des scores/exports.
- Dépendance : lots 0–1.
- Rollback : double lecture temporaire et colonnes additives ; ne supprimer l'ancien
  chemin qu'après validation sur copie de production.

### Lot 3 — Stratégies de calcul et contrat de résultat

- Objectif : encapsuler EPC sans modifier ses formules, puis permettre d'autres méthodes.
- Fichiers : `analysis/registry.py`, `analysis/epc.py`, contrats typés de résultat ; adapter
  `analysis_for()` en façade.
- Données : identifiants de stratégie versionnés dans la révision de modèle, paramètres
  validés ; aucune formule arbitraire exécutable en DB.
- Routes : `/analysis`, consolidation et rapports conservent leur contrat initial via un
  adaptateur EPC.
- Risque : critique (mathématiques et consolidation).
- Tests : golden tests décimaux, N=0/N=1/N>1, incomplets, manquants, plusieurs groupes,
  graduation aux bornes, ordre/pondération des domaines.
- Dépendance : lot 2.
- Rollback : feature flag par modèle, ancien moteur EPC disponible jusqu'à parité parfaite.

### Lot 4 — Profil participant composable et politique d'anonymat

- Objectif : définir champs `single_choice`, `multi_choice`, `text`, `number`, etc., et
  stocker valeurs typées sans casser les participants existants.
- Tables : `profile_schemas`, `profile_field_definitions`, `participant_profile_values`
  (ou JSON validé avec index secondaire), révision épinglée à campagne ; booléen/politique
  d'anonymat explicite.
- Routes : définition du profil, formulaire participant dynamique, lecture contrôlée.
- Fichiers : backend validation/normalisation, composants de formulaire frontend.
- Risque : élevé (vie privée, reprise, compatibilité).
- Tests : migrations participants existants, obligatoire/facultatif, anonymat, reprise,
  types/cardinalité, aucune fuite relais/export.
- Dépendance : lots 1–2.
- Rollback : profil désactivé par défaut, tables additives, rendu historique inchangé.

### Lot 5 — Filtrage, désagrégation et comparaison génériques

- Objectif : tout champ catégoriel autorisé devient dimension analytique ; filtres et
  comparaisons réutilisent le même moteur et les réponses individuelles.
- Tables : métadonnées `is_dimension`, ordre/libellés d'options, éventuels index ; pas de
  duplication d'agrégats au départ.
- Routes : extension additive de `/analysis` et `/consolidate` avec filtres/dimensions ;
  endpoint de dimensions disponibles.
- Fichiers : query builder sûr, cohortes, UI filtres/comparaison, exports filtrés.
- Risque : critique pour confidentialité et biais petits N.
- Tests : multi-choix, valeur absente, groupes + sous-populations, incomplets exclus,
  parité sans filtre, interdiction cross-campaign, seuil minimal configurable.
- Dépendance : lots 3–4.
- Rollback : paramètres opt-in ; chemin sans filtre inchangé.

### Lot 6 — Restitution et IA modulaires

- Objectif : graphiques/sections/export/prompts choisis selon les capacités du modèle ;
  conserver intégralement le rapport EPC actuel.
- Fichiers : composants de restitution, renderers XLSX/DOCX/web, assembleur de rapport,
  prompts IA construits depuis métadonnées du modèle.
- Routes : `report-data` enrichi d'un manifeste de restitution ; exports gardent URL et
  formats.
- Tables : éventuelles configurations de vue/rapport par révision, blocs IA typés.
- Risque : élevé, surtout mise en page et valeur terrain.
- Tests : snapshots structuraux/visuels, classeurs ouvrables, DOCX valide, PDF imprimable,
  IA désactivée/indisponible, aucun recalcul IA, exports anonymisés.
- Dépendance : lots 2–5.
- Rollback : renderer EPC historique sélectionné pour `model_key=epc_seneval`.

### Lot 7 — UX catalogue/configurateur et dépréciation contrôlée

- Objectif : parcours « modèle EPC », « CAD/OCDE », « vierge », profil, collecte et analyse
  compréhensibles, puis suppression graduelle des anciennes couches qualitatives.
- Fichiers : découpage `static/app.js` en modules, aide contextuelle, migrations de façade.
- Tables/routes : catalogue, assistants de configuration ; dépréciation documentée de
  `analysis_notes`/`recommendations` seulement après migration vérifiée.
- Risque : moyen/élevé (workflow utilisateurs).
- Tests : E2E pilote/relais/participant, clavier/mobile/impression, création complète de
  chaque modèle, reprise et suppression isolée.
- Dépendance : tous les lots précédents.
- Rollback : anciens écrans conservés derrière flag jusqu'à recette terrain.

## 15. Risques principaux

1. **Historique questionnaire** : une édition de sous-ressource peut encore altérer un
   template utilisé ; toute nouvelle abstraction doit rendre les révisions immuables.
2. **Parité mathématique** : rendre configurable peut modifier subtilement normalisation,
   écart-type, arrondi, pondération ou traitement N=1/incomplets.
3. **Cloisonnement** : extraction des routes ou ajout de filtres peut contourner le garde
   basé sur le chemin et mélanger pilotes/campagnes.
4. **Vie privée** : désagrégation sur petits effectifs ou profils sensibles peut réidentifier
   un participant, même anonyme.
5. **Migrations** : `init_db()` s'exécute à chaque ouverture de DB/requête et les migrations
   ne sont pas versionnées ; il faut séparer démarrage, schéma et migration avant évolution.
6. **Monolithe frontend/backend** : nombreux globals et redéfinitions de fonctions JS ; une
   extraction massive rendrait les régressions difficiles à localiser.
7. **Double modèle qualitatif** : risque de perdre notes/recommandations anciennes lors de
   la convergence.
8. **Exports/IA** : consommateurs directs du résultat EPC et du schéma ; à adapter après un
   contrat stable, pas avant.

Le risque architectural principal est la combinaison « questionnaire/version + calcul » :
les données historiques référencent des IDs de domaines/indicateurs tandis que les règles
exécutées sont partiellement codées dans le monolithe. Une migration mal séquencée peut
préserver les réponses brutes tout en changeant leur interprétation ou leur restitution.

## 16. Tests de non-régression indispensables

- Conserver les 20 tests actuels comme plancher.
- Golden dataset EPC 7×70 avec résultats indicateur/domaine/global attendus à précision
  fixée et graduations à toutes les bornes.
- N=0, N=1, N=2+, réponses manquantes, incomplets, reprises et sur-réponses à objectif.
- Consolidation 1/N groupes : résultat identique au recalcul du pool individuel ; ordre des
  groupes sans effet ; rejet de modèles/révisions différents.
- Campagnes et groupes homonymes, cinq pilotes/campagnes, accès croisés GET/POST/PUT/DELETE.
- Tous les groupes créés avant/après modification utilisent la révision attendue ; ancienne
  campagne byte-for-byte/logiquement inchangée.
- Création, clone, import, modification, archivage/suppression de templates avec/sans usage.
- Token relais valide/invalide/régénéré, contenu minimal, kits ZIP et invalidation annoncée.
- Suppression forcée campagne/groupe : inventaire exact et aucune donnée hors cible.
- Participant anonyme/nommé, reprise au premier domaine incomplet, copie personnelle,
  impossibilité de lire une autre réponse.
- Exports JSON/CSV/XLSX/DOCX et rapport imprimé ; noms, feuilles, sections et valeurs.
- Qualitatif complet et relations cause/levier/recommandation/formation.
- IA désactivée, erreur fournisseur, validation humaine, absence de clé dans GET/export/log,
  scores identiques avec IA activée ou non.
- À partir du lot profil : validation de chaque type, multi-choix, migrations, dimensions,
  cohortes vides/petites, seuil anti-réidentification et parité analyse sans filtre.

## 17. Ordre recommandé pour démarrer

Claude Code devrait commencer par le lot 0, puis le lot 1 sans modifier les contrats. Le
premier changement de données doit être la fondation additive et testée du lot 2. Il ne
faut pas commencer par le configurateur UI, la désagrégation ou les rapports : ils
dépendent d'une identité de modèle/révision et d'un contrat de calcul stabilisés.

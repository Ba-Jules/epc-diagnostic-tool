# Moteur universel de diagnostic EPC / SENEVAL

Application locale de préparation, collecte et restitution d’ateliers de diagnostic organisationnel EPC/SENEVAL.

## Fonctionnalités principales

- Questionnaires versionnés et configurables, avec import/export de matrice XLSX.
- Sessions d’atelier et collecte multi-participants.
- Calcul des scores de capacité, consensus et notes graduées.
- Restitution graphique, sélection humaine des priorités et analyse qualitative.
- Recommandations, thèmes de formation, plan d’action et rapport final.
- Exports XLSX et Word ; le rapport écran (« Imprimer / Télécharger en PDF ») s'imprime ou s'enregistre en PDF depuis le navigateur, avec la même mise en forme.

## Prérequis

- Python 3.10 ou version plus récente.
- Le moteur lui-même (`app.py`) ne dépend que de la bibliothèque standard Python et de SQLite : il démarre sans rien installer de plus.
- Les exports XLSX et Word nécessitent respectivement `xlsxwriter` et `python-docx`. Sans ces paquets, l'application fonctionne normalement ; seuls les boutons d'export correspondants sont indisponibles. Le PDF ne dépend d'aucun paquet : il passe par l'impression du navigateur.
- `Pillow` est utilisé pour dessiner les graphiques insérés dans les exports Word et Excel. Sans lui, ces exports restent disponibles mais sans graphiques (texte et tableaux uniquement).

Pour disposer de tous les exports, installer ces deux paquets optionnels :

```powershell
python -m pip install -r requirements.txt
```

## Lancement

Depuis le dossier du projet :

```powershell
python app.py
```

Ouvrir ensuite [http://127.0.0.1:8000](http://127.0.0.1:8000).

## Assistant IA optionnel

Un assistant IA facultatif peut être activé depuis Configuration (« 5 · Assistant IA »). Il est désactivé par défaut, et **l'application fonctionne intégralement sans lui** — aucun calcul EPC (capacité, consensus, standardisation, graduation) ne dépend de l'IA.

- **Fournisseurs pris en charge** : Google Gemini, Groq, OpenRouter, Cerebras, OpenAI, Anthropic Claude, DeepSeek, xAI Grok. Chacun est étiqueté GRATUIT / ESSAI / PAYANT selon l'accès **API** réel (à distinguer d'un chatbot gratuit, dont l'API peut être payante). Ces statuts et les URL de création de clé reflètent l'état connu à la mise en place de la fonctionnalité et méritent une vérification ponctuelle si un fournisseur change ses conditions.
- **Clé API** : saisie une seule fois dans Configuration, elle est stockée côté serveur dans la base SQLite et n'est **jamais** renvoyée par l'API, jamais présente dans le HTML/JS servi au navigateur, jamais visible du participant, jamais journalisée, jamais exportée. Toute requête passe par le serveur (`Navigateur → Backend → Fournisseur IA`), jamais directement du navigateur vers le fournisseur.
- **Points d'usage** : lecture croisée capacité/consensus du diagnostic, préparation de l'analyse d'une priorité, hypothèses de causes/conséquences/leviers, recommandations (fondées uniquement sur les causes et leviers déjà validés par le groupe), besoins de formation, structuration du plan d'action, et synthèse rédactionnelle du rapport final (par section ou en une fois).
- **Aucune automatisation** : chaque suggestion s'affiche dans un encart « ✦ » distinct, modifiable avant d'être retenue. Rien n'est enregistré tant que le modérateur ne clique pas explicitement sur « Retenir » — l'IA ne valide ni ne sélectionne jamais de cause, de recommandation ou de priorité à la place du groupe.
- **Résilience** : en cas d'indisponibilité, de clé invalide ou de quota atteint, un message clair s'affiche (« Assistant IA momentanément indisponible. Vous pouvez poursuivre l'atelier normalement. ») sans jamais interrompre l'atelier.
- **Rapport final** : les textes IA retenus apparaissent dans une section « Synthèse assistée par IA » distincte des données et scores EPC, aussi bien à l'écran que dans les exports Word (section dédiée) et Excel (feuille « Synthèse_IA »).

## Structure

- `app.py` : serveur local, base SQLite, API et calculs.
- `static/` : interface web, graphiques et styles d’impression.
- `tests.py` : tests automatisés du moteur.
- `archive/` et les classeurs/documents à la racine : références méthodologiques et historiques EPC/SENEVAL.
- `data/` : données locales d’exécution, non versionnées.

La recette terrain avec les utilisateurs et ateliers réels reste à effectuer.

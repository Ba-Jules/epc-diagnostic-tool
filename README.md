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

## Structure

- `app.py` : serveur local, base SQLite, API et calculs.
- `static/` : interface web, graphiques et styles d’impression.
- `tests.py` : tests automatisés du moteur.
- `archive/` et les classeurs/documents à la racine : références méthodologiques et historiques EPC/SENEVAL.
- `data/` : données locales d’exécution, non versionnées.

La recette terrain avec les utilisateurs et ateliers réels reste à effectuer.

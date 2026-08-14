# Moteur universel de diagnostic EPC / SENEVAL

Application locale de préparation, collecte et restitution d’ateliers de diagnostic organisationnel EPC/SENEVAL.

## Fonctionnalités principales

- Questionnaires versionnés et configurables, avec import/export de matrice XLSX.
- Sessions d’atelier et collecte multi-participants.
- Calcul des scores de capacité, consensus et notes graduées.
- Restitution graphique, sélection humaine des priorités et analyse qualitative.
- Recommandations, thèmes de formation, plan d’action et rapport final.
- Exports XLSX, Word et PDF ; le rapport écran peut être imprimé ou enregistré en PDF depuis le navigateur.

## Prérequis

- Python 3.10 ou version plus récente.
- Les bibliothèques `xlsxwriter`, `python-docx` et `reportlab` sont nécessaires pour les exports associés.

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

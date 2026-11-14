# Note sur l'utilisation de l'IA

## Outil utilisé

J'ai utilisé Claude Code, l'assistant de programmation d'Anthropic, directement dans VS Code. Il peut lire les fichiers du projet, écrire du code, l'exécuter et lire les résultats, ce qui permet de travailler par allers-retours rapides.

## Répartition du travail

J'ai défini ce que l'outil devait faire, à partir de l'énoncé, et l'IA a écrit le premier jet :

- les trois fichiers Excel synthétiques et le tableau `anomalies_attendues.xlsx` ;
- `app.py`, qui contient les paramètres, les contrôles, le reporting Excel et l'application Streamlit ;
- `test_controls.py`, le test de bout en bout ;
- une première version du README.

J'ai ensuite relu le code, vérifié les seuils et les sévérités dans les paramètres, contrôlé les données et testé l'outil. J'ai aussi fait modifier plusieurs points, détaillés plus bas.

## Prompts principaux

### 1. Le démarrage

> « Crée un environnement virtuel. J'ai mis dans consigne.txt ce que je dois faire. Fais quelque chose de simple et crée les fichiers Excel comme demandé. »

L'énoncé m'a servi de cahier des charges. Pour les données, j'ai demandé :

- les trois fichiers Excel avec les colonnes de l'énoncé :
  - `positions.xlsx` : ValuationDate, Portfolio, ISIN, SecurityName, AssetType, Currency, Nominal, BookPrice, MarketPrice, Rating, MaturityDate ;
  - `previous_prices.xlsx` : ISIN, PreviousDate, PreviousPrice, PriceSource ;
  - `cashflows.xlsx` : Portfolio, ISIN, CashflowDate, CashflowType, Currency, Amount, Status ;
- plusieurs portefeuilles, devises, types d'actifs (CLO Debt, CLO Equity, ABS), ratings et maturités ;
- des erreurs introduites volontairement :
  - un ISIN manquant ;
  - une position en double ;
  - un prix de marché manquant ;
  - un rating manquant ;
  - des nouveaux titres sans historique de prix ;
  - des variations de prix importantes, une par type d'actif ;
  - des nominaux nuls ou négatifs ;
  - des prix nuls ou négatifs ;
  - un cashflow sans position correspondante ;
  - un cashflow en double ;
  - des cashflows dont la date est passée mais qui sont restés en Forecast ;
- les cases en erreur surlignées en jaune fluo, pour les repérer dès l'ouverture des fichiers ;
- un fichier `anomalies_attendues.xlsx` qui liste chaque erreur introduite, sa ligne et le résultat attendu du contrôle.

J'ai demandé directement les fichiers Excel, sans garder de script pour les fabriquer. Un programme de génération de données aurait alourdi le projet sans rien apporter au livrable : ce qui compte, ce sont les fichiers et la liste des anomalies attendues.

### 2. Le regroupement du code

> « Mets tout dans un seul fichier Python, avec en commentaire chaque point de la consigne : charger et valider les inputs, vérifier les colonnes obligatoires, etc. Supprime les fichiers inutiles. »

## Ce que j'ai corrigé ou challengé

### Mes choix

1. L'organisation du code. L'IA avait découpé l'outil en quatre fichiers qui s'appellent entre eux : paramètres, contrôles, reporting et application. C'est une organisation classique, mais difficile à suivre pour quelqu'un qui découvre le projet. J'ai préféré un seul fichier, `app.py`, qui suit l'ordre de l'énoncé, avec chaque phrase de la consigne en commentaire au-dessus du code qui la traite. On peut ainsi vérifier point par point que tout est couvert, et retrouver rapidement l'endroit à modifier.
2. Les paramètres regroupés en haut du fichier. Seuils de variation, sévérités, colonnes obligatoires : tout est au même endroit. Changer le seuil de la CLO Equity ou ajouter une colonne à contrôler se fait en modifiant une seule ligne.
3. La suppression du script de génération des données, pour la raison donnée plus haut.

### Les erreurs de l'IA

Elles ont été repérées en ouvrant les fichiers produits, pas en relisant le code.

1. Le format des dates. Le code tournait sans erreur, mais dans Excel les dates s'affichaient `2026-11-12 00:00:00` au lieu de `12/11/2026` : l'option de format utilisée n'avait aucun effet. Le format est maintenant appliqué cellule par cellule. C'est le meilleur exemple de ce qu'il faut vérifier soi-même : pas de message d'erreur, seul le fichier montrait le problème.
2. La répartition des anomalies. Dans la première version, les quatre anomalies de cashflows tombaient toutes dans le même portefeuille. Le filtre par portefeuille ne pouvait donc pas être vraiment testé. Elles ont été réparties sur les quatre portefeuilles.
3. Un chiffre mal affiché. Dans l'onglet Data Quality, le nombre de lignes contrôlées s'affichait `38.0` au lieu de `38`.

## Comment j'ai vérifié les résultats

- Le test automatique (`python test_controls.py`). Chaque anomalie introduite est notée dans `anomalies_attendues.xlsx`, avec sa ligne et le résultat attendu. Le test vérifie que l'outil retrouve exactement ces anomalies, sans en oublier ni en ajouter. Résultat : 17 trouvées sur 17 attendues.
- Des cas pièges. Trois lignes ne doivent pas être signalées : une CLO Equity qui monte de 10 % (seuil à 15 %), une dette qui monte de 5 % pile (l'énoncé dit « supérieures à 5 % ») et un même titre détenu par deux portefeuilles, ce qui est normal. Elles montrent que l'outil ne remonte pas de fausses alertes.
- Un calcul refait à la main. Un prix qui passe de 98,00 à 104,37 donne +6,50 %, la valeur affichée dans le reporting.
- Les fichiers Excel relus : onglets, filtres, formats, couleurs. C'est cette relecture qui a révélé l'erreur sur les dates.
- Les cases jaunes des fichiers d'entrée comparées au reporting : chaque erreur surlignée se retrouve dans l'onglet Exceptions.
- L'application testée. Avec le filtre Critical, 7 anomalies sur 17. En abaissant le seuil de la CLO Equity à 9 %, la hausse de 10 % est signalée et on passe à 18 anomalies : le paramétrage fonctionne.
- Des fichiers abîmés exprès : colonne obligatoire supprimée, texte à la place d'un prix, fichier de cashflows vide. L'outil affiche un message clair au lieu de planter.

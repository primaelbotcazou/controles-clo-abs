# Outil de contrôle des positions, prix et cashflows CLO / ABS

Cet outil charge trois fichiers Excel, contrôle la qualité des données avec Python, puis affiche les résultats dans une application Streamlit et dans un reporting Excel.

Toutes les données sont synthétiques : les noms de deals, les ISIN, les prix et les cashflows sont inventés.

Application en ligne : URL à compléter

Code source : https://github.com/primaelbotcazou/controles-clo-abs


## Comment ça marche

1. L'utilisateur charge les trois fichiers Excel : positions.xlsx, previous_prices.xlsx et cashflows.xlsx. Il peut aussi cocher « Utiliser les fichiers d'exemple » pour prendre ceux du dossier data.
2. L'application vérifie que chaque fichier se lit bien et contient toutes les colonnes obligatoires. S'il manque une colonne, elle dit laquelle et bloque le traitement.
3. L'utilisateur peut modifier les seuils de variation de prix et l'horizon des cashflows dans la barre de gauche.
4. Il clique sur « Lancer les traitements ». L'outil fait tous les contrôles et classe chaque anomalie en Critical, Warning ou Information.
5. Les résultats s'affichent dans cinq onglets : Executive Summary, Exceptions, Price Movements, Cashflows 30D et Data Quality. Les anomalies peuvent être filtrées par portefeuille, type d'actif, sévérité ou type d'anomalie.
6. Un bouton permet de télécharger le reporting Excel, avec les mêmes cinq onglets.


## Lancer l'outil sur son ordinateur

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Sous Windows, la deuxième ligne devient : .venv\Scripts\activate



## Les fichiers

app.py : tout l'outil, dans un seul fichier, en quatre parties qui suivent l'ordre de l'énoncé : les paramètres, les traitements et contrôles, le reporting Excel et l'application Streamlit.

data : les trois fichiers Excel d'entrée. Les cases en erreur sont surlignées en jaune.

anomalies_attendues.xlsx : la liste des anomalies introduites volontairement, avec la ligne concernée et le résultat attendu.

reporting_controles.xlsx : le reporting généré par l'outil sur ces données.

PLAN_CONSTRUCTION.md : la note sur l'utilisation de l'IA, avec mes demandes à l'IA étape par étape.

requirements.txt : les librairies à installer.


## Les données

38 positions réparties dans 4 portefeuilles, en 3 devises (EUR, USD, GBP), sur 3 types d'actifs (CLO Debt, CLO Equity, ABS). Les ratings vont de AAA à BB, et NR pour l'equity de CLO, qui n'est pas notée.

30 prix précédents et 167 cashflows (coupons, principal et frais), avec les statuts Paid, Confirmed et Forecast.

La date de valorisation est le 12/11/2026. Les prix précédents sont ceux du 11/11/2026.

17 anomalies ont été introduites volontairement :

7 Critical : 1 ISIN manquant, 1 prix de marché manquant, 1 position en double, 2 nominaux nuls ou négatifs, 2 prix nuls ou négatifs.

8 Warning : 1 rating manquant, 3 variations de prix trop fortes (une par type d'actif), 1 cashflow sans position, 1 cashflow en double, 2 cashflows passés restés en Forecast.

2 Information : 2 nouveaux titres sans prix de la veille.

Trois lignes pièges ne doivent pas être signalées : une CLO Equity qui monte de 10 % (son seuil est de 15 %), une CLO Debt qui monte de 5 % pile (on signale seulement au-dessus de 5 %) et un même titre détenu par deux portefeuilles (ce n'est pas un doublon).


## Les contrôles

Colonnes obligatoires : si une colonne manque, le traitement ne se lance pas.

Données manquantes : un ISIN ou un prix de marché vide est Critical, un rating vide est Warning.

Valeurs nulles ou négatives : un nominal, un prix comptable ou un prix de marché inférieur ou égal à 0 est Critical.

Doublon de position : même portefeuille et même ISIN qu'une ligne précédente. Critical, car la position compterait deux fois.

Variation de prix : écart entre le prix du jour et le prix de la veille, en valeur et en pourcentage. Au-delà de 5 % pour la CLO Debt et l'ABS, et de 15 % pour la CLO Equity, c'est un Warning.

Nouveau titre : un ISIN sans prix de la veille. Information, car ce n'est pas une erreur, mais sa variation ne peut pas être contrôlée.

Cashflow sans position : un cashflow dont le portefeuille et l'ISIN ne correspondent à aucune position. Warning.

Doublon de cashflow : même portefeuille, ISIN, date, type, devise et montant qu'une ligne précédente. Warning.

Cashflow passé en Forecast : un cashflow dont la date est passée mais qui est toujours prévu, et pas payé. Warning.

Cashflows à venir : tous les cashflows entre la date de valorisation et les 30 jours qui suivent.

La logique des sévérités :
Critical, la donnée est fausse ou inutilisable et les totaux sont faux, il faut corriger ;
Warning, la donnée est suspecte, quelqu'un doit vérifier ;
Information, ce n'est pas une erreur, c'est juste à savoir.


## Modifier les règles

Toutes les règles sont dans la partie PARAMÈTRES, en haut de app.py. On les change sans toucher au code des contrôles :

PRICE_THRESHOLDS : les seuils de variation par type d'actif. Ils sont aussi modifiables dans l'application.
CASHFLOW_HORIZON_DAYS : l'horizon des cashflows à venir, 30 jours par défaut. Aussi modifiable dans l'application.
MISSING_CHECKS : les colonnes qui doivent être remplies. Ajouter MaturityDate à cette liste crée un nouveau contrôle.
POSITIVE_COLUMNS : les colonnes qui doivent être positives.
SEVERITY : la sévérité de chaque contrôle.


## Hypothèses

L'énoncé parle d'un onglet « Cashflows 3D ». Je l'ai compris comme 30D, puisque le contrôle demandé porte sur les 30 prochains jours.

La date de référence est la date de valorisation du fichier de positions, pas la date du jour. Le résultat est donc le même quel que soit le jour où l'on lance l'outil.

Le seuil est strict : une variation de 5 % pile n'est pas signalée, puisque l'énoncé dit « supérieures à 5 % ».

Un même ISIN dans deux portefeuilles est normal. Seule la deuxième ligne d'un vrai doublon est signalée.

Si un des deux prix est vide, nul ou négatif, la variation n'est pas calculée. L'erreur est déjà signalée par le contrôle sur le prix, il n'y a pas de deuxième alerte.

Une valeur illisible, par exemple du texte à la place d'un prix, est traitée comme une donnée manquante.

Le rating NR est accepté, car l'equity de CLO n'est pas notée. Seul un rating vide est signalé.

L'outil signale les anomalies mais ne corrige pas les données. Les totaux sont calculés sur les données telles qu'elles sont.

Les fichiers ne contiennent pas de taux de change. Le nominal total est donc donné sans conversion, avec le détail par devise.


## Limites et ce que j'aurais ajouté avec plus de temps

Une conversion de devises, avec un fichier de taux de change, pour avoir un nominal total dans une seule devise.

Un contrôle des valeurs autorisées, par exemple un type d'actif ou une devise qui n'existe pas.

Des contrôles de cohérence entre fichiers : une maturité déjà passée, une devise de cashflow différente de celle de la position.

Le signalement des doublons dans le fichier des prix précédents. Aujourd'hui, l'outil garde simplement la dernière ligne.

Un historique des contrôles d'un jour à l'autre.

Des tests automatiques livrés avec le projet, un par contrôle.

Un accès protégé par mot de passe à l'application.


## Vérification

Pendant la construction, un script de test a comparé les anomalies trouvées par l'outil à la liste de anomalies_attendues.xlsx. Cette liste a été remplie au moment où chaque erreur a été introduite, indépendamment du code. Le test échouait s'il manquait une anomalie ou s'il y en avait une en trop, ce qui vérifiait aussi que les lignes pièges ne sont pas signalées. Il ne fait pas partie des livrables.

On peut refaire la vérification dans l'application : cocher « Utiliser les fichiers d'exemple », lancer les traitements, puis comparer l'onglet Exceptions à anomalies_attendues.xlsx.

Résultat : 17 anomalies trouvées sur 17 attendues.


## Mise en ligne

L'application est hébergée sur Streamlit Community Cloud, à partir du dépôt GitHub. Sur share.streamlit.io, on se connecte avec son compte GitHub, on choisit le dépôt controles-clo-abs, la branche main et le fichier app.py. Streamlit installe les librairies et lance l'application.

L'application en ligne n'utilise que les données synthétiques du dossier data. Pour la tester sans rien charger, cocher « Utiliser les fichiers d'exemple ».

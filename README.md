# Outil de contrôle positions / prix / cashflows — CLO & ABS

Prototype d'étude de cas. L'outil charge trois fichiers Excel, contrôle les données en Python, puis restitue les résultats dans une application Streamlit et dans un reporting Excel généré automatiquement.

Toutes les données sont **synthétiques** : noms de deals, ISIN, prix et cashflows sont inventés.

Application en ligne : _URL à compléter_

Code source : https://github.com/flutchi22/controles-clo-abs

## Lancer l'outil

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows : .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Dans l'application : charger les trois fichiers du dossier `data/` (ou cocher « Utiliser les fichiers d'exemple »), ajuster les seuils si besoin, puis cliquer sur « Lancer les traitements ». Le reporting Excel se télécharge avec le bouton en haut de la page.

Autres commandes :

| Commande | Effet |
|---|---|
| `python test_controls.py` | vérifie que les anomalies détectées sont exactement les anomalies attendues, puis génère `reporting_controles.xlsx` |

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `data/positions.xlsx`, `data/previous_prices.xlsx`, `data/cashflows.xlsx` | les trois fichiers d'entrée synthétiques. Les anomalies introduites y sont surlignées en jaune pour les repérer |
| `anomalies_attendues.xlsx` | anomalies volontairement introduites et résultat attendu de chaque contrôle |
| `reporting_controles.xlsx` | reporting Excel généré par l'outil sur ces données |
| `app.py` | tout l'outil dans un seul fichier, dans l'ordre de l'énoncé : paramètres, traitements et contrôles, reporting Excel, application Streamlit |
| `test_controls.py` | test de bout en bout |
| `NOTE_IA.md` | note sur l'utilisation de l'IA |

## Données synthétiques

- 38 positions, 4 portefeuilles, 3 devises (EUR, USD, GBP), 3 types d'actifs, 6 ratings (AAA, AA, A, BBB, BB, plus NR pour l'equity de CLO).
- 30 prix précédents, 167 cashflows (Coupon, Principal, Fee ; statuts Paid, Confirmed, Forecast).
- Date de valorisation : 12/11/2026. Prix précédents au 11/11/2026.
- 5 titres sont détenus par deux portefeuilles : c'est une situation normale, qui ne doit pas être prise pour un doublon.

### Anomalies introduites

Le détail ligne par ligne est dans `anomalies_attendues.xlsx`. Résumé :

| Anomalie | Nombre | Sévérité attendue |
|---|---|---|
| ISIN manquant | 1 | Critical |
| Prix de marché manquant | 1 | Critical |
| Doublon de position | 1 | Critical |
| Nominal nul ou négatif | 2 | Critical |
| Prix nul ou négatif (un MarketPrice, un BookPrice) | 2 | Critical |
| Rating manquant | 1 | Warning |
| Variation de prix importante (une par type d'actif) | 3 | Warning |
| Cashflow sans position correspondante | 1 | Warning |
| Doublon de cashflow | 1 | Warning |
| Cashflow passé resté en Forecast | 2 | Warning |
| Nouveau titre sans historique de prix | 2 | Information |
| **Total** | **17** | 7 Critical, 8 Warning, 2 Information |

Trois **cas témoins** ne doivent pas être signalés : une equity à +10 % (sous le seuil de 15 %), une dette à +5 % pile (le seuil est strict) et un ISIN présent dans deux portefeuilles.

## Contrôles

| Contrôle | Règle | Sévérité |
|---|---|---|
| Colonnes obligatoires | toutes les colonnes attendues sont présentes, sinon le traitement est bloqué | Critical |
| Données manquantes | `ISIN`, `MarketPrice`, `Rating` vides ou illisibles | Critical / Critical / Warning |
| Doublon de position | même `Portfolio` + `ISIN` qu'une ligne précédente | Critical |
| Valeur nulle ou négative | `Nominal`, `BookPrice`, `MarketPrice` ≤ 0 | Critical |
| Variation de prix | \|variation en %\| > seuil : 5 % pour CLO Debt et ABS, 15 % pour CLO Equity | Warning |
| Nouveau titre | ISIN absent de `previous_prices.xlsx` | Information |
| Cashflow sans position | couple `Portfolio` + `ISIN` absent des positions | Warning |
| Doublon de cashflow | mêmes portefeuille, ISIN, date, type, devise et montant | Warning |
| Cashflow passé en Forecast | date antérieure à la date de valorisation, statut encore `Forecast` | Warning |

Logique des sévérités : **Critical** = la donnée est inutilisable ou fausse les totaux ; **Warning** = à vérifier par un utilisateur ; **Information** = à savoir, sans action obligatoire.

## Paramétrage

Les règles sont regroupées dans le bloc « PARAMÈTRES » en haut de `app.py`, pas dans le code des contrôles :

- `PRICE_THRESHOLDS` : seuil de variation par type d'actif (aussi modifiable dans la barre latérale de l'application) ;
- `CASHFLOW_HORIZON_DAYS` : horizon des cashflows à venir (aussi modifiable dans l'application) ;
- `MISSING_CHECKS` et `POSITIVE_COLUMNS` : colonnes à contrôler. Ajouter `"MaturityDate"` à `MISSING_CHECKS` crée un nouveau contrôle sans toucher au reste ;
- `SEVERITY` : sévérité de chaque contrôle ;
- `POSITION_KEY`, `CASHFLOW_KEY` : colonnes qui définissent un doublon.

## Hypothèses

- L'onglet « Cashflows 3D » de l'énoncé est lu comme « 30D » (cashflows des 30 prochains jours). L'onglet porte le nom de l'horizon choisi : `Cashflows 30D`.
- La date de référence des contrôles est la `ValuationDate` la plus récente de `positions.xlsx`, et non la date du jour : le résultat est identique quel que soit le jour d'exécution.
- Variation = (`MarketPrice` − `PreviousPrice`) / `PreviousPrice`. Le seuil est strict : une variation d'exactement 5 % n'est pas signalée.
- Un type d'actif inconnu reçoit le seuil par défaut de 5 %.
- Le même ISIN dans deux portefeuilles est normal. Seule la ligne en double est signalée, pas la première occurrence.
- Les mouvements de prix sont calculés par ligne de position : un titre détenu par deux portefeuilles apparaît deux fois.
- Si le prix du jour ou le prix précédent est manquant, nul ou négatif, la variation est « Non calculable ». L'anomalie est déjà remontée par le contrôle de prix, il n'y a pas de deuxième alerte.
- Une valeur illisible (du texte dans un prix, une date invalide) est traitée comme une donnée manquante.
- Le rating `NR` est une valeur valide (equity de CLO, non notée) : seul un rating vide est signalé.
- Les cashflows à venir sont ceux dont la date est comprise entre la date de valorisation et J+30, bornes incluses, tous statuts confondus.
- L'outil signale les anomalies, il ne corrige pas les données. Les totaux (nominal, cashflows) sont calculés sur les données brutes, lignes en anomalie comprises.
- Sans taux de change dans les fichiers, le nominal total est donné sans conversion, avec le détail par devise.

## Limites et améliorations possibles

- Pas de conversion de devises : il faudrait un fichier de taux de change pour un nominal total homogène.
- Pas de contrôle des valeurs autorisées (`AssetType`, `Currency`, `CashflowType`, `Status`).
- Pas de contrôle de cohérence entre fichiers : maturité déjà dépassée, devise du cashflow différente de celle de la position, prix différents pour un même ISIN selon le portefeuille.
- Les doublons de `previous_prices.xlsx` ne sont pas signalés : la dernière ligne de chaque ISIN est retenue.
- Un seul jeu de fichiers à la fois, sans historique des contrôles d'un jour à l'autre.
- Un seul test de bout en bout. Avec plus de temps : un test par contrôle et des tests sur des fichiers mal formés.
- Pas d'authentification sur l'application.

## Vérification

`python test_controls.py` compare les anomalies détectées à celles de `anomalies_attendues.xlsx` (fichier, ligne Excel, type d'anomalie, sévérité). Ce tableau a été rempli au moment où chaque anomalie a été introduite dans les données : il ne dépend pas du code des contrôles. Le test échoue s'il manque une anomalie ou s'il y a une fausse alerte, cas témoins compris.

## Déploiement

L'application est hébergée sur Streamlit Community Cloud, à partir de ce dépôt GitHub :

1. sur share.streamlit.io, se connecter avec le compte GitHub ;
2. « Create app », choisir le dépôt `controles-clo-abs`, la branche `main` et le fichier `app.py` ;
3. Streamlit installe les librairies de `requirements.txt` et lance l'application.

L'application en ligne n'utilise que les données synthétiques du dossier `data/`. Dans l'application, cocher « Utiliser les fichiers d'exemple » pour la tester sans rien charger.

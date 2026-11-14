# Utilisation de l'IA et plan de construction

## L'outil utilisé

J'ai utilisé Claude Code, l'assistant de programmation d'Anthropic, directement dans VS Code. Il lit les fichiers du projet, écrit le code, le lance et lit les résultats.

## Comment j'ai travaillé

J'ai défini ce que l'outil devait faire, à partir de l'énoncé. L'IA a écrit le premier jet : les fichiers Excel, app.py, le test et le README. J'ai ensuite relu le code, vérifié les seuils et les sévérités, contrôlé les données, testé l'outil et fait modifier plusieurs points.

Ci-dessous, mes demandes à l'IA, remises dans l'ordre et reformulées étape par étape. Pour chacune, j'explique pourquoi je l'ai faite comme ça.

Pour toutes les étapes : code simple, pas de classes, commentaires en français, uniquement pandas, openpyxl et streamlit.

---

## Prompt 1 — L'environnement

Crée un environnement virtuel `.venv` et installe streamlit, pandas et openpyxl. Écris un `requirements.txt` avec les versions exactes : `streamlit==1.67.0`, `pandas==3.1.1`, `openpyxl==3.1.5`.

Pourquoi : je fige les versions pour que l'outil tourne avec les mêmes versions sur un autre poste.

---

## Prompt 2 — Les fichiers Excel d'entrée

Crée directement trois fichiers Excel dans `data/` : `positions.xlsx`, `previous_prices.xlsx`, `cashflows.xlsx`, avec les colonnes de l'énoncé. Je ne veux pas de script de génération dans le projet, seulement les fichiers.

D'abord des données propres

Fichier 1 : positions.xlsx
Le fichier devra contenir au minimum les colonnes suivantes :
ValuationDate : date de valorisation
Portfolio : identifiant du portefeuille
ISIN : identifiant du titre
SecurityName : nom du titre
AssetType : CLO Debt, CLO Equity ou ABS
Currency : devise
Nominal : nominal détenu
BookPrice : prix comptable
MarketPrice : prix de marché
Rating : notation du titre
MaturityDate : date de maturité

Fichier 2 : previous_prices.xlsx
Le fichier devra contenir au minimum :
ISIN
PreviousDate
PreviousPrice
PriceSource

Fichier 3 : cashflows.xlsx
Le fichier devra contenir au minimum :
Portfolio
ISIN
CashflowDate
CashflowType : Coupon, Principal ou Fee
Currency
Amount
Status : Forecast, Confirmed ou Paid

Ensuite les anomalies
ISIN manquant ;
doublon de position ;
prix de marché manquant ;
rating manquant ;
nouveau titre sans historique de prix ;
variation importante de prix ;
nominal nul ou négatif ;
prix nul ou négatif ;
cashflow sans position correspondante ;
doublon de cashflow ;
cashflow dont la date est passée mais dont le statut reste Forecast.


Ajoute aussi trois cas pièges, qui ne doivent pas être signalés :
une CLO Equity à +10 % (sous le seuil de 15 %) ;
une CLO Debt à +5 % pile ;
un même ISIN dans deux portefeuilles.

Surligne en jaune fluo les cases en erreur.

Crée un fichier `anomalies_attendues.xlsx` : une ligne par anomalie, avec le fichier, la ligne, le type d'anomalie et la sévérité attendue.

Pourquoi : je sais exactement ce que l'outil doit trouver. Les cas pièges vérifient qu'il ne signale pas tout et n'importe quoi.

---

## Prompt 3 — `app.py` : les paramètres

Crée `app.py`, en quatre parties : Paramètres, Traitements et contrôles, Reporting Excel, Application Streamlit. Commence par les paramètres, en haut du fichier :

colonnes obligatoires de chaque fichier ;
seuils de variation : 5 % pour CLO Debt et ABS, 15 % pour CLO Equity ;
horizon des cashflows : 30 jours ;
colonnes qui doivent être remplies : ISIN, MarketPrice, Rating ;
colonnes qui doivent être positives : Nominal, BookPrice, MarketPrice ;
sévérité de chaque contrôle : Critical, Warning ou Information.

Pourquoi : si on veut changer un seuil en direct. Il suffit de modifier une ligne en haut du fichier.

Sévérités :

Critical : la donnée est fausse ou inutilisable, donc les calculs qui l'utilisent sont faux. Un nominal négatif ou un prix à 0 n'existe pas pour une position détenue : c'est forcément une erreur de saisie. Sans ISIN, on ne sait pas quel titre on détient. Sans prix de marché, on ne peut pas valoriser la position. Une position en double compte deux fois dans le nominal total. Il faut corriger avant d'utiliser le reporting.

Warning : la donnée est peut-être juste, mais elle est suspecte. Une CLO Debt qui bouge de plus de 5 % en un jour, c'est rare : il peut y avoir un vrai événement de marché, ou une erreur de prix. Un cashflow passé encore en Forecast : soit il a été payé et le statut n'a pas été mis à jour, soit le paiement n'est pas arrivé. Un rating manquant ou un cashflow sans position sont dans le même cas. Quelqu'un doit aller vérifier, mais le reste du reporting reste utilisable.

Information : ce n'est pas une erreur, juste un fait utile à connaître. Un nouveau titre n'a pas de prix de la veille, c'est normal puisqu'il vient d'être acheté. On ne peut simplement pas calculer sa variation : l'information prévient que cette ligne n'a pas été contrôlée. Aucune action n'est nécessaire.


## Prompt 4 — `app.py` : les traitements et contrôles

Écris les contrôles, avec en commentaire au-dessus de chaque bloc la phrase de l'énoncé qu'il traite :

charger et valider les inputs ;
vérifier la présence des colonnes obligatoires ;
détecter les anomalies de qualité de données ;
détecter les doublons ;
identifier les données manquantes ;
comparer les prix du jour aux prix précédents ;
calculer les variations absolues et en pourcentage ;
signaler les variations supérieures au seuil ;
identifier les nouveaux titres sans historique de prix ;
identifier les cashflows des 30 prochains jours ;
classifier les anomalies en Critical, Warning ou Information.

Pourquoi : on vérifie point par point que toute la consigne est couverte.

Mes règles :
la date de référence est la date de valorisation du fichier, pas la date du jour ;
le seuil est strict : 5 % pile n'est pas signalé ;
un doublon, c'est même portefeuille + même ISIN (deux fonds peuvent détenir le même titre) ;
une valeur illisible (du texte dans un prix) est traitée comme manquante.

---

## Prompt 5 — `app.py` : le reporting Excel

Génère un fichier Excel avec les onglets Executive Summary, Exceptions, Price Movements, Cashflows 30D et Data Quality, avec les colonnes de l'énoncé.

Rends-le directement exploitable : filtres, en-têtes figés, colonnes lisibles, dates en JJ/MM/AAAA, couleur selon la sévérité (rouge Critical, orange Warning, bleu Information).

Pourquoi : on repère les anomalies importantes sans rien configurer. L'énoncé écrit « Cashflows 3D », je le lis comme 30D.

---

## Prompt 6 — `app.py` : l'application Streamlit

Écris l'interface, avec en commentaire chaque point de l'énoncé :

charger les trois fichiers Excel d'entrée (ou cocher « Utiliser les fichiers d'exemple ») ;
contrôler que les fichiers et colonnes attendus sont présents ;
modifier les seuils ;
lancer les traitements ;
afficher un Executive Summary, le nombre de positions et d'anomalies par sévérité ;
visualiser les anomalies, avec des filtres par portefeuille, type d'actif, sévérité et type d'anomalie ;
visualiser les principaux mouvements de prix ;
visualiser les cashflows à venir ;
télécharger le reporting Excel.

Pourquoi : les fichiers d'exemple permettent de tester l'application sans avoir à charger les fichiers. Les onglets sont les mêmes que dans l'Excel.

---

## Prompt 7 — Le test

Écris `test_controls.py` : il lance les contrôles sur les fichiers de `data/` et vérifie que les anomalies trouvées sont exactement celles de `anomalies_attendues.xlsx`, ni plus ni moins. Si tout est bon, il génère `reporting_controles.xlsx`.

Ce test m'a servi pendant la construction. Je ne l'ai pas gardé dans les livrables, qui se limitent à ce que demande l'énoncé.

Pourquoi : trouver 17 anomalies ne prouve rien si ce ne sont pas les bonnes. Le test vérifie aussi que les cas pièges ne sont pas signalés.

---

## Prompt 8 — Vérifier et documenter

Lance le test, ouvre les fichiers Excel, parcours l'application et essaie des fichiers abîmés (colonne supprimée, texte dans un prix). Puis écris le README et ce document.

Résultats attendus :
17 anomalies trouvées sur 17 attendues (7 Critical, 8 Warning, 2 Information) ;
seuil CLO Equity abaissé à 9 % : 18 anomalies, le cas piège à +10 % est alors signalé.

Pourquoi : un code qui tourne sans erreur n'est pas forcément juste. L'erreur sur le format des dates ne se voyait qu'en ouvrant le fichier.


---

## Prompt 9 — Mettre en ligne

Mets le projet sur mon GitHub, sans le dossier .venv. Ensuite je déploie l'application sur share.streamlit.io avec app.py, et je mets l'URL dans le README.

Pourquoi : l'énoncé recommande une application accessible depuis un navigateur, sans rien installer.

---

## Les cas pièges

Ce sont des lignes qui ressemblent à des erreurs mais qui sont normales. L'outil ne doit pas les signaler.

J'en ai mis trois, dans positions.xlsx :
ligne 6 : une CLO Equity qui monte de 10 %. Normal, son seuil est de 15 %.
ligne 14 : une CLO Debt qui monte de 5 % pile. Normal, on signale seulement au-dessus de 5 %.
ligne 34 : un titre détenu par deux portefeuilles. Normal, ce n'est pas un doublon.

Pourquoi ces lignes : elles sont pile à la limite des règles. Si le code a une erreur, c'est là qu'elle se voit.

Le test vérifie que ces trois lignes ne sont pas signalées.

---

## Ce que j'ai corrigé ou challengé

Mes choix :
l'IA avait découpé l'outil en quatre fichiers qui s'appellent entre eux. J'ai préféré un seul fichier, app.py, qui suit l'ordre de l'énoncé, avec chaque phrase de la consigne en commentaire. C'est plus facile à suivre et à modifier ;
les paramètres sont regroupés en haut du fichier : un seuil se change en modifiant une seule ligne ;
l'IA avait fait un script pour fabriquer les données. Je l'ai supprimé : ce qui compte, ce sont les fichiers Excel et la liste des anomalies attendues.

Les erreurs de l'IA, repérées en ouvrant les fichiers produits :
les dates s'affichaient 2026-11-12 00:00:00 au lieu de 12/11/2026 dans Excel. Le code tournait sans erreur, seul le fichier montrait le problème ;
les quatre anomalies de cashflows tombaient toutes dans le même portefeuille, donc le filtre par portefeuille ne pouvait pas être testé. Elles ont été réparties ;
dans l'onglet Data Quality, un nombre s'affichait 38.0 au lieu de 38.

---

## Comment j'ai vérifié les résultats

Le test, lancé pendant la construction : l'outil retrouve exactement les 17 anomalies de anomalies_attendues.xlsx, sans en oublier ni en ajouter.
Les cas pièges : aucun des trois n'est signalé.
Un calcul refait à la main : un prix qui passe de 98,00 à 104,37 donne bien +6,50 %.
Les fichiers Excel relus : onglets, filtres, formats, couleurs. C'est comme ça que l'erreur sur les dates a été vue.
Les cases jaunes des fichiers d'entrée se retrouvent toutes dans l'onglet Exceptions.
L'application : en baissant le seuil de la CLO Equity à 9 %, on passe de 17 à 18 anomalies. Le paramétrage fonctionne.
Des fichiers abîmés exprès (colonne supprimée, texte dans un prix, fichier vide) : l'outil affiche un message clair au lieu de planter.

"""Outil de contrôle positions / prix / cashflows — CLO & ABS.

Tout l'outil est dans ce fichier, dans l'ordre de l'énoncé :
    PARAMÈTRES                          règles et seuils, modifiables sans toucher au reste
    3.   TRAITEMENTS ET CONTRÔLES       chargement, validation, contrôles
    4.B  REPORTING EXCEL                fichier de restitution
    4.A  APPLICATION STREAMLIT          interface

Lancer : streamlit run app.py
"""
import io
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill

# ======================================================================================
# PARAMÈTRES
# « Les règles et seuils devront, autant que possible, être paramétrables »
# ======================================================================================

# Colonnes obligatoires de chaque fichier d'entrée
REQUIRED_COLUMNS = {
    "positions": ["ValuationDate", "Portfolio", "ISIN", "SecurityName", "AssetType", "Currency",
                  "Nominal", "BookPrice", "MarketPrice", "Rating", "MaturityDate"],
    "previous_prices": ["ISIN", "PreviousDate", "PreviousPrice", "PriceSource"],
    "cashflows": ["Portfolio", "ISIN", "CashflowDate", "CashflowType", "Currency", "Amount", "Status"],
}

# Type des colonnes (les autres colonnes sont du texte)
DATE_COLUMNS = ["ValuationDate", "MaturityDate", "PreviousDate", "CashflowDate"]
NUMERIC_COLUMNS = ["Nominal", "BookPrice", "MarketPrice", "PreviousPrice", "Amount"]

# Seuil de variation de prix (en %) par type d'actif : au-delà, la variation est signalée
PRICE_THRESHOLDS = {"CLO Debt": 5.0, "ABS": 5.0, "CLO Equity": 15.0}
DEFAULT_PRICE_THRESHOLD = 5.0  # utilisé si le type d'actif n'est pas dans la liste

# Nombre de jours pour les « cashflows à venir »
CASHFLOW_HORIZON_DAYS = 30

# Colonnes de positions.xlsx qui doivent être remplies
MISSING_CHECKS = ["ISIN", "MarketPrice", "Rating"]

# Colonnes de positions.xlsx qui doivent être strictement positives
POSITIVE_COLUMNS = ["Nominal", "BookPrice", "MarketPrice"]

# Colonnes qui définissent un doublon
POSITION_KEY = ["Portfolio", "ISIN"]
CASHFLOW_KEY = ["Portfolio", "ISIN", "CashflowDate", "CashflowType", "Currency", "Amount"]

# Sévérité de chaque contrôle
SEVERITY_ORDER = ["Critical", "Warning", "Information"]
SEVERITY = {
    "ISIN manquant": "Critical",
    "MarketPrice manquant": "Critical",
    "Rating manquant": "Warning",
    "Doublon de position": "Critical",
    "Nominal nul ou négatif": "Critical",
    "BookPrice nul ou négatif": "Critical",
    "MarketPrice nul ou négatif": "Critical",
    "Variation de prix importante": "Warning",
    "Nouveau titre sans historique": "Information",
    "Cashflow sans position": "Warning",
    "Doublon de cashflow": "Warning",
    "Cashflow passé en Forecast": "Warning",
}
DEFAULT_SEVERITY = "Warning"  # pour un contrôle qui ne serait pas dans la liste ci-dessus

# Colonnes du tableau d'exceptions
EXCEPTION_COLUMNS = ["Source", "Row", "Portfolio", "ISIN", "SecurityName", "AssetType",
                     "AnomalyType", "ObservedValue", "Severity", "Comment"]

# Mise en forme du reporting Excel : couleur selon la valeur de la cellule, format selon la colonne
RED, ORANGE, BLUE, GREY = "F4B6B6", "FFE699", "BDD7EE", "D9D9D9"
COLORS = {"Critical": RED, "Warning": ORANGE, "Information": BLUE,
          "Dépassement": RED, "KO": RED, "Nouveau titre": BLUE, "Non calculable": GREY}
NUMBER_FORMATS = {"Value": "#,##0", "Amount": "#,##0.00", "MarketPrice": "0.00", "PreviousPrice": "0.00",
                  "AbsChange": "+0.00;-0.00", "PctChange": '+0.00"%";-0.00"%"', "Threshold": '0.0"%"'}


# ======================================================================================
# 3. TRAITEMENTS ET CONTRÔLES PYTHON
# ======================================================================================

# charger et valider les inputs ;
# vérifier la présence des colonnes obligatoires ;
def load_file(file, name):
    """Lit un fichier Excel. Renvoie (tableau, liste des colonnes obligatoires manquantes)."""
    if hasattr(file, "seek"):
        file.seek(0)  # un fichier chargé dans Streamlit peut avoir déjà été lu
    df = pd.read_excel(file)
    df.columns = df.columns.astype(str).str.strip()
    missing = [c for c in REQUIRED_COLUMNS[name] if c not in df.columns]
    return df, missing


def clean(df, name):
    """Garde les colonnes obligatoires et les met au bon type.

    Une valeur illisible (texte dans un prix, date invalide...) devient vide
    et sera donc signalée comme donnée manquante.
    """
    out = df[REQUIRED_COLUMNS[name]].copy()
    for col in out.columns:
        if col in DATE_COLUMNS:
            out[col] = pd.to_datetime(out[col], errors="coerce")
        elif col in NUMERIC_COLUMNS:
            out[col] = pd.to_numeric(out[col], errors="coerce")
        else:
            text = out[col].astype("string").str.strip()
            out[col] = text.replace("", pd.NA)  # cellule vide ou espaces = manquant
    out["Row"] = out.index + 2  # numéro de ligne dans Excel (la ligne 1 contient les en-têtes)
    return out


def run_controls(positions, previous_prices, cashflows, price_thresholds=None, horizon_days=None):
    """Lance tous les contrôles. Renvoie un dictionnaire de tableaux (un par onglet du reporting)."""
    thresholds = price_thresholds or PRICE_THRESHOLDS
    horizon = horizon_days or CASHFLOW_HORIZON_DAYS

    pos = clean(positions, "positions")
    prev = clean(previous_prices, "previous_prices")
    cf = clean(cashflows, "cashflows")

    # Date de référence de tous les contrôles = date de valorisation du fichier de positions
    val_date = pos["ValuationDate"].max()
    if pd.isna(val_date):
        raise ValueError("Aucune ValuationDate valide dans positions.xlsx : impossible de lancer les contrôles.")

    # On ajoute le nom et le type du titre aux cashflows (pour l'affichage et les filtres)
    names = pos.dropna(subset=["ISIN"]).drop_duplicates("ISIN")[["ISIN", "SecurityName", "AssetType"]]
    cf = cf.merge(names, on="ISIN", how="left")

    sizes = {"positions.xlsx": len(pos), "cashflows.xlsx": len(cf)}
    exceptions = []  # un tableau d'anomalies par contrôle
    controls = [{"Control": "Colonnes obligatoires présentes", "File": "les 3 fichiers",
                 "Severity": "Critical", "RowsChecked": len(pos) + len(prev) + len(cf), "Anomalies": 0}]

    # classifier les anomalies, par exemple en Critical, Warning ou Information.
    def add(rows, source, anomaly, observed, comment):
        """Enregistre un contrôle et les lignes en anomalie qu'il a trouvées (rows peut être vide)."""
        severity = SEVERITY.get(anomaly, DEFAULT_SEVERITY)
        controls.append({"Control": anomaly, "File": source, "Severity": severity,
                         "RowsChecked": sizes[source], "Anomalies": len(rows)})
        if len(rows) > 0:
            out = rows.reindex(columns=EXCEPTION_COLUMNS)
            out["Source"], out["AnomalyType"], out["Severity"] = source, anomaly, severity
            out["ObservedValue"], out["Comment"] = observed, comment
            exceptions.append(out)

    # ---------------------------------------------------------------------------------
    # détecter les anomalies de qualité de données ;
    # ---------------------------------------------------------------------------------
    # Nominal ou prix nul ou négatif
    for col in POSITIVE_COLUMNS:
        bad = pos[pos[col] <= 0]
        add(bad, "positions.xlsx", f"{col} nul ou négatif",
            bad[col].map("{:,.2f}".format), f"{col} doit être strictement positif.")

    # Cashflow sans position correspondante (couple Portfolio + ISIN inconnu)
    pos_keys = (pos["Portfolio"] + "|" + pos["ISIN"]).dropna()
    orphan = cf[~(cf["Portfolio"] + "|" + cf["ISIN"]).isin(pos_keys)]
    add(orphan, "cashflows.xlsx", "Cashflow sans position",
        orphan["Portfolio"] + " / " + orphan["ISIN"],
        "Aucune position pour ce couple Portfolio + ISIN.")

    # Cashflow dont la date est passée mais dont le statut reste Forecast
    late = cf[(cf["CashflowDate"] < val_date) & (cf["Status"].str.lower() == "forecast")]
    add(late, "cashflows.xlsx", "Cashflow passé en Forecast",
        late["CashflowDate"].dt.strftime("%d/%m/%Y"),
        f"Date antérieure à la date de valorisation ({val_date:%d/%m/%Y}) mais statut toujours Forecast.")

    # ---------------------------------------------------------------------------------
    # détecter les doublons ;
    # ---------------------------------------------------------------------------------
    # Position : même Portfolio + ISIN (le même ISIN dans deux portefeuilles est normal)
    dup = pos[pos["ISIN"].notna() & pos.duplicated(POSITION_KEY)]
    add(dup, "positions.xlsx", "Doublon de position",
        dup["Portfolio"] + " / " + dup["ISIN"],
        "Même Portfolio + ISIN qu'une ligne précédente : risque de double comptage.")

    # Cashflow : même portefeuille, ISIN, date, type, devise et montant
    dup_cf = cf[cf.duplicated(CASHFLOW_KEY)]
    add(dup_cf, "cashflows.xlsx", "Doublon de cashflow",
        dup_cf["Amount"].map("{:,.2f}".format),
        "Identique à une ligne précédente (portefeuille, ISIN, date, type, devise, montant).")

    # ---------------------------------------------------------------------------------
    # identifier les données manquantes ;
    # ---------------------------------------------------------------------------------
    for col in MISSING_CHECKS:
        add(pos[pos[col].isna()], "positions.xlsx", f"{col} manquant",
            "(vide)", f"La colonne {col} n'est pas renseignée (ou sa valeur est illisible).")

    # ---------------------------------------------------------------------------------
    # comparer les prix du jour aux prix précédents ;
    # ---------------------------------------------------------------------------------
    prev = prev.dropna(subset=["ISIN"]).drop_duplicates("ISIN", keep="last")  # un seul prix précédent par ISIN
    pm = pos.dropna(subset=["ISIN"]).merge(prev.drop(columns="Row"), on="ISIN", how="left")

    # calculer les variations absolues et en pourcentage ;
    usable = (pm["MarketPrice"] > 0) & (pm["PreviousPrice"] > 0)  # les deux prix sont exploitables
    pm["AbsChange"] = (pm["MarketPrice"] - pm["PreviousPrice"]).where(usable).round(4)
    pm["PctChange"] = (pm["AbsChange"] / pm["PreviousPrice"] * 100).round(4)

    # signaler les variations supérieures à 5 % pour CLO Debt et ABS et 15 % pour CLO Equity ;
    pm["Threshold"] = pm["AssetType"].map(thresholds).fillna(DEFAULT_PRICE_THRESHOLD).astype(float)
    breach = pm["PctChange"].abs() > pm["Threshold"]  # strictement supérieur au seuil
    moves = pm[breach]
    add(moves, "positions.xlsx", "Variation de prix importante",
        moves["PctChange"].map("{:+.2f} %".format),
        moves.apply(lambda r: f"Prix {r.PreviousPrice:.2f} -> {r.MarketPrice:.2f}, "
                              f"seuil {r.Threshold:g} % pour {r.AssetType}.", axis=1))

    # identifier les nouveaux titres sans historique de prix ;
    is_new = ~pm["ISIN"].isin(prev["ISIN"])  # ISIN absent du fichier de prix précédents
    add(pm[is_new], "positions.xlsx", "Nouveau titre sans historique",
        "Pas de prix précédent", "ISIN absent de previous_prices.xlsx : variation non calculable.")

    # Statut du contrôle de prix, ligne par ligne
    pm["Status"] = "OK"
    pm.loc[~usable, "Status"] = "Non calculable"  # prix manquant, nul ou négatif (déjà signalé plus haut)
    pm.loc[is_new, "Status"] = "Nouveau titre"
    pm.loc[breach, "Status"] = "Dépassement"

    # ---------------------------------------------------------------------------------
    # identifier les cashflows attendus au cours des 30 prochains jours ;
    # ---------------------------------------------------------------------------------
    # entre la date de valorisation et J + horizon, bornes incluses
    cf_next = cf[cf["CashflowDate"].between(val_date, val_date + pd.Timedelta(days=horizon))]
    cf_next = cf_next.sort_values(["CashflowDate", "Portfolio", "ISIN"])

    # ---------------------------------------------------------------------------------
    # Tableaux de restitution (un par onglet)
    # ---------------------------------------------------------------------------------
    # Exceptions, triées par sévérité (Critical en premier)
    exc = pd.concat(exceptions) if exceptions else pd.DataFrame(columns=EXCEPTION_COLUMNS)
    rank = {severity: i for i, severity in enumerate(SEVERITY_ORDER)}
    exc = exc.sort_values(["Source", "Row"]).sort_values("Severity", key=lambda s: s.map(rank), kind="stable")
    exc = exc.reset_index(drop=True)

    # Data Quality : une ligne par contrôle effectué
    dq = pd.DataFrame(controls)
    dq["Status"] = ["KO" if n > 0 else "OK" for n in dq["Anomalies"]]

    # Executive Summary
    by_severity = exc["Severity"].value_counts()
    summary = {
        "Date de valorisation": val_date.strftime("%d/%m/%Y"),
        "Nombre de positions analysées": len(pos),
        "Nombre de portefeuilles": pos["Portfolio"].nunique(),
        "Nominal total (toutes devises, sans conversion FX)": pos["Nominal"].sum(),
    }
    for currency, amount in pos.groupby("Currency")["Nominal"].sum().items():
        summary[f"Nominal total {currency}"] = amount
    summary["Nombre total d'anomalies"] = len(exc)
    for severity in SEVERITY_ORDER:
        summary[f"Anomalies {severity}"] = int(by_severity.get(severity, 0))
    summary["Variations de prix hors seuil"] = int(breach.sum())
    summary["Nouveaux titres sans historique"] = int(is_new.sum())
    summary[f"Cashflows à venir ({horizon} jours)"] = len(cf_next)
    for currency, amount in cf_next.groupby("Currency")["Amount"].sum().items():
        summary[f"Montant net des cashflows à venir {currency}"] = amount

    pm = pm.sort_values("PctChange", key=abs, ascending=False)  # plus gros mouvements en premier
    return {
        "valuation_date": val_date,
        "horizon": horizon,
        "summary": pd.DataFrame({"Indicator": list(summary), "Value": list(summary.values())}),
        "exceptions": exc,
        "price_movements": pm[["Portfolio", "ISIN", "SecurityName", "AssetType", "MarketPrice", "PreviousPrice",
                               "PreviousDate", "AbsChange", "PctChange", "Threshold", "Status"]],
        "cashflows": cf_next[["Portfolio", "ISIN", "SecurityName", "CashflowDate", "CashflowType",
                              "Currency", "Amount", "Status"]],
        "data_quality": dq,
    }


# ======================================================================================
# 4.B REPORTING EXCEL
# ======================================================================================

def build_report(results):
    """Construit le fichier Excel en mémoire et renvoie son contenu (bytes)."""
    sheets = {
        "Executive Summary": results["summary"],                  # positions, nominal, anomalies par sévérité
        "Exceptions": results["exceptions"],                      # une ligne par anomalie, avec commentaire
        "Price Movements": results["price_movements"],            # prix, variations, seuil, statut
        f"Cashflows {results['horizon']}D": results["cashflows"],  # « Cashflows 3D » de l'énoncé, lu comme 30D
        "Data Quality": results["data_quality"],                  # contrôles effectués et nombre d'anomalies
    }
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
            format_sheet(writer.sheets[name])
    return buffer.getvalue()


def format_sheet(ws):
    """Fichier directement exploitable : filtres, formats cohérents, colonnes lisibles, exceptions en couleur."""
    ws.freeze_panes = "A2"              # ligne d'en-têtes figée
    ws.auto_filter.ref = ws.dimensions  # filtres
    for column in ws.iter_cols():
        header = column[0]
        header.font = Font(bold=True, color="FFFFFF")
        header.fill = PatternFill("solid", fgColor="1F4E78")
        width = max(len(str(cell.value)) for cell in column if cell.value is not None)
        ws.column_dimensions[header.column_letter].width = min(width + 4, 80)
        for cell in column[1:]:
            if cell.is_date:
                cell.number_format = "DD/MM/YYYY"
            if header.value in NUMBER_FORMATS and isinstance(cell.value, (int, float)):
                cell.number_format = NUMBER_FORMATS[header.value]
            if cell.value in COLORS:  # identification claire des exceptions importantes
                cell.fill = PatternFill("solid", fgColor=COLORS[cell.value])


# ======================================================================================
# 4.A APPLICATION STREAMLIT
# ======================================================================================

def main():
    date_format = {col: st.column_config.DateColumn(format="DD/MM/YYYY") for col in DATE_COLUMNS}

    def show(df):
        """Affiche un tableau avec les dates au format français."""
        st.dataframe(df, hide_index=True, column_config=date_format)

    st.set_page_config(page_title="Contrôles CLO / ABS", layout="wide")
    st.title("Contrôles positions, prix et cashflows — CLO / ABS")
    st.caption("Prototype d'étude de cas. Données 100 % synthétiques.")

    with st.sidebar:
        # charger les trois fichiers Excel d'entrée ;
        st.header("1. Fichiers d'entrée")
        use_sample = st.checkbox("Utiliser les fichiers d'exemple")
        files = {}
        for name in REQUIRED_COLUMNS:
            if use_sample:
                files[name] = Path(__file__).parent / "data" / f"{name}.xlsx"
            else:
                files[name] = st.file_uploader(f"{name}.xlsx", type="xlsx")

        # règles et seuils paramétrables depuis l'interface
        st.header("2. Paramètres")
        thresholds = {
            asset: st.number_input(f"Seuil de variation {asset} (%)", min_value=0.0, value=value, step=0.5)
            for asset, value in PRICE_THRESHOLDS.items()
        }
        horizon = st.number_input("Horizon des cashflows (jours)", min_value=1, value=CASHFLOW_HORIZON_DAYS)

        st.header("3. Traitement")
        launch = st.button("Lancer les traitements", type="primary")

    # contrôler que les fichiers et colonnes attendus sont présents ;
    st.subheader("Contrôle des fichiers")
    data = {}
    for box, (name, file) in zip(st.columns(3), files.items()):
        with box:
            if file is None:
                st.info(f"{name}.xlsx : en attente du fichier")
                continue
            try:
                df, missing = load_file(file, name)
            except Exception as error:  # fichier corrompu ou qui n'est pas un vrai Excel
                st.error(f"{name}.xlsx : fichier illisible ({error})")
                continue
            if missing:
                st.error(f"{name}.xlsx : colonnes obligatoires manquantes : {', '.join(missing)}")
            else:
                st.success(f"{name}.xlsx : {len(df)} lignes, colonnes obligatoires présentes")
                data[name] = df

    # lancer les traitements ;
    if launch:
        st.session_state.pop("results", None)  # on efface les résultats précédents
        if len(data) < 3:
            st.error("Les trois fichiers doivent être chargés et valides avant de lancer les traitements.")
        else:
            try:
                results = run_controls(**data, price_thresholds=thresholds, horizon_days=int(horizon))
                st.session_state["results"] = results
                st.session_state["report"] = build_report(results)
            except ValueError as error:
                st.error(str(error))

    if "results" not in st.session_state:
        st.info("Chargez les trois fichiers (ou cochez les fichiers d'exemple) "
                "puis cliquez sur « Lancer les traitements ».")
        st.stop()

    res = st.session_state["results"]
    exc, pm, cf = res["exceptions"], res["price_movements"], res["cashflows"]
    summary = dict(zip(res["summary"]["Indicator"], res["summary"]["Value"]))

    # télécharger le reporting Excel généré par l'application.
    st.download_button(
        "Télécharger le reporting Excel",
        st.session_state["report"],
        file_name=f"reporting_controles_{res['valuation_date']:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab_summary, tab_exc, tab_pm, tab_cf, tab_dq = st.tabs(
        ["Executive Summary", "Exceptions", "Price Movements", f"Cashflows {res['horizon']}D", "Data Quality"])

    # afficher un Executive Summary ;
    with tab_summary:
        boxes = st.columns(5)
        # afficher le nombre total de positions analysées ;
        boxes[0].metric("Positions analysées", summary["Nombre de positions analysées"])
        boxes[1].metric("Anomalies", summary["Nombre total d'anomalies"])
        # afficher le nombre d'exceptions par niveau de sévérité ;
        for box, severity in zip(boxes[2:], SEVERITY_ORDER):
            box.metric(severity, summary[f"Anomalies {severity}"])

        left, right = st.columns(2)
        with left:
            st.markdown("**Indicateurs**")
            display = res["summary"].copy()  # nombres avec séparateur de milliers
            display["Value"] = [v if isinstance(v, str) else f"{v:,.0f}".replace(",", " ") for v in display["Value"]]
            show(display)
        with right:
            st.markdown("**Anomalies par type**")
            st.bar_chart(exc["AnomalyType"].value_counts(), horizontal=True)

    # visualiser les principales anomalies ;
    with tab_exc:
        # filtrer les anomalies par portefeuille, type d'actif, sévérité ou autre critère pertinent ;
        view = exc
        for box, col in zip(st.columns(4), ["Portfolio", "AssetType", "Severity", "AnomalyType"]):
            chosen = box.multiselect(col, sorted(exc[col].dropna().unique()))
            if chosen:  # aucun choix = pas de filtre
                view = view[view[col].isin(chosen)]
        st.write(f"{len(view)} anomalie(s) affichée(s) sur {len(exc)}")
        colored = view.style.map(lambda v: f"background-color: #{COLORS[v]}; color: black", subset=["Severity"])
        st.dataframe(colored.format(na_rep=""), hide_index=True)

    # visualiser les principaux mouvements de prix ;
    with tab_pm:
        st.markdown("**10 plus fortes variations (en %)**")
        top = pm.dropna(subset=["PctChange"]).drop_duplicates("ISIN").head(10)
        st.bar_chart(top, x="SecurityName", y="PctChange", horizontal=True)
        only_breaches = st.checkbox("Afficher uniquement les dépassements de seuil")
        show(pm[pm["Status"] == "Dépassement"] if only_breaches else pm)

    # visualiser les cashflows à venir ;
    with tab_cf:
        end = res["valuation_date"] + pd.Timedelta(days=res["horizon"])
        st.markdown(f"**Cashflows attendus du {res['valuation_date']:%d/%m/%Y} au {end:%d/%m/%Y}**")
        if cf.empty:
            st.info("Aucun cashflow sur la période.")
        else:
            per_day = cf.assign(Date=cf["CashflowDate"].dt.strftime("%Y-%m-%d"))
            per_day = per_day.groupby(["Date", "Currency"], as_index=False)["Amount"].sum()
            st.bar_chart(per_day, x="Date", y="Amount", color="Currency")
            show(cf)

    # Data Quality : contrôles effectués et nombre d'anomalies par contrôle
    with tab_dq:
        show(res["data_quality"])


if __name__ == "__main__":  # Streamlit exécute ce fichier comme script principal
    main()

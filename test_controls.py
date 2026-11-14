"""Vérifie que les contrôles trouvent exactement les anomalies listées dans anomalies_attendues.xlsx,
puis génère reporting_controles.xlsx.

Lancer : python test_controls.py
"""
from pathlib import Path

import pandas as pd

from app import REQUIRED_COLUMNS, build_report, load_file, run_controls

data = {name: load_file(f"data/{name}.xlsx", name)[0] for name in REQUIRED_COLUMNS}
results = run_controls(**data)
exc = results["exceptions"]

# 1. Anomalies détectées == anomalies attendues (ni oubli, ni fausse alerte, cas témoins compris)
expected = pd.read_excel("anomalies_attendues.xlsx")
expected = expected[expected["ExpectedDetection"] == "Oui"]
wanted = set(zip(expected["File"], expected["ExcelRow"], expected["ExpectedAnomalyType"], expected["ExpectedSeverity"]))
found = set(zip(exc["Source"], exc["Row"], exc["AnomalyType"], exc["Severity"]))
assert found == wanted, f"\nNon détectées : {wanted - found}\nEn trop : {found - wanted}"

# 2. Le calcul de variation est juste sur un cas connu (98.00 -> 104.37 = +6.50 %)
moves = exc[exc["AnomalyType"] == "Variation de prix importante"]
assert "+6.50 %" in set(moves["ObservedValue"]), moves["ObservedValue"].tolist()

# 3. Les seuils sont bien paramétrables : avec un seuil à 50 %, plus aucune variation signalée
relaxed = run_controls(**data, price_thresholds={"CLO Debt": 50, "ABS": 50, "CLO Equity": 50})["exceptions"]
assert "Variation de prix importante" not in set(relaxed["AnomalyType"])

print(f"OK : {len(found)} anomalies détectées, identiques aux {len(wanted)} attendues.")

Path("reporting_controles.xlsx").write_bytes(build_report(results))
print("reporting_controles.xlsx généré")

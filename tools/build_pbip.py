"""
Writes the Power BI project (PBIP) for the FurniturePro report.

    python tools/build_pbip.py

Creates powerbi/FurniturePro.pbip with its semantic model (Power Query + DAX, in tools/model.py)
and its report pages (tools/report.py). Run it again after changing either file.
"""
import json
import sys
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import model, report  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
NAME = "FurniturePro"
OUT = ROOT / "powerbi"


def platform(kind):
    return json.dumps({"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                       "metadata": {"type": kind, "displayName": NAME},
                       "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"furniturepro/{kind}"))}}, indent=2) + "\n"


def main():
    sm, rp = OUT / f"{NAME}.SemanticModel", OUT / f"{NAME}.Report"
    import shutil
    if (sm / "definition").exists():
        shutil.rmtree(sm / "definition")
    model.build_model(sm)
    pages = report.build_report(rp, NAME, ROOT / "tools" / "Fluent2-CY26SU10.json")
    model.write(sm / ".platform", platform("SemanticModel"))
    model.write(rp / ".platform", platform("Report"))
    model.write(OUT / f"{NAME}.pbip", json.dumps({"version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                                  "settings": {"enableAutoRecovery": True}}, indent=2) + "\n")
    print(f"{len(model.TABLES)} tables, {len(model.MEASURES)} measures, {len(pages)} pages, {sum(len(p.visuals) for p in pages)} visuals")


if __name__ == "__main__":
    main()

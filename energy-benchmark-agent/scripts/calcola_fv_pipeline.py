"""
Pipeline completa per il calcolo dell'intervento fotovoltaico.

1. Triggera il workflow GitHub per le chiamate PVGIS
2. Attende il completamento e legge i risultati
3. Esegue il calcolo benchmark localmente
4. Genera il report .docx

Uso:
    python scripts/calcola_fv_pipeline.py \
        --modello Modello_energetico.xlsx \
        --consumi Analisi_Consumi.xlsx \
        --pod IT001E02519429 \
        --pun 0.108 \
        --superfici superfici.json \
        --github-token <TOKEN> \
        --out Report_FV.docx
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import uuid
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.parser import parse_excel as load_energy_model
from interventions.fotovoltaico import SuperficieFV, calcola
from reporting.build_report import costruisci_relazione

GITHUB_REPO   = "Grafone000/intervention-agent"
GITHUB_BRANCH = "claude/cool-hawking-ycftnx"
WORKFLOW_FILE = "pvgis_fetch.yml"
API_BASE      = "https://api.github.com"


def _gh(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}


def trigger_workflow(token: str, inputs: dict) -> None:
    url = f"{API_BASE}/repos/{GITHUB_REPO}/actions/workflows/{WORKFLOW_FILE}/dispatches"
    r = requests.post(url, headers=_gh(token),
                      json={"ref": GITHUB_BRANCH, "inputs": inputs})
    if r.status_code not in (200, 204):
        sys.exit(f"[ERRORE] Trigger workflow fallito ({r.status_code}): {r.text}")
    print("  Workflow triggerato.")


def wait_for_workflow(token: str, run_id_marker: str, timeout: int = 300) -> None:
    url = f"{API_BASE}/repos/{GITHUB_REPO}/actions/workflows/{WORKFLOW_FILE}/runs"
    print("  Attendo completamento workflow", end="", flush=True)
    deadline = time.time() + timeout
    while time.time() < deadline:
        r = requests.get(url, headers=_gh(token),
                         params={"branch": GITHUB_BRANCH, "per_page": 5})
        r.raise_for_status()
        runs = r.json().get("workflow_runs", [])
        for run in runs:
            if run_id_marker in (run.get("display_title") or "") or \
               run_id_marker in (run.get("head_commit", {}).get("message") or ""):
                status     = run["status"]
                conclusion = run["conclusion"]
                if status == "completed":
                    if conclusion != "success":
                        sys.exit(f"\n[ERRORE] Workflow terminato con: {conclusion}")
                    print(" completato.")
                    return
                print(".", end="", flush=True)
                break
        time.sleep(8)
    sys.exit("\n[ERRORE] Timeout attesa workflow.")


def read_pvgis_result(token: str, pvgis_run_id: str) -> dict:
    path = f"pvgis_results/{pvgis_run_id}.json"
    url  = f"{API_BASE}/repos/{GITHUB_REPO}/contents/{path}"
    r = requests.get(url, headers=_gh(token), params={"ref": GITHUB_BRANCH})
    if r.status_code == 404:
        sys.exit(f"[ERRORE] File risultati non trovato: {path}")
    r.raise_for_status()
    import base64
    content = base64.b64decode(r.json()["content"]).decode()
    return json.loads(content)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Pipeline calcolo FV con PVGIS via GitHub Actions")
    p.add_argument("--modello",       required=True)
    p.add_argument("--consumi",       required=True)
    p.add_argument("--pod",           required=True, nargs="+")
    p.add_argument("--pun",           type=float, required=True)
    p.add_argument("--superfici",     required=True)
    p.add_argument("--github-token",  required=True)
    p.add_argument("--out",           default="Report_FV.docx")
    p.add_argument("--prezzo-kwh",    type=float, default=None)
    p.add_argument("--anno-pvgis",    type=int,   default=2023)
    p.add_argument("--loss",          type=float, default=14.0)
    p.add_argument("--fabbricati",    nargs="*",  default=None)
    p.add_argument("--edifici",       nargs="*",  default=None)
    p.add_argument("--timeout",       type=int,   default=300)
    return p.parse_args()


def main() -> None:
    args = parse_args()

    modello_path = Path(args.modello)
    consumi_path = Path(args.consumi)
    sup_path     = Path(args.superfici)

    for p, label in [(modello_path, "modello"), (consumi_path, "consumi"), (sup_path, "superfici")]:
        if not p.exists():
            sys.exit(f"[ERRORE] File {label} non trovato: {p}")

    print(f"Carico modello energetico: {modello_path}")
    model = load_energy_model(modello_path)

    with open(sup_path, encoding="utf-8") as f:
        raw_sup = json.load(f)
    superfici = [SuperficieFV(**s) for s in raw_sup]
    print(f"Superfici: {len(superfici)}")
    for s in superfici:
        print(f"  {s.id}: {s.peak_power_kwp:.2f} kWp, slope={s.slope}, az={s.azimuth}")

    pvgis_run_id = str(uuid.uuid4())[:8]
    print(f"\n[1/4] Triggero workflow PVGIS (run_id={pvgis_run_id})...")
    trigger_workflow(
        token=args.github_token,
        inputs={
            "run_id":         pvgis_run_id,
            "lat":            str(model.lat),
            "lon":            str(model.lon),
            "superfici_json": json.dumps([s.model_dump() for s in superfici]),
            "anno_pvgis":     str(args.anno_pvgis),
            "loss":           str(args.loss),
        },
    )

    print(f"[2/4] Attendo completamento workflow (timeout {args.timeout}s)...")
    time.sleep(5)
    wait_for_workflow(args.github_token, pvgis_run_id, timeout=args.timeout)

    print("[3/4] Leggo risultati PVGIS dal repo...")
    pvgis_data = read_pvgis_result(args.github_token, pvgis_run_id)
    produzione_oraria = pvgis_data["produzione_oraria_totale"]
    print(f"  Produzione totale: {pvgis_data['e_prodotta_totale_kwh']:,.0f} kWh/anno")
    for riga in pvgis_data["superfici"]:
        print(f"  {riga['superficie_id']}: {riga['e_prodotta_kwh']:,.0f} kWh "
              f"({riga['ore_equivalenti']:.0f} h eq)")

    print("\n[4/4] Calcolo benchmark e generazione report...")
    prezzo_kwh = args.prezzo_kwh or model.parametri.get("prezzo_kwh", 0.21)

    parametri: dict = {
        "lat":                             model.lat,
        "lon":                             model.lon,
        "superfici":                       [s.model_dump() for s in superfici],
        "percorso_consumi":                str(consumi_path),
        "pod_selezionati":                 args.pod,
        "pun_euro_kwh":                    args.pun,
        "prezzo_kwh":                      prezzo_kwh,
        "anno_pvgis":                      args.anno_pvgis,
        "loss":                            args.loss,
        "fabbricati":                      args.fabbricati or [],
        "edifici":                         args.edifici or [],
        "_produzione_oraria_precalcolata": produzione_oraria,
        "_pvgis_righe":                    pvgis_data["superfici"],
    }

    result = calcola(model, parametri)

    sep = '─' * 40
    print(f"\n  {sep}")
    print(f"  Potenza totale:      {result.potenza_totale_kwp:.2f} kWp")
    print(f"  Produzione annua:    {result.e_prodotta_kwh:,.0f} kWh")
    print(f"  Autoconsumo:         {result.e_autoconsumata_kwh:,.0f} kWh  ({result.quota_autoconsumo*100:.1f}%)")
    print(f"  Immissione in rete:  {result.e_immessa_kwh:,.0f} kWh")
    print(f"  Investimento totale: {result.investimento_totale_euro:,.0f} EUR")
    print(f"  VAN (att.):          {result.van_attualizzato.van:,.0f} EUR")
    print(f"  TR semplice:         {result.van_semplice.tr:.2f} anni")
    if result.van_attualizzato.tir:
        print(f"  TIR:                 {result.van_attualizzato.tir*100:.2f}%")
    print(f"  {sep}\n")

    out_path = Path(args.out)
    costruisci_relazione([result], str(out_path))
    print(f"Report salvato: {out_path.resolve()}")


if __name__ == "__main__":
    main()

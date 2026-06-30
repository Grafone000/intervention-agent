"""
Script CLI per il calcolo dell'intervento fotovoltaico.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.parser import parse_excel as load_energy_model
from interventions.fotovoltaico import calcola, SuperficieFV
from reporting.build_report import costruisci_relazione


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Calcolo intervento fotovoltaico")
    p.add_argument("--modello",   required=True)
    p.add_argument("--consumi",   required=True)
    p.add_argument("--pod",       required=True, nargs="+")
    p.add_argument("--pun",       type=float, required=True)
    p.add_argument("--out",       default="Report_FV.docx")
    p.add_argument("--superfici", default=None)
    p.add_argument("--prezzo-kwh", type=float, default=None)
    p.add_argument("--anno-pvgis", type=int, default=2023)
    p.add_argument("--loss",      type=float, default=14.0)
    p.add_argument("--fabbricati", nargs="*", default=None)
    p.add_argument("--edifici",    nargs="*", default=None)
    return p.parse_args()


def main() -> None:
    args = parse_args()
    modello_path = Path(args.modello)
    consumi_path = Path(args.consumi)
    out_path     = Path(args.out)

    if not modello_path.exists():
        sys.exit(f"[ERRORE] File modello non trovato: {modello_path}")
    if not consumi_path.exists():
        sys.exit(f"[ERRORE] File consumi non trovato: {consumi_path}")

    print(f"Carico modello energetico: {modello_path}")
    model = load_energy_model(modello_path)

    if args.superfici:
        sup_path = Path(args.superfici)
        if not sup_path.exists():
            sys.exit(f"[ERRORE] File superfici non trovato: {sup_path}")
        with open(sup_path, encoding="utf-8") as f:
            raw = json.load(f)
        superfici = [SuperficieFV(**s) for s in raw]
        print(f"Superfici caricate da JSON: {len(superfici)} superficie/i")
    else:
        raw_sup = model.parametri.get("superfici")
        if not raw_sup:
            sys.exit(
                "[ERRORE] Nessuna superficie trovata nel modello energetico. "
                "Usa --superfici <file.json> per fornirle esternamente."
            )
        superfici = [SuperficieFV(**s) for s in raw_sup]
        print(f"Superfici lette dal modello: {len(superfici)} superficie/i")

    prezzo_kwh = args.prezzo_kwh or model.parametri.get("prezzo_kwh", 0.21)

    parametri: dict = {
        "lat":              model.lat,
        "lon":              model.lon,
        "superfici":        [s.model_dump() for s in superfici],
        "percorso_consumi": str(consumi_path),
        "pod_selezionati":  args.pod,
        "pun_euro_kwh":     args.pun,
        "prezzo_kwh":       prezzo_kwh,
        "anno_pvgis":       args.anno_pvgis,
        "loss":             args.loss,
        "fabbricati":       args.fabbricati or [],
        "edifici":          args.edifici or [],
    }

    print(f"\nAvvio calcolo FV — {len(superfici)} superficie/i, POD: {args.pod}")
    print(f"Lat: {model.lat}, Lon: {model.lon}, PUN: {args.pun} €/kWh, Anno PVGIS: {args.anno_pvgis}")
    print("Chiamate PVGIS in corso...\n")

    result = calcola(model, parametri)

    print(f"  Potenza totale:      {result.potenza_totale_kwp:.2f} kWp")
    print(f"  Produzione annua:    {result.e_prodotta_kwh:,.0f} kWh")
    print(f"  Autoconsumo:         {result.e_autoconsumata_kwh:,.0f} kWh ({result.quota_autoconsumo*100:.1f}%)")
    print(f"  Immissione in rete:  {result.e_immessa_kwh:,.0f} kWh")
    print(f"  Investimento totale: {result.investimento_totale_euro:,.0f} €")
    print(f"  VAN (att.):          {result.van_attualizzato.van:,.0f} €")
    print(f"  TR semplice:         {result.van_semplice.tr:.2f} anni")
    print(f"  TIR:                 {result.van_attualizzato.tir*100:.2f}%" if result.van_attualizzato.tir else "  TIR: n.d.")
    print()

    print(f"Generazione report: {out_path}")
    costruisci_relazione([result], str(out_path))
    print(f"Report salvato: {out_path}")


if __name__ == "__main__":
    main()

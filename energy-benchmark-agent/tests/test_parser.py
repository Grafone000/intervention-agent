"""
Test del parser Excel → EnergyModel.
Usa il file reale: Modello_energetico_Salerno_rev0.xlsx
"""

import pytest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.parser import parse_excel
from core.energy_model import EnergyModel, Utenza

EXCEL_PATH = Path("/root/.claude/uploads/c4385e63-15b9-5d14-8081-c1145e8ff6c3/7ba488ef-Modello_energetico_Salerno_rev0.xlsx")


def test_parsing_completo():
    """Il file reale deve essere parsato senza errori bloccanti."""
    model = parse_excel(EXCEL_PATH, regione="Marche")
    assert isinstance(model, EnergyModel)
    assert len(model.utenze) > 0
    assert model.regione == "Marche"
    assert model.nome_progetto == "7ba488ef-Modello_energetico_Salerno_rev0"


def test_conteggio_utenze():
    """Deve parsare esattamente 465 utenze (righe con Utenza valorizzata)."""
    model = parse_excel(EXCEL_PATH)
    assert len(model.utenze) == 465


def test_utenze_illuminazione():
    """Le utenze di illuminazione devono essere 179."""
    model = parse_excel(EXCEL_PATH)
    ill = model.utenze_illuminazione()
    assert len(ill) == 179
    assert all(u.uso_energetico == "Illuminazione" for u in ill)


def test_utenze_relamping_escluse_led():
    """Le utenze con tipologia contenente 'LED' devono essere escluse dal relamping."""
    model = parse_excel(EXCEL_PATH)
    candidati = model.utenze_relamping()
    for u in candidati:
        assert u.tipologia is not None
        assert "led" not in u.tipologia.lower(), f"LED non escluso: {u.tipologia}"


def test_ricalcolo_cella_errata():
    """La riga con h_eq='stile ' (riga 3) deve avere h_eq ricalcolato correttamente."""
    model = parse_excel(EXCEL_PATH)
    # Riga 3: Pompa di calore - C - CDZ01, h/d=11, d/w=5, w/y=26, f_util=0.55
    riga_errata = next(
        u for u in model.utenze if u.utenza == "Pompa di calore - C - CDZ01"
        and abs(u.potenza_unitaria_kw - 3.51) < 0.01
    )
    assert abs(riga_errata.h_equivalenti - 786.5) < 1.0, (
        f"h_eq atteso ~786.5, ottenuto {riga_errata.h_equivalenti}"
    )
    assert any("ricalcolat" in n for n in riga_errata.note_parser), (
        "Manca nota di ricalcolo nelle note_parser"
    )


def test_tipologia_mancante_warning():
    """Utenze illuminazione senza tipologia devono avere nota nel note_parser."""
    model = parse_excel(EXCEL_PATH)
    senza_tipo = [
        u for u in model.utenze_illuminazione() if u.tipologia is None
    ]
    for u in senza_tipo:
        assert any("Tipologia mancante" in n for n in u.note_parser), (
            f"Manca warning Tipologia mancante per: {u.utenza} (POD {u.pod})"
        )


def test_file_non_trovato():
    """Deve sollevare FileNotFoundError se il file non esiste."""
    with pytest.raises(FileNotFoundError):
        parse_excel("/percorso/inesistente/file.xlsx")


def test_consumo_coerente():
    """Il consumo deve essere ≈ Pot_totale × h_eq × f_contemp per ogni utenza."""
    model = parse_excel(EXCEL_PATH)
    for u in model.utenze[:20]:  # verifica prime 20 righe
        atteso = u.potenza_totale_kw * u.h_equivalenti * u.fatt_contemporaneita
        assert abs(u.consumo_kwh - atteso) < 1.0, (
            f"{u.utenza}: consumo={u.consumo_kwh:.2f}, atteso={atteso:.2f}"
        )

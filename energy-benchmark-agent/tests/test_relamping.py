"""
Test dell'intervento di relamping.
Usa il file reale: Modello_energetico_Salerno_rev0.xlsx
"""

import pytest
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.parser import parse_excel
from interventions.relamping import calcola, RelampingResult, NOME_INTERVENTO
from interventions.relamping_map import normalizza_tipologia, converti_lampada

EXCEL = Path("/root/.claude/uploads/c4385e63-15b9-5d14-8081-c1145e8ff6c3/7ba488ef-Modello_energetico_Salerno_rev0.xlsx")


@pytest.fixture(scope="module")
def result() -> RelampingResult:
    model = parse_excel(EXCEL)
    return calcola(model, {})


# --- Mappa sinonimi ---

def test_normalizzazione_plafoniera_neon():
    assert normalizza_tipologia("Plafoniera tipo Neon") == "plafoniera_neon"

def test_normalizzazione_faretto():
    assert normalizza_tipologia("Faretti") == "faretto_HID"
    assert normalizza_tipologia("Faro ioduri") == "faretto_HID"

def test_normalizzazione_lampione():
    assert normalizza_tipologia("Lampione SAP") == "lampione_SAP"

def test_normalizzazione_sconosciuta():
    from interventions.relamping_map import FALLBACK_CATEGORIA
    assert normalizza_tipologia("XYZ sconosciuta 2000") == FALLBACK_CATEGORIA

def test_led_esclusi_dalla_mappa():
    # LED non devono mai arrivare alla mappa — filtrati da utenze_relamping()
    model = parse_excel(EXCEL)
    for u in model.utenze_relamping():
        assert "led" not in u.tipologia.lower()


# --- Conversione potenze ---

def test_conversione_plafoniera_neon_036():
    """Plafoniera Neon 36W: riduzione 50% → 18W."""
    conv = converti_lampada("Plafoniera tipo Neon", 0.036)
    assert abs(conv.p_post_kw - 0.018) < 1e-4
    assert conv.costo_lampada_euro == 25.0

def test_conversione_plafoniera_neon_058():
    """Plafoniera Neon 58W: riduzione 52% → 27.84W."""
    conv = converti_lampada("Plafoniera tipo Neon", 0.058)
    atteso = 0.058 * (1 - 52.0 / 100)
    assert abs(conv.p_post_kw - atteso) < 1e-4

def test_interpolazione_faretto_HID():
    """Faretto HID 0.35 kW: interpolazione tra 0.25 (40%) e 0.40 (62%)."""
    conv = converti_lampada("Faretti", 0.350)
    frac = (0.350 - 0.250) / (0.400 - 0.250)
    riduzione_attesa = 40.0 + frac * (62.0 - 40.0)
    p_post_atteso = 0.350 * (1 - riduzione_attesa / 100)
    assert abs(conv.p_post_kw - p_post_atteso) < 1e-4
    assert conv.confidenza == "Media"

def test_estrapolazione_segnalata():
    """Potenza fuori range deve avere confidenza Bassa."""
    conv = converti_lampada("Faretti", 5.0)
    assert conv.confidenza == "Bassa"
    assert any("estrapolat" in n for n in conv.note)


# --- Calcoli aggregati ---

def test_result_tipo(result):
    assert isinstance(result, RelampingResult)
    assert result.nome_intervento == NOME_INTERVENTO

def test_numero_righe(result):
    """Devono esserci esattamente le utenze relamping del modello."""
    model = parse_excel(EXCEL)
    assert len(result.righe) == len(model.utenze_relamping())

def test_consumo_ante_coerente(result):
    """consumo_ante = P_tot_ante × f_cont × h_eq per ogni riga."""
    for r in result.righe[:10]:
        atteso = r.p_tot_ante_kw * r.fatt_contemporaneita * r.h_equivalenti
        assert abs(r.consumo_ante_kwh - atteso) < 0.01, f"{r.tipologia_originale}: {r.consumo_ante_kwh} vs {atteso}"

def test_risparmio_coerente(result):
    """risparmio_ee = consumo_ante - consumo_post."""
    for r in result.righe[:10]:
        assert abs(r.risparmio_ee_kwh - (r.consumo_ante_kwh - r.consumo_post_kwh)) < 0.01

def test_cb_h_post(result):
    """h_post = E_post / P_tot_post."""
    cb = result.cb
    atteso = cb.e_post_kwh / cb.p_tot_post_kw
    assert abs(cb.h_post - atteso) < 0.01

def test_cb_tep_floor(result):
    """tep_floor deve essere floor(tep_risparmiati)."""
    cb = result.cb
    assert cb.tep_floor == math.floor(cb.tep_risparmiati)

def test_cb_incentivo_annuo(result):
    """incentivo_annuo = tep_floor × 250."""
    cb = result.cb
    assert abs(cb.incentivo_annuo - cb.tep_floor * 250.0) < 0.01


# --- VAN e indicatori economici ---

def test_van_fc_anno8_uguale_van(result):
    """Il VAN deve essere uguale all'FC dell'anno 8."""
    for scenario in [
        result.van_attualizzato_con_incentivi,
        result.van_attualizzato_senza_incentivi,
        result.van_semplice_con_incentivi,
        result.van_semplice_senza_incentivi,
    ]:
        assert abs(scenario.van - scenario.flussi_cassa[-1]) < 0.1, scenario.nome

def test_van_con_incentivi_maggiore_senza(result):
    """Il VAN con incentivi deve essere >= quello senza incentivi."""
    assert result.van_attualizzato_con_incentivi.van >= result.van_attualizzato_senza_incentivi.van
    assert result.van_semplice_con_incentivi.van >= result.van_semplice_senza_incentivi.van

def test_tr_senza_incentivi(result):
    """TR senza incentivi = Inv / risparmio_annuo."""
    s = result.van_attualizzato_senza_incentivi
    atteso = s.investimento / s.risparmio_annuo_euro
    assert abs(s.tr - atteso) < 0.01

def test_tir_solo_attualizzati(result):
    """TIR calcolato solo per i casi attualizzati."""
    assert result.van_attualizzato_con_incentivi.tir is not None
    assert result.van_attualizzato_senza_incentivi.tir is not None
    assert result.van_semplice_con_incentivi.tir is None
    assert result.van_semplice_senza_incentivi.tir is None

def test_indice_profitto(result):
    """IP = VAN / Investimento."""
    for s in [result.van_attualizzato_con_incentivi, result.van_attualizzato_senza_incentivi]:
        atteso = s.van / s.investimento
        assert abs(s.indice_profitto - atteso) < 1e-3, s.nome

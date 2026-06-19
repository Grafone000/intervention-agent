"""
Schema dati canonico per il modello energetico.

Contratto tra il parser Excel e i tool di intervento:
ogni tool riceve un EnergyModel e non sa nulla del formato sorgente.

Basato sul file reale: Modello_energetico_Salerno_rev0.xlsx, foglio ModEE.
"""

from __future__ import annotations
from typing import Optional, List
from pydantic import BaseModel, Field


class Utenza(BaseModel):
    """
    Singola utenza energetica censita nel modello.
    Copre tutte le tipologie (illuminazione, climatizzazione, FM, ecc.).
    I tool di intervento filtrano sul campo `uso_energetico`.
    """

    # --- Identificazione ---
    utenza: str = Field(description="Nome/descrizione utenza (col C)")
    tipologia: Optional[str] = Field(None, description="Tipo apparecchio (col D). Può essere None.")
    pod: str = Field(description="Codice POD punto di prelievo (col E)")
    edificio_zona: str = Field(description="Edificio o zona di pertinenza (col F)")
    competenza: str = Field(description="Competenza aziendale principale (col K)")

    # --- Classificazione energetica ---
    uso_energetico: str = Field(
        description=(
            "Categoria d'uso: 'Illuminazione', 'Climatizzazione Civile - Generazione', "
            "'Climatizzazione Civile - Distribuzione', 'Climatizzazione Apparati Tecnologici', "
            "'FM Utenze', 'Apparati tecnologici', 'ACS'. (col M)"
        )
    )

    # --- Dati tecnici ---
    n_utenze: int = Field(description="Numero di corpi/apparecchi identici (col N)")
    potenza_unitaria_kw: float = Field(description="Potenza del singolo corpo in kW (col P)")
    potenza_totale_kw: float = Field(description="Potenza totale = P × N / COP (col Q)")

    # --- Parametri orari ---
    fatt_utilizzo: float = Field(description="Fattore di utilizzo 0-1 (col S)")
    fatt_contemporaneita: float = Field(description="Fattore di contemporaneità 0-1 (col T)")
    ore_giorno: float = Field(description="Ore di funzionamento al giorno h/d (col U)")
    giorni_settimana: float = Field(description="Giorni di funzionamento a settimana d/w (col V)")
    settimane_anno: float = Field(description="Settimane di funzionamento all'anno w/y (col W)")

    # --- Risultati energetici (ricalcolati dal parser se cella errata) ---
    ore_anno: float = Field(description="h/y = h/d × d/w × w/y (col X)")
    h_equivalenti: float = Field(description="h_eq = h/y × f_utilizzo (col Y)")
    consumo_kwh: float = Field(description="Consumo = Pot_tot × h_eq × f_contemp (col Z)")
    tep: float = Field(description="Energia primaria in tep = Consumo × 0.000187 (col AB)")
    costi_euro: float = Field(description="Costo energetico = Consumo × 0.21 €/kWh (col AC)")
    co2_kg: float = Field(description="CO2 emessa = Consumo × 294.784 / 1000 (col AD)")

    # --- Warning interni ---
    note_parser: List[str] = Field(
        default_factory=list,
        description="Warning generati dal parser (dati ricalcolati, ambiguità, ecc.)"
    )


class EnergyModel(BaseModel):
    """
    Modello energetico completo di un sito, prodotto dal parser e consumato dai tool.
    """
    nome_progetto: Optional[str] = Field(None, description="Nome del file/progetto sorgente")
    regione: Optional[str] = Field(None, description="Regione (parametro esterno, per prezziario FV)")
    utenze: List[Utenza] = Field(default_factory=list)

    def utenze_illuminazione(self) -> List[Utenza]:
        """Restituisce solo le utenze con uso_energetico == 'Illuminazione'."""
        return [u for u in self.utenze if u.uso_energetico == "Illuminazione"]

    def utenze_relamping(self) -> List[Utenza]:
        """
        Utenze candidate al relamping:
        - uso_energetico == 'Illuminazione'
        - tipologia NON contiene 'LED' (già sostituite)
        - tipologia non è None (altrimenti non mappabile)
        """
        candidati = []
        for u in self.utenze_illuminazione():
            if u.tipologia is None:
                continue
            if "led" in u.tipologia.lower():
                continue
            candidati.append(u)
        return candidati

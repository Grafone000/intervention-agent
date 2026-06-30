# Direttive di progetto — Energy Benchmark Agent

Convenzioni di calcolo da rispettare per gli interventi. Aggiornare questo file
quando l'utente fissa una nuova regola.

## Intervento Fotovoltaico (FTV)

- **Prezziario impianto: SEMPRE quello regionale delle Marche**, indipendentemente
  dalla regione del sito. Codificato in `interventions/fotovoltaico.py`
  (`_REGIONE_PREZZIARIO_FV = "marche"`). Non usare la regione del progetto né il
  fallback di mercato.
- **Investimento iniziale** = costo impianto (da prezziario Marche) + **15%** di
  progettazione.
- **Manutenzione annua** = **1%** di (impianto + progettazione), sottratta dal
  flusso di cassa.
- **Flusso netto annuo** = autoconsumo·prezzo_energia + immissione·PUN − manutenzione.
  - PUN è fornito dall'utente, **diverso** dal prezzo dell'energia.
  - Immissione in rete = produzione − autoconsumo.
- **Vita utile = 20 anni**; tasso di sconto = 6%.
- **Nessun incentivo e nessun Certificato Bianco** per il FTV.
- CO₂/TEP riferiti al **mancato prelievo da rete** (autoconsumo), coerenti con la
  tabella Emissioni PRE/POST del POD.
- Producibilità via PVGIS (una chiamata per superficie, somma oraria). PVGIS
  richiede accesso di rete a `re.jrc.ec.europa.eu`.
- Selezione POD: i consumi si leggono dal foglio `Input POD orario ATTIVA`
  (codici POD nell'header dalla colonna DZ/130); si sommano solo i POD richiesti.

## Intervento Relamping

- Vita utile = **8 anni**.
- Usa i Certificati Bianchi (250 €/tep) e gli scenari con/senza incentivi.

## Nota

La vita utile è specifica per intervento (relamping 8, FTV 20); altri interventi
possono avere durate diverse.

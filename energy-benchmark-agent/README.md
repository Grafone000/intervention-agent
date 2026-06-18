# Energy Benchmark Agent

Agente per il calcolo del benchmark di interventi di efficientamento energetico.
Prende in input un modello energetico (file Excel con utenze, consumi, POD, etichette edificio),
calcola il benchmark per diversi interventi (fotovoltaico, relamping, coibentazione, ecc.)
e genera una relazione finale in `.docx` con descrizione interventi, tabelle e grafici.

L'architettura è a plugin: aggiungere un nuovo intervento richiede solo l'aggiunta di un file
in `interventions/`, senza modificare codice esistente.

## Struttura

```
energy-benchmark-agent/
├── core/           # Schema dati canonico, parser Excel, registry interventi
├── interventions/  # Un file per intervento (plugin architecture)
├── pricing/        # Lookup prezzi per regione/anno da CSV
├── reporting/      # Generazione grafici e documento .docx
├── skills/         # SKILL.md per orchestratore e template interventi
├── tests/          # Test suite
└── data/           # File Excel di lavoro (verificare .gitignore per dati sensibili)
```

## Stato attuale

**Scaffolding iniziale** — logica di business non ancora implementata.
I file contengono strutture dati placeholder (Pydantic), interfacce e firme di funzione.
La logica reale (parser Excel, formule di calcolo, prezzi) verrà aggiunta negli step successivi.

"""
Auto-discovery dei tool di intervento energetico.

Scansiona la cartella interventions/, trova tutti i moduli che espongono
una funzione calcola() compatibile con il protocollo BaseIntervention,
e li registra automaticamente.

ESTENSIBILITÀ: per aggiungere un nuovo intervento basta creare un file
in interventions/ — registry.py non va mai modificato.
"""

import importlib
import pkgutil
from typing import Callable, Dict
from pathlib import Path

# Moduli da escludere dall'auto-discovery
_EXCLUDED = {"base", "_template", "__init__"}


def discover_interventions() -> Dict[str, Callable]:
    """
    Scansiona il package `interventions` e restituisce un dizionario
    {nome_intervento: funzione_calcola}.

    Un modulo viene incluso se:
    - il suo nome non è in _EXCLUDED
    - espone una funzione `calcola` al top-level
    - espone un attributo `NOME_INTERVENTO: str` al top-level

    Returns:
        Dict mapping nome_intervento -> funzione calcola()
    """
    import interventions as _pkg

    registry: Dict[str, Callable] = {}
    pkg_path = Path(_pkg.__file__).parent

    for finder, module_name, is_pkg in pkgutil.iter_modules([str(pkg_path)]):
        if module_name in _EXCLUDED:
            continue
        full_name = f"interventions.{module_name}"
        try:
            module = importlib.import_module(full_name)
        except ImportError as e:
            print(f"[registry] Impossibile importare {full_name}: {e}")
            continue

        if not hasattr(module, "calcola"):
            print(f"[registry] {full_name} non espone 'calcola', ignorato.")
            continue

        nome = getattr(module, "NOME_INTERVENTO", module_name)
        registry[nome] = module.calcola

    return registry


# Singleton: caricato una volta sola all'import
REGISTRY: Dict[str, Callable] = discover_interventions()

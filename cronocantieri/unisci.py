"""Fusione di più cronoprogrammi di impresa in un unico cronoprogramma di cantiere.

Il project manager riceve n file (uno per impresa/lotto). Questo modulo li
importa tutti e li assembla in un singolo Cantiere, dove ogni file diventa un
Lotto. Le date di inizio/fine del cantiere sono dedotte dagli estremi dei lotti
e l'avanzamento (di lotto e globale) viene calcolato automaticamente
dall'analisi, pesato sugli importi.
"""

from __future__ import annotations

from pathlib import Path

from .importa import CronoprogrammaImportato, importa_file
from .models import Cantiere


def unisci_cronoprogrammi(
    percorsi: list[str | Path],
    cantiere_nome: str | None = None,
    tipo: str = "pubblico",
    committente: str = "",
) -> tuple[Cantiere, list[CronoprogrammaImportato]]:
    """Importa i file indicati e li unisce in un unico Cantiere.

    Se `cantiere_nome` non è specificato, usa il nome del cantiere letto dal
    primo file. Restituisce il Cantiere e la lista dei risultati di import
    (utile per un riepilogo a video).
    """
    if not percorsi:
        raise ValueError("Nessun file da unire.")

    importati = [importa_file(p) for p in percorsi]
    nome = cantiere_nome or importati[0].cantiere_nome

    cantiere = Cantiere(
        nome=nome,
        tipo=tipo,
        committente=committente,
        descrizione="Cronoprogramma totale generato dall'unione dei file imprese",
        lotti=[imp.lotto for imp in importati],
    )

    # date complessive dedotte dai lotti
    inizi = [l.data_inizio for l in cantiere.lotti if l.data_inizio]
    fini = [l.data_fine_prevista for l in cantiere.lotti if l.data_fine_prevista]
    cantiere.data_inizio = min(inizi) if inizi else None
    cantiere.data_fine_prevista = max(fini) if fini else None

    return cantiere, importati

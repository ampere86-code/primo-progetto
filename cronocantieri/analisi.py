"""Analisi dell'avanzamento e rilevamento dei ritardi.

Idea di base
------------
Per ogni lotto conosciamo data di inizio e fine prevista. Alla data di
oggi (o a una data scelta) possiamo calcolare quanta parte del tempo è
trascorsa: questo è l'avanzamento "atteso" se il lavoro procedesse in
modo lineare.

    avanzamento_atteso = tempo_trascorso / durata_totale * 100

Confrontando l'avanzamento ATTESO con quello REALE (avanzamento_pct)
otteniamo lo scostamento:

    scostamento = avanzamento_reale - avanzamento_atteso

  * scostamento negativo  -> l'impresa è IN RITARDO
  * scostamento ~ 0        -> in linea
  * scostamento positivo  -> in anticipo

L'avanzamento globale del cantiere è la media degli avanzamenti dei
lotti PESATA sull'importo di contratto: un lotto da 1.000.000 € pesa più
di uno da 50.000 €.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from .models import Cantiere, Lotto


@dataclass
class StatoLotto:
    """Risultato dell'analisi di un singolo lotto a una certa data."""

    lotto: Lotto
    avanzamento_reale: float
    avanzamento_atteso: float
    scostamento: float          # reale - atteso (negativo = ritardo)
    in_ritardo: bool
    importo_maturato: float     # dall'ultimo SAL, se presente

    @property
    def giudizio(self) -> str:
        if self.scostamento < -10:
            return "RITARDO GRAVE"
        if self.scostamento < -2:
            return "In ritardo"
        if self.scostamento > 10:
            return "In netto anticipo"
        if self.scostamento > 2:
            return "In anticipo"
        return "In linea"


@dataclass
class StatoCantiere:
    """Risultato dell'analisi complessiva di un cantiere."""

    cantiere: Cantiere
    avanzamento_globale_reale: float
    avanzamento_globale_atteso: float
    scostamento_globale: float
    stati_lotti: list[StatoLotto]
    importo_maturato_totale: float

    def lotti_in_ritardo(self) -> list[StatoLotto]:
        return [s for s in self.stati_lotti if s.in_ritardo]

    def lotto_piu_in_ritardo(self) -> StatoLotto | None:
        """Il lotto con lo scostamento peggiore (più negativo)."""
        if not self.stati_lotti:
            return None
        return min(self.stati_lotti, key=lambda s: s.scostamento)


def avanzamento_atteso(lotto: Lotto, alla_data: date) -> float:
    """% di avanzamento attesa per il lotto alla data indicata (0-100).

    Restituisce 0 se non sono note le date, 100 se la data supera la fine.
    """
    inizio = lotto.data_inizio
    fine = lotto.data_fine_prevista
    if not inizio or not fine or fine <= inizio:
        return 0.0
    if alla_data <= inizio:
        return 0.0
    if alla_data >= fine:
        return 100.0
    durata = (fine - inizio).days
    trascorso = (alla_data - inizio).days
    return round(trascorso / durata * 100, 1)


def analizza_lotto(lotto: Lotto, alla_data: date | None = None) -> StatoLotto:
    """Calcola lo stato di avanzamento di un singolo lotto."""
    alla_data = alla_data or date.today()
    atteso = avanzamento_atteso(lotto, alla_data)
    reale = lotto.avanzamento_pct
    scostamento = round(reale - atteso, 1)
    ultimo = lotto.ultimo_sal()
    return StatoLotto(
        lotto=lotto,
        avanzamento_reale=reale,
        avanzamento_atteso=atteso,
        scostamento=scostamento,
        in_ritardo=scostamento < -2,
        importo_maturato=ultimo.importo if ultimo else 0.0,
    )


def analizza_cantiere(cantiere: Cantiere, alla_data: date | None = None) -> StatoCantiere:
    """Calcola lo stato complessivo del cantiere e di tutti i suoi lotti."""
    alla_data = alla_data or date.today()
    stati = [analizza_lotto(l, alla_data) for l in cantiere.lotti]

    peso_totale = sum(l.importo_contratto for l in cantiere.lotti)

    def media_pesata(valore) -> float:
        if peso_totale > 0:
            somma = sum(
                valore(s) * s.lotto.importo_contratto for s in stati
            )
            return round(somma / peso_totale, 1)
        # senza importi, media semplice
        if stati:
            return round(sum(valore(s) for s in stati) / len(stati), 1)
        return 0.0

    glob_reale = media_pesata(lambda s: s.avanzamento_reale)
    glob_atteso = media_pesata(lambda s: s.avanzamento_atteso)

    return StatoCantiere(
        cantiere=cantiere,
        avanzamento_globale_reale=glob_reale,
        avanzamento_globale_atteso=glob_atteso,
        scostamento_globale=round(glob_reale - glob_atteso, 1),
        stati_lotti=stati,
        importo_maturato_totale=round(sum(s.importo_maturato for s in stati), 2),
    )

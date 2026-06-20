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

from .models import Attivita, Cantiere, Lotto


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


def _avanzamento_atteso_periodo(
    inizio: date | None, fine: date | None, alla_data: date
) -> float:
    """% attesa fra due date (progresso lineare). 0 se date mancanti/non valide."""
    if not inizio or not fine or fine <= inizio:
        return 0.0
    if alla_data <= inizio:
        return 0.0
    if alla_data >= fine:
        return 100.0
    return round((alla_data - inizio).days / (fine - inizio).days * 100, 1)


def avanzamento_da_attivita(lotto: Lotto, alla_data: date) -> float | None:
    """Calcola la % del lotto a partire dalle sue attività.

    Per ogni attività usa la % dichiarata, se presente; altrimenti la stima
    dalle date (progresso lineare). Le attività sono pesate sull'importo;
    se gli importi non sono disponibili, si usa la durata in giorni; in
    mancanza anche di quella, media semplice.

    Restituisce None se il lotto non ha attività (così l'analisi ricade
    sulla % manuale del lotto).
    """
    att = lotto.attivita
    if not att:
        return None

    pesi = [a.importo for a in att]
    if sum(pesi) <= 0:
        pesi = [a.durata_giorni() or 1 for a in att]
    totale = sum(pesi) or len(att)

    somma = 0.0
    for a, peso in zip(att, pesi):
        av = a.avanzamento_pct
        if av is None:  # non dichiarata: stima dalle date
            av = _avanzamento_atteso_periodo(a.data_inizio, a.data_fine, alla_data)
        somma += av * peso
    return round(somma / totale, 1)


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
    # se il lotto ha attività, la % reale è calcolata da esse; altrimenti
    # si usa la % impostata manualmente sul lotto.
    da_attivita = avanzamento_da_attivita(lotto, alla_data)
    reale = da_attivita if da_attivita is not None else lotto.avanzamento_pct
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

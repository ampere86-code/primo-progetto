"""Modello dati del cronoprogramma e salvataggio su file JSON.

Gerarchia:
    Cantiere (un appalto/commessa)
      └── Lotto (porzione di lavoro affidata a una singola Impresa)
            └── SAL (Stato di Avanzamento Lavori, con data, % e importo)

I dati sono salvati in JSON: un formato di testo leggibile e portabile,
che puoi copiare su qualsiasi PC insieme al programma.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from datetime import date
from pathlib import Path


def _parse_date(valore: str | None) -> date | None:
    """Converte una stringa 'AAAA-MM-GG' in date. Restituisce None se vuota."""
    if not valore:
        return None
    return date.fromisoformat(valore)


def _format_date(valore: date | None) -> str | None:
    """Converte una date in stringa 'AAAA-MM-GG'. Restituisce None se vuota."""
    return valore.isoformat() if valore else None


@dataclass
class SAL:
    """Stato di Avanzamento Lavori registrato a una certa data."""

    numero: int                 # progressivo del SAL (1, 2, 3...)
    data: date                  # data di emissione del SAL
    avanzamento_pct: float      # % di completamento dichiarata a quella data
    importo: float = 0.0        # valore economico maturato (euro)
    note: str = ""

    def to_dict(self) -> dict:
        d = asdict(self)
        d["data"] = _format_date(self.data)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "SAL":
        return cls(
            numero=d["numero"],
            data=_parse_date(d["data"]),
            avanzamento_pct=d.get("avanzamento_pct", 0.0),
            importo=d.get("importo", 0.0),
            note=d.get("note", ""),
        )


@dataclass
class Attivita:
    """Singola lavorazione/voce di un cronoprogramma di lotto.

    È il livello di dettaglio inviato dalle imprese: ogni riga del loro
    cronoprogramma (es. "Scavi", "Getto fondazioni", "Posa quadri") con le
    proprie date, l'importo (peso) e, se nota, la % di completamento.
    """

    nome: str
    data_inizio: date | None = None
    data_fine: date | None = None
    importo: float = 0.0                 # peso dell'attività (euro)
    avanzamento_pct: float | None = None  # None = non dichiarata (si stima dalle date)
    note: str = ""

    def durata_giorni(self) -> int:
        if self.data_inizio and self.data_fine and self.data_fine > self.data_inizio:
            return (self.data_fine - self.data_inizio).days
        return 0

    def to_dict(self) -> dict:
        return {
            "nome": self.nome,
            "data_inizio": _format_date(self.data_inizio),
            "data_fine": _format_date(self.data_fine),
            "importo": self.importo,
            "avanzamento_pct": self.avanzamento_pct,
            "note": self.note,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Attivita":
        return cls(
            nome=d["nome"],
            data_inizio=_parse_date(d.get("data_inizio")),
            data_fine=_parse_date(d.get("data_fine")),
            importo=d.get("importo", 0.0),
            avanzamento_pct=d.get("avanzamento_pct"),
            note=d.get("note", ""),
        )


@dataclass
class Lotto:
    """Porzione di cantiere affidata a una singola impresa.

    Più lotti lavorano in parallelo sullo stesso cantiere; insieme
    portano al completamento dell'appalto.
    """

    nome: str                       # es. "Lotto 2 - Impianti elettrici"
    impresa: str                    # impresa che esegue il lotto
    categoria: str = ""             # es. "edile", "impiantistico", "serramenti"
    responsabile: str = ""          # referente / direttore tecnico
    data_inizio: date | None = None
    data_fine_prevista: date | None = None
    importo_contratto: float = 0.0  # importo del lotto (euro), usato come peso
    avanzamento_pct: float = 0.0    # % attuale (manuale o ricalcolata dalle attività)
    note: str = ""
    sal: list[SAL] = field(default_factory=list)
    attivita: list[Attivita] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "nome": self.nome,
            "impresa": self.impresa,
            "categoria": self.categoria,
            "responsabile": self.responsabile,
            "data_inizio": _format_date(self.data_inizio),
            "data_fine_prevista": _format_date(self.data_fine_prevista),
            "importo_contratto": self.importo_contratto,
            "avanzamento_pct": self.avanzamento_pct,
            "note": self.note,
            "sal": [s.to_dict() for s in self.sal],
            "attivita": [a.to_dict() for a in self.attivita],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Lotto":
        return cls(
            nome=d["nome"],
            impresa=d["impresa"],
            categoria=d.get("categoria", ""),
            responsabile=d.get("responsabile", ""),
            data_inizio=_parse_date(d.get("data_inizio")),
            data_fine_prevista=_parse_date(d.get("data_fine_prevista")),
            importo_contratto=d.get("importo_contratto", 0.0),
            avanzamento_pct=d.get("avanzamento_pct", 0.0),
            note=d.get("note", ""),
            sal=[SAL.from_dict(s) for s in d.get("sal", [])],
            attivita=[Attivita.from_dict(a) for a in d.get("attivita", [])],
        )

    def ultimo_sal(self) -> SAL | None:
        """Restituisce il SAL più recente per numero, o None se non ce ne sono."""
        if not self.sal:
            return None
        return max(self.sal, key=lambda s: s.numero)

    def date_estremi_attivita(self) -> tuple[date | None, date | None]:
        """Data minima di inizio e massima di fine fra le attività del lotto."""
        inizi = [a.data_inizio for a in self.attivita if a.data_inizio]
        fini = [a.data_fine for a in self.attivita if a.data_fine]
        return (min(inizi) if inizi else None, max(fini) if fini else None)


@dataclass
class Cantiere:
    """Un cantiere / appalto, composto da uno o più lotti."""

    nome: str
    tipo: str = "pubblico"          # "pubblico" o "privato"
    descrizione: str = ""
    committente: str = ""
    data_inizio: date | None = None
    data_fine_prevista: date | None = None
    note: str = ""
    lotti: list[Lotto] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "nome": self.nome,
            "tipo": self.tipo,
            "descrizione": self.descrizione,
            "committente": self.committente,
            "data_inizio": _format_date(self.data_inizio),
            "data_fine_prevista": _format_date(self.data_fine_prevista),
            "note": self.note,
            "lotti": [l.to_dict() for l in self.lotti],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Cantiere":
        return cls(
            nome=d["nome"],
            tipo=d.get("tipo", "pubblico"),
            descrizione=d.get("descrizione", ""),
            committente=d.get("committente", ""),
            data_inizio=_parse_date(d.get("data_inizio")),
            data_fine_prevista=_parse_date(d.get("data_fine_prevista")),
            note=d.get("note", ""),
            lotti=[Lotto.from_dict(l) for l in d.get("lotti", [])],
        )

    def importo_totale(self) -> float:
        """Somma degli importi di contratto di tutti i lotti."""
        return sum(l.importo_contratto for l in self.lotti)


class Archivio:
    """Contenitore di tutti i cantieri, con salvataggio/caricamento su file."""

    LIMITE_CANTIERI = 20

    def __init__(self, percorso: str | Path = "dati/cantieri.json"):
        self.percorso = Path(percorso)
        self.cantieri: list[Cantiere] = []

    def carica(self) -> "Archivio":
        """Carica i cantieri dal file JSON, se esiste."""
        if self.percorso.exists():
            dati = json.loads(self.percorso.read_text(encoding="utf-8"))
            self.cantieri = [Cantiere.from_dict(c) for c in dati.get("cantieri", [])]
        return self

    def salva(self) -> None:
        """Salva tutti i cantieri sul file JSON (creando la cartella se serve)."""
        self.percorso.parent.mkdir(parents=True, exist_ok=True)
        dati = {"cantieri": [c.to_dict() for c in self.cantieri]}
        self.percorso.write_text(
            json.dumps(dati, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    def aggiungi_cantiere(self, cantiere: Cantiere) -> None:
        if len(self.cantieri) >= self.LIMITE_CANTIERI:
            raise ValueError(
                f"Raggiunto il limite di {self.LIMITE_CANTIERI} cantieri."
            )
        self.cantieri.append(cantiere)

    def trova(self, nome: str) -> Cantiere | None:
        for c in self.cantieri:
            if c.nome.lower() == nome.lower():
                return c
        return None

"""Esportazione del cronoprogramma in un file Excel con diagramma di Gantt.

Il file generato contiene, per ogni cantiere:
  * un foglio "Riepilogo" con l'andamento globale e per impresa;
  * un foglio "Gantt" con le barre temporali dei lotti, colorate in base
    al ritardo (verde = in linea, giallo = lieve ritardo, rosso = grave).

Usa la libreria openpyxl (installabile con: pip install openpyxl).
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .analisi import StatoCantiere, analizza_cantiere
from .models import Archivio, Cantiere

# Colori (esadecimale ARGB) usati nelle barre del Gantt
_VERDE = PatternFill("solid", fgColor="63BE7B")
_GIALLO = PatternFill("solid", fgColor="FFD666")
_ROSSO = PatternFill("solid", fgColor="F8696B")
_GRIGIO = PatternFill("solid", fgColor="D9D9D9")
_INTESTAZIONE = PatternFill("solid", fgColor="305496")

_BORDO = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
_BIANCO_BOLD = Font(bold=True, color="FFFFFF")
_BOLD = Font(bold=True)


def _colore_per_scostamento(scostamento: float) -> PatternFill:
    if scostamento < -10:
        return _ROSSO
    if scostamento < -2:
        return _GIALLO
    return _VERDE


def _foglio_riepilogo(wb: Workbook, stato: StatoCantiere) -> None:
    ws = wb.create_sheet(title=_nome_foglio(stato.cantiere.nome, "Riepilogo"))
    c = stato.cantiere

    ws["A1"] = f"CANTIERE: {c.nome}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Tipo: {c.tipo}   Committente: {c.committente}"
    ws["A3"] = (
        f"Avanzamento globale: {stato.avanzamento_globale_reale}% "
        f"(atteso {stato.avanzamento_globale_atteso}%)  "
        f"Scostamento: {stato.scostamento_globale:+}%"
    )
    ws["A3"].font = _BOLD
    ws["A4"] = f"Importo maturato (somma ultimi SAL): € {stato.importo_maturato_totale:,.2f}"

    intestazioni = [
        "Lotto", "Impresa", "Categoria", "Importo €",
        "Avanz. reale %", "Avanz. atteso %", "Scostamento %", "Giudizio",
        "Ultimo SAL €",
    ]
    riga0 = 6
    for col, testo in enumerate(intestazioni, start=1):
        cell = ws.cell(row=riga0, column=col, value=testo)
        cell.fill = _INTESTAZIONE
        cell.font = _BIANCO_BOLD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = _BORDO

    for i, s in enumerate(stato.stati_lotti, start=1):
        r = riga0 + i
        valori = [
            s.lotto.nome, s.lotto.impresa, s.lotto.categoria,
            s.lotto.importo_contratto, s.avanzamento_reale,
            s.avanzamento_atteso, s.scostamento, s.giudizio,
            s.importo_maturato,
        ]
        for col, val in enumerate(valori, start=1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.border = _BORDO
        # colora la cella del giudizio in base al ritardo
        ws.cell(row=r, column=8).fill = _colore_per_scostamento(s.scostamento)

    for col in range(1, len(intestazioni) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16


def _foglio_gantt(wb: Workbook, stato: StatoCantiere) -> None:
    ws = wb.create_sheet(title=_nome_foglio(stato.cantiere.nome, "Gantt"))
    lotti = [s.lotto for s in stato.stati_lotti]
    date_inizio = [l.data_inizio for l in lotti if l.data_inizio]
    date_fine = [l.data_fine_prevista for l in lotti if l.data_fine_prevista]
    if not date_inizio or not date_fine:
        ws["A1"] = "Date non sufficienti per generare il Gantt."
        return

    inizio_min = min(date_inizio)
    fine_max = max(date_fine)
    # scala settimanale per non avere troppe colonne
    giorni_totali = (fine_max - inizio_min).days or 1
    passo = max(1, giorni_totali // 40)  # ~40 colonne max

    col_base = 3  # le prime 2 colonne sono Lotto/Impresa
    ws.cell(row=1, column=1, value="Lotto").font = _BIANCO_BOLD
    ws.cell(row=1, column=2, value="Impresa").font = _BIANCO_BOLD
    ws.cell(row=1, column=1).fill = _INTESTAZIONE
    ws.cell(row=1, column=2).fill = _INTESTAZIONE

    # intestazione temporale
    n_colonne = giorni_totali // passo + 1
    for k in range(n_colonne):
        giorno = inizio_min + timedelta(days=k * passo)
        cell = ws.cell(row=1, column=col_base + k, value=giorno.strftime("%d/%m"))
        cell.fill = _INTESTAZIONE
        cell.font = _BIANCO_BOLD
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col_base + k)].width = 4

    oggi = date.today()
    for i, s in enumerate(stato.stati_lotti, start=1):
        r = i + 1
        l = s.lotto
        ws.cell(row=r, column=1, value=l.nome).border = _BORDO
        ws.cell(row=r, column=2, value=l.impresa).border = _BORDO
        if not l.data_inizio or not l.data_fine_prevista:
            continue
        colore = _colore_per_scostamento(s.scostamento)
        for k in range(n_colonne):
            giorno = inizio_min + timedelta(days=k * passo)
            cell = ws.cell(row=r, column=col_base + k)
            if l.data_inizio <= giorno <= l.data_fine_prevista:
                cell.fill = colore
                cell.border = _BORDO

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 20

    # legenda
    rl = len(stato.stati_lotti) + 3
    ws.cell(row=rl, column=1, value="Legenda:").font = _BOLD
    for j, (txt, fill) in enumerate(
        [("In linea", _VERDE), ("Ritardo lieve", _GIALLO), ("Ritardo grave", _ROSSO)]
    ):
        c = ws.cell(row=rl + 1 + j, column=1, value=txt)
        ws.cell(row=rl + 1 + j, column=2).fill = fill


def _nome_foglio(nome_cantiere: str, suffisso: str) -> str:
    """Excel limita i nomi dei fogli a 31 caratteri e vieta alcuni simboli."""
    pulito = "".join(ch for ch in nome_cantiere if ch not in r'[]:*?/\\')
    base = f"{pulito[:20]} {suffisso}"
    return base[:31]


def esporta(archivio: Archivio, percorso: str | Path = "export/cronoprogramma.xlsx",
            alla_data: date | None = None) -> Path:
    """Genera il file Excel con un foglio Riepilogo + Gantt per ogni cantiere."""
    wb = Workbook()
    wb.remove(wb.active)  # rimuove il foglio vuoto di default

    if not archivio.cantieri:
        ws = wb.create_sheet("Vuoto")
        ws["A1"] = "Nessun cantiere presente."
    else:
        for cantiere in archivio.cantieri:
            stato = analizza_cantiere(cantiere, alla_data)
            _foglio_riepilogo(wb, stato)
            _foglio_gantt(wb, stato)

    percorso = Path(percorso)
    percorso.parent.mkdir(parents=True, exist_ok=True)
    wb.save(percorso)
    return percorso

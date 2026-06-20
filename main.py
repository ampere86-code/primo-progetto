"""Cronocantieri - programma a menu per gestire i cronoprogrammi dei cantieri.

Avvio:
    python main.py

I dati vengono salvati in  dati/cantieri.json  e l'export Excel viene
generato in  export/cronoprogramma.xlsx
"""

from __future__ import annotations

from datetime import date

from cronocantieri.analisi import analizza_cantiere
from cronocantieri.esporta_excel import esporta, genera_template_impresa
from cronocantieri.models import SAL, Archivio, Cantiere, Lotto
from cronocantieri.unisci import unisci_cronoprogrammi


# ---------------------------------------------------------------- input utili

def chiedi(testo: str, obbligatorio: bool = True) -> str:
    while True:
        valore = input(testo).strip()
        if valore or not obbligatorio:
            return valore
        print("  ! Campo obbligatorio, riprova.")


def chiedi_float(testo: str, default: float = 0.0) -> float:
    while True:
        valore = input(testo).strip().replace(",", ".")
        if not valore:
            return default
        try:
            return float(valore)
        except ValueError:
            print("  ! Inserisci un numero (es. 12500.50).")


def chiedi_data(testo: str, obbligatorio: bool = False) -> date | None:
    while True:
        valore = input(testo).strip()
        if not valore:
            if obbligatorio:
                print("  ! Data obbligatoria.")
                continue
            return None
        try:
            return date.fromisoformat(valore)
        except ValueError:
            print("  ! Formato data: AAAA-MM-GG (es. 2026-03-15).")


# ----------------------------------------------------------------- operazioni

def nuovo_cantiere(archivio: Archivio) -> None:
    print("\n--- NUOVO CANTIERE ---")
    nome = chiedi("Nome cantiere: ")
    if archivio.trova(nome):
        print("  ! Esiste già un cantiere con questo nome.")
        return
    tipo = chiedi("Tipo (pubblico/privato) [pubblico]: ", obbligatorio=False) or "pubblico"
    cantiere = Cantiere(
        nome=nome,
        tipo=tipo,
        committente=chiedi("Committente: ", obbligatorio=False),
        descrizione=chiedi("Descrizione: ", obbligatorio=False),
        data_inizio=chiedi_data("Data inizio (AAAA-MM-GG): "),
        data_fine_prevista=chiedi_data("Data fine prevista (AAAA-MM-GG): "),
    )
    try:
        archivio.aggiungi_cantiere(cantiere)
    except ValueError as e:
        print(f"  ! {e}")
        return
    archivio.salva()
    print(f"  Cantiere '{nome}' creato.")


def seleziona_cantiere(archivio: Archivio) -> Cantiere | None:
    if not archivio.cantieri:
        print("  Nessun cantiere presente.")
        return None
    print("\nCantieri disponibili:")
    for i, c in enumerate(archivio.cantieri, 1):
        print(f"  {i}. {c.nome}  ({len(c.lotti)} lotti)")
    scelta = chiedi("Numero cantiere: ", obbligatorio=False)
    if not scelta.isdigit() or not (1 <= int(scelta) <= len(archivio.cantieri)):
        print("  ! Scelta non valida.")
        return None
    return archivio.cantieri[int(scelta) - 1]


def nuovo_lotto(archivio: Archivio) -> None:
    cantiere = seleziona_cantiere(archivio)
    if not cantiere:
        return
    print(f"\n--- NUOVO LOTTO per '{cantiere.nome}' ---")
    lotto = Lotto(
        nome=chiedi("Nome lotto (es. 'Lotto 1 - Opere edili'): "),
        impresa=chiedi("Impresa esecutrice: "),
        categoria=chiedi("Categoria (edile/impiantistico/...): ", obbligatorio=False),
        responsabile=chiedi("Responsabile: ", obbligatorio=False),
        data_inizio=chiedi_data("Data inizio (AAAA-MM-GG): "),
        data_fine_prevista=chiedi_data("Data fine prevista (AAAA-MM-GG): "),
        importo_contratto=chiedi_float("Importo contratto € [0]: "),
        avanzamento_pct=chiedi_float("Avanzamento attuale % [0]: "),
        note=chiedi("Note: ", obbligatorio=False),
    )
    cantiere.lotti.append(lotto)
    archivio.salva()
    print(f"  Lotto '{lotto.nome}' aggiunto.")


def aggiorna_avanzamento(archivio: Archivio) -> None:
    cantiere = seleziona_cantiere(archivio)
    if not cantiere or not cantiere.lotti:
        print("  Nessun lotto da aggiornare.")
        return
    print("\nLotti:")
    for i, l in enumerate(cantiere.lotti, 1):
        print(f"  {i}. {l.nome} - {l.impresa}  ({l.avanzamento_pct}%)")
    scelta = chiedi("Numero lotto: ", obbligatorio=False)
    if not scelta.isdigit() or not (1 <= int(scelta) <= len(cantiere.lotti)):
        print("  ! Scelta non valida.")
        return
    lotto = cantiere.lotti[int(scelta) - 1]
    lotto.avanzamento_pct = chiedi_float(
        f"Nuovo avanzamento % [{lotto.avanzamento_pct}]: ", lotto.avanzamento_pct
    )
    archivio.salva()
    print("  Avanzamento aggiornato.")


def registra_sal(archivio: Archivio) -> None:
    cantiere = seleziona_cantiere(archivio)
    if not cantiere or not cantiere.lotti:
        print("  Nessun lotto disponibile.")
        return
    print("\nLotti:")
    for i, l in enumerate(cantiere.lotti, 1):
        print(f"  {i}. {l.nome} - {l.impresa}")
    scelta = chiedi("Numero lotto: ", obbligatorio=False)
    if not scelta.isdigit() or not (1 <= int(scelta) <= len(cantiere.lotti)):
        print("  ! Scelta non valida.")
        return
    lotto = cantiere.lotti[int(scelta) - 1]
    numero = (lotto.ultimo_sal().numero + 1) if lotto.ultimo_sal() else 1
    print(f"\n--- SAL n.{numero} per '{lotto.nome}' ---")
    sal = SAL(
        numero=numero,
        data=chiedi_data("Data SAL (AAAA-MM-GG): ", obbligatorio=True),
        avanzamento_pct=chiedi_float("Avanzamento dichiarato %: "),
        importo=chiedi_float("Importo maturato € : "),
        note=chiedi("Note: ", obbligatorio=False),
    )
    lotto.sal.append(sal)
    # aggiorna anche l'avanzamento corrente del lotto al valore del SAL
    lotto.avanzamento_pct = sal.avanzamento_pct
    archivio.salva()
    print(f"  SAL n.{numero} registrato (€ {sal.importo:,.2f}).")


def mostra_analisi(archivio: Archivio) -> None:
    cantiere = seleziona_cantiere(archivio)
    if not cantiere:
        return
    stato = analizza_cantiere(cantiere)
    print(f"\n===== ANALISI: {cantiere.nome} =====")
    print(
        f"Avanzamento globale: {stato.avanzamento_globale_reale}% "
        f"(atteso {stato.avanzamento_globale_atteso}%)  "
        f"scostamento {stato.scostamento_globale:+}%"
    )
    print(f"Importo maturato totale: € {stato.importo_maturato_totale:,.2f}")
    print("-" * 60)
    for s in stato.stati_lotti:
        print(
            f"  {s.lotto.impresa:<22} {s.lotto.nome[:28]:<28} "
            f"reale {s.avanzamento_reale:>5}% / atteso {s.avanzamento_atteso:>5}%  "
            f"[{s.giudizio}]"
        )
    peggiore = stato.lotto_piu_in_ritardo()
    if peggiore and peggiore.in_ritardo:
        print("-" * 60)
        print(
            f"  >> Impresa più in ritardo: {peggiore.lotto.impresa} "
            f"({peggiore.scostamento:+}% rispetto al previsto)"
        )


def importa_da_imprese(archivio: Archivio) -> None:
    print("\n--- IMPORTA CRONOPROGRAMMI IMPRESE (Excel/PDF) ---")
    print("Inserisci i percorsi dei file, uno per riga. Riga vuota per terminare.")
    print("(Puoi trascinare il file nella finestra per incollarne il percorso.)")
    percorsi: list[str] = []
    while True:
        p = input(f"  File {len(percorsi) + 1}: ").strip().strip('"')
        if not p:
            break
        percorsi.append(p)
    if not percorsi:
        print("  Nessun file indicato.")
        return
    nome = chiedi("Nome cantiere [vuoto = leggi dai file]: ", obbligatorio=False)
    committente = chiedi("Committente: ", obbligatorio=False)
    try:
        cantiere, importati = unisci_cronoprogrammi(
            percorsi, cantiere_nome=nome or None, committente=committente
        )
    except Exception as e:  # errori di lettura/formato
        print(f"  ! Errore durante l'import: {e}")
        return

    if archivio.trova(cantiere.nome):
        print(f"  ! Esiste già un cantiere '{cantiere.nome}'. Rinominalo e riprova.")
        return
    try:
        archivio.aggiungi_cantiere(cantiere)
    except ValueError as e:
        print(f"  ! {e}")
        return
    archivio.salva()
    print(f"\n  Cantiere '{cantiere.nome}' creato unendo {len(importati)} file:")
    for imp in importati:
        print(f"    - {imp.file_origine}: {imp.lotto.impresa} "
              f"({imp.righe_lette} attività)")
    print("  Usa la voce 5) per vedere l'analisi con gli avanzamenti calcolati.")


def genera_modello(archivio: Archivio) -> None:
    percorso = genera_template_impresa()
    print(f"  Modello per le imprese generato: {percorso.resolve()}")
    print("  Invialo alle imprese: una volta compilato, importalo con la voce 8).")


def esporta_excel(archivio: Archivio) -> None:
    if not archivio.cantieri:
        print("  Nessun cantiere da esportare.")
        return
    percorso = esporta(archivio)
    print(f"  File Excel generato: {percorso.resolve()}")


# ---------------------------------------------------------------------- menu

MENU = """
============================================
   CRONOCANTIERI - gestione cronoprogrammi
============================================
  1) Nuovo cantiere
  2) Aggiungi lotto/impresa a un cantiere
  3) Aggiorna avanzamento di un lotto
  4) Registra un SAL
  5) Analisi cantiere (ritardi)
  6) Esporta in Excel (con Gantt)
  7) Elenco cantieri
  8) Importa cronoprogrammi imprese (Excel/PDF) e crea cantiere
  9) Genera modello Excel da inviare alle imprese
  0) Esci
"""


def elenco(archivio: Archivio) -> None:
    if not archivio.cantieri:
        print("  Nessun cantiere presente.")
        return
    for c in archivio.cantieri:
        print(f"\n  {c.nome}  [{c.tipo}]  - {len(c.lotti)} lotti")
        for l in c.lotti:
            print(f"     - {l.nome} ({l.impresa}) {l.avanzamento_pct}%")


def main() -> None:
    archivio = Archivio().carica()
    azioni = {
        "1": nuovo_cantiere,
        "2": nuovo_lotto,
        "3": aggiorna_avanzamento,
        "4": registra_sal,
        "5": mostra_analisi,
        "6": esporta_excel,
        "7": elenco,
        "8": importa_da_imprese,
        "9": genera_modello,
    }
    while True:
        print(MENU)
        scelta = input("Scelta: ").strip()
        if scelta == "0":
            print("Arrivederci!")
            break
        azione = azioni.get(scelta)
        if azione:
            azione(archivio)
        else:
            print("  ! Scelta non valida.")


if __name__ == "__main__":
    main()

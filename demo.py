"""Crea dati di esempio per provare subito il programma.

Avvio:
    python demo.py

Genera un cantiere con 3 imprese su 3 lotti in parallelo, una delle
quali in ritardo, salva i dati e produce l'Excel con il Gantt.
"""

from datetime import date

from cronocantieri.analisi import analizza_cantiere
from cronocantieri.esporta_excel import esporta
from cronocantieri.models import SAL, Archivio, Cantiere, Lotto


def main() -> None:
    cantiere = Cantiere(
        nome="Scuola Media Via Roma",
        tipo="pubblico",
        committente="Comune di Esempio",
        descrizione="Ristrutturazione e adeguamento impianti",
        data_inizio=date(2026, 1, 7),
        data_fine_prevista=date(2026, 9, 30),
        lotti=[
            Lotto(
                nome="Lotto 1 - Opere edili",
                impresa="Rossi Costruzioni S.r.l.",
                categoria="edile",
                responsabile="Geom. Rossi",
                data_inizio=date(2026, 1, 7),
                data_fine_prevista=date(2026, 7, 31),
                importo_contratto=850000,
                avanzamento_pct=58,
                note="In linea con il cronoprogramma.",
                sal=[
                    SAL(1, date(2026, 3, 31), 30, 255000, "Primo SAL regolare"),
                    SAL(2, date(2026, 5, 31), 58, 493000, "Secondo SAL"),
                ],
            ),
            Lotto(
                nome="Lotto 2 - Impianti elettrici e meccanici",
                impresa="Bianchi Impianti S.p.A.",
                categoria="impiantistico",
                responsabile="P.I. Bianchi",
                data_inizio=date(2026, 3, 1),
                data_fine_prevista=date(2026, 9, 15),
                importo_contratto=620000,
                avanzamento_pct=20,  # in ritardo
                note="Ritardo per forniture quadri elettrici.",
                sal=[
                    SAL(1, date(2026, 5, 31), 20, 124000, "Avvio rallentato"),
                ],
            ),
            Lotto(
                nome="Lotto 3 - Serramenti e facciate",
                impresa="Verde Infissi S.r.l.",
                categoria="serramenti",
                responsabile="Sig. Verde",
                data_inizio=date(2026, 4, 1),
                data_fine_prevista=date(2026, 8, 31),
                importo_contratto=300000,
                avanzamento_pct=45,
                note="In anticipo sulla tabella di marcia.",
                sal=[
                    SAL(1, date(2026, 5, 31), 45, 135000, "Posa avviata in anticipo"),
                ],
            ),
        ],
    )

    archivio = Archivio()
    archivio.aggiungi_cantiere(cantiere)
    archivio.salva()
    print(f"Dati salvati in: {archivio.percorso.resolve()}")

    # analisi a video alla data del 1 giugno 2026
    stato = analizza_cantiere(cantiere, date(2026, 6, 1))
    print(f"\nAnalisi al 01/06/2026 - {cantiere.nome}")
    print(
        f"Avanzamento globale: {stato.avanzamento_globale_reale}% "
        f"(atteso {stato.avanzamento_globale_atteso}%)  "
        f"scostamento {stato.scostamento_globale:+}%"
    )
    for s in stato.stati_lotti:
        print(
            f"  {s.lotto.impresa:<26} reale {s.avanzamento_reale:>5}% / "
            f"atteso {s.avanzamento_atteso:>5}%  [{s.giudizio}]"
        )
    peggiore = stato.lotto_piu_in_ritardo()
    if peggiore and peggiore.in_ritardo:
        print(f"\n>> Impresa piu in ritardo: {peggiore.lotto.impresa} "
              f"({peggiore.scostamento:+}%)")

    percorso = esporta(archivio, alla_data=date(2026, 6, 1))
    print(f"\nExcel generato in: {percorso.resolve()}")


if __name__ == "__main__":
    main()

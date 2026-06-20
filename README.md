# Cronocantieri

Programma in Python per creare e gestire i **cronoprogrammi dei cantieri**
edili e impiantistici, pubblici e privati.

Pensato per appalti in cui **più imprese lavorano in parallelo su lotti
diversi** della stessa commessa: il programma analizza l'andamento della
singola impresa e quello globale del cantiere, evidenziando **quale impresa
è in ritardo** rispetto alle altre.

## Funzionalità

- Gestione fino a **20 cantieri**, ciascuno con più lotti/imprese.
- **Import dei cronoprogrammi delle imprese da Excel o PDF**: il project
  manager carica gli n file inviati dalle imprese e il sistema genera in
  automatico il **cronoprogramma totale** del cantiere, organizzato per
  lotto e impresa.
- **Calcolo automatico delle percentuali** di avanzamento (di lotto e
  globale) a partire dalle attività importate: usa la % dichiarata nel file
  dove presente, altrimenti la stima dalle date.
- Avanzamento espresso in **percentuale**.
- Registrazione dei **SAL** (Stati di Avanzamento Lavori) con data e importo.
- **Note** su cantieri, lotti, attività e SAL.
- Rilevamento automatico dei **ritardi**: confronto fra avanzamento reale e
  avanzamento atteso in base al tempo trascorso.
- Avanzamento globale del cantiere **pesato sugli importi** dei lotti.
- Esportazione in **Excel** con **diagramma di Gantt** colorato per ritardo.
- Dati salvati in **JSON**: portabili su qualsiasi PC.

## Requisiti

- Python 3.10 o superiore
- Libreria `openpyxl` (per l'export Excel)

Installazione delle dipendenze:

```bash
pip install -r requirements.txt
```

## Uso

Avvia il menu interattivo:

```bash
python main.py
```

Voci del menu:

1. Nuovo cantiere
2. Aggiungi lotto/impresa a un cantiere
3. Aggiorna avanzamento di un lotto
4. Registra un SAL
5. Analisi cantiere (ritardi)
6. Esporta in Excel (con Gantt)
7. Elenco cantieri
8. Importa cronoprogrammi imprese (Excel/PDF) e crea cantiere
9. Genera modello Excel da inviare alle imprese

### Flusso tipico del project manager

1. Voce **9**: genera il **modello Excel precompilato** (riferimenti cantiere
   e attività standard già inseriti) e invialo a ogni impresa.
2. Ogni impresa compila date/importi/% per le attività e indica il
   **Lotto / Commessa**, poi restituisce il file in Excel o PDF.
3. Voce **8**: carica gli n file ricevuti. Il sistema li unisce nel
   cronoprogramma totale del cantiere e calcola le percentuali.
4. Voce **5**: analizza ritardi e avanzamento globale.
5. Voce **6**: esporta l'Excel per la stazione appaltante. Contiene tre fogli:
   - **Riepilogo**: andamento per lotto/impresa;
   - **Cronoprogramma generale**: tutti i lotti con il dettaglio delle
     attività e il Gantt mensile colorato per ritardo;
   - **Gantt**: vista sintetica per lotto.

### Provare subito con dati di esempio

```bash
python demo.py          # un cantiere creato a mano, con 3 imprese
python demo_import.py   # genera 3 file Excel di imprese, li importa e li unisce
```

Entrambi mostrano l'analisi a video e generano `export/cronoprogramma.xlsx`.

### Formato dei file delle imprese

Il file (Excel o PDF) deve contenere, nelle prime righe, alcune etichette
di intestazione e poi una tabella delle attività:

```
Cantiere:     Scuola Media Via Roma
Impresa:      Rossi Costruzioni S.r.l.
Lotto:        Lotto 1 - Opere edili
Categoria:    edile

Attività            | Data inizio | Data fine  | Importo | Avanzamento %
Scavi e fondazioni  | 2026-01-07  | 2026-02-28 | 150000  | 100
Struttura in c.a.   | 2026-03-01  | 2026-05-31 | 500000  | 60
```

Il lettore riconosce le colonne dalle parole chiave nell'intestazione,
quindi tollera piccole variazioni di nome e posizione. La colonna
**Avanzamento %** può essere lasciata vuota: in tal caso la percentuale
viene stimata dalle date.

## Struttura del progetto

```
main.py                      menu interattivo
demo.py                      dati di esempio (cantiere creato a mano)
demo_import.py               dati di esempio (import + fusione file imprese)
requirements.txt             dipendenze
cronocantieri/
  models.py                  modello dati (Cantiere, Lotto, Attività, SAL) + JSON
  analisi.py                 calcolo avanzamenti e ritardi
  importa.py                 lettura cronoprogrammi imprese da Excel/PDF
  unisci.py                  fusione di più file nel cronoprogramma totale
  esporta_excel.py           export Excel con Gantt + modello per le imprese
dati/                        dati salvati (generati, non versionati)
export/                      file Excel generati (non versionati)
import_imprese/              file di esempio delle imprese (non versionati)
```

## Come funziona il calcolo del ritardo

Per ogni lotto, conoscendo data di inizio e fine prevista, si calcola
l'avanzamento *atteso* alla data odierna ipotizzando un progresso lineare:

```
avanzamento_atteso = tempo_trascorso / durata_totale * 100
scostamento        = avanzamento_reale - avanzamento_atteso
```

Uno scostamento negativo indica un'impresa **in ritardo**. L'avanzamento
globale del cantiere è la media degli avanzamenti dei lotti, pesata sui
rispettivi importi di contratto.

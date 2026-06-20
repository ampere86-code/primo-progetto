# Cronocantieri

Programma in Python per creare e gestire i **cronoprogrammi dei cantieri**
edili e impiantistici, pubblici e privati.

Pensato per appalti in cui **più imprese lavorano in parallelo su lotti
diversi** della stessa commessa: il programma analizza l'andamento della
singola impresa e quello globale del cantiere, evidenziando **quale impresa
è in ritardo** rispetto alle altre.

## Funzionalità

- Gestione fino a **20 cantieri**, ciascuno con più lotti/imprese.
- Avanzamento espresso in **percentuale**.
- Registrazione dei **SAL** (Stati di Avanzamento Lavori) con data e importo.
- **Note** su cantieri, lotti e SAL.
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

### Provare subito con dati di esempio

```bash
python demo.py
```

Crea un cantiere con 3 imprese su 3 lotti (di cui una in ritardo), mostra
l'analisi a video e genera il file Excel in `export/cronoprogramma.xlsx`.

## Struttura del progetto

```
main.py                      menu interattivo
demo.py                      dati di esempio
requirements.txt             dipendenze
cronocantieri/
  models.py                  modello dati (Cantiere, Lotto, SAL) + salvataggio JSON
  analisi.py                 calcolo avanzamenti e ritardi
  esporta_excel.py           export Excel con Gantt
dati/                        dati salvati (generati, non versionati)
export/                      file Excel generati (non versionati)
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

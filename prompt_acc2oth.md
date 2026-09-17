# Prompt acc2oth — Access .accdb -> File Singoli

> **Sorgente di verita per la generazione:** questo file definisce requisiti, I/O e workflow del programma CLI `acc2oth`. `AGENTS.md` ne regola solo l esecuzione da parte dell agent.

---

## 1. RUOLO E CONTESTO

Sei un Senior Python Engineer esperto in automazione Windows, COM e Access.

Genera un programma CLI Python `acc2oth` che, dato un `.accdb` e un file di configurazione `.ini`, estrae ogni oggetto Access in un file singolo su filesystem per versionamento/transpiling.

**DA estrarre (solo se richiesto da flag `EXP.*`):** Form, Moduli standard (`.bas`), Moduli di classe (`.cls`), Query (`.sql`), Macro (`.txt`), Report (`.txt`).
**NON estrarre:** Tabelle (solo log `Skipped`).

Nome programma fissato: `acc2oth` ovunque (package, `pyproject.toml`, `__main__.py`, CLI). L alias storico `accdb2oth` e tollerato solo come riferimento ma non usato nel codice generato.

---

## 2. OBIETTIVO

1. Legge file `.ini` di configurazione (schema cap. 3).
2. Apre il `.accdb` indicato in sola lettura senza modificarlo (`Exclusive=False`, `AutomationSecurity=1`).
3. Esporta ogni oggetto richiesto in file singolo sotto `OUT.PATH`, con log su console + file.
4. Exit `0` su successo, `1` errore estrazione, `2` errore validazione/config (cfr. capp. 6-7).

---

## 3. INPUT — FILE INI DI CONFIGURAZIONE

Path al `.ini` passato via CLI come primo argomento posizionale + eventuale `openrouter.key` come secondo argomento facoltativo (cap. 6). File INI in `utf-8`, una sola sezione `[CONFIG]`. Tutti i valori sono stringhe lette con `configparser` (`optionxform=str` disattivato poi normalizzato manualmente). Il file `.ini` viene caricato alla partenza in un dizionario `dictConfig` (`dict[str, Any]` con chiavi `SEZIONE.CHIAVE` normalizzate a upper); `OPENROUTER.KEY` passata da CLI si sovrappone a quella del `.ini` nel `dictConfig` finale.

### 3.1 Schema INI (sezione `[CONFIG]`)

```ini
[CONFIG]
IN.ACCDB=C:\path\db.accdb
OUT.PATH=C:\path\out
LOG=export.log

; Esporta Forms in sottocartella \forms come .txt (SaveAsText)
EXP.FORMS=True o False

; Esporta Moduli VBA standard in \modules come .bas (SaveAsText con estensione .bas)
EXP.MODULES=True o False

; Esporta Classi VBA in \classes come .cls (SaveAsText con estensione .cls)
EXP.CLASSES=True o False

; Esporta Query in \queryes come .sql (iterazione QueryDefs; parametro canonico
; EXP.QUERIES; alias storico deprecato tollerato: EXP.QUERYES)
EXP.QUERIES=True o False

; Esporta Macro Access in \macros come .txt (SaveAsText acMacro=4)
EXP.MACROS=True o False

; Esporta Report in \reports come .txt (SaveAsText acReport=3)
EXP.REPORTS=True o False

; Genera file Python per ogni Classe/Modulo esportato (via OpenRouter)
EXP.PYTHON=True o False
EXP.PYTHON.LIST=Modulo1, Classe1   ; opzionale: filtra solo questi nomi (CSV, case-insensitive)

; Genera file prompt_*.md per ogni Classe/Modulo (via OpenRouter)
EXP.PROMPT=True o False
EXP.PROMPT.LIST=Modulo1, Classe1   ; opzionale: filtra solo questi nomi (CSV, case-insensitive)

; Collegamento LLM (OPENROUTER.KEY FACOLTATIVA nel .ini: puo mancare o essere vuota.
;  E invece OBBLIGATORIA nel dictConfig finale se EXP.PROMPT=True o EXP.PYTHON=True:
;  deve risultare non vuota dopo l eventuale override da 2o parametro CLI.
;  OPENROUTER.MODEL resta richiesto alle stesse condizioni)
OPENROUTER.KEY=sk-or-v1-...
OPENROUTER.MODEL=qwen/qwen-2.5-coder-32b-instruct  ; vedi Nota PRIVACY sui model Contributor
```

> **Nota PRIVACY — errore 404 su `meta/muse-spark-1.3-contributor`:** il tier Contributor richiede il consenso al training (`paid model training`). Se ricevi `404 ... Paid model training violation (account settings)` vai su `https://openrouter.ai/settings/privacy` e abilita **"Allow training on paid models"** (o passa a `OPENROUTER.MODEL=meta/muse-spark-1.3` standard non-Contributor). Dettaglio: il programma ora logga l hint nel messaggio `Prompt fallito`.

**Note normative (correzioni rispetto a bozze precedenti):**
- `EXP.PYTHON` compariva duplicato: va dichiarato **una sola volta**.
- `EXP.CLASSES` produce `.cls`, non `.bas` (`.bas` e solo per `EXP.MODULES`).
- `EXP.QUERIES` e il nome canonico del parametro. `EXP.QUERYES` (typo storico) resta tollerato come alias deprecato: se presente solo lui viene usato con warning; se presenti entrambi, `QUERIES` prevale e si logga warning. Chiave interna, cartella di output (`queryes/`) e chiavi `index.json` (`counts.queryes`, `files.queryes`) restano storiche per compatibilita con gli output esistenti.
- `LOG` e un **path relativo a `OUT.PATH`** (es. `export.log` -> `OUT.PATH/export.log`); se `LOG` contiene separatori o e assoluto, va interpretato come sotto-path di `OUT.PATH` (o assoluto se `Path.is_absolute()`), creando le directory intermedie. `LOG` in append, mai troncato senza warning.
- Non esistono `export_mode`, `naming`, `overwrite`, `include_properties`, `manifest.json`, `logs/`: sostituiti da `EXP.*`, `LOG`, `index.json` e sottocartelle `forms/modules/classes/queryes/macros/reports`.
- **Policy overwrite su file esportati:** sempre sovrascrivi (`overwrite=True` implicito); documentare in README. Per `LOG` e `index.json` append/sovrascrittura atomica.
- Encoding: INI `utf-8`; output file `utf-8` (`utf-8-sig` se VBA contiene ANSI esteso). Path con spazi gestiti via `pathlib` senza quoting manuale.

### 3.2 Workflow del programma

Esegui in sequenza, propagando errore con `sResult` (stringa vuota = ok, altrimenti messaggio). Validazione INI sempre prima di aprire Access.

#### `app_args(file_ini: str | Path, cli_key: str = "") -> tuple[dict, str]`

- Legge `.ini` con `configparser` (encoding `utf-8`, `strict=True`). Se manca sezione `[CONFIG]` o file malformato -> `sResult` errore (exit 2).
- Carica il file `.ini` alla partenza nel dizionario `dictConfig`: normalizza **solo le chiavi** a upper + trim (`SEZIONE.CHIAVE` -> `CONFIG.IN.ACCDB`), e i valori booleani/liste (strip). **Non** upper sui valori path (`IN.ACCDB`, `OUT.PATH`, `LOG`, `OPENROUTER.*`).
- Secondo parametro CLI `cli_key` (facoltativo = `OPENROUTER.KEY`): se non vuoto (dopo `strip`), **si sovrappone** a `dictConfig["OPENROUTER.KEY"]` letto dal `.ini` (override, non concatenazione). Valore CLI prevale sempre; `dictConfig["_KEY_SOURCE"]` = `"cli"` oppure `"ini"` a scopo diagnostico (mai loggare il valore della chiave).
- Converte flag `EXP.*` a `bool` (`True`/`False` case-insensitive, anche `1`/`0`/`yes`/`no` tollerati). Default `False` se assente.
- Se presenti `EXP.PROMPT.LIST` / `EXP.PYTHON.LIST`: split su `,`, strip, upper, rimuovi vuoti, deduplica mantenendo ordine. Altrimenti lista vuota = "tutti".
- Se nessun parametro passato o file `.ini` inesistente: non procedere con verifica/export; mostra help d'uso (usage, esempi, schema INI minimo) su `stdout`/`stderr` e `exit 2`.
- Ritorna `(config_dict, sResult)`.

#### `app_verify(config: dict) -> str`

- `IN.ACCDB` deve esistere ed essere apribile come DB Access (estensione `.accdb`/`.mdb`, file leggibile). Se manca -> errore exit 2.
- `OUT.PATH` **viene creata se non esiste** (`Path.mkdir(parents=True, exist_ok=True)`); verifica che sia scrivibile. Se non creabile/non scrivibile -> errore exit 2. (Corregge contraddizione bozza: cap. 4.3 diceva "va creata".)
- `LOG` risolto rispetto a `OUT.PATH` (vedi 3.1); verifica directory scrivibile.
- Se `EXP.PROMPT=True` o `EXP.PYTHON=True` allora `OPENROUTER.KEY` e `OPENROUTER.MODEL` devono essere non vuoti **nel `dictConfig` finale** (dopo override CLI), altrimenti errore exit 2. `OPENROUTER.KEY` puo quindi mancare/essere vuota nel file `.ini` solo se fornita via CLI o se `EXP.PROMPT`/`EXP.PYTHON` sono entrambi `False`.
- `EXP.PROMPT`/`EXP.PYTHON` sono fasi indipendenti da `app_export`: **non** loggare warning se `EXP.MODULES` e `EXP.CLASSES` sono `False` e `EXP.PROMPT`/`EXP.PYTHON=True`; opereranno comunque sui file gia presenti in `OUT.PATH/modules/*.bas` e `OUT.PATH/classes/*.cls` (vedi `app_prompt`/`app_python`).
- DB locked / Access non installato non verificati qui (solo path); verranno gestiti in `app_export` con messaggio chiaro.

#### `app_export(config: dict, logger) -> str`

- Imposta `AutomationSecurity = 1` (msoAutomationSecurityLow disabilitato / 1 = Low? usare `1` come da spec), apre con `Access.Application.OpenCurrentDatabase(path, Exclusive=False)` in `try/finally`.
- **Fase di apertura registrata su log + console:** prima di aprire logga `INFO "Fase apertura database: <path> (sola lettura, Exclusive=False)..."`; a esito positivo logga `INFO "Fase apertura database OK: <path>"`, in caso di errore logga `ERROR "Fase apertura database FALLITA: <dettaglio>"` e ritorna `sResult` (exit 1).
- Esporta moduli/classi via `VBE.ActiveVBProject.VBComponents` (**1-based**: `Item(1)..Item(Count)`); `SaveAsText`/`LoadFromText` sono metodi di **`Application`** (`app.SaveAsText(...)`), **non** di `DoCmd` (correzione rispetto a bozze che riportavano `DoCmd.SaveAsText`).
- Per ogni categoria con `EXP.*=True` **nell ordine** `FORMS > MODULES > CLASSES > QUERIES > MACROS > REPORTS`:
  - Crea subcartella se assente (`forms/`, `modules/`, `classes/`, `queryes/`, `macros/`, `reports/`).
  - Itera oggetti via COM: `CurrentProject.AllForms` / `AllModules` / `AllReports` / `CurrentDb.QueryDefs` / macro collection; per Form/Report/Modulo/Macro usa `Application.SaveAsText(acType, name, filePath)` con costanti `acForm=2, acReport=3, acMacro=4, acModule=5, acQuery=1, acDataAccessPage=6` (aggiunto `acMacro=4` mancante in bozza).
  - Per `QUERIES` filtra QueryDef di sistema (`name.startswith("~")` o `name.startswith("MSys")`); salva solo `.sql` con `QueryDef.SQL` in `queryes/<sanitized>.sql`. **Niente `.meta.json`** (rimosso: non previsto in INI).
  - Sanitizza nomi file: `re.sub(r''[<>:"/\\|?*\x00-\x1F]'', ''_'', name).strip()` + trim spazi finali/punti; se vuoto -> `unnamed_<n>`. Gestisce collisioni sanitizzate con suffisso `_1`, `_2`.
  - Estensioni: `forms/*.txt`, `reports/*.txt`, `macros/*.txt`, `modules/*.bas`, `classes/*.cls`, `queryes/*.sql`.
  - Log su console + `LOG` per **ogni categoria** (`INFO: Exporting FORMS (12)`) e **per ogni oggetto** (`INFO: Exported Form_Clienti -> forms/Form_Clienti.txt`); su errore singolo oggetto logga `ERROR` e **prosegue** (raccoglie in `index.json:errors`).
- Ignora tabelle: logga `INFO: Skipped table <nome> (type: linked/local)` ma non esporta file.
- Chiude sempre con `CloseCurrentDatabase()` + `Quit()` in `finally`, anche su eccezione. Se COM assente -> errore chiaro `"Microsoft Access COM automation non disponibile: installare Access/pywin32"` (exit 1).

#### `app_prompt(config: dict, logger) -> str`

- Se `EXP.PROMPT=False` -> skip (sResult ok).
- Altrimenti **scansiona il filesystem** in `OUT.PATH/modules/*.bas` e `OUT.PATH/classes/*.cls` (file gia presenti, non `state["files"]` di `app_export`). Questo permette di generare prompt anche quando `EXP.MODULES`/`EXP.CLASSES=False` o su riesecuzione senza ri-esportazione. Filtra con `EXP.PROMPT.LIST` (CSV upper, case-insensitive sullo stem senza estensione): se vuoto -> tutti i `.bas`/`.cls` trovati, se valorizzato -> solo quelli il cui stem upper e nella lista.
- **Diagnostica obbligatoria per ogni chiamata OpenRouter**: prima della `POST` logga `INFO "OpenRouter PROMPT start: <Nome> (model=<MODEL>, src=<relpath>, <N> chars) -> <dest>"`; alla fine logga `INFO "OpenRouter PROMPT done: <Nome> in <Mm SSs> | model=<effettivo, con eventuale fallback> | tokens prompt=<p> completion=<c> total=<t> | cost=$<costo> | -> <dest>"`. Durata da `time.monotonic()` formattata `Xm SSs`; token/costo da `response.usage.{prompt_tokens,completion_tokens,total_tokens,cost}` (`_openrouter_chat` ritorna `(content, usage)`). Su errore HTTP/logica logga `ERROR` e prosegue (non abortisce).

#### `app_python(config: dict, logger) -> str`

- Analogo a `app_prompt` ma con `EXP.PYTHON` / `EXP.PYTHON.LIST`; scansiona `OUT.PATH/modules/*.bas` e `OUT.PATH/classes/*.cls` sul filesystem (non `state`). Genera `.py` corrispondente per ogni `.bas`/`.cls` filtrato (stessa cartella, stesso basename con estensione `.py`). `EXP.PROMPT.LIST` e `EXP.PYTHON.LIST` filtrano **solo** queste due fasi e operano su filesystem, non sull export appena eseguito.
- **Diagnostica identica a `app_prompt`** con prefisso `OpenRouter PYTHON start/done` (operazione, model effettivo, src, chars prima; durata `Mm SSs`, token prompt/completion/total, costo `$` dopo).

#### `app_end(config: dict, export_state: dict) -> str`

- Scrive `index.json` in `OUT.PATH` (accanto a `LOG`, non `manifest.json`) con schema:

```json
{
  "source": "C:/path/db.accdb",
  "exported_at": "2026-05-13T10:30:00Z",
  "config_ini": "C:/path/config.ini",
  "counts": { "forms": 12, "modules": 5, "classes": 4, "queryes": 20, "macros": 2, "reports": 3, "skipped_tables": 15 },
  "files": {
    "forms":   [{ "name": "Form_Clienti", "file_object": "forms/Form_Clienti.txt", "file_prompt": null }],
    "modules": [{ "name": "modUtility",   "file_object": "modules/modUtility.bas", "file_prompt": "modules/prompt_modUtility.md", "file_python": "modules/modUtility.py" }],
    "classes": [{ "name": "clsGestore",   "file_object": "classes/clsGestore.cls", "file_prompt": "classes/prompt_clsGestore.md", "file_python": null }],
    "queryes": [{ "name": "qryClienti",   "file_object": "queryes/qryClienti.sql", "file_prompt": null }],
    "macros":  [{ "name": "AutoExec",     "file_object": "macros/AutoExec.txt",    "file_prompt": null }],
    "reports": [{ "name": "Report_Fattura","file_object":"reports/Report_Fattura.txt","file_prompt": null }]
  },
  "errors": [{ "object_type": "Form", "object_name": "Form_Bad", "error": "SaveAsText failed: ..." }]
}
```

- `file_prompt` / `file_python` valorizzati solo se generati; altrimenti `null`. Path relativi a `OUT.PATH` con `/`.
- **Merge con `index.json` precedente**: se esiste, fonde le entry per `(categoria, name upper)` preservando lo storico (utile quando il run non esporta ma genera solo prompt/python da filesystem); le entry correnti aggiornano `file_object`/`file_prompt`/`file_python` se valorizzati. `counts` ricalcolati sul merge; `skipped_tables` del run se >0 altrimenti storico.
- Scrittura atomica (`tmp` + `replace`), encoding `utf-8`.

---

## 4. REQUISITI FUNZIONALI

### 4.1 Form / Report / Moduli / Classi / Macro

- Usa COM: `win32com.client.Dispatch("Access.Application")` + `Application.SaveAsText`.
- Costanti: `acForm=2, acReport=3, acMacro=4, acModule=5, acQuery=1, acDataAccessPage=6`.
- Sanitizza nomi file come sopra; gestisci nomi con spazi/caratteri non validi e collisioni.

### 4.2 Query

- Itera `CurrentDb.QueryDefs`, ignora di sistema (`~*`, `MSys*`).
- Salva `SQL` in `queryes/<Nome>.sql` (`utf-8`). **Niente `.meta.json`** (rimosso rispetto a bozza: non previsto in INI, era legato a `include_properties` inesistente). Se in futuro serve, documentare estensione in README.

### 4.3 Struttura Output

```
<OUT.PATH>/
  export.log            # path da LOG (relativo a OUT.PATH), in append
  index.json            # riepilogo (non manifest.json)
  forms/Form_Clienti.txt
  modules/modUtility.bas
  modules/prompt_modUtility.md   # se EXP.PROMPT
  modules/modUtility.py          # se EXP.PYTHON
  classes/clsGestore.cls
  classes/prompt_clsGestore.md
  queryes/qryClienti.sql
  macros/AutoExec.txt
  reports/Report_Fattura.txt
```

`index.json` schema come in `app_end`. Nessun file per tabelle. `LOG` contiene riga per categoria + per oggetto.

---

## 5. REQUISITI TECNICI

- **Linguaggio:** Python 3.10+, type hints obbligatori.
- **Dipendenze:** `pywin32` (`win32com.client`), `requests` (per OpenRouter), `argparse` (stdlib), `pathlib`, `logging`, `configparser`, `json`, `dataclasses`, `re`. Niente dipendenze superflue. `comtypes` opzionale come fallback.
- **OS:** Windows 10/11 con Microsoft Access installato (requisito COM). Se Access/COM non disponibile, fallisci con messaggio chiaro (exit 1).
- **Sicurezza:** Apri DB `Exclusive=False`, `AutomationSecurity=1`, chiudi sempre `Quit()` in `try/finally` anche su eccezione.
- **Encoding:** INI `utf-8`; output `utf-8` (`utf-8-sig` se VBA ANSI).
- **Struttura codice:** `src/acc2oth/__main__.py` (entry), `cli.py` (argparse posizionale -> `app_args`), `app.py` (orchestratore `app_args/verify/export/prompt/python/end` con `sResult`), `config.py` (load/validate INI -> dict), `extractor.py` (helper COM `AccessExtractor`), `llm.py` (OpenRouter), `utils.py` (sanitize, ensure_dir, lista). Funzioni `app_*` ben isolate e testabili.

---

## 6. CLI

```
usage: acc2oth <file.ini> [openrouter.key]
  file.ini              Path al file INI di configurazione (posizionale, required)
  openrouter.key        Chiave OpenRouter (posizionale, facoltativo): se passata,
                        si sovrappone a OPENROUTER.KEY del file .ini nel dictConfig
```

Avvio:

```powershell
python -m acc2oth ./config/test.ini
python -m acc2oth C:\path\config.ini
python -m acc2oth C:\path\config.ini sk-or-v1-TUA_CHIAVE
python acc2oth.py ./config/test.ini sk-or-v1-TUA_CHIAVE
```

Comportamento:
- Primo argomento posizionale `file.ini` (required). Secondo argomento posizionale `openrouter.key` (facoltativo): se presente e non vuoto, sovrascrive `OPENROUTER.KEY` del `.ini` nel `dictConfig` finale.
- **Non** introdurre flag extra (`--config`, `--dry-run`, `--verbose`) se non documentati come divergenza in README (la specifica INI non li prevede). Se aggiunti, devono essere opzionali e non rompere `python -m acc2oth file.ini`.
- Se invocato senza parametri o con `file.ini` inesistente: stampa errore + help completo (usage, esempi, schema INI minimo di cap. 3.1) e `exit 2`.
- Valida INI contro schema cap. 3 prima di aprire Access; su errore stampa su `stderr` e `exit 2`.
- Su successo `exit 0`, su errore estrazione `exit 1`, su errore validazione/config `exit 2`.

---

## 7. ERRORI E LOG

- `logging` su console + file indicato da `LOG` (sotto `OUT.PATH`, in append). Formato: `%(asctime)s [%(levelname)s] %(message)s`.
- Livelli: `INFO` (oggetti estratti/skipped, categorie), `WARNING` (nomi sanitizzati, alias QUERYES/QUERIES, overwrite implicito), `ERROR` (fallimento singolo oggetto — **continua** con gli altri, raccogli in `index.json.errors`).
- Non interrompere l intera estrazione per un singolo oggetto fallito.
- Gestisci esplicitamente: file `.accdb` inesistente/non apribile, DB locked (`Could not use ... already in use`), Access/COM non installato, `OUT.PATH` non scrivibile, `LOG` non scrivibile, INI malformato / sezione `[CONFIG]` mancante, `OPENROUTER.*` mancanti nel `dictConfig` finale quando richiesti (`EXP.PROMPT`/`EXP.PYTHON=True`, tenendo conto dell override CLI), HTTP 4xx/5xx OpenRouter (retry 1 volta poi logga errore e prosegue), sanitizzazione/collisione nomi.

---

## 8. CRITERI DI ACCETTAZIONE (per validare il codice generato)

- [ ] `python -m acc2oth example.ini` crea `OUT.PATH/{forms,modules,classes,queryes,macros,reports}/` + `LOG` + `index.json` con schema cap. 3.2
- [ ] Ogni Form/Report/Modulo/Classe/Query/Macro del DB compare come file singolo con estensione corretta (conteggio vs `AllForms`/`AllModules`/`QueryDefs` filtrate); nessun file per tabelle
- [ ] `LOG` contiene riga per categoria + per ogni oggetto esportato
- [ ] `index.json` valido con `name`/`file_object`/`file_prompt`/`file_python` per categoria e `errors` per fallimenti singoli
- [ ] `EXP.PROMPT.LIST` / `EXP.PYTHON.LIST` filtrano solo generazione prompt/python, non l export grezzo
- [ ] Sanitizzazione nomi e filtro QueryDef di sistema (`~*`, `MSys*`)
- [ ] Chiusura pulita di `Access.Application.Quit()` anche su eccezione (verifica Task Manager)
- [ ] Testato su `.accdb` di esempio con almeno 2 Form, 1 Modulo standard, 1 Classe, 2 Query, 1 Macro, 1 Report
- [ ] Exit code 0/1/2 come da cap. 6-7; validazione INI prima di aprire Access
- [ ] Nessun `manifest.json`/`logs/`/`export_mode`/`naming` residuo non documentato

---

## APPENDICE — SaveAsText snippet (riferimento)

```python
import win32com.client

acForm, acReport, acMacro, acModule, acQuery = 2, 3, 4, 5, 1

app = win32com.client.Dispatch("Access.Application")
try:
    app.AutomationSecurity = 1  # msoAutomationSecurityLow
    app.OpenCurrentDatabase(r"C:\path\db.accdb", Exclusive=False)
    # NOTA: SaveAsText e un metodo di Application, non di DoCmd.
    app.SaveAsText(acForm,   "Form_Clienti", r"C:\out\forms\Form_Clienti.txt")
    app.SaveAsText(acModule, "modUtility",   r"C:\out\modules\modUtility.bas")
    app.SaveAsText(acQuery,  "qryClienti",   r"C:\out\queryes\qryClienti.sql")  # oppure via QueryDefs.SQL
    app.SaveAsText(acMacro,  "AutoExec",     r"C:\out\macros\AutoExec.txt")
    app.SaveAsText(acReport, "Report_Fattura", r"C:\out\reports\Report_Fattura.txt")
finally:
    try: app.CloseCurrentDatabase()
    except Exception: pass
    app.Quit()
```

`SaveAsText`/`LoadFromText` sono metodi non documentati ma stabili da Access 2000+ e sono lo standard de-facto per versionamento Access/VCS. Per query preferibile iterare `CurrentDb.QueryDefs` e scrivere `QueryDef.SQL` direttamente.

---

## APPENDICE — Esempio conversione BAS/CLS -> Prompt e -> Python (OpenRouter)

### BAS/CLS -> Prompt (System Prompt / specifica per rigenerazione)

```python
import os
import requests

def convert_bas_to_generation_prompt(
    bas_file_path: str,
    api_key: str,
    model_id: str = "qwen/qwen-2.5-coder-32b-instruct",
) -> str:
    """Legge .bas/.cls (VBA), lo invia a OpenRouter e restituisce un System Prompt
    dettagliato per rigenerare la procedura in Python. Non genera codice, solo specifica."""
    if not os.path.exists(bas_file_path):
        raise FileNotFoundError(f"File non trovato: {bas_file_path}")
    with open(bas_file_path, "r", encoding="utf-8", errors="ignore") as f:
        vba_code = f.read()
    url = "https://openrouter.ai/api/v1/chat/completions"
    system_instruction = (
        "Sei uno specialista in refactoring da VBA/Access a Python. "
        "Analizza il codice del modulo .bas/.cls fornitoti. NON generare codice Python. "
        "Il tuo compito e creare un PROMPT DI GENERAZIONE (System Prompt / Specifica Tecnica) "
        "estremamente dettagliato che permettera a uno sviluppatore o a un altro LLM di riscrivere "
        "da zero la logica in Python.\n\n"
        "Il prompt prodotto deve definire:\n"
        "1. Scopo e workflow del modulo.\n"
        "2. Input/Output (tabelle, query, file esterni gestiti).\n"
        "3. Regole di business, formule, cicli e gestione eccezioni.\n"
        "4. Librerie e architettura Python suggerite (es. pandas, openpyxl, sqlalchemy)."
    )
    payload = {
        "model": model_id,
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": f"Ecco il codice VBA (.bas/.cls):\n\n```vba\n{vba_code}\n```"},
        ],
        "temperature": 0.2,
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    response = requests.post(url, json=payload, headers=headers, timeout=60)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]

if __name__ == "__main__":
    MY_API_KEY = os.environ.get("OPENROUTER_API_KEY", "sk-or-v1-TUA_KEY")
    print(convert_bas_to_generation_prompt("ModuloEsempio.bas", MY_API_KEY))
```

### BAS/CLS -> Python (transpiling diretto)

Variante per `EXP.PYTHON`: stessa struttura ma `system_instruction` = "Sei un transpiler VBA->Python. Traduci il modulo in Python 3.10+ idiomatico, preservando logica, con type hints e gestione errori..." e il risultato va scritto come `.py`. Gestire `response.raise_for_status()` con retry singolo su 429/5xx.

---

## NOTE PER L AGENT

- Prima di generare, leggi anche `ricerche.md` (limiti transpiling Access->Python/Java, mapping DAO/ADODB -> SQLAlchemy, Form Win32 -> framework).
- Documenta in README ogni assunzione fatta su punti ambigui (sanitizzazione, overwrite implicito, LOG relativo, alias QUERYES).
# mcp4office

> 🇮🇹 **Italiano** | 🇬🇧 **[English](README.en.md)**

Server **MCP (Model Context Protocol)** locale scritto in **C# (.NET Framework 4.8)** per pilotare **Microsoft Word, Excel, PowerPoint, Outlook classico** e leggere documenti **PDF / Acrobat** direttamente dal proprio client AI — senza Python, senza Node e senza chiavi API aggiuntive.

- **Pagina download e guida (GitHub Pages):** https://pai-office-automation.github.io/mcp4office/
- **Ultima release (`v0.5.7 beta`):** https://github.com/pai-office-automation/mcp4office/releases/tag/v0.5.7
- **Privacy e sicurezza:** https://pai-office-automation.github.io/mcp4office/privacy.html

---

## A cosa serve in dettaglio e con quali connettori lavora

`mcp4office` fa da **ponte locale** tra un client/agente AI esterno e le applicazioni di Office già aperte sul PC, esponendo **32 tool** di lettura, confronto e scrittura controllata. È progettato per flussi di lavoro conversazionali e **multi-applicazione** (ad esempio: leggere un PDF, confrontare due versioni di un Word o di un foglio Excel, aggiornare tabelle, generare slide PowerPoint con grafici e preparare bozze in Outlook nella stessa chat, senza copia-incolla).

- **Model-free e senza chiavi API:** `mcp4office` non integra alcun modello AI, non effettua chiamate API verso provider esterni e non richiede chiavi API a consumo. Sfrutta il modello e l'abbonamento già inclusi nel client AI che stai usando.
- **Connettori lato AI (protocollo MCP su `stdio`):** comunica via **JSON-RPC 2.0 su `stdio`** (nessuna porta di rete aperta). Il Setup configura automaticamente **Antigravity / Antigravity IDE**, **Claude Desktop** (classico e Microsoft Store/MSIX), **Claude Code** e **Codex**; è utilizzabile con qualsiasi altro client compatibile con MCP tramite configurazione manuale.
- **Connettori lato Office e documenti:** utilizza automazione **COM locale** sulle istanze già aperte di **Word, Excel, PowerPoint, Outlook classico** e **Adobe Acrobat** (Standard/Pro), più lettura diretta di **file PDF locali** (anche senza Acrobat) entro le cartelle consentite.

---

## Differenze rispetto a Personal AI Assistant (`plugin_word`)

Entrambi i progetti integrano l'AI con Microsoft Office su Windows, ma adottano due architetture e modalità d'uso opposte:

| Aspetto | `mcp4office` (questo progetto) | Personal AI Assistant (`plugin_word`) |
| --- | --- | --- |
| **Dove si usa** | **Dall'esterno di Office**, nella chat del tuo client AI (Antigravity, Claude Desktop/Code, Codex, ecc.) | **Dentro Office**, come scheda *Personal AI* (pannello laterale in Word/Excel/PowerPoint e pulsanti in Outlook) |
| **Modello AI e costi** | **Model-free**: nessun modello integrato e nessuna chiave API; usa l'abbonamento del tuo client MCP | **BYOK (chiave API personale)**: chiamate dirette a pagamento/piano gratuito verso le API del provider, oppure modello locale (Ollama / `llama-server` / LM Studio) |
| **Connettori verso l'AI** | **MCP (JSON-RPC 2.0 su `stdio`)** verso client/agenti compatibili | **HTTPS REST diretti** verso Google Gemini, OpenAI, Anthropic Claude, Qwen, MiniMax, DeepSeek, OpenRouter o endpoint locale OpenAI-compatibile |
| **Connettori verso Office** | **COM locale** su istanze aperte di Word, Excel, PowerPoint, Outlook classico, Acrobat + lettura diretta **PDF locali** | **Office.js** (WebView2) per Word, Excel, PowerPoint + add-in **VSTO (.NET 4.8)** per Outlook classico (nessun supporto PDF/Acrobat) |
| **Tipo di utilizzo** | Flussi **agentici e cross-app** (confronti tra documenti, lettura PDF, manipolazione multi-file, automazione in chat) | Azioni **rapide e guidate** sul documento aperto (riscrivi, traduci, formule, analisi, slide, bozze email) senza uscire da Office |
| **Diritti di installazione** | 100% profilo utente, **nessun UAC / amministratore** | Profilo utente per Outlook; richiede **1 conferma amministratore (UAC)** per la cartella condivisa di Word/Excel/PowerPoint |

In sintesi: scegli **`mcp4office`** se usi già un client MCP (come Antigravity, Claude Desktop o Codex), non vuoi gestire chiavi API e desideri operazioni multi-documento/PDF; scegli **Personal AI Assistant** se preferisci lavorare direttamente dentro Office con pulsanti dedicati usando la tua chiave API (o un modello locale).

---

## Installazione automatica (1 clic, senza amministratore)

Scarica ed esegui [`mcp4office-Setup-0.5.7.exe`](https://github.com/pai-office-automation/mcp4office/releases/tag/v0.5.7) (oppure `dist\mcp4office-Setup-0.5.7.exe`, generato da `build.ps1`):
- **Nota SmartScreen:** l'eseguibile locale non è firmato con certificato commerciale; Windows può mostrare l'avviso SmartScreen (*Ulteriori informazioni → Esegui comunque*).
- Si installa nel profilo utente (`%LOCALAPPDATA%\Programs\mcp4office`), **senza UAC**.
- Se `mcp4office.exe` è in esecuzione (aperto da un client AI), il Setup avvisa e chiede di chiudere i client prima di proseguire.
- Se è già installato, rimuove prima i vecchi binari **preservando configurazioni MCP e opzioni utente** (es. `--sola-lettura`, `--pdf-roots`) tramite marcatore `.upgrading`, poi aggiorna e verifica tutti i client.
- Configura automaticamente i client rilevati (**Antigravity**, **Claude Desktop** sia classico sia Store/MSIX, **Codex**) tramite `mcp4office.exe --configure-clients`, creando una copia di sicurezza `<file>.mcp4office.bak` alla prima modifica di file preesistenti e mostrando a schermo l'esito reale per ciascun client.
- Alla **disinstallazione** (da *App installate* di Windows) rimuove solo le voci e i file creati dall'installer; le eventuali copie di sicurezza `<file>.mcp4office.bak` accanto ai config preesistenti vengono intenzionalmente conservate per sicurezza.
- Dopo l'installazione, **riavvia i client AI aperti** per caricare i 32 tool.

---

## Configurazione manuale (senza Setup o per opzioni personalizzate)

Opzioni CLI supportate da `mcp4office.exe`:
- `--sola-lettura` / `--read-only`: espone solo i 16 tool di lettura.
- `--pdf-roots "C:\cartella1;D:\cartella2"`: cartelle consentite per PDF e immagini (alternativa alla variabile `MCP4OFFICE_PDF_ROOTS`; default `Documenti`, `Desktop`, `Download`).
- `--self-test`: esegue tutti i test interni senza toccare Office.

**Antigravity** (`%USERPROFILE%\.gemini\config\mcp_config.json` e/o `.gemini\antigravity-ide\mcp_config.json`) o **Claude Desktop** (`%APPDATA%\Claude\claude_desktop_config.json` oppure `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "mcp4office": {
      "command": "C:\\Users\\<utente>\\AppData\\Local\\Programs\\mcp4office\\mcp4office.exe",
      "args": []
    }
  }
}
```

**Codex** (`%USERPROFILE%\.codex\config.toml`):

```toml
[mcp_servers.mcp4office]
command = "C:\\Users\\<utente>\\AppData\\Local\\Programs\\mcp4office\\mcp4office.exe"
args = []
startup_timeout_sec = 20
tool_timeout_sec = 60
```

---

## Trasparenza, privacy e sicurezza

- **Trasporto `stdio` locale ≠ elaborazione AI locale:** `mcp4office.exe` gira sul tuo PC senza aprire porte di rete, ma **i testi letti dai documenti/email/PDF vengono inviati al modello AI del client** secondo le condizioni del tuo abbonamento.
- **Dati non fidati (*Prompt injection*):** documenti, email e PDF sono *dati*, mai istruzioni. Verifica sempre **destinatari (`To`/`CC`)** nelle bozze Outlook e i bersagli prima di confermare modifiche.
- **Protezioni principali (v0.4.0):**
  - Validazione rigorosa di schema JSON e lifecycle MCP (`-32002` prima di `initialize`, `-32602` su parametri mancanti/invalidi).
  - Ancoraggio esplicito del bersaglio (`expected_doc`, `expected_table_token`, `expected_workbook`, `expected_sheet`, `expected_presentation`, `slide_id`, `entry_id`/`store_id`) + `state_token` SHA-256 tipizzato.
  - **Word:** `Range.Duplicate` + `StoryType`, blocco selezioni multi-cella (`Cells.Count > 1`), modifica sicura di singole celle tabella via `word_edit_table_cell`, creazione tabelle via `word_insert_table`, `UndoRecord` *best-effort* per singolo `Ctrl+Z`.
  - **Excel:** **la scrittura via COM azzera `Ctrl+Z`.** Lettura/scrittura in blocco, date ISO (`parse_iso_dates=true`), testo letterale di default (`allow_formulas=false`), blocco sovrascrittura con `state_token` tipizzato, rollback verificato tramite rilettura, nuovi fogli (`excel_add_sheet`) e grafici nativi collegati all'intervallo (`excel_add_chart`).
  - **PowerPoint:** `SlideID` stabile (`FindBySlideID`), layout `title_only` e area contenuto dinamica (sicura anche su 4:3 senza sovrapposizioni), grafica nativa con colori/font del tema raggruppata per un solo clic, grafici nativi con escape letterale di categorie/serie e cleanup automatico su errore.
  - **Outlook:** **solo bozze** aperte a schermo (`Display()`), firma preservata (`InsertAfterBody`), mai invio automatico.
  - **PDF e immagini locali:** radici consentite, blocco UNC/device path e symlink/junction prima di `File.Exists` (consentendo tag cloud OneDrive), limiti 50 MB PDF / 25 MB immagini / 100 pagine per chiamata.

---

## Elenco dei 32 tool esposti

Con `--sola-lettura` (o `--read-only`) vengono registrati solo i 16 tool di lettura.

### Microsoft Word (11 tool)
- `word_status` *(lettura)*, `word_read_selection` *(lettura)*, `word_read_document` *(lettura, supporta `doc_id`)*, `word_compare_documents` *(lettura, diff a parole + opzionale `CompareDocuments`)*
- `word_replace_selection` *(scrittura)*, `word_insert_text` *(scrittura)*
- `word_list_tables` *(lettura)*, `word_read_table` *(lettura)*, `word_edit_table_cell` *(scrittura)*, `word_batch_edit_cells` *(scrittura multipla fino a 200 celle in 1 Ctrl+Z)*, `word_insert_table` *(scrittura)*

### Microsoft Excel (6 tool)
- `excel_status` *(lettura)*, `excel_read_range` *(lettura, supporta `workbook`)*, `excel_compare_sheets` *(lettura, confronto cella per cella valori + formule)*
- `excel_write_cells` *(scrittura)*, `excel_add_sheet` *(scrittura)*, `excel_add_chart` *(scrittura)*

### Microsoft PowerPoint (8 tool)
- `ppt_status` *(lettura)*, `ppt_read_slides` *(lettura)*
- `ppt_add_slide` *(scrittura)*, `ppt_set_slide_notes` *(scrittura)*
- `ppt_add_table` *(scrittura)*, `ppt_add_visual_blocks` *(scrittura)*, `ppt_add_chart` *(scrittura)*, `ppt_add_picture` *(scrittura)*

### Microsoft Outlook classico (4 tool)
- `outlook_list_selected` *(lettura)*, `outlook_list_inbox` *(lettura)*
- `outlook_create_draft` *(bozza)*, `outlook_create_reply_draft` *(bozza)*

### PDF e Adobe Acrobat (3 tool)
- `pdf_info` *(lettura)*, `pdf_read_file` *(lettura)*, `acrobat_read_active` *(lettura)*

---

## Compilazione

```powershell
.\build.ps1
```

---

## Limiti noti

- **Outlook Object Model Guard:** da processo esterno Outlook può mostrare l'avviso di sicurezza se l'antivirus non è riconosciuto nel *Centro protezione → Accesso a livello di codice*.
- **Outlook classico only:** il "Nuovo Outlook" non espone COM.
- **Excel `Ctrl+Z`:** azzerato dalla scrittura COM.
- **Acrobat Reader gratuito:** non espone COM; usare `pdf_read_file` con il percorso del file.
- **Stessi privilegi UAC:** Office e client AI devono girare entrambi come utente normale.

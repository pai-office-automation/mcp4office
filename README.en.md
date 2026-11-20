# mcp4office

> 🇬🇧 **English** | 🇮🇹 **[Italiano](README.md)**

Local **MCP (Model Context Protocol)** server written in **C# (.NET Framework 4.8)** to drive **Microsoft Word, Excel, PowerPoint, classic Outlook** (via Windows COM on already-open instances) and read **PDF / Adobe Acrobat** documents using your existing flat-rate AI subscription on **Antigravity**, **Claude Desktop** (classic and Microsoft Store/MSIX), **Claude Code**, or **Codex** — **no Python, no Node.js, and no pay-per-token API keys required**.

- **Download & Guide (GitHub Pages):** https://pai-office-automation.github.io/mcp4office/
- **Latest Release (`v0.5.7 beta`):** https://github.com/pai-office-automation/mcp4office/releases/tag/v0.5.7
- **Privacy & Security:** https://pai-office-automation.github.io/mcp4office/privacy.html
- **Prefer an in-Office sidebar using your own API key?** See [Personal AI Assistant for Office](https://github.com/pai-office-automation/personal-ai-assistant-office).

---

## One-Click Installation (No Admin / No UAC)

Download and run `mcp4office-Setup-0.5.7.exe` from the [latest release](https://github.com/pai-office-automation/mcp4office/releases/latest) (or build locally with `.\build.ps1`):

- **SmartScreen note:** the installer is not signed with a commercial certificate; Windows may show a SmartScreen prompt (*More info → Run anyway*).
- Installs per-user into `%LOCALAPPDATA%\Programs\mcp4office` (**no admin rights / no UAC**).
- If `mcp4office.exe` is currently running (held open by an AI client), Setup alerts you to close open AI clients first.
- On upgrade, preserves your existing MCP config and custom CLI flags (e.g. `--sola-lettura`, `--pdf-roots`).
- Automatically configures detected clients (**Antigravity**, **Claude Desktop** classic & Store/MSIX, **Claude Code**, **Codex**) via `mcp4office.exe --configure-clients`, creating a one-time `<file>.mcp4office.bak` backup of pre-existing configs.
- Uninstalling from Windows *Installed Apps* cleanly removes only the entries/files created by the installer (keeping `.mcp4office.bak` backups intact).
- Restart your AI client after installation to load all **32 tools**.

---

## CLI Options & Manual Configuration

Supported CLI flags for `mcp4office.exe`:
- `--sola-lettura` / `--read-only`: exposes only the **16 read-only tools**.
- `--pdf-roots "C:\folder1;D:\folder2"`: allowed root directories for PDF and image files (alternative to `MCP4OFFICE_PDF_ROOTS`; defaults to `Documents`, `Desktop`, `Downloads`).
- `--self-test`: runs all 10 built-in verification suites without touching Office.
- `--list-tools`: prints the registered MCP tool catalog.

**Antigravity / Claude Desktop (JSON):**
```json
{
  "mcpServers": {
    "mcp4office": {
      "command": "C:\\Users\\<user>\\AppData\\Local\\Programs\\mcp4office\\mcp4office.exe",
      "args": []
    }
  }
}
```

**Codex (`%USERPROFILE%\.codex\config.toml`):**
```toml
[mcp_servers.mcp4office]
command = "C:\\Users\\<user>\\AppData\\Local\\Programs\\mcp4office\\mcp4office.exe"
args = []
startup_timeout_sec = 20
tool_timeout_sec = 60
```

---

## Transparency, Privacy & Safety

- **Local `stdio` transport ≠ local AI processing:** `mcp4office.exe` opens zero network ports and sends no telemetry, but **text read from your documents/emails/PDFs is passed to your AI client's model** under your client's subscription terms.
- **Untrusted data (Prompt injection):** documents, emails, and PDFs are *data*, never instructions. Always review Outlook draft recipients (`To`/`CC`) and proposed edits before confirming.
- **Built-in safeguards:**
  - Strict JSON-RPC 2.0 lifecycle and input schema validation (`-32002` before `initialize`, `-32602` on invalid/missing arguments).
  - Explicit target anchoring (`expected_doc`, `expected_table_token`, `expected_workbook`, `expected_sheet`, `expected_presentation`, `slide_id`, `entry_id`/`store_id`) + typed SHA-256 `state_token` checks before any overwrite.
  - **Word:** `Duplicate + SetRange` preserving the original Story (MainText, Header, Footer, TextBox, Footnote), multi-cell selection guard, in-place formatting preservation (paragraph styles + mixed intra-paragraph bold/italic/color), batch editing up to 200 table cells in a single `Ctrl+Z` (`UndoRecord`) with partial-failure reporting.
  - **Excel:** **COM writes clear Excel's `Ctrl+Z` undo stack.** Bulk read/write, ISO date parsing (`parse_iso_dates=true`), literal text by default (`allow_formulas=false`), typed overwrite token + verified rollback, keyed sheet diff (`key_column`, `has_header`, duplicate-key & missing-key reporting).
  - **PowerPoint:** stable `SlideID` (`FindBySlideID`), theme-aware visual blocks, native charts with literal label escaping and cleanup on error.
  - **Outlook:** **drafts only**, opened on screen (`Display()`), signature preserved, never sends automatically.
  - **PDF & local images:** allowed roots only, blocks UNC/device paths and symlinks/junctions before `File.Exists` (allowing OneDrive cloud files), size/page caps (50 MB PDF / 25 MB image / 100 pages per call).

---

## 32 Exposed Tools (16 Read-Only)

- **Microsoft Word (11):** `word_status`, `word_read_selection`, `word_read_document`, `word_compare_documents`, `word_replace_selection`, `word_insert_text`, `word_list_tables`, `word_read_table`, `word_edit_table_cell`, `word_batch_edit_cells`, `word_insert_table`
- **Microsoft Excel (6):** `excel_status`, `excel_read_range`, `excel_compare_sheets`, `excel_write_cells`, `excel_add_sheet`, `excel_add_chart`
- **Microsoft PowerPoint (8):** `ppt_status`, `ppt_read_slides`, `ppt_add_slide`, `ppt_set_slide_notes`, `ppt_add_table`, `ppt_add_visual_blocks`, `ppt_add_chart`, `ppt_add_picture`
- **Microsoft Outlook classic (4):** `outlook_list_selected`, `outlook_list_inbox`, `outlook_create_draft`, `outlook_create_reply_draft`
- **PDF & Adobe Acrobat (3):** `pdf_info`, `pdf_read_file`, `acrobat_read_active`

---

## Building from Source

```powershell
.\build.ps1
```

Requires .NET SDK (targets `net48` via `Microsoft.NETFramework.ReferenceAssemblies`) and Inno Setup 6 (`ISCC.exe`) to produce `dist\mcp4office-Setup-<version>.exe` + `SHA256SUMS.txt`.

## License

[MIT License](LICENSE). Independent open-source project, not affiliated with Microsoft, Google, OpenAI, Anthropic, or Adobe.

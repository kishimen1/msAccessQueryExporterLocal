# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Windows-only utility that extracts schema and queries from Microsoft Access databases (.accdb, .mdb) and visualizes them in an interactive browser-based viewer. No build tools, package managers, or server required.

**Entry point:** `AccessQueryExporter.bat`
**Stack:** Windows Batch, VBScript (DAO), vanilla HTML/CSS/JavaScript

## How It Works

1. `AccessQueryExporter.bat` — Opens a PowerShell file dialog, validates the selected `.accdb`/`.mdb` file, and invokes the VBScript extractor via `cscript`.
2. `ExtractToJSON.vbs` — Uses DAO (tries `DAO.DBEngine.120` first, falls back to `DAO.DBEngine.36`) to extract tables, fields, relationships, and queries. Outputs:
   - `{filename}_result.json` — structured data for the viewer
   - `{filename}_クエリ一覧.txt` — human-readable text report
3. `viewer.html` — Single-file offline web app that loads the JSON and provides filtering, search, SQL formatting, and dark/light theming. Opened automatically by the batch file.

## Architecture Notes

**Three files, strict separation of layers:**
- **Extraction layer** (`ExtractToJSON.vbs`): database I/O only; outputs platform-neutral JSON
- **Data layer** (JSON output): schema defined in VBScript — `{filename, exported_at, tables[], relationships[], queries[]}`
- **Presentation layer** (`viewer.html`): self-contained SPA; no server, no build step

**viewer.html internals (1974 lines, single file):**
- All CSS in `<style>`, all JS in `<script>` — no bundler
- Theme persistence via `localStorage` key `aqe-theme`
- Custom SQL formatter (`formatSQL`) and parser (`parseSQL`) — no external SQL library
- Optional CDN: Google Fonts + Highlight.js (graceful degradation if offline)
- Key JS functions: `loadJsonFile`, `renderResults`, `formatSQL`, `parseSQL`, `applyFilters`, `downloadAll`, `resetApp`

**Query type numeric mapping** (used in VBScript and viewer):
- 0=選択, 1=クロス集計, 2=削除, 3=更新, 4=追加, 5=テーブル作成, 6=データ定義, 7=パススルー, 8=実行可能, 9=サブフォーム/サブレポート

## Running

```bat
# Interactive (file dialog opens)
AccessQueryExporter.bat

# Direct invocation
AccessQueryExporter.bat "C:\path\to\database.accdb"
```

Requires Windows with DAO library (installed with Microsoft Office / Access Runtime).

## Localization

All UI text, comments, and documentation are in Japanese. Maintain this convention when editing.

---
name: access-vba-reserved-words
description: >-
  Detect and avoid Microsoft Access / VBA reserved words in identifiers.
  Use when naming variables, fields/controls, modules, procedures, queries, or SQL columns.
  Guide learners to safe alternatives and consistent conventions.
license: CC-BY-4.0
---

# Access/VBA Reserved Words – Naming Safety Skill

## Integration Hint
Use this skill **together with** the repository’s custom instructions (e.g., `.github/instructions/vba-naming.instructions.md`).
The instructions provide **always-on guardrails** for naming; this skill provides **on-demand** scanning, reporting, and renaming guidance.


## What this skill does
- **Scans identifiers** in code and SQL for case-insensitive collisions with VBA keywords/operators, built‑in function names (`InStr`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`), Access/DAO names (`Form`, `Fields`, `Recordset`), and ACE/Jet SQL keywords (`SELECT`, `FROM`, `WHERE`, `JOIN`, etc.).
- **Reports findings** with file/line references and proposes **safer replacements** (e.g., `Date` → `SaleDate`, `Name` → `PersonName`, `Text` → `BodyText`).
- **Applies conventions**: CamelCase, no spaces/specials, optional Leszynski/Reddick prefixes (`strName`, `lngCount`, `dtmStart`, `frmMain`, `qrySales`).

> Access treats names **case-insensitively**, so reusing these terms as identifiers often leads to compile/runtime errors or ambiguous references. (Authoritative guidance and lists: Microsoft Access reserved words & symbols; ACE/Jet SQL reserved words; VBA keywords/spec; Allen Browne’s problem names.)  
> See References section.

## How to run the scan
1. Ensure PowerShell is available (Windows/macOS/Linux with PowerShell 7+).
2. From the repo root, execute:

```powershell
.\.github\skills\access-vba-reserved-words\scripts\scan_reserved_words.ps1 `
  -Path . `
  -CsvPath .github\skills\access-vba-reserved-words\data\reserved_words_full.csv `
  -Extensions ".bas,.cls,.frm,.vba,.sql,.txt"

## Purpose
Help developers avoid case-insensitive collisions with:
- **VBA keywords & operators** (`Dim`, `If`, `Select`, `Function`, `ByVal`, `And`, etc.).
- **Built-in VBA function names** (`InStr`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`, etc.).
- **Access/DAO object model names** (`Form`, `Forms`, `Report`, `Section`, `Field`, `Fields`, `Recordset`, etc.).
- **ACE/Jet SQL keywords** (`SELECT`, `FROM`, `WHERE`, `JOIN`, `GROUP BY`, `ORDER BY`, `TABLE`, `INDEX`, `UNION`, etc.).
- **Special characters/symbols** (spaces, `'`, `"`, `.`, `!`, `?`, `*`, `+`, `-`, `=`, `<`, `>`, `#`, `%`, `$`, `&`, `@`, `\`, `/`, `^`, `~`, `{}`, `[]`, `()`).
> Access and the database engine treat names **case-insensitively**; reusing these terms as identifiers often leads to compile or runtime errors.

## Procedure
1. **Scan identifiers** in the current context (variables, procedure names, control names, field names, query columns, and embedded SQL).
2. **Flag exact case-insensitive matches** to reserved words and report each occurrence with the location.
3. **Suggest safe replacements** using descriptive names:
   - `InStr` → `posInStr` or `searchPos`
   - `Date` → `SaleDate` / `CreatedDate`
   - `Name` → `PersonName` / `ItemName`
   - `Text` → `BodyText` / `MessageText`
4. **Apply conventions**:
   - Use CamelCase, no spaces or special characters.
   - Prefer Leszynski/Reddick-style type prefixes for variables (e.g., `strName`, `lngCount`, `dtmStart`) and object prefixes (`frmMain`, `qrySales`).
5. **If renaming Access objects** already in use, prefer a refactor rather than bracketing; `[Date]` can work but renaming avoids subtle bugs.
6. **(Optional) Run a repository scan** when requested to list offenders and propose bulk renames.

## Common Offenders (teach by example)
- Identifiers: `date`, `time`, `now`, `value`, `name`, `text`, `year`
- VBA functions used as names: `instr`, `left`, `right`, `mid`, `len`
- Access/DAO object names: `form`, `forms`, `report`, `section`, `field`, `fields`
- SQL words: `select`, `from`, `where`, `join`, `order`, `group`, `union`, `table`, `index`

## Naming Recommendations
- ✅ Prefer descriptive, specific names: `SaleDate`, `IsEligible`, `TotalAmount`.
- ✅ Use CamelCase; consider type/object prefixes (`str`, `lng`, `dtm`, `frm`, `qry`) for clarity.
- ❌ Avoid reserved words and special characters in any identifier.
- ❌ Avoid generic names like `Data`, `Info`, `Temp` when they collide with engine terms.

## Example Prompts (to trigger this skill)
- “Scan this module and flag any **Access/VBA reserved-word** variable names.”
- “Suggest safe replacements for fields named `Date` and `Text` across forms.”
- “Check my SQL queries for column names that collide with ACE/Jet reserved words.”

## References
- Microsoft: Avoid using reserved words and symbols in Access
- Microsoft: SQL reserved words (ACE/Jet engine)
- VBA language keywords & reserved identifiers
- Allen Browne: Problem names & reserved words in Access
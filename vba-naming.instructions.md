---
description: Always-on guardrails for Access/VBA naming
applyTo:
  - "**/*.bas"
  - "**/*.cls"
  - "**/*.frm"
  - "**/*.vba"
---

# Access/VBA Naming Guardrails

- **Do not** use **reserved words** for identifiers (variables, procedures, modules, forms/reports, controls, fields, query columns, or SQL).
- Treat reserved words as **case-insensitive**; `instr` collides with `InStr()` and must be renamed.
- Prefer **CamelCase** and **no spaces/special characters** in names.
- Prefer **Leszynski/Reddick-style prefixes** (e.g., `strName`, `lngCount`, `dtmStart`; `frmMain`, `qrySales`) for clarity.
- If a collision already exists, **rename** rather than relying on brackets (e.g., `[Date]`).
- When generating code, **ask for replacements** if the user suggests a reserved word (e.g., propose `SaleDate` instead of `Date`).

> Reference lists: Microsoft Access reserved words & symbols, ACE/Jet SQL reserved words, and VBA keywords/specification.
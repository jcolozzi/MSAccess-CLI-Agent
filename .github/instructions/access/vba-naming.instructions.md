---
description: Always-on guardrails for Access/VBA naming
applyTo: "**/*.{bas,cls,frm,vba}"
---

# Access/VBA Naming Guardrails

## Core Rules

### Reserved Words
Do not use reserved words for identifiers (variables, procedures, modules, forms/reports, controls, fields, or query columns).
- Reserved words are **case-insensitive**: `instr` collides with `InStr()` and must be renamed.
- **Rename** rather than bracket escaping (prefer `SaleDate` over `[Date]`).

### Naming Standards
- Names must start with a letter, use no spaces/special characters, and stay concise.
- Prefer CamelCase and Leszynski/Reddick-style prefixes.
- Use `Option Explicit` in every module.

### Refactoring Existing Code
- Refactor to comply if the change does not break existing functionality and is reasonably scoped (< 2 hours).
- For legacy code with external constraints, document the exception and reason.
- Prioritize new code and critical paths over legacy code.

## Access-Specific Collision Hazards

- Avoid ACE/Jet SQL reserved words as table, field, or query-column names.
- Avoid Access object model names as identifiers: `Forms`, `Reports`, `Controls`, `Modules`, `Pages`.
- Avoid common ambiguous names that collide in Access SQL/VBA: `Name`, `Value`, `Date`, `Count`, `Index`.

## Prefix Guidance

- Forms/reports/queries/tables/controls: `frmMain`, `rptSales`, `qrySales`, `tblCustomers`, `ctlSearch`.
- Variables: `strName`, `lngCount`, `dblRate`, `blnFound`, `dtmStart`, `objItem`, `arrData`.
- Modules/classes: `modUtilities`, `clsLogger`.

## Procedure Naming

- Use PascalCase and verb-first names: `CalculateTotal`, `LoadCustomer`, `SaveRecord`.
- Keep event handlers in `ObjectName_EventName` format.

> Reference lists: Microsoft Access reserved words and symbols, ACE/Jet SQL reserved words, and VBA keywords/specification.
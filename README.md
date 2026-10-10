# MULTIVAC Purchase Order

A lightweight web application that fills `Template_Purchase_Order_2026.xlsx` directly from a MULTIVAC quotation PDF with zero loss of workbook fidelity.

## Architecture

- **Server-Side Gemini API (`gemini-3.8-flash`)**:
  - Independent double-extraction (Extractor pass + Auditor pass) with field-level reconciliation.
  - If differences are detected, a Senior Auditor pass reconciles the positions.
  - Ambiguities or mismatches halt execution safely with an amber stop screen.
  - Secure server-side processing via Express (`/api/extract`) and Vercel serverless function (`api/extract.ts`).
  - Automatic exponential backoff retry and fallback to `gemini-3.1-flash-lite`.

- **Package-Level Direct XML OpenXML Editing**:
  - Uses `JSZip` and DOMParser/XMLSerializer to edit only the target sheets (`Order list` and `PO Multivac`), comments, VML shapes, and workbook defined names.
  - All other sheets, drawings, printer settings, logos, and custom XML remain byte-identical.
  - Never uses load-and-save Excel libraries (ExcelJS, SheetJS, openpyxl) that drop or damage custom workbook features.

- **HyperFormula Calculations & Verification**:
  - Automatically evaluates formulas and caches computed values into `<v>`.
  - Comprehensive 4-stage verification before download is unlocked.

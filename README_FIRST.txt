MO32 Crane Compliance - Auto Check (v7.3)

Fixes:
  - "Try demo" no longer writes to session_state; uses a direct sample DataFrame instead.
  - PDF export remains ASCII-safe (no Helvetica Unicode crashes).

Features:
  - Web form (no Excel), photos per crane, contradiction checks, due-soon, PASS/ATTENTION/FAIL.
  - Saves each submission under mo32_cases/<timestamp> with CSV, results, photos, DOCX & PDF.

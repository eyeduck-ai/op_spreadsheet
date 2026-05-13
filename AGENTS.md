# Project Guidance

- This repository contains a Google Apps Script surgery scheduling tool for the `OP` sheet.
- System-managed headers must use English names. User-facing clinical headers may remain Chinese.
- Always locate fields by row-1 header names; do not rely on fixed column letters or positions.
- System columns must be hidden and protected where practical. Current system columns: `CalendarEventId`, `CalendarSheetWriteUpdated`.
- Calendar titles must preserve field slots as `chartNo | patientName | condition`; do not collapse empty middle fields.
- Duplicate `CalendarEventId` values must be treated as a sync diagnostic, not silently resolved to the first row.
- Any change requiring existing sheet data to move, rename, or be reinterpreted must include an explicit migration function and a `維護工具` menu entry.
- Migrations must preserve old data on conflict and report what was copied, removed, or left for manual review.
- Keep `日曆eventID` and `日曆apiEventID` as legacy migration aliases only, not as canonical headers.
- Any change intended to be deployed to GAS must bump `CONFIG.VERSION` in `code.js`; use `YYYY.MM.DD`, or `YYYY.MM.DD.2` and so on for multiple deployed changes on the same day.
- If README or menu examples mention the visible version, update those references in the same change.
- Before deployment, run `node --check code.js`, `node --check tests/smoke.js`, `node tests/smoke.js`, and `git diff --check`.

# Project Guidance

- This repository contains a Google Apps Script surgery scheduling and follow-up tool for `FU` and `YYYYMM` monthly sheets.
- System-managed headers must use English names. User-facing clinical headers may remain Chinese.
- Always locate fields by row-1 header names; do not rely on fixed column letters or positions.
- System columns must be hidden. Do not automatically create, remove, or adjust protected ranges; preserve any owner-managed protection settings. FU and monthly sheets use `CalendarEventId` for Calendar binding and `ScheduleSyncId` for permanent patient/date-header identity. Preserve both columns across native sorts, moves, copies to archive, and header remapping.
- GAS must not call Calendar or CalendarApp; only the designated external worker owns event mutations. Keep legacy trigger entrypoints safe no-ops.
- Keep `tests/fixtures/sheet_contract.json` aligned with the platform's matching contract; run `node tests/sheet_contract.js` before deployment.
- Duplicate `ScheduleSyncId` or `CalendarEventId` values must be treated as a sync diagnostic, not silently resolved to the first row.
- Existing-data identity/header migrations belong to the main program and independent worker's automatic setup. Do not require GAS initialization, migration or handover menus. Keep only historical installed trigger names as harmless compatibility entrypoints; obsolete menus and internal scaffolding should be removed. Optional format repair works on existing sheets only.
- Monthly patient hospital cells are optional manual content. Resolve campus from the nearest valid day header; never automatically fill or overwrite the patient cell. FU hospital fields retain their own semantics.
- `OP` is a legacy sheet name only. Do not automatically rename it; prompt users to manually rename `OP` to `All`.
- Migrations must preserve old data on conflict and report what was copied, removed, or left for manual review.
- Keep `日曆eventID` and `日曆apiEventID` as legacy migration aliases only, not as canonical headers.
- Any change intended to be deployed to GAS must bump `CONFIG.VERSION` in `code.js`; use `YYYY.MM.DD`, or `YYYY.MM.DD.2` and so on for multiple deployed changes on the same day.
- If README or menu examples mention the visible version, update those references in the same change.
- Before deployment, run `node --check code.js`, `node --check tests/smoke.js`, `node tests/smoke.js`, and `git diff --check`.

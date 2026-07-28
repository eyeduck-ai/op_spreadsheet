/**
 * 一次性安全清理：舊遷移備份分頁已完成其用途，且原試算表刻意允許
 * 免登入瀏覽，故不能把隱藏分頁當成隱私邊界。正式刪除前會先建立並
 * 回讀驗證一份私人整檔備份；同時只修剪 FU／月表尾端的空白列。
 */
function fingerprintLegacyBackupSheet_(sheet) {
  const lastRow = Math.max(1, sheet.getLastRow());
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const chunkRows = 200;
  let digest = sha256Hex_(JSON.stringify({
    name: sheet.getName(),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns(),
    lastRow,
    lastColumn
  }));
  for (let row = 1; row <= lastRow; row += chunkRows) {
    const count = Math.min(chunkRows, lastRow - row + 1);
    const range = sheet.getRange(row, 1, count, lastColumn);
    digest = sha256Hex_(digest + JSON.stringify([
      range.getValues(),
      range.getFormulas(),
      range.getNotes()
    ]));
  }
  return digest;
}

function findTrailingManagedGridNoteRows_(sheet, keepRows) {
  const firstRow = Number(keepRows) + 1;
  const maxRows = sheet.getMaxRows();
  if (firstRow > maxRows) return [];
  const maxColumns = Math.max(1, sheet.getMaxColumns());
  const chunkRows = 250;
  const result = [];
  for (let row = firstRow; row <= maxRows; row += chunkRows) {
    const count = Math.min(chunkRows, maxRows - row + 1);
    const notes = sheet.getRange(row, 1, count, maxColumns).getNotes();
    notes.forEach((values, index) => {
      if (values.some(value => toCellText_(value))) {
        result.push(row + index);
      }
    });
  }
  return result;
}

function buildManagedGridTrimPlan_(spreadsheet) {
  const sheets = [];
  const fu = getMainTrackingSheet_(spreadsheet);
  if (fu) sheets.push({ sheet: fu, minimumRows: FU_GRID_RETAIN_ROWS });
  spreadsheet.getSheets()
    .filter(sheet => isMonthlySheetName_(sheet.getName()))
    .forEach(sheet => {
      sheets.push({
        sheet,
        minimumRows: MONTHLY_GRID_RETAIN_ROWS
      });
    });
  return sheets.map(item => {
    const sheet = item.sheet;
    const maxRows = sheet.getMaxRows();
    const lastRow = Math.max(1, sheet.getLastRow());
    const keepRows = Math.min(
      maxRows,
      Math.max(item.minimumRows, lastRow + MANAGED_GRID_ROW_BUFFER)
    );
    const deleteCount = Math.max(0, maxRows - keepRows);
    const noteRows = deleteCount
      ? findTrailingManagedGridNoteRows_(sheet, keepRows)
      : [];
    return {
      sheetId: sheet.getSheetId(),
      sheetName: sheet.getName(),
      maxRows,
      lastRow,
      keepRows,
      deleteCount,
      noteRows
    };
  });
}

function buildLegacyBackupAndGridCleanupPreview_(spreadsheet) {
  const backupSheets = LEGACY_BACKUP_SHEET_NAMES
    .map(name => spreadsheet.getSheetByName(name))
    .filter(Boolean)
    .map(sheet => ({
      sheetId: sheet.getSheetId(),
      sheetName: sheet.getName(),
      hidden: sheet.isSheetHidden(),
      maxRows: sheet.getMaxRows(),
      maxColumns: sheet.getMaxColumns(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      fingerprint: fingerprintLegacyBackupSheet_(sheet)
    }));
  const gridPlans = buildManagedGridTrimPlan_(spreadsheet);
  const blockingNotes = gridPlans
    .filter(plan => plan.noteRows.length)
    .map(plan => ({
      sheetName: plan.sheetName,
      rows: plan.noteRows
    }));
  const fingerprint = sha256Hex_(JSON.stringify({
    spreadsheetId: spreadsheet.getId(),
    backupSheets,
    gridPlans: gridPlans.map(plan => ({
      sheetId: plan.sheetId,
      sheetName: plan.sheetName,
      maxRows: plan.maxRows,
      lastRow: plan.lastRow,
      keepRows: plan.keepRows,
      deleteCount: plan.deleteCount,
      noteRows: plan.noteRows
    }))
  }));
  return {
    version: 1,
    spreadsheetId: spreadsheet.getId(),
    createdAt: new Date().toISOString(),
    backupSheets,
    gridPlans,
    blockingNotes,
    fingerprint,
    deleteSheetCount: backupSheets.length,
    deleteGridRowCount: gridPlans.reduce(
      (total, plan) => total + plan.deleteCount,
      0
    )
  };
}

function formatLegacyBackupAndGridCleanupPreview_(preview) {
  const backupLines = preview.backupSheets.length
    ? preview.backupSheets.map(item => {
      return `• ${item.sheetName}（${item.maxRows}×${item.maxColumns}）`;
    })
    : ['• 沒有待移除的舊備份分頁'];
  const trimLines = preview.gridPlans
    .filter(plan => plan.deleteCount)
    .map(plan => {
      return `• ${plan.sheetName}：${plan.maxRows} → ` +
        `${plan.keepRows} 列（移除 ${plan.deleteCount} 個尾端空白列）`;
    });
  const blockedLines = preview.blockingNotes.map(item => {
    return `• ${item.sheetName} 的尾端空白區仍有 note：第 ` +
      `${item.rows.join('、')} 列`;
  });
  return [
    `待移除舊備份分頁：${preview.deleteSheetCount} 個`,
    ...backupLines,
    '',
    `可修剪尾端空白列：${preview.deleteGridRowCount} 列`,
    ...(trimLines.length ? trimLines : ['• 不需要修剪']),
    ...(blockedLines.length
      ? ['', '阻擋項目（不會執行）：', ...blockedLines]
      : []),
    '',
    '執行時會先建立私人整檔備份並回讀驗證；不呼叫 Calendar API。'
  ].join('\n');
}

function previewLegacyBackupAndGridCleanup() {
  return runMenuAction_('預覽舊備份分頁與空白列清理', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const preview = buildLegacyBackupAndGridCleanupPreview_(spreadsheet);
    PropertiesService.getScriptProperties().setProperty(
      LEGACY_BACKUP_CLEANUP_PREVIEW_PROPERTY,
      JSON.stringify({
        version: preview.version,
        spreadsheetId: preview.spreadsheetId,
        createdAt: preview.createdAt,
        fingerprint: preview.fingerprint
      })
    );
    SpreadsheetApp.getUi().alert(
      '舊備份分頁與空白列清理預覽',
      formatLegacyBackupAndGridCleanupPreview_(preview),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return {
      ok: !preview.blockingNotes.length,
      ...preview
    };
  });
}

function assertLegacyCleanupBackupMatches_(backupSpreadsheet, preview) {
  preview.backupSheets.forEach(expected => {
    const backupSheet = backupSpreadsheet.getSheetByName(expected.sheetName);
    if (!backupSheet) {
      throw new Error(`私人整檔備份缺少 ${expected.sheetName}。`);
    }
    if (
      backupSheet.getMaxRows() !== expected.maxRows ||
      backupSheet.getMaxColumns() !== expected.maxColumns ||
      backupSheet.getLastRow() !== expected.lastRow ||
      backupSheet.getLastColumn() !== expected.lastColumn ||
      fingerprintLegacyBackupSheet_(backupSheet) !== expected.fingerprint
    ) {
      throw new Error(
        `${expected.sheetName} 在私人整檔備份中的回讀指紋不一致。`
      );
    }
  });
  return true;
}

function executeLegacyBackupAndGridCleanup() {
  return runMenuAction_('執行舊備份分頁與空白列清理', () => {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const stored = safeJsonParse_(
      PropertiesService.getScriptProperties().getProperty(
        LEGACY_BACKUP_CLEANUP_PREVIEW_PROPERTY
      ),
      null
    );
    if (!stored || stored.spreadsheetId !== spreadsheet.getId()) {
      throw new Error('尚未建立有效預覽；請先執行預覽。');
    }
    const beforeLock = buildLegacyBackupAndGridCleanupPreview_(spreadsheet);
    if (beforeLock.blockingNotes.length) {
      throw new Error(
        '尾端待修剪區域仍含 note；為避免資料遺失，本次不執行。'
      );
    }
    if (stored.fingerprint !== beforeLock.fingerprint) {
      throw new Error(
        '試算表在預覽後已有變動；請重新執行預覽後再確認。'
      );
    }
    if (!beforeLock.deleteSheetCount && !beforeLock.deleteGridRowCount) {
      ui.alert('目前沒有需要清理的舊備份分頁或尾端空白列。');
      return { ok: true, noOp: true };
    }
    const response = ui.alert(
      '確認安全清理',
      formatLegacyBackupAndGridCleanupPreview_(beforeLock) +
        '\n\n確認建立私人整檔備份後執行嗎？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return { ok: false, cancelled: true };
    }

    return withCalendarSyncLock_(() => {
      const preview = buildLegacyBackupAndGridCleanupPreview_(spreadsheet);
      if (
        preview.blockingNotes.length ||
        preview.fingerprint !== stored.fingerprint
      ) {
        throw new Error(
          '資料在取得鎖之前已變動；未建立備份或刪除任何內容，請重新預覽。'
        );
      }
      const backup = createPrivateSpreadsheetBackup_(spreadsheet, []);
      PropertiesService.getScriptProperties().setProperty(
        LEGACY_BACKUP_CLEANUP_BACKUP_PROPERTY,
        JSON.stringify({
          status: 'BACKUP_VERIFIED',
          fileId: backup.fileId,
          name: backup.name,
          createdAt: new Date().toISOString(),
          sourceFingerprint: preview.fingerprint
        })
      );
      const backupSpreadsheet = SpreadsheetApp.openById(backup.fileId);
      assertLegacyCleanupBackupMatches_(backupSpreadsheet, preview);

      preview.backupSheets.forEach(item => {
        const sheet = spreadsheet.getSheetByName(item.sheetName);
        if (sheet) spreadsheet.deleteSheet(sheet);
      });
      preview.gridPlans.forEach(plan => {
        if (!plan.deleteCount) return;
        const sheet = getSheetById_(spreadsheet, plan.sheetId);
        if (!sheet || sheet.getName() !== plan.sheetName) {
          throw new Error(`${plan.sheetName} 在修剪前已被移動或刪除。`);
        }
        sheet.deleteRows(plan.keepRows + 1, plan.deleteCount);
      });
      SpreadsheetApp.flush();

      const remainingBackups = LEGACY_BACKUP_SHEET_NAMES.filter(name => {
        return Boolean(spreadsheet.getSheetByName(name));
      });
      if (remainingBackups.length) {
        throw new Error(
          `仍有舊備份分頁未移除：${remainingBackups.join('、')}。`
        );
      }
      preview.gridPlans.forEach(plan => {
        const sheet = getSheetById_(spreadsheet, plan.sheetId);
        if (sheet && sheet.getMaxRows() !== plan.keepRows) {
          throw new Error(`${plan.sheetName} 空白列修剪回讀驗證失敗。`);
        }
      });
      PropertiesService.getScriptProperties().setProperty(
        LEGACY_BACKUP_CLEANUP_BACKUP_PROPERTY,
        JSON.stringify({
          status: 'COMPLETE',
          fileId: backup.fileId,
          name: backup.name,
          completedAt: new Date().toISOString(),
          sourceFingerprint: preview.fingerprint,
          removedSheets: preview.backupSheets.map(item => item.sheetName),
          removedGridRows: preview.deleteGridRowCount
        })
      );
      PropertiesService.getScriptProperties()
        .deleteProperty(LEGACY_BACKUP_CLEANUP_PREVIEW_PROPERTY);
      ui.alert(
        '安全清理完成',
        `私人整檔備份：${backup.name}\n` +
          `已移除 ${preview.deleteSheetCount} 個舊備份分頁，並修剪 ` +
          `${preview.deleteGridRowCount} 個尾端空白列。\n` +
          'Calendar 事件與同步索引均未改寫。',
        ui.ButtonSet.OK
      );
      return {
        ok: true,
        backup,
        removedSheets: preview.backupSheets.map(item => item.sheetName),
        removedGridRows: preview.deleteGridRowCount
      };
    });
  });
}

/** Durable two-phase GAS archive. Full plans live only in private Drive files. */
function readNativeArchiveBook_(spreadsheetId) {
  const book = Sheets.Spreadsheets.get(spreadsheetId, { fields: 'spreadsheetId,properties,sheets(properties,protectedRanges,charts,conditionalFormats),developerMetadata,namedRanges' });
  book.developerMetadata = book.developerMetadata || [];
  (book.sheets || []).forEach(sheet => {
    sheet.rows = []; sheet.merges = []; sheet.column_widths = []; sheet.row_metadata = [];
    const grid = sheet.properties.gridProperties;
    if (!grid) return;
    if (grid.columnCount > 1000) throw new Error('分頁超過 1,000 欄，請先確認刀表。');
    const title = sheet.properties.title.replace(/'/g, "''"), last = columnToLetter_(grid.columnCount);
    for (let start = 0; start < grid.rowCount; start += 400) {
      const response = Sheets.Spreadsheets.get(spreadsheetId, { ranges: [`'${title}'!A${start + 1}:${last}${Math.min(start + 400, grid.rowCount)}`],
        fields: 'sheets(merges,data(startRow,startColumn,columnMetadata(pixelSize,hiddenByUser),rowMetadata(pixelSize,hiddenByUser),rowData(values(userEnteredValue,effectiveValue,formattedValue,userEnteredFormat,effectiveFormat,dataValidation,note,textFormatRuns))))' });
      const part = (response.sheets || [])[0] || {};
      sheet.merges = part.merges || [];
      (part.data || []).forEach(block => {
        const offset = block.startRow || 0;
        (block.rowData || []).forEach((row, index) => { while (sheet.rows.length <= offset + index) sheet.rows.push([]); sheet.rows[offset + index] = row.values || []; });
        if (block.columnMetadata && !sheet.column_widths.length) sheet.column_widths = block.columnMetadata;
        (block.rowMetadata || []).forEach((value, index) => { while (sheet.row_metadata.length <= offset + index) sheet.row_metadata.push({}); sheet.row_metadata[offset + index] = value; });
      });
    }
  });
  return book;
}
function nativeArchiveLock_(book) {
  const matches = (book.developerMetadata || []).filter(item => item.metadataKey === NATIVE_ARCHIVE_LOCK_KEY || item.metadataId === NATIVE_ARCHIVE_LOCK_ID);
  if (matches.length > 1 || matches.length && (matches[0].metadataKey !== NATIVE_ARCHIVE_LOCK_KEY || matches[0].metadataId !== NATIVE_ARCHIVE_LOCK_ID)) throw new Error('封存互斥標記衝突，保留所有標記供核對。');
  return matches.length ? matches[0].metadataValue : '';
}
function assertNativeArchiveAvailable_(owner, book) {
  const current = book || Sheets.Spreadsheets.get(SpreadsheetApp.getActiveSpreadsheet().getId(), { fields: 'developerMetadata' });
  const active = nativeArchiveLock_(current);
  if (active && active !== owner) throw new Error('有未完成封存，請回原封存入口核對並完成；目前暫停自動寫入。');
  return active;
}
function nativeLockRequest_(operationId, remove) {
  const metadata = { metadataId: NATIVE_ARCHIVE_LOCK_ID, metadataKey: NATIVE_ARCHIVE_LOCK_KEY, metadataValue: operationId, visibility: 'DOCUMENT' };
  return remove ? { deleteDeveloperMetadata: { dataFilter: { developerMetadataLookup: metadata } } }
    : { createDeveloperMetadata: { developerMetadata: { ...metadata, location: { spreadsheet: true } } } };
}
function nativeWitnessRequest_(operationId, phase, remove) {
  const metadata = { metadataKey: NATIVE_ARCHIVE_WITNESS, metadataValue: `${operationId}:${phase}`, visibility: 'DOCUMENT' };
  return remove ? { deleteDeveloperMetadata: { dataFilter: { developerMetadataLookup: metadata } } }
    : { createDeveloperMetadata: { developerMetadata: { ...metadata, location: { spreadsheet: true } } } };
}
function nativePrivateFile_(file) {
  if (typeof file.setShareableByEditors === 'function') file.setShareableByEditors(false);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  file.getEditors().forEach(user => file.removeEditor(user)); file.getViewers().forEach(user => file.removeViewer(user));
  verifyPrivateBackup_(file); return file;
}
function nativeLoadPlan_(transaction) {
  const file = DriveApp.getFileById(transaction.planFileId); verifyPrivateBackup_(file);
  const blob = file.getName().endsWith('.gz') ? Utilities.ungzip(file.getBlob()) : file.getBlob();
  const plan = JSON.parse(blob.getDataAsString('UTF-8'));
  if (plan.transaction_id !== transaction.operationId) throw new Error('私人封存計畫識別碼不一致。');
  return plan;
}
function nativePublicArchive_(transaction) {
  const summary = transaction.summary || nativeLoadPlan_(transaction).summary;
  const changes = transaction.changes || summary.months.map(item => `${item.month} → ${item.target}：${item.patients} 筆排程、${item.blocks} 個刀日${item.backfill ? '；補入較早月份' : ''}${item.migrated_existing_format ? '；備份並固定既有條件格式，舊取消狀態仍待核對' : ''}`);
  return { ok: !['blocked', 'uncertain'].includes(transaction.status), operationId: transaction.operationId,
    status: transaction.status, months: transaction.months, summary,
    changes, message: transaction.message || '',
    canCopy: ['preview', 'locked', 'backup_verified'].includes(transaction.status), canFinalize: transaction.status === 'copied',
    canCancel: !transaction.copyAttempted && transaction.status !== 'applied' };
}
function nativeSaveArchive_(transaction, status, message) {
  transaction.status = status; transaction.message = message; transaction.updatedAt = new Date().toISOString();
  writeAnnualArchiveTransaction_(transaction); return nativePublicArchive_(transaction);
}
function nativeAssertOriginal_(plan, book) {
  const find = title => book.sheets.find(sheet => sheet.properties.title === title);
  if (Object.keys(plan.sources).some(title => !find(title) || !nativeSame_(plan.sources[title], find(title))) ||
      Object.keys(plan.original_targets).some(title => plan.original_targets[title] ? !find(title) || !nativeSame_(plan.original_targets[title], find(title)) : Boolean(find(title)))) throw new Error('來源或既有封存已改變，未寫入，請重新核對預覽。');
  nativeFormulaChecks_(book, Object.keys(plan.sources)); nativeAssertIdentities_(book);
}
function previewNativeArchive_(spreadsheet, payload) {
  const previous = readAnnualArchiveTransaction_();
  if (previous && previous.version !== 2) throw new Error('有舊版未完成交易，保留來源與私人備份；請先人工核對，不自動續跑舊版刪除流程。');
  if (previous && !['preview', 'applied', 'rejected'].includes(previous.status)) return nativePublicArchive_(previous);
  const book = readNativeArchiveBook_(spreadsheet.getId()); assertNativeArchiveAvailable_('', book);
  assertNoLegacyArchiveState_();
  const operationId = Utilities.getUuid(), plan = planNativeArchive_(book, payload.months, formatMonthKey_(new Date()), operationId);
  // Native formats are repetitive; compress private journals and omit the duplicate full book.
  const savedPlan = { ...plan }; delete savedPlan.book;
  const blob = Utilities.newBlob(JSON.stringify(savedPlan), 'application/json', `刀表封存計畫_${operationId}.json`);
  const file = nativePrivateFile_(DriveApp.createFile(Utilities.gzip(blob)));
  const transaction = { version: 2, operationId, spreadsheetId: spreadsheet.getId(), months: Object.keys(plan.sources),
    status: 'preview', planFileId: file.getId(), backupFileId: '', summary: plan.summary, startedAt: new Date().toISOString(),
    changes: plan.summary.months.map(item => `${item.month} → ${item.target}：${item.patients} 筆排程、${item.blocks} 個刀日${item.backfill ? '；補入較早月份' : ''}${item.migrated_existing_format ? '；備份並固定既有條件格式，舊取消狀態仍待核對' : ''}`) };
  return nativeSaveArchive_(transaction, 'preview', '預覽已保存；下一步先建立私人整檔備份，再複製並核對。原月份仍保留。');
}
function checkNativeArchive_(spreadsheet, transaction) {
  const plan = nativeLoadPlan_(transaction), book = readNativeArchiveBook_(spreadsheet.getId());
  const marker = phase => (book.developerMetadata || []).some(item => item.metadataKey === NATIVE_ARCHIVE_WITNESS && item.metadataValue === `${transaction.operationId}:${phase}`);
  if (transaction.phase === 'delete' && (marker('delete') || transaction.deleteConfirmed)) {
    if (!verifyNativeArchive_(plan, book, true)) return nativeSaveArchive_(transaction, 'blocked', '移除批次已送達，但封存內容或結構變動；保留備份供核對。');
    transaction.deleteConfirmed = true; nativeSaveArchive_(transaction, 'applied', '封存完成，原月份已移除並回讀核對。');
    try { Sheets.Spreadsheets.batchUpdate({ requests: ['copy', 'delete'].map(phase => nativeWitnessRequest_(transaction.operationId, phase, true)) }, spreadsheet.getId()); } catch (_) { /* confirmed receipts remain durable */ }
    return nativePublicArchive_(transaction);
  }
  if (marker('copy') || transaction.copyConfirmed) {
    if (nativeArchiveLock_(book) !== transaction.operationId || !verifyNativeArchive_(plan, book, false)) return nativeSaveArchive_(transaction, 'blocked', '來源、封存副本或互斥標記不一致，來源月份保留。');
    transaction.copyConfirmed = true;
    return nativeSaveArchive_(transaction, 'copied', '封存副本與私人備份已核對。請再次確認後才移除原月份；並存期間同步暫停。');
  }
  if (!transaction.copyAttempted && nativeArchiveLock_(book) === transaction.operationId) return nativeSaveArchive_(transaction, 'locked', '已取得封存互斥標記；可明確繼續建立備份與複製，尚未送出複製批次。');
  return nativeSaveArchive_(transaction, 'uncertain', '尚無法確認批次完成標記；未重送或解除互斥標記，來源保留供核對。');
}
function copyNativeArchive_(spreadsheet, transaction) {
  if (!['preview', 'locked', 'backup_verified'].includes(transaction.status)) return checkNativeArchive_(spreadsheet, transaction);
  const plan = nativeLoadPlan_(transaction);
  let book = readNativeArchiveBook_(spreadsheet.getId()); assertNativeArchiveAvailable_(transaction.operationId, book);
  nativeAssertOriginal_(plan, book);
  if (transaction.status === 'preview') {
    nativeSaveArchive_(transaction, 'uncertain', '正在取得互斥標記；中斷後請核對原交易。');
    Sheets.Spreadsheets.batchUpdate({ requests: [nativeLockRequest_(transaction.operationId, false)] }, spreadsheet.getId());
    nativeSaveArchive_(transaction, 'locked', '已取得封存互斥標記，正在保存私人備份。');
  }
  if (!transaction.backupFileId) {
    const file = DriveApp.getFileById(spreadsheet.getId()).makeCopy(`手術排程系統_Backup_${transaction.operationId}`);
    transaction.backupFileId = file.getId(); transaction.backupName = file.getName(); writeAnnualArchiveTransaction_(transaction);
    nativePrivateFile_(file);
  }
  verifyPrivateBackup_(DriveApp.getFileById(transaction.backupFileId));
  const backup = readNativeArchiveBook_(transaction.backupFileId);
  if (JSON.stringify(backup.sheets.map(sheet => sheet.properties.title).sort()) !== JSON.stringify(book.sheets.map(sheet => sheet.properties.title).sort())) throw new Error('私人整檔備份缺少分頁，來源保留。');
  [...Object.values(plan.sources), ...Object.values(plan.original_targets).filter(Boolean)].forEach(source => {
    const copied = backup.sheets.find(sheet => sheet.properties.title === source.properties.title);
    if (!copied || !nativeSame_(source, copied)) throw new Error('私人整檔備份原生內容回讀不一致，來源保留。');
  });
  nativeSaveArchive_(transaction, 'backup_verified', '私人整檔備份已核對，正在複製年度封存。');
  book = readNativeArchiveBook_(spreadsheet.getId()); nativeAssertOriginal_(plan, book);
  if (nativeArchiveLock_(book) !== transaction.operationId) throw new Error('封存互斥標記遺失，未複製。');
  transaction.copyAttempted = true; transaction.phase = 'copy'; nativeSaveArchive_(transaction, 'uncertain', '複製批次已保存，正在送出；中斷時只核對，不重送。');
  Sheets.Spreadsheets.batchUpdate({ requests: [...plan.requests, nativeWitnessRequest_(transaction.operationId, 'copy', false)] }, spreadsheet.getId());
  return checkNativeArchive_(spreadsheet, transaction);
}
function finalizeNativeArchive_(spreadsheet, transaction, payload) {
  if (transaction.phase === 'delete') return checkNativeArchive_(spreadsheet, transaction);
  if (transaction.status !== 'copied' || payload.confirmedPermanentDeletion !== true) throw new Error('請先核對封存副本，再明確確認移除原月份。');
  const plan = nativeLoadPlan_(transaction), book = readNativeArchiveBook_(spreadsheet.getId());
  if (nativeArchiveLock_(book) !== transaction.operationId || !verifyNativeArchive_(plan, book, false)) throw new Error('封存副本或來源已改變，未移除來源。');
  nativeFormulaChecks_(book, transaction.months);
  const requests = Object.values(plan.sources).map(source => ({ deleteSheet: { sheetId: source.properties.sheetId } }));
  requests.push(nativeWitnessRequest_(transaction.operationId, 'delete', false), nativeLockRequest_(transaction.operationId, true));
  transaction.phase = 'delete'; nativeSaveArchive_(transaction, 'uncertain', '移除批次已保存，正在送出；中斷時只核對，不重送。');
  Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheet.getId());
  return checkNativeArchive_(spreadsheet, transaction);
}
function cancelUncopiedNativeArchive_(spreadsheet, transaction) {
  const book = readNativeArchiveBook_(spreadsheet.getId());
  assertNativeArchiveAvailable_(transaction.operationId, book);
  if (transaction.copyAttempted || (book.developerMetadata || []).some(item => item.metadataKey === NATIVE_ARCHIVE_WITNESS && item.metadataValue === `${transaction.operationId}:copy`)) throw new Error('複製批次已送出或結果不明；只可核對，不能取消或解除互斥。');
  if (nativeArchiveLock_(book) === transaction.operationId) {
    Sheets.Spreadsheets.batchUpdate({ requests: [nativeLockRequest_(transaction.operationId, true)] }, spreadsheet.getId());
    if (nativeArchiveLock_(readNativeArchiveBook_(spreadsheet.getId())) === transaction.operationId) throw new Error('取消互斥標記尚未確認，請核對原交易。');
  }
  return nativeSaveArchive_(transaction, 'rejected', '尚未送出複製批次，已明確取消；來源與私人備份保留，可重新預覽。');
}
function submitNativeArchive_(payload) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(), action = payload.action || 'preview';
  return withSheetOperationLock_(() => {
    if (action === 'preview') return previewNativeArchive_(spreadsheet, payload);
    const transaction = readAnnualArchiveTransaction_();
    if (!transaction || transaction.version !== 2 || transaction.spreadsheetId !== spreadsheet.getId() || transaction.operationId !== payload.operationId) throw new Error('請開啟原封存交易核對，不建立另一批重複封存。');
    try {
      if (action === 'check') return checkNativeArchive_(spreadsheet, transaction);
      if (action === 'cancel') return cancelUncopiedNativeArchive_(spreadsheet, transaction);
      if (action === 'copy') return copyNativeArchive_(spreadsheet, transaction);
      if (action === 'finalize') return finalizeNativeArchive_(spreadsheet, transaction, payload);
      throw new Error('封存操作不存在。');
    } catch (error) {
      return nativeSaveArchive_(transaction, transaction.copyAttempted || transaction.phase === 'delete' ? 'uncertain' : transaction.status,
        `${error.message || error}\n原交易、備份與來源保留；請先核對，未自動重送。`);
    }
  }, payload.operationId || '');
}

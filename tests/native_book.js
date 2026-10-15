'use strict';
// Executes native requests; never replaces the workbook with a planner's result.
const clone = value => JSON.parse(JSON.stringify(value));
function execute(book, requests) {
  const next = clone(book), sheet = id => next.sheets.find(item => item.properties.sheetId === id);
  next.developerMetadata ||= [];
  for (const request of requests) {
    if (request.addSheet) {
      const p = request.addSheet.properties;
      if (next.sheets.some(item => item.properties.sheetId === p.sheetId || item.properties.title === p.title)) throw new Error('duplicate sheet');
      next.sheets.push({ properties: clone(p), rows: [], column_widths: [], row_metadata: [], merges: [] });
    } else if (request.deleteSheet) next.sheets = next.sheets.filter(item => item.properties.sheetId !== request.deleteSheet.sheetId);
    else if (request.createDeveloperMetadata) {
      const item = clone(request.createDeveloperMetadata.developerMetadata);
      if (item.metadataId === undefined) item.metadataId = Math.max(0, ...next.developerMetadata.filter(m => m.metadataId < 100000).map(m => m.metadataId)) + 1;
      if (next.developerMetadata.some(m => m.metadataId === item.metadataId)) throw new Error('duplicate metadata');
      next.developerMetadata.push(item);
    } else if (request.deleteDeveloperMetadata) {
      const lookup = request.deleteDeveloperMetadata.dataFilter.developerMetadataLookup;
      next.developerMetadata = next.developerMetadata.filter(item => !Object.entries(lookup).every(([key, value]) => item[key] === value));
    } else if (request.appendDimension) {
      const item = request.appendDimension;
      sheet(item.sheetId).properties.gridProperties[item.dimension === 'ROWS' ? 'rowCount' : 'columnCount'] += item.length;
    } else if (request.insertDimension) {
      const range = request.insertDimension.range, target = sheet(range.sheetId), count = range.endIndex - range.startIndex;
      target.rows.splice(range.startIndex, 0, ...Array.from({ length: count }, () => []));
      target.row_metadata ||= [];
      while (target.row_metadata.length < range.startIndex) target.row_metadata.push({});
      target.row_metadata.splice(range.startIndex, 0, ...Array.from({ length: count }, () => ({})));
      target.properties.gridProperties.rowCount += count;
      const ranges = [...(target.merges || [])];
      (target.protectedRanges || []).forEach(item => ranges.push(item.range || {}, ...(item.unprotectedRanges || [])));
      ranges.forEach(item => { if ((item.startRowIndex || 0) >= range.startIndex) item.startRowIndex = (item.startRowIndex || 0) + count; if ((item.endRowIndex || 0) > range.startIndex) item.endRowIndex += count; });
    } else if (request.updateDimensionProperties) {
      const { range, properties } = request.updateDimensionProperties, target = sheet(range.sheetId), key = range.dimension === 'ROWS' ? 'row_metadata' : 'column_widths';
      target[key] ||= [];
      while (target[key].length < range.endIndex) target[key].push({});
      for (let index = range.startIndex; index < range.endIndex; index++) Object.assign(target[key][index], clone(properties));
    } else if (request.updateCells) {
      const { range, rows, fields } = request.updateCells, target = sheet(range.sheetId);
      while (target.rows.length < range.endRowIndex) target.rows.push([]);
      rows.forEach((row, offset) => row.values.forEach((value, column) => {
        const index = column + (range.startColumnIndex || 0), line = target.rows[range.startRowIndex + offset];
        while (line.length <= index) line.push({});
        for (const field of fields.split(',')) { if (field in value) line[index][field] = clone(value[field]); else delete line[index][field]; }
        if (fields.includes('userEnteredValue')) { delete line[index].formattedValue; delete line[index].effectiveValue; }
      }));
    } else if (request.mergeCells) (sheet(request.mergeCells.range.sheetId).merges ||= []).push(clone(request.mergeCells.range));
    else if (request.deleteConditionalFormatRule) sheet(request.deleteConditionalFormatRule.sheetId).conditionalFormats.splice(request.deleteConditionalFormatRule.index, 1);
    else throw new Error('Unsupported request: ' + Object.keys(request));
  }
  return next;
}
module.exports = { execute, clone };

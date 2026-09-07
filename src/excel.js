import ExcelJS from 'exceljs';
import { invoke } from '@tauri-apps/api/core';
import { readFile, writeFile } from '@tauri-apps/plugin-fs';

async function readFileBytes(path) {
  return await readFile(path);
}

async function writeToFile(path, uint8array) {
  await writeFile(path, uint8array);
  return path;
}

const THIN_BORDER = {
  top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
};

const SEVERITY_STYLES = {
  critical: { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9F1239' } } },
  high: { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE11D48' } } },
  medium: { font: { bold: true, color: { argb: 'FF451A03' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } } },
  low: { font: { bold: true, color: { argb: 'FF022C22' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34D399' } } },
  unknown: { font: { color: { argb: 'FF64748B' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } } },
};

const EOL_STYLES = {
  '已EOL': { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE11D48' } } },
  '维护中': { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF059669' } } },
  '即将EOL': { font: { bold: true, color: { argb: 'FF451A03' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } } },
  'AI判定': { font: { bold: true, color: { argb: 'FF4F46E5' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF2FF' } } },
  '未知': { font: { color: { argb: 'FF64748B' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } } },
  '待研判': { font: { bold: true, color: { argb: 'FF6366F1' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } } },
};

const REPORT_FONT = { name: '宋体', size: 11 };

const REPORT_BORDER = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
};

const REPORT_BORDER_NO_TOP = {
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
};

const REPORT_COL_WIDTHS = [30.25, 66.0, 46.0, 16.25, 19.125, 12.875, 8.875, 31.625];

function applyHeaderStyle(cell) {
  cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = THIN_BORDER;
}

function applyDataStyle(cell, opts) {
  opts = opts || {};
  cell.border = THIN_BORDER;
  cell.font = { name: 'Arial', size: 10, bold: !!opts.bold };
  cell.alignment = { horizontal: opts.align || 'left', vertical: 'middle', wrapText: !!opts.wrap };
  if (opts.fontColor) cell.font.color = { argb: opts.fontColor };
}

function applySeverityStyle(cell, severity) {
  const style = SEVERITY_STYLES[severity] || SEVERITY_STYLES.unknown;
  cell.font = { name: 'Arial', size: 10, ...style.font };
  cell.fill = style.fill;
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = THIN_BORDER;
}

function applyEOLStyle(cell, status) {
  const style = EOL_STYLES[status] || EOL_STYLES['未知'];
  cell.font = { name: 'Arial', size: 10, ...style.font };
  cell.fill = style.fill;
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = THIN_BORDER;
}

function normalizeEOLStatus(status) {
  if (!status) return '未知';
  if (status.includes('EOL') && !status.includes('维护') && !status.includes('即将')) return '已EOL';
  if (status.includes('维护')) return '维护中';
  if (status.includes('即将')) return '即将EOL';
  if (status.includes('AI')) return 'AI判定';
  return '未知';
}

export async function importFromExcel(filePath) {
  const bytes = await readFileBytes(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);

  let sheet = workbook.getWorksheet('检出组件信息');
  if (!sheet) {
    workbook.eachWorksheet(function (ws) {
      if (sheet) return;
      const headerRow = ws.getRow(1);
      let hasName = false, hasVersion = false;
      headerRow.eachCell(function (cell) {
        const text = String(cell.value || '').trim();
        if (text.includes('组件名称') || (text.includes('组件') && text.includes('名称'))) hasName = true;
        if (text.includes('版本号') || text.includes('版本')) hasVersion = true;
      });
      if (hasName && hasVersion) sheet = ws;
    });
  }
  if (!sheet) sheet = workbook.getWorksheet(1);
  if (!sheet) throw new Error('Excel 文件中没有工作表');

  const headerRow = sheet.getRow(1);
  let nameCol = null, versionCol = null;
  headerRow.eachCell(function (cell, colNumber) {
    const text = String(cell.value || '').trim();
    if (text === '组件名称') nameCol = colNumber;
    else if (text === '版本号') versionCol = colNumber;
  });

  if (!nameCol) {
    let hasGenericName = false;
    headerRow.eachCell(function (cell, colNumber) {
      const text = String(cell.value || '').trim();
      if (text.includes('组件') || text.includes('名称') || text.includes('name') || text.includes('Component')) {
        nameCol = colNumber;
        hasGenericName = true;
      } else if (text.includes('版本') || text.includes('version') || text.includes('Version')) {
        versionCol = colNumber;
      }
    });
    if (!hasGenericName) nameCol = 4;
  }
  if (!versionCol) versionCol = 5;

  const items = [];
  for (let i = 2; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);
    const name = String(row.getCell(nameCol).value || '').trim();
    const version = String(row.getCell(versionCol).value || '').trim();
    if (!name || !version) continue;
    items.push({ name: name, version: version, vendor: '' });
  }

  return items;
}

const DATA_SHEET_KEYWORDS = ['检出组件信息', '检出路径列表', '漏洞信息', '许可证', '敏感信息', '合规风险', '风险文件'];

function isDataSheet(sheetName) {
  for (let i = 0; i < DATA_SHEET_KEYWORDS.length; i++) {
    if (sheetName.indexOf(DATA_SHEET_KEYWORDS[i]) !== -1) return true;
  }
  return false;
}

function detectHeaderRow(sheet) {
  for (let r = 1; r <= Math.min(3, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    let textCount = 0;
    let hasComponentField = false;
    row.eachCell(function (cell) {
      const val = String(cell.value || '').trim();
      if (val.length > 0) textCount++;
      if (val.indexOf('组件') !== -1 || val.indexOf('版本') !== -1 || val.indexOf('名称') !== -1) hasComponentField = true;
    });
    if (textCount >= 3 && hasComponentField) return r;
  }
  return 1;
}

export async function mergeExcelFiles(filePaths, outputPath) {
  if (!filePaths || filePaths.length === 0) throw new Error('请至少选择一个 Excel 文件');
  if (filePaths.length === 1) throw new Error('请选择至少两个 Excel 文件进行合并');

  const firstBytes = await readFileBytes(filePaths[0]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(firstBytes);

  const mergeStats = {};

  for (let f = 1; f < filePaths.length; f++) {
    const bytes = await readFileBytes(filePaths[f]);
    const sourceWb = new ExcelJS.Workbook();
    await sourceWb.xlsx.load(bytes);

    for (let s = 0; s < workbook.worksheets.length; s++) {
      const targetSheet = workbook.worksheets[s];
      const sheetName = targetSheet.name;
      if (!isDataSheet(sheetName)) continue;

      const sourceSheet = sourceWb.getWorksheet(sheetName);
      if (!sourceSheet) continue;

      const headerRow = detectHeaderRow(targetSheet);
      const srcHeaderRow = detectHeaderRow(sourceSheet);

      if (!mergeStats[sheetName]) mergeStats[sheetName] = { targetHeader: headerRow, addedRows: 0, files: 0 };
      mergeStats[sheetName].files++;

      const targetStartRow = targetSheet.rowCount + 1;
      let added = 0;

      for (let r = srcHeaderRow + 1; r <= sourceSheet.rowCount; r++) {
        const srcRow = sourceSheet.getRow(r);
        let hasData = false;
        srcRow.eachCell(function (cell) {
          if (cell.value !== null && cell.value !== undefined && String(cell.value).trim() !== '') hasData = true;
        });
        if (!hasData) continue;

        const targetRowNum = targetStartRow + added;
        const targetRow = targetSheet.getRow(targetRowNum);

        srcRow.eachCell(function (cell, colNumber) {
          const targetCell = targetRow.getCell(colNumber);
          if (cell.type === ExcelJS.ValueType.Formula) {
            targetCell.value = cell.result;
          } else {
            targetCell.value = cell.value;
          }
          if (cell.style) {
            targetCell.font = cell.font ? Object.assign({}, cell.font) : targetCell.font;
            targetCell.fill = cell.fill ? Object.assign({}, cell.fill) : targetCell.fill;
            targetCell.border = cell.border ? Object.assign({}, cell.border) : targetCell.border;
            targetCell.alignment = cell.alignment ? Object.assign({}, cell.alignment) : targetCell.alignment;
          }
        });
        targetRow.commit();
        added++;
      }
      mergeStats[sheetName].addedRows += added;
    }
  }

  const componentList = buildMergeReportSheet(workbook);

  const buffer = await workbook.xlsx.writeBuffer();
  const uint8 = new Uint8Array(buffer);
  await writeToFile(outputPath, uint8);

  return {
    totalFiles: filePaths.length,
    sheets: Object.keys(mergeStats).length,
    stats: mergeStats,
    components: componentList,
  };
}

function parseVulnCount(text) {
  const result = { total: 0, critical: 0, high: 0, medium: 0, low: 0 };
  if (!text) return result;

  const totalMatch = text.match(/^(\d+)/);
  if (totalMatch) result.total = parseInt(totalMatch[1], 10);

  const critMatch = text.match(/严重[：:]\s*(\d+)/);
  if (critMatch) result.critical = parseInt(critMatch[1], 10);

  const highMatch = text.match(/高危[：:]\s*(\d+)/);
  if (highMatch) result.high = parseInt(highMatch[1], 10);

  const medMatch = text.match(/中危[：:]\s*(\d+)/);
  if (medMatch) result.medium = parseInt(medMatch[1], 10);

  const lowMatch = text.match(/低危[：:]\s*(\d+)/);
  if (lowMatch) result.low = parseInt(lowMatch[1], 10);

  if (result.total === 0 && (result.critical + result.high + result.medium + result.low > 0)) {
    result.total = result.critical + result.high + result.medium + result.low;
  }

  return result;
}

function buildVulnsArray(parsed) {
  const arr = [];
  for (let i = 0; i < parsed.critical; i++) arr.push({ severity: 'critical' });
  for (let i = 0; i < parsed.high; i++) arr.push({ severity: 'high' });
  for (let i = 0; i < parsed.medium; i++) arr.push({ severity: 'medium' });
  for (let i = 0; i < parsed.low; i++) arr.push({ severity: 'low' });
  return arr;
}

function buildMergeReportSheet(workbook) {
  const compSheet = workbook.getWorksheet('检出组件信息');
  if (!compSheet) return;

  const headerRow = detectHeaderRow(compSheet);
  const colMap = {};
  const headerCells = compSheet.getRow(headerRow);
  headerCells.eachCell(function (cell, colNumber) {
    const val = String(cell.value || '').trim();
    colMap[val] = colNumber;
  });

  const nameCol = colMap['组件名称'] || 4;
  const versionCol = colMap['版本号'] || 5;
  const vulnCountCol = colMap['漏洞数'] || 12;

  const appName = readAppNameFromTemplate(workbook) || '合并应用';

  const allData = [];
  for (let r = headerRow + 1; r <= compSheet.rowCount; r++) {
    const row = compSheet.getRow(r);
    const name = String(row.getCell(nameCol).value || '').trim();
    const version = String(row.getCell(versionCol).value || '').trim();
    if (!name || !version) continue;

    const vulnCountText = String(row.getCell(vulnCountCol).value || '').trim();

    const parsedVuln = parseVulnCount(vulnCountText);
    const vulns = buildVulnsArray(parsedVuln);

    const eolStatus = '待研判';

    allData.push({
      fullName: name,
      version: version,
      vulns: vulns,
      eolStatus: eolStatus,
    });
  }

  const existing = workbook.getWorksheet('应用安全检测报告');
  if (existing) workbook.removeWorksheet(existing.id);

  const navSheet = workbook.getWorksheet('导航');
  const navIdx = navSheet ? workbook.worksheets.indexOf(navSheet) : 0;

  const sheet = workbook.addWorksheet('应用安全检测报告');
  const wsIdx = workbook.worksheets.indexOf(sheet);
  workbook.worksheets.splice(wsIdx, 1);
  workbook.worksheets.splice(navIdx, 0, sheet);

  buildReportSheet(sheet, appName, allData);

  return allData.map(function (item) {
    return { name: item.fullName, version: item.version };
  });
}

export async function exportToExcel(reportData, batchResults, filePath, templatePath) {
  if (templatePath) {
    await exportFromTemplate(reportData, batchResults, filePath, templatePath);
  } else {
    await exportFromScratch(reportData, batchResults, filePath);
  }
}

function getEOLStatusText(data) {
  if (data.eolSource === 'internal') return '维护中';
  if (data.eolData && Array.isArray(data.eolData) && data.eolData.length > 0) {
    var eolVal = data.eolData[0].eol;
    if (eolVal === false) return '维护中';
    if (eolVal === true || eolVal === 'true') return '已EOL';
    if (typeof eolVal === 'string' && eolVal.trim() !== '') {
      var d = new Date(eolVal);
      if (!isNaN(d.getTime()) && d < new Date()) return '已EOL';
    }
    return '维护中';
  }
  if (data.aiEol) return data.aiEol.eolStatus || '已EOL';
  return '已EOL';
}

function countComponents(compSheet) {
  if (!compSheet) return 0;
  const headerRow = detectHeaderRow(compSheet);
  const colMap = {};
  const headerCells = compSheet.getRow(headerRow);
  headerCells.eachCell(function (cell, colNumber) {
    const val = String(cell.value || '').trim();
    colMap[val] = colNumber;
  });
  const nameCol = colMap['组件名称'] || 4;
  const versionCol = colMap['版本号'] || 5;
  let count = 0;
  for (let r = headerRow + 1; r <= compSheet.rowCount; r++) {
    const row = compSheet.getRow(r);
    const name = String(row.getCell(nameCol).value || '').trim();
    const version = String(row.getCell(versionCol).value || '').trim();
    if (name && version) count++;
  }
  return count;
}

function countDetailRows(sheet) {
  const headerRow = detectHeaderRow(sheet);
  if (!headerRow) return 0;
  const colMap = {};
  const headerCells = sheet.getRow(headerRow);
  headerCells.eachCell(function (cell, colNumber) {
    const val = String(cell.value || '').trim();
    colMap[val] = colNumber;
  });
  const nameCol = colMap['组件名称'] || 2;
  let count = 0;
  for (let r = headerRow + 1; r <= sheet.rowCount; r++) {
    const name = String(sheet.getRow(r).getCell(nameCol).value || '').trim();
    if (name) count++;
  }
  return count;
}

function readAppNameFromTemplate(workbook) {
  const sheet = workbook.getWorksheet('应用信息');
  if (!sheet || sheet.rowCount < 2) return '';
  const cell = sheet.getRow(2).getCell(1);
  return String(cell.value || '').trim();
}

async function exportFromTemplate(reportData, batchResults, filePath, templatePath) {
  const bytes = await readFileBytes(templatePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);

  const eolMap = {};
  const eolList = [];
  if (batchResults && batchResults.length > 0) {
    for (let i = 0; i < batchResults.length; i++) {
      const r = batchResults[i];
      const fullName = r.fullName || r.name || '';
      const name = r.name || '';
      const version = r.version || '';
      const status = r.eolStatus || '维护中';
      eolList.push({ fullName: fullName, name: name, version: version, status: status });
      eolMap[fullName + '|' + version] = status;
      if (name && name !== fullName) eolMap[name + '|' + version] = status;
      var colonIdx = fullName.indexOf(':');
      if (colonIdx !== -1) {
        var shortName = fullName.substring(colonIdx + 1).trim();
        if (shortName && shortName !== fullName) eolMap[shortName + '|' + version] = status;
      }
    }
  } else {
    const key = (reportData.rawName || reportData.name) + '|' + reportData.version;
    eolMap[key] = getEOLStatusText(reportData);
    eolList.push({ fullName: reportData.rawName || reportData.name, name: reportData.name, version: reportData.version, status: getEOLStatusText(reportData) });
  }

  const navSheet = workbook.getWorksheet('导航');
  const navIdx = navSheet ? workbook.worksheets.indexOf(navSheet) : 0;

  const existing = workbook.getWorksheet('应用安全检测报告');
  const compSheet = workbook.getWorksheet('检出组件信息');
  const compCount = countComponents(compSheet);
  const reportRows = existing ? countDetailRows(existing) : 0;

  if (existing && reportRows === compCount && compCount > 0) {
    updateEOLColumn(existing, eolMap, eolList, workbook);
    var curIdx = workbook.worksheets.indexOf(existing);
    if (curIdx !== navIdx - 1) {
      workbook.worksheets.splice(curIdx, 1);
      var targetIdx = navIdx > curIdx ? navIdx - 1 : navIdx;
      workbook.worksheets.splice(targetIdx, 0, existing);
    }
  } else {
    if (existing) workbook.removeWorksheet(existing.id);
    if (compSheet) {
      const newSheet = workbook.addWorksheet('应用安全检测报告');
      var wsIdx = workbook.worksheets.indexOf(newSheet);
      workbook.worksheets.splice(wsIdx, 1);
      var insertIdx = navSheet ? workbook.worksheets.indexOf(navSheet) : 0;
      workbook.worksheets.splice(insertIdx, 0, newSheet);
      buildReportSheetFromTemplate(newSheet, workbook, eolMap, eolList);
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const uint8 = new Uint8Array(buffer);
  await writeToFile(filePath, uint8);
}

function updateEOLColumn(sheet, eolMap, eolList, workbook) {
  const compSheet = workbook.getWorksheet('检出组件信息');
  const compVulnList = [];
  if (compSheet) {
    const compHeaderRow = detectHeaderRow(compSheet);
    const compColMap = {};
    const compHeaderCells = compSheet.getRow(compHeaderRow);
    compHeaderCells.eachCell(function (cell, colNumber) {
      const val = String(cell.value || '').trim();
      compColMap[val] = colNumber;
    });
    const compNameCol = compColMap['组件名称'] || 4;
    const compVersionCol = compColMap['版本号'] || 5;
    const vulnCountCol = compColMap['漏洞数'] || 12;

    for (let r = compHeaderRow + 1; r <= compSheet.rowCount; r++) {
      const row = compSheet.getRow(r);
      const name = String(row.getCell(compNameCol).value || '').trim();
      const version = String(row.getCell(compVersionCol).value || '').trim();
      if (!name || !version) continue;
      const vulnText = String(row.getCell(vulnCountCol).value || '').trim();
      const parsed = parseVulnCount(vulnText);
      compVulnList.push({ name: name, version: version, parsed: parsed });
    }
  }

  const headerRow = detectHeaderRow(sheet);
  if (!headerRow) return;

  const colMap = {};
  const headerCells = sheet.getRow(headerRow);
  headerCells.eachCell(function (cell, colNumber) {
    const val = String(cell.value || '').trim();
    colMap[val] = colNumber;
  });

  const nameCol = colMap['组件名称'] || 2;
  const versionCol = colMap['组件版本'] || 3;
  const eolCol = colMap['是否EOS'] || 4;
  const highVulnCol = colMap['是否存在中高危漏洞'] || 5;
  const anyVulnCol = colMap['是否存在漏洞'] || 6;
  const passCol = colMap['是否通过'] || 7;

  let hasEOL = false, hasHighVuln = false, hasAnyVuln = false, allPass = true;
  let allPending = true;
  let detailCount = 0;

  const seqCol = colMap['序号'] || 1;

  let compIdx = 0;
  for (let r = headerRow + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const name = String(row.getCell(nameCol).value || '').trim();
    if (!name) continue;

    const seqVal = row.getCell(seqCol).value;
    if (seqVal !== null && seqVal !== undefined && seqVal !== '') {
      detailCount++;
    }

    const version = String(row.getCell(versionCol).value || '').trim();
    const key = name + '|' + version;
    let eolStatus = eolMap[key];
    if (!eolStatus) {
      var colonIdx = name.indexOf(':');
      if (colonIdx !== -1) {
        var shortName = name.substring(colonIdx + 1).trim();
        eolStatus = eolMap[shortName + '|' + version];
      }
    }
    if (!eolStatus && compIdx < eolList.length) {
      eolStatus = eolList[compIdx].status;
    }
    if (eolStatus) {
      row.getCell(eolCol).value = eolStatus;
      if (eolStatus === '已EOL' || eolStatus === '即将EOL') hasEOL = true;
      if (eolStatus !== '待研判' && eolStatus !== '未知') allPending = false;
    }

    let vuln = null;
    for (let i = 0; i < compVulnList.length; i++) {
      if (compVulnList[i].name === name && compVulnList[i].version === version) {
        vuln = compVulnList[i].parsed;
        break;
      }
    }
    if (!vuln && compIdx < compVulnList.length) {
      vuln = compVulnList[compIdx].parsed;
    }
    compIdx++;

    if (vuln) {
      const hasHigh = vuln.critical > 0 || vuln.high > 0 || vuln.medium > 0;
      row.getCell(highVulnCol).value = hasHigh ? '是' : '否';
      row.getCell(anyVulnCol).value = vuln.total > 0 ? '是' : '否';
      row.getCell(passCol).value = hasHigh ? '否' : '是';
      if (hasHigh) { hasHighVuln = true; allPass = false; }
      if (vuln.total > 0) hasAnyVuln = true;
    }

    row.commit();
  }

  if (detailCount === 0) {
    for (let r = headerRow + 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const name = String(row.getCell(nameCol).value || '').trim();
      if (name) detailCount++;
    }
  }

  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const label = String(row.getCell(1).value || '').trim();
    if (label === '应用名称') continue;
    if (label === '组件数量') {
      row.getCell(2).value = detailCount;
      row.commit();
    } else if (label === '是否存在EOS组件') {
      const hasAny = Object.values(eolMap).some(function (s) { return s === '已EOL' || s === '即将EOL'; });
      const hasPending = Object.values(eolMap).some(function (s) { return s === '待研判' || s === '未知'; });
      row.getCell(4).value = hasPending ? '待研判' : (hasAny ? '是' : '否');
      row.commit();
    } else if (label === '是否存在中高危漏洞') {
      row.getCell(4).value = hasHighVuln ? '是' : '否';
      row.commit();
    } else if (label === '是否存在漏洞') {
      row.getCell(4).value = hasAnyVuln ? '是' : '否';
      row.commit();
    } else if (label === '是否通过') {
      row.getCell(4).value = allPending ? '待研判' : (allPass ? '是' : '否');
      row.commit();
    }
  }
}

function buildReportSheetFromTemplate(sheet, workbook, eolMap, eolList) {
  sheet.columns = REPORT_COL_WIDTHS.map(function (w) { return { width: w }; });

  const appName = readAppNameFromTemplate(workbook) || '未知应用';
  const compSheet = workbook.getWorksheet('检出组件信息');
  const components = [];
  if (compSheet) {
    const headerRow = detectHeaderRow(compSheet);
    const colMap = {};
    const headerCells = compSheet.getRow(headerRow);
    headerCells.eachCell(function (cell, colNumber) {
      const val = String(cell.value || '').trim();
      colMap[val] = colNumber;
    });
    const nameCol = colMap['组件名称'] || 4;
    const versionCol = colMap['版本号'] || 5;
    const vulnCountCol = colMap['漏洞数'] || 12;

    for (let r = headerRow + 1; r <= compSheet.rowCount; r++) {
      const row = compSheet.getRow(r);
      const name = String(row.getCell(nameCol).value || '').trim();
      const version = String(row.getCell(versionCol).value || '').trim();
      if (!name || !version) continue;
      const vulnText = String(row.getCell(vulnCountCol).value || '').trim();
      const parsed = parseVulnCount(vulnText);
      const key = name + '|' + version;
      let eolStatus = eolMap[key];
      if (!eolStatus) {
        var colonIdx = name.indexOf(':');
        if (colonIdx !== -1) {
          var shortName = name.substring(colonIdx + 1).trim();
          eolStatus = eolMap[shortName + '|' + version];
        }
      }
      if (!eolStatus && components.length < eolList.length) {
        eolStatus = eolList[components.length].status;
      }
      components.push({
        name: name, version: version,
        eolStatus: eolStatus || '已EOL',
        vulnCount: parsed.total,
        hasHigh: parsed.critical > 0 || parsed.high > 0 || parsed.medium > 0,
      });
    }
  }

  let hasEOL = false, hasHighVuln = false, hasAnyVuln = false, allPass = true;
  let allPending = true;
  for (let i = 0; i < components.length; i++) {
    const c = components[i];
    if (c.eolStatus === '已EOL' || c.eolStatus === '即将EOL') hasEOL = true;
    if (c.eolStatus !== '待研判' && c.eolStatus !== '未知') allPending = false;
    if (c.hasHigh) { hasHighVuln = true; allPass = false; }
    if (c.vulnCount > 0) hasAnyVuln = true;
  }

  sheet.mergeCells('A1:H1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = '软件成分分析情况总结表';
  titleCell.font = REPORT_FONT;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const summaryHeaders = ['应用名称', '组件数量', '系统类型', '是否存在EOS组件', '是否存在中高危漏洞', '是否存在漏洞', '是否通过', '备注'];
  const headerRow2 = sheet.getRow(2);
  for (let i = 0; i < summaryHeaders.length; i++) {
    const cell = headerRow2.getCell(i + 1);
    cell.value = summaryHeaders[i];
    cell.font = REPORT_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = REPORT_BORDER_NO_TOP;
  }
  headerRow2.commit();

  const eolSummary = allPending ? '待研判' : (hasEOL ? '是' : '否');
  const summaryData = [appName, components.length, '内网系统', eolSummary, hasHighVuln ? '是' : '否', hasAnyVuln ? '是' : '否', allPass ? '是' : '否', ''];
  const dataRow3 = sheet.getRow(3);
  for (let i = 0; i < summaryData.length; i++) {
    const cell = dataRow3.getCell(i + 1);
    cell.value = summaryData[i];
    cell.font = REPORT_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: i === 0 };
    cell.border = REPORT_BORDER;
  }
  dataRow3.commit();

  sheet.mergeCells('A5:H5');
  const detailTitleCell = sheet.getCell('A5');
  detailTitleCell.value = '软件成分分析情况明细表';
  detailTitleCell.font = REPORT_FONT;
  detailTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const detailHeaders = ['序号', '组件名称', '组件版本', '是否EOS', '是否存在中高危漏洞', '是否存在漏洞', '是否通过', '备注'];
  const headerRow6 = sheet.getRow(6);
  for (let i = 0; i < detailHeaders.length; i++) {
    const cell = headerRow6.getCell(i + 1);
    cell.value = detailHeaders[i];
    cell.font = REPORT_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = REPORT_BORDER_NO_TOP;
  }
  headerRow6.commit();

  for (let idx = 0; idx < components.length; idx++) {
    const c = components[idx];
    const row = sheet.getRow(7 + idx);

    row.getCell(1).value = idx + 1;
    row.getCell(1).font = REPORT_FONT;
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(1).border = REPORT_BORDER;

    row.getCell(2).value = c.name;
    row.getCell(2).font = REPORT_FONT;
    row.getCell(2).alignment = { vertical: 'middle' };

    row.getCell(3).value = c.version;
    row.getCell(3).font = REPORT_FONT;
    row.getCell(3).alignment = { vertical: 'middle' };

    row.getCell(4).value = c.eolStatus;
    row.getCell(4).font = REPORT_FONT;
    row.getCell(4).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    row.getCell(5).value = c.hasHigh ? '是' : '否';
    row.getCell(5).font = REPORT_FONT;
    row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(5).border = REPORT_BORDER;

    row.getCell(6).value = c.vulnCount > 0 ? '是' : '否';
    row.getCell(6).font = REPORT_FONT;
    row.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(6).border = REPORT_BORDER;

    row.getCell(7).value = c.hasHigh ? '否' : '是';
    row.getCell(7).font = REPORT_FONT;
    row.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(7).border = REPORT_BORDER;

    row.getCell(8).value = '';
    row.getCell(8).font = REPORT_FONT;

    row.commit();
  }
}

function buildReportSheet(sheet, appName, allData) {
  sheet.columns = REPORT_COL_WIDTHS.map(function (w) { return { width: w }; });

  let hasEOL = false, hasHighVuln = false, hasAnyVuln = false, allPass = true;
  let allPending = true;
  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const d = item.reportData || item;
    const vulns = d.vulns || item.vulns || [];
    const crit = vulns.filter(function (v) { return v.severity === 'critical'; }).length;
    const high = vulns.filter(function (v) { return v.severity === 'high'; }).length;
    const med = vulns.filter(function (v) { return v.severity === 'medium'; }).length;
    const eolStatus = item.eolStatus || getEOLStatusText(d);
    if (eolStatus === '已EOL' || eolStatus === '即将EOL') hasEOL = true;
    if (eolStatus !== '待研判') allPending = false;
    if (crit > 0 || high > 0 || med > 0) { hasHighVuln = true; allPass = false; }
    if (vulns.length > 0) hasAnyVuln = true;
  }

  sheet.mergeCells('A1:H1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = '软件成分分析情况总结表';
  titleCell.font = REPORT_FONT;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.border = { left: REPORT_BORDER.left };

  const summaryHeaders = ['应用名称', '组件数量', '系统类型', '是否存在EOS组件', '是否存在中高危漏洞', '是否存在漏洞', '是否通过', '备注'];
  const headerRow2 = sheet.getRow(2);
  for (let i = 0; i < summaryHeaders.length; i++) {
    const cell = headerRow2.getCell(i + 1);
    cell.value = summaryHeaders[i];
    cell.font = REPORT_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = REPORT_BORDER_NO_TOP;
  }
  headerRow2.commit();

  const eolSummary = allPending ? '待研判' : (hasEOL ? '是' : '否');
  const summaryData = [appName, allData.length, '互联网系统', eolSummary, hasHighVuln ? '是' : '否', hasAnyVuln ? '是' : '否', allPass ? '是' : '否', ''];
  const dataRow3 = sheet.getRow(3);
  for (let i = 0; i < summaryData.length; i++) {
    const cell = dataRow3.getCell(i + 1);
    cell.value = summaryData[i];
    cell.font = REPORT_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: i === 0 };
    cell.border = REPORT_BORDER;
  }
  dataRow3.commit();

  sheet.mergeCells('A5:H5');
  const detailTitleCell = sheet.getCell('A5');
  detailTitleCell.value = '软件成分分析情况明细表';
  detailTitleCell.font = REPORT_FONT;
  detailTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  detailTitleCell.border = { left: REPORT_BORDER.left };

  const detailHeaders = ['序号', '组件名称', '组件版本', '是否EOS', '是否存在中高危漏洞', '是否存在漏洞', '是否通过', '备注'];
  const headerRow6 = sheet.getRow(6);
  for (let i = 0; i < detailHeaders.length; i++) {
    const cell = headerRow6.getCell(i + 1);
    cell.value = detailHeaders[i];
    cell.font = REPORT_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = REPORT_BORDER_NO_TOP;
  }
  headerRow6.commit();

  for (let idx = 0; idx < allData.length; idx++) {
    const item = allData[idx];
    const d = item.reportData || item;
    const vulns = d.vulns || item.vulns || [];
    const crit = vulns.filter(function (v) { return v.severity === 'critical'; }).length;
    const high = vulns.filter(function (v) { return v.severity === 'high'; }).length;
    const med = vulns.filter(function (v) { return v.severity === 'medium'; }).length;
    const eolStatus = item.eolStatus || getEOLStatusText(d);
    const hasHigh = crit > 0 || high > 0 || med > 0;

    const row = sheet.getRow(7 + idx);

    const cellA = row.getCell(1);
    cellA.value = idx + 1;
    cellA.font = REPORT_FONT;
    cellA.alignment = { horizontal: 'center', vertical: 'middle' };
    cellA.border = REPORT_BORDER;

    const cellB = row.getCell(2);
    cellB.value = item.fullName || d.name || item.name;
    cellB.font = REPORT_FONT;
    cellB.alignment = { vertical: 'middle' };

    const cellC = row.getCell(3);
    cellC.value = item.version || d.version;
    cellC.font = REPORT_FONT;
    cellC.alignment = { vertical: 'middle' };

    const cellD = row.getCell(4);
    cellD.value = eolStatus;
    cellD.font = REPORT_FONT;
    cellD.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    const cellE = row.getCell(5);
    cellE.value = hasHigh ? '是' : '否';
    cellE.font = REPORT_FONT;
    cellE.alignment = { horizontal: 'center', vertical: 'middle' };
    cellE.border = REPORT_BORDER;

    const cellF = row.getCell(6);
    cellF.value = vulns.length > 0 ? '是' : '否';
    cellF.font = REPORT_FONT;
    cellF.alignment = { horizontal: 'center', vertical: 'middle' };
    cellF.border = REPORT_BORDER;

    const cellG = row.getCell(7);
    cellG.value = hasHigh ? '否' : '是';
    cellG.font = REPORT_FONT;
    cellG.alignment = { horizontal: 'center', vertical: 'middle' };
    cellG.border = REPORT_BORDER;

    const cellH = row.getCell(8);
    cellH.value = '';
    cellH.font = REPORT_FONT;
    cellH.alignment = { vertical: 'middle' };

    row.commit();
  }
}

async function exportFromScratch(reportData, batchResults, filePath) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = '组件安全研判';
  workbook.created = new Date();

  const hasBatch = batchResults && batchResults.length > 0;
  const allData = hasBatch ? batchResults : [{
    name: reportData.name, version: reportData.version, vendor: reportData.vendor,
    fullName: reportData.rawName || reportData.name, vulns: reportData.vulns || [],
    eolStatus: getEOLStatusText(reportData), eolData: reportData.eolData, aiEol: reportData.aiEol,
    eolSource: reportData.eolSource, reportData: reportData,
  }];

  buildSummarySheet(workbook, allData);
  buildVulnDetailSheet(workbook, allData);
  buildEOLDetailSheet(workbook, allData);

  const buffer = await workbook.xlsx.writeBuffer();
  const uint8 = new Uint8Array(buffer);
  await writeToFile(filePath, uint8);
}

function buildSummarySheet(workbook, allData) {
  const sheet = workbook.addWorksheet('概览', {
    views: [{ state: 'frozen', ySplit: 1 }],
    columns: [
      { width: 36 }, { width: 12 }, { width: 10 }, { width: 8 },
      { width: 8 }, { width: 8 }, { width: 8 }, { width: 12 },
      { width: 20 }, { width: 50 },
    ],
  });

  const headers = ['组件名称', '版本', '漏洞总数', '严重', '高危', '中危', '低危', 'EOL状态', '推荐安全版本', '判定依据'];
  const headerRow = sheet.addRow(headers);
  headerRow.height = 28;
  headerRow.eachCell(function (cell) { applyHeaderStyle(cell); });

  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const d = item.reportData || item;
    const vulns = d.vulns || item.vulns || [];
    const crit = vulns.filter(function (v) { return v.severity === 'critical'; }).length;
    const high = vulns.filter(function (v) { return v.severity === 'high'; }).length;
    const med = vulns.filter(function (v) { return v.severity === 'medium'; }).length;
    const low = vulns.filter(function (v) { return v.severity === 'low'; }).length;
    const eolStatus = item.eolStatus || getEOLStatusText(d);

    let safeVersion = '';
    let rationale = '';
    if (d.eolData && Array.isArray(d.eolData) && d.eolData.length > 0) {
      const active = d.eolData.find(function (v) { return v.eol === false; });
      if (active) safeVersion = active.latest || active.cycle;
      rationale = '数据来源：endoflife.date';
    } else if (d.aiEol) {
      safeVersion = d.aiEol.latestSafeVersion || '';
      rationale = d.aiEol.rationale || '';
    } else if (d.eolSource === 'internal') {
      safeVersion = d.version || '';
      rationale = '内部自研组件，由我方维护';
    }

    const row = sheet.addRow([
      item.fullName || d.name || item.name,
      item.version || d.version,
      vulns.length, crit, high, med, low,
      eolStatus, safeVersion, rationale,
    ]);

    applyDataStyle(row.getCell(1), { wrap: true });
    applyDataStyle(row.getCell(2));
    for (let c = 3; c <= 7; c++) applyDataStyle(row.getCell(c), { align: 'center' });
    applyEOLStyle(row.getCell(8), normalizeEOLStatus(eolStatus));
    applyDataStyle(row.getCell(9));
    applyDataStyle(row.getCell(10), { wrap: true });

    if (crit > 0 || high > 0 || med > 0) {
      row.getCell(4).font = Object.assign({}, row.getCell(4).font, { bold: true, color: { argb: 'FFE11D48' } });
      row.getCell(5).font = Object.assign({}, row.getCell(5).font, { bold: true, color: { argb: 'FFE11D48' } });
    }
  }
}

function buildVulnDetailSheet(workbook, allData) {
  const sheet = workbook.addWorksheet('漏洞详情', {
    views: [{ state: 'frozen', ySplit: 1 }],
    columns: [
      { width: 36 }, { width: 12 }, { width: 20 }, { width: 10 },
      { width: 10 }, { width: 60 },
    ],
  });

  const headers = ['组件名称', '版本', 'CVE ID', '严重等级', 'CVSS分数', '摘要'];
  const headerRow = sheet.addRow(headers);
  headerRow.height = 28;
  headerRow.eachCell(function (cell) { applyHeaderStyle(cell); });

  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const d = item.reportData || item;
    const vulns = d.vulns || item.vulns || [];
    const name = item.fullName || d.name || item.name;
    const version = item.version || d.version;

    if (vulns.length === 0) {
      const row = sheet.addRow([name, version, '无已知漏洞', '-', '-', '查询覆盖有限']);
      applyDataStyle(row.getCell(1), { wrap: true });
      for (let c = 2; c <= 6; c++) applyDataStyle(row.getCell(c), { align: c <= 4 ? 'center' : 'left' });
      continue;
    }

    for (let j = 0; j < vulns.length; j++) {
      const v = vulns[j];
      const row = sheet.addRow([name, version, v.id, v.severity.toUpperCase(), v.score, v.summary]);
      applyDataStyle(row.getCell(1), { wrap: true });
      applyDataStyle(row.getCell(2));
      applyDataStyle(row.getCell(3));
      applySeverityStyle(row.getCell(4), v.severity);
      applyDataStyle(row.getCell(5), { align: 'center' });
      applyDataStyle(row.getCell(6), { wrap: true });
    }
  }
}

function buildEOLDetailSheet(workbook, allData) {
  const sheet = workbook.addWorksheet('EOL详情', {
    views: [{ state: 'frozen', ySplit: 1 }],
    columns: [
      { width: 36 }, { width: 12 }, { width: 12 }, { width: 15 },
      { width: 40 }, { width: 20 }, { width: 50 },
    ],
  });

  const headers = ['组件名称', '版本', 'EOL状态', '数据来源', '当前版本状态', '推荐安全版本', '判定依据'];
  const headerRow = sheet.addRow(headers);
  headerRow.height = 28;
  headerRow.eachCell(function (cell) { applyHeaderStyle(cell); });

  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const d = item.reportData || item;
    const name = item.fullName || d.name || item.name;
    const version = item.version || d.version;
    const eolStatus = item.eolStatus || getEOLStatusText(d);
    const eolSource = d.eolSource || item.eolSource || '未知';

    let currentStatus = '';
    let safeVersion = '';
    let rationale = '';

    if (d.eolData && Array.isArray(d.eolData) && d.eolData.length > 0) {
      const active = d.eolData.find(function (v) { return v.eol === false; });
      if (active) {
        currentStatus = '版本 ' + active.cycle + ' 仍在维护';
        safeVersion = active.latest || active.cycle;
      }
      rationale = '数据来源：endoflife.date';
    } else if (d.aiEol) {
      currentStatus = d.aiEol.currentVersionStatus || '';
      safeVersion = d.aiEol.latestSafeVersion || '';
      rationale = d.aiEol.rationale || '';
    } else if (d.eolSource === 'internal') {
      currentStatus = '内部自研组件';
      safeVersion = version;
      rationale = '内部维护，不适用外部 EOL 判定';
    }

    const row = sheet.addRow([name, version, eolStatus, eolSource, currentStatus, safeVersion, rationale]);
    applyDataStyle(row.getCell(1), { wrap: true });
    applyDataStyle(row.getCell(2));
    applyEOLStyle(row.getCell(3), normalizeEOLStatus(eolStatus));
    applyDataStyle(row.getCell(4));
    applyDataStyle(row.getCell(5), { wrap: true });
    applyDataStyle(row.getCell(6));
    applyDataStyle(row.getCell(7), { wrap: true });
  }
}

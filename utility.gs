let lastValueCache = new Map();
let mergedRangeCache = new Map();
let mergedRangeCachePopulated = false;

function cachedMergedCellValues(sheet, row, col) {
  if (!mergedRangeCachePopulated) {
    const dataRange = sheet.getDataRange();
    const mergedRanges = dataRange.getMergedRanges();
    Logger.log(`Found ${mergedRanges.length} merged ranges.`);
    for (const range of mergedRanges) {
      const topLeftCell = range.getCell(1, 1);
      const value = topLeftCell.getValue();
      for (let r = range.getRow(); r <= range.getLastRow(); r++) {
        for (let c = range.getColumn(); c <= range.getLastColumn(); c++) {
          mergedRangeCache.set(`${r},${c}`, value);
        }
      }
    }
    mergedRangeCachePopulated = true;
  }
  return mergedRangeCache.get(`${row},${col}`);
}

// Function to get merged cell value with caching and fallback to last non-blank value
function getCellValue(sheet, row, col) {
  Logger.log(`${new Date().toISOString()} getting a value for ${row}, ${col}`);
  let value = cachedMergedCellValues(sheet, row, col);
  if (value === undefined) {
    value = sheet.getRange(row, col).getValue();
  }

  // Handle fallback to last non-blank value for specific columns
  if (!value && lastValueCache.has(col)) {
    value = lastValueCache.get(col);
  } else {
    // Update lastValueCache for current column if the value is not empty
    if (value && columnsFromLastAvalableValue.includes(col)) {
      lastValueCache.set(col, value);
    }
  }
  Logger.log(`${new Date().toISOString()} returning a value for ${row}, ${col}: ${value}`);
  return value;
}

function createProductDescription(description, productCare, material, sizeTable, madeIn) {
  // Replace placeholders in the HTML template
  const template = HtmlService.createHtmlOutputFromFile('ProductDescriptionTemplate')
    .getContent();

  // Escape special characters for HTML
  const escapeHtml = (text) => {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  };

  // Replace newline characters with <br>
  const replaceNewlines = (text) => {
    return text.replace(/\n/g, '<br>\n');
  };

  // Replace placeholders with actual content
  const populatedTemplate = template
    .replace(/\${DESCRIPTION}/g, replaceNewlines(escapeHtml(description)))
    .replace(/\${PRODUCTCARE}/g, replaceNewlines(escapeHtml(productCare)))
    .replace(/\${MATERIAL}/g, replaceNewlines(escapeHtml(material)))
    .replace(/\${SIZE_TABLE}/g, sizeTable) // sizeTable is already HTML
    .replace(/\${MADEIN}/g, replaceNewlines(escapeHtml(madeIn)));

  return populatedTemplate;
}

function createHtmlTableFromDynamicText(sizeData) {
  if (!sizeData || sizeData.length === 0) return '';

  const headersSet = new Set();
  const sizeMeasurementsMap = new Map();

  // Regex to match label and measurement pairs
  const regex = /([^\d\s]+)\s*(\d+(\.\d+)?)/g;

  // Process each entry in sizeData
  sizeData.forEach(({ sizeLabel, sizeMeasurement }) => {
    const measurements = {};
    let match;

    // Extract label and measurements from the sizeMeasurement text
    while ((match = regex.exec(sizeMeasurement)) !== null) {
      const label = match[1];
      const measurement = match[2];
      headersSet.add(label);            // Collect unique headers
      measurements[label] = measurement; // Map label to measurement
    }

    sizeMeasurementsMap.set(sizeLabel, measurements);
  });

  // Generate the table with headers from headersSet and data from sizeMeasurementsMap
  const headers = Array.from(headersSet);
  let tableHtml = '<table><thead><tr><th></th>';

  headers.forEach(header => {
    tableHtml += `<th>${header}</th>`;
  });
  tableHtml += '</tr></thead><tbody>';

  sizeMeasurementsMap.forEach((measurements, sizeLabel) => {
    tableHtml += `<tr><td>${sizeLabel}</td>`;
    headers.forEach(header => {
      tableHtml += `<td>${measurements[header] || ''}</td>`;
    });
    tableHtml += '</tr>';
  });

  tableHtml += '</tbody></table>';
  return tableHtml;
}

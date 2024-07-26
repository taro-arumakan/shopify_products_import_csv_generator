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
  Logger.log(`${new Date().toISOString()} returning a value for ${row}, ${col}`);
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
    .replace(/\${MATERIAL}/g, escapeHtml(material))
    .replace(/\${SIZE_TABLE}/g, sizeTable) // Assuming sizeTable is HTML and doesn't need escaping
    .replace(/\${MADEIN}/g, escapeHtml(madeIn));
  
  return populatedTemplate;
}

function createHtmlTableFromDynamicText(text) {
  const lines = text.trim().split('\n');
  const headings = ['サイズ'];
  const rows = [];

  lines.forEach((line) => {
    // Check for the size and value format
    const parts = line.match(/\[(\w+)\]\s*(.*)/);
    if (!parts) {
      // Skip if parts are not found
      return;
    }

    const size = parts[1];
    const values = parts[2].match(/(\S+\s+\d+\.?\d*)/g);

    if (headings.length === 1 && values) {
      // Extract headings from values
      values.forEach(value => {
        const heading = value.split(/\s+/)[0];
        if (!headings.includes(heading)) {
          headings.push(heading);
        }
      });
    }

    const row = [size];
    headings.slice(1).forEach(heading => {
      const value = values.find(v => v.startsWith(heading));
      row.push(value ? value.split(/\s+/)[1] : '');
    });
    rows.push(row);
  });

  // Handle case where rows might be empty
  if (rows.length === 0) {
    return '<p>No size data available.</p>';
  }

  let table = '<table><thead><tr>';
  headings.forEach(heading => {
    table += `<th>${heading}</th>`;
  });
  table += '</tr></thead><tbody>';
  rows.forEach(row => {
    table += '<tr>';
    row.forEach(cell => {
      table += `<td>${cell}</td>`;
    });
    table += '</tr>';
  });
  table += '</tbody></table>';

  return table;
}


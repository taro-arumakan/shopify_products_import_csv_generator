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
    .replace(/\${MATERIAL}/g, replaceNewlines(escapeHtml(material)))
    .replace(/\${SIZE_TABLE}/g, sizeTable) // Assuming sizeTable is HTML and doesn't need escaping
    .replace(/\${MADEIN}/g, escapeHtml(madeIn));
  
  return populatedTemplate;
}

function createHtmlTableFromDynamicText(text) {
  const lines = text.trim().split('\n');
  const headings = [''];
  const rows = [];

  lines.forEach((line, index) => {
    const parts = line.match(/\[(\w+)\]\s*(.*)/); // Check for the size and value format
    if (!parts) {
      return; // Skip if parts are not found
    }

    const size = parts[1];
    const values = parts[2].match(/(\D+\s*\d+\.?\d*)/g);

    if (headings.length === 1 && values) {
      // Extract headings from values
      values.forEach(value => {
        const heading = value.match(/^\D+/)?.[0]?.replace('/', '').trim();
        if (heading && !headings.includes(heading)) {
          headings.push(heading);
        }
      });
    }

    if (values) {
      const row = [size];
      values.forEach(value => {
        const valueParts = value.match(/(\D+)\s*(\d+\.?\d*)/);
        if (valueParts) {
          const numericValue = valueParts[2];
          row.push(numericValue);
        }
      });

      // Fill remaining columns if row length is less than headings
      while (row.length < headings.length) {
        row.push('');
      }
      rows.push(row);
    }
  });

  // Handle case where rows might be empty
  if (rows.length === 0) {
    return `<p>サイズ: ${text}</p>`;
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

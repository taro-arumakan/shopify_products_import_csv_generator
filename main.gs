// Column indexes, 0 base
const columnIndexes = {
  description: 15, // Column P
  productCare: 17, // Column R
  material: 18, // Column S
  sizeTable: 19, // Column T
  title: 2, // Column C
  option1Value: 3, // Column D
  option2Value: 4, // Column E
  variantSku: 5, // Column F
  variantInventoryQty: 11, // Column L
  variantPrice: 10, // Column K
  madeIn: 20, // Column U
  category: 4, // Column G
  collection: 5 // column H
};

 // Column indexes, 1 base
const columnsFromLastAvalableValue = [
  3,
  4
]

function createProductImportCsvSheet(sourceSheetName, headerRowsToSkip) {
  Logger.log(`${new Date(new Date().getTime()).toISOString()} starting to process ${sourceSheetName}`);
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('Source sheet not found.');
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const csvData = [];

  // CSV header
  const csvHeader = [
    'Handle', 'Title', 'Body (HTML)', 'Vendor', 'Tags', 'Published', 'Option1 Name',
    'Option1 Value', 'Option2 Name', 'Option2 Value', 'Variant SKU', 'Variant Inventory Tracker', 'Variant Inventory Qty',
    'Variant Inventory Policy', 'Variant Fulfillment Service', 'Variant Price',
    'Variant Requires Shipping', 'Variant Taxable'
  ];
  csvData.push(csvHeader);

  const descriptionCache = new Map();
  const sizeTableCache = new Map();
  const lastValueCache = new Map(); // Key: column index, Value: last non-empty value

  for (let i = headerRowsToSkip; i < data.length; i++) {
    const row = data[i];

    const handle = '=googletranslate(B' + (i + 2 - headerRowsToSkip) + ',"ja","en")';
    const title = getCellValue(sourceSheet, i + 1, columnIndexes.title + 1);
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title}`);

    // Get merged values
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title} - option1Value`);
    const option1Value = getCellValue(sourceSheet, i + 1, columnIndexes.option1Value + 1);
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title} - description`);
    const description = getCellValue(sourceSheet, i + 1, columnIndexes.description + 1);
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title} - productCare`);
    const productCare = getCellValue(sourceSheet, i + 1, columnIndexes.productCare + 1);
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title} - material`);
    const material = getCellValue(sourceSheet, i + 1, columnIndexes.material + 1);
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title} - size`);
    const sizeTableText = getCellValue(sourceSheet, i + 1, columnIndexes.sizeTable + 1);
    Logger.log(`${new Date(new Date().getTime()).toISOString()} processing ${title} - Made In`);
    const madeIn = getCellValue(sourceSheet, i + 1, columnIndexes.madeIn + 1);

    Logger.log(`${new Date(new Date().getTime()).toISOString()} caching description`);

    // Generate HTML for description and size table, avoid duplication
    const descriptionHtml = descriptionCache.has(description)
      ? descriptionCache.get(description)
      : createProductDescription(description, productCare, material, createHtmlTableFromDynamicText(sizeTableText), madeIn);
    descriptionCache.set(description, descriptionHtml);

    Logger.log(`${new Date(new Date().getTime()).toISOString()} caching size table`);

    const sizeTableHtml = sizeTableCache.has(sizeTableText)
      ? sizeTableCache.get(sizeTableText)
      : createHtmlTableFromDynamicText(sizeTableText);
    sizeTableCache.set(sizeTableText, sizeTableHtml);

    const bodyHtml = descriptionHtml;

    Logger.log(`${new Date(new Date().getTime()).toISOString()} done generating body html`);

    const tags = row[columnIndexes.category + 1] + ', ' + row[columnIndexes.collection + 1]; // Column F & G
    const variantSku = row[columnIndexes.variantSku];
    const variantInventoryQty = row[columnIndexes.variantInventoryQty];
    const variantPrice = row[columnIndexes.variantPrice];
    const option2Value = row[columnIndexes.option2Value];

    Logger.log(`${new Date(new Date().getTime()).toISOString()} adding a csv row`);
    const csvRow = [
      handle, title, bodyHtml, 'KUME', tags, 'True', 'カラー', option1Value, 'サイズ', option2Value, variantSku, 'shopify',
      variantInventoryQty, 'deny', 'manual', variantPrice, 'True', 'True'
    ];
    csvData.push(csvRow);

    Logger.log(`${new Date(new Date().getTime()).toISOString()} done adding a csv row`);
  }

  const newSheetName = 'Product Import CSV';
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
  if (newSheet) {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(newSheet);
  }
  newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newSheetName);
  newSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
}

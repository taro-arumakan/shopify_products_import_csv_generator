const activateProducts = true;
const vendor = 'GBH'

const gbhColumIndexes = {
  releaseDate: 1,           // Column B
  collection: 2,            // Column C
  category: 3,              // Column D
  category2: 4,             // Column E
  title: 5,                 // Column F
  option1Value: 6,          // Column G
  option2Value: 7,          // Column H
  variantSku: 8,            // Column I
  variantPrice: 11,         // Column L
  variantInventoryQty: 12,  // Column M
  description: 16,          // Column Q
  productCare: 18,          // Column S
  sizeTable: 20,            // Column U
  material: 21,             // Column V
  madeIn: 22,               // Column W
}

// Column indexes, 0 base
const kumeColumnIndexes = {
  releaseDate: 1,           // Column B
  title: 2,                 // Column C
  option1Value: 3,          // Column D
  option2Value: 4,          // Column E
  variantSku: 5,            // Column F
  category: 6,              // Column G
  collection: 7,            // Column H
  variantPrice: 10,         // Column K
  variantInventoryQty: 11,  // Column L
  description: 16,          // Column Q
  productCare: 18,          // Column S
  material: 19,             // Column T
  sizeTable: 20,            // Column U
  madeIn: 21,               // Column V
};

const columnIndexes = vendor == 'GBH' ? gbhColumIndexes : kumeColumnIndexes;

// Column indexes, 1 base
const columnsFromLastAvalableValue = [
  3,
  4
]

function populateProductDescription(sourceSheet, headerRowsToSkip) {
  productDescriptionMap = {};
  const data = sourceSheet.getDataRange().getValues();
  let sizeData = {};
  for (let i = headerRowsToSkip; i < data.length; i++) {
    const row = data[i];
    const title = getCellValue(sourceSheet, i + 1, columnIndexes.title + 1);
    const variantSize = row[columnIndexes.option2Value].trim();
    // Get merged values
    const option1Value = getCellValue(sourceSheet, i + 1, columnIndexes.option1Value + 1);
    const sizeTableText = getCellValue(sourceSheet, i + 1, columnIndexes.sizeTable + 1);
    const category = getCellValue(sourceSheet, i + 1, columnIndexes.category + 1);
    const collection = getCellValue(sourceSheet, i + 1, columnIndexes.collection + 1);

    // Use a "Set" approach to store unique size measurements per product
    if (!sizeData[title]) {
      sizeData[title] = [];
    }
    if (sizeTableText) {
      sizeData[title].push({ sizeLabel: variantSize, sizeMeasurement: sizeTableText });
    }

    // Process the product once all its rows are handled
    const nextRow = data[i + 1] || [];
    const nextTitle = nextRow[columnIndexes.title];
    const isLastVariant = !nextTitle || nextTitle !== title;
    let bodyHtml = '';
    if (isLastVariant) {
      // Combine unique sizes and measurements into a single text
      const sizeTableHtml = createHtmlTableFromDynamicText(sizeData[title]);
      console.log(sizeTableHtml);

      // Get product-level fields for the description
      const description = getCellValue(sourceSheet, i + 1, columnIndexes.description + 1);
      const productCare = getCellValue(sourceSheet, i + 1, columnIndexes.productCare + 1);
      const material = getCellValue(sourceSheet, i + 1, columnIndexes.material + 1);
      const madeIn = getCellValue(sourceSheet, i + 1, columnIndexes.madeIn + 1);

      // Create description HTML with the consolidated size table
      const descriptionHtml = createProductDescription(description, productCare, material, sizeTableHtml, madeIn);

      productDescriptionMap[title] = descriptionHtml;
    }
  }
  return productDescriptionMap;
}

function createProductImportCsvSheet(sourceSheetName, headerRowsToSkip) {
  Logger.log(`${new Date(new Date().getTime()).toISOString()} starting to process ${sourceSheetName}`);
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('Source sheet not found.');
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const csvData = [];

  const csvHeader = [
    'Handle', 'Title', 'Body (HTML)', 'Vendor', 'Tags', 'Published', 'Option1 Name',
    'Option1 Value', 'Option2 Name', 'Option2 Value', 'Variant SKU', 'Variant Inventory Tracker', 'Variant Inventory Qty',
    'Variant Inventory Policy', 'Variant Fulfillment Service', 'Variant Price',
    'Variant Requires Shipping', 'Variant Taxable', 'Status'
  ];
  csvData.push(csvHeader);

  const productDescriptionMap = populateProductDescription(sourceSheet, headerRowsToSkip);

  for (let i = headerRowsToSkip; i < data.length; i++) {
    const row = data[i];
    const handle = '=googletranslate(B' + (i + 2 - headerRowsToSkip) + ',"ja","en")';
    const title = getCellValue(sourceSheet, i + 1, columnIndexes.title + 1);

    if (!title) {
      throw new Error('no title');
    }

    const variantSize = row[columnIndexes.option2Value].trim();
    Logger.log(`${new Date(new Date().getTime()).toISOString()} --- processing ${title}`);

    // Get merged values
    const option1Value = getCellValue(sourceSheet, i + 1, columnIndexes.option1Value + 1);
    let category = getCellValue(sourceSheet, i + 1, columnIndexes.category + 1);
    if (columnIndexes.category2) {
        category = `${category}, ${getCellValue(sourceSheet, i + 1, columnIndexes.category2 + 1)}`;
    }
    const collection = getCellValue(sourceSheet, i + 1, columnIndexes.collection + 1);

    bodyHtml = productDescriptionMap[title];
    let tags;
    let status;
    if (activateProducts) {
      tags = `${category}, ${collection}`;
      status = 'active';
    } else {
      const releaseDate = getCellValue(sourceSheet, i + 1, columnIndexes.releaseDate + 1).replace('\n', '');
      tags = `${releaseDate}, ${category}, ${collection}`;
      status = 'draft';
    }
    tags = `${tags}, new`;

    const variantSku = row[columnIndexes.variantSku];
    const variantInventoryQty = row[columnIndexes.variantInventoryQty];
    const variantPrice = row[columnIndexes.variantPrice];

    const csvRow = [
      handle, title, bodyHtml, vendor, tags, 'True', 'カラー', option1Value.trim(), 'サイズ', variantSize.trim(),
      variantSku.trim(), 'shopify', variantInventoryQty, 'deny', 'manual', variantPrice, 'True', 'True', status
    ];
    csvData.push(csvRow);
  }

  const newSheetName = 'Product Import CSV';
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
  if (newSheet) {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(newSheet);
  }
  newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newSheetName);
  newSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
}

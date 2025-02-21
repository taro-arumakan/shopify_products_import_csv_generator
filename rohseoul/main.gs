const activateProducts = false;
const vendor = 'rohseoul'     // alvana, KUME, GBH, rohseoul


const columnIndexes = {     // Column indexes, 0 base
  releaseDate: 2,           // Column C
  title: 3,                 // Column D
  variantSku: 4,            // Column E
  collection: 5,            // Column F
  category: 6,              // Column G
  category2: 7,             // Column H
  option1Value: 8,          // Column I, Color
  // option2Value: 11,         // Column L, Size
  variantPrice: 11,         // Column L
  variantInventoryQty: 12,  // Column M
  description: 16,          // Column Q
  sizeTable: 19,            // Column T
  material: 20,             // Column U
  madeIn: 21,               // Column V
  // productCare: 5,           // Column F
};

function populateProductDescription(sourceSheet, headerRowsToSkip) {
  productDescriptionMap = {};
  const data = sourceSheet.getDataRange().getValues();
  let sizeData = {};
  for (let i = headerRowsToSkip; i < data.length; i++) {
    const row = data[i];
    const variantSku = row[columnIndexes.variantSku];
    if (!variantSku) {
      console.log(`populateProductDescription - breaking at row: ${i}`);
      break
    }
    const title = getCellValue(sourceSheet, i + 1, columnIndexes.title + 1);
    let variantSize = "";

    // Get merged values
    const sizeTableText = getCellValue(sourceSheet, i + 1, columnIndexes.sizeTable + 1);

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
    if (isLastVariant) {
      // Combine unique sizes and measurements into a single text
      const sizeTableHtml = createHtmlTableFromDynamicText(sizeData[title]);
      console.log(sizeTableHtml);

      // Get product-level fields for the description
      const description = getCellValue(sourceSheet, i + 1, columnIndexes.description + 1);
      let productCare = ''
      console.log('getting productCare');
      if (columnIndexes.productCare) {
        productCare = getCellValue(sourceSheet, i + 1, columnIndexes.productCare + 1);
      } else {
        productCare = `革表面に跡や汚れなどが残る場合がありますが、天然皮革の特徴である不良ではございませんのでご了承ください。また、時間経過により金属の装飾や革の色が変化する場合がございますが、製品の欠陥ではありません。あらかじめご了承ください。
1: 熱や直射日光に長時間さらされると革に変色が生じることがありますのでご注意ください。
2: 変形の恐れがありますので、無理のない内容量でご使用ください。
3: 水に弱い素材です。濡れた場合は柔らかい布で水気を除去した後、乾燥させてください。
4: 使用しないときはダストバッグに入れ、涼しく風通しのいい場所で保管してください。
5: アルコール、オイル、香水、化粧品などにより製品が損傷することがありますので、ご使用の際はご注意ください。`;
      }
      console.log(`productCare: ${productCare}`);
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
    'Handle', 'Title', 'Body (HTML)', 'Vendor', 'Tags', 'Published', 'Option1 Name', 'Option1 Value',
    'Variant SKU', 'Variant Inventory Tracker', 'Variant Inventory Qty',
    'Variant Inventory Policy', 'Variant Fulfillment Service', 'Variant Price',
    'Variant Requires Shipping', 'Variant Taxable', 'Status'
  ];
  csvData.push(csvHeader);

  console.log(`Populating product description map first`);
  const productDescriptionMap = populateProductDescription(sourceSheet, headerRowsToSkip);

  console.log(`Starting to process each product`);
  for (let i = headerRowsToSkip; i < data.length; i++) {
    const row = data[i];
    const handle = '=googletranslate(B' + (i + 2 - headerRowsToSkip) + ',"ja","en")';
    const title = getCellValue(sourceSheet, i + 1, columnIndexes.title + 1);

    if (!title) {
      throw new Error('no title');
    }

    Logger.log(`${new Date(new Date().getTime()).toISOString()} --- processing ${title}`);

    // Get merged values
    const option1Value = getCellValue(sourceSheet, i + 1, columnIndexes.option1Value + 1);
    let tags;     // category, category2, and collection
    tags = getCellValue(sourceSheet, i + 1, columnIndexes.category + 1);
    if (columnIndexes.category2) {
      tags = `${tags}, ${getCellValue(sourceSheet, i + 1, columnIndexes.category2 + 1)}`;
    }
    if (columnIndexes.collection) {
      tags = `${tags}, ${getCellValue(sourceSheet, i + 1, columnIndexes.collection + 1)}`;
    }

    bodyHtml = productDescriptionMap[title];
    let status;
    if (activateProducts) {
      status = 'active';
    } else {
      if (columnIndexes.releaseDate) {
        tags = `${tags}, ${getCellValue(sourceSheet, i + 1, columnIndexes.releaseDate + 1).replace('\n', '')}`;
      }
      status = 'draft';
    }
    tags = `${tags}, new`;

    const variantSku = row[columnIndexes.variantSku];
    if (!variantSku) {
      console.log(`breaking at row: ${i}`);
      break
    }
    const variantInventoryQty = row[columnIndexes.variantInventoryQty];
    const variantPrice = getCellValue(sourceSheet, i + 1, columnIndexes.variantPrice + 1);

    const csvRow = [
      handle, title, bodyHtml, vendor, tags, 'True', 'カラー', option1Value.trim(),
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

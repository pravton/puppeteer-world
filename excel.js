var Excel = require('exceljs');
const data = require('./covert');


function updateFile(data) {
  data
  var workbook = new Excel.Workbook();

  workbook.csv.readFile('./data.csv')
    .then(function () {
      console.log('Loading the worksheet...');
      var worksheet = worksheet = workbook.getWorksheet(1);
      workbook.removeWorksheet(worksheet.id);
      console.log('Deleted old content...');
      workbook.addWorksheet('My Sheet');
      var worksheet = worksheet = workbook.getWorksheet(1);
      const rowsL = worksheet.getColumn(1);
      const rowsCount = rowsL['_worksheet']['_rows'].length;
      worksheet.spliceRows(1, rowsCount);
      console.log('Row count before update:' + rowsCount);
      worksheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Title', key: 'title', width: 50 },
        { header: 'Variation', key: 'variationName', width: 50 },       
        { header: 'Price', key: 'price', width: 32 },
        { header: 'Stock Status', key: 'stockStatus', width: 10 },
        { header: 'Single Product', key: 'singleProduct', width: 10 },
        ];
        
        // Add a row by sparse Array (assign to columns A, E & I)

        const rows = [];
        data.forEach((product, idx) => {
          if(product.singleProduct) {
            const productData = {}; 

            productData.id = idx;
            productData.title = product.title;
            productData.variationName = 'Single Product';
            productData.price = product.price;
            productData.stockStatus = product.stock;
            productData.singleProduct = 'true';

            rows.push(productData);

          } else {
            for(let variation of product.variations) {
              const productData = {};
              productData.id = idx;
              productData.title = product.title;
              productData.variationName = variation.variation;
              productData.price = variation.price;
              productData.stockStatus = variation.stockStatus;
              productData.singleProduct = 'false';

              rows.push(productData);
            }
          }
          
        });

        // console.log(rows);
        worksheet.addRows(rows);
        
        //write in File
        var strFilename = "data.csv";
        workbook.csv.writeFile(strFilename)
        .then(function() {
        console.log("File Updated with new Data!!!");
        });
    })
}

updateFile(data);
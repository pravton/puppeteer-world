const puppeteer = require('puppeteer');
const axios = require('axios').default;
const parseString = require('xml2js').parseString;
var fs = require('fs')
const xml2js = require('xml2js');
var Excel = require('exceljs');

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
        console.log("File Updated with New Data!!!");
        });
    })
}


const findPageDetailsfunction = async () => {
  const finalResult = [];
  let linksList = [];
  async function getData() {
    try {
      const sitemap = await axios.get('https://laroccakitchen.com/sitemap.xml');
      parseString(sitemap.data, async function (err, result) {
        const productsLink = result.sitemapindex.sitemap[0].loc[0];
        const response = await axios.get(productsLink);
        // console.log(response.data);
        parseString(response.data, function (err, result) {
          const urlData = result.urlset.url;
          const urls = [];
          for (let url of urlData) {
            if(url.loc[0] !== 'https://laroccakitchen.com/') {
              // console.log(url.loc[0]);
              urls.push(url.loc[0]);
            }
          }
          console.log(`Loading ${urls.length} URLs for checking...`);
          linksList = urls;
        });
      });
    } catch (error) {
      console.error(error);
    }
  }

  getData();

  const browser = await puppeteer.launch({
    headless: true,
    args: [`--window-size=1920,1080`],
    defaultViewport: {
      width:1920,
      height:1080
    }
  });
    const page = await browser.newPage();
    await page.goto('https://laroccakitchen.com/', {
      waitUntil: 'networkidle0',
    });
    await page.click('a[href="https://laroccakitchen.com"]');
    await page.goto('https://laroccakitchen.com/collections/seasonal', {
      waitUntil: 'networkidle0',
    });

    console.log('Product checking has been started. Please wait until this window close itself!');
    
//   let linksList = ['https://laroccakitchen.com/products/lemon-meringue-tart',
//   'https://laroccakitchen.com/products/gift-card',
// 'https://laroccakitchen.com/collections/cakes/products/la-rocca-chocolate-fudge-cake']


    for (let link of linksList) {
      console.log(link);
      await page.goto(link, {
        waitUntil: 'networkidle0',
      });
      let resObj = await page.evaluate(() => {
        const event = new Event('change');
        let pageTitle = document.querySelector('h1.product-single__title').textContent.replace(/(\r\n|\n|\r)/gm, "").trim();
        const allVariations = document.querySelectorAll('.variant-input-wrap[data-index="option1"] select option');
        let variations = []; 
        let price;
        let inStock = true;
        let singleProduct = false;
        if(allVariations.length) {
          for (let i = 0; i < allVariations.length; i++) {
            const varObj = {
              variation: '',
              price: '',
              stockStatus: true,
            }
            document.querySelector('.variant-input-wrap select').selectedIndex = i;
            document.querySelector('.variant-input-wrap select').dispatchEvent(event);
            varObj.variation = document.querySelector('.variant-input-wrap[data-index="option1"] select').value;
            varObj.price = document.querySelector('.product__price').textContent.replace(/(\r\n|\n|\r)/gm, "").trim();
            if(document.querySelector('.add-to-cart[disabled]')) {
              varObj.stockStatus = false;
            }
            variations.push(varObj);
          }
        } else {
            singleProduct = true;
            price = document.querySelector('.product__price').textContent.replace(/(\r\n|\n|\r)/gm, "").trim();
            if(document.querySelector('.add-to-cart[disabled]')) {
              inStock = false;
            } 
        }

        let res = {};

        if(singleProduct) {
          res = {
            title: pageTitle,
            variations: variations,
            singleProduct: singleProduct,
            price: price,
            stock: inStock
          };
        } else {
          res = {
            title: pageTitle,
            variations: variations,
            singleProduct: singleProduct,
          };
        }
      
        return res;
      });
      finalResult.push(resObj);
    }
    // console.log(finalResult);
    browser.close();
    return finalResult;
}

const renderResults = async function() {
  const result = await findPageDetailsfunction();
  updateFile(result);
}

renderResults();

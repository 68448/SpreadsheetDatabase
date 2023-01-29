const Excel = require('exceljs');
const fs = require('fs');
var convert = require('xml-js');
const axios = require('axios');
const { JSONFromTable } = require('jsonfromtable')




//Pioner login
let pronet_login = 'Sales@4k-monitor.ru'
let pronet_password = '4k-monitor.ru'
//Подключение к бд
const import_database = require("./database.js");
var database = import_database.database()




//Получение html страницы и парсинг ее
async function html_parse(){

    const response = await axios.get('https://f5store.ru/opt/dx/'); 
    const obj = JSONFromTable.fromString(response.data)
    
    var results = []
    for (let i = 0; i < obj.length; i++){
        var product = []

        var pn = obj[i]['Модель'].split(' ')

        product.pn = String(pn.slice(-1))
        product.price = obj[i]['Скидка 25%']
        product.rrcprice = obj[i]['Цена сайта']

        results.push(product)
    }
    return results
      
}

// Получение информации о продукте OSC
async function osc_get_product_information(products, products_pn){

    const response = await axios.post(
        'https://connector.b2b.ocs.ru/api/v2/catalog/products/batch',
        products,
        {
            params: {
                'shipmentcity': '\u041C\u043E\u0441\u043A\u0432\u0430'
            },
            headers: {
                'accept': 'application/json',
                'X-API-Key': 'aPo8FHEFJlF@I2E7B_sebdTcyarp?t',
                'Content-Type': 'application/json'
            }
        }
    );    
    
    // Обработка данных из OCS
    var products = []  
    const res = response     

    if (res.data.result != null){     
        for(let i = 0; i < products_pn.length; i++){
            var product = {}
            for(let j = 0; j < res.data.result.length; j++){
                if( products_pn[i] == res.data.result[j].product.partNumber ){                        
                    product.name = res.data.result[j].product.itemNameRus + res.data.result[j].product.productName
                    product.price_rrc = ' '
                    product.price_ocs = String(res.data.result[j].price.order.value)
                    product.pn = res.data.result[j].product.productKey
                    product.sort_filed = res.data.result[j].product.partNumber
    
                    let location = res.data.result[j].locations[0].location
                    let quantity = res.data.result[j].locations[0].quantity.value
                    let quantity_stock = res.data.result[j].locations[0].quantity.isGreatThan
    
                    if ( quantity == 1 && quantity_stock == true ) product.quantity = 'Да'
                    if ( quantity == 0 && quantity_stock == true ) product.quantity = 'По запросу'
                    if ( quantity > 1 && quantity_stock == true ) product.quantity = 'Да'
                    if ( quantity > 1 && quantity_stock == false ) product.quantity = 'Да'

                }
            }
            products.push(product)
        }


    }else{
        products.push('Товаров в OCS не найдено')
    }

    return products
}

async function create_OCS_Pronet_object(rows){

    const resultObj = rows.map(x => (
        {
            name: x['Модель игрового кресла(name)'],
            pn: x['P/N'],
            ocs: (x['P/N OCS'] != '') ? x['P/N OCS'] : '',
            pronet: (x['P/N Pronet'] != '') ? x['P/N Pronet'] : ''
        }
    ))

    return resultObj;
}

function mapOrder(a, order, key) {
    const map = order.reduce((r, v, i) => ((r[v] = i), r), {})
    return a.sort((a, b) => map[a[key]] - map[b[key]])
}

// Получение токена Pronet
async function pronet_get_token(login, pass){

    const response = await axios.get(`http://png.pronetgroup.ru:116/api/Login/GetLogon?login=${ login }&password=${ pass }`); 
  
    return response

}

//Получение информации о продукте Pronet
async function pronet_get_product_information(token, productsList, ocsProducts){

    const response = await axios.get(`http://png.pronetgroup.ru:116/api/Item/GetItems?token=${ token }&listId=${ productsList }`); 

    //Обработка данных из Pronet
    var pronet_products = []
    const pronet_res = response
   
    if(pronet_res.data != null){
        for(let i = 0; i < pronet_res.data.length; i++){
            var product = {}
            if(pronet_res.data[i] != null){
                if(ocsProducts[i] == null){
                    product.name = pronet_res.data[i].NameRus
                }
                product.price = pronet_res.data[i].PriceRUR
                product.pn = pronet_res.data[i].Id
                if (pronet_res.data[i].Available == 0){
                    product.quantity = 'Нет'
                }else{
                    product.quantity = 'Да'
                }
                product.sort_filed = ''
                if( pronet_res.data[i].PartNumber != null ){
                    product.sort_filed = pronet_res.data[i].PartNumber
                }else{
                    product.sort_filed = pronet_res.data[i].Model
                }
                
            }else{
                product.price = ' '
                product.pn = ' '
                product.quantity = ' '
                product.sort_filed = ' '
            }

            pronet_products.push(product)
        }
    }


    return pronet_products
}

//Получение данных из xlsx файла
function get_xlsx_data(filename){

    var workbook = new Excel.Workbook(); 
  
    return workbook.xlsx.readFile(filename)
  }
//Получение цен из xlsx  
function get_prices(list, headerRowIDX, data, startIDX, articleName, priceName){

var worksheet = data.getWorksheet(list);
worksheet.getRow(headerRowIDX).eachCell((cell, colNumber) => {
    worksheet.getColumn(colNumber).key = cell.text;
});
const rows = worksheet.actualRowCount

var results = []
for (let i = startIDX; i <= rows; i++){
    let result = {}
    result.pn = worksheet.getRow(i).getCell(articleName).value
    result.price = worksheet.getRow(i).getCell(priceName).value
    results.push(result)
}

return results
}
//Получения наличия из xlsx
function get_stock(list, headerRowIDX , data, startIDX, articleName, stockName){
var worksheet = data.getWorksheet(list);
worksheet.getRow(headerRowIDX).eachCell((cell, colNumber) => {
    worksheet.getColumn(colNumber).key = cell.text;
});
const rows = worksheet.actualRowCount

var results = []
for (let i = startIDX; i <= rows; i++){
    let result = {}
    result.pn = worksheet.getRow(i).getCell(articleName).value
    if(
        String(worksheet.getRow(i).getCell(stockName).value).toUpperCase  == 'НЕТ'
        || String(worksheet.getRow(i).getCell(stockName).value).toUpperCase  == 'НЕТ В НАЛИЧИИ'
        || String(worksheet.getRow(i).getCell(stockName).value) == 0
        ){
        result.stock = 'Нет'
    }else{
        result.stock = 'Да'
    }
    
    results.push(result)
}

return results  
}
//Получение цен из xml
function get_prices_from_xml(filename){

const data = fs.readFileSync(filename,'utf8');
var options = {compact: true, spaces: 4};
var result1 = convert.xml2json(data, options);
result1 = JSON.parse(result1);

const offers = result1.yml_catalog.shop.offers.offer

var results = []
for (let i = 0; i < offers.length; i++){
    var offer = {}
    offer.pn = offers[i].vendorCode._text
    offer.price = offers[i].price._text
    results.push(offer)
}


return results

}
//Получение наличия из xml
function get_stock_from_xml(filename){

const data = fs.readFileSync(filename,'utf8');
var options = {compact: true, spaces: 4};
var result1 = convert.xml2json(data, options);
result1 = JSON.parse(result1);

const offers = result1.yml_catalog.shop.offers.offer

var results = []
for (let i = 0; i < offers.length; i++){
    var offer = {}
    offer.pn = offers[i].vendorCode._text
    if(offers[i].quantity._text != 0){
        offer.stock = 'Да'
    }else{
        offer.stock = 'Нет'
    }
    results.push(offer)
}


return results

}
//Задержка в одну секунду для печати всего документа  
function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

//Подключение к Spreadsheet
const { GoogleSpreadsheet } = require('google-spreadsheet');

async function SpreadsheetConnect(
    dxracer_prices_rrc,
    arroziKronox_prices_rrc,
    bf_prices_rrc,
    dxracer_prices,
    bf_prices,
    api_ocs_products,
    api_pronet_products,
    DXRacer_stock,
    foxgamer_stock,
    bf_cougar_stock,
    f5store_products
){
    const doc = new GoogleSpreadsheet('1hTBcj3zxiWWDGB7TWseOBggDSK-ts3s8FiGAnOmMMqU');
    const creds = require('./oauth/excel-database-370410-c4fefc8555ff.json');
    await doc.useServiceAccountAuth(creds);
    await doc.loadInfo();

    const sheet = doc.sheetsByTitle['База по креслам'];
    const rows = await sheet.getRows();
    const pns = rows.map(x => x['P/N'])
    const names = rows.map(x => x['Модель игрового кресла(name)'])
    const pnsOcs = rows.map(x => x['P/N OCS'])
    const pnsPronet = rows.map(x => x['P/N Pronet'])

    console.log('База по креслам ' + rows.length)

    //Подключение к таблице API

    const APiSheet = doc.sheetsByTitle['API'];
    var APIrows = await APiSheet.getRows();
    const pns_api = APIrows.map(x => x['P/N'])

    console.log('Api таблица ' + APIrows.length)

    var InfoArray = {
        dxracerIDX: 0,
        foxgamerIDX: 0,
        cougarIDX: 0,
        f5IDX: 0,
        ocsIDX: 0,
        pronetIDX: 0,

        dxracerStock: 0,
        foxgamerStock: 0,
        cougarStock: 0,
        fpStock: 0,
        ocsStock: 0,
        pronetStock: 0
    }

    for (let i = 0; i < rows.length; i++){

        let search_item = pns[i]

        if( i >= APIrows.length ){

            var newProduct = {}

            try{
                newProduct['Название кресел'] = names[i]
            }catch(e){newProduct['Название кресел'] = 'Не найден'}

            try{
                newProduct['P/N'] = pns[i]
            }catch(e){newProduct['P/N'] = 'Не найден'}

            try{
                newProduct['P/N OCS'] = pnsOcs[i]
            }catch(e){newProduct['P/N OCS'] = 'Не найден'}

            try{
                newProduct['P/N Pronet'] = pnsPronet[i]
            }catch(e){newProduct['P/N Pronet'] = 'Не найден'}

            try{
                let price = dxracer_prices_rrc.find(city => city.pn === search_item).price
                newProduct['РРЦ DXRacer'] = Number(price)
                InfoArray.dxracerIDX++
            }catch(e){
                newProduct['РРЦ DXRacer'] = 'Не найден'
            }

            try{
                let price = arroziKronox_prices_rrc.find(city => city.pn === search_item).price
                newProduct['РРЦ Foxgamer'] = Number(price)
                InfoArray.foxgamerIDX++
            }catch(e){
                newProduct['РРЦ Foxgamer'] = 'Не найден'
            }
            
            try{
                let price = bf_prices_rrc.find(city => city.pn === search_item).price
                newProduct['РРЦ Cougar/ Zone51'] = Number(price)
                InfoArray.cougarIDX++
            }catch(e){
                newProduct['РРЦ Cougar/ Zone51'] = 'Не найден'
            }

            try{
                let price = f5store_products.find(city => city.pn === search_item).rrcprice
                newProduct['РРЦ F5'] = Number(price)
                InfoArray.f5IDX++
            }catch(e){
                newProduct['РРЦ F5'] = 'Не найден'
            }

            try{
                let price = dxracer_prices.find(city => city.pn === search_item).price.result
                newProduct['Закупка DXRacer.su'] = Number(price)
            }catch(e){
                newProduct['Закупка DXRacer.su']  = 'Не найден'
            }
            
            try{
                let price = arroziKronox_prices_rrc.find(city => city.pn === search_item).price
                newProduct['Закупка Фоксгеймер'] = Number(price)
            }catch(e){
                newProduct['Закупка Фоксгеймер'] = 'Не найден'
            }

            try{
                let price = bf_prices.find(city => city.pn === search_item).price
                newProduct['Закупка Бизнес Фабрика'] = Number(price)     
            }catch(e){
                newProduct['Закупка Бизнес Фабрика'] = 'Не найден'
            }

            try{
                let price = api_ocs_products.find(city => city.sort_filed === search_item).price_ocs
                newProduct['Закупка OCS'] = Number(price) 
                InfoArray.ocsIDX++
            }catch(e){
                newProduct['Закупка OCS'] = 'Не найден'
            }
            
            try{
                let price = api_pronet_products.find(city => city.sort_filed === search_item).price
                newProduct['Закупка Pronet'] = Number(price) 
                InfoArray.pronetIDX++
            }catch(e){
                newProduct['Закупка Pronet'] = 'Не найден'
            } 
            
            try{
                let price = f5store_products.find(city => city.pn === search_item).price
                newProduct['Закупка F5store'] = Number(price) 
                newProduct['Наличие F5Store'] = 'Да'
                InfoArray.fpStock++           
            }catch(e){
                newProduct['Закупка F5store'] = 'Не найден'
                newProduct['Наличие F5Store'] = 'Не найден'
            } 

            try{
                let stock = DXRacer_stock.find(city => city.pn === search_item).stock
                newProduct['Наличие DXRacer.su'] = stock     
                InfoArray.dxracerStock++     
            }catch(e){
                newProduct['Наличие DXRacer.su'] = 'Не найден'
            }       

            try{
                let stock = foxgamer_stock.find(city => city.pn === search_item).stock
                newProduct['Наличие Foxgamer.ru'] = stock   
                InfoArray.foxgamerStock++
            }catch(e){
                newProduct['Наличие Foxgamer.ru'] = 'Не найден'
            }

            try{
                let stock = bf_cougar_stock.find(city => city.pn === search_item).stock
                newProduct['Наличие Бизнес Фабрика'] = stock 
                InfoArray.cougarStock++
            }catch(e){
                newProduct['Наличие Бизнес Фабрика'] = 'Не найден'
            }    

            try{
                let stock = api_ocs_products.find(city => city.sort_filed === search_item).quantity
                newProduct['Наличие OCS'] = stock 
                InfoArray.ocsStock++
            }catch(e){
                newProduct['Наличие OCS'] = 'Не найден'
            }
            
            try{
                let stock = api_pronet_products.find(city => city.sort_filed === search_item).quantity
                newProduct['Наличие Pronet'] = stock 
                InfoArray.pronetStock++
            }catch(e){
                newProduct['Наличие Pronet'] = 'Не найден'
            }  

            await APiSheet.addRow(newProduct)
                
                
        }else{

            APIrows[i]['Название кресел'] = names[i]
            APIrows[i]['P/N'] = pns[i]
            APIrows[i]['P/N OCS'] = pnsOcs[i]
            APIrows[i]['P/N Pronet'] = pnsPronet[i]

            try{
                let price = dxracer_prices_rrc.find(city => city.pn === search_item).price
                APIrows[i]['РРЦ DXRacer'] = Number(price)
                InfoArray.dxracerIDX++
            }catch(e){
                APIrows[i]['РРЦ DXRacer'] = 'Не найден'
            }

            try{
                let price = arroziKronox_prices_rrc.find(city => city.pn === search_item).price
                APIrows[i]['РРЦ Foxgamer'] = Number(price)
                InfoArray.foxgamerIDX++
            }catch(e){
                APIrows[i]['РРЦ Foxgamer'] = 'Не найден'
            }
            
            try{
                let price = bf_prices_rrc.find(city => city.pn === search_item).price
                APIrows[i]['РРЦ Cougar/ Zone51'] = Number(price)
                InfoArray.cougarIDX++
            }catch(e){
                APIrows[i]['РРЦ Cougar/ Zone51'] = 'Не найден'
            }

            try{
                let price = f5store_products.find(city => city.pn === search_item).rrcprice
                APIrows[i]['РРЦ F5'] = Number(price)
                InfoArray.f5IDX++
            }catch(e){
                APIrows[i]['РРЦ F5'] = 'Не найден'
            }

            try{
                let price = dxracer_prices.find(city => city.pn === search_item).price.result
                APIrows[i]['Закупка DXRacer.su'] = Number(price)
            }catch(e){
                APIrows[i]['Закупка DXRacer.su']  = 'Не найден'
            }
            
            try{
                let price = arroziKronox_prices_rrc.find(city => city.pn === search_item).price
                APIrows[i]['Закупка Фоксгеймер'] = Number(price)
            }catch(e){
                APIrows[i]['Закупка Фоксгеймер'] = 'Не найден'
            }

            try{
                let price = bf_prices.find(city => city.pn === search_item).price
                APIrows[i]['Закупка Бизнес Фабрика'] = Number(price)     
            }catch(e){
                APIrows[i]['Закупка Бизнес Фабрика'] = 'Не найден'
            }

            try{
                let price = api_ocs_products.find(city => city.sort_filed === search_item).price_ocs
                APIrows[i]['Закупка OCS'] = Number(price) 
                InfoArray.ocsIDX++
            }catch(e){
                APIrows[i]['Закупка OCS'] = 'Не найден'
            }
            
            try{
                let price = api_pronet_products.find(city => city.sort_filed === search_item).price
                APIrows[i]['Закупка Pronet'] = Number(price) 
                InfoArray.pronetIDX++
            }catch(e){
                APIrows[i]['Закупка Pronet'] = 'Не найден'
            } 
            
            try{
                let price = f5store_products.find(city => city.pn === search_item).price
                APIrows[i]['Закупка F5store'] = Number(price) 
                APIrows[i]['Наличие F5Store'] = 'Да'
                InfoArray.fpStock++           
            }catch(e){
                APIrows[i]['Закупка F5store'] = 'Не найден'
                APIrows[i]['Наличие F5Store'] = 'Не найден'
            } 

            try{
                let stock = DXRacer_stock.find(city => city.pn === search_item).stock
                APIrows[i]['Наличие DXRacer.su'] = stock     
                InfoArray.dxracerStock++     
            }catch(e){
                APIrows[i]['Наличие DXRacer.su'] = 'Не найден'
            }       

            try{
                let stock = foxgamer_stock.find(city => city.pn === search_item).stock
                APIrows[i]['Наличие Foxgamer.ru'] = stock   
                InfoArray.foxgamerStock++
            }catch(e){
                APIrows[i]['Наличие Foxgamer.ru'] = 'Не найден'
            }

            try{
                let stock = bf_cougar_stock.find(city => city.pn === search_item).stock
                APIrows[i]['Наличие Бизнес Фабрика'] = stock 
                InfoArray.cougarStock++
            }catch(e){
                APIrows[i]['Наличие Бизнес Фабрика'] = 'Не найден'
            }    

            try{
                let stock = api_ocs_products.find(city => city.sort_filed === search_item).quantity
                APIrows[i]['Наличие OCS'] = stock 
                InfoArray.ocsStock++
            }catch(e){
                APIrows[i]['Наличие OCS'] = 'Не найден'
            }
            
            try{
                let stock = api_pronet_products.find(city => city.sort_filed === search_item).quantity
                APIrows[i]['Наличие Pronet'] = stock 
                InfoArray.pronetStock++
            }catch(e){
                APIrows[i]['Наличие Pronet'] = 'Не найден'
            } 

            await APIrows[i].save()
        }

        await delay(1000);
    }

    try{
        const infoSheet = doc.sheetsByTitle['Проверка'];
        const rowsNew = await infoSheet.getRows();

        rowsNew[0]['DxRacer общее'] = InfoArray.dxracerIDX
        rowsNew[0]['Foxgamer общее'] = InfoArray.foxgamerIDX
        rowsNew[0]['БФ общее'] = InfoArray.cougarIDX
        rowsNew[0]['F5 общее'] = InfoArray.f5IDX
        rowsNew[0]['OCS общее'] = InfoArray.ocsIDX
        rowsNew[0]['Pronet общее'] = InfoArray.pronetIDX

        rowsNew[0]['DxRacer наличие'] = InfoArray.dxracerStock
        rowsNew[0]['Foxgamer наличие'] = InfoArray.foxgamerStock
        rowsNew[0]['БФ наличие'] = InfoArray.cougarStock
        rowsNew[0]['F5 наличие'] = InfoArray.fpStock
        rowsNew[0]['OCS наличие'] = InfoArray.ocsStock
        rowsNew[0]['Pronet наличие'] = InfoArray.pronetStock

        await rowsNew[0].save()

    }catch(e){
        console.log(e)
    }

}





module.exports.pronet_get_token = pronet_get_token(pronet_login, pronet_password)
module.exports.osc_get_product_information = osc_get_product_information 
module.exports.html_parse = html_parse
module.exports.pronet_get_product_information = pronet_get_product_information
module.exports.get_xlsx_data = get_xlsx_data
module.exports.get_prices = get_prices
module.exports.get_stock = get_stock
module.exports.get_prices_from_xml = get_prices_from_xml
module.exports.get_stock_from_xml = get_stock_from_xml
module.exports.SpreadsheetConnect = SpreadsheetConnect
module.exports.create_OCS_Pronet_object = create_OCS_Pronet_object
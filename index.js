//Настройка запросов для управления базы данных кресел
const import_main = require("./spreadsheet.js");

//Настройки хранения файлов
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, "export")
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname)
    }
})

var uploadFields = multer({ storage : storage }).fields([
    { name: 'exportFilesDxracer', maxCount: 1 },
    { name: 'exportFilesCougar', maxCount: 1 },
    { name: 'exportFilesFoxgamer', maxCount: 1 },
	{ name: 'exportFilesarroziKronox', maxCount: 1 },
  ]);

//Главная страница сервера
//Ренденр главной страницы загрузки прайсов http://89.108.65.126:3000/excel/
app.get('/excel', function (request, response) {
    response.render('sd_index');
})

//Обновить файлы
// Рендер страницы приложения файлов для обновления прайсов http://89.108.65.126:3000/excel/update
app.get('/excel/update', function (request, response) {
    response.render('sd_upload');
})

//Обновление файлов и главной таблицы
// POST ответ, на приложеннные файлы с GET запроса http://89.108.65.126:3000/excel/update
app.post('/excel/update', async (req, res) =>{

	try{

		uploadFields(req,res,function(err) { //Получение приложенных файлов
			if(err) {
				console.log("Error uploading file Foxgamer. " + err);
			}
			console.log("Files are uploaded");
		});
        //Подключение к Spreadsheet документу
		const doc = new GoogleSpreadsheet('1hTBcj3zxiWWDGB7TWseOBggDSK-ts3s8FiGAnOmMMqU');
		const creds = require('./oauth/excel-database-370410-c4fefc8555ff.json');
		await doc.useServiceAccountAuth(creds);
		await doc.loadInfo();
		
		const sheet = doc.sheetsByTitle['База по креслам'];
		const rows = await sheet.getRows(); //Получение всех строк на листе

		const pnsDatabase = await import_main.create_OCS_Pronet_object(rows)


		var ocs_products_ids = pnsDatabase.map(data => data.ocs) //Получение партномеров OCS
		var pronet_products_ids = pnsDatabase.map(data => data.pronet).join() //Получение партномеров Pronet
		var products_pn = pnsDatabase.map(data => data.pn) //Получение партномеров

		const ocs_products = await import_main.osc_get_product_information(ocs_products_ids, products_pn) //Получение продуктов с ocs

		const f5store_products = await import_main.html_parse() //Получение продуктов с f5store
	

		var tokenData = await import_main.pronet_get_token //Получение токена для pronet
		const token = tokenData.data.Token
	
		const pronet_products = await import_main.pronet_get_product_information(token, pronet_products_ids, ocs_products) //Получение продуктов с pronet
	
		

		//Получение данных с DXRacer
		const dxracerData = await import_main.get_xlsx_data('export/dxracer.xlsx')
		var DXRacer_prices_rrc = import_main.get_prices('DXRacer', 2 ,dxracerData, 3, 'Артикул', 'РРЦ, руб') //Получение цен с прайс листа DXRacer
		var DXRacer_prices = import_main.get_prices('DXRacer', 2 ,dxracerData, 3, 'Артикул', 'Цена со скидкой, руб') //Получение цен с прайс листа DXRacer
		var DXRacer_stock = import_main.get_stock('DXRacer', 2 ,dxracerData, 3, 'Артикул', 'Наличие*') //Получение наличия с прайс листа DXRacer
		//Arozzi Карнокс РРЦ
		const arozziKronoxData = await import_main.get_xlsx_data('export/arroziKronox.xlsx')
		var arozzi_prices_rrc = import_main.get_prices('fullProducts', 1 ,arozziKronoxData, 2, 'Артикул', 'РРЦ') //Получение цен с прайс листа arozzi и Kronox
		var foxgamer_stock = import_main.get_stock('fullProducts', 1 ,arozziKronoxData, 2, 'Артикул', 'Наличие')  //Получение наличия с прайс листа foxgamer
		//Получение данных с Cougar
		const BFData = await import_main.get_xlsx_data('export/cougar.xlsx')
		var bf_prices_rrc = import_main.get_prices('Cougar', 1 ,BFData, 2, 'Код', 'ЦЕНА РРЦ') //Получение цен ррц с прайс листа Cougar
		var bf_prices = import_main.get_prices('Cougar', 1 ,BFData, 2, 'Код', 'Цена ОПТ') //Получение цен с прайс листа Cougar
		var bf_cougar_stock = import_main.get_stock('Cougar', 1 ,BFData, 2, 'Код', 'Остаток') //Получение остатка с прайс листа Cougar

		import_main.SpreadsheetConnect(
			DXRacer_prices_rrc, //РРЦ цена DXRacer
			arozzi_prices_rrc, //РРЦ цена Arozzi и Kronox
			bf_prices_rrc, //РРЦ цена Cougar и Zone 51
		
			DXRacer_prices, // цена DXRacer
			bf_prices, // Цена Бизнес фабрика
			ocs_products, // Цена OCS
			pronet_products, // Цена Pronet
		
			DXRacer_stock, //Наличие DxRacer
			foxgamer_stock, //Наличие FoxGamer
			bf_cougar_stock, // Наличие Бизнес Фабрика
		
			f5store_products //F5store   
			)
			
		console.log('finish')
		res.end('sucess')
	}catch(e){
		console.log(e)
		res.json(e)
	}

})




//Создание YML фида
app.get('/excel/yml', async (req, res) =>{

	try{
		const doc = new GoogleSpreadsheet('1hTBcj3zxiWWDGB7TWseOBggDSK-ts3s8FiGAnOmMMqU');
		const creds = require('./oauth/excel-database-370410-c4fefc8555ff.json');
		await doc.useServiceAccountAuth(creds);
		await doc.loadInfo();
		const sheet = doc.sheetsByTitle['База по креслам'];
		const rows = await sheet.getRows();  

		//Получение всех продуктов с листа       
		var products = []	
		
		for (let i = 0; i < rows.length; i++){


			if( rows[i]['Ссылка на сайт(url)'] == '' || rows[i]['НАЛИЧИЕ ИЗ API ТАБЛИЦЫ(api_stock)'] == 'нет' || rows[i]['Изображение товара(pictures)'] == '' ) continue

			var product = {
				'@id': rows[i]['P/N']
			}

			if(rows[i]['ЦЕНА ИЗ API ТАБЛИЦЫ(api_price)']){
				let price = rows[i]['ЦЕНА ИЗ API ТАБЛИЦЫ(api_price)']
				price = price.split(/\s+/).join('')
				product.price = price

			} 
			if(rows[i]['Ссылка на сайт(url)']) product.url = rows[i]['Ссылка на сайт(url)']
			product.currencyId = 'RUR'
			product.categoryId = 1
			if( rows[i]['Изображение товара(pictures)'] ) product.picture = rows[i]['Изображение товара(pictures)']
			if( rows[i]['Описание(description)'] ) product.description = rows[i]['Описание(description)']   
			product.delivery = true		
			if( rows[i]['Модель игрового кресла(name)'] && rows[i]['Серия(seria)'] ) product.name = rows[i]['Модель игрового кресла(name)'] + ' ' + rows[i]['Серия(seria)']
			if( rows[i]['Модель(model)'] ) product.model = rows[i]['Модель(model)']
			if( rows[i]['Бренд(brand)'] ) product.vendor = rows[i]['Бренд(brand)']
			if( rows[i]['P/N'] ) product.vendorCode = rows[i]['P/N']   
			product.typePrefix = 'Компьютерное игровое кресло' 
			product.manufacturer_warranty = true
			product.available = true
			product.param = []
			if( rows[i]['Цвет(color)'] ) product.param.push({'@name': 'Цвет', '#text': rows[i]['Цвет(color)']})
			if( rows[i]['Материал отделки(material_otd)'] ) product.param.push({'@name': 'Материал отделки', '#text': rows[i]['Материал отделки(material_otd)']})
			if( rows[i]['Механизм качания(mechanism_kach)'] ) product.param.push({'@name': 'Механизм качания', '#text': rows[i]['Механизм качания(mechanism_kach)']})
			if( rows[i]['Рекомендуемый вес пользователя(recommended_ves)'] ) product.param.push({'@name': 'Рекомендуемый вес пользователя', '#text': rows[i]['Рекомендуемый вес пользователя(recommended_ves)']})
			if( rows[i]['Рекомендуемый рост пользователя(recommended_rost)'] ) product.param.push({'@name': 'Рекомендуемый рост пользователя', '#text': rows[i]['Рекомендуемый рост пользователя(recommended_rost)']})
			if( rows[i]['Макс. вес пользователя(max_ves)'] ) product.param.push({'@name': 'Максимальный вес пользователя', '#text': rows[i]['Макс. вес пользователя(max_ves)']})
			if( rows[i]['Макс. рост пользователя(max_rost)'] ) product.param.push({'@name': 'Максимальный рост пользователя', '#text': rows[i]['Макс. рост пользователя(max_rost)']})
			if( rows[i]['Тип подлокотников(armrest type)'] ) product.param.push({'@name': 'Тип подлокотников', '#text': rows[i]['Тип подлокотников(armrest type)']})
			if( rows[i]['Угол наклона спинки (Max)'] ) product.param.push({'@name': 'Угол наклона спинки (Max)', '#text': rows[i]['Угол наклона спинки (Max)']})
			if( rows[i]['Наличие подушек(pillows_stock)'] ) product.param.push({'@name': 'Наличие подушек', '#text': rows[i]['Наличие подушек(pillows_stock)']})
			if( rows[i]['Крепление поясничной подушки(lumbar_support)'] ) product.param.push({'@name': 'Крепление поясничной подушки', '#text': rows[i]['Крепление поясничной подушки(lumbar_support)']})
			if( rows[i]['Наличие регулируемой поясничной поддержки(lumbar_support_stock)'] ) product.param.push({'@name': 'Наличие регулируемой поясничной поддержки', '#text': rows[i]['Наличие регулируемой поясничной поддержки(lumbar_support_stock)']})
			if( rows[i]['Газлифт(gaslift)'] ) product.param.push({'@name': 'Газлифт', '#text': rows[i]['Газлифт(gaslift)']})
			if( rows[i]['Материал крестовины(material_cross)'] ) product.param.push({'@name': 'Материал крестовины', '#text': rows[i]['Материал крестовины(material_cross)']})
			if( rows[i]['Размер колесиков кресла(wheel_size)'] ) product.param.push({'@name': 'Размер колесиков кресла', '#text': rows[i]['Размер колесиков кресла(wheel_size)']})
			if( rows[i]['Гарантия(garanty)'] ) product.param.push({'@name': 'Гарантия', '#text': rows[i]['Гарантия(garanty)']})
			if( rows[i]['Вес брутто(ves_brutto)'] ) product.param.push({'@name': 'Вес брутто', '#text': rows[i]['Вес брутто(ves_brutto)']})
			if( rows[i]['Вес нетто'] ) product.param.push({'@name': 'Вес нетто', '#text': rows[i]['Вес нетто']})
			if( rows[i]['Габариты упаковки (ДxШxВ)(gabarity)'] ) product.param.push({'@name': 'Габариты упаковки (ДxШxВ)', '#text': rows[i]['Габариты упаковки (ДxШxВ)(gabarity)']})
	
			products.push(product)
		}		

		var obj = {
			yml_catalog: 
			{
				'@date': new Date().toISOString(),
				shop: {
					name: { '#text': 'Boiling Machine'},                    
					company: { '#text': 'Boiling Machine'},
					url: { '#text': 'https://boiling-machine.ru/'},
					platform: { '#text': 'Wordpress'},
					currencies: {
						currency: {
							'@id': 'RUB',
							'@rate': 1
						}
					},
					categories:{
						category: {
							'@id': '1',
							'#text': 'Игровые кресла'
						}
					},
					offers: {
						offer: []
					}
					
				}
			}
		}
		
		for ( let i = 0; i < products.length; i++){
			obj.yml_catalog.shop.offers.offer.push(products[i])
		}
		
		var xml = builder.create(obj,{version: '1.0', encoding: 'UTF-8', date: new Date()}).end({ pretty: true});
	
		res.header("Content-Type", "application/xml");
		res.status(200).send(xml);

	}catch(e){
		console.log(e)
		res.json(e)
	}
})
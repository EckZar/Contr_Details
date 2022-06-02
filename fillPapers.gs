function onOpen(){

  SpreadsheetApp.getUi().createMenu("__MENU__")
  .addItem("oAuth", "showSidebar")
  .addItem("Обновить список организаций", "fillCompany")
  .addItem("Обновить «Статья БК»", "fillCostItem")
  .addItem("Обновить «Статья ДДС»", "fillCashFlowItem")
  .addItem("Обновить «ОИДП»", "fillObject")
  .addItem("Обновить «ЦФО по БФ»", "fillFinancialResponsibilityCenterForBF")
  .addItem("Обновить «ЦФО-заказчик»", "fillFinancialResponsibilityCenter")
  .addItem("Обновить «Контрагентов»", "fillContragents")
  .addToUi()
}

//  Вставляем array на лист sheetName в текущей таблице. 
//  Предварительно очищаем лист sheetName, по размеру вставляемого массива, перед тем как вставить массив. 
//  Очистка идет от 2 ячейки первого столбца по ширине массива, до конца таблицы.
function pasteArr(sheetName, array){ 
  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  try
  {
    main.getRange(2, 1, main.getLastRow()-1, array[0].length).clear();
  }
  catch(e){}
  main.getRange(2, 1, array.length, array[0].length).setValues(array);
}

function fillCompany(){ // Обновляем лист организации
  
  let hide = SpreadsheetApp.openById("").getSheetByName("Организации в периметре консолидации");
  let range = hide.getRange(2, 1, hide.getLastRow()-1, 3).getValues();
  pasteArr("Организация", range);  

}



function fillCostItem(){ // Статья БК

  let accessToken = getDriveService().getAccessToken();
  let costItem = getRequest(accessToken, "api/v2/CostItem", 1000, 0);
  pasteArr("Статья БК", costItem.data.filter(function(e){ return e.isGroup === false && !e.deleted}).map(function(arr){ return [arr.name, arr.isIncome ? "Приход":"Расход",arr.guid1C]}));
  
}


function fillCashFlowItem(){ // Статья ДДС

  let service = getDriveService();
  let accessToken = service.getAccessToken();
  
  let cashFlowItem = getRequest(accessToken, "api/v2/CashFlowItem", 1000, 0); // Получаем список статей. MainActivityKind содержит UUID
  let activityKind = getRequest(accessToken, "api/v2/ActivityKind", 1000, 0); // Получаем список наименований MainActivityKind

  activityKind = activityKind.data;
  
  pasteArr("Статья ДДС", cashFlowItem.data // В массиве объектов cashFlowItem
  .filter(function(e) // фильтруем массив объектов где isGroup false
    {
      return e.isGroup === false && !e.deleted;
    })
  .map(function(arr) // из отфильтрованного массива объектов возвращаем значения name, isIncome, activityKind, guid1C
    {
      return [arr.name, 
              arr.isIncome ? "Приход":"Расход", 
              activityKind
                .filter(function(q) // возвращаем объект с совпадающим ключом
                  {
                    return q.guid1C === arr.mainActivityKind;
                  })
                .map(function(obj) // возвращаем только наименование
                  {
                    return obj.name;
                  }), 
              arr.guid1C]
    }));  

}

function fillObject(){ // ОИДП

  let service = getDriveService();
  let accessToken = service.getAccessToken()
  
  let dateObj = new Date(2018, 0, 1); 
  let date = new Date(2020, 0, 1); 

  let company = getRequest(accessToken, "api/v2/Company", 1000, 0);
  // pasteArr("Организация", company.data.map(function(arr){ return [arr.name, arr.guid1C]}));

  


  let objectName = getRequest(accessToken, "api/v2/ObjectOther", 3000, 0);

  let array = objectName.data
    .filter(function(e)
    {      
      let compDate = new Date(e.endDate);
      return compDate >= date || !e.endDate && !e.isGroup;
    })
    .map(function(obj)
    {
      return [obj.name, 
              company.data
                .filter(function(e)
                { 
                  return e.guid1C === obj.companyGuid1C;
                })
                .map(function(objj)
                {
                  return objj.name;
                }),              
              obj.hierarchicalCode, 
              obj.guid1C,
              "ObjectOther"]
    });  


  let object = getRequest(accessToken, "api/v2/Object", 1, 0);
  let meta = object.meta;
  
  let end = Math.floor(meta.totalCount / 1000) + 1; 
  
  for( let i = 0 ; i<end ; i++ )
  {

    let offset = i * 1000;
 
    let object = getRequest(accessToken, "api/v2/Object", 1000, offset);
    
    let tempArr = object.data
                    .filter(function(e) // Из объектов фильтруем даты начала старше 1 января 2018 года
                    {
                      
                      let compDate = new Date(e.startDate);                  

                      return (compDate >= dateObj || e.guid1C === "e6175738-e0d7") && !e.deleted;

                    })
                    .map(function(arr)
                    {
                      return [arr.name,
                              objectName.data.filter(function(q)
                              {
                                return q.guid1C === arr.guid1C; 
                              })
                              .map(function(obj)
                              {
                                return obj.name;
                              }),
                              arr.code1C,
                              arr.guid1C,
                              "Object"]
                    });


    

    if(tempArr.length>0)
    {
      array = array.concat(tempArr);
    }  
  }

  pasteArr("ОИДП", array);

}


function fillFinancialResponsibilityCenterForBF(){ // ЦФО по БФ

  let service = getDriveService();
  let accessToken = service.getAccessToken()
  
  let finance = getRequest(accessToken, "api/v2/FinancialResponsibilityCenterForBF", 1000, 0);
  finance = finance.data.filter(function(e){ return !e.deleted && !e.isGroup}).map(function(obj){ return [obj.name, obj.guid1C]});

  pasteArr("ЦФО по БФ", finance);

}


function fillFinancialResponsibilityCenter(){ // ЦФО заказчик

  hide = SpreadsheetApp.openById("1L0WJqa30").getSheetByName("ЦФО-заказчик - добавить");
  let specialTwo = hide.getRange(2, 1, hide.getLastRow()-1, 2).getValues();

  let service = getDriveService();
  let accessToken = service.getAccessToken()
  
  let finance = getRequest(accessToken, "api/v2/FinancialResponsibilityCenter", 1000, 0);
  finance = finance.data.filter(function(e){ return !e.deleted && !e.isGroup}).map(function(obj){ return [obj.name, obj.guid1C]}).concat(specialTwo);
  

  pasteArr("ЦФО-заказчик", finance);

}


function fillContr(){ // Контрагенты

  let service = getDriveService();
  let accessToken = service.getAccessToken()
  
  let contr = getRequest(accessToken, "api/v2/ContractorInPerimetr", 1500, 0);  
  
  let guids = contr.data;
  
  let uniques = guids.map(function(obj)
                    {
                      return obj.contractorGuid1C
                    })
                .filter(onlyUnique)
                .map(function(arr)
                    { 
                      return arr = { 
                                    "contractorGuid1C":arr,
                                    "dates":[]
                                  }
                    });


  uniques.forEach(function(item, i)
                  {

                    guids.filter(function(e)
                                {
                                  return e.contractorGuid1C === item.contractorGuid1C;                                  
                                })
                         .forEach(function(jtem)
                                {
                                  item.dates.push(jtem.period);                                  
                                });

                  });

  uniques.forEach(function(item, i)
                  {                   
                      item.dates = new Date(Math.max.apply(null,item.dates));
                  });

  contr = guids.filter(function(e)
                      {
                        return uniques.filter(function(q)
                                              { 
                                                return q.companyGuid1C == e.companyGuid1C && new Date(q.period) == new Date(e.dates)
                                              }) 
                                && e.inPerimetr
                      })
                      .map(function(obj)
                      { 
                        return obj.contractorGuid1C;
                      });

  Logger.log(contr.length)

  //================================================================
  
  let array = getContractors(contr);


  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Контрагент");
  try{
    main.getRange(101, 1, main.getLastRow()-100, 4).clear();
  }
  catch(e){}
  
   main.getRange(101, 1, array.length, 4).setValues(array);
}



function fillContragents(){

  let fromS = SpreadsheetApp.openById("1L0WJqKtw4VYva30").getSheetByName("Контрагенты");
  let range = fromS.getRange(2, 1, fromS.getLastRow()-1, 4).getValues();

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Контрагент");

  try
  {
    main.getRange(101, 1, main.getLastRow()-100, range[0].length).clear();
  }
  catch(e){}

  main.getRange(101, 1, range.length, range[0].length).setValues(range);

}



function onlyUnique(value, index, self) {
    return self.indexOf(value) === index;
}










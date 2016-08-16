// Define constants
// http://mis.twse.com.tw/stock/api/getStock.jsp?ch=t00.tw
// http://mis.twse.com.tw/stock/api/getStock.jsp?ch=2330.tw
var EXPECTED_PRICE_1 = 16;
var EXPECTED_PRICE_2 = 20;
var EXPECTED_PRICE_3 = 30;

var TWSE_HOST           = "mis.twse.com.tw";
var STOCK_REVENUE       = mainUrl+"/z/zc/zch/zch_";  // 營收
var STOCK_PROFILE_URL   = mainUrl+"/z/zc/zcx/zcx_";  // 基本資料
var STOCK_EPS_URL       = mainUrl+"/z/zc/zcc/zcc_";  // 股利
var STOCK_NEWS          = mainUrl+"/z/zc/zci/zci_";  // 重大行事曆
var STOCK_INFO_URL      = "http://isin.twse.com.tw/isin/class_main.jsp?market=1&issuetype=1&Page=1&chklike=Yhklike=Y";
var ALL_AFTER_TRADNING_URL = 'http://goodinfo.tw/StockInfo/StockDividendPolicyList.asp?MARKET_CAT=%E5%85%A8%E9%83%A8&INDUSTRY_CAT=%E8%B3%87%E8%A8%8A%E6%9C%8D%E5%8B%99%E6%A5%AD&YEAR=2015';

var START_COLUMN    = 3;  // 從 C 欄開始放股票代號
var STOCK_NAME      = 1;   
var STOCK_NO_ROW    = 2;  // 股票代號
var PROFILE_ROW     = 3;  // 個股速覽
var REVENUE_ROW     = 4;  // 營收
var NEWS_ROW        = 5;  // 重大新聞
var EPS_ROW         = 6;  // 股利
var PRICE_ROW       = 8;  // 股價
var START_ROW       = 3;

var stockInfo = [  
  { name: "股票代號", id:"c" },
  { name: "股票名稱", id:"n" },
  { name: "個股速覽", id: "profileLink" },
  { name: "重大行事曆", id: "newsLink" },  
  { name: "營收", id: "revenueLink" },
  { name: "股利", id: "dividendLink" },  
  { name: "今日價格", id: "z" },
  { name: "漲跌", id: "change" },
  { name: "5年平均\n現金股利", id: "avgDividend5Years" },
  { name: "便宜價", id: "expectedPrice1" },
  { name: "合理價", id: "expectedPrice2" },
  { name: "昂貴價", id: "expectedPrice3" },
  { name: "股本(億)\n(>50億)", id: "" },
  { name: "每股淨值\n(>15)", id: "" },
  { name: "營業毛利率\n(>20%)", id: "" },
  { name: "股東權益\n報酬率(>5%)", id: "" },
  { name: "股價/淨值比\n(<2)", id: "priceBookRatio" },
  { name: "本益比\n(<12)", id: "priceEarningsRatio" },
  { name: "負債比例\n(<50%)", id: "" },
  { name: "抵稅率\n(>20%)", id: "" },
  { name: "今年\n殖利率", id: "cashDividendYield" },
  { name: "5年平均\n殖利率", id: "" },
];

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = SpreadsheetApp.getActiveSheet();
  
function renderStocks() {
  
  var title = [];
  for (var i=0; i<stockInfo.length; i++) {
    title.push(stockInfo[i].name); 
  }

  sheet
    .getRange(1, 1, 1, stockInfo.length)
    .setValues([title])
    .setHorizontalAlignment("center");
 
  var row = 1;  
  var stocks = [];
  var startRow = 0;
  while (1) {
    if (row > 1) {
      startRow = row;
      stockNoRange = sheet.getRange(row, 1, 1, 1);
      var stockNo = stockNoRange.getDisplayValue();
      if (stockNo == "")
        break;    
      stocks.push(stockNo);
    }
    row++;
  }
  
//  var currentStocks = buildFakeStocks(stocks);
//  var currentStocks = getCurrentPrice(stocks);
  var currentStocks = getCurrentStockStats(stocks);
  if (currentStocks) {
    var pasteRange = [];
    for (var i=0; i<currentStocks.length; i++) {
      var cs = currentStocks[i];
      var stockRange = [];
      for (var j=0; j<stockInfo.length; j++) {
        var si = stockInfo[j];
        if (si.id != '') {
          stockRange.push(cs[si.id]);
        } else {
          stockRange.push("");
        }
      }
      pasteRange.push(stockRange);
      
      //if (!ss.getSheetByName(cs.n))
      //  ss.insertSheet(cs.n);
    }
    
    var firstSheet = ss.getSheetByName("總覽");
    if (!firstSheet) {
      firstSheet = ss.getSheets()[0];
    }
    firstSheet.activate(); //跳回總覽 or 第一個 sheet
    Logger.log('length = %s %s %s', stocks.length, stockInfo.length, pasteRange.length);
    sheet
      .getRange(2, 1, stocks.length, stockInfo.length)
      .setHorizontalAlignment("center")
      .setValues(pasteRange);
  }
}

function demoMain() {
  var x = JSON.parse(getProperty('2606'))
  var y = calculateAverageDividend(x.dividends);
  Logger.log(y);
}

function getCurrentStockStats(stocks) {
  var stockMap = {};
  var scriptProperties = PropertiesService.getScriptProperties();
  for (var i=0; i<stocks.length; i++) {
    var stockProps = getProperty(stocks[i]);    
    if (stockProps) {
      stockProps = JSON.parse(stockProps);
    } else {
      stockProps = {
        'stockId': stocks[i],
        'stockName': '',
        'z': '',
        'change': '',
        'profile': '',
      }
    }
    var data = {};
    data['c'] = stockProps.stockId;
    data['n'] = stockProps.stockName;
    data['z'] = stockProps.z;
    data['change'] = stockProps.change;
    data['avgDividend5Years'] = calculateAverageDividend(stockProps['dividends']);
    data['cashDividendYield'] = stockProps.cashDividendYield;
    data['priceEarningsRatio'] = stockProps.priceEarningsRatio;
    data['priceBookRatio'] = stockProps.priceBookRatio;
    data['dividend'] = stockProps['dividends']['2015'].toFixed(3);
    data['expectedPrice1'] = data['avgDividend5Years'] ? (EXPECTED_PRICE_1 * data['avgDividend5Years']).toFixed(2) : '-';
    data['expectedPrice2'] = data['avgDividend5Years'] ? (EXPECTED_PRICE_2 * data['avgDividend5Years']).toFixed(2) : '-';
    data['expectedPrice3'] = data['avgDividend5Years'] ? (EXPECTED_PRICE_3 * data['avgDividend5Years']).toFixed(2) : '-';
    
    var stock = new Stock(stockProps.stockId, stockProps.stockName);
    data['profileLink'] = stock.getProfileLink();
    data['newsLink'] = stock.getNewsLink();
    data['revenueLink'] = stock.getRevenueLink();
    data['dividendLink'] = stock.getDividentLink();
    //data['cashDividendYield'] = stockProps['dividents']['2015'];
    stockMap[stocks[i]] = data;      
  }
  
  var currentPrice = [];
  for(var key in stockMap) {
    currentPrice.push(stockMap[key]);
  }
  
  if (currentPrice != null) {
    return currentPrice;
  }
  
  return currentPrice;
}

function getCurrentPrice(stocks) {
    try {
      // Use IP directly to avoid DNS failure
      //var host = "163.29.17.179";
      var host = "mis.twse.com.tw";
      
      // Poke TWSE homepage to get session id, e.g.JSESSIONID=5B050F7AF3A3CD64091F772D7D589A82; Path=/stock
      var resForSession = UrlFetchApp.fetch("http://"+host+"/stock/index.jsp?lang=zh-tw");     
      var headers = {
        "Cookie" : resForSession.getHeaders()["Set-Cookie"]
      }
      
      var options = {
        "method" : "get",
        "headers" : headers
      };
         
      var stockMap = {};
      var ex_ch = "";
      for (var i=0; i<stocks.length; i++) {
        ex_ch += "tse_" + stocks[i] + ".tw%7C";
        ex_ch += "otc_" + stocks[i] + ".tw%7C";        
        stockMap[stocks[i]] = 0;
      }

      var afterTradingInfo = getAfterMarketTrading();
      Logger.log("http://"+host+"/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch=" + ex_ch);
      var response = UrlFetchApp.fetch("http://"+host+"/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch=" + ex_ch, options);
      var jsonData = JSON.parse(response.getContentText());            
      var scriptProperties  = PropertiesService.getScriptProperties();

      for (var i=0; i<jsonData.msgArray.length; i++) {
        var data = jsonData.msgArray[i];
        var stockId = data.c;
        var stockProps = scriptProperties.getProperty(stockId);
     
        data['profile'] = "=HYPERLINK(\"" + yuantaInfo['profile'].url + data.c + ".djhtm\",\"" + yuantaInfo['profile'].name + "\")"
        data['news'] = "=HYPERLINK(\"" + yuantaInfo['news'].url + data.c + ".djhtm\",\"" + yuantaInfo['news'].name + "\")"
        data['revenue'] = "=HYPERLINK(\"" + yuantaInfo['revenue'].url + data.c + ".djhtm\",\"" + yuantaInfo['revenue'].name + "\")"
        data['eps'] = "=HYPERLINK(\"" + yuantaInfo['eps'].url + data.c + ".djhtm\",\"" + yuantaInfo['eps'].name + "\")"
        
        for (var k in afterTradingInfo[data.c]) {
          data[k] = afterTradingInfo[data.c][k];
        }
        Logger.log(stockProps)
        data['cashDividendYield'] = stockProps['dividents']['2015'];
        stockMap[data.c] = data;        
      }
      
      var currentPrice = [];
      for(var key in stockMap) {
        currentPrice.push(stockMap[key]);
      }
      
      if (currentPrice != null) {
        return currentPrice;
      }
    } catch (error) {
      /* Somwthing's wrong... */
      Logger.log(error);
      ss.toast(error);
      return false;
    }
    return false;
}

function getAllStockAfterTrading() {
  var html = UrlFetchApp.fetch(ALL_AFTER_TRADNING_URL).getContentText('UTF-8');
  var pattern = /<div id="divDetail"[^>]*>((.|[\n\r])*)<\/div>/im;
  var t = pattern.exec(html);
  
  var doc = XmlService.parse(html);
  var html = doc.getRootElement();
  var matches = {};
  var atom = XmlService.getNamespace('http://www.w3.org/2005/Atom'); 
  var trElements = getElementsByTagName(html, 'tr');
  for (var i=0; i<trElements.length; i++) {
    var tr = trElements[i];
    var labels = [];
    var tdElements = getElementsByTagName(tr, 'td');
    if (tdElements.length >= 6) {
      var stockId = tdElements[0].getText();
      var dividendPerShare = tdElements[3].getText();
      var cashDividendYield = tdElements[4].getText();
      var priceBookRatio = tdElements[5].getText();
  
      matches[stockId] = {
        stockId: stockId,
        dividendPerShare: dividendPerShare,
        cashDividendYield: cashDividendYield,
        priceBookRatio: priceBookRatio,  // 股價淨值比
      };
    }
  }
  
  return matches;
}

function goToSheet(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  SpreadsheetApp.setActiveSheet(sheet);
}

function onOpen() { 
  //renderStockProfile();
  renderStocks();
}

function onChange(e){
  Logger.log("onChange event fired " + e.changeType);
  if (e.changeType == 'EDIT') {
    renderStocks();
  }
}

function onFormSumit(e) {
  Logger.log("onFormSumit event fired " + e.changeType);
}

function onEdit(e) {
  Logger.log("onEdit event fired " + e.value);
    
  if (e.value != e.oldValue) {  
    checkCellValue(e);
  }
  
  if (parseInt(e.value)) {
      renderStocks();
  }
}

function onClickHandler(e) {
  Logger.log("onClickHandler event fired " + e.value);
}

function onMouseClick(e) {
  Logger.log("onMouseClick event fired " + e.value);
}

function checkCellValue(e) {
  if (!e.value || e.value == 0) {  // Delete if value is zero or empty
    e.source.getActiveSheet().deleteRow(e.range.rowStart);
  }
}


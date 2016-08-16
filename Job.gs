var MAX_NUM_SEND_YQL = 100;
var AFTER_MARKET_1_URL  = 'http://www.twse.com.tw/ch/trading/exchange/BWIBBU/BWIBBU_d.php';
var AFTER_TRADNING_URL  = 'http://www.tpex.org.tw/web/stock/aftertrading/peratio_analysis/pera_print.php';
var DIVIDEND_HISTORY_URL = 'http://stock.wespai.com/p/5625'; // 撿股網-歷年股利股息
var TEN_RATE_URL         = 'http://stock.wespai.com/tenrate'; // 撿股網-近10年配息配股

// 上市公司 即時資訊
function mergeCurrentPrice(stocks) {
  var yfinance = "http://query.yahooapis.com/v1/public/yql?" +
                 "format=json&env=store://datatables.org/alltableswithkeys&" +
                 "q=select * from yahoo.finance.quote where symbol in ";
  
  var maps = [];
  var i = 1;
  for (var key in stocks) {
    maps.push(key + '.TW');
    var has_next_page = (Object.size(stocks) / MAX_NUM_SEND_YQL) != (i / MAX_NUM_SEND_YQL);
    if ((i % MAX_NUM_SEND_YQL == 0) || !has_next_page) {
      if (maps.length == 1) {
        maps.push('2002.TW');
      }
      var url = yfinance + "('" + maps.join("','") + "')";
      var html = UrlFetchApp.fetch(url).getContentText();
      var data = JSON.parse(html).query.results.quote;
      for (var k in data) {
        var d = data[k];
        if (d) {
          var stockId = d['symbol'].replace(/\.TW/, '');
          if (stockId) {
            var s = JSON.parse(stocks[stockId]);
            s['z'] = d['LastTradePriceOnly'];
            s['change'] = d['Change'];
            stocks[stockId] = JSON.stringify(s);
          }
        }
      }      
      maps = [];
    }
    i++;
  }
  return stocks;
}

// 上市公司 個股日本益比、殖利率及股價淨值比
function mergeRoe1(stocks) {
  var html = '';
  try {
    html = UrlFetchApp.fetch(AFTER_MARKET_1_URL).getContentText('big5');
  } catch (error) {
    ss.toast('mergeRoe(): ' + error);
    return {};
  }
  var pattern = new RegExp("<input.*?id='html'.*?\/>", 'i');
  var t = pattern.exec(html);
  // Strip non well-formed tags for XMLService
  html = /value="([^"]*)"/.exec(t[0]);
  html = html[1].replace(/<\s*(\w+).*?>/ig, '<$1>');
  var doc = XmlService.parse(html);
  var html = doc.getRootElement();
  var matches = {};
  var atom = XmlService.getNamespace('http://www.w3.org/2005/Atom'); 
  var trElements = getElementsByTagName(html, 'tr');

  for (var i=0; i<trElements.length; i++) {
    var tr = trElements[i];
    var labels = [];
    var tdElements = getElementsByTagName(tr, 'td');
    if (tdElements.length >= 4) {
      var stockId = tdElements[0].getText().trim();
      if (!( /^\d+$/.test(stockId))) {
        continue;
      }
      
      if (stocks[stockId]) {
        var s = JSON.parse(stocks[stockId]);
        s['priceEarningsRatio'] = tdElements[2].getText().trim(), // 本益比
        s['cashDividendYield'] = tdElements[3].getText().trim(), // 殖利率
        s['priceBookRatio'] = tdElements[4].getText().trim(), // 股價淨值比
        matches[stockId] = s;
        stocks[stockId] = JSON.stringify(s);
      }
    }
  }
  return matches;  
}

function getStockMarket1() {
  return getStockProfile(MARKET_LIST_1);
}

function getStockMarket2() {
  
  var html = UrlFetchApp.fetch(CURRENT_MARKET_2).getContentText();
  var data = JSON.parse(html);
  
  var matches = {};
  for (var i=0; i<data.aaData.length; i++) {
    var stock = data.aaData[i];
    var match = {
        stockId: stock[0],
        stockName: stock[1],
        change: stock[3],
        z: stock[7],
    }
    matches[match.stockId] = JSON.stringify(match);
  }
  
  return matches;
}

// 上市、上櫃公司基本資料 代碼、名稱
function getStockProfile(url) {
  var html = UrlFetchApp.fetch(url).getContentText('big5');
  var pattern = new RegExp("<table class='h4'.*?>((.|[\n\r])*)<\/table>", 'ig');
  var t = pattern.exec(html);

  // Strip non well-formed tags for XMLService
  html = t[0].replace(/<\s*(\w+).*?>/ig, '<$1>')

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
      var stockId = tdElements[2].getText().trim();
      if (!( /^\d+$/.test(stockId))) {
        continue;
      }
      var dividendPerShare = tdElements[3].getText().trim();
      var cashDividendYield = tdElements[4].getText().trim();
      var priceBookRatio = tdElements[5].getText().trim();
      var stockName = tdElements[3].getText().trim();
      var stock = new Stock(stockId, stockName);
      //var dividents = getDividend(stock);
      
      var match = {
        stockId: stockId,
        stockName: stockName,
      };
      matches[stockId] = JSON.stringify(match);
    }
  }
  return matches;
}

// 上櫃公司基本資料
function getAfterMarketTrading() {
  var html = UrlFetchApp.fetch(AFTER_TRADNING_URL).getContentText('UTF-8');

  // Strip non well-formed tags for XMLService
  html = html
    .replace(/<link[^>]*>/ig, '')
    .replace(/<meta[^>]*>/ig, '')
    .replace(/&nbsp;/ig, '')
    .replace(/<thead.*?[^>].*?<\/thead>/ig, '')
    .replace(/<tfoot.*?[^>].*?<\/tfoot>/ig, '')
  
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
      matches[stockId] = {
        stockId: tdElements[0].getText(),
        priceEarningsRatio: tdElements[2].getText(),
        dividendPerShare: tdElements[3].getText(),
        cashDividendYield: tdElements[4].getText(),
        priceBookRatio: tdElements[5].getText(),  // 股價淨值比
      };
    }
  }
  
  return matches;
}

function mergeCurrentMarket2(data) {
  var stocks = getAfterMarketTrading();
  
  for (var key in stocks) {
    var s = stocks[key];
    if (data[key]) {
      var d = JSON.parse(data[key]);
      if (d) {
        d['cashDividendYield'] = s.cashDividendYield;
        d['priceBookRatio'] = s.priceBookRatio;
        d['priceEarningsRatio'] = s.priceEarningsRatio;
      }
      data[key] = JSON.stringify(d);
    }
  }
  return data;
}

// 近 9 年股息、股利
function mergeDividends(stocks) {
  var html = UrlFetchApp.fetch(DIVIDEND_HISTORY_URL).getContentText();  
  var start = /<table class=\"display\" id=\"example\">/.exec(html);
  var end = /<\/table>/.exec(html);
  
  var html = html.substring(start.index, end.index + '</table>'.length);  
  html = html.replace(/<\s*(\w+).*?>/ig, '<$1>')
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
      var stockId = tdElements[0].getText().trim();
      if (!( /^\d+$/.test(stockId))) {
        continue;
      }      
      if (stocks[stockId]) {
        var match = JSON.parse(stocks[stockId]);
        if (match) {
          var dividends = {};
          dividends['2015'] = Number(tdElements[3].getText().trim()) + Number(tdElements[4].getText().trim());
          dividends['2014'] = Number(tdElements[5].getText().trim()) + Number(tdElements[6].getText().trim());
          dividends['2013'] = Number(tdElements[7].getText().trim()) + Number(tdElements[8].getText().trim());
          dividends['2012'] = Number(tdElements[9].getText().trim()) + Number(tdElements[10].getText().trim());
          dividends['2011'] = Number(tdElements[11].getText().trim()) + Number(tdElements[12].getText().trim());
          dividends['2010'] = Number(tdElements[13].getText().trim()) + Number(tdElements[14].getText().trim());
          dividends['2009'] = Number(tdElements[15].getText().trim()) + Number(tdElements[16].getText().trim());
          dividends['2008'] = Number(tdElements[17].getText().trim()) + Number(tdElements[18].getText().trim());
          dividends['2007'] = Number(tdElements[19].getText().trim()) + Number(tdElements[20].getText().trim()); 
          match['dividends'] = dividends;    
          stocks[stockId] = JSON.stringify(match);
          //matches[stockId]['dividends'] = match;
        }
      }
    }
  }
  return matches;
}

function demoJob() {
  var scriptProperties = PropertiesService.getScriptProperties();
//  scriptProperties.deleteAllProperties();
  // 暫存上市公司資訊
//  var data = getStockMarket1();
//  mergeCurrentPrice(data);
//  mergeRoe1(data);
//  mergeDividends(data);
//  scriptProperties.setProperties(data);
//  Logger.log(JSON.parse(scriptProperties.getProperty('8114')));

  var userProperties = PropertiesService.getUserProperties();
  // 暫存上櫃公司資訊  
  var data2 = getStockMarket2();
  mergeCurrentMarket2(data2);
  mergeDividends(data2);
  userProperties.setProperties(data2);
  
  renderStocks();
}

function testJob() {
  var scriptProperties = PropertiesService.getScriptProperties();

  Logger.log(JSON.parse(getProperty('8114')));
  Logger.log(JSON.parse(getProperty('3702')));  
}



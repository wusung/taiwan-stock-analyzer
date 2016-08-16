var CURRENT_MARKET_2 = 'http://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw';
var MARKET_LIST_1 = 'http://isin.twse.com.tw/isin/class_main.jsp?market=1&issuetype=1';
var MARKET_LIST_2 = 'http://isin.twse.com.tw/isin/class_main.jsp?market=2&issuetype=4';
var bot = 'fund.bot.com.tw';
var yuanta = 'jdata.yuanta.com.tw';
var mainUrl = bot;
var yuantaInfo = {
  'revenue': { 'name': '營收', 'url': mainUrl+"/z/zc/zch/zch_"},  
  'profile': { 'name': '個股速覽', 'url': mainUrl+"/z/zc/zcx/zcx_"},
  'eps': { 'name': '股利', 'url': mainUrl+"/z/zc/zcc/zcc_"}, 
  'dividend': { 'name': '股利', 'url': mainUrl+"/z/zc/zcc/zcc_"}, 
  'news': { 'name': '重大行事曆', 'url': mainUrl+"/z/zc/zci/zci_"}, 
  '4': "http://isin.twse.com.tw/isin/class_main.jsp?market=1&issuetype=1&Page=1&chklike=Yhklike=Y"
};

function Stock(stockId, stockName) {
  this.dividendUrlPrefix = yuantaInfo.dividend.url;
  this.revenueUrlPrefix = yuantaInfo.revenue.url;
  this.profileUrlPrefix = yuantaInfo.profile.url;
  this.newsUrlPrefix = yuantaInfo.news.url;
  this.stockId = stockId;
  this.stockName = stockName;
  this.ext = '.djhtm';
  
  this.getDividendUrl = function() {
    return this.dividendUrlPrefix + this.stockId + this.ext;
  }
  this.getRevenueUrl = function() {
    return this.revenueUrlPrefix + this.stockId + this.ext;
  }
  this.getProfileUrl = function() {
    return this.profileUrlPrefix + this.stockId + this.ext;
  }
  this.getNewsUrl = function() {
    return this.newsUrlPrefix + this.stockId + this.ext;
  }
  
  this.getDividentLink = function() {
    return "=HYPERLINK(\"" + this.getDividendUrl() + "\",\"" + "股利" + "\")";
  }
   this.getRevenueLink = function() {
    return "=HYPERLINK(\"" + this.getRevenueUrl() + "\",\"" + '營收' + "\")";
  }
  this.getProfileLink = function() {
    return "=HYPERLINK(\"" + this.getProfileUrl() + "\",\"" + '個股速覽' + "\")";
  }
  this.getNewsLink = function() {
    return "=HYPERLINK(\"" + this.getNewsUrl() + "\",\"" + '重大行事曆' + "\")";
  }

  this.setStockId = function(stockId) {
    this.stockId = stockId; 
  }
  this.setStockName = function(stockName) {
    this.stockName = stockName; 
  }
}

function getDividend(stock) {
  var html = UrlFetchApp.fetch(stock.getDividendUrl()).getContentText('big5');   
  var html = findElementByClassName(html, 't01');

  // Strip non well-formed tags for XMLService
  html = html
             .replace(/<\s*(\w+).*?>/ig, '<$1>')
             .replace(/<br>/ig, '')

  var matches = {};
  var trElements = html.split(/<\/tr>/);
  var regex = /<td>(.*?)<\/td>/i;
  for (var i=0; i<trElements.length; i++) {
    var tr = trElements[i].trim();
    if (!tr.match(/^<td>.+?/i))
      continue;
  
    var tdElements = findElementsByTagName(tr, 'td');
    if (!tdElements)
      continue;
    if (tdElements.length >= 7) {
      var m = regex.exec(tdElements[0], 'i');
      if (!m) 
        continue;
      var year = m[1];
      if (!(/^\d+$/.test(year))) {
        continue;
      }
      var reg = regex;
      var dividend = /<td>(.*?)<\/td>/.exec(tdElements[7])[1];
      var match = {
        year: year,
        dividend: dividend,
      };
      matches[year] = match;      
    }
  }
  return matches;  
}

function demoStock() {
  var stock = new Stock('3702', '中華電');
  var r = getDividend(stock);
  Logger.log(r);
}

//var scriptProperties = PropertiesService.getScriptProperties();
//var userProperties = PropertiesService.getUserProperties();
//var documentProperties = PropertiesService.getDocumentProperties();  


function getElementsByTagName(element, tagName) {  
  var data = [];
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();     
    if( elt !=null && elt.getName()== tagName) data.push(elt);      
  }
  return data;
}

function getElementsByClassName(element, classToFind) {  
  var data = [];
  var descendants = element.getDescendants();
  descendants.push(element);  
  for(i in descendants) {
    var elt = descendants[i].asElement();
    if(elt != null) {
      var classes = elt.getAttribute('class');
      if(classes != null) {
        classes = classes.getValue();
        if(classes == classToFind) data.push(elt);
        else {
          classes = classes.split(' ');
          for(j in classes) {
            if(classes[j] == classToFind) {
              data.push(elt);
              break;
            }
          }
        }
      }
    }
  }
  return data;
}

function findElementByClassName(html, id) {
  var pattern = new RegExp("<table class=\"" + id + "\".*?>((.|[\n\r])+?)<\/table>", 'ig');  
  var matches = html.match(pattern);
  return matches[0];
}

function bindStockValues(stock) {
  for (var i=0; i<stockInfo.length; i++) {
    
  }
}

Object.size = function(obj) {
    var size = 0, key;
    for (key in obj) {
        if (obj.hasOwnProperty(key)) size++;
    }
    return size;  
};

function findElementsBetweenEndTagName(html, tagName) {

  var regex = /<\/tr>/;
//  var regex = new RegExp("^</tr>((.|[\n\r])+?)", 'ig');  
  var matches = [];
  var x = html.split(/<\/tr>/);
  Logger.log('x');
  Logger.log(x);
  while (match = regex.exec(html)) {
    // full match is in match[0], whereas captured groups are in ...[1], ...[2], etc.
    matches.push(match[0]);
  }
  
  Logger.log(matches);

//  var regex = new RegExp("</tr>((.|[\n\r])+?)</tr>", 'igm');  
////  var regex = /<\/tr>((.|[\n\r])+?)<\/tr>/ig;
////  var regex = /<\/tr>(.*?)<\/tr>/ig;
//  var matches = html.match(regex);
  Logger.log(matches);
  return matches;
}

function findElementsByTagName(html, tagName) {
  var pattern = new RegExp("<td>((.|[\n\r])+?)<\/td>", 'igm');  
  var matches = html.match(pattern);
  return matches;
}

function getProperty(key) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var userProperties = PropertiesService.getUserProperties();

  var value = scriptProperties.getProperty(key);
  if (value) {
    return value;
  }

  value = userProperties.getProperty(key);
  if (value) {
    return value;
  }
  
  return null;
}

function calculateAverageDividend(dividends) {
  if (!dividends) {
    return;
  }
  var arrs = [];
  for (var key in dividends) {
    arrs.push(key);
  }
  
  Array.sort(arrs);
  var sum = 0;  
  for (var i=arrs.length-1; i>arrs.length-6; i--) {
    Logger.log(dividends[arrs[i]]);
    sum += Number(dividends[arrs[i]]);  
  }
  return Number(sum / 5).toFixed(3);
}

function demoUtils() {
  var x = JSON.parse(getProperty('4904'))
  var y = calculateAverageDividend(x.dividends);
  Logger.log('avg=' + y);
}



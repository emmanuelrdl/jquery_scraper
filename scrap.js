var fs = require('fs');
var jsdom = require("node-jsdom");
var excelbuilder = require('msexcel-builder');

getPageData(1);

var result = [];
function getPageData(i) {
  jsdom.env({
    url: 'https://xxx.xxxx.com/ventes_immobilieres/offres/ile_de_france/hauts_de_seine/?o=' + i.toString(),
    scripts: ["http://code.jquery.com/jquery.js"],
    done: function (errors, window) {
      var $ = window.$;
      $('li a section').each(function(index, section) {
        var title  = $(section).find(".item_title").text().replace(/\n/g, '').trim();
        var price  = $(section).find(".item_price").text().replace(/\s|\n|\€/g, '');
        var city   = $(section).find("p[itemscope]").text().replace(/\s|\n/g, '').trim();
        var slashIndex = ""
        if (city.length > 1) {
          slashIndex = city.indexOf("/")
          city       = city.substring(0, slashIndex)
        }
        var item = {title: title, price: price, city: city}
        result.push(item);
      })
       if (i <= 3) {
        getPageData(i + 1)
      } else {
        writeFile(result)
      }
     
    }
  });
}

function writeFile(data) {
  var workbook = excelbuilder.createWorkbook('./', 'output.xlsx')
  var sheet = workbook.createSheet('sheet', 3, data.length);
  sheet.set(1, 1, 'title');
  sheet.set(2, 1, 'price');
  sheet.set(3, 1, 'city');
  var n = 2
  data.forEach(function(value){
   if (n <= data.length) {
    sheet.set(1, n , value.title);
    sheet.set(2, n , value.price);
    sheet.set(3, n , value.city);
    n++
   };
  });
  workbook.save(function(err){
    if (err)
      throw err;
    else
      console.log('file is saved');
  });

}


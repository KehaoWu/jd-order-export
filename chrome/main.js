var data = []
var colnames = []

function Workbook() {
  if(!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

function datenum(v, date1904) {
  if(date1904) v+=1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  for(var R = 0; R != data.length; ++R) {
    for(var C = 0; C != data[R].length; ++C) {
      if(range.s.r > R) range.s.r = R;
      if(range.s.c > C) range.s.c = C;
      if(range.e.r < R) range.e.r = R;
      if(range.e.c < C) range.e.c = C;
      var cell = {v: data[R][C] };
      if(cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

      if(typeof cell.v === 'number') cell.t = 'n';
      else if(typeof cell.v === 'boolean') cell.t = 'b';
      else if(cell.v instanceof Date) {
        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      }
      else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }
  if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

var filename = function(){
  var date = new Date()
  var year = date.getYear() + 1900
  var month = date.getMonth()
  var day = date.getDay()
  var hour = date.getHours()
  var minute = date.getMinutes()
  var second = date.getSeconds()
  var title = $(".time-txt").html()
  var title = ""
  var filename = "jd.com.orders." + title + "." + year + month + day + hour + minute + second + ".xlsx"
  return filename
}

var save = function(dataxls){
    var wb = new Workbook()
    var title = $(".time-txt").html()
    var title = "a"
    wb.SheetNames.push(title)
    wb.Sheets[title] = sheet_from_array_of_arrays(dataxls)

    var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'})
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), filename())
}

var checkFirstPage = function(){
    if($(".prev-disabled").html() != undefined){
        window.localStorage.clear('data');
        bindButton();
        return true
    }else{
        return false
    }
}

var checkLastPage = function(){
    if($(".next-disabled").html() != undefined){
        return true
    }else{
        return false
    }
}



var viewItem = function(urf){
    window.localStorage.setItem("lock", "true")
    window.open(nextUrl)
}

var _retrieve = function(){
    const tbodies = $(".td-void > tbody")
    var items = []
    var date, dealno, status, amount, cosignee, pname, pnumber
    tbodies.each(function(value, item){
        date = $(item).find('.dealtime').html()
        dealno = $(item).find('.number>a').html()
        status = $(item).find('.order-status').html()
        amount = $(item).find('.amount > span').html()
        cosignee = $(item).find('.txt').html()
        pname = $(item).find('.p-name > a').html()
        url = $(item).find('.p-name > a').attr("href")
        pnumber = $(item).find('.goods-number').html()
        if(pname != undefined){
            /*
            setTimeout(function(){
                window.open(url)
            }, 1000)
            */
            items.push({
                date: date,
                dealno: dealno,
                status: status.replace(/\s*/,''),
                amount: amount.replace('总额 ¥', ''),
                cosignee: cosignee,
                pname: pname,
                pnumber: pnumber.replace('x', '').replace(/\s*/,''),
                url: url,
                item: [date, pname, pnumber.replace('x', '').replace(/\s*/,''),
                    dealno, cosignee, amount.replace('总额 ¥', ''), status.replace(/\s*/,''), url]
            })
        }
        pname = undefined
    })
    var data = loadData();
    data[getCurrentPage()] = items
    window.localStorage.setItem("data", JSON.stringify(data))
}

var loadData = function(){
    const data = window.localStorage.getItem("data");
    if (data == undefined){
        return {}
    }else{
        return JSON.parse(data)
    }
}

var fetchItem = function(url){
    var items = JSON.parse(window.localStorage.getItem('items'));
    if(items[url] == undefined){
        return undefined
    }else{
        const category = items[url]['category'];
        const visit = items[url]['visit'];
        if(visit >= 1){
            return category
        }else{
            window.open(url)
            return undefined
        }
    }
}



var patch = function(){
      var items = JSON.parse(window.localStorage.getItem('data'));
      const keys = Object.keys(items)
      for(var i = 0; i < keys.length; i++){
          for(var j = 0; j < items[keys[i]].length; j++){
              const url = items[keys[i]][j]['url']
              $.ajax({
                url: url,
                type: 'get',
                success: function(data){
                    var _items = JSON.parse(window.localStorage.getItem('data'));
                    const keys = Object.keys(_items);
                    console.log(keys)
                    console.log("i:" + i + ", j:" + j)
                    const category = $(data).find(".breadcrumb > strong > a").html()
                    console.log(category)
                    console.log(_items[keys[i]])
                    if(_items[keys[i]][j]['item'].length < 8){
                        _items[keys[i]][j]['item'].push(category)
                        window.localStorage.setItem('data', JSON.stringify(_items))
                    }
                },
                async: false
              })

          }
      }
}

var saveData = function(){
    //patch()
    const data = window.localStorage.getItem('data');
    if (data != undefined){
        const xlsxdata = JSON.parse(data);
        const keys = Object.keys(xlsxdata);
        if(validator(keys)){
            var items = [];
            for(var i = 0; i < keys.length; i++){
                var _items = [];
                for(var j = 0; j < xlsxdata[keys[i]].length; j++){
                    _items.push(xlsxdata[keys[i]][j]['item'])
                }
                items = items.concat(_items)
            }
            console.log(items)
            save(items)
        }
    }
}

var validator = function(keys){
    let _keys = [];
    for(var i = 0; i < keys.length; i++){
        _keys.push(i + 1)
    }
    return _keys.toString() == keys.toString()
}

var getCurrentPage = function(){
    return $(".current").html()
}

var findNextPage = function(){
    return $('.current').next().attr("href")
}

var bindButton = function(){
    const btn = $("<button>").html("导出报告").css("margin-right", "16px").click(retrieve);
    $(".order-detail-txt").before(btn);
}

var retrieve = function(){
    const lastPage = checkLastPage();
    console.log("This is last page or not? " + lastPage)
    _retrieve();
    if (lastPage == false){
        const nextUrl = findNextPage()
        console.log("Next page is " + nextUrl)
        setTimeout(function(){
            window.location.replace(nextUrl)
        }, 1000)

    } else {
        saveData();
    }
}

$(document).ready(function(){
    /*
        翻页，从 prev-disabled 一直翻到 next-disabled
    */

    const firstPage = checkFirstPage();
    console.log("This is first page or not? " + firstPage)
    if(!firstPage){
        retrieve()
    }
})

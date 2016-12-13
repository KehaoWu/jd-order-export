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
  var filename = "jd.com.orders." + year + month + day + hour + minute + second + ".xlsx"
  return(filename)
}

var save = function(dataxls){
    var wb = new Workbook()
    wb.SheetNames.push("订单")
    wb.Sheets["订单"] = sheet_from_array_of_arrays(dataxls)
    wb.SheetNames.push("Raw")

    var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'})
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), filename())
}

var checkFirstPage = function(){
    if($(".prev-disabled").html() != undefined){
        window.localStorage.clear('data');
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

var retrieve = function(){
    const tbodies = $(".td-void > tbody")
    var items = []
    var dataxls = []
    var date, dealno, status, amount, cosignee, pname, pnumber
    tbodies.each(function(value, item){
        date = $(item).find('.dealtime').html()
        dealno = $(item).find('.number>a').html()
        status = $(item).find('.order-status').html()
        amount = $(item).find('.amount > span').html()
        cosignee = $(item).find('.txt').html()
        pname = $(item).find('.p-name > a').html()
        console.log(pname)
        pnumber = $(item).find('.goods-number').html()
        if(pname != undefined){
            items.push({
                date: date,
                dealno: dealno,
                status: status.replace(/\s*/,''),
                amount: amount.replace('总额 ¥', ''),
                cosignee: cosignee,
                pname: pname,
                pnumber: pnumber.replace('x', '').replace(/\s*/,''),
                item: [date, pname, pnumber, dealno, cosignee, amount, status]
            })
            dataxls.push([date, pname, pnumber.replace('x', '').replace(/\s*/,''), 
                    dealno, cosignee, amount.replace('总额 ¥', ''), status.replace(/\s*/,'')])
        }
        pname = undefined
    })
    var data = loadData();
    data[getCurrentPage()] = dataxls
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

var saveData = function(){
    const data = window.localStorage.getItem('data');
    if (data != undefined){
        const xlsxdata = JSON.parse(data);
        const keys = Object.keys(xlsxdata);
        if(validator(keys)){
            var items = [];
            for(var i = 0; i < keys.length; i++){
                items = items.concat(xlsxdata[keys[i]])
            }
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

$(document).ready(function(){
    /*
        翻页，从 prev-disabled 一直翻到 next-disabled
    */
    const firstPage = checkFirstPage();
    const lastPage = checkLastPage();
    console.log("This is first page or not? " + firstPage)
    console.log("This is last page or not? " + lastPage)
    retrieve();
    if (lastPage == false){
        const nextUrl = findNextPage()
        console.log("Next page is " + nextUrl)
        window.open(nextUrl)
    } else {
        saveData();
    }
})
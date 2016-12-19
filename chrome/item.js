var getItems = function(){
    const data = window.localStorage.getItem("items");
    if(data == undefined){
      return {}
    } else {
      return JSON.parse(data)
    }
}

var setItems = function(items){
    window.localStorage.setItem("items", JSON.stringify(items))
}

$(document).ready(function(){
    var category = $(".breadcrumb > strong").html();
    var items = getItems()
    items[window.location.pathname] = {"category": category, "visit": 1}
    setItems(items)
    //window.close()
})
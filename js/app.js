var app = {}
app.ajaxRequest = function(url) {
    $.ajax({
        dataType: "script",
        url : '/js/'+url+'.js',
        cache: false,
        complete : function(response){
            if(response.status==200){
                //console.log(response)          
            }else{
                //console.log(response)
            }
        }
    })
}

app.ajaxRequest('parser');

$(document).on('click','#reload',function(){
    console.clear();
    $('#file').val('');
    $('#result').html('Ready');
    app.ajaxRequest('parser');    
})
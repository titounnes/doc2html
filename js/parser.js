var result = $('#result');
var docParser = {};
docParser.read = false;
docParser.styles = {};
$('#file').on('change', function(evt){
    if(evt.target.files.length==0){
        return false;
    }
    $.each(sessionStorage, function(id){
        delete sessionStorage[id];
    })
    docParser.handleFile(evt.target.files[0])
})

docParser.handleFile = function(f){
    var title = $('<h4>', {
        text: f.name ? f.name : '',
    })
    var fileContent = $('<div>');
    
    result.html(title);

    result.append(fileContent)
    docParser.read = true;
    docParser.component(f);    
}

docParser.component = function(f){
    JSZip.loadAsync(f)
    .then(function(part){
        part.file('word/styles.xml').async('string')
        .then(function(text){
            var parser = new DOMParser();
            docParser.style(parser.parseFromString(text.replace(/w:/g,'').replace(/wp:/g,''), 'text/xml').childNodes[0].childNodes)
        })
        part.file('word/document.xml').async('string')
        .then(function(text){
            var parser = new DOMParser();
            docParser.body(parser.parseFromString(text.replace(/w:/g,'').replace(/wp:/g,''), 'text/xml').childNodes[0].childNodes[0].childNodes)
        })
        var media = part.file(/media/);
        if(media.length>0){
            media.forEach(function(image, i){
                part.file(image.name).async('blob')
                .then(function(text){
                    if(typeof sessionStorage['#img_'+(i+1)] == 'undefined'){
                        var reader = new FileReader();
                        reader.readAsDataURL(text)
                        reader.onload = function(){
                            sessionStorage['#img_'+(i+1)] = true;
                            $('#img_'+(i+1)).attr('src', this.result)
                        }
                    }
                })
            })
        }
    })
}

docParser.style = function(xml){
    xml.forEach(function(style){
        if(style.nodeName=='style'){
            if(typeof style.attributes[2] != 'undefined'){
                docParser.styles[style.attributes[2].nodeValue.replace('-','')] = style.childNodes[0].attributes[0].nodeValue.replace('-',''); 
            }else if(typeof style.attributes[1] != 'undefined'){
                docParser.styles[style.attributes[1].nodeValue.replace('-','')] = style.childNodes[0].attributes[0].nodeValue.replace('-','');
            }            
        }
        
    })
}

docParser.body = function(xml){
    xml.forEach(function(p, pIndex){
        if(p.nodeName=='p'){
            if(typeof $('#p_'+pIndex).html() == 'undefined'){
                $('#result').append($('<p>', {
                    id : 'p_'+pIndex,
                }))
            }
            p.childNodes.forEach(function(pc,pi){    
                if(pi==0){
                    //atributes paragraf
                    pc.childNodes.forEach(function(pAttr){
                        var test = docParser[pAttr.nodeName];
                        if(typeof test == 'function'){
                            test('#p_'+pIndex, pAttr.attributes);
                        }else{
                            //console.log(pAttr)
                        }
                    })                
                }else{
                    //console.log(pc.childNodes, pc.childNodes.length)
                    if(pc.childNodes.length>1){
                        pc.childNodes.forEach(function(pData, pDataIndex){
                            if(pDataIndex==0){
                                //atributes text
                                if(typeof $('#pr_'+pIndex+'_'+pi).html() == 'undefined'){
                                    $('#p_'+pIndex).append($('<span>',{
                                        id : 'pr_'+pIndex+'_'+pi
                                    }))  
                                }
                                var test = docParser[pData.nodeName];
                                if(typeof test == 'function'){
                                    if(pData.childNodes && pData.childNodes.length>0){
                                        test('#pr_'+pIndex+'_'+pi, pData.childNodes)
                                    }                            
                                }else{
                                    console.log(pData, pData.nodeName)
                                }    
                            }else{
                                //text
                                var test = docParser[pData.nodeName];
                                if(typeof test=='function'){
                                    test('#pr_'+pIndex+'_'+pi, pData.childNodes)
                                }else{
                                    console.log(pData, pData.nodeName)
                                }
                            }
                        })
                    }else{
                        if(pc.childNodes[0] && pc.childNodes[0].childNodes[0]){
                            var test = docParser[pc.childNodes[0].childNodes[0].nodeName.replace('#','')];
                            if(typeof test=='function'){
                                if(typeof $('#pr_'+pIndex+'_'+pi).html() == 'undefined'){
                                    $('#p_'+pIndex).append($('<span>',{
                                        id : 'pr_'+pIndex+'_'+pi
                                    }))    
                                }
                                test('#pr_'+pIndex+'_'+pi, pc.childNodes[0].childNodes[0].nodeValue)
                            }else{
                                //console.log(test, pc.nodeName[0])
                            }
                        }
                    }
                }
            })
        }else if(p.nodeName=='tbl'){
            if(typeof $('#tbl_'+pIndex).html() == 'undefined'){
                $('#result').append($('<table>', {
                    id : 'tbl_'+pIndex,
                    'style': 'margin-left:10%;width:80%',
                }))
            }
            p.childNodes.forEach(function(tr,iTr){
                if(tr.nodeName=='tr'){
                    if(typeof $('#tr_'+pIndex+'_'+iTr).html() == 'undefined'){
                        $('#tbl_'+pIndex).append($('<tr>',{
                            id: 'tr_'+pIndex+'_'+iTr,
                        }))
                    }
                    tr.childNodes.forEach(function(tc, iTc){
                        if(tc.nodeName == 'tc'){
                            if(typeof $('#tc_'+pIndex+'_'+iTr+'_'+iTc).html() == 'undefined'){
                                $('#tr_'+pIndex+'_'+iTr).append($('<td>',{
                                    id: 'tc_'+pIndex+'_'+iTr+'_'+iTc,
                                }))
                            }
                            if(typeof tc.childNodes[1].childNodes[1] == 'undefined'){
                                docParser.tc('#tc_'+pIndex+'_'+iTr+'_'+iTc, ' ', tc.childNodes[0].childNodes[1], '#tr_'+pIndex+'_'+iTr)
                                //console.log('#tr_'+pIndex+'_'+iTr,tc.childNodes[0].childNodes[1])
                            }else{
                                docParser.tc('#tc_'+pIndex+'_'+iTr+'_'+iTc, tc.childNodes[1].childNodes[1].childNodes, tc.childNodes[0].childNodes[1], '#tr_'+pIndex+'_'+iTr)
                            }                            
                        }
                    })                    
                }
            })
            
        }
    })
}

docParser.pStyle = function(target, attr){
    if(isNaN(attr[0].nodeValue)){
        $(target).addClass(attr[0].nodeValue);
    }else{
        $(target).addClass(docParser.styles[attr[0].nodeValue]);
    }
    
}

docParser.rPr = function(target, attr){
    var st = {};
    if(attr.length>1){
        attr.forEach(function(rPr,i){
            var test = docParser[rPr.nodeName];
            if(typeof test =='function'){
                test(target, rPr);
            }else{
                console.log(rPr.nodeName)
            }
        })
    }else{
        if(attr[0]){
            var test = docParser[attr[0].nodeName];
            if(typeof test =='function'){
                test(target, attr[0]);
            }else{
                console.log(attr[0].nodeName)
            }
        }
    }
    $(target).css(st);
}

docParser.vertAlign = function(target, val){
    $(target).css({
        'vertical-align' : val.attributes[0].nodeValue.replace('script',''),
        'font-size' : '60%',
    })
}

docParser.b = function(target){
    $(target).css({
        'font-weight' :'bold',
    })
}

docParser.bCs = function(target){
    docParser.b(target)
}

docParser.i = function(target){
    $(target).css({
        'font-style' : 'italic',
    })
}

docParser.iCs = function(target){
    docParser.i(target);
}

docParser.u = function(target){
    $(target).css({
        'text-decoration' : 'underline',
    })
}

docParser.rFonts = function(target, val){
    if(val.attributes && val.attributes.length >1){
        $(target).css({
            'font-family' : val.attributes[val.attributes.length-1].nodeValue,
        })
    }
}

docParser.sz = function(target, val){
    $(target).css({
        'font-size' : val.attributes[0].nodeValue,
    })
}

docParser.szCs = function(target, val){
    docParser.sz(target,val)
}

docParser.color = function(target, val){
    $(target).css({
        'color' : '#'+val.attributes[0].nodeValue,
    })
}

docParser.caps = function(target, val){
    if(val.attributes[0]){
        $(target).css({
            //'color' : '#'+val.attributes[0].nodeValue,
        })    
    }
}

docParser.smallCaps = function(target, val){
    if(val.attributes[0]){
        $(target).css({
            //'color' : '#'+val.attributes[0].nodeValue,
        })    
    }
}

docParser.lang = function(target, val){
    /*if(val.attributes[0]){
        $(target).css({
            //'color' : '#'+val.attributes[0].nodeValue,
        })    
    }*/
}

docParser.kern = function(target, val){
    /*if(val.attributes[0]){
        $(target).css({
            //'color' : '#'+val.attributes[0].nodeValue,
        })    
    }*/
}

docParser.r = function(target, data){
    console.log(data)
    if(data.length<2){
        return false;
    }
    if(data[0]){
        //$(target).html(data[0])
    }
}

docParser.t = function(target, data){
    if(data[0]){
        $(target).html(data[0])
    }
}

docParser.text = function(target, data){
    if(data){
        $(target).html(data)
    }
}

docParser.drawing = function(target, data){
    var id, cx, cy = '';
    $.each(data[0].childNodes, function(i, node){
        if(node.nodeName=='docPr'){
            id = node.attributes[0].nodeValue
        }else if(node.nodeName=='extent'){
            cx = node.attributes[0].nodeValue;
            cy = node.attributes[1].nodeValue;
        }
        
    })
    if(typeof $('#img_'+id).html() == 'undefined'){
        $(target).append($('<img>',{
            id: 'img_'+id,
            style: 'width:'+cx/(1024*10)+'px;height:'+cy/(1024*10)+'px;align:center',
        }))
    }
}

docParser.tc = function(target, data, border, parent){
    if(data == ' '){
        $(target).html('&nbsp;')
        border.childNodes.forEach(function(tpr){
            switch(tpr.nodeName){
                case 'top' : 
                    $(parent).addClass('borderTop');
                    break;
                case 'bottom':
                    $(parent).addClass('borderBottom');
                    break;
            }
            
        })
    }else{
        $(target).html(data[1].childNodes[0].nodeValue)
        border.childNodes.forEach(function(tpr){
            switch(tpr.nodeName){
                case 'top' : 
                    $(parent).addClass('borderTop');
                    break;
                case 'bottom':
                    $(parent).addClass('borderBottom');
                    break;
            }
            
        })

    }    
}
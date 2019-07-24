<script src="/js/jquery.min.js"></script>
<script src="/js/jszip.min.js"></script>
<h3>Choose File</h3>
<input type="file" id="file" name="file" multiple /><br>
<div id="result_block" class="hidden">
    <h3>Content</h3>
    <div id="result"></div>
</div>
<script>
    var $result = $('#result');
    $('#file').on('change', function(evt){
        $result.html('');
        $('#result_block').removeClass('hidden').addClass('show');
        
        function handleFile(f){
            var $title = $('<h4>', {
                text: f.name
            })
            var $fileContent = $('<div>');
            
            $result.append($title);

            $result.append($fileContent)

            var dateBefore = new Date();

            JSZip.loadAsync(f)
            .then(function(zip){
                var dateAfter = new Date();
                $title.append($('<span>', {
                    'class' : 'small',
                    text: '(loaded word/styles.xml in '+ (dateAfter - dateBefore) +' ms)'
                }))
                return zip.file('word/styles.xml').async('string')
            }).then(function(text){
                $('head').append($('<style>'))
                text = text.replace(/w:/g,'');
                    
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(text, 'text/xml');
                var style = xmlDoc.childNodes[0].childNodes;
                var st = '';
                                
                style.forEach(function(s, sIndex){
                    if(s.childNodes.length>0){
                        if(s.nodeName=='docDefaults'){
                            if(s.childNodes.length>0){
                                s.childNodes.forEach(function(dd, ddIndex){
                                    if(dd.childNodes[0].childNodes.length>0){
                                        var ddStyle = dd.childNodes[0].childNodes[0];
                                        switch(ddStyle.nodeName){
                                            case 'rFonts' :
                                                var font = ddStyle.attributes.cs.nodeValue;//.split(' ')[0];
                                                $('style').append('p{font-family:'+font+'}')
                                                break;
                                            default :
                                            console.log(ddStyle.nodeName)
                                                break;
                                        }
                                    }
                                })
                            }

                        }else if(s.nodeName=='style'){
                            if(s.childNodes.length>0){
                                s.childNodes.forEach(function(dd, ddIndex){
                                    switch(dd.nodeName){
                                        case 'name' :
                                            st += '.'+dd.attributes[0].nodeValue; 
                                            break;
                                        case 'pPr' : 
                                            dd.childNodes.forEach(function(atr,atrIndex){
                                                switch(atr.nodeName){
                                                    case 'jc' :
                                                        st +='text-align:'+atr.attributes[0].nodeValue+'}';
                                                        break;
                                                    default:
                                                        console.log(atr.nodeName);
                                                        break;
                                                }
                                                
                                            })
                                            if(typeof dd.childNodes[2]!='undefined'){
                                                //console.log(dd.childNodes[2]);
                                            }
                                            break;
                                        default : //console.log(dd.nodeName);
                                            break; 
                                    }


                                    
                                })
                            }
                        }else{
                            //console.log(s.nodeName)
                        }
                        //s.forEach(function(ss, ssIndex){
                            //console.log(ss.childNodes, ssIndex)
                        //})
                    }
                    console.log(st)                    
                })
                return 1;
                $.each($.parseHTML(text), function(i, style){
                    console.log(style.nodeName)
                    if(style.nodeName=='STYLES'){
                        //console.log(s)
                        $.each($.parseHTML(style.innerHTML), function(j, s){
                            
                            console.log(s)
                        })
                    }
                })
                
            })

            JSZip.loadAsync(f)
            .then(function(zip){
                var dateAfter = new Date();
                $title.append($('<span>', {
                    'class' : 'small',
                    text: '(loaded word/document.xml in '+ (dateAfter - dateBefore) +' ms)'
                }))
                return zip.file('word/document.xml').async('string')                    
            }).then(function(text){
                text = text.replace(/w:/g,'');
                var no = 0;

                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(text, 'text/xml');
                var paragraph = xmlDoc.childNodes[0].childNodes[0].childNodes;
                
                paragraph.forEach(function(p, pIndex){
                    $('#result').append($('<p>', {
                        id : 'par_'+pIndex,
                    }))
                    p.childNodes.forEach(function(r, rIndex){
                        if(rIndex==0){
                            r.childNodes.forEach(function(s, sIndex){
                                var attrLength = s.attributes.length;
                                if(attrLength>0){
                                    for(i=0;i<attrLength;i++){
                                        var aObj = s.attributes[i];
                                        var aId = '#par_'+pIndex;
                                        var val = aObj.nodeValue;
                                        var aName = aObj.nodeName; 
                                        var fh = 24;
                                        switch(aName){
                                            case 'val' : 
                                                $(aId).addClass(val);
                                                break;
                                            case 'before' :
                                                $(aId).css({'margin-top':(val)/fh+'px'})
                                                break;
                                            case 'after' :
                                                $(aId).css({'margin-bottom':(val)/fh+'px'})
                                                break;
                                            case 'left' :
                                                if(val > 0){
                                                    $(aId).css({'margin-left':(val)/fh+'px'})
                                                }
                                                break;
                                            case 'right' :
                                                if(val > 0){
                                                    $(aId).css({'margin-right':(val)/fh+'px'})
                                                }
                                                break;
                                            case 'hanging' :
                                                if(val > 0){
                                                    $(aId).css({'text-indent':'-'+$(aId).css('margin-left')})
                                                }
                                                break;
                                            default : console.log(aName, s);
                                                break;
                                        }
                                    }  
                                }
                            })                            
                        }else{
                            var rLength = r.childNodes.length;
                            var aId = '#par_'+pIndex;
                            var sId = 's_'+pIndex+'_'+rIndex;
                            $(aId).append($('<span>',{
                                id: sId,
                            }))
                            if(rLength > 1){
                                var rStyle = r.childNodes[0].childNodes 
                                if(rStyle.length > 0){
                                    rStyle.forEach(function(s, sIndex){
                                        var sTag = s.nodeName;
                                        switch(sTag){
                                            case 'i' :
                                            case 'iCs' :
                                                $('#'+sId).css({'font-style':'italic'});
                                                break;
                                            case 'b' :
                                            case 'bCs' :
                                                $('#'+sId).css({'font-weight':'bold'});
                                                break;
                                            case 'u' : 
                                                $('#'+sId).css({'text-decoration':'underline'});
                                                break;
                                            case 'vertAlign' :
                                                var va = {
                                                    subscript: 'sub',
                                                    superscript : 'super'
                                                }
                                                $('#'+sId).css({'vertical-align':va[s.attributes[0].nodeValue],'font-size':'60%'});
                                                break;
                                            case 'rFonts' :
                                                var sAttr = s.attributes.length;
                                                if(sAttr>0){
                                                    var font = s.attributes[sAttr-1].nodeValue.split(' ')[0];
                                                    $('#'+sId).css({'font-family':font});
                                                }
                                                break;
                                            case 'color' :
                                                var cAttr = s.attributes.length;
                                                if(cAttr>0){
                                                    $('#'+sId).css({'color':'#'+s.attributes[0].nodeValue});
                                                }
                                                break;
                                            case 'sz' :
                                            case 'szCs' :
                                                var szAttr = s.attributes.length;
                                                if(cAttr>0){
                                                    $('#'+sId).css({'font-size':s.attributes[0].nodeValue});
                                                }
                                                break;
                                            default : 
                                                //console.log(sTag, s);
                                                break;
                                        }
                                    })
                                }
                                for(i=1;i<rLength;i++){
                                    $('#'+sId).append(r.childNodes[i].childNodes[0].nodeValue);
                                }
                            }else if(rLength==1){
                                if(r.childNodes.length>1){
                                    console.log(r)
                                }
                                
                            }
                            
                        }
                    })
                })
                return 1;
                for(i=0; i < pLength; i++){
                    $('#result').append($('<p>', {
                        id : 'par_'+i,
                    }))


                    console.log(paragraph[i])
                }
                //console.log(body.childNodes.length)
                $.each(body, function(pKey, pValue){
                    //console.log(pValue.childNodes)
                })
                return 1;

                $.each($.parseHTML(text), function(i, doc){
                    if(doc.nodeName=='DOCUMENT'){
                        $.each($.parseHTML(doc.innerHTML), function(j, p){
                            if(p.nodeName=='P'){
                                var par = 'p_'+no;
                                $fileContent.append($('<p>', {
                                    id: par,
                                }))
                                $.each($.parseHTML(p.innerHTML), function(k, r){
                                    if(r.nodeName=='PPR'){
                                        $.each($.parseHTML(r.innerHTML), function(l, s){
                                            switch(s.nodeName){
                                                case 'PSTYLE' : 
                                                    $('#'+par).addClass($(s).attr('val'));
                                                    break;
                                                case 'SPACING' :
                                                    if($(s).attr('before')>0){
                                                        $('#'+par).css({'margin-top':$(s).attr('before')/24+'px'})
                                                    }
                                                    if($(s).attr('after')>0){
                                                        $('#'+par).css({'margin-bottom':$(s).attr('after')/24+'px'})
                                                    }
                                                    break;
                                                case 'RPR' :
                                                    $('#'+par).append(s.innerHTML)
                                                    break;
                                                case 'NUMPR' :
                                                    //Citation
                                                    //console.log(s.innerHTML);
                                                    break;
                                                case 'IND' :
                                                    if($(s).attr('left')*1>0 && $(s).attr('hanging')*1>0){
                                                        $('#'+par).css({'margin-left':$(s).attr('left')/(24)+'px','text-indent':$(s).attr('left')/(-24)+'px'})
                                                    }
                                                    break;
                                                case 'JC' :
                                                    $('#'+par).css({'text-align':'justify'})
                                                    break;
                                                case 'BIDI' :
                                                    //console.log(s);
                                                    break;
                                                case 'WIDOWCONTROL' :
                                                    //console.log(s);
                                                    break;
                                                default : console.log(s.nodeName);
                                                    break;
                                            }
                                        })
                                    }
                                    else if(r.nodeName=='R'){
                                        var noo = 0;
                                        $.each($.parseHTML(r.innerHTML), function(l, t){
                                            var row = 's_'+no;
                                            if(typeof $('#'+row).html()=='undefined'){
                                                $('#'+par).append($('<span>',{
                                                    id: row,
                                                }))
                                            }
                                            if(t.nodeName=='RPR'){
                                                $.each($(t.innerHTML), function(m, u){
                                                    switch(u.nodeName){
                                                        case 'I' :
                                                            $('#'+row).css({'font-style':'italic'});
                                                            break;
                                                        case 'B' :
                                                            $('#'+row).css({'font-weight':'bold'});
                                                            break;
                                                        case 'T' : 
                                                            $('#'+row).append(u.innerHTML);
                                                            noo++;
                                                            break;
                                                        case 'VERTALIGN' :
                                                            if($(u).attr('val')=='superscript'){
                                                                $('#'+row).css({'vertical-align':'super','font-size':'80%'});
                                                            }else if($(u).attr('val')=='subscript'){
                                                                $('#'+row).css({'vertical-align':'sub','font-size':'80%'});
                                                            }
                                                            break;
                                                        default :
                                                            //console.log(u.nodeName);
                                                            //console.log(u)
                                                            break;
                                                    }
                                                })
                                            }else if(t.nodeName=='T'){ 
                                                $('#'+row).append(t.innerHTML);    
                                                no++;  
                                                
                                            }
                                            
                                        })
                                        //$('#no_'+no).append(r.innerHTML)
                                        //console.log(r.innerHTML)
                                    }
                                })
                            }
                        })
                        
                    }
                    //console.log(i, v, v.nodeName)
                })
                /*parser = new DOMParser();
                xmlDoc = parser.parseFromString(text,"text/html");
                var body = xmlDoc.getElementsByTagName("body")[0];*/
                /*$.each($.parseHTML(body), function(j, v) {
                    console.log(v)
                })*/
                /*
                $.each(body, function(i, p){
                    //console.log(i, p)
                    if(i=='innerText'){
                        //console.log(p)
                    }
                    if(i=='innerHTML'){
                        console.log(p)
                    }
                    

                })*/
                /*$fileContent.append($('<div>',{
                    'class' : 'small',
                    text : text,
                }))*/
                //console.log(body)
            })

            JSZip.loadAsync(f)
            .then(function(zip){
                var dateAfter = new Date();
                $title.append($('<span>', {
                    'class' : 'small',
                    text: '(loaded word/styles.xml in '+ (dateAfter - dateBefore) +' ms)'
                }))
                return zip.file('word/styles.xml').async('string')
            }).then(function(text){
                //console.log(text)
            })

            
        }/*,function(evt){
            $result.append($('<div>', {
                'class' : 'alert alert-danger',
                text: 'Error Reading',
            }))
        }*/

        var files = evt.target.files;
        for(var i= 0; i < files.length; i++){
            handleFile(files[i])
        }
    })

</script>
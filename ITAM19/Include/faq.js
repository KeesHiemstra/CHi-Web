function createStyleElement(cssString){
  var newStyle=document.createElement("style");
  newStyle.setAttribute("type","text/css");
  if(newStyle.styleSheet){
    newStyle.styleSheet.cssText=cssString;
  }else{
    var cssText=document.createTextNode(cssString);
    newStyle.appendChild(cssText);
  }
  document.getElementsByTagName('head')[0].appendChild(newStyle);
}

function addEvent(obj,type,fn){
  if(obj.addEventListener)
    obj.addEventListener(type,fn,false);
  else if(obj.attachEvent){
    obj["e"+type+fn]=fn;
    obj[type+fn]=function(){
      obj["e"+type+fn](window.event);
    }
    obj.attachEvent("on"+type,obj[type+fn]);
  }
}

var displayFAQ=function(){
  var createFAQLinks=function(){
    //Get all dd elements in this page
    var ddTag=document.getElementsByTagName('dd');
    //Leave this function if there are no dd elements in this page
    if(ddTag.length==0)return;

    //Get all dt elements in this page
    var dtTag=document.getElementsByTagName('dt');
    //Leave this function if there are no dt elements in this page
    if(dtTag.length==0)return;

    //Make links for all dt elements in the page
    for(var i=0;i<dtTag.length;i++){
      var dtText=dtTag[i].firstChild.nodeValue;
      dtTag[i].removeChild(dtTag[i].firstChild);
      var aTag=document.createElement('a');
      aTag.setAttribute('href','#');
      var aTagText=document.createTextNode(dtText);
      aTag.appendChild(aTagText);
      dtTag[i].appendChild(aTag);
      dtTag[i].firstChild.onclick=function(){toggleEachAnswer(this);return false;}
    }
  }
  //This will create the Open all/Close all link
  var createToggleAll=function(){
    var toggleAllLink=document.createElement('a');
    var toggleText=document.createTextNode(i18n.TEXT_OPEN_ALL);
    toggleAllLink.appendChild(toggleText);
    toggleAllLink.href='#';
    toggleAllLink.id='toggleAll';
    toggleAllLink.onclick=function(){toggleAllAnswers();return false;}

    //Insert Open all/Close all link before the tag with ID "first-FAQ_header"
    if(document.getElementById('first-FAQ_header')){
      var hdr=document.getElementById('first-FAQ_header');
      hdr.parentNode.insertBefore(toggleAllLink,hdr);
    }
  }

  var toggleEachAnswer=function(elem){
    if(elem.parentNode.nextSibling.nodeType==3&&!/\S/.test(elem.parentNode.nextSibling.nodeValue)){
      var answer=elem.parentNode.nextSibling.nextSibling;
    }else{
      var answer=elem.parentNode.nextSibling;
    }
    answer.style.display=(answer.style.display=='block')?'none':'block';
  }

  var toggleAllAnswers=function(){
    ddTag=document.getElementsByTagName('dd');
    if(document.getElementById('toggleAll').firstChild.nodeValue==i18n.TEXT_OPEN_ALL){
      for(var i=0;i<ddTag.length;i++){ddTag[i].style.display='block';}
      document.getElementById('toggleAll').firstChild.nodeValue=i18n.TEXT_CLOSE_ALL;
    }else{
      for(var i=0;i<ddTag.length;i++){ddTag[i].style.display='none';}
      document.getElementById('toggleAll').firstChild.nodeValue=i18n.TEXT_OPEN_ALL;
    }
  }

  return{
    init:function(){
      if(!document.getElementById||!document.getElementsByTagName)return;
      createFAQLinks();
      createToggleAll();
    }
  }
}();

addEvent(window,"load",displayFAQ.init);
createStyleElement("dd { display: none; } dl dt { font-weight: normal; }");

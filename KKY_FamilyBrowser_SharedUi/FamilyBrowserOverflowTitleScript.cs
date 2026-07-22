internal static class FamilyBrowserOverflowTitleScript
{
	public static string Script()
	{
		return @"(function(w,d){
  if(w.KKYFBOverflowTitles&&w.KKYFBOverflowTitles.version)return;
  var generated='data-kkyfb-overflow-title';
  var generatedText='data-kkyfb-overflow-title-text';
  var installed=false;
  function trim(v){return String(v||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'');}
  function tagOf(el){return el&&el.tagName?String(el.tagName).toLowerCase():'';}
  function styleOf(el){try{return w.getComputedStyle?w.getComputedStyle(el,null):el.currentStyle;}catch(e){return null;}}
  function styleValue(style,camel,dashed){if(!style)return '';try{return String(style[camel]||(style.getPropertyValue&&style.getPropertyValue(dashed))||'').toLowerCase();}catch(e){return '';}}
  function textOf(el){
    if(!el)return '';
    var explicit=el.getAttribute('data-full-text')||el.getAttribute('data-kkyfb-full-text')||'';
    if(explicit)return trim(explicit);
    var tag=tagOf(el);
    if(tag==='input'||tag==='textarea')return trim(el.value||'');
    if(tag==='select'){
      try{var option=el.options&&el.selectedIndex>=0?el.options[el.selectedIndex]:null;return trim(option?(option.text||option.innerText||option.value):'');}catch(e){}
    }
    return trim(el.innerText||el.textContent||'');
  }
  function isExcluded(el){
    var tag=tagOf(el);
    if(!tag||tag==='html'||tag==='body'||tag==='head'||tag==='script'||tag==='style'||tag==='meta'||tag==='link'||tag==='img'||tag==='svg'||tag==='path'||tag==='br')return true;
    if(el.getAttribute&&el.getAttribute('data-kkyfb-no-overflow-title')==='true')return true;
    return false;
  }
  function hasAuthorTitle(el){
    var title=el&&el.getAttribute?el.getAttribute('title'):'';
    return !!title&&el.getAttribute(generated)!=='1';
  }
  function isClipped(el){
    try{
      if(!el||el.clientWidth<1||el.clientHeight<1)return false;
      return el.scrollWidth>el.clientWidth+1||el.scrollHeight>el.clientHeight+1;
    }catch(e){return false;}
  }
  function isTextOverflowCandidate(el){
    if(isExcluded(el))return false;
    var style=styleOf(el);
    var overflow=styleValue(style,'overflow','overflow');
    var overflowX=styleValue(style,'overflowX','overflow-x');
    var overflowY=styleValue(style,'overflowY','overflow-y');
    var textOverflow=styleValue(style,'textOverflow','text-overflow');
    var whiteSpace=styleValue(style,'whiteSpace','white-space');
    if(textOverflow.indexOf('ellipsis')>=0)return true;
    var hidden=overflow.indexOf('hidden')>=0||overflowX.indexOf('hidden')>=0||overflowY.indexOf('hidden')>=0||overflow.indexOf('clip')>=0||overflowX.indexOf('clip')>=0||overflowY.indexOf('clip')>=0;
    if(!hidden)return false;
    var tag=tagOf(el);
    return whiteSpace.indexOf('nowrap')>=0||tag==='td'||tag==='th'||tag==='a'||tag==='button'||tag==='label'||tag==='input'||tag==='select';
  }
  function clearGenerated(el){
    if(!el||!el.getAttribute||el.getAttribute(generated)!=='1')return;
    try{el.removeAttribute('title');el.removeAttribute(generated);el.removeAttribute(generatedText);}catch(e){}
  }
  function update(el){
    if(!el||!el.getAttribute||isExcluded(el))return 0;
    if(hasAuthorTitle(el))return 2;
    var clipped=isTextOverflowCandidate(el)&&isClipped(el);
    if(!clipped){clearGenerated(el);return 0;}
    var text=textOf(el);
    if(!text){clearGenerated(el);return 0;}
    el.setAttribute('title',text);
    el.setAttribute(generated,'1');
    el.setAttribute(generatedText,text);
    return 1;
  }
  function refresh(root,generatedOnly){
    root=root||d.body||d;
    var all=root.getElementsByTagName?root.getElementsByTagName('*'):[];
    if(root.nodeType===1&&(!generatedOnly||root.getAttribute(generated)==='1'))update(root);
    for(var i=0;i<all.length;i++){
      if(generatedOnly&&all[i].getAttribute(generated)!=='1')continue;
      update(all[i]);
    }
    return true;
  }
  function inspectEvent(e){
    e=e||w.event;
    var el=e?(e.target||e.srcElement):null;
    for(var depth=0;el&&el!==d.body&&depth<7;depth++,el=el.parentNode){
      var result=update(el);
      if(result===1||result===2)return;
    }
  }
  function onResize(){w.setTimeout(function(){refresh(d.body||d,true);},0);}
  function install(){
    if(installed)return;
    installed=true;
    if(d.addEventListener){d.addEventListener('mouseover',inspectEvent,true);d.addEventListener('focusin',inspectEvent,true);w.addEventListener('resize',onResize,false);}
    else if(d.attachEvent){d.attachEvent('onmouseover',inspectEvent);d.attachEvent('onfocusin',inspectEvent);w.attachEvent('onresize',onResize);}
  }
  w.KKYFBOverflowTitles={version:'20260715',install:install,refresh:function(root){return refresh(root||d.body||d,false);},refreshGenerated:function(){return refresh(d.body||d,true);},update:update};
  w.refreshOverflowTitles=function(root){return w.KKYFBOverflowTitles.refresh(root);};
  w.applyControlTitles=function(){install();return true;};
  install();
})(window,document);";
	}
}

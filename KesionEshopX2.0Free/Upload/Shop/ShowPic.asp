<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.commoncls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

dim ks:set ks=new publiccls
Dim ID:ID=KS.ChkClng(KS.S("ID"))
Dim titleStr,PhotoStr,imgs,n,curr
If ID=0 Then KS.Die "Error!"
Dim RS:Set RS=Server.CreateObject("adodb.recordset")
RS.Open "SELECT TOP 100 a.title,b.* FROM KS_Product A Inner Join KS_ProImages B On A.ID=B.ProID WHERE a.ID=" & ID & " order by b.orderid,b.id",Conn,1,1
Do While Not RS.Eof
     titleStr = rs("title")
     photostr=photostr &"<li><a href=""" & RS("BigPicUrl") & """ hidefocus=""true""><img src=""" & rs("SmallPicUrl") & """ width=""59"" height=""80""  title=""" & rs("groupname") &""" alt=""" & rs("groupname") & """ style=""cursor:pointer;""/></a></li>"
	 if imgs="" then
	 imgs="'" & rs("bigpicurl") & "'"
	 else
	 imgs=imgs &",'" & rs("bigpicurl") & "'"
	 end if
	 if lcase(request("u"))=lcase(rs("bigpicurl")) then curr=n
	 n=n+1
RS.MoveNext
Loop
RS.Close
Set RS=Nothing

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1"><link rel="stylesheet" type="text/css" media="screen" href="/Css/photoslider.css" />
<title>
	<%=titleStr %> 清晰图片
</title>
<script type="text/javascript" src="../KS_Inc/jQuery.JS"></script>

<script>

$(document).ready(function(){
 try{
	document.execCommand("BackgroundImageCache",false,true);
}catch(e)
{}
function $$(){
	var elements=[];
	for(var i=0;i<arguments.length;i++){
		var element=arguments[i];
		if(typeof element=="string"){
			element=document.getElementById(element);
		}
		if(arguments.length==1){
			return element;
		}
		elements.push(element);
	}
	return elements;
}
function getStyle(obj,option){
	if(obj.currentStyle){
		var value=obj.currentStyle[option];
		if(value=="auto")value="0px";
	}else{
		var value=document.defaultView.getComputedStyle(obj,null)[option];
	}
	return value;
}
Function.prototype.bind=function(object){
	var method=this;
	return function(){
		method.apply(object,arguments);	
	}
}
var Class={
	create:function(){
		return function(){
			this.initialize.apply(this,arguments);	
		}
	}
}
var Scroll=Class.create();
Scroll.prototype={
	initialize:function(parent){
	    if($$(parent) == null)
	        return;
		this.parent=$$(parent);
		this.step=2;
		this.speed=30;
		var newIe= $.browser.msie && /MSIE 8.0/.test(navigator.userAgent);
		if(newIe)
		{
		    this.speed=15;
		}
		this.flagLeft=true;
		var obj_ul=this.parent.getElementsByTagName("ul")[0];
		var arr_li=obj_ul.getElementsByTagName("li");
		var obj_li=arr_li[0];
		try{
		var li_width=obj_li.offsetWidth;
		li_width+=parseInt(getStyle(obj_li,"marginLeft"))+parseInt(getStyle(obj_li,"marginRight"));	
		var ul_width=li_width*(arr_li.length);
		ul_width+=parseInt(getStyle(obj_ul,"paddingLeft"))+parseInt(getStyle(obj_ul,"paddingRight"));
		ul_width+=parseInt(getStyle(obj_ul,"marginLeft"))+parseInt(getStyle(obj_ul,"marginRight"));
		this.maxWidth=ul_width; //1547*2=3094
		obj_ul.parentNode.style.width=ul_width*2+"px"; //6188
		obj_ul.innerHTML+=obj_ul.innerHTML;
		if(this.parent.scrollLeft==0)
		{
		   $("#dpyleft img").removeClass("hand").attr("src","../images/default/BLGray.png"); 
		}
        if(this.parent.scrollLeft>=this.maxWidth-886-30){
           $("#dpyright img").removeClass("hand").attr("src","../images/default/BRGray.png");
           return;
        }
		if(arguments.length>1){
			this.arrowLeft=$$(arguments[1]);
			this.arrowLeft.onmouseover=function(){
				this.moveLeft();
				this.flagLeft=true;
			}.bind(this);
			this.arrowLeft.onmouseout=this.stop.bind(this);
		}
		if(arguments.length>2){
			this.arrowRight=$$(arguments[2]);
			this.arrowRight.onmouseover=function(){
				this.moveRight();
				this.flagLeft=false;
			}.bind(this);
			this.arrowRight.onmouseout=this.stop.bind(this);
		}
		}catch(e){
		}
	},
	moveLeft:function(){
        if(this.parent.scrollLeft==0){
			this.stop();
			$("#dpyleft img").removeClass("hand").attr("src","../images/default/BLGray.png");
		}else{
			this.parent.scrollLeft-=this.step;
		}
		if(!$("#dpyright img").hasClass("hand"))
		{
		    $("#dpyright img").addClass("hand").attr("src","../images/default/dpy__r3_c49.png"); 
		}
		this.timer=setTimeout(this.moveLeft.bind(this),this.speed);
	},
	moveRight:function(){
		if(this.parent.scrollLeft>this.maxWidth-886-30){
			this.stop();
			$("#dpyright img").removeClass("hand").attr("src","../images/default/BRGray.png");
		}else{
			this.parent.scrollLeft+=this.step;
		}
		if(!$("#dpyleft img").hasClass("hand"))
		{
		    $("#dpyleft img").addClass("hand").attr("src","../images/default/dpy__r3_c4.png"); 
		}
		this.timer=setTimeout(this.moveRight.bind(this),this.speed);
	},
	stop:function(){
		clearTimeout(this.timer);	
	},
	start:function(){
		if(this.flagLeft){
			this.moveLeft();
		}else{
			this.moveRight();
		}
	}
}
<%if n>10 Then '大于10张才出现滚动效果%>
new Scroll("dpyscroll","dpyleft","dpyright");
<%end if%>
}        
);
</script>

<script type="text/javascript">
$(document).ready(function(){
	$("#thumbnail li a").click(function(){
		$("#large img").hide().attr({"src": $(this).attr("href"), "title": $("> img", this).attr("title")});
		return false;
	});
	$("#large>img").load(function(){$("#large>img:hidden").fadeIn("slow")});

	$("#thumbnail li a img").each(function(){
	    $(this).click(function(){
	       
	       $("#thumbnail li a img").each(function(){
	          $(this).css({"border":"1px solid #efefef","padding":"2px","filter": "Alpha(Opacity=100)"});
	       });
	    
	      $(this).css({"border":"2px solid #ff6600","padding":"1px","filter": "Alpha(Opacity=30)"});
	    });
	});
});
</script>

<style type="text/css">
 body{font-size:12px;padding:0px;margin:0px}
 ul,li{margin:0px;padding:0px}
 li{list-style-type:none}
 img{border:0px}
 h2{text-align:center;font-size:16px;margin:5px;}
#large{clear:both;text-align:center;margin:0 auto;margin-top:20px}
.topbg{height:150px;background:url('images/bg_headerwide.gif') no-repeat left center;}
.gdbjColor{height:100px}
.gdbjColor a:visited{color:#e6e6e6;}
.gdbjColor a:hover{color:#e6e6e6;}
.hand{ cursor:pointer;}
#dpyscroll{width:886px;float:left;display:inline;overflow:hidden; height:86px; margin:10px 0;}
#dpyscroll ul{ list-style:none; margin:0;padding-left:0px;}
#dpyscroll ul li{ width:91px; text-align:left;float:left; height:86px;}
#dpyscroll img{ border:1px solid #efefef;padding:2px}
#large img{border:1px solid #efefef;padding:2px}
</style>
</head>
<body>

    <%if titlestr<>"" then
	   titlestr="<font color=red>“" & titlestr & "”</font>"
	end if%>
    <h2>查看<%=titleStr %>清晰图片</h2>
    <div id="main">
	<%If imgs<>"" then%>
     <div class="topbg" id="topbg">
		<div class="gdbjColor" id="dpyTopScroll">
		   <div style="width:47px; float:left;" id="dpyleft">
		     <img style="padding:42px 30px 0 10px;" class="hand" src="../images/default/BLGray.png" alt="left"/></div>
		     <div id="dpyscroll">
		      <div id="divcontainer">
		        <ul id="thumbnail">
                   <%=photoStr%>
		         </ul>
		       </div>
		     </div>
		     <div style="width:47px; float:left;" id="dpyright"><img style="padding:42px 10px 0 30px;" src="../images/default/dpy__r3_c49.png" class="hand" alt="right"  /></div>
		</div>
    </div>
	<%end if%>
    
    <div id="large">
       <img id="img1" src="<%=ks.CheckXSS(Request.QueryString("u"))%>" alt=""/>
    </div>
       
    </div>
<div style="text-align:center;">
   </div>
<script type="text/javascript">
var Util = {};
Util.Event = {
    stop: function(ent){           
        var e = ent||window.event;
        if (e.preventDefault){
          e.preventDefault();
          e.stopPropagation();
        } 
        else{
          e.returnValue = false;
          e.cancelBubble = true;
        }
    },
    add:function(elem,name,fn,useCapture){
        if (name == 'keypress' &&
        (navigator.appVersion.match(/Konqueror|Safari|KHTML/)
        || elem.attachEvent))
            name = 'keydown';
        if(elem.addEventListener){
            elem.addEventListener(name,fn,useCapture);
        }
        if(elem.attachEvent){
            elem.attachEvent("on"+name,fn);
        }
    },
    getEvent:function() {
        if (window.event) {
            return this.formatEvent(window.event);
        } else {
            return this.getEvent.caller.arguments[0];
        }
    },
    formatEvent:function (oEvent) {
        if (document.all) {
            oEvent.charCode = (oEvent.type == "keypress") ? oEvent.keyCode : 0;
            oEvent.eventPhase = 2;
            oEvent.isChar = (oEvent.charCode > 0);
            oEvent.pageX = oEvent.clientX + document.body.scrollLeft;
            oEvent.pageY = oEvent.clientY + document.body.scrollTop;
            oEvent.layerX = oEvent.offsetX;
            oEvent.layerY = oEvent.offsetY;
            oEvent.preventDefault = function () {
                this.returnValue = false;
            }
            


            if (oEvent.type == "mouseout") {
                oEvent.relatedTarget = oEvent.toElement;
            } else if (oEvent.type == "mouseover") {
                oEvent.relatedTarget = oEvent.fromElement;
            }
            oEvent.stopPropagation = function () {
                this.cancelBubble = true;
            };
            oEvent.target = oEvent.srcElement;
            oEvent.time = (new Date).getTime();
        }
        return oEvent;
    }
}
function $$(element) {
	return document.getElementById(element);
}

var arrowImage1 = new Image();arrowImage1.src = "../images/default/arrow001.gif";
var arrowImage2 = new Image();arrowImage2.src = "../images/default/arrow002.gif";
var NextPageTips = function(obj){
    
    var str = new String('\
                                <div style="width:103px;height:27px; text-align:center;"><img id="cursorImg"  src="../images/default/arrow001.gif" /></div>\
                                <div style="width:103px;height:20px; border:1px solid #ffffff;filter:Alpha(Opacity=70);-moz-opacity: 0.8">\
                                   <div style="width:101px;height:18px;border:1px solid #000000;filter:Alpha(Opacity=60);-moz-opacity: 0.8">\
                                     <div style="width:100%;height:100%; background:#000000; filter:Alpha(Opacity=60);-moz-opacity: 0.6">\
                                     </div>\
                                   </div>\
                                </div>\
                                <span id="NextPageTipsSpan" style="font-size:13px; position:relative; top:-20px;left:8px;color:#ffffff;" ></span>\
                                ');
                                

    Util.Event.add(obj,"mousemove",function(){
    
       var ObjectX = 0;
       ObjectX = Util.Event.getEvent().layerX;
       
            
        if($$('NextPageTips')==null) {
			var oDiv = document.createElement("div");
			oDiv.style.position = "absolute";
			oDiv.style.left = Util.Event.getEvent().pageX + "px";
			oDiv.style.top = Util.Event.getEvent().pageY  + "px";

			oDiv.id = "NextPageTips";
			oDiv.style.height="20px";
			oDiv.style.width="103px";
			document.body.appendChild(oDiv);
			
			$$('NextPageTips').innerHTML = str;
		}
            
		$$('NextPageTips').style.left = Util.Event.getEvent().pageX - 45 + "px";
		$$('NextPageTips').style.top = Util.Event.getEvent().pageY + 10 + "px";
		if(document.all)
		{
		     Util.Event.stop(); 
		}

		var image = new Image();
		image.src = Util.Event.getEvent().target.src; 
		width = image.width;
                    
		 if(ObjectX<Math.floor(width/2)) {
			$$('cursorImg').src = arrowImage1.src;
			   
			$$('NextPageTipsSpan').innerHTML = "点击查看上一张";
			Util.Event.getEvent().target.onclick = function(){
				prePic();
			}
		 }
		 else
		 {
			$$('cursorImg').src = arrowImage2.src;
			$$('NextPageTipsSpan').innerHTML = "点击查看下一张";
			Util.Event.getEvent().target.onclick = function(){
			    nextPic();
			}
		 }
    },false);
                        
	Util.Event.add(obj,"mouseout",function(){
	   if($$('NextPageTips')!=null)
		  document.body.removeChild($$('NextPageTips'));
	},false);                                
};
function prePic() {
	if (i==0) alert('已经是第一张了');
	else img.src = imgs[i--];
}
function nextPic() {
	if (i==imgs.length) alert('已经是最后一张了');
	else img.src = imgs[i++];
}

imgs = new Array(<%=imgs%>);

var img = $$('img1');
img.style.cursor = "url(transMouse.cur),auto";
i = 1;
<%if request("u")="" then%>
img.src = imgs[i];
<%end if%>
<%if imgs<>"" then%>
new NextPageTips(img);
<%end if%>
</script>

<script>document.onkeydown=chang_page;function chang_page(event){var e=window.event||event;var eObj=e.srcElement||e.target;var oTname=eObj.tagName.toLowerCase();if(oTname=='input' || oTname=='textarea' || oTname=='form')return;	event = event ? event : (window.event ? window.event : null);if(event.keyCode==37||event.keyCode==33){prePic()}	if (event.keyCode==39 ||event.keyCode==34){nextPic()}}</script>

</body>
</html>
<%
Set KS=Nothing
CloseConn
%>
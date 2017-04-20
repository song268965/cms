lastScrollY = 0;
function heartBeat(){
var diffY;
if (document.documentElement && document.documentElement.scrollTop)
diffY = document.documentElement.scrollTop;
else if (document.body)
diffY = document.body.scrollTop
else
{/*Netscape stuff*/}
percent=.1*(diffY-lastScrollY);
if(percent>0)percent=Math.ceil(percent);
else percent=Math.floor(percent);
document.getElementById("leftDiv").style.top = parseInt(document.getElementById("leftDiv").style.top)+percent+"px";
document.getElementById("rightDiv").style.top = parseInt(document.getElementById("rightDiv").style.top)+percent+"px";
lastScrollY=lastScrollY+percent;
}
//下面这段删除后，对联将不跟随屏幕而移动。
window.setInterval("heartBeat()",1);
//-->
//关闭按钮
function close_left1(){left1.style.visibility='hidden';}
function close_right1(){right1.style.visibility='hidden';}
//显示样式
document.writeln("<style type=\"text\/css\">");
document.writeln("#leftDiv,#rightDiv{width:100px;height:300px;background-color:#fff;position:absolute;}");
document.writeln(".itemFloat{width:100px;height:auto;line-height:5px}");
document.writeln(".itemFloat img{width:100px;height:300px;}");
document.writeln("<\/style>");
//以下为主要内容
document.writeln("<div id=\"leftDiv\" style=\"top:100px;left:5px\">");
//------左侧各块开始
//---L1
document.writeln("<div id=\"left1\" class=\"itemFloat\">");
if(0==0 || (0==1 && checkDate33('2014/7/22'))){
document.writeln("<span onclick=\"addHits33(0,27)\"><a href=\"http://www.kesion.com\" target=\"_blank\"><img  alt=\"33333333\"  border=\"0\"  src=\"http://bbs.kesion.com/images/ks_mnkc.jpg\"></a></span><br/>");
}
document.writeln("<br><a href=\"javascript:close_left1();\" title=\"关闭上面的广告\">×<\/a><br><br><br><br>");
document.writeln("<\/div>");
//------左侧各块结束
document.writeln("<\/div>");
document.writeln("<div id=\"rightDiv\" style=\"top:100px;right:5px\">");
//------右侧各块结束
//---R1
document.writeln("<div id=\"right1\" class=\"itemFloat\">");
if(1==0 || (1==1 && checkDate33('2014/7/22'))){
document.writeln("<span onclick=\"addHits33(0,28)\"><a href=\"http://11111\" target=\"_blank\"><img  alt=\"ssss\"  border=\"0\"  src=\"http://bbs.kesion.com/images/ks_tg.jpg\"></a></span><br/>");
}
document.writeln("<br><a href=\"javascript:close_right1();\" title=\"关闭上面的广告\">×<\/a><br><br><br><br>");
document.writeln("<\/div>");
//------右侧各块结束
document.writeln("<\/div>");
function addHits33(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate33(date_arr){
 var date=new Date();
 date_arr=date_arr.replace(/\//g,"-").split("-");
var year=parseInt(date_arr[0]);
var month=parseInt(date_arr[1])-1;
var day=0;
if (date_arr[2].indexOf(" ")!=-1)
day=parseInt(date_arr[2].split(" ")[0]);
else
day=parseInt(date_arr[2]);
var date1=new Date(year,month,day);
if(date.valueOf()>date1.valueOf())
 return false;
else
 return true
}

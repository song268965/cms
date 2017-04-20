document.write("<span id='s31'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a31=new Array();
var t31=new Array();
var ts31=new Array();
var allowclick31=new Array();
var id31=new Array();
a31[0]="<span onclick=\"addHits31(0,26)\"><a href=\"http://www.kesion.com\" target=\"_blank\"><img  alt=\"科汛NET版\"  border=\"0\"  height=60  width=998  src=\"/images/2013121819393566584.gif\"></a></span>";
t31[0]=0;
ts31[0]="2014-6-19";
allowclick31[0]=0;
id31[0]=26;
var temp31=new Array();
var k=0;
for(var i=0;i<a31.length;i++){
if (t31[i]==1){
if (checkDate31(ts31[i])){
	temp31[k++]=a31[i];
}
	}else{
 temp31[k++]=a31[i];
}
}
if (temp31.length>0){
GetRandom(temp31.length);
var index31=GetRandomn-1;
if (allowclick31[index31]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick31[index31]+'&id='+id31[index31],function(){  
$('#s31').html(a31[index31]);
if (data.isEnd=='1'){
$('#s31').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s31').find("span").click(function(){return false; });

 } });
}else{
$('#s31').html(a31[index31]);
}
}
function addHits31(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate31(date_arr){
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

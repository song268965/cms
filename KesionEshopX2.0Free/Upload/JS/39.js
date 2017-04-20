document.write("<span id='s39'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a39=new Array();
var t39=new Array();
var ts39=new Array();
var allowclick39=new Array();
var id39=new Array();
a39[0]="<span onclick=\"addHits39(0,38)\"><a href=\"http://\" target=\"_blank\"><img  alt=\"广告横幅\"  border=\"0\"  src=\"/images/20151071452.jpg\"></a></span>";
t39[0]=0;
ts39[0]="2017-3-16";
allowclick39[0]=0;
id39[0]=38;
var temp39=new Array();
var k=0;
for(var i=0;i<a39.length;i++){
if (t39[i]==1){
if (checkDate39(ts39[i])){
	temp39[k++]=a39[i];
}
	}else{
 temp39[k++]=a39[i];
}
}
if (temp39.length>0){
GetRandom(temp39.length);
var index39=GetRandomn-1;
if (allowclick39[index39]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick39[index39]+'&id='+id39[index39],function(){  
$('#s39').html(a39[index39]);
if (data.isEnd=='1'){
$('#s39').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s39').find("span").click(function(){return false; });

 } });
}else{
$('#s39').html(a39[index39]);
}
}
function addHits39(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate39(date_arr){
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

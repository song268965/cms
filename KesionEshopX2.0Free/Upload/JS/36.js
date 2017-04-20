document.write("<span id='s36'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a36=new Array();
var t36=new Array();
var ts36=new Array();
var allowclick36=new Array();
var id36=new Array();
a36[0]="<span onclick=\"addHits36(0,40)\"><a href=\"http://\" target=\"_blank\"><img  alt=\"x1.5\"  border=\"0\"  src=\"http://192.168.0.2:95/images/20151022033.jpg\"></a></span>";
t36[0]=0;
ts36[0]="2016-12-30";
allowclick36[0]=0;
id36[0]=40;
var temp36=new Array();
var k=0;
for(var i=0;i<a36.length;i++){
if (t36[i]==1){
if (checkDate36(ts36[i])){
	temp36[k++]=a36[i];
}
	}else{
 temp36[k++]=a36[i];
}
}
if (temp36.length>0){
GetRandom(temp36.length);
var index36=GetRandomn-1;
if (allowclick36[index36]>0){ 
jQuery.getScript('http://192.168.0.2:95/plus/ads/showA.asp?action=loadjs&times='+allowclick36[index36]+'&id='+id36[index36],function(){  
$('#s36').html(a36[index36]);
if (data.isEnd=='1'){
$('#s36').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s36').find("span").click(function(){return false; });

 } });
}else{
$('#s36').html(a36[index36]);
}
}
function addHits36(c,id){if(c==1){try{jQuery.getScript('http://192.168.0.2:95/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate36(date_arr){
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

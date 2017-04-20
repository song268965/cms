document.write("<span id='s38'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a38=new Array();
var t38=new Array();
var ts38=new Array();
var allowclick38=new Array();
var id38=new Array();
a38[0]="<span onclick=\"addHits38(0,37)\"><a href=\"http://\" target=\"_blank\"><img  alt=\"广告横幅\"  border=\"0\"  src=\"/images/20151061454.jpg\"></a></span>";
t38[0]=0;
ts38[0]="2016-12-30";
allowclick38[0]=0;
id38[0]=37;
var temp38=new Array();
var k=0;
for(var i=0;i<a38.length;i++){
if (t38[i]==1){
if (checkDate38(ts38[i])){
	temp38[k++]=a38[i];
}
	}else{
 temp38[k++]=a38[i];
}
}
if (temp38.length>0){
GetRandom(temp38.length);
var index38=GetRandomn-1;
if (allowclick38[index38]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick38[index38]+'&id='+id38[index38],function(){  
$('#s38').html(a38[index38]);
if (data.isEnd=='1'){
$('#s38').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s38').find("span").click(function(){return false; });

 } });
}else{
$('#s38').html(a38[index38]);
}
}
function addHits38(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate38(date_arr){
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

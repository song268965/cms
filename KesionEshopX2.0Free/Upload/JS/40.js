document.write("<span id='s40'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a40=new Array();
var t40=new Array();
var ts40=new Array();
var allowclick40=new Array();
var id40=new Array();
a40[0]="<span onclick=\"addHits40(0,39)\"><a href=\"http://\" target=\"_blank\"><img  alt=\"广告横幅\"  border=\"0\"  src=\"http://demo.kesion.com/images/20151022033.jpg\"></a></span>";
t40[0]=0;
ts40[0]="2015-10-8";
allowclick40[0]=0;
id40[0]=39;
var temp40=new Array();
var k=0;
for(var i=0;i<a40.length;i++){
if (t40[i]==1){
if (checkDate40(ts40[i])){
	temp40[k++]=a40[i];
}
	}else{
 temp40[k++]=a40[i];
}
}
if (temp40.length>0){
GetRandom(temp40.length);
var index40=GetRandomn-1;
if (allowclick40[index40]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick40[index40]+'&id='+id40[index40],function(){  
$('#s40').html(a40[index40]);
if (data.isEnd=='1'){
$('#s40').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s40').find("span").click(function(){return false; });

 } });
}else{
$('#s40').html(a40[index40]);
}
}
function addHits40(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate40(date_arr){
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

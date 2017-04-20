document.write("<span id='s28'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a28=new Array();
var t28=new Array();
var ts28=new Array();
var allowclick28=new Array();
var id28=new Array();
a28[0]="<span onclick=\"addHits28(1,23)\"><a href=\"http://www.kesion.com/e/\" target=\"_blank\"><img  alt=\"科汛网校\"  border=\"0\"  height=90  width=250  src=\/images/250-90.gif\"></a></span>";
t28[0]=0;
ts28[0]="2015-9-17";
allowclick28[0]=8;
id28[0]=23;
var temp28=new Array();
var k=0;
for(var i=0;i<a28.length;i++){
if (t28[i]==1){
if (checkDate28(ts28[i])){
	temp28[k++]=a28[i];
}
	}else{
 temp28[k++]=a28[i];
}
}
if (temp28.length>0){
GetRandom(temp28.length);
var index28=GetRandomn-1;
if (allowclick28[index28]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick28[index28]+'&id='+id28[index28],function(){  
$('#s28').html(a28[index28]);
if (data.isEnd=='1'){
$('#s28').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s28').find("span").click(function(){return false; });

 } });
}else{
$('#s28').html(a28[index28]);
}
}
function addHits28(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate28(date_arr){
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

document.write("<span id='s37'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a37=new Array();
var t37=new Array();
var ts37=new Array();
var allowclick37=new Array();
var id37=new Array();
a37[0]="<span onclick=\"addHits37(0,36)\"><a href=\"http://\" target=\"_blank\"><img  alt=\"广告\"  border=\"0\"  src=\"/images/20151050947.jpg\"></a></span>";
t37[0]=0;
ts37[0]="2016-12-30";
allowclick37[0]=0;
id37[0]=36;
var temp37=new Array();
var k=0;
for(var i=0;i<a37.length;i++){
if (t37[i]==1){
if (checkDate37(ts37[i])){
	temp37[k++]=a37[i];
}
	}else{
 temp37[k++]=a37[i];
}
}
if (temp37.length>0){
GetRandom(temp37.length);
var index37=GetRandomn-1;
if (allowclick37[index37]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick37[index37]+'&id='+id37[index37],function(){  
$('#s37').html(a37[index37]);
if (data.isEnd=='1'){
$('#s37').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s37').find("span").click(function(){return false; });

 } });
}else{
$('#s37').html(a37[index37]);
}
}
function addHits37(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate37(date_arr){
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

document.write("<span id='s29'></span>");
var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a29=new Array();
var t29=new Array();
var ts29=new Array();
var allowclick29=new Array();
var id29=new Array();
a29[0]="<span onclick=\"addHits29(0,24)\"><a href=\"http://www.kesion.com\" target=\"_blank\"><img  alt=\"KESIONCMS X1\"  border=\"0\"  height=80  width=998  src=\"/images/cmsbanner.png\"></a></span>";
t29[0]=0;
ts29[0]="2014-6-19";
allowclick29[0]=0;
id29[0]=24;
a29[1]="<span onclick=\"addHits29(0,34)\"><a href=\"http://192.168.0.10/images/ad680.gif\" target=\"_blank\"><img  alt=\"广告\"  border=\"0\"  src=\"http://demo.kesion.com/images/ad680.gif\"></a></span>";
t29[1]=0;
ts29[1]="2015-9-30";
allowclick29[1]=0;
id29[1]=34;
var temp29=new Array();
var k=0;
for(var i=0;i<a29.length;i++){
if (t29[i]==1){
if (checkDate29(ts29[i])){
	temp29[k++]=a29[i];
}
	}else{
 temp29[k++]=a29[i];
}
}
if (temp29.length>0){
GetRandom(temp29.length);
var index29=GetRandomn-1;
if (allowclick29[index29]>0){ 
jQuery.getScript('/plus/ads/showA.asp?action=loadjs&times='+allowclick29[index29]+'&id='+id29[index29],function(){  
$('#s29').html(a29[index29]);
if (data.isEnd=='1'){
$('#s29').find("a").attr("href","javascript:;").click(function(){ alert("对不起，该广告今天已达到点击上限！"); return false; });
$('#s29').find("span").click(function(){return false; });

 } });
}else{
$('#s29').html(a29[index29]);
}
}
function addHits29(c,id){if(c==1){try{jQuery.getScript('/plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate29(date_arr){
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

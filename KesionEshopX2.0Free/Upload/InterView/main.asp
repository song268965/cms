<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="session.asp"-->
<%
Dim KS:Set KS=New Publiccls
Dim ID:ID=KS.ChkClng(request("id"))
if id=0 then id=KS.ChkClng(KS.C("InterViewID"))
if id=0 then ks.die "<script>alert('参数出错!');window.close();</script>"
dim rs:set rs=conn.execute("select top 1 * from KS_InterView Where ID=" & id)
if rs.eof and rs.bof then
 rs.close
 set rs=nothing
 ks.die "<script>alert('对不起，访谈主题不存在!');window.close();</script>"
end if
if rs("locked")="1" then 
 rs.close
 set rs=nothing
 ks.die "<script>alert('对不起，该访谈已结束!');window.close();</script>"
end if
dim title:title=rs("title")
dim host:host=rs("host")
dim guests:guests=rs("guests")
rs.close
set rs=nothing


%>
<!DOCTYPE html>
<html>
<head>
<title><%=KS.Setting(0)%>---在线访谈系统</title>
<script type="text/JavaScript" src="../ks_inc/jquery.js"></script>
<script type="text/JavaScript" src="../ks_inc/common.js"></script>
<script src="../KS_Inc/DatePicker/WdatePicker.js"></script>
<script type="text/javascript" src="../ks_Inc/lhgdialog.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<style type="text/css"> 
	html{color:#000;font-family:Arial,sans-serif;font-size:12px;}
	h1, h2, h3, h4, h5, h6, h7, p, ul, ol,div,span, dl, dt, dd, li, body,em,i, form, input,i,cite, button, img, cite, strong,    em,label,fieldset,pre,code,blockquote, table, td, th ,tr{ padding:0; margin:0;outline:0 none;}
	img, table, td, th ,tr { border:0;}
	address,caption,cite,code,dfn,em,th,var{font-style:normal;font-weight:normal;}
	select,img,select{font-size:12px;vertical-align:middle;color:#666; font-family:Arial,sans-serif}
	.checkbox{vertical-align:middle;margin-right:5px;margin-top:-2px; margin-bottom:1px;}
	textarea{font-size:12px;color:#666; font-family:Arial,sans-serif}
	.textbox{font-size:12px;color:#666; height:23px;line-height:23px;font-family:Arial,sans-serif}
	table{border-collapse:collapse;border-spacing:0;}
	ul, ol, li { list-style-type:none;}
	a { color:#0082cb; text-decoration:none;}
	a:hover{text-decoration:none;}
	ul:after,.clearfix:after { content: "."; display: block; height: 0; clear: both; visibility: hidden; }/* 不适合用clear时使用 */
	ul,.clearfix{ zoom:1;}
	.clear{clear:both;font-size:0px; line-height:0px;height:1px;overflow:hidden;}/*  空白占位  */
	body {margin:0 auto;font-size:12px; background:#E0F1FB;color:#666;position:relative}



    .head{background:url(images/bg_head.png);height:50px;font-weight:bold;color:#fff;font-size:14px}
    .head span{padding-left:20px;color:#333333;font-size:12px}
	.foot{background:url(images/bg_head.png);height:40px;color:#fff;line-height:30px}
	.foot a{color:#ffffff;}
	.foot a:visited{ color:#fff;}
	.maintitle{border-bottom:1px solid #999;line-height:25px;font-weight:bold;background:#f1f1f1;height:25px;padding-left:5px}
	.maintitle span{font-weight:normal;}
	.mytable{}
	.mytable .title td{background:#f1f1f1;font-weight:bold;height:25px;line-height:25px;}
	.mytable .splittd td{border-bottom:1px dashed #999;height:25px;}
	.jb td{padding:4px;font-size:14px;background:#FFFFCC}
	.zrr td{padding:4px;font-size:14px;background:#FFFFFF;border-bottom:1px solid #999;border-top:1px solid #999}
</style>
</head>
<body style="overflow:hidden" scroll="no" onResize="resize()">
<script>
 function resize(){
  var h=$(window).height()-$("#topframe").height()-$("#bottomframe").height();
  $("#mainframe").height(h);
  $("#interviewMsg").attr("style","overflow-x:hidden;overflow-y:auto;height:"+(h-26)+"px");
  $("#interviewText").attr("style","overflow-x:hidden;overflow-y:auto;height:"+(h-26)+"px");
 }
 $(document).ready(function(){
   resize();
  showTextRecord();
  showMsg();

  
 });

var timer1=null;
var timer=null; 
var editTime=0;
var editTime2=0;

function setEditTime(){
	 if ($("#editDateBtn").val()=="保存批量修改时间"){
		 editTime=0;
		 $("#editDateBtn").val("批量修改时间");
		 $("#myform1").submit();
	 }else{
		 editTime=1;
		 window.clearTimeout(timer1); 
		 $("#editDateBtn").val("保存批量修改时间");
		 showTextRecord();
	 }
}
function setEditTime2(){
	 if ($("#editDateBtn2").val()=="保存批量修改时间"){
		 editTime2=0;
		 $("#editDateBtn2").val("批量修改时间");
		 $("#myform2").submit();
	 }else{
		 editTime2=1;
		 window.clearTimeout(timer); 
		 $("#editDateBtn2").val("保存批量修改时间");
		 showMsg();
	 }
}

//显示文字实录
function showTextRecord(){
 $.ajax({
			url: "adminajax.asp",
			cache: false,
			data: "action=loadtextrecord&id=<%=id%>&editTime="+editTime,
			success: function(r){
			r=unescape(r);
			$("#interviewText").html(r)
		 }
	 });
	if (editTime==0){
		 if ($("#textrefresh").attr("checked")){
		 timer1=setTimeout('showTextRecord();',parseInt($("#textpertime").val())*1000);
		 }
	 }
}
//显示网友留言
function showMsg(){
	 $.ajax({
			url: "adminajax.asp",
			cache: false,
			data: "action=loadmsg&id=<%=id%>&editTime2="+editTime2,
			success: function(r){
			r=unescape(r);
			$("#interviewMsg").html(r)
									
		 }
	 });
	if (editTime2==0){ 
	if ($("#msgrefresh").attr("checked")){
	 timer=setTimeout('showMsg();',parseInt($("#pertime").val())*1000);
	 }
	 }

}
function SetRefresh(v){
 if (v==true){
  timer=setTimeout('showMsg();', parseInt($("#pertime").val())*1000);
 }else{
  window.clearTimeout(timer); 
 }
}
function SetTextRefresh(v){
if (v==true){
  timer1=setTimeout('showTextRecord();', parseInt($("#textpertime").val())*1000);
 }else{
  window.clearTimeout(timer1); 
 }
}
function mselect(id){
  $("#role").val("网友");
  $("#content").val("【网友："+$('#r'+id).html()+"】"+$('#m'+id).html());
}
function delMsg(id){
  if (confirm('确定删除吗？')){
   $.ajax({
			url: "adminajax.asp",
			cache: false,
			data: "action=delmsg&id=<%=id%>&msgid="+id,
			success: function(r){
			if (r=='success'){
			 $.dialog.alert('恭喜删除成功！',function(){
			showMsg();
			 });
			}else{
			 $.dialog.alert('error!');
			}
		 }
	 });
  }
}
function verifyMsg(id,v){
   $.ajax({
			url: "adminajax.asp",
			cache: false,
			data: "action=verifymsg&id=<%=id%>&msgid="+id+"&v="+v,
			success: function(r){
			if (r=='success'){
			 $.dialog.alert('恭喜操作成功！',function(){
			showMsg();
			 });
			}else{
			 $.dialog.alert('error!');
			}
		 }
	 });
}
function postTextRecord(){
 var role=$("#role").val();
 var content=$("#content").val();
 if ($("#adddate").val()==''){
  $.dialog.alert('请输入直播时间！ ',function(){
  $("#adddate").focus();
  });
 }
 if (content==''){
  $.dialog.alert('请输入文字实录内容！ ',function(){
  $("#content").focus();
  });
 }
    $.ajax({
			url: "adminajax.asp",
			cache: false,
			data: "action=savetextrecord&adddate="+$("#adddate").val()+"&id=<%=id%>&role="+escape(role)+"&content="+escape(content),
			success: function(r){
			if (r=='success'){
			 $.dialog.alert('实录内容提交成功！',function(){
			 $("#content").val('')
			 showTextRecord();
			 });;
			}else{
			 $.dialog.alert(unescape(r));
			}
		 }
	 });
 
 
}
function delTextRecord(id){
if (confirm('确定删除该文字实录内容吗？')){
   $.ajax({
			url: "adminajax.asp",
			cache: false,
			data: "action=deltextrecord&id=<%=id%>&msgid="+id,
			success: function(r){
			if (r=='success'){
			 $.dialog.alert('恭喜删除成功！',function(){
			 showTextRecord()
			 });
			}else{
			 $.dialog.alert('error!');
			}
		 }
	 });
  }
}

</script>
<iframe src="about:blank" name="hidframe" id="hidframe" style="display:none"></iframe>
 <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td id="topframe" class="head">
	
	您好<font style='color:red'><%
	 if Not KS.IsNul(KS.C("AdminName")) then
	  response.write KS.C("AdminName")
	 Else
	  response.write KS.C("InterViewUserName")
	 End If%></font>！欢迎登录在线访谈系统!
	当前访谈主题：<%=title%>
	
	</td>
	<td class="head">
	<span style="float:right;padding-right:20px">
	<a href="login.asp?action=LoginOut&id=<%=id%>" onClick="return(confirm('确定安全退出吗？'));" style="color:#fff;font-weight:bold;cursor:pointer">安全退出</a></span><span style="font-weight:normal;color:#fff">主持人：<%=host%>  嘉宾：<%=guests%></span>
	</td>
  </tr>
  <tr>
    <td id="mainframe" valign="top" style="border-right:1px solid #999" width="60%">
	 <div class="maintitle">文字实录：
	 
	 <span style="float:right;padding-right:20px;">
	 <input type="button" value="批量修改时间" id="editDateBtn" onclick="setEditTime();" class="button"/>
	 </span>
	  <span>
	  <label><input type='checkbox' name='textrefresh' id='textrefresh' value='1' onClick="SetTextRefresh(this.checked)" checked/>自动刷新</label>
	   &nbsp;&nbsp;间隔：<select name="textpertime" id="textpertime">
	    <option value="5">5秒</option>
	    <option value="10">10秒</option>
	    <option value="15">15秒</option>
	    <option value="20">20秒</option>
	    <option value="25">25秒</option>
	    <option value="30">30秒</option>
	   </select>
	   <input type="button" value="刷新" onClick="showTextRecord()" class="button"/>
	   
	   
	 </span>
	 
	 </div>
	 <form name="myform1" id="myform1" action="savedate.asp" method="post" target="hidframe">
	  <input type="hidden" value="savedate" name="action"/>
	  <input type="hidden" value="<%=ID%>" name="id"/>
	  <div id="interviewText">loading...</div>
	 </form>
	</td>
    <td valign="top">
	 <div class="maintitle">网友留言：
	 <span style="float:right;padding-right:20px;">
	 <input type="button" value="批量修改时间" id="editDateBtn2" onclick="setEditTime2();" class="button"/>
	 </span>
	 
	 <span>
	  <label><input type='checkbox' name='msgrefresh' id='msgrefresh' value='1' onClick="SetRefresh(this.checked)" checked/>自动刷新</label>
	   &nbsp;&nbsp;间隔：<select name="pertime" id="pertime">
	    <option value="5">5秒</option>
	    <option value="10">10秒</option>
	    <option value="15">15秒</option>
	    <option value="20">20秒</option>
	    <option value="25">25秒</option>
	    <option value="30">30秒</option>
	   </select>
	   <input type="button" value="刷新" onClick="showMsg()" class="button"/>
	 </span>
	 </div>
	 <form name="myform2" id="myform2" action="savedate.asp" method="post" target="hidframe">
	  <input type="hidden" value="savedate2" name="action"/>
	  <input type="hidden" value="<%=ID%>" name="id"/>
	 <div id="interviewMsg" class="interviewMsg">
	 loading...
	 </div>
	</form>
	 
	 
	</td>
  </tr>
  <tr>
    <td id="bottomframe" colspan="2" class="foot">
	
	
	<strong>访谈身份：</strong>
	<select name="role" id="role">
	     
		 <option value="主持人">主持人</option>
			<%
									 if not ks.isnul(guests) then
									   dim i,garr:garr=split(guests,",")
									   for i=0 to ubound(garr)
									    if not ks.isnul(garr(i)) then
										 if KS.C("InterRole")=garr(i) then
										  response.write "<option value='" & garr(i) &"' selected>" & garr(i) &"</option>"
										 else
									       response.write "<option value='" & garr(i) &"'>" & garr(i) &"</option>"
										 end if
										end if
									   next
									 end if
		  %>
		<option value="网友">网友</option>							 
	  </select>
	直播时间：<input type="text" name="adddate" id="adddate" value="<%=now%>" class="textbox" onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"/>
	直播内容：<textarea name="content" id="content" style="margin-top:5px;width:600px;height:30px"></textarea>
	
	<input type="button" value=" 发 言 " style="padding:2px" class="button" onClick="postTextRecord()"/>
	</td>
  </tr>
</table>

 
</body>
</html>


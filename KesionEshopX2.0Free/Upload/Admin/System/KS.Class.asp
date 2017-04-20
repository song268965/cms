<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<!--#include file="../../KS_Cls/Kesion.ClassCls.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
Server.ScriptTimeOut=9999999
Dim KS:Set KS = New PublicCls
Dim KSCls:Set KSCls=New ManageCls
Dim Action,TitleColor,ChannelDir,strOption,ChannelPath
Dim RsObj,i,Flag,HtmlFileDir,ClassDir,strClassDir,MyFolderName,TotalNum
Dim moduleid,UseHtml,IsCreateHtml,strClass,sModuleName,FolderID,Go,TempStr
FolderID = Trim(Request("FolderID")):If FolderID = "" Then FolderID = "0"
Go=KS.G("Go")
Action=KS.G("Action")
Dim MaxPerPage,TotalPut
MaxPerPage=20

If Action="" Then
	if request("channelid")<>"" then  session("class_modelid")=request("channelid"):session("class_from")="" Else session("class_modelid")="":session("class_from")="all"
	if request("from")="all" then session("class_modelid")=""
End If
MyFolderName=KS.GetClassName(KS.ChkClng(session("class_modelid")))

If Action="ExtSub" Then
	response.cachecontrol="no-cache"
	response.addHeader "pragma","no-cache"
	response.expires=-1
	response.expiresAbsolute=now-1
	Response.CharSet="utf-8"
	Call SubTreeList(KS.G("TN"))
	Response.End()
End If
	If Not KS.ReturnPowerResult(KS.ChkClng(KS.S("ChannelID")), "M010001") Then                  '栏目权限检查
	Call KS.ReturnErr(1, "")   
	Response.End()
	End iF
	Response.CharSet="utf-8"

Select case Action
 Case "Add","Edit"  CreateClass 
 Case "Del"   DelClass
 Case "DelInfo"  DelInfo
 Case "Unite" ClassHead: Unite
 Case "UniteSave"  UniteSave
 Case "MoveInfo" ClassHead :MoveInfo 
 Case "DoMoveToClass" DoMoveToClass
 Case "Attribute" ClassHead: Attribute
 Case "DoBatch" AttributeSave
 Case "OrderOne" ClassHead :OrderOne
 Case "DoUpOrderSave" DoUpOrderSave
 Case "DoDownOrderSave" DoDownOrderSave
 Case "OrderN" ClassHead:OrderN
 Case "DoUpOrderNSave" DoUpOrderNSave
 Case "DoDownOrderNSave" DoDownOrderNSave
 Case "MoveClass" MoveClass
 Case "MoveSave" MoveSave
 Case "DoOrder" ClassHead :DoOrder
 Case Else ClassHead : MainPage
End Select

Sub ClassHead()
%><!DOCTYPE html>
<html>
<link href="../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript" src="../../KS_Inc/Jquery.js"></script>
<script language="JavaScript" src="../../KS_Inc/common.js"></script>
<head>
<script language="JavaScript">
var oids='';
function ExtSub(ID){
 if ($('#sub'+ID).html()==''||$('#sub'+ID).html()==undefined){
	  if (oids==''){ //记录打开状态的ID
	   oids=ID;
	  }else{
	   oids=oids+','+ID;
	  }
	 $('#C'+ID).attr('src','../images/folder/open.gif');
	 $('#sub'+ID).html('<div style="padding-left:20px"><img src=../images/loading.gif>子<%=MyFolderName%>加载中...</div>');
	 $(parent.document).find("#ajaxmsg").toggle();
	 $.ajax({
	   type: "POST",
	   url: "KS.Class.asp",
	   data: "tn="+ID+"&action=ExtSub&channelid=<%=request("channelid")%>",
	   success: function(data){
			$(parent.document).find("#ajaxmsg").toggle();
			$("#sub"+ID).html(data);
	   }
	});
}else{
	 var oarr=oids.split(',');
	 oids='';
	 for(var i=0;i<oarr.length;i++){
	  if (oarr[i]!=ID){
	   if (oids==''){
		oids=oarr[i];
	   }else{
		oids+=','+oarr[i];
		}
	  }
	 }
	 $('#sub'+ID).html('');
	 $('#C'+ID).attr('src','../images/folder/close.gif');
 }
 $("#oids").val(oids);
}
function o(o,id){
	if (o!=null){
		 if (o.prev().prev().attr('src').indexOf("Open.gif")==-1){
		  o.prev().prev().attr('src','../images/folder/Open.gif');
		  $('.my'+id).show();
		 }else{
		  o.prev().prev().attr('src','../images/folder/Close.gif');
		  $('.my'+id).hide();
		 }
	}
}
function CreateHtml()
{   var ids=get_Ids(document.myform);
	if (ids!='')
	    top.openWin('发布选中的<%=MyFolderName%>','Include/RefreshHtmlSave.Asp?Types=Folder&RefreshFlag=IDS&ID='+ids,false,530,108);
	else 
		top.$.dialog.alert('请选择要发布的<%=MyFolderName%>!');
}		

function CreateClass(FolderID)
{
  location.href='KS.Class.asp?Action=Add&channelid=<%=request("channelid")%>&Go=Class&FolderID='+FolderID+'&oids='+oids;;
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<font color=red>添加<%=MyFolderName%></font>")+'&ButtonSymbol=Go&Go=Class&FolderID='+FolderID;
}
function DelClass(){
  var ids=get_Ids(document.myform);
  if (ids!=''){
   top.$.dialog.confirm("删除<%=MyFolderName%>操作将删除此<%=MyFolderName%>中的所有子<%=MyFolderName%>和文档，并且不能恢复！确定要删除此<%=MyFolderName%>吗?",function(){
   $("#myform").submit();
  },function(){
  });
}else{
	 top.$.dialog.alert('请选择要删除的<%=MyFolderName%>!');
 }
}
function EditClass(FolderID)
{
 location.href='KS.Class.asp?Page=<%=CurrentPage%>&channelid=<%=request("channelid")%>&a=<%=request("a")%>&Action=Edit&Go=Class&FolderID='+FolderID+'&oids='+oids;
 $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<font color=red>编辑<%=MyFolderName%></font>")+'&ButtonSymbol=GoSave&Go=Class&FolderID='+FolderID;
}
function MoveClass(FolderID){
	  top.openWin('移动栏目','System/KS.Class.asp?action=MoveClass&ID='+FolderID,false,530,280);
}
function AddInfo(BasicType,C_id,ClassID)
{ 
  switch (BasicType)
  {
   case 1:location.href='../Article/KS.Article.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
   case 2:location.href='../Photo/KS.Picture.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
   case 3:location.href='../DownLoad/KS.Down.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
   case 4:location.href='../Flash/KS.Flash.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
   case 5:location.href='../Shop/KS.Shop.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
   case 7:location.href='../Movie/KS.Movie.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
   case 8:location.href='../Supply/KS.Supply.asp?ChannelID='+C_id+'&Action=Add&FolderID='+ClassID; 
   break;
  }
   $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ChannelID='+C_id+'&OpStr='+escape("添加信息")+'&ButtonSymbol=AddInfo&FolderID='+ClassID; 
}
function DelInfo(C_id,FolderID){if(confirm('清空<%=MyFolderName%>将把<%=MyFolderName%>（包括子<%=MyFolderName%>）的所有文档删除！确定要清空此<%=MyFolderName%>吗？')){
 location.href='KS.Class.asp?ChannelID='+C_id+'&Action=DelInfo&Go=Class&FolderID='+FolderID;
}}
function UniteClass(){
  location.href='KS.Class.asp?Action=Unite';
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<%=MyFolderName%>管理 >> <font color=red>合并<%=MyFolderName%></font>")+'&ButtonSymbol=Disabled';}
function OrderOne(){
  location.href='KS.Class.asp?Action=OrderOne';
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<%=MyFolderName%>管理 >> <font color=red>一级<%=MyFolderName%>排序</font>")+'&ButtonSymbol=Disabled';}
function OrderN(){
  location.href='KS.Class.asp?Action=OrderN';
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<%=MyFolderName%>管理 >> <font color=red>N级<%=MyFolderName%>排序</font>")+'&ButtonSymbol=Disabled';}
function MoveClassInfo(){
  location.href='KS.Class.asp?Action=MoveInfo';
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<%=MyFolderName%>管理 >> <font color=red>移动<%=MyFolderName%></font>")+'&ButtonSymbol=Disabled';}
function SetAttribute(){
  location.href='KS.Class.asp?Action=Attribute';
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<%=MyFolderName%>管理 >> <font color=red>批量设置</font>")+'&ButtonSymbol=Disabled';}
function DoOrder(){
  location.href='KS.Class.asp?Action=DoOrder';
  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("<%=MyFolderName%>管理 >> <font color=red>一键更新栏目排序</font>")+'&ButtonSymbol=Disabled';}
function ConfirmUnite(){
 if ($("#FolderID1 option:selected").val()==undefined)
 {
  alert("请选择源<%=MyFolderName%>!");
  return false;
 }
 if ($("#FolderID2 option:selected").val()==undefined)
 {
  alert("请选择目标<%=MyFolderName%>!");
  return false;
 }
  if ($("#FolderID1 option:selected").val()==$("#FolderID2 option:selected").val())
  {    alert('请不要在相同<%=MyFolderName%>内进行操作！'); 
      $("select[name=FolderID2]").focus(); 
	  return false;
  } 
  $("form[name=myform]").submit();
}
$(function(){
 <%if request("oids")<>"" then%>
  var roids='<%=request("oids")%>';
  if (roids!=''){
    var oarr=roids.split(',');
	for(var i=0;i<oarr.length;i++){
	 ExtSub(oarr[i]);
	}
  }
 <%end if%>
});
</script>
</head>
<body>
      <% If KS.S("From")<>"main" Then  
  With KS
    .echo "<ul id='menu_top' class='menu_top_fixed'>"
	.echo "<li onclick='CreateClass(0);' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加"& MyFolderName&"</span></li>"
	.echo "<li onclick='DelClass();' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon delete'></i>删除"& MyFolderName&"</span></li>"
	if session("class_from")="all" or (request("channelid")<>"" and KS.C_S(request("channelid"),7)<>0) Then
	.echo "<li onclick='CreateHtml();' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon move'></i>发布"& MyFolderName&"</span></li>"
	end if
	if session("class_from")="all" Then
	.echo "<li onclick='UniteClass();' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon merge'></i>"& MyFolderName&"合并</span></li>"
	.echo "<li onclick='OrderOne();' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon sorting'></i>一级排序</span></li>"
	.echo "<li onclick='OrderN();' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon sorting'></i>N级排序</span></li>"
	.echo "<li class='parent' onclick='MoveClassInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add2'></i>文档移动</span></li>"
	.echo "<li onclick='javascript:SetAttribute()' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>批量设置</span></li>"
	.echo "<li onclick='javascript:DoOrder()' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon num'></i>更新排序</span></li>"
	.echo "<li class='parent' onclick=""location.href='KS.Class.asp';"""
	if KS.G("Action")="" Then .echo " disabled"
	.echo"><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>回上一级</span></li>"
	end if

	.echo "<div class=""quicktz""><a href='?a=extall&from=" & request("from") & "&channelid=" & request("channelid") &"'><i class='icon merge'></i>全部展开</a></div>"
	.echo "</ul>"
	 .echo "<div class=""menu_top_fixed_height""></div>"
 End With
	End If
End Sub

'创建栏目
Public Sub CreateClass()
    If KS.G("Flag")="Save" Then
		 Call ClassAddSave()
	Else
		 Dim KMO:Set KMO = New ClassCls
		 Call KMO.GetAddChannelFolder(Action,FolderID, "KS.Class.asp?Action=" & Action & "&Flag=Save&Go="&Go&"&FolderID=" & FolderID)
		 Set KMO = Nothing
   End If
End Sub

'保存栏目的新建
Sub ClassAddSave()
	Dim KMO:Set KMO = New ClassCls
    Call KMO.ChannelFolderAddSave (Go)
	Set KMO = Nothing
End Sub

Sub DelClass()

	Dim K, ID, ParentID, OrderID,Root,Depth,CurrPath,RS,FolderID,Sql, Folder, ClassType,C_ID,RSC
	FolderID=KS.G("ID")
	If FolderID="" Then KS.AlertHintScript "对不起，您没有选择要删除的"& MyFolderName&"!"
	FolderID=Replace(Replace(FolderID,",","','")," ","")
	Set RSC=Server.CreateObject("ADODB.RECORDSET")
	RSC.Open "Select ID From KS_Class Where ID in('" & FolderID & "') order by root,folderorder",conn,1,1
    Do While Not RSC.Eof 
   	 Set RS=Server.CreateObject("ADODB.Recordset")
	 
	 'on error resume next
	 Sql = "select * from KS_Class where ts Like '%" & RSC(0) & ",%' order by root,folderorder desc"
	 if err then 
	  err.clear
	  exit do
	 end if
	 
		 RS.Open Sql, conn, 1, 1
			 Do While Not RS.Eof 
			    ID=RS("ID")
				C_ID=RS("ChannelID")
				ParentID = RS("TN")
				Depth=RS("tj")
				OrderID=RS("FolderOrder")
				Root=RS("Root")
				ClassType=RS("ClassType")
				Folder = RS("folder")
				Folder = Left(Folder, Len(Folder) - 1)
			 
				 If ClassType="1" Then
						If KS.C_S(C_ID,8) = "/" Or KS.C_S(C_ID,8) = "\" Then
						  CurrPath = KS.Setting(3) & Folder
						Else
						  CurrPath = KS.Setting(3) & KS.C_S(C_ID,8) & Folder
						End If
					If KS.FoundInarr("ask,club,space,job,item,config,admin,ks_cls,ks_inc,api,plus,editor",lcase(CurrPath),",")=false  Then
					  If (KS.DeleteFolder(CurrPath) = False) Then
					  ' Call KS.Alert("Delete Folder Error!", "KS.Class.asp")
					  ' Exit Sub
					  End If
				   End If
				 End IF
			  IF KS.ChkClng(KS.C_S(C_id,6))<=8 Then
				  conn.Execute ("Delete From KS_ItemInfoR Where (ChannelID=" & C_ID &" and InfoID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "')) Or (RelativeChannelID=" & C_ID & " And RelativeID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "'))")
				  conn.Execute ("Delete From KS_Comment Where ChannelID=" & C_ID &" and InfoID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "')")
				  conn.Execute ("Delete From KS_SpecialR Where ChannelID=" & C_ID &" and InfoID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "')")
				  conn.Execute ("Delete From KS_Digg Where ChannelID=" & C_ID &" and InfoID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "')")
				  conn.Execute ("Delete From KS_DiggList Where ChannelID=" & C_ID &" and InfoID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "')")
				  '删除栏目下信息的关联上传文件
				  conn.Execute ("Delete From KS_UploadFiles Where ChannelID=" & C_ID &" and InfoID in(select id from " & KS.C_S(C_ID,2) & " where tid='" & ID & "')")
				 '删除栏目的关联上传文件
				  Conn.Execute("Delete From [KS_UploadFiles] Where ChannelID=1000 and infoid=" & RS("ClassID"))
				  conn.Execute ("Delete From " & KS.C_S(C_ID,2) &" Where tid='" & ID & "'")
				  conn.Execute ("Delete From KS_ItemInfo Where ChannelID=" & C_ID &" and Tid='" & ID & "'")
				  End If
			 
			 if (Depth > 1) Then 
			  Conn.Execute("Update ks_Class set Child=Child-1 where ID='" & ParentID & "'")
		      Conn.Execute ("update ks_class set FolderOrder=FolderOrder-1 where FolderOrder>" & OrderID & " and root=" & Root)
			 End If

             '从缓存中去除
			 dim childNode:set childNode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & ID & "']")
			 If Not childNode Is Nothing Then childNode.parentNode.removeChild(childNode)
			 RS.MoveNext
            Loop
			 RS.Close
			 Conn.Execute("delete from KS_Class where ts Like '%" & RSC(0) & ",%'")
		RSC.MoveNext
	 Loop	 
			 
			 Set RS = Nothing
			 Dim Url:Url=REQUEST.ServerVariables("HTTP_REFERER")
			 if instr(lcase(url),"index.asp")<>0 then url="system/KS.Class.asp"
			 If Instr(Url,"?")=0 Then Url=Url&"?oids=" & request("oids") Else Url=Url&"&oids=" & request("oids")
			 ks.die "<script>top.$.dialog.alert('恭喜，删除成功!',function(){ location.href='" & Url &"';});</script>"
End Sub

Sub DelInfo()
			  Dim K, CurrPath, ArticleDir, FolderID,C_Id
			  Dim PageArr, TotalPage, I, CurrPathAndName, FExt, Fname
			  Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			  C_ID=KS.ChkClng(KS.S("ChannelID"))
			   RS.Open "Select * FROM " & KS.C_S(C_ID,2) &" Where Tid In (" & KS.GetFolderTid(KS.G("FolderID")) & ")", conn, 1, 1
			  Do While Not RS.Eof
				  '删除评论
				  conn.Execute ("Delete From KS_Comment Where ChannelID=" & C_ID &" and InfoID=" & RS("ID"))

				  FolderID = Trim(RS("Tid"))
                  on error resume next
				  FExt = Mid(Trim(RS("Fname")), InStrRev(Trim(RS("Fname")), ".")) '分离出扩展名
				  Fname = Replace(Trim(RS("Fname")), FExt, "")                    '分离出文件名 如 2005/9-10/1254ddd
				  
				 '删除物理文件
				  Dim FolderRS:Set FolderRS=Server.CreateObject("ADODB.Recordset")
				  FolderRS.Open "Select Folder From KS_Class WHERE ID='" & FolderID & "'", conn, 1, 1
				  CurrPath = Replace(KS.Setting(3) & KS.C_S(C_ID,8) & FolderRS("Folder"),"//","/")
				  If KS.C_S(C_ID,6)=1 Then
					  PageArr = Split(RS("ArticleContent"), "[NextPage]")
				  ElseIf  KS.C_S(C_ID,6)=2 Then
				      PageArr = Split(RS("PicUrls"), "|||")
				  End If
					  TotalPage = UBound(PageArr) + 1
					  If TotalPage > 1 Then
						For I = LBound(PageArr) To UBound(PageArr)
						 If I = 0 Then
						  CurrPathAndName = CurrPath & RS("Fname")
						 Else
						  CurrPathAndName = CurrPath & Fname & "_" & (I + 1) & FExt
						 End If
						 Call KS.DeleteFile(CurrPathAndName)
						Next
					  Else
					   CurrPathAndName = CurrPath & RS("Fname")
					   Call KS.DeleteFile(CurrPathAndName)
					  End If
				  FolderRS.Close
			  RS.MoveNext
             Loop
 			Set RS = Nothing
			conn.Execute ("Delete From " & KS.C_S(C_ID,2) &" Where Tid In (" & KS.GetFolderTid(KS.G("FolderID")) & ")")
			conn.Execute ("Delete From KS_ItemInfo Where Tid='" &  KS.G("FolderID") & "'")
			KS.Echo "<script>location.href='KS.Class.asp'</script>"
End Sub

'移动栏目
Sub MoveClass()
     Dim ID:ID=KS.G("ID")
	 %>
	 <!DOCTYPE html>
	<html>
	<link href="../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<script language="JavaScript" src="../../KS_Inc/Jquery.js"></script>
	<script language="JavaScript" src="../../KS_Inc/common.js"></script>
	<script>
	function ConfirmMovie(){
	  if (confirm('此操作不可逆，确认执行移动操作吗？')){
	  return true;
	  }else{
	   return false;
	  }
	}
	</script>
	<head>
	<body>
	 <%
 With KS

	.echo "<form action='KS.Class.asp?action=MoveSave&ID=" & ID &"' name='myform' method='post'>"
	.echo " <table border='0' cellpadding='3' cellspacing='1'  width='100%' align='center'>"
	.echo " <tr class='sort'>"
	.echo " <td style='text-align:left'> 将"& MyFolderName&"[" & KS.C_C(ID,1) &"]移动到 </td>"
	.echo "</tr>" & vbNewLine

	
	.echo " <tr class='tdbg'>"
	.echo " <td><select name='FolderID2' id='FolderID2'><option value='0'>无（作为频道)</option>" & KS.LoadClassOption(0,false) & "</select></td>"
	.echo "</tr>" & vbNewLine
	.echo " <tr class='sort'>"
	.echo "<td style='text-align:left'> <input type='submit' onclick='return(ConfirmMovie())' class='button' value='确定移动'>&nbsp;&nbsp;<input type='button' onclick='javascript:top.box.close();' class='button' value='取消返加'></td>"
	.echo "</tr>" & vbNewLine
    .echo "</table>"
	.echo "</form>"
	.echo "<br/><Br/><div class='attention'><strong>注意事项：</strong><br>" & _
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本操作不可逆，请慎重操作！！！<br>" & _
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;不能在同一个"& MyFolderName&"内进行操作，不能将一个"& MyFolderName&"移动到其下属"& MyFolderName&"中。<br>" & _
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;移动后您所指定的"& MyFolderName&"（或者包括其下属"& MyFolderName&"）将转移到目标"& MyFolderName&"中。</div>"
  End With
End Sub

'保存栏目移动
Sub MoveSave()
  Dim ID:ID=KS.S("ID")
  Dim ID2:ID2=KS.S("FolderID2")
  If KS.IsNul(ID) Or KS.IsNul(ID2) Then
     Call KS.AlertDoFun("参数出错！","history.back()")
  End If
  If ID=ID2 Then
     Call KS.AlertDoFun("不能在同一个"& MyFolderName&"里移动！","history.back()")
  End If
  
  Dim SQLStr
  If DataBaseType=1 Then
    SqlStr="select top 1 id From KS_Class Where ID='" & ID2 & "' and ','+convert(nvarchar(4000),ts)+',' LIKE '%," & id &",%'"
  Else
    SQLStr="select top 1 id From KS_Class Where ID='" & ID2 & "' and ','+TS+',' LIKE '%," & id &",%'"
  End If
  if Not Conn.Execute(SqlStr).Eof Then
     Call KS.AlertDoFun("对不起，不能移动到子"& MyFolderName&"下，请重新选择！","history.back()")
  End If
  
  If KS.C_C(ID,13)=ID2 Then
     Call KS.AlertDoFun("对不起，你没有选择要移动的位置，请重新选择！","history.back()")
  End If
  
  If DataBaseType=1 Then
    SqlStr="select count(1) From  KS_Class Where ','+convert(nvarchar(4000),ts)+',' LIKE '%," & id &",%'"
  Else
    SQLStr="select count(1) From  KS_Class Where ','+TS+',' LIKE '%," & id &",%'"
  End If
  Dim Num:Num=Conn.Execute(SQLStr)(0)
  Dim FolderOrder,root

  '改变源栏目剩下的栏目序号
  Root=Conn.Execute("select top 1 root From KS_Class Where ID='" & ID & "'")(0)
  if Conn.Execute("select top 1 id From KS_Class Where tn='" & ID & "'").eof then  '没有子栏目
   FolderOrder=KS.ChkClng(Conn.Execute("select top 1 FolderOrder From KS_Class Where ID='" & ID & "'")(0))
  else
   FolderOrder=KS.ChkClng(Conn.Execute("select max(FolderOrder) From KS_Class Where tn='" & ID & "'")(0))
  end if
  If  FolderOrder>0 Then
   Conn.Execute("Update KS_Class Set FolderOrder=FolderOrder-" & Num &" Where Root=" & Root &" and FolderOrder>" & FolderOrder&"")  '更新目标栏目大于当前栏目的排序号
  End If


  
  '改变目标栏目的序号
  If ID2<>"0" Then
	  Root=Conn.Execute("select top 1 root From KS_Class Where ID='" & ID2 & "'")(0)
	  if Conn.Execute("select top 1 id From KS_Class Where tn='" & ID2 & "'").eof then  '没有子栏目
	   FolderOrder=KS.ChkClng(Conn.Execute("select top 1 FolderOrder From KS_Class Where ID='" & ID2 & "'")(0))
	  else
	   FolderOrder=KS.ChkClng(Conn.Execute("select max(FolderOrder) From KS_Class Where tn='" & ID2 & "'")(0))
	  end if
	  if  FolderOrder>0 Then
	  Conn.Execute("Update KS_Class Set FolderOrder=FolderOrder+" & Num &" Where Root=" & Root &" and FolderOrder>" & FolderOrder&"")  '更新目标栏目大于当前栏目的排序号
	  End If
  Else 
      FolderOrder=0
  End If

 '改变当前栏目
  Call DoMoveClass(ID,ID2,FolderOrder)
  
  '改变子栏目
  Call DoMoveLoop(ID,FolderOrder)

   Application(KS.SiteSN&"_class")=""
   Application(KS.SiteSN&"_classpath")=""
   Call KS.AlertDoFun("恭喜"& MyFolderName&"移动成功！","top.frames['MainFrame'].location.reload();top.box.close()")
End Sub

Sub DoMoveLoop(ParentID,FolderOrder)
  '改变子栏目
  Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
  RSS.Open "select * From KS_Class Where Tn='" & ParentID & "'",conn,1,1
  Do While Not RSS.Eof
	 Call DoMoveClass(RSS("ID"),RSS("TN"),FolderOrder)
	 Call DoMoveLoop(RSS("id"),FolderOrder)
  RSS.MoveNext
  Loop
  RSS.Close
  Set RSS=Nothing
End Sub

'执行移动栏目操作
'参数 ID 当前栏目 ID2目标栏目
Sub DoMoveClass(ID,ID2,FolderOrder)
  Dim TS,Folder,TJ,Root
  If ID2<>"0" Then
	  Dim RS:Set RS=Conn.Execute("select top 1 * From KS_Class Where ID='" & id2 &"'")
	  If RS.Eof And RS.Bof Then
		 RS.Close
		 Set RS=Nothing
		 Call KS.AlertDoFun("对不起，找不到目标"& MyFolderName&"！","history.back()")
	  End If
	  TS=RS("Ts")
	  Folder=RS("Folder")
	  TJ=KS.ChkClng(RS("TJ"))
	  Root=KS.ChkClng(RS("Root"))
	  RS.Close
      Set RS=Nothing
      FolderOrder=FolderOrder+1
 Else
      Root=KS.ChkClng(Conn.Execute("select max(root) From KS_Class")(0))+1
	  Tj=0
	  Folder=""
	  TS=""
	  FolderOrder=0
 End If

 
  Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
  RSS.Open "select top 1 * From KS_Class Where ID='" & ID & "'",conn,1,3
	  RSS("TN")=ID2
	  RSS("TJ")=TJ+1
	  RSS("FolderOrder")=FolderOrder
	  Dim NewFolder
	  if rss("classType")=1 Then
	   NewFolder=Folder & Split(RSS("Folder"),"/")(Ubound(Split(RSS("Folder"),"/"))-1) &"/"
	  ElseiF rss("ClassType")=2 Then
	   NewFolder=RSS("Folder")
	  Else
	     if instr(RSS("Folder"),"/")=0 then
	       NewFolder=Folder & RSS("Folder")
		 else
	       NewFolder=Folder & Split(RSS("Folder"),"/")(Ubound(Split(RSS("Folder"),"/")))
		 End If
	  End If
	  RSS("Folder")=NewFolder
	  Dim NewTS:NewTS=TS & RSS("id")&","
	  RSS("TS")=NewTS
	  RSS("Root")=Root
  RSS.Update
  RSS.Close
  
End Sub




		
'一键更新栏目排序		
Sub DoOrder()
		 With KS  

				.echo "<div class='pageCont2'><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
				.echo "<tr> "
				.echo "<td bgcolor=000000>"
				.echo " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				.echo "<tr> "
				.echo "<td bgcolor=ffffff height=9><img src=""../images/114_r2_c2.jpg"" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table>"
				.echo "</td></tr></table>"
				.echo "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				.echo "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				.echo "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
				.echo "</table>"
			
			 .echo ("<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
			 .echo (" <tr>")
			 .echo ("   <td height=""50"">")
			 .echo ("     <div align=""center""> ")
			 .echo ("       </div></td>")
			 .echo ("   </tr>")
			 .echo ("</table>")
			 .echo ("</div>")
			.echo ("</div>")


       Dim StartRefreshTime:StartRefreshTime = Timer()
       Dim root:root = 0
       TotalNum = Conn.Execute("select count(1) From KS_Class")(0)
	   Dim NowNum:NowNum=0
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select * from ks_class where tn='0' order by root,classid",conn,1,1
			If RS.Eof And RS.Bof  Then
			     RS.Close
				 Set RS=Nothing
                 Call KS.AlertDoFun("对不起，你没有添加任何"& MyFolderName&"！","history.back()")
			ELSE
				Do While Not RS.Eof 
				  Dim I:I=0
				  Root=Root+1
				  NowNum=NowNum+1
				  Dim Child:Child=Conn.Execute("select count(1) From KS_Class Where Tn='" & RS("ID") &"'")(0) 
				  Conn.Execute("Update KS_Class Set TJ=1,Root=" & Root & ",FolderOrder=" & I & ",Child=" & Child &" Where ClassID=" & RS("ClassID"))
				  Call UpdateChild(rs("id"), root, I, 1,NowNum)
				  Call InnerJS(NowNum,TotalNum,MyFolderName)
				RS.MoveNext
				Loop
			End If
				RS.Close
				Set RS=Nothing
				
				
		Application(KS.SiteSN&"_class")=""
                .echo "<script>"
				.echo "img2.width=400;" & vbCrLf
				.echo "txt2.innerHTML=""更新"& MyFolderName&"排序结束！100"";" & vbCrLf
				.echo "txt3.innerHTML=""总共更新了 <font color=red><b>" & TotalNum & "</b></font> 个,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br><br><input name='button1' type='button' onclick=\""javascript:location.href='KS.Class.asp'; $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('<font color=red>" & MyFolderName&"管理</font>')+'&ButtonSymbol=Disabled';\"" class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				.die ""
	End With
End Sub

Sub UpdateChild(ID, root, orderid,byval depth,NowNum)
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "select * From KS_Class Where Tn='" & ID & "' order by FolderOrder,classid",conn,1,1
  depth=depth+1
  Do While Not RS.Eof
    NowNum=NowNum+1
    OrderID=OrderID+1
	Dim Child:Child=Conn.Execute("select count(1) From KS_Class Where Tn='" & RS("ID") &"'")(0) 
    Conn.Execute("Update KS_Class Set Root=" & Root & ",TJ=" & Depth &",FolderOrder=" &OrderID& ",Child=" & Child&" Where ClassID=" & RS("ClassID"))
	Call UpdateChild(rs("id"), root,orderid, depth,NowNum)
    Call InnerJS(NowNum,TotalNum,MyFolderName)
  RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
      
End Sub


Sub InnerJS(NowNum,TotalNum,itemname)
		  With KS
				.echo "<script>"
				.echo "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.echo "txt2.innerHTML=""生成进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.echo "txt3.innerHTML=""总共需要更新 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在更新第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				Response.Flush
		  End With
End Sub


Sub Unite()
    With KS
	 .echo "<script language='javascript'>" & vbcrlf
     .echo "$(document).ready(function(){" &vbcrlf
	 .echo " $('#channelids').change(function(){" &vbcrlf
	 .echo " if ($(this).val()!=0){" & vbcrlf
	 .echo "  $(parent.document).find(""#ajaxmsg"").toggle();" & vbcrlf
	 .echo "  $.get('../../plus/ajaxs.asp',{action:'GetClassOption',from:'label',channelid:$(this).val()},function(data){" & vbcrlf
	 .echo "    $(parent.document).find(""#ajaxmsg"").toggle();" & vbcrlf
	 .echo "    $('select[name=FolderID1]').empty().append(unescape(data));" & vbcrlf
	 .echo "    $('select[name=FolderID2]').empty().append(unescape(data));" & vbcrlf
	 .echo "      }" & vbcrlf
	 .echo "    );" & vbcrlf
	 .echo "  }" &vbcrlf
	 .echo " });" & vbcrlf
	 .echo "})"&vbcrlf
	 .echo "</script>"
	 .echo "<div class='pageCont2'>"
	.echo "<form action='KS.Class.asp?action=UniteSave' name='myform' method='post'>"
	.echo " <table border='0' cellpadding='3' cellspacing='1'  width='100%' align='center'>"
	.echo " <tr class='sort'>"
	.echo " <td>" & MyFolderName &"合并 </td>"
	.echo "</tr>" & vbNewLine
	.echo " <tr class='tdbg'>"
	.echo " <td height='40'><strong>选择模型</strong><select id='channelids' name='channelid'>"
	.echo " <option value='0'>---请选择模型---</option>"
	.LoadChannelOption 0
	
	.echo "</select></td>"
	.echo "</tr>" & vbNewLine
	
	.echo " <tr class='tdbg'>"
	.echo " <td height=150>"
	.echo "   <table border='0' cellspacing='0' cellpadding='0'><tr><td><strong>将"& MyFolderName&"</strong></td><td><select name='FolderID1' id='FolderID1' size='8' style='width:200px'>" & KS.LoadClassOption(1,false) & "</select></td>"
	.echo "   <td><strong>合并到</strong></td><td><select style='width:200px' name='FolderID2' id='FolderID2' size='8'>" & KS.LoadClassOption(1,false) & "</select></td></tr></table></td>"
	.echo "</tr>" & vbNewLine
	.echo " <tr class='sort'>"
	.echo "<td align='center'><input type='button' onclick='return(ConfirmUnite())' class='button' value='确定合并'>&nbsp;&nbsp;<input type='button' onclick='javascript:location.href=""KS.Class.asp"";' class='button' value='取消返回'></td>"
	.echo "</tr>" & vbNewLine
    .echo "</table>"
	.echo "</form>"
	.echo "<div class='attention'><strong>注意事项：</strong><br>" & _
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本操作不可逆，请慎重操作！！！<br>" & _
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;不能在同一个"& MyFolderName&"内进行操作，不能将一个"& MyFolderName&"合并到其下属"& MyFolderName&"中。<br>" & _
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;合并后您所指定的"& MyFolderName&"（或者包括其下属"& MyFolderName&"）将被删除，所有文档将转移到目标"& MyFolderName&"中。</div>"
	.echo "</form>"
  End With
End Sub

Sub UniteSave()
 Dim  CurrPath,ChannelID
 Dim FolderID1:FolderID1=KS.G("FolderID1")
 Dim FolderID2:FolderID2=KS.G("FolderID2")
 If FolderID1<> FolderID2 Then
   If Not Conn.Execute("Select ID From KS_Class Where TS Like '%" & FolderID1 & "%' And ID='" & FolderID2 & "'").Eof Then
     Call KS.AlertHintScript("不能将一个"& MyFolderName&"合并到其下属"& MyFolderName&"中！")
	 Exit Sub
   Else
    '得到当前栏目信息
	Dim ParentID,ParentPath,Depth
    Dim rsc:Set rsc = Conn.Execute("select ID,tn,ts,tj from KS_Class where ID='" & FolderID1 & "'")
    If rsc.BOF And rsc.EOF Then
        RSC.Close:Set RSC=Nothing
		KS.AlertHintScript "找不到指定的"& MyFolderName&"，可能已经被删除！"
        Exit Sub
    End If
    ParentID = rsc(1)
    ParentPath = rsc(2)
    Depth = rsc(3)
	RSC.Close:Set RSC=Nothing
   
   
     Dim RS:Set RS=Server.CreateObject("Adodb.recordset")
	 RS.Open "Select * From KS_Class Where TS Like '%" & FolderID1 & "%'",conn,1,3
	 Do While Not RS.Eof
	       ChannelID=RS("ChannelID")
           Conn.Execute("Update " & KS.C_S(ChannelID,2)  & " Set Tid='" & FolderID2 & "' Where Tid='" & FolderID1 &"'")
           Conn.Execute("Update [KS_ItemInfo] Set Tid='" & FolderID2 & "' Where Tid='" & FolderID1 &"'")
             If KS.C_S(RS("ChannelID"),8) = "/" Or KS.C_S(ChannelID,8) = "\" Then
			  CurrPath = KS.Setting(3) & RS("Folder")
			 Else
			  CurrPath = KS.Setting(3) & KS.C_S(ChannelID,8) & RS("Folder")
			 End If
			 If (KS.DeleteFolder(CurrPath) = False) Then
			   Call KS.AlertHintScript("Delete Folder Error!")
			   Exit Sub
			 End If
			  '从缓存中去除
			 dim childNode:set childNode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & RS("ID") & "']")
			 childNode.parentNode.removeChild(childNode)
	  RS.Delete
	  RS.MoveNext
	 Loop
	 RS.Close:Set RS=Nothing
	 '更新其原来所属栏目的子栏目数，排序相当于剪枝而不需考虑
    If ParentID <> "0" And Not KS.IsNul(ParentID) Then
        Conn.Execute ("update KS_Class set Child=Child-1 where ID='" & ParentID &"'")
    End If
	 
   End If
 End If
    Call KS.AlertHintScript("恭喜，网站"& MyFolderName&"已成功合并！")
End Sub

Sub MoveInfo()
 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
 If ChannelID=0 Then ChannelID=1
%>
 <script language="javascript">
  $(document).ready(function(){
   $("#channelids").change(function(){
     if ($(this).val()!=0){
	  $(parent.document).find("#ajaxmsg").toggle();
	  $.get("../../plus/ajaxs.asp",{action:"GetClassOption",from:"label",channelid:$(this).val()},function(data){
	     $(parent.document).find("#ajaxmsg").toggle();
	     $("select[name=BatchClassID]").empty();
		 $("select[name=BatchClassID]").append(unescape(data));
		 $("select[name=tClassID]").empty();
		 $("select[name=tClassID]").append(unescape(data));
		 $("input[name=ChannelID]").val($("#channelids").val());
	   });
	 }
   });
  })
function SelectAll(){
  $("select[name=BatchClassID]>option").each(function(){
   $(this).attr("selected",true);
  })
}
function UnSelectAll(){
  $("select[name=BatchClassID]>option").each(function(){
   $(this).attr("selected",false);
  })
}
 </script>
 <div class="pageCont2">
 <div class="tabTitle">批量移动信息</div>
 <table width='100%' border='0' align='center' cellpadding='1' cellspacing='1'>   
 <form method='POST' name='myform' action='KS.Class.asp?From=<%=KS.S("From")%>' target='_self'>
   <tr class='tdbg' <%if KS.S("From")="main" Then KS.Echo " style='display:none'"%>>
	 <td height='40' colspan='4'><strong>选择模型</strong>
	 <select id='channelids' name='channelids'>
	 <option value='0'>---请选择模型---</option>
	 <%
	KS.LoadChannelOption 0
	%>
	</select></td>
   </tr>   
   <tr align='left' class='tdbg'>      <td valign='top' width='350'>     
   <%if KS.S("From")="main" Then
     KS.Echo "<span style=''>"
	 Else
     KS.Echo "<span style='display:none'>"
	 End If
	 %><input type='radio' name='InfoType' value='1'<%if KS.S("From")="main" Then KS.Echo " checked"%>>指定信息ID：<input type='text' name='BatchInfoID' value='<%=Replace(KS.G("ID")," ","")%>' class='textbox' size='20'><br> </span>      
	  <input type='radio' name='InfoType' value='2' <%if KS.S("From")<>"main" Then KS.Echo " checked"%>>指定<%=MyFolderName%>的信息：<br>
	  <select name='BatchClassID' size='2' multiple style='height:300px;width:300px;'><%=KS.LoadClassOption(ChannelID,false)%></select><br>        <input type='button' name='Submit' value='选定所有' class='button' onclick='SelectAll()'>        <input type='button' class='button' name='Submit' value='取消所选' onclick='UnSelectAll()'>      </td>      <td align='center' >移动到&gt;&gt;</td>      <td valign='top'>      <div>目标<%=MyFolderName%></div><select name='tClassID' size='2' style='height:360px;width:300px;'><%=KS.LoadClassOption(ChannelID,false)%></select>      </td>    </tr>  </table>  <p align='center'>  
	  <input name='Action' type='hidden' id='Action' value='MoveToClass'>    
	  <input name='ChannelID' type='hidden' id='ChannelID' value='<%=ChannelID%>'>
	  <input name='add' type='submit'  class='button' id='Add' value=' 执行批处理 ' style='cursor:pointer;' onClick="$('input[name=Action]').val('DoMoveToClass');">&nbsp;    
	  <%if KS.S("From")="main" Then%>
	  <input name='Cancel' type='button' id='Cancel' value=' 取消关闭 ' onClick="top.box.close();" class='button' style='cursor:pointer;'>
	  <%else%>
	  <input name='Cancel' type='button' id='Cancel' value=' 取消返回 ' onClick="history.back();" class='button' style='cursor:pointer;'>
	  <%end if%>
	    </p></form>
        </div>
<%
End Sub
Sub DoMoveToClass()
		 Dim BatchClassID:BatchClassID=Replace(KS.G("BatchClassID")," ","")
		 Dim tClassID:tClassID=KS.G("tClassID")
		 if TclassID="" Then KS.AlertHintScript "请选择目标"& MyFolderName&"!"
		 Dim InfoType:InfoType=Replace(KS.G("InfoType")," ","")
		 Dim BatchInfoID:BatchInfoID=KS.G("BatchInfoID")
		 Dim ChannelID,FolderTidList,I
		 ChannelID=KS.ChkClng(KS.G("ChannelID"))
		 If InfoType=1 Then
		   If KS.FilterIDs(BatchInfoID)="" Then 
		    KS.AlertHintScript "请输入要移动的文档ID列表!"
			Response.End()
		   Else
		   Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Tid='" & tClassID & "' where ID In(" & KS.FilterIDs(BatchInfoID) &")")
		   Conn.Execute("Update KS_ItemInfo Set Tid='" & tClassID & "' where ChannelID=" & ChannelID & " and InfoID In(" & KS.FilterIDs(BatchInfoID) &")")
		   End If
		 Else
		   
		   If BatchClassID="" Then 
		    KS.AlertHintScript "请选择要移动"& MyFolderName&"!"
			Response.End()
		   End if
		   BatchClassID=Split(BatchClassID,",")
		   For i=0 To Ubound(BatchClassID)
		     If FolderTidList="" Then
			 FolderTidList=GetFolderTid(BatchClassID(i))
			 Else
		     FolderTidList=FolderTidList &","&GetFolderTid(BatchClassID(i))
			 End If
		   Next
		   Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Tid='" & tClassID & "' Where Tid In (" & FolderTidList &")") 
		   Conn.Execute("Update KS_ItemInfo Set Tid='" & tClassID & "' Where Tid In (" & FolderTidList &")") 
		 End IF
		 If KS.S("From")="" Then
		  KS.AlertHintScript "恭喜,文档批量移动成功!"
		 Else
		  KS.Echo ("<script>alert('恭喜，成功批量移动指定的文档到目标"& MyFolderName&"!');top.box.close();</script>")
		 End If
End Sub

Function GetFolderTid(FolderID)
			Dim I,Tid,SQL
			Dim RS:Set RS=Conn.Execute("Select ID From KS_Class Where DelTF=0 AND TS LIKE '%" & FolderID & "%'")
			 If RS.EOF Then	 GetFolderTid="'0'":RS.Close:Set RS=Nothing:Exit Function
			 SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
             For I=0 To Ubound(SQL,2)
				  Tid = Tid & "'" & Trim(SQL(0,I)) & "',"
			 Next
			Tid = Left(Trim(Tid), Len(Trim(Tid)) - 1) '去掉最后一个逗号
			GetFolderTid = Tid
End Function


Sub Attribute()
    Dim ChannelID:ChannelID=1
    Dim KSCls:Set KSCls=New ManageCls
	KS.Echo "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
	KS.Echo "<script src=""../../KS_Inc/Jquery.js"" language=""JavaScript""></script>"
	KS.Echo "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>"
	KS.Echo "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
	Dim NowDate:NowDate = Now()
	Dim YearStr:YearStr = CStr(Year(NowDate))
	Dim MonthStr:MonthStr = CStr(Month(NowDate))
	Dim DayStr:DayStr = CStr(Day(NowDate))


%> 
<script language="javascript">
  $(document).ready(function(){
   $("#channelids").change(function(){
     if ($(this).val()!=0){
	  $(parent.document).find("#ajaxmsg").toggle();
	  $.get("../../plus/ajaxs.asp",{action:"GetClassOption",from:'label',channelid:$(this).val()},function(data){
	     $(parent.document).find("#ajaxmsg").toggle();
	     $("select[name=ClassID]").empty().append(unescape(data));
		 $("input[name=ChannelID]").val($("#channelids").val());
	   });
	 }
   });
  })
  function SelectAll(){
   $("#ClassID>option").each(function(){
     $(this).attr("selected",true);
   }) 
  }
 function UnSelectAll(){
   $("#ClassID>option").each(function(){
     $(this).attr("selected",false);
    }) 
}

</script> 
<div class="pageCont2">
  <FORM name=form1 action=KS.Class.asp method=post>
<table cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
    <tr class=sort>
      <td align=middle colSpan=3 height=22><strong>批量设置<%=MyFolderName%>属性</strong></td>
    </tr>
    <tr class=tdbg>
      <td vAlign=top width=200><font color=red>提示：</font>可以按住“Shift”<br />或“Ctrl”键进行多个<%=MyFolderName%>的选择<br />

<select id='channelids' name='channelids'>
	 <option value='0'>---请选择模型---</option>
	 <%
	 KS.LoadChannelOption 0
	%>
	</select>
<Select style="WIDTH: 200px; HEIGHT: 380px" multiple size=2 name="ClassID" id="ClassID">

 <%=KS.LoadClassOption(ChannelID,false)%>
</Select>
<div align=center>
   <Input onclick=SelectAll() type=button class="button" value="选定所有<%=MyFolderName%>" name=Submit><br />
   <Input onclick=UnSelectAll() type=button value="取消选定<%=MyFolderName%>" class="button" name=Submit></div>
   </td>
      <td vAlign=top><br />
		<div class=tab-page id=ClassAttrPane>
		<SCRIPT type=text/javascript>
			   var tabPane1 = new WebFXTabPane( document.getElementById( "ClassAttrPane" ), 1 )
		</SCRIPT>
				 
			<div class=tab-page id=site-page1>
			<H2 class=tab><%=MyFolderName%>选项</H2>
				<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "site-page1" ) );
				</SCRIPT>
                  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyTopFlag'></td>
					<td height='30' width='200' align='right' class='clefttitle'><strong><%=MyFolderName%>顶部导航：</strong></td>
					<td height='28'>&nbsp;<input name="TopFlag" type="radio" value="1" checked>显示 <input name="TopFlag" type="radio" value="0">不显示              
				   </td>          
				  </tr>
				 <%If KS.WSetting(0)="1" Then%>
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyWapSwitch'></td>
					<td height='30' width='200' align='right' class='clefttitle'><strong><%=MyFolderName%>3G状态：</strong></td>
					<td height='28'>&nbsp;<input name="WapSwitch" type="radio" value="1" checked>显示 <input name="WapSwitch" type="radio" value="0">不显示              
				   </td>          
				  </tr>   
				 <%End If%> 
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyMailTF'></td>
					<td height='30' width='200' align='right' class='clefttitle'><strong><%=MyFolderName%>允许邮件订阅：</strong></td>
					<td height='28'>&nbsp;<input name="MailTF" type="radio" value="1" checked>允许 <input name="MailTF" type="radio" value="0">不允许              
				   </td>          
				  </tr>    
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyChannelTemplateID'></td>
					<td height='30' align='right'  width='200' class='clefttitle'><strong>
		频道模板：</strong> </td>
					<td height='28'><b>
					  <input type="text" class="textbox" name='ChannelTemplateID' id='ChannelTemplateID' size="30">&nbsp;<%=KSCls.Get_KS_T_C("$('#ChannelTemplateID')[0]")%></select>
				    </td>
				 </tr>
				 <%If KS.WSetting(0)="1" Then%>
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='WAPModifyChannelTemplateID'></td>
					<td height='30' align='right'  width='200' class='clefttitle'><strong>
		3G频道模板：</strong> </td>
					<td height='28'><b>
					  <input type="text" class="textbox" name='WAPChannelTemplateID' id='WAPChannelTemplateID' size="30">&nbsp;<%=KSCls.Get_KS_T_C("$('#WAPChannelTemplateID')[0]")%></select>
				    </td>
				 </tr>
				 <%End If%>
				 
				 
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyFolderTemplateID'></td>
					<td height='30' align='right'  width='200' class='clefttitle'><strong>
		<%=MyFolderName%>模板：</strong> </td>
					<td height='28'><b>
					  <input type="text" class="textbox" name='FolderTemplateID' id='FolderTemplateID' size='30'>&nbsp;<%=KSCls.Get_KS_T_C("$('#FolderTemplateID')[0]")%></select>
				    </td>
				 </tr>
				 <%If KS.WSetting(0)="1" Then%>
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='WAPModifyFolderTemplateID'></td>
					<td height='30' align='right'  width='200' class='clefttitle'><strong>
		3G<%=MyFolderName%>模板：</strong> </td>
					<td height='28'><b>
					  <input type="text" class="textbox" name='WAPFolderTemplateID' id='WAPFolderTemplateID' size='30'>&nbsp;<%=KSCls.Get_KS_T_C("$('#WAPFolderTemplateID')[0]")%></select>
				    </td>
				 </tr>
				 <%End If%>
				 
				 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				 <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyFolderFsoIndex'></td>
				 <td align=right  class='clefttitle' width='200'><strong>
					 生成的<%=MyFolderName%>首页文件：</strong>
		</td><td>      <select name='FolderFsoIndex' id='select2' class='textbox'>
					   <option value='index.html'>index.html</option>
					   <option value='index.htm' selected>index.htm</option>
					   <option value='index.shtm'>index.shtm</option>
					   <option value='index.shtml'>index.shtml</option>
					   <option value='default.html'>default.html</option>
					   <option value='default.htm'>default.htm</option>
					   <option value='default.shtm'>default.shtm</option>
					   <option value='default.shtml'>default.shtml</option>
					   <option value='index.asp'>index.asp</option>
					   <option value='default.asp'>index.asp</option>
					   <option value="index.html" selected>index.html</option>             </select>
					 </td>
				 </tr>
				 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				   <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyTemplateID'></td>
				   <td height='30' align='right'  class='clefttitle' width='200'><strong>内容页模板：</strong></td>
				   <td height='28'>
					  <input type="text" class="textbox" name='TemplateID' id='TemplateID' size='30'>&nbsp;<%=KSCls.Get_KS_T_C("$('#TemplateID')[0]")%></select>    </td></tr>   
			<%If KS.WSetting(0)="1" Then%>
				 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				   <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='WAPModifyTemplateID'></td>
				   <td height='30' align='right'  class='clefttitle' width='200'><strong>3G内容页模板：</strong></td>
				   <td height='28'>
					  <input type="text" class="textbox" name='WAPTemplateID' id='WAPTemplateID' size='30'>&nbsp;<%=KSCls.Get_KS_T_C("$('#WAPTemplateID')[0]")%></select>    </td></tr>   
			<%End If%>
					  
					       
		     <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
			     <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyFnameType'></td>
			     <td height='28'   width='200' align=right class='clefttitle'><strong>生成的页扩展名：</strong>         </td><td>             <input type='text' ID='FnameType' name='FnameType' value='.html' size='15' class="textbox"/> <-<select name='FnameTypes'  class='upfile' onChange="$('#FnameType').val(this.value);">
               <option value='.html' selected>.html</option>
               <option value='.htm'>.htm</option>
               <option value='.shtm'>.shtm</option>
               <option value='.shtml'>.shtml</option>
               <option value='.asp'>.asp</option>
             </select>
					</td>
				</tr>
				<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">          
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyFsoType'></td>
				  <td height='30' align='right' width='200' class='clefttitle'><strong>生成路径格式：</strong></td>
				  <td height='28'> <select style='width:200;' name='FsoType' id='select5' onChange='SelectFsoType(options[selectedIndex].value);'>
					<option value="1"><%=YearStr%>/<%=MonthStr%>-<%=DayStr%>/RE</option>
					<option value="2"><%=YearStr%>/<%=MonthStr%>/<%=DayStr%>/RE</option>
					<option value="3"><%=YearStr%>-<%=MonthStr%>-<%=DayStr%>/RE</option>
					<option value="4"><%=YearStr%>/<%=MonthStr%>/RE</option>
					<option value="5"><%=YearStr%>-<%=MonthStr%>/RE</option>
					<option value="6"><%=YearStr%><%=MonthStr%><%=DayStr%>/RE</option>
					<option value="7"><%=YearStr%>/RE</option>
					<option value="8"><%=YearStr%><%=MonthStr%><%=DayStr%>RE</option>
					<Option value="9" Selected>RE</Option>
					<option value="10">SCE</option><option value="11">新闻IDE</option>            
		         </select> </td>        
				 </tr>
				 
<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">          
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyAd'></td>
				  <td height='30' align='right' width='200' class='clefttitle'><strong>设置画中画：</strong></td>
				  <td height='28'> 
				  <input onClick="$('#Ad').show();" name="ShowADTF" type="radio" value="1">显示 <input onClick="$('#Ad').hide();" name="ShowADTF" type="radio" value="0" checked>不显示 <table  style="display:none" id="Ad" class="border" style="margin:5px" border="0" cellpadding="5" cellspacing="1"><tr class="tdbg"><td width="22%"><div align="right">画中画参数设置：</div></td><td width="78%"><input class="textbox" name="AdParam" type="text" id="AdParam" size="20" maxlength="20" value="250,left,300,300">(插入位置在内容前多少字,左(left)右(right),宽度,高度：如500,left,300,300)</td></tr><tr class="tdbg"><td><div align="right">广告类型：</div></td><td><input onClick="$('#adcodearea').hide();$('#adimgarea').show();" name="ADType" type="radio" value="1" checked>图片/Flash <input onClick="$('#adimgarea').hide();$('#adcodearea').show();" name="ADType" type="radio" value="2">代码广告（支持Google广告)</td></tr><tbody id='adcodearea' style='display:none'><tr class="tdbg"><td><div align="right">广告代码：<br><font color=red>支持HTML语法</font></div></td><td><textarea style='height:60px' name="AdCode" class="textbox" cols='60' rows=6></textarea></td></tr></tbody><tbody id='adimgarea'><tr class="tdbg"><td><div align="right">图片地址：</div></td><td><input name="AdUrl" class="textbox" type="text" id="AdUrl"  size="36" maxlength="250" value=""> <input class="button"  type='button' name='Submit' value='选择图片或FLASH' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=1&CurrPath=<%=KS.GetUpfilesDir()%>',550,290,window,$('#AdUrl')[0]);"> </td></tr><tr class="tdbg"><td><div align="right">链接地址：</div></td><td><input name="AdLinkUrl" type="text" class="textbox" id="AdLinkUrl"  size="36" maxlength="250" value="http://www.kesion.com">仅对图片有效</td></tr></tbody> </table>
				  
				   </td>        
				 </tr>				 
				 
				
				 </table>
				</div>
		 
		<div class=tab-page id=site-page>
			<H2 class=tab>权限选项</H2>
				<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "site-page" ) );
				</SCRIPT>
              <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
                <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyClassPurview'></td>
                  <td  class='clefttitle' width=200><strong>浏览/查看权限：</strong></td>
                  <td>
                    &nbsp;<input name='ClassPurview' type='radio' value='0' checked>              开放<%=MyFolderName%>&nbsp;&nbsp;<font color=red>任何人（包括游客）可以浏览和查看此<%=MyFolderName%>下的信息。</font><br>              &nbsp;<INPUT type='radio'  name='ClassPurview' value='1'>
              半开放<%=MyFolderName%>&nbsp;&nbsp;<font color=red>任何人（包括游客）都可以浏览。游客不可查看，其他会员根据会员组的<%=MyFolderName%>权限设置决定是否可以查看。</font><br/>              &nbsp;<INPUT type='radio'  name='ClassPurview' value='2'>
              认证<%=MyFolderName%>&nbsp;&nbsp;<font color=red>游客不能浏览和查看，其他会员根据会员组的<%=MyFolderName%>权限设置决定是否可以浏览和查看。</font>
                  </td>
                </tr>
                <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyGroupID'></td>
                  <td class='clefttitle' width=200><div><strong>允许查看此<%=MyFolderName%>下信息的会员组：</strong></div><font color=blue>如果<%=MyFolderName%>是“认证<%=MyFolderName%>”，请在此设置允许查看此<%=MyFolderName%>下信息的会员组,如果在信息中设置了查看权限，则以信息中的权限设置优先</font></td>
                  <td><%=KS.GetUserGroup_CheckBox("GroupID","",3)%></td>
                </tr>
                
                <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyReadPoint'></td>
                  <td  class='clefttitle' width=200><strong>默认阅读信息所需点数：</strong><br><font color=blue>如果在信息中设置了阅读点数，则以信息中的点数设置优先</font></td>
                  <td>&nbsp;<input name='ReadPoint' type='text' id='ReadPoint'  value='0' size='6' class='textbox' style='text-align:center'> 　免费阅读请设为 "<font color=red>0</font>"，否则有权限的会员阅读该<%=MyFolderName%>下的信息时将消耗相应点数，游客将无法阅读。</td>
                </tr>
				 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				 	 <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyDividePercent'></td>

            <td height='60' align='center' class='clefttitle'><strong>默认与投稿者的分成比率：</strong></td>
            <td height='28'>&nbsp;<input name='DividePercent' class="textbox" type='text' value='0' size='6' class='upfile' style='text-align:center'>% 系统将根据这里设置的分成比率将收成分给投稿者。建议设成10的整数倍!</td>          </tr>
                <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyChargeType'></td>
                  <td  class='clefttitle' width=200><strong>默认阅读信息重复收费：</strong><br><font color=blue>如果在信息中设置了阅读点数，则以信息中的点数设置优先</font></td>
                  <td>&nbsp;<input name='ChargeType' type='radio' value='0'  checked >不重复收费(如果信息需扣点数才能查看，建议使用)<br>&nbsp;<input name='ChargeType' type='radio' value='1'>距离上次收费时间 <input name='PitchTime' type='text' class='textbox' value='12' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='2'>会员重复阅信息 &nbsp;<input name='ReadTimes' type='text' class='textbox' value='10' size='8' maxlength='8' style='text-align:center'> 页次后重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='3'>上述两者都满足时重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='4'>上述两者任一个满足时就重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='5'>每阅读一页次就重复收费一次（建议不要使用,多页信息将扣多次点数）</td>
                </tr>
				</table>
			</div>

            <div class=tab-page id=tg-page>
			<H2 class=tab>投稿选项</H2>
				<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "tg-page" ) );
				</SCRIPT>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
				  
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyPubTF'></td>
					<td height='30' align='right' width='200' class='clefttitle'><strong>允许在本<%=MyFolderName%>发布文档：</strong><br><font color=blue>当<%=MyFolderName%>不是终级<%=MyFolderName%>时,建议选择不允许</font></td>
					<td height='28'<input name="PubTF" type="radio" value="1" checked><input name="PubTF" type="radio" value="1" checked>允许<input name="PubTF" type="radio" value="0">不允许
					
					 </td>          
				 </tr>
				  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				    <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyCommentTF'></td>
					<td height='30' align='right' width='200' class='clefttitle'><strong><%=MyFolderName%>是否允许投稿：</strong></td>
					<td height='28'>①<input name="CommentTF" type="radio" value="0">不允许<br>②<input name="CommentTF" type="radio" value="1" checked>允许<br> ③<input name="CommentTF" type="radio" value="2" checked>允许所有人投稿<font color=red>(包括游客)</font><br>④<input name="CommentTF" type="radio" value="3" checked>只允许指定用户组的会员投稿<br>  
					
					 </td>          
				 </tr>
				 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				  <td width='40' align='center' class='clefttitle'><Input type='checkbox' value='1' name='ModifyAllowArrGroupID'></td>
                  <td  class='clefttitle' width=200><strong>允许此<%=MyFolderName%>下投稿的会员组：</strong><br><font color=blue>当上面选择④时，请在此设置允许在此<%=MyFolderName%>下投稿的会员组</font><font color=blue>如果该<%=MyFolderName%>允许投稿，请在此设置允许在此<%=MyFolderName%>下投稿的会员组</font></td>
                  <td><%=KS.GetUserGroup_CheckBox("AllowArrGroupID","",3)%></td>
                </tr>
			  </table>
		    </div>
           
<br /><B>说明：</B><br />1、若要批量修改某个属性的值，请先选中其左侧的复选框，然后再设定属性值。<br />2、这里显示的属性值都是系统默认值，与所选<%=MyFolderName%>的已有属性无关<br />
<p align=center>
  <Input id=Action type=hidden value="DoBatch" name=Action>
  <Input id=ChannelID type="hidden" value=<%=ChannelID%> name="ChannelID"> 
  <Input style="CURSOR: hand" type=submit value="执行批处理" class="button" name=Submit>&nbsp;
        <Input id=Cancel style="CURSOR: hand" class="button" onClick="window.location.href='KS.Class.asp?ChannelID=<%=ChannelID%>'" type=button value=" 取 消 " name=Cancel></p>
		</td>
    </tr>
  </table>
</FORM>
</div>
<%
End Sub

Sub AttributeSave()
   Dim I,ClassID:ClassID=Replace(Request.Form("ClassID")," ","")
   Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
   Dim ClassIDArr:ClassIDArr=Split(ClassID,",")
   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   For I=0 To Ubound(ClassIDArr)
     RS.Open "Select * From KS_Class Where ID='" & ClassIDArr(I) & "'",conn,1,3
	 If Not RS.Eof Then
	   If KS.ChkClng(KS.G("ModifyTopFlag"))=1 Then RS("TopFlag")=KS.ChkClng(KS.G("TopFlag"))
	   If KS.ChkClng(KS.G("ModifyCommentTF"))=1 Then 
	    RS("CommentTF")=KS.ChkClng(KS.G("CommentTF"))
		RS("AllowArrGroupID")=KS.G("AllowArrGroupID")
	   End If
	   If KS.ChkClng(KS.G("ModifyChannelTemplateID"))=1 Then
	     If RS("TN")="0" Then RS("FolderTemplateID")=KS.G("ChannelTemplateID") 
	   End If
	   If KS.ChkClng(KS.G("WAPModifyChannelTemplateID"))=1 Then
	     If RS("TN")="0" Then RS("WAPFolderTemplateID")=KS.G("WAPChannelTemplateID") 
	   End If
	   
	   If KS.ChkClng(KS.G("ModifyFolderTemplateID"))=1 Then
	    If rs("TN")<>"0" Then RS("FolderTemplateID")=KS.G("FolderTemplateID")
	   End If
	   If KS.ChkClng(KS.G("WAPModifyFolderTemplateID"))=1 Then
	    If rs("TN")<>"0" Then RS("WAPFolderTemplateID")=KS.G("WAPFolderTemplateID")
	   End If
	   
	   If KS.ChkClng(KS.G("ModifyWapSwitch"))=1 Then RS("WapSwitch")=KS.ChkClng(KS.G("WapSwitch"))
	   If KS.ChkClng(KS.G("ModifyMailTF"))=1 Then RS("MailTF")=KS.ChkClng(KS.G("MailTF"))
	   
	   If KS.ChkClng(KS.G("ModifyFolderFsoIndex"))=1 Then RS("FolderFsoIndex")=Request("FolderFsoIndex")
	   If KS.ChkClng(KS.G("ModifyTemplateID"))=1 Then 
	      RS("TemplateID")=KS.G("TemplateID")
		  Conn.Execute("Update " &KS.C_S(ChannelID,2) & " Set TemplateID='"& KS.G("TemplateID") & "' Where Tid='" &ClassIDArr(I) &"'")
	   End If
	   If KS.ChkClng(KS.G("WAPModifyTemplateID"))=1 Then 
	      RS("WAPTemplateID")=KS.G("WAPTemplateID")
		  If KS.C_S(ChannelID,6)=1 or KS.C_S(ChannelID,6)=2 or KS.C_S(ChannelID,6)=3 or KS.C_S(ChannelID,6)=5 then
		  Conn.Execute("Update " &KS.C_S(ChannelID,2) & " Set WAPTemplateID='"& KS.G("WAPTemplateID") & "' Where Tid='" &ClassIDArr(I) &"'")
		  end if
	   End If
	   
	   If KS.ChkClng(KS.G("ModifyPubTF"))=1 Then RS("PubTF")=KS.ChkClng(KS.G("PubTf"))
	   
	   If KS.ChkClng(KS.G("ModifyFnameType"))=1 Then RS("FnameType") = KS.G("FnameType")
	   If KS.ChkClng(KS.G("ModifyFsoType"))=1 Then RS("FsoType")=KS.ChkClng(KS.G("FsoType"))
	   
	   If KS.ChkClng(KS.G("ModifyClassPurview"))=1 Then RS("ClassPurview")=KS.ChkClng(KS.G("ClassPurview"))
	   If KS.ChkClng(KS.G("ModifyGroupID"))=1 Then RS("DefaultArrGroupID")=Request("GroupID")
	   If KS.ChkClng(KS.G("ModifyAllowArrGroupID"))=1 Then RS("AllowArrGroupID")=Request("AllowArrGroupID")
	   If KS.ChkClng(KS.G("ModifyReadPoint"))=1 Then  RS("DefaultReadPoint")=KS.ChkClng(KS.G("ReadPoint"))
	   If KS.ChkClng(KS.G("ModifyDividePercent"))=1 Then 
	            Dim DividePercent:DividePercent=KS.G("DividePercent")
				If Not IsNumeric(DividePercent) Then
				 DividePercent=0
				End If
	      RS("DefaultDividePercent")=DividePercent
	   End If
	   If KS.ChkClng(KS.G("ModifyChargeType"))=1 Then 
	     RS("DefaultChargeType")=KS.ChkClng(KS.G("ChargeType"))
		 RS("DefaultPitchTime")=KS.ChkClng(KS.G("PitchTime"))
		 RS("DefaultReadTimes")=KS.ChkClng(KS.G("ReadTimes"))
	   End If
	   If KS.ChkClng(KS.G("ModifyAd"))=1 Then
	       Dim AdPa,AdParam
		   AdParam=KS.G("AdParam")
	      If KS.G("ADtype")="2" then
			AdPa=KS.ChkClng(KS.G("ShowADTF")) & "%ks%" & AdParam &"%ks%" & KS.G("ADType") & "%ks%" & Request.Form("AdCode") & "%ks%"
		  else
			AdPa=KS.ChkClng(KS.G("ShowADTF")) & "%ks%" & AdParam &"%ks%" & KS.G("ADType") & "%ks%"& KS.G("AdUrl") & "%ks%" & KS.G("AdLinkUrl")
		  end if
		  Dim ClassBasicInfo:ClassBasicInfo=RS("ClassBasicInfo")
		  Dim Carr:CArr=Split(ClassBasicInfo,"||||")
		  Dim ii,NStr
		  For II=0 To Ubound(Carr)
		    If II=0 Then
			  Nstr=Carr(0)
			ElseIf II=4 Then
			 Nstr=Nstr & "||||" & AdPa
			Else
			 Nstr=Nstr & "||||" & Carr(II)
			End If
		  Next
		  RS("ClassBasicInfo")=Nstr
	   End If
	   
	   RS.Update
	 End If
	 RS.Close
   Next
   Set RS=Nothing
   KS.AlertHintScript "恭喜," & MyFolderName &"批量设置成功!"
End Sub

Sub ShowChannelOption()
		  With KS
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				 if (Node.SelectSingleNode("@ks21").text="1" and KS.ChkClng(Node.SelectSingleNode("@ks0").text)<9) or (Node.SelectSingleNode("@ks0").text="5" And KS.SSetting(0)<>"0") Then
				  if request("mid")=Node.SelectSingleNode("@ks0").text or request("channelid")=Node.SelectSingleNode("@ks0").text then
				   .echo "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected>" &Node.SelectSingleNode("@ks1").text & "</option>"
				  else
				   .echo "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" &Node.SelectSingleNode("@ks1").text & "</option>"
				  end if
			    End If
			next
         End With
End Sub


Sub MainPage()
   With KS
	.echo "<div class='pageCont2'><form name='myform' id='myform' action='KS.Class.asp' method='post'>"
	if session("class_modelid")="" or request("channelid")="" then
    .echo "<div style='float:right;margin-top:10px'> &nbsp;"
	.echo "<select name='sc' onchange=""location.href='?from=all&a=" & request("a") & "&mid='+this.value;"" class='md10'>"
	.echo "<option value=''>---按模型查看管理---</option>"
	ShowChannelOption
	.echo "</select>"
	.echo "</div>"
	.echo "<div class='tabTitle'>栏目管理</div>"
	end if
	

	.echo "<input type='hidden' name='oids' id='oids' value=''/>"
   	.echo " <table border='0' cellpadding='0' cellspacing='0'  width='100%' style=""table-layout:fixed;"" align='center'>"
	.echo " <tr class='sort'>"
	.echo " <td style=""width:40px;padding-left:10px;text-align:left""><input type=""checkbox"" name=""cbox"" onclick=""if (this.checked){Select(0)}else{Select(2)}""/></td>"
	.echo " <td style=""width:100px;padding-left:2px;text-align:left"">ID</td>"
	.echo " <td style=""width:100px;padding-left:2px;text-align:left"">ClassID</td>"
	.echo " <td style=""text-align:left;padding-left:10px;"">"& MyFolderName&"名称 </td>"
	.echo "</tr>" & vbNewLine
	.echo "<input type='hidden' name='action' value='Del' id='action'>"
	If KS.C("SuperTF")<>1Then
	Dim Param:Param=" And ID IN('" & replace(KS.C("PowerList"),",","','") &"')"
	End If
	
	Dim ClassXML,Node,ClassType,TypeStr,ID
	Dim Sqlstr
	If KS.G("A")="extall" Then
	 param=" where 1=1"
	Else
	 param = " where tj=1"
	End If
	if request("channelid")<>"" then param=param &" and a.channelid=" & ks.chkclng(request("channelid"))
	if request("mid")<>"" then param=param &" and a.channelid=" & ks.chkclng(request("mid"))
	SQLstr = "select a.ID,a.FolderName,a.classid,a.FolderOrder,a.ClassType,a.ChannelID,a.tj,a.tn,a.adminpurview,a.ts,a.child from KS_Class a inner join ks_channel b on a.channelid=b.channelid " & Param & " and (b.channelstatus=1 or a.channelid=5) Order BY root,folderorder"
    
	Dim RS:Set Rs = Server.CreateObject("adodb.recordset")
	Rs.Open SQLstr, Conn, 1, 1
	If RS.Eof And RS.Bof Then
	.echo "<tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'""><td class='splittd' height='50' style='text-align:center' colspan=4>还没有添加"& MyFolderName&"！</td></tr>"
	Else
	        totalPut = rs.recordcount
		    If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
				RS.Move (CurrentPage - 1) * MaxPerPage
			End If
	       Set ClassXML=KS.ArrayToXML(RS.GetRows(MaxPerPage),RS,"row","xmlroot")
		   
		   	If IsObject(ClassXML) Then
			For Each Node In ClassXML.DocumentElement.SelectNodes("row")
				ID=Node.SelectSingleNode("@id").text
				ClassType=Node.SelectSingleNode("@classtype").text
				If KS.C("SuperTF")=1 or KS.FoundInArr(Node.SelectSingleNode("@adminpurview").text,KS.C("GroupID"),",") or Instr(KS.C("ModelPower"),KS.C_S(Node.SelectSingleNode("@channelid").text,10)&"1")>0 Then 
					if ClassType="2" Then
					 TypeStr="<font color=blue>(外)</font>"
					ElseIf ClassType="3" Then
					 TypeStr="<font color=green>(单)</font>"
					Else
					 TypeStr=""
					End If
					If KS.G("A")="extall" Then
					  if Node.SelectSingleNode("@tj").text="1"  Then
						.echo "<tr height='20' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">" &vbcrlf
						else
						dim nn,myclass:myclass=""
						dim tsarr:tsarr=split(Node.SelectSingleNode("@ts").text,",")
						for nn=0 to ubound(tsarr)-2
						 if myclass="" then
						  myclass="my" &tsarr(nn)
						 else
						  myclass=myclass & " my" &tsarr(nn)
						 end if
						next
						
						.echo "<tr class='list " &myclass&"' onmouseout=""$(this).addClass('list');$(this).removeClass('listmouseover')"" onmouseover=""$(this).removeClass('list');$(this).addClass('listmouseover')"">"&vbcrlf
						end if
					Else
					.echo "<tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					End If
					.echo " <td class='splittd' style=""width:40px;padding-left:10px"">"
					.echo "  <input type='checkbox' name='id' id='c" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@id").text & "'></td>" & vbNewLine
					.echo " <td class='splittd' style=""width:100px;padding-left:2px"">"
					.echo " " & Node.SelectSingleNode("@id").text
					.echo " </td>" & vbNewLine
					.echo " <td class='splittd' style=""width:100px;padding-left:2px"">"
					.echo " " & Node.SelectSingleNode("@classid").text
					.echo " </td>" & vbNewLine
					.echo " <td style='padding-left:5px' class='splittd'>"
					if Node.SelectSingleNode("@tj").text="1" Then
						If KS.G("A")="extall" Then
						.echo "<img id='C" & ID & "' src='../images/folder/Open.gif' align='absmiddle'>"
						.echo "<a href=""javascript:;"" onclick=""o($(this),'" & Node.SelectSingleNode("@id").text&"');"">" & Node.SelectSingleNode("@foldername").text & "</a>"
						Else
						.echo "<img id='C" & ID & "' src='../images/folder/Close.gif' align='absmiddle'>"
						.echo "<a href=""javascript:ExtSub('" & ID & "');"">" & Node.SelectSingleNode("@foldername").text & "</a>"
						End If
					Else
						Dim TJ,SpaceStr,k,Total
						SpaceStr=""
						TJ=Node.SelectSingleNode("@tj").text
						For k = 1 To TJ - 1
						  SpaceStr = SpaceStr & "──"
						Next
						If KS.G("A")="extall" Then
						  if Node.SelectSingleNode("@child").text=0 then
						   .echo "<img src='../images/folder/HR.gif' align='absmiddle'>" & SpaceStr & "" & Node.SelectSingleNode("@foldername").text
						  Else
						   .echo "<img src='../images/folder/HR.gif' align='absmiddle'>" & SpaceStr & "<a href=""javascript:;"" onclick=""$('.my" & Node.SelectSingleNode("@id").text&"').toggle()"">" & Node.SelectSingleNode("@foldername").text &"</a>"
						  End If
						Else
						.echo "<img src='../images/folder/HR.gif' align='absmiddle'>" & SpaceStr & "<a href=""javascript:ExtSub('" & ID & "');"">" & Node.SelectSingleNode("@foldername").text
						End If
					End If
					.echo TypeStr & "" & vbNewLine
					.echo " <span class=""noshow"">"
					.echo "<a href='" & KS.GetFolderPath(id) & "' target='_blank'>预览</a>"
					.echo " | <a href=""#"" onclick=""javascript:AddInfo(" & KS.C_S(Node.SelectSingleNode("@channelid").text,6) & "," & Node.SelectSingleNode("@channelid").text &",'" & ID & "');"">添加" & KS.C_S(Node.SelectSingleNode("@channelid").text,3) & "</a>"
					.echo " | <a href=""javascript:CreateClass('" & ID & "');"">添加子"& MyFolderName&"</a>"
					.echo " | <a href=""javascript:EditClass('" & ID & "');"">编辑"& MyFolderName&"</a>"
					
					.echo " | <a href=""javascript:;"" onclick=""if (confirm('删除"& MyFolderName&"操作将删除此"& MyFolderName&"中的所有子"& MyFolderName&"和文档，并且不能恢复！确定要删除此"& MyFolderName&"吗?')){ location.href='KS.Class.asp?ChannelID=" & Node.SelectSingleNode("@channelid").text & "&Action=Del&Go=Class&ID=" & ID & "&oids='+$('#oids').val();}"">删除"& MyFolderName&"</a>"
					.echo " | <a href=""javascript:MoveClass('" & ID & "');"">移动</a>"
					.echo " | <a href=""javascript:DelInfo(" & Node.SelectSingleNode("@channelid").text & ",'" & ID & "');"">清空</a>"
			
					.echo "</span> </td>" & vbNewLine
					.echo "</tr>" & vbNewLine
					.echo "<tr><td id='sub" & ID &"' colspan=4>"
					.echo "</td></tr>"
			   End If
			Next
			End If
		   
	End If
	Rs.Close
	Set Rs = Nothing

	.echo "</form></div>"
	.echo "</table>"
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	
  End With
End Sub

Sub SubTreeList(parentid)
      If KS.C("SuperTF")<>1 Then
	   '  Param=Param & " And ID IN('" & replace(Application(KS.C("GroupID")&"PowerList"),",","','") &"')"
	  End If
	  KS.LoadClassConfig()
       
	  Dim SubTypeList, p,SpaceStr, k, Total, Num,ID,TJ,SQL,N,SubClassXML,Node,TypeStr,ClassType
	  Num = 0
	  if request("channelid")<>"" then p=" && @ks12='" & request("channelid")&"'"
	  For Each Node In Application(ks.SiteSN&"_class").documentelement.selectnodes("class[@ks13='"&parentid&"'" &p&"]")
	    Num = Num + 1:SpaceStr = "":TJ = CInt(Node.SelectSingleNode("@ks10").text)
		For k = 1 To TJ - 1
		  SpaceStr = SpaceStr & "──"
		Next
	   ID = Node.SelectSingleNode("@ks0").text
	   ClassType=Node.SelectSingleNode("@ks14").text
		if ClassType="2" Then
		 TypeStr="<font color=blue>(外)</font>"
		ElseIf ClassType="3" Then
		 TypeStr="<font color=green>(单)</font>"
		Else
		 TypeStr=""
		End If
		With KS
		If (KS.C("SuperTF")=1 or KS.FoundInArr(Node.SelectSingleNode("@ks16").text,KS.C("GroupID"),",")) and (KS.C_S(Node.SelectSingleNode("@ks12").text,21)=1 or Node.SelectSingleNode("@ks12").text=5) or Instr(KS.C("ModelPower"),KS.C_S(Node.SelectSingleNode("@ks12").text,10)&"1")>0 Then 
		 .echo " <table border='0' cellpadding='0' cellspacing='0'  style=""table-layout:fixed;"" width='100%' align='center'>"
	     .echo " <tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		 .echo " <td style=""width:40px;padding-left:10px;text-align:left"" class='splittd'><input type='checkbox' name='id' id='c" & id & "' value='" & id & "'></td>" & vbNewLine
		 .echo " <td style=""width:100px;padding-left:2px;text-align:left"" class='splittd'>" & id & " </td>" & vbNewLine
		 .echo " <td style=""width:100px;padding-left:2px;text-align:left"" class='splittd'>" & Node.SelectSingleNode("@ks9").text & " </td>" & vbNewLine
		 .echo " <td style='padding-left:5px;' class='splittd'><img align='absmiddle' src='../images/folder/HR.gif'>" & SpaceStr & "<a href=""javascript:ExtSub('" & ID & "');"">" & Node.SelectSingleNode("@ks1").text& TypeStr & "</a>" & vbNewLine
		 .echo " <span class=""noshow"">"
		 .echo "<a href='" & KS.GetFolderPath(id) & "' target='_blank'>预览</a>"
		 .echo " | <a href=""#"" onclick=""javascript:AddInfo(" & KS.C_S(Node.SelectSingleNode("@ks12").text,6) & "," & Node.SelectSingleNode("@ks12").text & ",'" & ID & "');"">添加" & KS.C_S(Node.SelectSingleNode("@ks12").text,3) & "</a>"
		 .echo " | <a href=""javascript:CreateClass('" & ID & "');"">添加子"& MyFolderName&"</a>"
		 .echo " | <a href=""javascript:EditClass('" & ID & "');"">编辑" & MyFolderName &"</a>"
		 .echo " | <a href=""javascript:;"" onclick=""if (confirm('删除"& MyFolderName&"操作将删除此"& MyFolderName&"中的所有子"& MyFolderName&"和文档，并且不能恢复！确定要删除此"& MyFolderName&"吗?')){ location.href='KS.Class.asp?ChannelID=" & Node.SelectSingleNode("@ks12").text & "&Action=Del&Go=Class&ID=" & ID & "&oids='+$('#oids').val();}"">删除"& MyFolderName&"</a>"
		.echo " | <a href=""javascript:MoveClass('" & ID & "');"">移动</a>"
		 .echo " | <a href=""javascript:DelInfo(" & Node.SelectSingleNode("@ks12").text & ",'" & ID & "');"">清空</a>"
		 .echo " </span></td>" & vbNewLine
		 .echo "</tr>" & vbNewLine
		
		 .echo "<tr><td id='sub" & ID &"' colspan=4>"
		 If KS.G("A")="extall" Then	Call SubTreeList(ID)
		.echo "</td></tr>"
		 .echo "</table>"
		End If
	   End With
	   'Call SubTreeList(ID)
	 Next
	End Sub
	
	'一级栏目排序
	Sub OrderOne()
	   With KS
	   	.echo "<div class='pageCont2'>"
		.echo " <table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
		.echo " <tr class='sort'>"
		.echo " <td  style='padding-left:20px;text-align:left;'>"& MyFolderName&"名称 </td>"
		.echo " <td>序号</td>"
		.echo " <td>一级"& MyFolderName&"排序操作</td>"
		.echo "</tr>" & vbNewLine
		Dim SQLStr,ClassXml,Node,i,k
		SQLstr = "select a.ID,a.FolderName,a.FolderOrder,a.ClassType,a.ChannelID,a.tj,a.root from KS_Class a inner join KS_Channel B on a.channelid=b.channelid Where a.TJ=1 and b.channelstatus=1 Order BY a.root,a.folderorder"
		maxperpage=100
		Dim RS:Set Rs = Server.CreateObject("adodb.recordset")
		Rs.Open SQLstr, Conn, 1, 1
		If Not RS.Eof Then
				totalPut = rs.recordcount
				If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrentPage - 1) * MaxPerPage
				End If
				i=(currentpage-1)*maxperpage
			   Set ClassXML=KS.ArrayToXML(RS.GetRows(MaxPerPage),RS,"row","xmlroot")
			   For Each Node In ClassXML.DocumentElement.SelectNodes("row")
			    .echo "<tr>"
			    .echo "<td class='splittd' style='padding-left:20px'>" & Node.SelectSingleNode("@foldername").text & "</td>"
				.echo "<td class='splittd' align='center'>" & Node.SelectSingleNode("@root").text & "</td>"
				.echo "<td class='splittd'>"
				
				.echo "<table border='0' width='100%'><tr>"
				.echo "<form name='upform' action='KS.Class.asp?action=DoUpOrderSave' method='post'>"
				.echo "<input type='hidden' value='" & Node.SelectSingleNode("@root").text & "' name='croot'>"
				.echo "<td width='50%'>"
				if i<>0 then
				 .echo "<select name='MoveNum'><option value=0>↑向上移动</option>"
				 for k=1 to i
				 .echo "<option value=" & k &">" & k &"</option>"
				 next
				 .echo "</select> <input type='submit' value='修改' class='button'>"
				end if
				.echo "</td></form>"
				.echo "<form name='downform' action='KS.Class.asp?action=DoDownOrderSave' method='post'>"
				.echo "<input type='hidden' value='" & Node.SelectSingleNode("@root").text & "' name='croot'>"
				.echo "<td widht='100%'>"
				
				if i<>totalput-1 then
				 .echo "<select name='MoveNum'><option value=0>↓向下移动</option>"
				 for k=1 to totalput-i-1
				 .echo "<option value=" & k &">" & k &"</option>"
				 next
				 .echo "</select> <input type='submit' value='修改' class='button'>"
                end if
				.echo "</td></form>"
				.echo "</tr></table>"
				
				
				i=i+1
				
				.echo "</td>"
				.echo "</tr>"
			   Next
		End If
		.echo "<tr><td colspan=4>"
		Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.echo "</td></tr>"
		.echo "</table>"
		.echo "</div>"
	  End With
	End Sub
      
	Sub DoUpOrderSave()
	 Dim TRoot,i,Croot:croot=KS.ChkClng(Request("croot"))
	 Dim MoveNum:MoveNum=KS.ChkClng(Request("MoveNum"))
	 If MoveNum=0 Then KS.AlertHintScript "对不起,您没有选择位移量!"
	 Dim MaxRootID:MaxRootID=Conn.Execute("select max(Root) From KS_Class")(0)+1
	 '先将当前栏目移至最后，包括子栏目
	 Conn.Execute("Update KS_Class set Root=" & MaxRootID & " where Root=" & cRoot)
	 '然后将位于当前栏目以上的栏目的RootID依次加一，范围为要提升的数字
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "select * From KS_Class where tj=1 and Root<" & cRoot  &  " order by Root desc",conn,1,1
	 If Not RS.Eof Then
	     i=1
		 Do While Not RS.Eof
		  tRoot=rs("Root")       '得到要提升位置的RootID，包括子栏目
		  Conn.Execute("Update KS_Class set Root=Root+1 where Root=" & tRoot)
		  i=i+1
		  if i>MoveNum Then Exit Do
		  RS.MoveNext
		 Loop
		 '然后再将当前栏目从最后移到相应位置，包括子栏目
		 Conn.Execute("Update KS_Class set Root=" & tRoot & " where Root=" & MaxRootID)
	 End If
	 	 RS.CLose
	 Set RS=Nothing
      KS.AlertHintScript "恭喜,上移成功!"
	End Sub
	
	Sub DoDownOrderSave()
	 Dim TRoot,i,Croot:croot=KS.ChkClng(Request("croot"))
	 Dim MoveNum:MoveNum=KS.ChkClng(Request("MoveNum"))
	 If MoveNum=0 Then KS.AlertHintScript "对不起,您没有选择位移量!"
      Dim MaxRootID: MaxRootID = KS.ChkClng(Conn.Execute("select max(Root) From KS_Class")(0)) + 1
	  '先将当前栏目移至最后，包括子栏目
	  Conn.Execute("Update KS_Class set Root=" & MaxRootID & " where Root=" & cRoot)
      '然后将位于当前栏目以上的栏目的RootID依次减一，范围为要提升的数字
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "select * From KS_Class where tj=1 and Root>" & cRoot  &  " order by Root",conn,1,1
	  If Not RS.Eof Then
	       i=1
	     Do While NOT rs.eOF
           tRoot = rs("Root") '得到要提升位置的RootID，包括子栏目
           Conn.Execute("Update KS_Class set Root=Root-1 where Root=" & tRoot)
		   i = i + 1
           if (i > MoveNum) then exit do
		   RS.MoveNext
		 Loop
		 '然后再将当前栏目从最后移到相应位置，包括子栏目
		 Conn.Execute("Update KS_Class set Root=" & tRoot & " where Root=" & MaxRootID)
	  End If
	 	 RS.CLose
	 Set RS=Nothing
      KS.AlertHintScript "恭喜,下移成功!"
	End Sub
	
	'N级栏目排序
	Sub OrderN()
     With KS
	 	.echo "<div class='pageCont2'>"
		.echo " <table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
		.echo " <tr class='sort'>"
		.echo " <td  style='padding-left:20px;text-align:left'>栏目名称 </td>"
		.echo " <td>序号</td>"
		.echo " <td>N级栏目排序操作</td>"
		.echo "</tr>" & vbNewLine
		Dim SQLStr,ClassXml,Node,i,k
		SQLstr = "select a.ID,a.FolderName,a.FolderOrder,a.ClassType,a.ChannelID,a.tj,a.root,a.tn,a.ts,a.child from KS_Class  a inner join KS_Channel B On a.channelid=B.channelid where b.channelstatus=1 Order BY a.root,a.folderorder"
		maxperpage=100
		Dim RS:Set Rs = Server.CreateObject("adodb.recordset")
		Rs.Open SQLstr, Conn, 1, 1
		If Not RS.Eof Then
				totalPut = rs.recordcount
				If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrentPage - 1) * MaxPerPage
				End If
				i=(currentpage-1)*maxperpage
			   Set ClassXML=KS.ArrayToXML(RS.GetRows(MaxPerPage),RS,"row","xmlroot")
			   For Each Node In ClassXML.DocumentElement.SelectNodes("row")
			  '  or KS.ChkClng(Node.SelectSingleNode("@child").text)<>0
				if Node.SelectSingleNode("@tj").text="1"  Then
			    .echo "<tr>" &vbcrlf
				else
				dim nn,myclass:myclass=""
				dim tsarr:tsarr=split(Node.SelectSingleNode("@ts").text,",")
				for nn=0 to ubound(tsarr)-2
				 if myclass="" then
				  myclass="my" &tsarr(nn)
				 else
				  myclass=myclass & " my" &tsarr(nn)
				 end if
				next
				
			    .echo "<tr class='" &myclass&"'>"&vbcrlf
				end if
			    .echo "<td class='splittd'  style='padding-left:20px'>" 
				if Node.SelectSingleNode("@tj").text="1" Then
					.echo "<img src='../images/folder/Open.gif' align='absmiddle'>"
					.echo "<a href=""javascript:;"" onclick=""o($(this),'" & Node.SelectSingleNode("@id").text&"');""><strong>" & Node.SelectSingleNode("@foldername").text & "</strong></a>"
				Else
				    Dim TJ,SpaceStr,Total
					SpaceStr=""
				    TJ=Node.SelectSingleNode("@tj").text
					For k = 1 To TJ - 1
					   SpaceStr = SpaceStr & "──"
					Next
					if Node.SelectSingleNode("@child").text=0 then
					.echo "<img src='../images/folder/HR.gif' align='absmiddle'>" & SpaceStr & Node.SelectSingleNode("@foldername").text
					else
					.echo "<img src='../images/folder/HR.gif' align='absmiddle'>" & SpaceStr & "<a href=""javascript:;"" onclick=""$('.my" & Node.SelectSingleNode("@id").text&"').toggle()"">" & Node.SelectSingleNode("@foldername").text&"</a>"
					end if
				End If
				
				.echo "</td>"
				.echo "<td class='splittd' align='center'>" & Node.SelectSingleNode("@folderorder").text & "</td>"
				.echo "<td class='splittd'>"
				
				if Node.SelectSingleNode("@tj").text="1" Then
				    .echo "&nbsp;"
				Else
					.echo "<table border='0' width='100%'><tr>"
					.echo "<form name='upform' action='KS.Class.asp?action=DoUpOrderNSave' method='post'>"
					.echo "<input type='hidden' value='" & Node.SelectSingleNode("@id").text & "' name='id'>"
					.echo "<td width='50%'>"
					
					'如果不是一级栏目，则算出相同深度的栏目数目，得到该栏目在相同深度的栏目中所处位置（之上或者之下的栏目数）
					'所能提升最大幅度应为For i=1 to 该版之上的版面数
					Dim Trs,UpMoveNum,DownMoveNum
					Set trs = Conn.Execute("select count(ID) from KS_Class where TN='" & Node.SelectSingleNode("@tn").text & "' and FolderOrder<" & Node.SelectSingleNode("@folderorder").text & "")
					UpMoveNum = trs(0)
					If KS.IsNul(UpMoveNum) Then UpMoveNum = 0
					
					
					if UpMoveNum>0 then
					 .echo "<select name='MoveNum'><option value=0>↑向上移动</option>"
					 for k=1 to UpMoveNum
					 .echo "<option value=" & k &">" & k &"</option>"
					 next
					 .echo "</select> <input type='submit' value='修改' class='button'>"
					end if
					.echo "</td></form>"
					.echo "<form name='downform' action='KS.Class.asp?action=DoDownOrderNSave' method='post'>"
					.echo "<input type='hidden' value='" & Node.SelectSingleNode("@id").text & "' name='id'>"
					.echo "<td widht='100%'>"
					
					'所能降低最大幅度应为For i=1 to 该版之下的版面数
					Set trs = Conn.Execute("select count(ID) from KS_Class where tn='" & Node.SelectSingleNode("@tn").text & "' and Folderorder>" & Node.SelectSingleNode("@folderorder").text & "")
					DownMoveNum = trs(0)
					If KS.IsNul(DownMoveNum) Then DownMoveNum = 0
					
					if DownMoveNum>0 then
					 .echo "<select name='MoveNum'><option value=0>↓向下移动</option>"
					 for k=1 to DownMoveNum
					 .echo "<option value=" & k &">" & k &"</option>"
					 next
					 .echo "</select> <input type='submit' value='修改' class='button'>"
					end if
					.echo "</td></form>"
				    .echo "</tr></table>"
				End If
				
				i=i+1
				
				.echo "</td>"&vbcrlf
				.echo "</tr>"&vbcrlf
			   Next
		End If
		.echo "<tr><td colspan=4>"
		Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.echo "</td></tr>"
		.echo "</table>"
		.echo "</div>"
	  End With

	End Sub
	
	Sub DoUpOrderNSave()
	 Dim ID:ID=KS.G("ID")
	 Dim MoveNum:MoveNum=KS.ChkClng(Request.Form("MoveNum"))
	 If ID="" Then KS.AlertHintScript "参数错误!"
	 If MoveNum=0 Then KS.AlertHintScript "对不起,您没有选择位移量!"
	 
	Dim parentID,OrderID,ParentPath,Child,sql, tOrderID,rs, trs, moveupnum, oldorders
    
    '要移动的栏目信息
    Set rs = Conn.Execute("select tn,folderOrder,ts,Child from KS_Class where ID='" & ID & "'")
	If RS.Eof Then 
	  RS.Close:Set RS=Nothing
	  KS.AlertHintScript "对不起,参数传递出错啦!"
	End If
    ParentID = rs(0)
    OrderID = rs(1)
    ParentPath = rs(2)
    Child = rs(3)
    rs.Close
    Set rs = Nothing
	
    '获得要移动的栏目的所有子栏目数，然后加1（栏目本身），得到排序增加数（即其上栏目的OrderID增加数AddOrderNum）
    If Child > 0 Then
        Set rs = Conn.Execute("select count(*) from KS_Class where TS like '%" & ParentPath & "%'")
        oldorders = rs(0) +1
        rs.Close
        Set rs = Nothing
    Else
        oldorders = 1
    End If
    
    '和该栏目同级且排序在其之上的栏目------更新其排序，范围为要提升的数字oldorders
    sql = "select ID,FolderOrder,Child,ts from KS_Class where tn='" & ParentID & "' and FolderOrder<" & OrderID & " order by FolderOrder desc"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 3
    i = 0
    Do While Not rs.EOF
        tOrderID = rs(1)
        Conn.Execute ("update KS_Class set FolderOrder=FolderOrder+" & oldorders & " where id='" & rs(0) & "'")
        If rs(2) > 0 Then
            Set trs = Conn.Execute("select ID,FolderOrder from KS_Class where ts like '%" & rs(3) & "%' and id<>'" &rs(0) &"' order by FolderOrder")
            If Not (trs.BOF And trs.EOF) Then
                Do While Not trs.EOF
                    Conn.Execute ("update KS_Class set FolderOrder=FolderOrder+" & oldorders & " where ID='" & trs(0) &"'")
                    trs.MoveNext
                Loop
            End If
            trs.Close
            Set trs = Nothing
        End If
        i = i + 1
        If i >= MoveNum Then
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    '更新所要排序的栏目的序号
    Conn.Execute ("update KS_Class set FolderOrder=" & tOrderID & " where ID='" &ID &"'")
    '如果有下属栏目，则更新其下属栏目排序
    If Child > 0 Then
        i = 1
        Set rs = Conn.Execute("select ID from KS_Class where ts like '%" & ParentPath & "%' and id<>'" & id & "' order by FolderOrder")
        Do While Not rs.EOF
            Conn.Execute ("update KS_Class set FolderOrder=" & tOrderID + i & " where ID='" & rs(0)&"'")
            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    
    KS.AlertHintScript "恭喜,上移成功!"
	End Sub
	
	Sub DoDownOrderNSave()
		 Dim ID:ID=KS.G("ID")
		 Dim MoveNum:MoveNum=KS.ChkClng(Request.Form("MoveNum"))
		 If ID="" Then KS.AlertHintScript "参数错误!"
		 If MoveNum=0 Then KS.AlertHintScript "对不起,您没有选择位移量!"
	
		Dim parentID,OrderID,ParentPath,Child,sql, tOrderID,rs, ii,trs, moveupnum, oldorders
		'要移动的栏目信息
		Set rs = Conn.Execute("select tn,folderOrder,ts,Child from KS_Class where ID='" & ID & "'")
		If RS.Eof Then 
		  RS.Close:Set RS=Nothing
		  KS.AlertHintScript "对不起,参数传递出错啦!"
		End If
		ParentID = rs(0)
		OrderID = rs(1)
		ParentPath = rs(2)
		Child = rs(3)
		rs.Close
		Set rs = Nothing

		'和该栏目同级且排序在其之下的栏目------更新其排序，范围为要下降的数字
			sql = "select ID,FolderOrder,child,ts from KS_Class where tn='" & ParentID & "' and FolderOrder>" & OrderID & " order by FolderOrder"
			Set rs = Server.CreateObject("adodb.recordset")
			rs.Open sql, Conn, 1, 3
			i = 0    '同级栏目
			ii = 0   '同级栏目和子栏目
			Do While Not rs.EOF
				Conn.Execute ("update KS_Class set FolderOrder=" & OrderID + ii & " where ID='" & rs(0) &"'")
				If rs(2) > 0 Then
					Set trs = Conn.Execute("select ID,FolderOrder from KS_Class where ts like '%" & rs(3) & "%' and id<>'"&rs(0) &"' order by FolderOrder")
					If Not (trs.BOF And trs.EOF) Then
						Do While Not trs.EOF
							ii = ii + 1
							Conn.Execute ("update KS_Class set FolderOrder=" & OrderID + ii & " where ID='" & trs(0)&"'")
							trs.MoveNext
						Loop
					End If
					trs.Close
					Set trs = Nothing
				End If
				ii = ii + 1
				i = i + 1
				If i >= MoveNum Then
					Exit Do
				End If
				rs.MoveNext
			Loop
			rs.Close
			Set rs = Nothing
			
	  '更新所要排序的栏目的序号
    Conn.Execute ("update KS_Class set FolderOrder=" & OrderID + ii & " where ID='" & ID &"'")
    '如果有下属栏目，则更新其下属栏目排序
    If Child > 0 Then
        i = 1
        Set rs = Conn.Execute("select ID from KS_Class where TS like '%" & ParentPath & "%' And ID<>'" & ID & "' order by FolderOrder")
        Do While Not rs.EOF
            Conn.Execute ("update KS_Class set FolderOrder=" & OrderID + ii + i & " where ID='" & rs(0)&"'")
            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If

	
		KS.AlertHintScript "恭喜,下移成功!"
	End Sub
%>

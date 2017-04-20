<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_Field
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Field
        Private KS,Action,ChannelID,Page,ItemName,TableName,KSCls
		Private I, totalPut, FieldSql, FieldRS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm, ShowOnUserForm,ShowOnClubForm,Options,OrderID,AllowFileExt,MaxFileSize,Width,Height,EditorType,ShowUnit,UnitOptions,ParentFieldName,MaxLength,GroupID

		Private Sub Class_Initialize()
		  MaxPerPage =50
		  Set KSCls=New ManageCls
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             Action=KS.G("Action")
             ChannelID=KS.ChkClng(KS.G("ChannelID"))
			 
			 TableName=KS.C_S(ChannelID,2)
			 If ChannelID=101 Then
			  TableName="KS_User"   : ItemName= "会员" '会员表
			 Else
			  ItemName=KS.C_S(ChannelID,3)
			 End If

			 if ChannelID=101 Then
		       If Not KS.ReturnPowerResult(0, "KMUA10012")  Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
			   End If
			 Else
		       If Not KS.ReturnPowerResult(0, "KSMM10003") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
			   End If
			 End If			 
			 
		With Response
		    .Write"<!DOCTYPE html><html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		
        .Write "<script language='JavaScript'>"
		.Write "var Page='" & CurrentPage & "';"
		.Write "var ItemName='" & ItemName & "';"
		.Write "var ChannelID=" & ChannelID & ";"
		.Write "</script>"
		.Write "<script src='../../KS_Inc/jquery.js'></script>"
		.Write "<script src='../../KS_Inc/common.js'></script>"
		
		if action="" then
		.Write "<script>"
		.Write "$(document).ready(function(){"
		%>
		$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
		$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
		<%
		.Write "})</script>"
		end if
		%>
		 <script language="javascript">
		function FieldGroup(){
		   location.href='KS.Field.asp?ChannelID='+ChannelID+'&Action=Group';
		   window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('模型管理 >> 模型字段管理 >> <font color=red>字段分组管理</font>')+'&ButtonSymbol=Go';
		}
		function FieldAdd(){
		   location.href='KS.Field.asp?ChannelID='+ChannelID+'&Action=Add&groupid=<%=KS.S("Groupid")%>';
		   window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('模型管理 >> 模型字段管理 >> <font color=red>新增'+ItemName+'自定义字段</font>')+'&ButtonSymbol=Go';
		}
		function EditField(id)
		{
		  location.href="KS.Field.asp?ChannelID="+ChannelID+"&Page="+Page+"&Action=Edit&ID="+id;
		  $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('模型管理 >> 模型字段管理 >> <font color=red>编辑'+ItemName+'自定义字段</font>')+'&ButtonSymbol=GoSave';

		  
		//  window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('模型管理 >> 模型字段管理 >> <font color=red>编辑'+ItemName+'自定义字段</font>')+'&ButtonSymbol=GoSave';
		}
		function DelField(id)
		{
		if (confirm('真的要删除该自定义字段吗?'))
		 location="KS.Field.asp?ChannelID="+ChannelID+"&Action=Del&Page="+Page+"&id="+id;
		  SelectedFile='';
		}
		function FieldControl(op)
		{   var alertmsg='';
			var SelectedFile=get_Ids(document.myform);
			if (SelectedFile!='')
			 {
			   if (op==1)
				{
				if (SelectedFile.indexOf(',')==-1) 
					EditField(SelectedFile)
				  else top.$.dialog.alert('一次只能编辑一个自定义字段!')	
				SelectedFile='';
				}	
			  else if (op==2)    
			   DelField(SelectedFile);
			 }
			else 
			 {
			 if (op==1)
			  alertmsg="编辑";
			 else if(op==2)
			  alertmsg="删除"; 
			 else
			  {
			  alertmsg="操作" 
			  }
			 top.$.dialog.alert('请选择要'+alertmsg+'的自定义字段');
			  }
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 78 : event.keyCode=0;event.returnValue=false;FieldAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;FieldControl(1);break;
			 case 68 : FieldControl(2);break;
		   }	
		else	
		{
		 //if (event.keyCode==46)FieldControl(2);
		 }
		}
		 </script>
		<%
		.Write "</head>"
		.Write "<body topmargin='0' leftmargin='0'  onkeydown='GetKeyDown();'>"
		.Write "<ul id='menu_top'  class='menu_top_fixed'>"
		If ChannelID<>9 and ChannelID<>101 Then
		.Write "<li class='parent' onclick=""FieldGroup();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon merge'></i>字段分组</span></li>"
		End If
		.Write "<li class='parent' onclick=""FieldAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>新增字段</span></li>"
		If Action<>"Group" Then
		.Write "<li class='parent' onclick=""FieldControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon write'></i>修改字段</span></li>"
		.Write "<li class='parent' onclick=""FieldControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon delete'></i>删除字段</span></li>"
		.Write "<li class='parent' onclick=""$('#action').val('setshowonform');$('#v').val('1');$('#myform').submit();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>设置后台显示</span></li>"
		.Write "<li class='parent' onclick=""$('#action').val('setshowonform');$('#v').val('0');$('#myform').submit();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>设置后台不显示</span></li>"
		.Write "<li class='parent' onclick=""$('#action').val('setshowonuserform');$('#v').val('1');$('#myform').submit();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>设置前台显示</span></li>"
		.Write "<li class='parent' onclick=""$('#action').val('setshowonuserform');$('#v').val('0');$('#myform').submit();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>设置前台不显示</span></li>"
		End If

		.Write "<li class='parent' onclick=""location.href='"
		If Action="Group" Then .Write "KS.Field.asp?channelid=" & ChannelID Else .Write "KS.Model.asp"
		.Write "';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>回上一级</span></li>"
		.Write "</ul>"		
		.Write "<div class=""menu_top_fixed_height""></div>"
		
		
			 
			 
			 Select Case Action
			  Case "Group"      Call FieldGroup()
			  Case "Add","Edit" Call FieldAddOrEdit(Action)
			  Case "Del"	    Call FieldDel()
			  Case "order"	    Call FieldOrder()
			  Case "AddSave"    Call FieldAddSave()
			  Case "EditSave"   Call FieldEditSave()
			  Case "setshowonform" Call setshowonform()
			  Case "setshowonuserform" Call setshowonuserform()
			  Case "setmustfill" Call setmustfill()
			  Case Else 	    Call FieldList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		 End With
		End Sub
		
		Sub FieldGroup()
		 %>
        <div class="pageCont2"> 
		  <form name="form1" method="post" action="?action=Group&x=a&channelid=<%=channelid%>">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center"  class="sort"> 
			<td style="width:130px"><strong>编号</strong></td>
			<td style="width:170px"><strong>组名称</strong></td>
			<td style="width:170px"><strong>排序</strong></td>
			<td></td>
		  </tr>
		  <%
		  dim orderid:orderid=KS.ChkClng(conn.execute("select max(orderid) from KS_FieldGroup Where ChannelID="&ChannelID)(0))+1
		  dim rs:set rs = conn.execute("select * from KS_FieldGroup Where ChannelID="&ChannelID &" order by orderid,id")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""4"" height=""25"" align=""center"" class=""tdbg"">本模型字段无分组!</td></tr>"
			else
			   do while not rs.eof%>
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class='splittd' style="width:130px;text-align:center" height="25"><%=rs("id")%> <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
				  <td class='splittd'><input name="GroupName<%=rs("id")%>" type="text" class="textbox" id="TypeName" value="<%=rs("GroupName")%>" size="25"></td>
				  <td class='splittd'><input name="OrderID<%=rs("id")%>" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="25"></td>
				  <td class='splittd'>
				  <span class="noshow">
				  <%if rs("issys")="0" then%>
				  <a href="javascript:;" onclick='top.$.dialog.confirm("删除字段分组将同时删除该分组下的所有字段且此操作不可逆，确定删除吗？",function(){window.location="System/KS.Field.asp?action=Group&x=c&id=<%=rs("id")%>&channelid=<%=channelid%>";})'>删除</a>
				  <%else%>
				  <%end if%>
				  </span>
				  </td>
				</tr>
		  <%rs.movenext
		   loop
		 End IF
		rs.close%>
		    <tr><td height="25" colspan="4" class='splittd'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="批量保存修改" class="button" /></td>
		    </tr>
			</form>
		<form action="?action=Group&x=b&channelid=<%=channelid%>" method="post" name="myform" id="form">
			<tr valign="middle"> 
			  <td class='splittd' height="25" style="width:130px;text-align:center">&nbsp;<strong>&gt;&gt;新增字段分组<<</strong></td>
			  <td class='splittd'><input name="GroupName" type="text" class="textbox" id="GroupName1" size="25"></td>
			  <td class='splittd'><input name="OrderID" type="text" value="<%=orderid%>" class="textbox" id="OrderID1" size="20">
			  </td>
			  <td class='splittd'> <input onclick="return(check());" name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
		</form>
      </table>
      </div>
      <script>
	   function check(){
		    if ($("#GroupName1").val()==''){
				top.dialog.$.alert('分组名称必须输入',function(){
				 $("#GroupName1").focus();
				});
				return false;
			}
		    if ($("#OrderID1").val()==''){
				top.dialog.$.alert('排序号必须输入',function(){
				 $("#OrderIDName1").focus();
				});
				return false;
			}
			return true;
	   }
	  </script>

		<% Select case request("x")
		   case "a"
		        Dim IDs:Ids=KS.FilterIds(KS.S("Id"))
				If Not KS.IsNul(IDS) Then
				  Dim I,IdsArr:IdsArr=Split(Ids,",")
				  For i=0 To Ubound(IdsArr)
				  conn.execute("Update KS_FieldGroup set GroupName='" & KS.G("GroupName" & IdsArr(i)) & "',OrderID=" & KS.ChkClng(KS.G("OrderID" & IdsArr(i))) & " where id="&KS.ChkClng(IdsArr(i)))
				  Next
				End If
				Application(KS.SiteSN & "_FieldGroupXml")=""
				KS.AlertHintScript "恭喜,修改成功!"
		   case "b"
		       If KS.G("GroupName")="" Then Response.Write "<script>top.$.dialog.alert('请输入组名称!',function(){history.back();});</script>":response.end
				conn.execute("Insert into KS_FieldGroup(GroupName,ChannelID,OrderID,isSys)values('" & KS.G("GroupName") & "',"& ChannelID& "," & KS.ChkClng(KS.G("OrderID")) & ",0)")
				Application(KS.SiteSN & "_FieldGroupXml")=""
				KS.AlertHintScript "恭喜,添加成功!"
		   case "c"
		        Dim RSS:Set RSS=Conn.Execute("select * From KS_Field Where GroupID=" & KS.ChkClng(KS.G("id")))
				If Not RSS.Eof Then
				  Do While Not RSS.Eof
                      Dim TableName:TableName=KS.C_S(RSS("ChannelID"),2)
 					  Conn.Execute("Alter Table "& TableName &" Drop column "& RSS("FieldName") &"")
					  If RSS("ShowUnit")="1" Then
					  Conn.Execute("Alter Table "& TableName &" Drop column "& RSS("FieldName") &"_Unit")
					  End if
				   RSS.MoveNext
				  Loop
				End IF
				RSS.Close
				Set RSS=Nothing
				 Conn.Execute("Delete From KS_Field Where GroupID=" & KS.ChkClng(KS.G("id")))
				 conn.execute("Delete from KS_FieldGroup where  issys=0 and id="&KS.ChkClng(KS.G("id")))
				 Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
				 Application(KS.SiteSN & "_FieldGroupXml")=""
				 Call KS.Alert("恭喜，删除成功!","System/KS.Field.asp?action=Group&channelID="&channelid)
		End Select
		
		End Sub
		
		Sub FieldList()
		With Response
		If ChannelID<>9 and ChannelID<>101 Then
			.Write "<div class='tabTitle'>模型字段管理</div><div class=""tabs_header mt0 pt0"">" &vbcrlf
			.Write " <ul class=""tabs"">" &vbcrlf
			.Write " <li"
			if KS.S("GroupID")="" then response.write " class='active'"
			.Write "><a href=""KS.Field.asp?" & KS.QueryParam("groupid") &"""><span>所有字段</span></a></li>"
			Call KS.LoadFieldGroupXML()
			 Dim Node
			 If IsObject(Application(KS.SiteSN & "_FieldGroupXml")) Then
			 For Each Node In Application(KS.SiteSN & "_FieldGroupXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &"]")
				If KS.ChkClng(KS.S("GroupID"))=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
				.Write "<li  class='active'><a href='KS.Field.asp?" & KS.QueryParam("groupid") & "&groupid=" & Node.SelectSingleNode("@id").text &"'><span>" &  Node.SelectSingleNode("@groupname").text &"</span></a></li>" 
				Else
				.Write "<li><a href='KS.Field.asp?" & KS.QueryParam("groupid") & "&groupid=" & Node.SelectSingleNode("@id").text &"'><span>" &  Node.SelectSingleNode("@groupname").text &"</span></a></li>" 
				End If
			 Next
			 End If
			.Write "</ul>"
			.Write "</div>"
		End If
		.Write "<div class='pageCont2 noRadTop'>"
		.Write "<div class='tabTitle'>会员字段管理</div>"
		.Write "<form action='KS.Field.asp?channelid=" & ChannelID&"&page="&CurrentPage &"' name='myform' id='myform' method='post'>"
		.Write "<input type='hidden' name='action' id='action' value='order'/>"
		.Write "<input type='hidden' name='v' id='v' value='0'/>"
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write " <tr class='sort'>"
		.Write "   <td width=""40"" align='center'><input type='checkbox' name='select' onclick=""if (this.checked){Select(0)}else{Select(2)}""/></td>"
		.Write "   <td align='center'>排序</td>"
		.Write "   <td width='100' align='center'>字段名称</td>"
		.Write "   <td align='center'>字段别名</td>"	
		If ChannelID<>9 and ChannelID<>101 Then	 .Write "   <td align='center'>分组</td>"
		.Write "   <td align='center'>归属模型</td>"
		.Write "   <td align='center'>字段类型</td>"
		.Write "   <td align='center'>后台显示</td>"
		.Write "   <td align='center'>前台显示</td>"
		.Write "   <td align='center'>必填</td>"
		.Write "   <td align='center'>↓管理操作</td>"
		.Write " </tr>"
		
		  Dim Param:Param=" Where a.ChannelID=" & ChannelID
		  If KS.G("GroupID")<>"" Then Param=Param & "  and a.groupid=" & KS.ChkClng(KS.G("GroupID"))
			 Set FieldRS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT a.*,B.GroupName FROM KS_Field a left join KS_FieldGroup b ON A.GroupID=B.ID" & Param & " order by a.orderid asc"
				   FieldRS.Open FieldSql, conn, 1, 1
				 If FieldRS.EOF And FieldRS.BOF Then
				   .Write "<tr><td align='center' colspan='10' class='splittd'>分组下没有字段！</td></tr>"
				 Else
					        totalPut = FieldRS.RecordCount
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									FieldRS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
			End If
		 .Write "<tr><td colspan='3' class='operatingBox'><input type='submit' onclick=""$('#action').val('order');"" class='button' value='批量保存排序'> </td></form>"
		 .Write " <td colspan='10' align='right'>"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "<br/><br/><br/></div>"
		.Write "</div>"
		End With
		End Sub
		Sub showContent()
		With Response
		Do While Not FieldRS.EOF
		 if KS.ChkClng(FieldRS("FieldType"))=0 or Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then
		 .Write "<tr name='sysfield' style='display:' class='list' id='u" & FieldRS("FieldID") &"' onclick=""chk_iddiv('" & FieldRS("FieldID") &"')"" onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
		 Else
		 .Write "<tr class='list' id='u" & FieldRS("FieldID") &"' onclick=""chk_iddiv('" & FieldRS("FieldID") &"')"" onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
		 End If
		 .Write "<td class='splittd' style='text-align:center'><input type='checkbox' onclick=""chk_iddiv('" & FieldRS("FieldID") &"')""  name='id' id='c" & FieldRS("FieldID") &"' value='" & FieldRS("FieldID") &"'></td>"
		 .Write "<td class='splittd' style='text-align:center'><input class='textbox' type='text' name='OrderID' style='width:40px;text-align:center' value='" & FieldRS("OrderID") &"'><input type='hidden' name='FieldID' value='" & FieldRS("FieldID") & "'></td>"
		 .Write "  <td class='splittd' nowrap><span FieldID='" & FieldRS("FieldID") & "' onDblClick=""EditField('" & FieldRS("FieldID") & "')""><img src='../Images/Field.gif' align='absmiddle'><span  style='cursor:default;'>" & FieldRS("FieldName") & "</span></span></td>"
		 .Write "   <td align='center' class='splittd' title='" & FieldRS("Title") &"'>" & KS.Gottopic(FieldRS("Title"),15) & " </td>"
		 If ChannelID<>9 and ChannelID<>101 Then .Write "   <td align='center' class='splittd' style='color:#999'>" & FieldRS("GroupName") & " </td>"
		 .Write "   <td align='center' class='splittd'><font color=#888888>"
		 If ChannelID=101 Then
		 .Write "会员系统"
		 Else
		  .Write KS.C_S(ChannelID,1) 
		 End If
		  .Write "</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>"
				 Select Case FieldRS("FieldType")
				  Case 1:.Write "单行文本(text)"
				  Case 2:.Write "文本(不支持HTML)"
				  Case 10:.Write "多行文本(支持HTML)"
				  Case 3:.Write "下拉列表(select)"
				  Case 4:.Write "数字(text)"
				  Case 5:.Write "日期(text)"
				  Case 6:.Write "单选框(radio)"
				  Case 7:.Write "复选框(checkbox)"
				  Case 8:.Write "电子邮箱(text)"
				  Case 9:.Write "文件(text)"
				  Case 11:.Write "联动菜单(text)"
				  Case 12: .Write "小数(text)"
				  Case 13: .Write "文档属性(checkbox)"
				  Case 14: .Write "绑定其它模型(select)"
				 End Select
		  If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then .Write "<font color=#cccccc>[系统]</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("ShowOnForm")=1 Then
		   .Write "<a title='设置为后台不显示' href='?channelid=" & channelid & "&action=setshowonform&id=" & FieldRS("FieldID") &"&v=0'><font color=red>是</font></a>"
		  Else
		   .Write "<a title='设置为后台显示' href='?channelid=" & channelid & "&action=setshowonform&id=" & FieldRS("FieldID") &"&v=1'><font color=green>否</font></a>"
		  End If
		 .Write " </td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("ShowOnUserForm")=1 Then
		   .Write "<a title='设置为前台不显示' href='?channelid=" & channelid & "&action=setshowonuserform&id=" & FieldRS("FieldID") &"&v=0'><font color=red>是</font></a>"
		  Else
		   .Write "<a title='设置为前台显示' href='?channelid=" & channelid & "&action=setshowonuserform&id=" & FieldRS("FieldID") &"&v=1'><font color=green>否</font></a>"
		  End If
		 .Write " </td>"
		  .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("MustFillTF")=1 Then
		   .Write "<a title='设置为必填' href='?channelid=" & channelid & "&action=setmustfill&id=" & FieldRS("FieldID") &"&v=0'><font color=red>是</font></a>"
		  Else
		   .Write "<a title='设置为选填' href='?channelid=" & channelid & "&action=setmustfill&id=" & FieldRS("FieldID") &"&v=1'><font color=green>否</font></a>"
		  End If
		 .Write " </td>"
		 .Write " <td align='center' class='splittd'><a href='javascript:EditField(" & FieldRS("FieldID") &");' class='setA'>修改</a>|"
		 If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then
		 .Write "<font color=#cccccc title='系统字段不允许删除' class='setA'>删除</font>"
		 Else
		 .Write "<a href='javascript:DelField(" & FieldRS("FieldID") &");' class='setA'>删除</a>"
		 End If
		 .Write "</td></tr>"
								I = I + 1
								If I >= MaxPerPage Then Exit Do
							   FieldRS.MoveNext
							   Loop
								FieldRS.Close
						 
         End With
		 End Sub
		 
		 Sub setshowonform()
		    dim id:id=KS.FilterIds(request("id"))
			if ks.isnul(id) then ks.alerthintscript "没有选择字段！"
			conn.execute("update KS_Field Set ShowOnForm=" & KS.ChkClng(Request("v")) & " Where FieldID in(" & ID &")")
			Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存

			response.Redirect request.ServerVariables("HTTP_REFERER")
		 End Sub
		 Sub setshowonuserform()
		    dim id:id=KS.FilterIds(request("id"))
			if ks.isnul(id) then ks.alerthintscript "没有选择字段！"
			conn.execute("update KS_Field Set ShowOnUserForm=" & KS.ChkClng(Request("v")) & " Where FieldID in(" & ID &")")
			Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
			response.Redirect request.ServerVariables("HTTP_REFERER")
		 End Sub
		 
		 Sub setmustfill()
		    dim id:id=KS.FilterIds(request("id"))
			if ks.isnul(id) then ks.alerthintscript "没有选择字段！"
			conn.execute("update KS_Field Set MustFillTF=" & KS.ChkClng(Request("v")) & " Where FieldID in(" & ID &")")
			Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
			response.Redirect request.ServerVariables("HTTP_REFERER")
		 End Sub
		 
		 Sub FieldAddOrEdit(OpType)
		 With Response
		  Dim FieldRS, FieldSql,OpAction,OpTempStr
		 ID = KS.G("ID")
		.Write "<!DOCTYPE html><html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='../../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<script src='../../KS_Inc/jquery.js'></script>"
		.Write "</head>"
		.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		
		 If Optype = "Edit" Then
		     OpAction="EditSave":OpTempStr="编辑"
			 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
			 FieldSql = "Select TOP 1 * From [KS_Field] Where FieldID=" & ID
			 FieldRS.Open FieldSql, conn, 1, 1
			 If Not FieldRS.EOF Then
				 FieldName = Trim(FieldRS("FieldName"))
				 ChannelID = FieldRS("ChannelID")
				 Title = Trim(FieldRS("Title"))
				 Tips = server.HTMLEncode(Trim(FieldRS("Tips")&""))
				 FieldType = Trim(FieldRS("FieldType"))
				 DefaultValue = Trim(FieldRS("DefaultValue"))
				 MustFillTF = FieldRS("MustFillTF")
				 ShowOnForm = FieldRS("ShowOnForm")
				 ShowOnUserForm=FieldRS("ShowOnUserForm")
				 ShowOnClubForm=FieldRS("ShowOnClubForm")
				 Options = Trim(FieldRS("Options"))
				 OrderID= FieldRS("OrderID")
				 AllowFileExt=FieldRS("AllowFileExt")
				 MaxFileSize=FieldRS("MaxFileSize")
				 Width=FieldRS("Width")
				 Height=FieldRS("Height")
				 MaxLength=FieldRS("MaxLength")
				 EditorType=FieldRS("EditorType")
				 ShowUnit=FieldRS("ShowUnit")
				 UnitOptions=FieldRS("UnitOptions")
				 ParentFieldName=FieldRS("ParentFieldName")
				 GroupID=FieldRS("GroupID")
			 End If
	  Else
	     FieldName="KS_":FieldType=1:MustFillTF=0:ShowOnForm=1:ShowOnUserForm=1:ShowOnClubForm=0:AllowFileExt="jpg|gif|png":MaxFileSize=1024:Width=350:Height=80:EditorType="Basic":ShowUnit=0:MaxLength=255:GroupID=KS.ChkClng(Request("GroupID"))
		 OpAction="AddSave":OpTempStr="添加"
		 OrderID=KS.ChkClng(Conn.Execute("Select Max(OrderID) From KS_Field Where GroupID=" & GroupID &" and ChannelID=" & ChannelID)(0))+1
	  End If
		 
		 If ChannelID=101 Then
		  OpTempStr=OpTempStr & "[会员系统]"
		Else
		  OpTempStr=OpTempStr & "[" & KS.C_S(ChannelID,1) &"]"
		End If
		 
		.Write "<div class='tabTitle'>" & OpTempStr &"自定义字段</div>"
		.Write "<form  action='KS.Field.asp?Action=" & OpAction &"' method='post' name='FieldForm' id='FieldForm' class='pageCont2'>"
		.Write "<input type='hidden' value='" & ChannelID & "' name='ChannelID'>"

        .Write "<dl class=""dtable"">" &vbcrlf
		
		If ChannelID=101 Then
		.Write "    <dd style='display:none'><div>所属分组：</div></dd>"
		Else
		.Write "    <dd><div class='left'>所属分组：</div>"
		End If
		 .Write "<select name=""GroupID"" style=""width:350px"">"
		 Call KS.LoadFieldGroupXML()
		 Dim Node
		 If IsObject(Application(KS.SiteSN & "_FieldGroupXml")) Then
		 For Each Node In Application(KS.SiteSN & "_FieldGroupXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &"]")
		    If KS.ChkClng(GroupID)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
		    .Write "<option value='" & Node.SelectSingleNode("@id").text &"' selected>" &  Node.SelectSingleNode("@groupname").text &"</option>" 
		    Else
			.Write "<option value='" & Node.SelectSingleNode("@id").text &"'>" &  Node.SelectSingleNode("@groupname").text &"</option>" 
			End If
		 Next
		 End If
		 .Write "</select>"
		 
		.Write "    </dd>"
    If FieldType="0" Then  '系统内置字段
		.Write "   <dd><div class='left'>字段名称：</div>"
		.Write "    <input class='textbox' name='FieldName' type='text' readonly id='FieldName' value='" & FieldName & "' size='50'> </dd>" &vbcrlf
		.Write "    <dd><div class='left'>字段别名：</div><input name='Title' style=""width:350px"" type='text' id='Title' size='30' class='textbox' value='" & Title & "'> *<span>便于在管理项目中显示请为字段取别名</div></dd>"
        .Write "    <dd><div class='left'>字段类型：</div>"
		.Write "    <input type='hidden' value=" & FieldType & " name='FieldType'><select name=""FieldType"" disabled><option>系统内置字段</option></select></dd>"	
		
		if lcase(FieldName)="otid" then
        .Write "    <dd><div class='left'>选择要绑定的其它模型的分类：</div>"
		.Write "    <select name=""DefaultValue"" style=""width:350px""><option value='0'>不绑定其它模型分类</option>"
		
		If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				  if (Node.SelectSingleNode("@ks21").text="1" and KS.ChkClng(Node.SelectSingleNode("@ks6").text)<9 and KS.ChkClng(Node.SelectSingleNode("@ks0").text)<>KS.ChkClng(ChannelID)) Then
				  Dim MyFolderName:MyFolderName=KS.M_C(KS.ChkClng(Node.SelectSingleNode("@ks0").text),26)
				  If KS.IsNul(MyFolderName) Then MyFolderName="栏目"
				  
				  if KS.ChkClng(DefaultValue)=KS.ChkClng(Node.SelectSingleNode("@ks0").text) then
				   .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected>" &Node.SelectSingleNode("@ks1").text & "的" & MyFolderName &"</option>"
				  else
				   .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" &Node.SelectSingleNode("@ks1").text  & "的" & MyFolderName &"</option>"
				  end if
			    End If
			next
		
		.Write "</select><span>绑定其它模型分类后，可以方便内容页文档之间的关联调用,绑定后请不要随意更改。</span></dd>"	
		end if
			
        .Write "    <dd><div class='left'>后台是否启用：</div>"
		.Write "    <input name='ShowOnForm' type='radio' id='ShowOnForm' value='1'"
		If ShowOnForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='0'"
		If ShowOnForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "    </dd>"
		.Write "    <dd><div class='left'>会员中心是否启用：</div>"
		.Write "   <input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='1'"
		If ShowOnUserForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='0'"
		If ShowOnUserForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write " <span>必须是启用，前台的会员中心才会显示</span>"
		.Write " </dd>"	
		
	Else
		.Write "   <dd><div class='left'>字段名称：</div> <input name='FieldName'  type='text' id='FieldName' value='" & FieldName & "' size='50'"
		If Optype = "Edit" Then .Write " readonly"
		.Write " class='textbox'> * <span>为了和系统字段区分，必须以""KS_""开头,在模板中可以通过""{$KS_字段名称}""进行调用，不能含有中文。</span>"
		.Write "    </dd>"
		.Write "    <dd><div class='left'>字段别名：</div>"
		.Write "     <input name='Title' type='text' id='Title' size='50' class='textbox' value='" & Title & "'> *<span>便于在管理项目中显示</span>"
		.Write "    </dd>"
		.Write "    <dd><div class='left'>附加提示：</div><textarea name='Tips'  id='Tips' class='textbox' style='width:340px;height:30px'>" & Tips & "</textarea><span class=""block"">在输入框旁边的提示信息,可以加入一些javascript事件</span>"
		.Write "    </dd>"
		.Write "    <dd><div class='left'>字段类型：</div>"
		If Optype = "Edit" Then
		.Write "     <input type='hidden' value=" & FieldType & " name='FieldType'><select style=""width:350px;"" name=""FieldType"" disabled>"
		else
		.Write "    <select name=""FieldType"" id='FieldType' style=""width:350px;"" onchange=""Setdisplay(this.value)"">"
		end if
		If ChannelID<>101 and channelid<>9 Then
		.Write " <option value=""13"""
		If FieldType=13 Then .Write " Selected"
		.Write ">文档属性(标签调用属性)</option>"
		End If
		.Write " <option value=""14"""
		If FieldType=14 Then .Write " Selected"
		.Write ">绑定其它模型</option>"

		.Write " <option value=""1"""
		If FieldType=1 Then .Write " Selected"
		.Write ">单行文本(text)</option>"
     	.Write " <option value=""2"""
		If FieldType=2 Then .Write " Selected"
		.Write ">多行文本(不支持HTML)</option>"
     	.Write " <option value=""10"""
		If FieldType=10 Then .Write " Selected"
		.Write ">多行文本(支持HTML)</option>"
		.Write " <option value=""3"""
		If FieldType=3 Then .Write " Selected"
		.Write ">下拉列表(select)</option>"
		.Write " <option value=""11"""
		If FieldType=11 Then .Write " selected"
		.Write " style='color:blue'>联动下拉列表</option>"
        .Write " <option value=""4"""
		If FieldType=4 Then .Write " Selected"
		.Write ">数字(text)</option>"
        .Write " <option value=""12"""
		If FieldType=12 Then .Write " Selected"
		.Write ">小数(text)</option>"
		.Write " <option value=""5"""
		If FieldType=5 Then .Write " Selected"
		.Write ">日期(text)</option>"
		.Write " <option value=""6"""
		If FieldType=6 Then .Write " Selected"
		.Write ">单选框(radio)</option>"
		.Write " <option value=""7"""
		If FieldType=7 Then .Write " Selected"
		.Write ">复选框(checkbox)</option>"
		.Write " <option value=""8"""
		If FieldType=8 Then .Write " Selected"
		.Write ">电子邮箱(text)</option>"
		.Write " <option value=""9"""
		If FieldType=9 Then .Write " Selected"
		.Write ">文件(text)</option>"
		
		.Write " </select>"
		
		
		.Write "  <font id='modelarea'>"
		.Write "    <select id=""BChannelID"" name=""BChannelID"""
		If Optype = "Edit" Then .Write " disabled"
		.Write "><option value='0'>请选择要绑定的模型</option>"
		If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				  if (Node.SelectSingleNode("@ks21").text="1" and KS.ChkClng(Node.SelectSingleNode("@ks6").text)<9 and KS.ChkClng(Node.SelectSingleNode("@ks0").text)<>KS.ChkClng(ChannelID)) Then
				 
				  if KS.ChkClng(DefaultValue)=KS.ChkClng(Node.SelectSingleNode("@ks0").text) then
				   .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected>" &Node.SelectSingleNode("@ks1").text & "</option>"
				  else
				   .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" &Node.SelectSingleNode("@ks1").text  &"</option>"
				  end if
			    End If
			next
		
		.Write "</select>"	
		.Write "</font>"
		
		
		
		.Write "<span>说明：一旦设定不能修改</span>"
		.Write " </dd>"
		
		
		
		
		.Write "  <font id='editorarea'>"
		.Write "    <dd><div class='left'>编辑器类型：</div>"
		.Write "   <input name='EditorType' type='text' id='EditorType' class='textbox' value='" & EditorType & "' size='10'>&nbsp;<select onchange=""$('#EditorType').val(this.value)"" name='selecteditor'><option value='Default'>Default</option><option value='NewsTool'>NewsTool</option><option value='Basic'>Basic</option></select><span>您可以打开/KS_Cls/EditorAPI.asp自定义编辑器类型</span>"
		.Write "       </dd>"
		.Write " </font>"
		
		
		.Write "<font id=""extarea"">"
		.Write "    <dd><div class='left'>允许上传的扩展名：</div>"
		.Write "   <input name='AllowFileExt' type='text' id='AllowFileExt' class='textbox' value='" & AllowFileExt & "' size='50'><span>多个扩展展名，请用逗号“|”隔开</span>"
		.Write "    </dd>"
		.Write "    <dd><div class='left'>允许上传的文件大小：</div>"
		.Write "   <input name='MaxFileSize' type='text' id='MaxFileSize' class='textbox' value='" & MaxFileSize & "' size='8' style='width:50px'>&nbsp;KB <span style='color:#ff0000'>*</span>  <span>提示：1 KB = 1024 Byte，1 MB = 1024 KB<span>  "
		.Write "    </dd>"
		.Write " </font>"
		
		
		.Write "<font id='showdefault'>"
		.Write "    <dd><div class='left'>默认值：</div>"
		.Write "    <textarea name='DefaultValue' id='DefaultValue' class='textbox' style='width:600px;height:100px'>" & server.HTMLEncode(DefaultValue&"") & "</textarea>&nbsp;<span id='darea'>多个默认选项，请用逗号“,”隔开</span>"
		If ChannelID<>101 Then
		 .Write "<br/><font id='dtips1'><font color=green>为便于会员获取默认值，可绑定表KS_User或KS_Enterprise的字段值<br>格式：表名|字段名 如：<font color=red>KS_User|RealName</font></font><br/><font color=blue>也可以将默认值设置为now或date取得当前时间</font></div><div id='dtips2'>输入<font color=red>“1”</font>，则添加文档时默认为该属性为选中状态</font>"
		End If
		.Write "    </dd>"
		.Write "</font>"
		
		.Write "<font id='showattrarea'>"
		
		.Write "    <dd id=""ldArea"" style='display:none'><div class='left'>所属父级字段：</div>"
		  Dim PRS
		  If KS.ChkClng(ID)<>0 Then
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_Field Where ChannelID=" & ChannelID& " and FieldType=11 And FieldID<>" & ID & " Order BY FieldID")
		  .Write "<select name='ParentFieldName'  style='width:350px' disabled>"
		  Else
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_Field Where ChannelID=" & ChannelID& " and FieldType=11 Order BY FieldID")
		  .Write "<select name='ParentFieldName' style='width:350px'>"
		  End If
		  .Write "<option value='0'>--作为一级联动--</option>"
		  Do While Not PRS.Eof
		      If PRS(0)=ParentFieldName Then
		      .Write "<option value='" & PRS(0) & "' selected>" & Prs(1) & "(" & PRS(0) & ")</option>"
			  Else
		      .Write "<option value='" & PRS(0) & "'>" & Prs(1) & "(" & PRS(0) & ")</option>"
			  End If
		  PRS.MoveNext
		  Loop
		  PRS.Close: Set PRS=Nothing
		.Write "      </select> <span>不选择表示一级联动字段，不能指定为下级联动字段，且一旦设定不能修改</span>"
		.Write "    </dd>"
		
		.Write "    <dd id=""OptionsArea"" style=""display:none"" >"
		.Write "     <div>列表选项：</div><textarea name='Options' style='height:70px' cols='50' rows='6' id='Options' class='textbox'>" & Options & "</textarea>"
		.Write "    <span class=""block""><font color=blue>每一行为一个列表选项</font><br>如果值和显示项不同可以用<font color=red>|</font>隔开<br>正确格式如：<font color=red>男</font> 或 <font color=red>0|男</font></span>"
		.Write "    </dd>"
		
		
		.Write "    <dd><div>是否显示下拉单位：</div>"
		 If Optype = "Edit" Then
		    If ShowUnit="1" Then .Write "是" Else .Write "否"
			.Write "<input type='hidden' name='ShowUnit' value='1'>"
		 Else
			.Write  "<input onclick=""$('#unitArea').show()"" name='ShowUnit' type='radio' id='ShowUnit' value='1'"
			If ShowUnit="1" Then .Write " Checked"
			.Write ">是"
			.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input onclick=""$('#unitArea').hide()"" name='ShowUnit' type='radio' id='ShowUnit' value='0'"
			If ShowUnit="0" Then .Write " Checked"
			.WRite ">否"
		 End If
		 .Write "<span>说明：一旦设定不能修改</span>"
		.Write "    </dd>"
		If ShowUnit="1" Then
		.Write "    <dd id=""unitArea"">"
	   else
		.Write "    <dd id=""unitArea"" style=""display:none"">"
	   end if
		.Write "      <div>下拉单位选项：</div><textarea name='UnitOptions' style='width:340px;height:70px' cols='20' rows='3' id='UnitOptions' class='textbox'>" & UnitOptions & "</textarea> "
		.Write "      <span class=""block"">每一行为一个列表选项<br/>如:件 个等</span>>"
		.Write "    </dd>"
		
		
		
		
		.Write "    <dd><div>是否必填：</div>"
		.Write "   <input name='MustFillTF' type='radio' id='MustFillTF' value='1'"
		If MustFillTF="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='MustFillTF' type='radio' id='MustFillTF' value='0'"
		If MustFillTF="0" Then .Write " Checked"
		.WRite ">否</dd>"
		.Write "    <dd><div>后台是否启用：</div>"
		.Write "   <input name='ShowOnForm' type='radio' id='ShowOnForm' value='1'"
		If ShowOnForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='0'"
		If ShowOnForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "    </dd>"
		.Write "    <dd><div>会员中心是否启用：</div><input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='1'"
		If ShowOnUserForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='0'"
		If ShowOnUserForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write " <span>必须是启用，前台的会员中心才会显示</span>"
		.Write "    </dd>"
		
		If  KS.ChkClng(KS.C_S(ChannelID,6))=1 Then
		.Write "    <dd><div>推送到论坛时显示：</div>"
		.Write "  <input name='ShowOnClubForm' type='radio' value='1'"
		If ShowOnClubForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnClubForm' type='radio' value='0'"
		If ShowOnClubForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "<span>指当文章被推送到论坛时是否显示该字段内容</span>"
		.Write "    </dd>"
		End If
		

		.Write "</font>"
    End If		
		.Write "    <dd><div class='left'>显示设置：</div>"
		.Write "    宽度<input name='Width' type='text' site='10' class='textbox' style='text-align:center;width:40px' id='Width' value='" & Width & "'>px <span>例如：200px</span>  长度<input name='MaxLength' type='text' site='10' class='textbox' style='text-align:center;width:40px' id='MaxLength' value='" & MaxLength & "'>个字符 <span>不限制请输入0</span><br><label style='display:none' id='heightarea'>高度<input name='Height' type='text' site='10' class='textbox' style='text-align:center;width:40px' id='Height' value='" & Height & "'>px <span>例如：100px</span></label>"
		.Write "    </dd>"	
					
		.Write "    <dd><div class='left'>排序序号：</div>"
		.Write "    <input name='OrderID' type='text' style='text-align:center' size='8' class='textbox' id='OrderID' value='" & OrderID & "'> <span>序号越小，排在越前面</span>"
		.Write "    </dd>"
	
		.Write "   <input type='hidden' value='" & ID & "' name='id'>"
		.Write "    <input type='hidden' value='" & Page & "' name='page'>"
		.Write "</dl></form>"
		
		
		
		 
		.Write "<Script Language='javascript'>"
		If FieldType<>"0" Then
		.Write "Setdisplay(" & FieldType & ");"
		.Write "function Setdisplay(s)"
		.Write  "{ if (s==3||s==6||s==7||s==11){ $('#OptionsArea').show();} else $('#OptionsArea').hide();if (s==7)$('#darea').show();else $('#darea').hide();if(s==9)$('#extarea').show();else $('#extarea').hide(); if(s==10)$('#editorarea').show();else $('#editorarea').hide();if (s==2||s==10) $('#heightarea').show();else $('#heightarea').hide();if(s==11) $('#ldArea').show(); else $('#ldArea').hide();if(s==13){ $('#showattrarea').hide();$('#dtips1').hide();$('#dtips2').show();}else{ $('#showattrarea').show();$('#dtips2').hide();$('#dtips1').show();}"
		.Write "if(s==13 || s==14){ $('#showdefault').hide();}else{ $('#showdefault').show(); }"
		.Write "if(s!=14){ $('#modelarea').hide();}else{ $('#modelarea').show(); }"
		
		.Write "}"
		End If
		.Write "function CheckForm()"
		.Write "{ "
		.Write "   if ($('#FieldName').val()==''||$('#FieldName').val().length<=1)"
		.Write "    {"
		.Write "     top.$.dialog.alert('请输入字段名称!',function(){$('#FieldName').focus();});"
		.Write "     return false;"
		.Write "    }"
		.Write "   if ($('#Title').val()=='')"
		.Write "    {"
		.Write "     top.$.dialog.alert('请输入字段标题!',function(){$('#Title').focus();});"
		.Write "     return false;"
		.Write "    }"
		.write "   if ($('#FieldType')[0]!=undefined){"
		.Write "   if ($('#FieldType option:selected').val()==14){"
		.Write "      if ($('#BChannelID option:selected').val()=='0'){  top.$.dialog.alert('请选择要绑定的模型!',function(){ $('#BChannelID').focus(); }); return false; } "
		.Write "   }"
		.write "}"
		.Write "    $('#FieldForm').submit();"
		.Write "}"
		.Write "</Script>"
		End With
		End Sub
		 
		 Sub FieldAddSave()
		 Dim FieldRS,ColumnType
		 FieldName = Trim(KS.G("FieldName"))
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 FieldType = KS.G("FieldType")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.ChkClng(KS.G("ShowOnForm"))
		 ShowOnUserForm=KS.ChkClng(KS.G("ShowOnUserForm"))
		 ShowOnClubForm=KS.ChkClng(KS.G("ShowOnClubForm"))
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 Width=KS.G("Width")
		 MaxLength=KS.ChkClng(KS.G("MaxLength"))
		 AllowFileExt=KS.G("AllowFileExt")
		 EditorType=KS.G("EditorType")
		 ShowUnit  =KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions=KS.G("UnitOptions")
		 ParentFieldName=KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"

		 If FieldName = "" Then Call KS.AlertHistory("请输入字段名称!", -1): Exit Sub
		 If KS.HasChinese(FieldName)  Then Call KS.AlertHistory("字段名称不能含有中文!", -1): Exit Sub
		 If Len(FieldName)<=3 Then Call KS.AlertHistory("字段名称长度必须大于3!", -1): Exit Sub
		 If Ucase(Left(FieldName,3))<>"KS_" Then Call KS.AlertHistory("字段名称格式有误，必须以""KS_开头""!", -1): Exit Sub
		 If Title="" Then Call KS.AlertHistory("字段标题必须输入!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" and lcase(DefaultValue)<>"now" and lcase(DefaultValue)<>"date" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认格式不正确，请输入正确的Email!",-1):Exit Sub
		 Select Case FieldType
		   Case 1,3,6,7,8,9,11
		     If MaxLength=0 Then
		     ColumnType="nvarchar(255)"
			 Else
		     ColumnType="nvarchar(" &MaxLength&")"
			 End If
		   Case 13
		    ColumnType="tinyint default 0"
		   Case 2,10
		     ColumnType="ntext"
		   Case 5
		     ColumnType="datetime"
		   Case 4,14
		     ColumnType="int"
		   Case 12
		     ColumnType="float"
		   Case else
		     Exit Sub
		 End Select
		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select top 1 * From [KS_Field] Where FieldName='" & FieldName & "' And ChannelID=" & KS.G("ChannelID")
		 FieldRS.Open FieldSql, conn, 3, 3
		 If FieldRS.EOF And FieldRS.BOF Then
		  FieldRS.AddNew
		  FieldRS("FieldName") = FieldName
		  FieldRS("ChannelID") = KS.G("ChannelID")
		  FieldRS("Title") = Title
		  FieldRS("Tips") = Tips
		  FieldRS("FieldType") = FieldType
		  If FieldType=14 Then
		   If  KS.ChkClng(KS.S("BChannelID"))=0 Then   Call KS.AlertHistory("请选择要调用的模型!", -1)
		   FieldRS("DefaultValue") = KS.ChkClng(KS.S("BChannelID"))
		  Else
		   FieldRS("DefaultValue") = DefaultValue
		  End If
		  FieldRS("MustFillTF") = MustFillTF
		  FieldRS("FieldType") = FieldType
		  FieldRS("ShowUnit")=ShowUnit
		  FieldRS("UnitOptions")=UnitOptions
		  FieldRS("ShowOnForm") = ShowOnForm
		  FieldRS("ShowOnUserForm")=ShowOnUserForm
		  FieldRS("ShowOnClubForm")=ShowOnClubForm
		  FieldRS("Options") = Options
		  FieldRS("OrderID")=KS.ChkClng(KS.G("OrderID"))
		  FieldRS("AllowFileExt")=KS.G("AllowFileExt")
		  FieldRS("MaxFileSize")=KS.ChkClng(KS.G("MaxFileSize"))
		  FieldRS("Width")=KS.ChkClng(KS.G("Width"))
		  FieldRS("Height")=KS.ChkClng(KS.G("Height"))
		  FieldRS("MaxLength")=MaxLength
		  FieldRS("EditorType")=EditorType
		  FieldRS("ParentFieldName")=ParentFieldName
		  FieldRS("GroupID")=KS.ChkClng(KS.G("GroupID"))
		  FieldRS.Update
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If FieldType=14 then
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&"_ChannelID "&ColumnType&"")
		  End If
		  If ShowUnit=1 Then  '增加单位字段
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		  
		  on error resume next
		  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5 Then
		  KS.ConnItem.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If ShowUnit=1 Then  '增加单位字段
		  KS.ConnItem.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		  if err then err.clear
		  
		  End If
		   Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
		 Response.Write ("<Script>top.$.dialog.confirm('字段增加成功,继续添加吗?',function() { location.href='System/KS.Field.asp?GroupID=" & KS.ChkClng(KS.G("GroupID"))&"&ChannelID=" & ChannelID& "&Action=Add';} ,function(){location.href='System/KS.Field.asp?ChannelID=" & ChannelID&"&Page="&Page &"&GroupID=" & KS.S("GroupID")&"';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=模型管理 >> <font color=#ff0000>模型字段管理</font>&ButtonSymbol=Disabled';});</script>")
		 Else
		   Call KS.AlertHistory("数据库中已存在该字段名称!", -1)
		   Exit Sub
		 End If
		 FieldRS.Close
		 End Sub
		 
		 Sub FieldEditSave()
		 With Response
		 ID = KS.G("ID")
		 FieldName = Trim(KS.G("FieldName"))
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.ChkClng(KS.G("MustFillTF"))
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.ChkClng(KS.G("ShowOnForm"))
		 ShowOnUserForm=KS.ChkClng(KS.G("ShowOnUserForm"))
		 ShowOnClubForm=KS.ChkClng(KS.G("ShowOnClubForm"))
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 OrderID   =KS.G("OrderID")
		 EditorType=KS.G("EditorType")
		 ShowUnit  =KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions=KS.G("UnitOptions")
		 ParentFieldName=KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"
		 
		 If Title="" Then Call KS.AlertHistory("字段标题必须输入!", -1): Exit Sub
		' If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" and lcase(DefaultValue)<>"now" and lcase(DefaultValue)<>"date" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		
		
		 If Not IsNumeric(OrderID) Then OrderID=0
		 
		 '修改字段长度
		 if (FieldType=1 or FieldType=3 or FieldType=6 or FieldType=7 or FieldType=8 or FieldType=9 or FieldType=11) then
		     Dim ColumnType
			 If KS.ChkClng(KS.G("MaxLength"))=0 Then
		     ColumnType="nvarchar(255)"
			 Else
		     ColumnType="nvarchar(" &KS.ChkClng(KS.G("MaxLength"))&")"
			 End If
			 on error resume next
			 Conn.Execute("Alter Table "&TableName&" Alter Column "&FieldName&" "&ColumnType&"")
			 if err then err.clear
		 end if

		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		  FieldSql = "Select top 1 * From [KS_Field] Where FieldID=" & ID 
		  FieldRS.Open FieldSql, conn, 1, 3
		  FieldRS("ChannelID") = KS.G("ChannelID")
		  FieldRS("Title") = Title
		  FieldRS("Tips") = Tips
		  FieldRS("DefaultValue") = DefaultValue
		  FieldRS("MustFillTF") = MustFillTF
		  If FieldRS("FieldType")=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		'  If FieldRS("FieldType")=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub

		 ' FieldRS("FieldType") = FieldType
		 ' FieldRS("ShowUnit")=ShowUnit
		  'FieldRS("ParentFieldName")=ParentFieldName
		  FieldRS("UnitOptions")=UnitOptions
		  FieldRS("ShowOnForm") = ShowOnForm
		  FieldRS("ShowOnUserForm")=ShowOnUserForm
		  FieldRS("ShowOnClubForm")=ShowOnClubForm
		  FieldRS("Options") = Options
		  FieldRS("OrderID")=OrderID
		  FieldRS("AllowFileExt")=KS.G("AllowFileExt")
		  FieldRS("MaxFileSize")=KS.ChkClng(KS.G("MaxFileSize"))
		  FieldRS("Width")=KS.ChkClng(KS.G("Width"))
		  FieldRS("Height")=KS.ChkClng(KS.G("Height"))
		  FieldRS("MaxLength")=KS.ChkClng(KS.G("MaxLength"))
		  FieldRS("EditorType")=EditorType
		  FieldRS("GroupID")=KS.ChkClng(KS.G("GroupID"))
		  FieldRS.Update
		  FieldRS.Close
		  on error resume next
	   	  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5 Then
			KS.ConnItem.Execute("Update KS_FieldItem Set FieldTitle='" & Title & "',OrderID=" & OrderID &" Where FieldID=" & ID)
          End If
		 .Write ("<form name=""split"" action=""../Post.Asp"" method=""GET"" target=""BottomFrame"">")
		 .Write ("<input type=""hidden"" name=""OpStr"" value=""模型管理 >> <font color=red>模型字段管理</font>"">")
		 .Write ("<input type=""hidden"" name=""ButtonSymbol"" value=""Disabled""></form>")
		 .Write ("<script language=""JavaScript"">document.split.submit();</script>")
		  Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
		  
		  KS.Die "<script> $.dialog.alert('字段修改成功!', function (){ location.href='" & KS.Setting(3) & KS.Setting(89) & "system/KS.Field.asp?ChannelID=" & ChannelID&"&Page=" & Page &"&GroupID=" & KS.S("GroupID") &"'; });</script>"
		 End With
		 End Sub
		 
		 Sub FieldDel()
		    On error resume Next
			Dim ID:ID = KS.FilterIds(KS.G("ID"))
			If ID="" Then KS.AlertHintScript "没有选择字段!"
			Dim DelId:DelId=""
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "Select FieldName,FieldType,ShowUnit,FieldID From KS_Field Where FieldID IN(" & ID & ")",Conn,1,1
			Do While Not RSObj.Eof 
			  If left(Lcase(RSObj(0)),3)<>"ks_" Then
			  Else
			      ID=RSObj("FieldID")
				  If DelID="" Then DelID=ID Else DelID=DelID &"," & ID
			      Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
				  If RSObj("ShowUnit")="1" Then
			      Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_Unit")
				  End if
			   	  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5  Then
					  KS.ConnItem.Execute("Delete From KS_FieldItem Where FieldID IN(" & ID & ")")
					  KS.ConnItem.Execute("Delete From KS_FieldRules Where FieldID IN(" & ID & ")")
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
					  If RSObj("ShowUnit")="1" Then
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_Unit")
					  End if
					  If RSObj("FieldType")=14 Then
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_ChannelID")
					  End If
				  End If
			  End If
			  RSObj.MoveNext
			Loop
			RSObj.Close:Set RSObj=Nothing
			if DelID<>"" Then Conn.Execute("Delete From KS_Field Where FieldID IN(" & DelID & ")")
			Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
			Response.Redirect "KS.Field.asp?ChannelID=" & ChannelID &"&Page=" & Page
		 End Sub
		 
		 Sub FieldOrder()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim OrderID:OrderID=KS.G("OrderID")
			  Dim I,FieldIDArr,OrderIDArr
			  FieldIDArr=Split(FieldID,",")
			  OrderIDArr=Split(OrderID,",")
			  For I=0 To Ubound(FieldIDArr)
			   Conn.Execute("update KS_Field Set OrderID=" & OrderIDArr(i) &" where FieldID=" & FieldIDArr(I))
			   on error resume next
			   If KS.C_S(ChannelID,6)=1 Then
				KS.ConnItem.Execute("Update KS_FieldItem Set OrderID=" & OrderIDArr(i) &" Where FieldID=" & FieldIDArr(I))
			   End If
			  Next
			  Call KS.CreateFieldXML(ChannelID,"") '生成xml缓存
			  ks.die "<script language=JavaScript>$.dialog.alert('批量保存字段排序成功！',function(){location.href='system/KS.Field.asp?ChannelID=" & channelid &"';});</script>"
			  
		 End Sub
End Class
%> 

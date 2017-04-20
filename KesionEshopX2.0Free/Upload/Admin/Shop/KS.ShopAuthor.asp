<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
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
Set KSCls = New Admin_Author
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Author
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,ChannelID,ItemName1,FlagName,Flag1Name,RS
		Private OriginName, ID, Sex, Birthday, Telphone, UnitName, UnitAddress, Zip, Email, QQ, HomePage, Note, OriginType
		
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		    CurrentPage = KS.ChkClng(KS.G("page"))
		    If CurrentPage=0 Then CurrentPage=1
			   ChannelID=KS.ChkClng(KS.G("ChannelID"))
             With KS
		 	    .echo "<!DOCTYPE html><html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo "<title>作者管理</title>"
				.echo "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
				.echo "<script language='JavaScript'>"
				.echo "var Page='" & CurrentPage & "';"
				.echo "var ChannelID=" & ChannelID & ";"
				.echo "</script>"
				.echo "<script src=""../../KS_Inc/Jquery.js""></script>" & vbCrLf
				.echo "<script src=""../../KS_Inc/common.js""></script>" & vbCrLf
				.echo "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>"
             Action=KS.G("Action")
			 
			 If ChannelID=0 Then
			   If Not KS.ReturnPowerResult(ChannelID, "KMST10016") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			   End iF
			 Else
				If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "20003") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF
             End if
			 
			 Page=KS.G("Page")
			 Select Case KS.C_S(ChannelID,6)
			  Case 0
			   ItemName1="作者":FlagName="作者姓名":Flag1Name="作者性别"
			  Case 3
			   ItemName1="开发商":FlagName="开发商":Flag1Name="作者性别"
			  Case 5
			   ItemName1="厂商":FlagName="厂商名称":Flag1Name="联系电话"
			 End Select
		With KS	 
		.echo "<body topmargin='0' leftmargin='0' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		.echo "<ul id='menu_top'>"
		 If Action="Add" Or Action="Edit" Then
		 .echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon save'></i>确定保存</span></li>"
		 .echo "<li class='parent' onclick=""location.href='?channelid=" & channelid &"';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消返回</span></li>"
		 Else
		.echo "<li class='parent' onClick=""AuthorAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>新增"&ItemName1 &"</span></li>"
		.echo "<li class='parent' onclick='AuthorControl(1);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon write'></i>修改"&ItemName1 &"</span></li>"
		.echo "<li class='parent' onclick='AuthorControl(2);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon delete'></i>删除"&ItemName1 &"</span></li>"
		End If
		.echo "</ul>"
	   End With

			 
			 Select Case Action
			  Case "Add"
			    Call AddOrEdit("Add")
			  Case "Edit"
			    Call AddOrEdit("Edit")
			  Case "Del"
			    Call AuthorDel()
			  Case "AddSave"
			    Call AuthorAddSave()
			  Case "EditSave"
			    Call AuthorEditSave()
			  Case Else
			   Call ShowMain()
			 End Select
			.echo "</body>"
			.echo "</html>"
			End With
		End Sub
		
		Sub ShowMain()
	%>
	   <script language="javascript">
	   	function set(v)
		{
			 if (v==1)
				 AuthorControl(1);
			 else if (v==2)
				 AuthorControl(2);
		}
		var box='';
		function AuthorAdd()
		{
			location.href='KS.ShopAuthor.asp?ChannelID='+ChannelID+'&Action=Add';
			$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("商城系统 >> <font color=red>新增厂商</font>")+'&ButtonSymbol=GO';

		}
		function EditAuthor(id)
		{
			location.href='KS.ShopAuthor.asp?ChannelID='+ChannelID+'&Action=Edit&ID='+id;
			$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("商城系统 >> <font color=red>编辑厂商</font>")+'&ButtonSymbol=GOSave';

			
		}
		function DelAuthor(id)
		{
		 if (confirm('真的要删除该作者吗?'))
		  location="KS.ShopAuthor.asp?ChannelID="+ChannelID+"&Action=Del&Page="+Page+"&id="+id;
		}
		function AuthorControl(op)
		{ var alertmsg='';
			var ids=get_Ids(document.myform);
			if (ids!='')
			 {
			   if (op==1)
				{
				if (ids.indexOf(',')==-1) 
					EditAuthor(ids)
				  else top.$.dialog.alertalert('一次只能编辑一个作者!')
				}	
			  else if (op==2)    
			  DelAuthor(ids);
			 }
			else 
			 {
			 if (op==1)
			  alertmsg="编辑";
			 else if(op==2)
			  alertmsg="删除"; 
			 else
			  {
			  WindowReload();
			  alertmsg="操作" 
			  }
			  top.$.dialog.alert('请选择要'+alertmsg+'的作者');
			  }
		}
		function GetKeyDown()
		{ event.returnValue=false;
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : Select(0);break;
			 case 78 : event.keyCode=0;event.returnValue=false;AuthorAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;AuthorControl(1);break;
			 case 68 : AuthorControl(2);break;
		   }	
		else	
		 if (event.keyCode==46)AuthorControl(2);
		}
	   </script>
	<%
	   With KS
		.echo "</head>"
		
		.echo "<div class='pageCont2'><table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		.echo ("<form name='myform' method='Post' action='?channelid="& channelid & "'>")
	    .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
		.echo "   <tr class='sort'>"
		.echo "          <td width='35' align='center'>选择</td>"
		.echo "          <td align='center'>" & FlagName &"</td>"
		.echo "          <td><div align='center'>" & Flag1Name & "</td>"
		.echo "          <td align='center'>电子邮箱</td>"
		.echo "          <td align='center'>添加时间</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
				   SqlStr = "SELECT * FROM [KS_Origin] Where ChannelID="& ChannelID& " AND OriginType=1 order by AddDate desc"
				   RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				   .echo "   <tr>"
		           .echo "          <td class=""splittd"" width='35' align='center' colspan='10'>还没有添加厂商！</td>"
				   .echo "</tr>"
				           totalPut=0
				 Else
					        totalPut = RS.RecordCount
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
			End If
		    .echo "</table>"
			.echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .echo ("<tr><td width='180' class='pt10'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	        .echo ("</td>")
	        .echo ("<td class='pt10'><select onchange='set(this.value)' name='setattribute' style='height:30px;'><option value=0>快速选项...</option><option value='1'>执行编辑</option><option value='2'>执行删除</option></select></td>")
	        .echo ("</form><td align='right'>")
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .echo ("</td></tr></table></div>")
		End With
		End Sub
		Sub showContent()
		  With KS
			Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
			.echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		    .echo "  <td  class='splittd' height='19'><span AuthorID='" & RS("ID") & "'  onDblClick=""EditAuthor(this.AuthorID)"">"
		    .echo "    <span style='cursor:default;'>"& RS("OriginName") & "</span></span> </td>"
		   
		   If ChannelID=5 Then
		   .echo " <td  class='splittd' align='center'>" & RS("Telphone") & " </td>"
		   Else
		   .echo " <td  class='splittd' align='center'>" & RS("Sex") & " </td>"
		   End If
		   .echo " <td  class='splittd' align='center'>&nbsp;" & RS("Email") & "</td>"
		   .echo " <td  class='splittd' align='center'>" & RS("AddDate") & " </td>"
		   .echo "</tr>"
							  I = I + 1
								If I >= MaxPerPage Then Exit Do
							   RS.MoveNext
							   Loop
								RS.Close
						  
		   End With
		 End Sub
		 
		 Sub AddOrEdit(OpType)
		  With KS
		   Dim RS, OriginSql
		   ID = Request("ID")
		
		  Action="AddSave"
		  Sex="男"
		  If OpType = "Edit" Then
			 Set RS = Server.CreateObject("ADODB.RECORDSET")
			 OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
			 RS.Open OriginSql, conn, 1, 1
			 If Not RS.EOF Then
			 OriginName = Trim(RS("OriginName"))
			 ChannelID = RS("ChannelID")
			 Sex = Trim(RS("Sex"))
			 Birthday = Trim(RS("Birthday"))
			 Telphone = Trim(RS("Telphone"))
			 UnitName = Trim(RS("UnitName"))
			 UnitAddress = Trim(RS("UnitAddress"))
			 Zip = Trim(RS("Zip"))
			 Email = Trim(RS("Email"))
			 QQ = Trim(RS("QQ"))
			 HomePage = Trim(RS("HomePage"))
			 Note = Trim(RS("Note"))
			 Action="EditSave"
		    End If
		
        End If
		.echo "<div class='pageCont2'><form  action='KS.ShopAuthor.asp?ID=" & ID &"&page=" & Page & "' method='post' name='AuthorForm' onsubmit='return(CheckForm())'>"
		.echo "     <input type='hidden' value='" & ChannelID & "' name='ChannelID'>"
		.echo "<dl class='dtable'>"
		.echo "    <dd>"
		.echo "      <div>" & FlagName &"：</div>"
		.echo "      <input name='OriginName' value='" & OriginName & "' size='50' type='text' id='OriginName' class='textbox'><span>如：厦门科汛软件有限公司</span>"
		.echo "    </dd>"
		
		if ChannelID<>5 then
		.echo "    <dd>"
		.echo "      <div>作者性别：</div>"
			 If Sex = "男" Then
			   .echo "<input name='Sex' type='radio' value='男' Checked> 男"
			  Else
			   .echo "<input name='Sex' type='radio' value='男'> 男"
			  End If
			  If Sex = "女" Then
			   .echo "<input name='Sex' type='radio' value='女' Checked> 女"
			  Else
			   .echo "<input name='Sex' type='radio' value='女'> 女"
			  End If
		.echo "       </dd>"
		.echo "    <dd>"
		.echo "        <div>出生日期：</div>"
		.echo "        <input name='Birthday' onclick=""WdatePicker({dateFmt:'yyyy-MM-dd'});"" type='text' id='Birthday' value='" & Birthday & "' class='textbox' size='15' readonly>"
		.echo "        <a href='#' onclick=""WdatePicker({el:$('input[name=Birthday]').get(0),dateFmt:'yyyy-MM-dd'});return false;""><img src='../images/date.gif' border='0' align='absmiddle' title='选择日期'></a>"
		.echo "        </dd>"
	   end if
		.echo "    <dd>"
		.echo "      <div>联系电话：</div>"
		.echo "     <input name='Telphone' type='text' value='" & Telphone & "' id='Telphone' class='textbox' size='50'><span>请输入厂商的联系电话</span>"
		.echo "    </dd>"
		
		if ChannelID<>5 then
		.echo "    <dd>"
		.echo "      <div>单位名称：</div>"
		.echo "      <input name='UnitName' type='text' id='UnitName' value='" & UnitName & "' class='textbox' size='50'>"
		.echo "   </dd>"
		end if 
		.echo "    <dd>"
		.echo "      <div>单位地址：</div>"
		.echo "      <input name='UnitAddress' type='text' size='50' id='UnitAddress' value='" & UnitAddress & "' class='textbox' size='50'><span>请输入厂商的联系地址</span>"
		.echo "    </dd>"
		.echo "    <dd><div>邮政编码：</div><input name='Zip' size='50' type='text' id='Zip' value='" & Zip & "' class='textbox'>"
		.echo "    </dd>"
		.echo "    <dd>"
		.echo "     <div>电子邮箱:</div>"
		.echo "     <input name='Email' type='text' id='Email' size='50' value='" & Email & "' class='textbox'><span>请输入厂商的电子邮件</span>"
		.echo "    </dd>"
		.echo "    <dd>"
		.echo "      <div>联系QQ：</div>"
		.echo "     <input name='QQ' type='text' id='QQ' size='50' value='" & QQ & "' class='textbox'>"
		.echo "    </dd>"
		.echo "    <dd>"
		.echo "       <div>主页地址:</div>"
		.echo "       <input name='HomePage' type='text' size='50' id='HomePage' value='" & HomePage & "' class='textbox' value='http://'>"
		.echo "    </dd>"
		.echo "    <dd>"
		.echo "      <div>备注说明：</div>"
		.echo "     <textarea name='Note' cols='50' rows='6' id='Note' class='textbox'>" & Note & "</textarea>"
		.echo "      </dd>"
		.echo "    <input type='hidden' value='" & Action & "' name='Action'>"
		.echo "    <input type='hidden' name='OriginType' value='1'>"
		.echo "</dl>"
		.echo "  </form></div>"

	
		.echo "<Script Language='javascript'>"
		.echo "function CheckForm()"
		.echo "{ var form=document.AuthorForm;"
		.echo "   if (form.OriginName.value=='')"
		.echo "    {"
		.echo "     top.$.dialog.alert('请输入" & FlagName &"!',function(){"
		.echo "     form.OriginName.focus();});"
		.echo "     return false;"
		.echo "    }"
		.echo "    if ((form.Zip.value!="""")&&((form.Zip.value.length>6)||(!is_number(form.Zip.value))))"
		.echo "    {"
		.echo "     top.$.dialog.alert('非法邮政编码!',function(){" &vbcrlf 
		.echo "     form.Zip.focus();});"
		.echo "     return false;"
		.echo "    }"
		.echo "    if (form.Email.value!="""")"
		.echo "    if(is_email(form.Email.value)==false)"
		.echo "    { top.$.dialog.alert('非法电子邮箱!',function(){" &vbcrlf
		.echo "     form.Email.focus();});"
		.echo "     return false;"
		.echo "    }"
		.echo "    form.submit();"
		.echo "}"
		.echo "</Script>"
        End With
		 End Sub
		 
		 Sub AuthorAddSave()
		 OriginName = Trim(Request.Form("OriginName"))
		 Sex = Trim(Request.Form("Sex"))
		 Birthday = Request.Form("Birthday")
		 Telphone = Trim(Request.Form("Telphone"))
		 UnitName = Trim(Request.Form("UnitName"))
		 UnitAddress = Trim(Request.Form("UnitAddress"))
		 Zip = Trim(Request.Form("Zip"))
		 Email = Trim(Request.Form("Email"))
		 QQ = Trim(Request.Form("QQ"))
		 HomePage = Trim(Request.Form("HomePage"))
		 Note = Trim(Request.Form("Note"))
		 OriginType = CInt(Request.Form("OriginType"))
		 
		 If OriginName = "" Then Call KS.AlertHistory("请输入作者姓名!", -1):Set KS = Nothing
		 Dim RS:Set RS = Server.CreateObject("ADODB.RECORDSET")
		 Dim OriginSQL:OriginSql = "Select * From [KS_Origin] Where OriginName='" & OriginName & "' And ChannelID=" & KS.G("ChannelID") & " And OriginType=1"
		 RS.Open OriginSql, conn, 3, 3
		 If RS.EOF And RS.BOF Then
		  RS.AddNew
		  RS("ID") = Year(Now) & Month(Now) & Day(Now) & KS.MakeRandom(5)
		  RS("OriginName") = OriginName
		  RS("ChannelID") = KS.G("ChannelID")
		  RS("Sex") = Sex
		   If Birthday <> "" Then
		  RS("Birthday") = Birthday
		   End If
		  RS("Telphone") = Telphone
		  RS("UnitName") = UnitName
		  RS("UnitAddress") = UnitAddress
		  RS("Zip") = Zip
		  RS("Email") = Email
		  RS("QQ") = QQ
		  RS("HomePage") = HomePage
		  RS("Note") = Note
		  RS("OriginType") = OriginType
		  RS("AddDate") = Now()
		  RS.Update
		  Set conn = Nothing
		  Call KS.ConfirmDoFun("厂商增加成功,继续添加吗?","location.href='shop/KS.ShopAuthor.asp?ChannelID="& ChannelID& "&Action=Add';","location.href='shop/KS.ShopAuthor.asp?ChannelID="& ChannelID& "';")
		 Else
		 Call KS.AlertHistory("数据库中已存在该作者!", -1)
		 Set KS = Nothing:.End
		 End If
		 RS.Close
		 End Sub
		 
		 Sub AuthorEditSave()
		 With Response
		 ID = Request("ID")
		 OriginName = Trim(Request.Form("OriginName"))
		 Sex = Trim(Request.Form("Sex"))
		 Birthday = Request.Form("Birthday")
		 Telphone = Trim(Request.Form("Telphone"))
		 UnitName = Trim(Request.Form("UnitName"))
		 UnitAddress = Trim(Request.Form("UnitAddress"))
		 Zip = Trim(Request.Form("Zip"))
		 Email = Trim(Request.Form("Email"))
		 QQ = Trim(Request.Form("QQ"))
		 HomePage = Trim(Request.Form("HomePage"))
		 Note = Trim(Request.Form("Note"))
		 OriginType = CInt(Request.Form("OriginType"))

		 If OriginName = "" Then Call KS.AlertHistory("请输入作者姓名!", -1)
		 Dim RS:Set RS = Server.CreateObject("ADODB.RECORDSET")
		 Dim OriginSQL:OriginSql = "Select top 1 * From [KS_Origin] Where ID='" & ID & "'"
		  RS.Open OriginSql, conn, 1, 3
		  RS("OriginName") = OriginName
		  RS("ChannelID") = KS.G("ChannelID")
		  RS("Sex") = Sex
		   If Birthday <> "" Then
		  RS("Birthday") = Birthday
		   End If
		  RS("Telphone") = Telphone
		  RS("UnitName") = UnitName
		  RS("UnitAddress") = UnitAddress
		  RS("Zip") = Zip
		  RS("Email") = Email
		  RS("QQ") = QQ
		  RS("HomePage") = HomePage
		  RS("Note") = Note
		  RS.Update
		  RS.Close
		  Set RS=Nothing
		  
		   KS.Echo ("<Script> top.$.dialog.alert('作者修改成功!',function(){ location.href='shop/KS.ShopAuthor.asp?channelid=5';});</script>")
		   Set conn = Nothing
		  End With
		 End Sub
		 
		 Sub AuthorDel()
			Dim ID:ID = KS.G("ID")
			ID = Replace(ID, ",", "','")
			ID = "'" & ID & "'"
			conn.Execute ("Delete From KS_Origin Where ID IN(" & ID & ")")
			Response.Redirect "KS.ShopAuthor.asp?ChannelID=" & ChannelID &"&Page=" & Page
		 End Sub
End Class
%> 

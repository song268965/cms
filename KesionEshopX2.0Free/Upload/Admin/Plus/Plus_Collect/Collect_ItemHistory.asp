<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Collect_ItemHistory
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemHistory
        Private KS
		Private KMCObj
		Private ConnItem
		Private i, totalPut, CurrentPage, SqlStr
		Private Rs, Sql, SqlItem, RSObj, Action, FoundErr, ErrMsg
		Private HistoryID, ItemID, ChannelID, ClassID, SpecialID, ArticleID, Title, CollecDate, NewsUrl, Result
		Private Arr_History, Arr_ArticleID, i_Arr, Del, Flag
		Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		If Not KS.ReturnPowerResult(0, "M0100082") Then
		  Response.Write "<script src='../../../ks_inc/jquery.js'></script>"
		  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
		  Call KS.ReturnErr(1, "")
		End If
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		'response.write "channelid=" & channelid
        CurrentPage = KS.ChkCLng(Request("page"))
		If CurrentPage=0 Then CurrentPage=1
		FoundErr = False
		Action = Trim(Request("Action"))
		If FoundErr <> True Then
		   Call Main
		Else
		   Call KS.AlertHistory(ErrMsg,-1)
		End If
		End Sub
		Sub Main()
		    Dim HistoryID:HistoryID = Trim(KS.G("HistoryID"))
			Dim Action:Action=KS.G("Action")
			Dim Page:Page = KS.G("Page")
		    If Action = "del" Then
			  HistoryID = Replace(HistoryID, " ", "")
			  If KS.IsNul(HistoryID) Then KS.AlertHintScript "对不起,没有选择记录!"
			  ConnItem.Execute ("Delete From KS_History Where HistoryID In(" & HistoryID & ")")
			 Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			ElseIf Action="DelSucceed" Then
			  ConnItem.Execute ("Delete From KS_History  Where  Result=True")
			  Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			ElseIf Action="DelFailure" Then
			  ConnItem.Execute ("Delete From KS_History  Where  Result=False")
			  Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			ElseIf Action = "delall" Then
			  ConnItem.Execute ("Delete From KS_History")
			 Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			End If
			
		 Response.Write "<!DOCTYPE html><html>"
		 Response.Write "<head>"
		Response.Write "<title>采集系统</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../../Include/Admin_Style.css"">"
		Response.Write "<script language=""JavaScript"">"
		Response.Write "var Page='" & CurrentPage & "';"
		Response.Write "</script>"
		Response.Write "<script language=""JavaScript"" src=""../../../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../../KS_Inc/jquery.js""></script>"
		%>
		<script>
		function DelRecords()
		{
		 var SelectedFile=get_Ids(document.myform);
		 if (SelectedFile!='')
		  {
		   if (confirm('真的要删除选中的记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=del&Page="+Page+"&HistoryID="+SelectedFile;
		  }
		 else
		  alert('请选择要删除的记录!');
		  SelectedFile='';
		}
		function DelSucceed()
		{
		 if (confirm('真的要清除所有成功记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=DelSucceed&Page="+Page;
		}
		function DelFailure()
		{
		 if (confirm('真的要清除所有记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=DelFailure&Page="+Page;
		}
		function DelAllRecords()
		{
		 if (confirm('真的要清除所有记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=delall&Page="+Page;
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 68 : DelRecords('');break;
			 case 83 : DelSucceed('');break;
			 case 70 : DelFailure('');break;
			 case 89 : event.keyCode=0;event.returnValue=false;DelAllRecords('');break;
		   }	
		else	
		 if (event.keyCode==46) DelRecords('');
		}
		function CheckAll(form)
			{
			  for (var i=0;i<form.elements.length;i++)
				{
				var e = form.elements[i];
				if (e.Name != "chkAll")
				   e.checked = form.chkAll.checked;
				}
			}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body topmargin=""0"" leftmargin=""0"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		Response.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemModify.asp?channelid=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>新建项目</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon choose'></i>过滤设置</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon audit'></i>审核入库</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon num'></i>历史记录</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>自定义字段</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>回上一级</span></li><li></li>"

			Response.Write "<div id='go'><select OnChange=""location.href=this.value"" style='width:120px' name='id'>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "'>快速查找历史记录</option>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "'>查看全部记录</option>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "&Action=Succeed'>查看成功记录</option>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "&Action=Failure'>查看失败记录</option>"
			
			Response.Write " </select>"
			Response.Write "</div>"
			Response.Write ("</ul>")
            

									
		Set RSObj = Server.CreateObject("adodb.recordset")
		'SqlItem = "select * From KS_History Where ChannelID=" & ChannelID
		SqlItem = "select * From KS_History"
		If Action = "Succeed" Then
		   SqlItem = SqlItem & "  where Result=True"
		   Flag = "成 功 记 录"
		ElseIf Action = "Failure" Then
		   SqlItem = SqlItem & " where Result=False"
		   Flag = "失 败 记 录"
		Else
		   Flag = "所 有 记 录"
		End If
		Response.Write "<div class='pageCont2'>"
		Response.Write "<div class='tabTitle'>采集历史记录</div>"
		Response.Write "  <table border=""0"" cellspacing=""0"" width=""100%"" cellpadding=""0"">"
		Response.Write "     <tr>"
		Response.Write "      <td height=""22"" nowrap align=""center"" class=sort>选择</td>"
		Response.Write "      <td align=""center"" class=sort>标题</td>"
		Response.Write "      <td align=""center"" class=sort>项目名称</td>"
		Response.Write "      <td align=""center"" class=sort>所属系统</td>"
		Response.Write "      <td align=""center"" class=sort>(频道)栏目</td>"
		Response.Write "      <td align=""center"" class=sort>来源</td>"
		 Response.Write "     <td align=""center"" class=sort>结果</td>"
		 Response.Write "    </tr>"
		
		If Request("page") <> "" Then
			CurrentPage = CInt(Request("Page"))
		Else
			CurrentPage = 1
		End If
		SqlItem = SqlItem & " order by HistoryID DESC"
		RSObj.Open SqlItem, ConnItem, 1, 1
		 If Not RSObj.EOF Then
					totalPut = RSObj.RecordCount
				 	If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
                    End If
					Call showContent
		 Else
		  Response.Write "<tr><td class='splittd' align='center' colspan='10'>没有采集记录！</td></tr>"
		 End If
		
		   Response.Write "<tr><td colspan=8 height='25' class='operatingBox' style='text-align:left'>&nbsp;&nbsp;<strong>选择：</strong><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a>&nbsp;&nbsp;<input type='button' value='批量删除' class='button' onclick=""DelRecords()"">&nbsp;<input type='button' onclick=""DelAllRecords();"" value='删除全部记录' class='button'>&nbsp;<input type='button' onclick=""DelSucceed();"" value='删除所有成功记录' class='button'>&nbsp;<input type='button' onclick=""DelFailure();"" value='删除所有失败记录' class='button'></td></tr>"
		   Response.Write "</form>"
			Response.Write ("<tr><td height=""22"" colspan=""10"" align=""right"">")
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		   Response.Write ("</td></tr>")
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		Sub showContent()
		   Response.Write "<form name='myform' method='Post' action='?Page=" & CurrentPage & "&channelid=" & channelid & "'>"
		 Do While Not RSObj.EOF
			 Response.Write "  <tr height=""23"" class='list' id='u" & RSObj("HistoryID") &"' onclick=""chk_iddiv('" & RSObj("HistoryID") &"')"" onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
		    Response.Write "<td class='splittd' style='text-align:center'><input name=""id"" type=""checkbox""  onclick=""chk_iddiv('" & RSObj("HistoryID") &"')"" id='c" & RSObj("HistoryID") &"'  value='" & RSObj("HistoryID") &"'></td>"
			 Response.Write (" <td class='splittd' width=""435"" height=""18"">")
				Response.Write "<span  HistoryID='" & RSObj("HistoryID") & "'><i class='icon photo'></i>"
				  Response.Write "  <span style='cursor:default;'>" & KS.GotTopic(RSObj("Title"), 40) & "</span></span>"
			  Response.Write ("</td> ")
			  Response.Write ("<td class='splittd' width=""214"" align=""center"">" & KMCObj.Collect_ShowItem_Name(RSObj("ItemID"), ConnItem) & "</td>")
			  Response.Write ("<td class='splittd' width=""123"" align=""center"">" & KS.C_S(ChannelID,1) & "</td>")
			  Response.Write ("<td class='splittd' width=""120"" align=""center"">" & KMCObj.Collect_ShowClass_Name(RSObj("ChannelID"), RSObj("ClassID")) & "</td>")
			  Response.Write ("<td class='splittd' width=""126"" align=""center""><a href=""" & RSObj("NewsUrl") & """ target=""_blank"" title=""" & RSObj("NewsUrl") & """>点击访问</a></td>")
			  Response.Write (" <td class='splittd' width=""87"" align=""center"">")
			  If RSObj("Result") = True Then
				   Response.Write "<font color=red>成功</font>"
				ElseIf RSObj("Result") = False Then
				   Response.Write "<font color=red>失败</font>"
				Else
				   Response.Write "<font color=red>异常</font>"
				End If
			  Response.Write (" </td></tr> ")
				   i = i + 1
				   If i >=MaxPerPage Then
					  Exit Do
				   End If
				RSObj.MoveNext
		   Loop
		RSObj.Close
		Set RSObj = Nothing
		End Sub
End Class
%> 

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
Set KSCls = New Collect_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_Main
        Private KS
		Private KMCObj
		Private ConnItem,ChannelID
		'=================================================================================================
		Private i
		Private totalPut
		Private CurrentPage
		Private SqlStr
		Private RSObj
		Private MaxPerPage
		'=================================================================================================
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
			
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			ChannelID=KS.ChkClng(KS.G("ChannelID"))
			If ChannelID=0 Then ChannelID=1
			If Not KS.ReturnPowerResult(0, "M010008") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
			
			'response.write channelid
			Select Case  KS.G("Action")
			 Case "Del"
			    Dim ItemID:ItemID = KS.FilterIds(Replace(KS.G("ItemID"), " ", ""))
				ConnItem.Execute ("Delete From KS_CollectItem Where ItemID In(" & ItemID & ")")
				ConnItem.Execute ("Delete From KS_FieldRules Where ItemID In(" & ItemID & ")")
				ConnItem.Execute ("Delete From KS_Filters Where ItemID In(" & ItemID & ")")
				ConnItem.Execute ("Delete From KS_History Where ItemID In(" & ItemID & ")")
				Response.Write "<script>alert('恭喜,采集项目删除成功!');location.href='" & request.servervariables("http_referer") & "';</script>"
			Case "Paste"
			 Call ItemPaste()
			case "delhistory"
			    ItemID = KS.FilterIds(replace(KS.G("ItemID"), " ", ""))
				ConnItem.Execute ("Delete From KS_History Where ItemID In(" & ItemID & ")")
				Response.Write "<script>alert('恭喜,采集历史记录清除成功!');location.href='" & request.servervariables("http_referer") & "';</script>"
			Case else
			 Call ItemList()
			End Select
          End Sub
		  
		  Sub ItemList()
			Response.Write "<!DOCTYPE html><html>"
			Response.Write "<head>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write "<title>采集项目管理</title>"
			Response.Write "<link href=""../../Include/Admin_Style.css"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script language=""JavaScript"">" & vbCrLf
			Response.Write "var Page='" & CurrentPage & "';" & vbCrLf
			Response.Write "</script>" & vbCrLf
			Response.Write "<script language=""JavaScript"" src=""../../../KS_Inc/jquery.js""></script>"
			Response.Write "<script language=""JavaScript"" src=""../../../KS_Inc/common.js""></script>"
			%>
			<script>
			
			function CreateCollectItem()
			{location.href='Collect_ItemModify.asp?channelid=<%=ChannelID%>';
			}
			function EditCollectItem()
			{
				var SelectedFile=get_Ids(document.myform);
			 if (SelectedFile!='')
			   if (SelectedFile.indexOf(',')==-1)
			   location.href='Collect_ItemModify.asp?ItemID='+SelectedFile;
			   else
			   alert('一次只能够编辑一个采集项目!'); 
			 else
			  alert('请选择要编辑的采集项目!');
			  SelectedFile='';
			}
			function DelCollectItem()
			{
			 var SelectedFile=get_Ids(document.myform);
			 if (SelectedFile!='')
			  {
			   if (confirm('真的要删除选中的采集项目吗?'))
				location="?ChannelID=<%=ChannelID%>&Action=Del&Page="+Page+"&ItemID="+SelectedFile;
			  }
			 else
			  alert('请选择要删除的采集项目!');
			  SelectedFile='';
			}
			function SetCollectItemPro()
			{
			var SelectedFile=get_Ids(document.myform);
			 if (SelectedFile!='')
			   if (SelectedFile.indexOf(',')==-1)
			   location.href='Collect_ItemAttribute.asp?ItemID='+SelectedFile;
			   else
			   alert('一次只能够设置一个采集项目的属性!'); 
			 else
			  alert('请选择要设置属性的采集项目!');
			  SelectedFile='';
			}
			function TestCollectItem()
			{
				var SelectedFile=get_Ids(document.myform);
			 if (SelectedFile!='')
			   if (SelectedFile.indexOf(',')==-1)
			   location.href='Collect_ItemModify5.asp?ItemID='+SelectedFile;
			   else
			   alert('一次只能够测试一个采集项目!'); 
			 else
			  alert('请选择要测试的采集项目!');
			  SelectedFile='';
			}
			
			function Paste1(id)
			{ 
			      var Param='';
				  Param='?Action=Paste&ChannelID=<%=ChannelID%>&Page='+Page;
				  Param+='&ContentID='+id;
				  location.href='Collect_Main.asp'+Param;
			}
			
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false;CreateCollectItem();break;
				 case 69 : event.keyCode=0;event.returnValue=false;EditCollectItem();break;
				 case 80 : event.keyCode=0;event.returnValue=false;SetCollectItemPro();break;
				 case 84 : event.keyCode=0;event.returnValue=false;TestCollectItem();break;
				 case 68 : DelCollectItem('');break;
				 case 67 : 
				   event.keyCode=0;event.returnValue=false;Copy();
					break;
				 case 86 : 
				   if (top.CommonCopyCut.ChannelID==100<%=channelid%> && top.CommonCopyCut.PasteTypeID!=0)
				   { event.keyCode=0;event.returnValue=false;Paste();}
				   else
					return;
					break;
			   }	
			else	
			 if (event.keyCode==46) DelCollectItem();
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
			Response.Write "<body topmargin=""0"" leftmargin=""0"" onselectstart=""return false;"">"
		    Response.Write "<ul id='menu_top'>"
			Response.Write "<li class='parent' onclick='CreateCollectItem();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>新建项目</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon choose'></i>过滤设置</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_SeoSetting.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>SEO伪原创设置</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon audit'></i>审核入库</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon num'></i>历史记录</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>自定义字段</span></li>"
			Response.Write "<li disabled class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>回上一级</span></li>"
			Response.Write ("</ul>")
			Response.Write "<div class=""pageCont2"">"
			Response.Write "<div class=""tabTitle"">采集项目管理</div>"
			Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write "  <tr>"
			Response.Write "    <td height=""22"" class=""sort""><div align=""center"">选择</div></td>"
			Response.Write "    <td  class=""sort""><div align=""center"">项目名称</div></td>"
			Response.Write "    <td class=""sort""><div align=""center""><span>采集(站点)地址</span></div></td>"
			Response.Write "    <td align=""center"" class=""sort"">采回栏目</td>"
			Response.Write "    <td class=""sort""><div align=""center"">上次采集</div></td>"
			Response.Write "    <td align=""center"" class=""sort"">状态</td>"
			Response.Write "    <td align=""center"" class=""sort"">操作</td>"
			Response.Write "  </tr>"
			   Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   RSObj.Open "select ItemID,ItemName,WebName,ListStr,ListPageType,ListPageStr2,ListPageID1,ListPageID2,ListPageStr3,ChannelID,ClassID,SpecialID,Flag From KS_CollectItem order by ItemID DESC", ConnItem, 1, 1
					 If Not RSObj.EOF Then
						totalPut = RSObj.RecordCount
			
								If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
								End If
								Call showContent
			Else
		  Response.Write "<tr><td class='splittd' align='center' colspan='10'>没有采集项目！</td></tr>"
		 End If
					 Response.Write "<tr><td colspan=12 style='text-align:left' class='operatingBox'>&nbsp;&nbsp;<strong>选择：</strong><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> <input type='submit' onclick='this.form.action=""Collect_ItemCollection.asp?ChannelID=" & ChannelID&"&Action=Start&CollecType=1"";' value='批量采集选中项' class='button'> <input type='button' class='button' value=' 批量删除 ' onclick=""DelCollectItem()""/></td></tr>"
					 Response.Write "</form>"
					 Response.Write "<tr><td height='26' colspan='16' align='right'>"
					 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
					 Response.Write "</tr>"
			Response.Write ("</table>")
			Response.Write ("</div>")
			Response.Write ("</body>")
			Response.Write ("</html>")
			
			End Sub
			Sub showContent()
			   Dim Rs, ItemCollecDate
			   Dim ItemID, ItemName, WebName, ChannelID, ClassID, SpecialID, ListStr, ListPageType, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3, Flag, ListUrl
			     Response.Write "<form name='myform' method='Post' action='Collect_ItemCollection.asp'>"
					Do While Not RSObj.EOF
					
					 ItemID = RSObj("ItemID")
				  ItemName = RSObj("ItemName")
				  WebName = RSObj("WebName")
				  ChannelID = RSObj("ChannelID")
				  ClassID = RSObj("ClassID")
				  SpecialID = RSObj("SpecialID")
				  ListStr = RSObj("ListStr")
				  ListPageType = RSObj("ListPageType")
				  ListPageStr2 = RSObj("ListPageStr2")
				  ListPageID1 = RSObj("ListPageID1")
				  ListPageID2 = RSObj("ListPageID2")
				  ListPageStr3 = RSObj("ListPageStr3")
				  Flag = RSObj("Flag")
				  If ListPageType = 0 Or ListPageType = 1 Then
						ListUrl = ListStr
				  ElseIf ListPageType = 2 Then
						ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1))
				  ElseIf ListPageType = 3 Then
						If InStr(ListPageStr3, "|") > 0 Then
						ListUrl = Left(ListPageStr3, InStr(ListPageStr3, "|") - 1)
					 Else
						   ListUrl = ListPageStr3
					 End If
				  End If
				  
                      Response.Write "  <tr height=""23"" class='list' id='u" & ItemID &"' onclick=""chk_iddiv('" & ItemID &"')"" onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
		              Response.Write "<td class='splittd' style='text-align:center'><input name=""id"" type=""checkbox""  onclick=""chk_iddiv('" & ItemID &"')"" id='c" & ItemID &"'  value='" & ItemID &"'></td>"
		              Response.Write "  <td class='splittd' style='text-align:left'><span ondblclick='EditCollectItem()' ItemID='" & ItemID & "'>"
					  Response.Write " <span style='cursor:default;'>" & KS.Gottopic(ItemName,25) & "</span></span></td>"
					  Response.Write "  <td class='splittd' style='text-align:left'><a href='" & ListUrl & "' target='_blank'>" & WebName & "</a></td>"
					  Response.Write "  <td  class='splittd' align='center'>" & KMCObj.Collect_ShowClass_Name(ChannelID, ClassID) & "</td>"
					  Response.Write "  <td  class='splittd' align='center'>"
			
					  '上次采集
					  Set Rs = ConnItem.Execute("select Top 1 CollecDate From KS_History Where ItemID=" & ItemID & " Order by HistoryID desc")
					  If Not Rs.EOF Then
						ItemCollecDate = Rs("CollecDate")
					  Else
						ItemCollecDate = ""
					  End If
					  Set Rs = Nothing
					 If ItemCollecDate <> "" Then
						Response.Write ItemCollecDate
					 Else
						Response.Write "尚无记录"
					 End If
					 
					  Response.Write " </td>"
					  
					 Response.Write "  <td  class='splittd' align='center'>"
					  '状态
					  If Flag = True Then
								Response.Write "√"
					  Else
							 Response.Write "<font color=red>×</font>"
					  End If
					  Response.Write "</td>"
					  Response.Write "<td  class='splittd' style='text-align:center'><a href='Collect_ItemCollection.asp?ChannelID=" & ChannelID&"&ItemID=" & itemid & "&Action=Start&NewsFalseNum=0&ImagesNumAll=0' class='setA'>采集</a>|<a href='Collect_ItemModify.asp?ItemID=" & itemid & "' class='setA'>编辑</a>|<a href='?ChannelID=" & ChannelID & "&Action=Del&Page=" & CurrentPage & "&ItemID=" & itemid & "' onclick=""return(confirm('确认删除采集项目吗？'));"" class='setA'>删除</a>|<a href='Collect_ItemModify5.asp?ItemID=" & itemid & "' class='setA'>测试</a>|<a href='Collect_ItemAttribute.asp?ItemID=" & itemid & "' class='setA'>属性</a>|<a href='?action=delhistory&itemid=" & itemid&"' title='清空采集历史记录!' onclick=""return(confirm('清空采集历史记录可能导致,重复采集!确定删除吗?'))"" class='setA'>清空采集记录</a>|<a href=""javascript:Paste1(" & itemid &")"" class='setA'>克隆</a></td>"
					  Response.Write "</tr>"

					i = i + 1
					  If i >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  ConnItem.Close
				 End Sub
				 
				 
				 '粘贴
				 Sub ItemPaste()
		 Dim DisplayMode, Page
		 Dim PasteTypeID, DestFolderID, SourceFolderID, FolderID, ContentID
		  DisplayMode = KS.G("DisplayMode")
		  Page = KS.G("Page")
		  PasteTypeID = KS.G("PasteTypeID")
		  DestFolderID = KS.G("DestFolderID")
		  SourceFolderID = KS.G("SourceFolderID")
		  FolderID = KS.G("FolderID")
		  ContentID = KS.G("ContentID")
		  If PasteTypeID = "" Then PasteTypeID = 0
		  If DestFolderID = "" Then DestFolderID = "0"
		  If FolderID = "" Then
			 FolderID = "0"
		  End If
		  If ContentID = "" Then
			 ContentID = "0"
		  Else
			 ContentID = "'" & Replace(ContentID, ",", "','") & "'"
		  End If
		  If ContentID = "" Then
			Call KS.AlertHistory("参数传递出错!", 1)
			Set KS = Nothing
			Exit Sub
		  End If
		  
			Call PasteByCopy(SourceFolderID, DestFolderID, FolderID, ContentID)
		  Response.Write "<script>location.href='Collect_main.asp?ChannelID=" & KS.G("ChannelID") & "&Page=" & Page & "';</script>"
		End Sub
		
		
		
		'过程:PasteByCopy复制粘贴
		'参数:SourceFolderID--源目录,DestFolderID--目标目录,FolderID---被复制的目录,ContentID---被复制的文件
		Sub PasteByCopy(SourceFolderID, DestFolderID, FolderID, ContentID)
		       Dim ItemName,RS,RSA,I,NewItemID
			   
			ContentID=Replace(Replace(ContentID,"'",""),"""","")
			if instr(contentid,",") then call KS.AlertHistory("对不起，一次只能复制一个项目!",-1):exit sub
			Set RS=Server.CreateObject("Adodb.Recordset")
			RS.Open "Select top 1 * From KS_CollectItem Where ItemID=" & ContentID,ConnItem,1,1
			IF RS.Eof And RS.Bof Then
			Call KS.AlertHistory("操作失败!", 1)
			 Exit Sub
			Else
			   ItemName = Trim(RS("ItemName"))
			   
			   Set RSA=Server.CreateObject("ADODB.RECORDSET")
			   RSA.Open "Select top 1 * From KS_CollectItem",ConnItem,1,3
			   RSA.AddNew
			     For I=0 To RS.Fields.count-1
				   if lcase(RS.Fields(i).name)="itemid" then
				   elseif lcase(RS.Fields(i).name)="itemname" then
				    RSA("ItemName") = GetNewTitle(RS.Fields(i).value)
				   else
				    RSA(RS.Fields(i).name) = RS.Fields(i).Value
				   end if
				 Next
			   RSA.Update
			   RSA.MoveLast
			   NewItemID=RSA("ItemID")
			   RSA.Close
			   Set RSA=Nothing
			End IF
			RS.Close
			'复制自定义字段
		    If NewItemID<>"" Then
				RS.Open "Select * from KS_FieldRules Where ItemID=" & ContentID,ConnItem,1,1
				If Not RS.Eof Then
				   Set RSA=Server.CreateObject("ADODB.RECORDSET")
				   Do While Not RS.Eof 
					   RSA.Open "Select top 1 * From KS_FieldRules where 1=0",ConnItem,1,3
					   RSA.AddNew
						 For I=0 To RS.Fields.count-1
						   if lcase(RS.Fields(i).name)="id" then
						   elseif lcase(RS.Fields(i).name)="itemid" then
							RSA("ItemID")=NewItemID
						   else
							RSA(RS.Fields(i).name) = RS.Fields(i).Value
						   end if
						 Next
					   RSA.Update
				       RSA.Close
					  RS.MoveNext
					Loop
				   Set RSA=Nothing
				End If
			 RS.Close
			 End If	
			Set RS=Nothing
		End Sub
		Function GetNewTitle(OriTitle)
			Dim RSC
			On Error Resume Next
			Set RSC = Server.CreateObject("Adodb.RecordSet")
			
				 RSC.Open "Select * From KS_CollectItem Where ItemName Like '复制%" & OriTitle & "' Order By ItemID Desc", connItem, 1, 1
				 If Not RSC.EOF Then
					RSC.MoveFirst
					If RSC.RecordCount = 1 Then
					   RSC.Close
					   Set RSC = Nothing
					  GetNewTitle = "复制(1) " & OriTitle
					  Exit Function
					Else
					  GetNewTitle = "复制(" & CInt(Left(Split(RSC("ItemName"), "(")(1), 1)) + 1 & ") " & OriTitle
					End If
					 RSC.Close
					 Set RSC = Nothing
				 Else
				  RSC.Close
				  Set RSC = Nothing
				  GetNewTitle = "复制 " & OriTitle
				  Exit Function
				 End If			  
		End Function
End Class
%> 

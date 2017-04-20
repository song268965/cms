<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.FunctionCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New MainCls
KSCls.Kesion()
Set KSCls = Nothing

Class MainCls
        Private KS,Action
		Private I, totalPut, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 10
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMB10001") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			 End iF
			Action=KS.G("Action")
			With Response
			If Request("Action")<>"Add" And Request("Action")<>"Edit" And Request("Action")<>"Merger" Then
			.Write "<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../../KS_Inc/jquery.js""></script>"
            End If
			Select Case Action
			 Case "Add","Edit" Call GuestBoardAddOrEdit()
			 Case "Save" Call GuestBoardSave()
			 Case "Del" Call GuestBoardDel()
			 Case "DelTopic" Call DelTopic()
			 Case "Merger" Call Merger()
			 Case "doMerger" Call doMerger()
			 Case Else
			   Call MainList()
			End Select
		  End With
	    End Sub
		
		Sub MainList()
		 With Response
			%>
			<script language="JavaScript">
			function GuestBoardAdd(parentid)
			{
				location.href='KS.GuestBoard.asp?Action=Add&parentid='+parentid;
				window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=版块管理中心 >> <font color=red>添加新版块</font>&ButtonSymbol=GO';
			}
			function GuestBoadMerger(){
				location.href='KS.GuestBoard.asp?Action=Merger';
				window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=版块管理中心 >> <font color=red>版块合并</font>&ButtonSymbol=GO';
			}
			function EditGuestBoard(id)
			{
				location="KS.GuestBoard.asp?Action=Edit&Page="+Page+"&Flag=Edit&GuestBoardID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=版块管理中心 >> <font color=red>编辑版块</font>&ButtonSymbol=GoSave';
			}
			function DelGuestBoard(id)
			{
			if (confirm('如果有子版块将同时被删除,真的要执行删除操作吗?'))
			 location="KS.GuestBoard.asp?Action=Del&Page="+Page+"&GuestBoardid="+id;
			   SelectedFile='';
			}
			function DelTopic(id){
			if (confirm('执行此操作将清空该版面面的所有主题和回复,此操作不可逆请慎重操作!!!'))
			 location="KS.GuestBoard.asp?Action=DelTopic&Page="+Page+"&GuestBoardid="+id;
			   SelectedFile='';
			}
			
			</script>
			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"">"
			  .Write "<ul id='menu_top' class='menu_top_fixed'>"
			  .Write "<li class='parent' onclick=""GuestBoardAdd(0);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加版块分区</span></li>"
			  .Write "<li class='parent' onclick=""location.href='../plus/Plus_Tool/KS.Tools.asp#Club';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon write'></i>重计论坛数据</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoadMerger();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon merge'></i>论坛版面合并</span></li>"
			  .Write "</ul>"
			  .Write "<div class=""menu_top_fixed_height""></div>"
			

			.Write "<div class='pageCont2'><table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr class='sort'>"			
			.Write "          <td height=""25"" width=""70%"" class=""sort"" align=""center"">版块名称</td>"
			.Write "          <td class=""sort""><div align=""center"">版主</div></td>"
			.Write "          <td align=""center"" class=""sort"">帖子数</td>"
			.Write "          <td width=""50"" class=""sort"" align=""center"">排序</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
			 SqlStr = "SELECT * FROM KS_GuestBoard Where ParentID=0 order by orderID,id"
			 RSObj.Open SqlStr, Conn, 1, 1
			 If RSObj.EOF And RSObj.BOF Then
			 Else
						        totalPut = RSObj.RecordCount
								If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
								End If
								Call showContent
			End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table></div>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			  Dim RS,I
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr class='list'  onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					  .Write "  <td class='splittd' height='20'>&nbsp; <span GuestBoardID='" & RSObj("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><img src='../Images/Field.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & RSObj("BoardName") & "</span></span> "
					  .Write "  <span class='noshow'><a href='javascript:GuestBoardAdd(" & rsobj("id") & ")'>添加分版</a> | <a href='javascript:EditGuestBoard(" & rsobj("id") & ")'>修改</a> | <a href='javascript:DelGuestBoard(" & rsobj("id") & ")'>删除</a></span></td>"
					  .Write "  <td class='splittd' align='center'>&nbsp;" & RSObj("master") & "&nbsp;</td>"
					  .Write "  <td class='splittd' align='center'></td>"
					  .Write "  <td class='splittd' align='center'>" & RSOBJ("OrderID") &"</td>"
					  .Write "</tr>"
					  Set RS=Conn.Execute("Select ID,BoardName,master,todaynum,postnum,topicnum,orderid From KS_GuestBoard Where ParentID=" & RSObj("ID") & " Order by orderid")
					  Do While not rs.eof
					  .Write "<tr  class='list'  onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					  .Write "  <td class='splittd' height='20'> &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;|- <span GuestBoardID='" & RS("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><i class='icon folder'></i>"
					  .Write "    <span style='cursor:default;'>" & RS("BoardName") & "</span></span> "
					  .Write "  <span class='noshow'><a href='javascript:EditGuestBoard(" & rs("id") & ")'>修改</a> | <a href='javascript:DelGuestBoard(" & rs("id") & ")'>删除</a>  | <a href='javascript:DelTopic(" & rs("id") & ")'>清空</a></span></td>"
					  If KS.IsNul(RS("master")) Then
					  .Write "  <td class='splittd' align='center' style='color:#777'>&nbsp;无&nbsp;</td>"
					  Else
					  .Write "  <td class='splittd' align='center' style='color:#777'>&nbsp;" & RS("master") & "&nbsp;</td>"
					  End If
					  .Write "  <td class='splittd' align='center' style='color:#777'>主题:<font Color=red>" & RS("topicnum") & "</font> 总数:<font Color=red>" & RS("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RS("OrderID") &"</td>"
					  .Write "</tr>"
					  rs.movenext
					  loop
					  rs.close
					  
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  .Write "<tr><td height='26' colspan='5' align='right'>"
					  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
				End With
			    Set RS=Nothing
		  End Sub
			
		  '版块合并
		  Sub Merger()
				With Response
			    .Write "<!DOCTYPE HTML>"
			    .Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				.Write "<title>版块管理</title>"
				.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<script src=""../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
				.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script type=""text/javascript"">"
				.Write " function check(){"
				.Write " if ($('#boardid1').val()==0 || $('#boardid2').val()==0){"
				.Write "    alert('请选择源版面及目标版面!');return false;"
				.Write "  }"
				.Write " if ($('#boardid1').val()==$('#boardid2').val()){"
				.Write "    alert('源版面和目标版面不能相同!');return false;"
				.Write "  }"
				.Write "  return true;"
				.Write "}"
				.Write "</script>"
				.Write "</head>"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>论坛版面合并</div>"
				.Write "<br>"
				.Write " <form name='myform' id='myform' action='KS.GuestBoard.asp' method='post'>"
				.Write " <input type='hidden' name='Action' value='doMerger'/>"
				.Write "<dl class=""dtable"">"
				.Write " <dd><div>源版面：</div>"
				.Write " <select name='boardid1' id='boardid1'>"
				.Write "  <option value='0'>---请选择---</option>"
				 Call KS.LoadClubBoard()
			     Dim node,Xml,n,Str
			     Set Xml=Application(KS.SiteSN&"_ClubBoard")
			     for each node in xml.documentelement.selectnodes("row[@parentid=0]")
				      .Write ("<OPTGROUP label=&nbsp;+" & node.selectsinglenode("@boardname").text & " </OPTGROUP>")
					for each n in xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
					  .Write ("<option value='" & N.SelectSingleNode("@id").text & "'>&nbsp;|-" & n.selectsinglenode("@boardname").text &"</option>")
					next
				next
				
				.Write "</select><span class='tips'>源版块的帖子全部转入目标版块，同时删除源版块</span></dd>"
				.Write " <dd><div>目标版面：</div>"
				.Write " <select name='boardid2' id='boardid2'>"
			    .Write "  <option value='0'>---请选择---</option>"
				 for each node in xml.documentelement.selectnodes("row[@parentid=0]")
				      .Write ("<OPTGROUP label=&nbsp;+" & node.selectsinglenode("@boardname").text & " </OPTGROUP>")
					for each n in xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
					  .Write ("<option value='" & N.SelectSingleNode("@id").text & "'>&nbsp;|-" & n.selectsinglenode("@boardname").text &"</option>")
					next
				next
				
				.Write "</select></dd>"
				.Write " <dd><input type='submit' value='确定合并' onclick='return(check())' class='button'/></dd>"
				.Write "</form>"
				.Write "</dl>"
             End With

		  End Sub
		  
		  Sub doMerger()
		    Dim BoardID1,BoardID2
			BoardID1=KS.ChkClng(KS.G("BoardID1"))
			BoardID2=KS.ChkClng(KS.G("BoardID2"))
			If BoardID1=0 Or BoardID2=0 Then
			  KS.AlertHintScript ("请选择要合并的源版面及目标版面!")
			End If
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_GuestBoard Where ID=" & BoardID1,conn,1,1
			If Not RS.Eof Then
			 Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum+" & RS("TodayNum")& ",TopicNum=TopicNum+" & RS("TopicNum") & ",PostNum=PostNum+" & rs("PostNum") & " Where id=" & BoardID2)
			 Conn.Execute("Update KS_GuestBook Set BoardID=" & BoardID2 & " Where BoardID=" &Boardid1)
			 Conn.Execute("Update KS_GuestCategory Set BoardID=" & BoardID2 & " Where BoardID=" &Boardid1)
			End If
			RS.Close
			Set RS=Nothing
			Conn.Execute("Delete From KS_GuestBoard Where ID=" & BoardID1)
			Application(KS.SiteSN&"_ClubBoard")=empty
			KS.AlertHintScript "恭喜，论坛版面合并成功!"
		  End Sub
		  
		  '添加修改版块
		  Sub GuestBoardAddOrEdit()
		  		Dim GuestBoardID, RSObj, SqlStr, Content, BoardName, Note, Master, AddDate,Flag, Page,OrderID,ParentID,BoardRules,Settings,SetArr,Locked,ShowOther
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					GuestBoardID = KS.G("GuestBoardID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT top 1 * FROM KS_GuestBoard Where ID=" & GuestBoardID
					RSObj.Open SqlStr, Conn, 1, 1
					  BoardName     = RSObj("BoardName")
					  Note    = RSObj("Note")
					  AddDate  = RSObj("AddDate")
					  Master  = RSObj("Master")
					  ParentID= RSObj("ParentID")
					  OrderID = RSObj("OrderID")
					  BoardRules=RSObj("BoardRules")
					  Locked = RSObj("Locked")
					  Settings=RSObj("Settings")&"$0$0$0$0$1$1$1$1$20$$1$1$10$1$0$0$0$1$1$20$20$0$0$0$0$1$1$1$1$20$$1$1$10$1$0$0$0$1$1$20$20$$$$$$$$$$$$$$$$$$$$$"
					RSObj.Close:Set RSObj = Nothing
				Else
				   Flag = "Add"
				   ParentID=KS.ChkClng(Request("Parentid"))
				   BoardRules="暂无版规" : Locked=0 : OrderID=0
				End If
				Settings=Settings&"1$0$0$1$1$1$1$1$1$20$$0$0$10$1$0$0$0$1$1$20$10$0$0$0$0$0$1000$50$0$1$1$1$1$1$1$0$jpg|gif|png$100$5$0$0$0$0$0$0$0$0$0$0$0$$0$0$0$$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$0$"
				SetArr=Split(Settings,"$")
				ShowOther=true
				If ParentID=0  Then 
				 ShowOther=false
				End If
				
				With Response
			    .Write "<!DOCTYPE HTML>"
			    .Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				.Write "<title>版块管理</title>"
				.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<script src=""../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
				.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>"
		        .Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write EchoUeditorHead
				.Write "<script language=""JavaScript"">" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.GuestBoardForm;" & vbCrLf
				.Write "  if (form.BoardName.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    top.$.dialog.alert('请输入版块名称!',function(){" & vbCrLf
				.Write "    form.BoardName.focus();});" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
			If ShowOther=true Then
				.Write "   if (form.Note.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    top.$.dialog.alert('请输入版块介绍!',function(){" & vbCrLf
				.Write "    form.Note.focus();});" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
			End If
				.Write "      if (form.OrderID.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    top.$.dialog.alert('请输入版块序号!',function(){" & vbCrLf
				.Write "    form.OrderID.focus();});" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   form.submit();"
				.Write "   return true;"
				.Write "}"
				.Write "//-->"
				.Write "</script>"
				.Write "</head>"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>"
				If Flag = "Edit" Then
				 .Write "修改版块"
				Else
				 .Write "添加版块"
				End If
	            .Write "</div>"
				.Write "<br>"
				
		If ShowOther=false Then
		        .Write "<dl class='dtable'>"
				.Write "  <form name=GuestBoardForm method=post action=""?Action=Save"">"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""GuestBoardID"" value=""" & GuestBoardID & """>"
				.Write "   <input type=""hidden"" name=""parentid"" value=""0"">"
				.Write "          <dd>"
				.Write "            <div>版块状态:</div>"
				.write "<input type=""radio"" name=""Locked"" value=""0"" "
				If KS.ChkClng(Locked) = 0 Then .Write (" checked")
				.Write ">"
				.Write "开放"
				.Write "  <input type=""radio"" name=""Locked"" value=""1"" "
				If KS.ChkClng(Locked) = 1 Then .Write (" checked")
				.Write ">"
				.Write "锁定"
				.Write "          </dd>"
				
				.Write "          <dd>"
				.Write "            <div>分区名称:</div>"
				.Write "              <input name=""BoardName"" type=""text"" id=""BoardName"" value=""" & BoardName & """ class=""textbox"" size='50'> 如：技术交流、健康咨询等"
				 .Write "</dd>"
				
				.Write "<dd>"
				.Write "  <div>分区广告代码:</div>"
				.Write "<textarea name=""Note"" cols='75' rows='6' class=""textbox"" style=""height:110px;width:550px"">" & Note &"</textarea><span class=""block"">Tips:可以留空表示不显示广告。否则在首页的分区下将显示广告，支持HTML语法。</span>"
				.Write "     </dd>"			
				
				
				.Write "     <dd>"
				.Write "            <div>下级子版块横排:</div>"
			
				.Write " <input name=""SetArr(52)"" type=""text"" size=""5"" value=""" & SetArr(52) &""" class=""textbox"" style=""text-align:center""> 个 <span>如果设置为 0，则按正常方式排列</span>"
				         
				.Write "          </dd>"
				.Write "          <dd>"
				.Write "            <div>排 序 号:</div>"
				.Write "              <input name=""OrderID"" size=""5"" style=""text-align:center"" type=""text"" value=""" & OrderID &""" class=""textbox""> <span>序号越小，排在越前面</span>"
				.Write "          </dd>"
				
				
			
				
				

				.Write "</dl>"
				.Write "</div>"
		
		Else
				.write "<div class=tab-page id=boardpanel>"
				.Write "  <form name=GuestBoardForm method=post action=""?Action=Save"">"
				.Write " <SCRIPT type=text/javascript>"& _
				"   var tabPane1 = new WebFXTabPane( document.getElementById( ""boardpanel"" ), 1 )"& _
				" </SCRIPT>"& _
					 
				" <div class=tab-page id=basic-page>"& _
				"  <H2 class=tab>基本信息</H2>"& _
				"	<SCRIPT type=text/javascript>"& _
				"				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"& _
				"	</SCRIPT>" 
				
				.Write "<dl class='dtable'>"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""GuestBoardID"" value=""" & GuestBoardID & """>"
				.Write "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.Write "          <dd>"
				.Write "            <div>版块状态:</div>"
				.write "<input type=""radio"" name=""Locked"" value=""0"" "
				If KS.ChkClng(Locked) = 0 Then .Write (" checked")
				.Write ">"
				.Write "开放"
				.Write "  <input type=""radio"" name=""Locked"" value=""1"" "
				If KS.ChkClng(Locked) = 1 Then .Write (" checked")
				.Write ">"
				.Write "锁定"
				.Write "          </dd>"
				
				.Write "     <dd>"
				.Write "            <div>上级版块:</div>"
				.Write "             <select name='parentid' style=""width:300px"">"
				   Dim RST:Set RST=Conn.Execute("Select ID,BoardName From KS_GuestBoard Where ParentID=0 order by orderid")
				   Do While Not RST.Eof
				     If trim(ParentID)=trim(RST(0)) Then
				     .Write "<option value='" & RST(0) & "' selected>" & RST(1) & "</option>"
					 Else
				     .Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
					 End If
				   RST.MoveNext
				   Loop
				   RST.Close
				   Set RST=Nothing
				.Write "             </select>"  
				
				.Write "          </dd>"
				
				.Write "          <dd>"
				.Write "            <div>捆绑主模型:</div>"
				.Write "             <select name=""SetArr(60)"" style=""width:300px"">"
				.Write "              <option value='0'>---不绑定任何模型---</option>"
				Dim ModelXML,Node,Pstr:Pstr="@ks21=1 and @ks6=1"
				If Not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel[" & Pstr & "]")
				
				  If trim(SetArr(60))=trim(Node.SelectSingleNode("@ks0").text) Then
				  .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  Else
				  .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  End If
				  
				Next
				.Write "            </select><span>Tips:如果没有特殊情况，请不要选择。绑定主模型后发帖可以调用主模型的字段,主模型的评论也将直接调用对应帖子的回复数据，但性能有所下降,一旦设定，建议不要更改。</span>"
				.Write "          </dd>"			
				
				
				.Write "          <dd>"
				.Write "            <div>版块名称:</div>"
				.Write "              <input name=""BoardName"" type=""text"" id=""BoardName"" value=""" & BoardName & """ class=""textbox"" size=""50""> <span>如，技术交流、健康咨询等。</span>"
				 .Write "</dd>"
				.Write "          <dd>"
				.Write "            <div>版块图标:</div>"
				.Write "              <input name=""SetArr(51)"" type=""text"" id=""SetArr51"" value=""" & SetArr(51) & """ class=""textbox"" size=""50""> <input class='button' type='button' name='Submit' value='选择图片地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath=" & KS.GetCommonUpFilesDir() & "',550,290,window,$('#SetArr51')[0]);"">"
				 .Write "</dd>"
				 .Write "<dd>"
				.Write "  <div>版块介绍:</div>"
				.Write "<textarea name=""Note"" cols='75' rows='6' class=""textbox"" style=""height:110px;width:550px"">" & Note &"</textarea>"
				.Write "          </dd>"			
				
				.Write "          <dd>"
				.Write "            <div>分页设置:</div>"
				.Write "              列表页每页显示<input name=""SetArr(20)"" type=""text""  value=""" & SetArr(20) &""" class=""textbox"" style=""width:50px;text-align:center""> 条记录  帖子页每页显示 <input name=""SetArr(21)"" type=""text""  value=""" & SetArr(21) &""" class=""textbox"" style=""width:50px;text-align:center""> 条回复记录"
				.Write "          </dd>"
				.Write "          <dd>"
				.Write "            <div>版面列表启用缓存:</div>"
				.Write "              列表页前<input name=""SetArr(69)"" type=""text""  value=""" & KS.ChkClng(SetArr(69)) &""" class=""textbox"" style=""width:50px;text-align:center""> 页启用缓存 <span class='tips'>tips:建议设置不超过100页，设置“0”不启用缓存。</span>"
				.Write "          </dd>"
				
				
				.Write "          <dd>"
				.Write "            <div>版面显示方式:</div>"
				.Write "            <label><input type='radio' name='setarr(65)'"
				If trim(SetArr(65))="0" Then .Write " checked"
				.Write " value='0'>标题列表（默认）</label>"
				.Write "            <label><input type='radio' name='setarr(65)'"
				If trim(SetArr(65))="1" Then .Write " checked"
				.Write " value='1'>显示发帖者头像方式</label>"
				
				.Write "              </dd>"				
				.Write "          <dd>"
				.Write "            <div>帖子显示方式:</div>"
				.Write "            <label><input type='radio' name='setarr(66)'"
				If trim(SetArr(66))="0" Then .Write " checked"
				.Write " value='0'>详细（默认）</label>"
				.Write "            <label><input type='radio' name='setarr(66)'"
				If trim(SetArr(66))="1" Then .Write " checked"
				.Write " value='1'>简洁</label>"
				
				.Write "             &nbsp;&nbsp;<span class='tips'>简洁模式不支持广告及签名。</span>"
				.Write "          </dd>"				
				.Write "          <dd>"
				.Write "            <div>新贴显示标记:</div>"
				.Write "            <input name=""SetArr(42)"" type=""text""  value=""" & SetArr(42) &""" class=""textbox"" style=""width:50px;text-align:center"">小时内有新回复的帖子显示<span style='color:red'>New</span>标志,不显示请输入0"
				.Write "              </dd>"
				
				.Write "          <dd>"
				.Write "            <div>热帖设置:</div>"
				.Write "              浏览数大于<input name=""SetArr(27)"" type=""text""  value=""" & SetArr(27) &""" class=""textbox"" style=""width:50px;text-align:center""> 次且回复数大于<input name=""SetArr(28)"" type=""text""  value=""" & SetArr(28) &""" class=""textbox"" style=""width:50px;text-align:center"">楼时自动转为热帖"
				.Write "          </dd>"

				
				.Write "          <dd>"
				.Write "            <div>本版版主:</div>"
				.Write "           <input type=""hidden"" name=""omaster"" value=""" & master &""">"
				.Write "              <input name=""Master"" type=""text"" id=""Master"" value=""" & Master &""" class=""textbox"" size=""50""> <span>多个版主请用英文逗号隔开</span>"
				.Write "          </dd>"
				.Write "          <dd>"
				.Write "            <div>排 序 号:</div>"
				.Write "              <input name=""OrderID"" size=""5"" type=""text"" value=""" & OrderID &""" class=""textbox""> <span>序号越小，排在越前面</span>"
				.Write "          </dd>"

				.Write "</table>"
				.Write "</div>"
				
			
				
				.Write "<div class=tab-page id=""formset"">"
		        .Write " <H2 class=tab>发帖&浏览</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""formset"" ) );"
			    .Write "</SCRIPT>"
				.Write "<dl class='dtable'>"
				.Write "          <dd>"
				.Write "            <div>新注册用户:</div><input type='text' class='textbox' style='text-align:center' name='setarr(9)' size=5 value='" & setarr(9) & "'> 分钟后才可以在本版块发布帖子</dd>"
				.Write "<dd><div>是否允许游客浏览查看:</div>"
				.write "<label><input type=""radio"" onclick=""$('#showpower').hide()"" name=""setarr(0)"" value=""1"" "
				If KS.ChkClng(SetArr(0)) = 1 Then .Write (" checked")
				.Write ">"
				.Write "允许</label>"
				.Write "  <label><input type=""radio"" onclick=""$('#showpower').show()"" name=""setarr(0)"" value=""0"" "
				If KS.ChkClng(SetArr(0)) = 0 Then .Write (" checked")
				.Write ">"
				.Write "不允许</label>"
				.Write "              </dd>"
				
				If KS.ChkClng(SetArr(0)) = 1 Then
				.Write "<div id='showpower' style='font-size:12px;font-weight:normal;display:none;padding:10px'>"
				Else
				.Write "<div id='showpower' style='font-size:12px;font-weight:normal;padding:10px'>"
				End If
				.Write "<FIELDSET>"
				.Write "<LEGEND align=left>设置浏览版面权限</LEGEND>"
				.Write "<dd>"
				.Write "            <div>浏览本版块的权限：<font>(当不允许游客浏览本版块时可以在此进一步设置权限)</font></div>"
				.Write "           <strong>1、限制会员组:</strong><font>(不限制请不要勾选)</font>"
				.Write KS.GetUserGroup_CheckBox("SetArr(1)",SetArr(1),5)
				
				.Write "            <div style='clear:both'></div><div><strong>2、认证会员:</strong><font>(允许进入此版块的会员,不限制请留空。否则只有认证会员才可以进入)</font></div>"
				.Write "           <textarea name='setarr(10)' style='width:600px;height:80px'>" & setarr(10) & "</textarea><span class=""block"">多个认证会员，请用英文逗号隔开，如kesion1,kesion2等。</span>"
				.Write "            <div style='clear:both'></div><div><strong>3、有效期限制</strong></div>"
				.Write "            <label><input type='radio' name='SetArr(54)'"
				If SetArr(54)="0" Then .Write " checked"
				.Write " value=0>不启用有效期限制</label><br/>"
				.Write "            <label><input type='radio' name='SetArr(54)'"
				If SetArr(54)="1" Then .Write " checked"
				.Write " value=1>满足以上两个条件的任一条件，还必须是有效期内的会员才可以进去</label><br/>"
				.Write "            <label><input type='radio' name='SetArr(54)'"
				If SetArr(54)="2" Then .Write " checked"
				.Write " value=2>不管是否满足以上两个条件，只要在有效期内的会员就可以进去</label><br/>"
				.Write "              </dd>"
				.Write "<dd>"
				.Write "            <div>积分/资金限制:</div>"
				.Write "            用户积分必须大于等于<input type='text' class='textbox' style='text-align:center' name='setarr(11)' size=5 value='" & setarr(11) & "'>个积分才可以进入此版块浏览及发帖<br/>用户资金必须大于等于<input type='text' style='text-align:center' name='setarr(12)' class='textbox' size=5 value='" & setarr(12) & "'>元才可以进入此版块浏览及发帖<br/><font color=blue>说明：如果启用有效期用户浏览，在有效期内的会员不受此限制!</font></dd>"
				.Write "</FIELDSET>"
				.Write "</div>"
				
				.Write "<dd>"
				.Write "            <div>允许发帖的会员组:</div><strong>1、允许发表主题的用户组：</strong><font>(不限制请不要勾选)</font><br/>"
				.Write KS.GetUserGroup_CheckBox("SetArr(2)",SetArr(2),5)
				.Write "   <div style=""clear:both""></div><strong>2、允许发表回复的用户组：</strong><font>(<不限制请不要勾选)</font><br/>"
				.Write KS.GetUserGroup_CheckBox("SetArr(62)",SetArr(62),5)         &" </dd>"
				.Write "<dd>"
				.Write "            <div>本版块发帖限制:</div>一天内每个会员最多只能发表<input type='text' class='textbox' style='text-align:center' name='setarr(13)' size=5 value='" & setarr(13) & "'>条主题 "
				
				.Write "发帖字数不少于<input type='text' class='textbox' style='text-align:center' name='setarr(40)' size=5 value='" & setarr(40) & "'>个字  发帖间隔时间<input type='text' style='text-align:center' name='setarr(41)' class='textbox' size=5 value='" & setarr(41) & "'>秒 <span style='color:green'>不限制请填0</span>"
				
				.Write "              </dd>"
				
				.Write "<dd>"
				.Write "            <div>本版块投票帖最多选项:</div><input type='text' style='text-align:center' name='setarr(64)' class='textbox' size=5 value='" & setarr(64) & "'>个投票选项,<span class='tips'>此版面不允许发投票帖，请输入“0”。</span>"
				
				.Write "              </dd>"
				
				
				
				.Write "<dd>"
				.Write "            <div>帖子操作选项:</div><label><input type='checkbox' name='setarr(14)'"
				If trim(SetArr(14))="1" Then .Write " checked"
				.Write " value='1'>允许回复自已的帖子</label>"
				.Write "           <label><input type='checkbox' name='setarr(29)'"
				If trim(SetArr(29))="1" Then .Write " checked"
				.Write " value='1'>允许编辑自已的帖子</label>"
				.Write "           <label><input type='checkbox' name='setarr(63)'"
				If trim(SetArr(63))="1" Then .Write " checked"
				.Write " value='1'>启用自动远程存图到本地</label>"
				.Write "           <label><input type='checkbox' name='setarr(67)'"
				If trim(SetArr(67))="1" Then .Write " checked"
				.Write " value='1'>同步主题到微博</label>"
				
				.Write "              </dd>"
				
				.Write "<dd>"
				.Write "            <div>帖子回复操作选项:</div> 帖子超过<input type='text'  <input type=""text""  class='textbox' style='text-align:center' size=""6"" value=""" & SetArr(91) & """ name=""SetArr(91)""  onKeyUp=""value=value.replace(/\D/g,'')"" onafterpaste=""value=value.replace(/\D/g,'')"" />天不能回复 <span style='color:green'>不限制请填0</span>"
				.Write "              </dd>"
				
				
				.Write "<dd>"
				.Write "            <div>会员发帖时间限制:</div> 发帖时间从<input type='text'  <input type=""text""  class='textbox' style='text-align:center' size=""6"" value=""" & SetArr(92) & """ name=""SetArr(92)""  onKeyUp=""value=value.replace(/\D/g,'')"" onafterpaste=""value=value.replace(/\D/g,'')"" /> - <input type=""text""  class='textbox' style='text-align:center' size=""6"" value=""" & SetArr(93) & """ name=""SetArr(93)""  onKeyUp=""value=value.replace(/\D/g,'')"" onafterpaste=""value=value.replace(/\D/g,'')"" />点不能发帖 <span style='color:green'>不限制请填0</span>"
				.Write "              </dd>"

				.Write "<dd>"
				.Write "            <div>发帖开启HTML支持:</div><label><input type='radio' name='setarr(59)'"
				If trim(SetArr(59))="0" Then .Write " checked"
				.Write " value='0'>不支持</label>"
				.Write "            <label><input type='radio' name='setarr(59)'"
				If trim(SetArr(59))="1" Then .Write " checked"
				.Write " value='1'>支持</label>"
				
				.Write "             <span style='color:#999'>开启后用户将可以使用html语法标记，有一定的安全隐患。 </span></dd>"				
				.Write "<dd>"
				.Write "            <div>发帖审核模式:</div><label><input type='radio' name='setarr(61)'"
				If trim(SetArr(61))="0" Then .Write " checked"
				.Write " value='0'>不启用审核</label><br/>"
				.Write "            <label><input type='radio' name='setarr(61)'"
				If trim(SetArr(61))="1" Then .Write " checked"
				.Write " value='1'>发表主题需要审核，回复不需要审核</label><br/>"
				.Write "            <label><input type='radio' name='setarr(61)'"
				If trim(SetArr(61))="2" Then .Write " checked"
				.Write " value='2'>发表主题和回复都需要审核</label><br/>"
				.Write "            <label><input type='radio' name='setarr(61)'"
				If trim(SetArr(61))="3" Then .Write " checked"
				.Write " value='3'>发表主题不需要审核，回复需要审核</label>"
				
				.Write "              </dd>"				
				.Write "<dd>"
				.Write "            <div>在本版块发帖需要填写验证码:</div><label><input type='radio' name='setarr(53)'"
				If trim(SetArr(53))="0" Then .Write " checked"
				.Write " value='0'>不需要</label>"
				.Write "            <label><input type='radio' name='setarr(53)'"
				If trim(SetArr(53))="1" Then .Write " checked"
				.Write " value='1'>需要</label>"
				
				.Write "              </dd>"
				
				
			.Write "    <dd>"
			.Write "      <div>允许会员上传附件：</div>"
			 .Write "    <input onclick=""document.getElementById('fj').style.display='';"" name=""SetArr(36)"" type=""radio"" value=""1"""
			 If SetArr(36)="1" Then .Write " Checked"
			 .Write ">允许 <input name=""SetArr(36)"" onclick=""document.getElementById('fj').style.display='none';"" type=""radio"" value=""0"""
			 If SetArr(36)="0" Then .Write " Checked"
			 .Write ">不允许"
			 If SetArr(36)="1" Then
			  .Write "<div id='fj'>"
			 Else
			  .Write "<div id='fj' style='display:none;'>"
			 End If
			 .Write "<FIELDSET>"
			.Write "<LEGEND align=left>设置上传参数</LEGEND>"

			 .Write "<font color=green>允许上传的文件类型：<input class='textbox' name=""SetArr(37)"" type=""text"" value=""" & SetArr(37) &""" size='30'>多个类型用|线隔开<br/>允许上传的文件大小：<input class='textbox' name=""SetArr(38)"" type=""text"" value=""" & SetArr(38) &""" style=""text-align:center"" size='8'>KB<br/>每天上传文件个数：<input class='textbox' name=""SetArr(39)"" type=""text"" value=""" & SetArr(39) &""" style=""text-align:center"" size='8'>个,不限制请填0<br/>"
			  .Write "<strong>如果上传的是图片，则自动增加水印<input type=""checkbox"" name=""SetArr(43)"" value=""1"""
			 if SetArr(43)="1" then .Write " checked"
			 .Write "/></strong></font><br/>"
			 .Write "<br/><strong>允许在此版块上传附件的用户组:</strong>(<font color=blue>不限制请不要勾选</font>)"
			 .Write KS.GetUserGroup_CheckBox("SetArr(17)",SetArr(17),5)
			 .Write "</div>"
			 .Write "</dd>"
			 .Write "</FIELDSET>"
	

				.Write "</dl>"
				.Write "</div>"
				
				.Write "<div class=tab-page id=""comments"">"
		        .Write " <H2 class=tab>帖子点评设置</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""comments"" ) );"
			    .Write "</SCRIPT>"
				.Write "<dl class='dtable'>"

				.Write "<dd>"
				.Write "            <div>每页显示点评条数:</div><input type='text' class='textbox' style='text-align:center' name='setarr(44)' size=5 value='" & setarr(44) & "'>条 <span style='color:green'>此版面不启用点评功能，请填“0”</span></dd>"
				.Write "<dd>"
				.Write "            <div>会员威望达到:</div><input type='text' class='textbox' style='text-align:center' name='setarr(45)' size=5 value='" & setarr(45) & "'>分 才可能对帖子进行点评 <span style='color:green'>为防止恶意点评攻击，建议只有达到一定威望的会员才能发表点评,不限制请输入0</span></dd>"
				.Write "<dd>"
				.Write "            <div>允许对主题进行点评:</div><label><input type='radio' name='setarr(46)'"
				If trim(SetArr(46))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(46)'"
				If trim(SetArr(46))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				.Write "              </dd>"				
				.Write "<dd>"
				.Write "            <div>允许对回复进行点评:</div><label><input type='radio' name='setarr(47)'"
				If trim(SetArr(47))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(47)'"
				If trim(SetArr(47))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				.Write "              </dd>"				
				.Write "<dd>"
				.Write "            <div>允许点评自己的帖子:</div><label><input type='radio' name='setarr(48)'"
				If trim(SetArr(48))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(48)'"
				If trim(SetArr(48))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				.Write "              </dd>"				
				.Write "<dd>"
				.Write "            <div>点评预置观点:</div><textarea name=""setarr(49)"" cols=""50"" rows=""3"">" & SetArr(49) & "</textarea>"
				.Write "             <br/><span style='color:green'>可选项，多个观点请用英文“,”号隔开，如""赞同,反对,中立""</span> </dd>"				
				.Write "<dd>"
				.Write "            <div>点评算入今日发帖数:</div><label><input type='radio' name='setarr(50)'"
				If trim(SetArr(50))="0" Then .Write " checked"
				.Write " value='0'>不计数</label>"
				.Write "            <label><input type='radio' name='setarr(50)'"
				If trim(SetArr(50))="1" Then .Write " checked"
				.Write " value='1'>计数</label>"
				.Write "              </dd>"				
				
                .Write "</dl>"
				.Write "</div>"				
				
				.Write "<div class=tab-page id=""scores"">"
		        .Write " <H2 class=tab>积分威望</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""scores"" ) );"
			    .Write "</SCRIPT>"
				.Write "<dl class='dtable'>"

				.Write "<dd>"
				.Write "            <div>下载附件最少达到积分:</div><input type='text' class='textbox' style='text-align:center' name='setarr(15)' size=5 value='" & setarr(15) & "'>个积分 <span style='color:green'>如果用户积分少于这里设置的最低积分值将不能下载,不限制请填0</span></dd>"
				.Write "<dd>"
				.Write "            <div>在此版块下载附件需消耗:</div><input type='text' class='textbox' style='text-align:center' name='setarr(16)' size=5 value='" & setarr(16) & "'>个积分 <span style='color:green'>24小时内重复下载只扣一次,不限制请填0</span></dd>"
				.Write "<dd>"
				.Write "            <div>本版块是否允许发出售帖:</div><label><input type='radio' name='setarr(55)'"
				If trim(SetArr(55))="0" Then .Write " checked"
				.Write " value='0' onclick=""$('#sale').hide()"">不允许</label>"
				.Write "            <label><input type='radio' name='setarr(55)'"
				If trim(SetArr(55))="1" Then .Write " checked"
				.Write " value='1' onclick=""$('#sale').show()"">允许</label></dd>"
				If trim(SetArr(55))="1" Then .Write "<tbody id='sale'>" Else  .Write "<tbody id='sale' style='display:none'>"
				.Write "<dd>"
				.Write "            <div>出售帖计费方式:</div><input type=""radio"" name=""SetArr(56)"" value=""0"" "
		If SetArr(56) = "0" Then .Write (" checked")
		.Write ">" & KS.Setting(45)
		.Write "          <input type=""radio"" name=""SetArr(56)"" value=""1"" "
		If SetArr(56) = "1" Then .Write (" checked")
		.Write ">资金(人民币)"		
		.Write "          <input type=""radio"" name=""SetArr(56)"" value=""2"" "
		If SetArr(56) = "2" Then .Write (" checked")
		.Write "> 积分   </dd>"
				.Write "<dd>"
				.Write "            <div>最高售价限:</div><input type=""text"" class=""textbox"" style='text-align:center' name=""SetArr(57)"" size='5' value=""" & SetArr(57) & """> <span style='color:green'>出售帖最高售价，不限制请输入0!</span>  </dd>"
				.Write "<dd>"
				.Write "            <div>分成比率:</div><input type=""text"" class=""textbox"" style='text-align:center' name=""SetArr(58)"" size='5' value=""" & SetArr(58) & """> % <span style='color:green'>系统将根据这里设置的分成比率将收成分给投稿者。建议设成10的整数倍!</span>  </dd>"
				.Write "</tbody>"
				

				.Write "<dd>"
				.Write "            <div>积分威望设置:</div>"
				%>
				<table width="80%" border="0">
  <tr>
    <td align="center">类型</td>
    <td align="center"><strong>发表主题</strong></td>
    <td align="center"><strong>发表回复</strong></td>
    <td align="center"><strong>置顶</strong></td>
    <td align="center"><strong>精华</strong></td>
    <td align="center"><strong>被删主题</strong></td>
    <td align="center"><strong>被删回复</strong></td>
  </tr>
  <tr>
    <td><strong>积分</strong></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(3)' size=5 value='<%=setarr(3)%>'></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(4)' size=5 value='<%=setarr(4)%>'></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(5)' size=5 value='<%=setarr(5)%>'></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(6)' size=5 value='<%=setarr(6)%>'></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(7)' size=5 value='<%=setarr(7)%>'></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(8)' size=5 value='<%=setarr(8)%>'></td>
  </tr>
  <tr>
    <td><strong>威望</strong></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(30)' size=5 value='<%=setarr(30)%>' /></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(31)' size=5 value='<%=setarr(31)%>' /></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(32)' size=5 value='<%=setarr(32)%>'/></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(33)' size=5 value='<%=setarr(33)%>'/></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(34)' size=5 value='<%=setarr(34)%>'/></td>
    <td><input type='text' class='textbox' style='text-align:center' name='setarr(35)' size=5 value='<%=setarr(35)%>'/></td>
  </tr>
</table>

				<%
				.Write "</dd>"
                .Write "</dl>"
				.Write "</div>"
				
				.Write "<div class=tab-page id=""boardrule"">"
		        .Write " <H2 class=tab>设置版规</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardrule"" ) );"
			    .Write "</SCRIPT>"
				.Write "<dl class='dtable'>"
				 .Write "<dd>"
				.Write "  <div>版 规<font color=blue>(可以留空)</font></div>"
				
				.Write EchoEditor("BoardRules",BoardRules,"Basic","96%","350px")
				
				.Write "            </dd>"
				.Write "</dl>"
				.Write "</div>"
				.Write "<div class=tab-page id=""boardclass"">"
		        .Write " <H2 class=tab>主题分类</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardclass"" ) );"
			    .Write "</SCRIPT>"
				.Write "<dl class='dtable'>"
				.Write "<dd>"
				.Write "  <div>启用主题分类:</div>"
				.Write "  <label><input type='radio' name='setarr(23)'"
				If trim(SetArr(23))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(23)'"
				If trim(SetArr(23))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>设置是否在本版块启用主题分类功能，您需要同时设定相应的分类选项，才能启用本功能</span>"
				.Write " </dd>"
				.Write "<dd>"
				.Write "  <div>发帖必须归类:</div>"
				.Write "  <label><input type='radio' name='setarr(24)'"
				If trim(SetArr(24))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(24)'"
				If trim(SetArr(24))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " <span>是否强制用户发表新主题时必须选择分类</span>"
				.Write " </dd>"
				.Write "<dd>"
				.Write "  <div>类别前缀:</div>"
				.Write "  <label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="0" Then .Write " checked"
				.Write " value='0'>不显示</label>  &nbsp;&nbsp;<span>是否在主题前面显示分类的名称</span>"
				.Write "            <br/><label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="1" Then .Write " checked"
				.Write " value='1'>只显示文字</label>"
				.Write "           <br/> <label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="2" Then .Write " checked"
				.Write " value='2'>只显示图标</label>"
				
				.Write "</dd>"
				
				.Write "<dd>"
				.Write "  <div>允许按类别浏览:</div>"
				.Write "  <label><input type='radio' name='setarr(26)'"
				If trim(SetArr(26))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(26)'"
				If trim(SetArr(26))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " <span>用户是否可以按照主题分类筛选浏览内容</span>"
				.Write " </dd>"
				.Write "<dd>"
				.Write "  <div>回帖时允许版主及管理员更改归类:</div>"
				.Write "  <label><input type='radio' name='setarr(68)'"
				If trim(SetArr(68))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(68)'"
				If trim(SetArr(68))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " </dd>"
				
				.Write "<dd>"
				.Write "<div>主题分类</div>"
				%>
<script type="text/JavaScript">
	var rowtypedata = [
		[
			[1,'', 'tdbg'],
			[1,'<div style="text-align:center">是</div>', 'tdbg'],
			[1,'<input type="text" class="textbox" size="2" name="categoryorder" value="0" />', 'tdbg'],
			[1,'<input type="text" class="textbox" name="categoryname"  size="30"/>', 'tdbg'],
			[1,'<input type="text" class="textbox" name="categoryicon" size="30"/>', 'tdbg'],
			[1,'', 'tdbg']
		],
	];

var addrowdirect = 0;
function addrow(obj, type) {
	var table = obj.parentNode.parentNode.parentNode.parentNode;
	if(!addrowdirect) {
		var row = table.insertRow(obj.parentNode.parentNode.parentNode.rowIndex);
	} else {
		var row = table.insertRow(obj.parentNode.parentNode.parentNode.rowIndex + 1);
	}
	var typedata = rowtypedata[type];
	for(var i = 0; i <= typedata.length - 1; i++) {
		var cell = row.insertCell(i);
		cell.colSpan = typedata[i][0];
		var tmp = typedata[i][1];
		if(typedata[i][2]) {
			cell.className = typedata[i][2];
		}
		tmp = tmp.replace(/\{(\d+)\}/g, function($1, $2) {return addrow.arguments[parseInt($2) + 1];});
		cell.innerHTML = tmp;
	}
	addrowdirect = 0;
}
</script>

<div id="threadtypes_manage">
<table cellspacing="1" width="80%" cellpadding="1" border="0">
<tr style='font-weight:bold;text-align:center' class="title">
<td height='22'>删除</td><td>启用</td><td>显示顺序</td><td>分类名称</td><td>前缀图标</td></tr>
<%
If GuestBoardID<>0 Then
  Dim RS:Set RS=Conn.Execute("Select * From KS_GuestCategory Where BoardID=" & GuestBoardID)
  Do While Not RS.Eof
    Response.Write "<tr><td align=""center""><input type=""hidden"" name=""categoryid"" value=""" &rs("categoryid") & """>"
	Response.Write "<input type=""checkbox"" value=""1"" onclick=""if (this.checked){return(confirm('确定删除该分类吗?'))}"" name=""categorydel" & RS("CategoryID") & """>"
	Response.Write "</td><td align=""center""><input type=""checkbox"" value=""1"" name=""categorystatus" & RS("CategoryID") & """ "
	if rs("status")="1" then response.write " checked"
	Response.Write "/>"
	response.write "<td><input type=""text"" class=""textbox"" size=""2"" name=""categoryorder"" value=""" & rs("orderid") &""" /></td>"
	response.write "<td><input type=""text"" class=""textbox"" name=""categoryname"" size=""30"" value=""" & rs("categoryname") &""" /></td>"
	response.write "<td><input type=""text"" class=""textbox"" name=""categoryicon""  size=""30"" value=""" & rs("ico") &""" /></td>"
	response.write "</tr>"
  RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
End If
%>
<tr><td colspan="6">
<div><img src="../images/accept.gif" align="absmiddle"/> <a href="javascript:;" onclick="addrow(this, 0)" class="addtr">添加分类</a></div>
</td>
</tr>
</table>
</div>

</dd>
				<%
                .Write "</dl>"
				.Write "</div>"
				
				.Write "<div class=tab-page id=""boardseo"">"
		        .Write " <H2 class=tab>版面SEO选项</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardseo"" ) );"
			    .Write "</SCRIPT>"
				.Write "<dl class='dtable'>"
				.Write "<dd>"
				.Write "  <div>SEO标题:</div>"
				.Write "  <label><input type='text' name='setarr(70)' style='width:350px' value='" & setarr(70) &"' class='textbox'></label>"
				.Write " &nbsp;&nbsp;<span class='tips'>提示：在模板里可以用{$Board.SeoTitle}进行调用。</span>"
				.Write " </dd>"
				.Write "<dd>"
				.Write "  <div>SEO关键字:</div>"
				.Write "  <label><textarea name='setarr(71)' class='textbox' style='width:350px;hieght:60px'>"& setarr(71) &"</textarea></label>"
				.Write " &nbsp;&nbsp;<span class='tips'>提示：在模板里可以用{$Board.SeoKeyWords}进行调用。</span>"
				.Write " </dd>"
				.Write "<dd>"
				.Write "  <div>SEO描述:</div>"
				.Write "  <label><textarea name='setarr(72)' class='textbox' style='width:350px;hieght:60px'>"& setarr(72) &"</textarea></label>"
				.Write " &nbsp;&nbsp;<span class='tips'>提示：在模板里可以用{$Board.SeoDescription}进行调用。</span>"
				.Write " </dd>"
                .Write "</dl>"
				.Write "</div>"
		End If		
								
				.Write "  </form>"
				.Write "</body>"
				.Write "</html>"
			 End With
		  End Sub
		  
		  '保存
		  Sub GuestBoardSave()
		    Dim categoryid:categoryid=KS.G("categoryid")&",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			Dim CategoryName:CategoryName=KS.G("CategoryName")
			Dim categoryorder:categoryorder=KS.G("categoryorder")
            Dim categoryicon:categoryicon=KS.G("categoryicon")
			Dim categorystatus:categorystatus=KS.G("categorystatus")
			Dim RS,CategoryNameArr,categoryorderArr,categoryiconArr,categorystatusArr,CategoryIDArr
			
			Dim GuestBoardID, RSObj, SqlStr, BoardName, Note, AddDate, Content, Master,Flag, Page, RSCheck,OrderID,ParentID,BoardRules,Settings,I,Locked
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			GuestBoardID = Request("GuestBoardID")
			BoardName = Replace(Replace(Request.Form("BoardName"), """", ""), "'", "")
			Note = Request.Form("Note")
			Master = Request.Form("Master")
			BoardRules=Request.Form("BoardRules")
			OrderID = KS.ChkClng(KS.G("OrderID"))
			ParentID = KS.Chkclng(Request.Form("ParentID"))
			Locked  = KS.ChkClng(Request.Form("Locked"))
			If BoardName = "" Then Call KS.AlertHistory("版块名称不能为空!", -1)
			'If Note = "" Then Call KS.AlertHistory("版块介绍不能为空!", -1)
			
			
			For I=0 To 100
			  If I=0 Then 
			   Settings=Request("setarr(" & i & ")") &"$"
			  Else
			   Settings=Settings  & Request("setarr(" & i & ")")& "$"
			  End If
			Next
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select top 1 ID From KS_GuestBoard Where ParentID=" & ParentID & " and BoardName='" & BoardName & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  Response.Write ("<script>top.$.dialog.alert('对不起,名称已存在!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT top 1 * FROM KS_GuestBoard Where 1=0", Conn, 1, 3
				RSObj.AddNew
				  RSObj("BoardName") = BoardName
				  RSObj("Note") = Note
				  RSObj("AddDate") = Now
				  RSObj("Master") = Master
				  RSObj("OrderID") =OrderID
				  RSObj("ParentID")=ParentID
				  RSObj("lastpost")="0$" & now & "$无$$$$$"
				  RSObj("TodayNum")=0
				  RSObj("PostNum")=0
				  RSObj("TopicNum")=0
				  RSObj("Locked")=Locked
				  RSObj("BoardRules")=BoardRules
				  RSObj("Settings")=Settings
				RSObj.Update
				GuestBoardID=RSObj("ID")
				 RSObj.Close
			If Not KS.IsNul(CategoryName) Then
			   CategoryNameArr=Split(Replace(CategoryName," ",""),",")
			   categoryorder=split(Replace(categoryorder," ",""),",")
			   categoryiconArr=split(Replace(categoryicon," ",""),",")
			   categorystatusArr=split(Replace(categorystatus," ",""),",")
			   Set RS=Server.CreateObject("ADODB.RECORDSET")
			   For I=0 To Ubound(CategoryNameArr) 
		          RS.Open "Select top 1 * From KS_GuestCategory",conn,1,3
				  RS.AddNew
				    RS("CategoryName")=CategoryNameArr(i)
					RS("OrderID")=KS.ChkClng(categoryorder(i))
					RS("Ico")=trim(categoryiconArr(i))
					RS("Status")=1
					RS("BoardID")=GuestBoardID
				  RS.Update
				  RS.Close
               Next
		   End If
				
				
			  End If
			   Set RSObj = Nothing
			   Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
			   Call KS.FileAssociation(1050,GuestBoardID,Settings&categoryicon,0)
			   Response.Write ("<script> top.$.dialog.confirm('版块添加成功!继续添加吗?',function(){location.href='" & KS.Setting(3) & KS.Setting(89) & "club/KS.GuestBoard.asp?Action=Add&parentid=" & ParentID &"';},function(){location.href='" & KS.Setting(3) & KS.Setting(89) & "club/KS.GuestBoard.asp';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=论坛系统 >> <font color=red>论坛版块管理</font>';});</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_GuestBoard Where  parentid=" & parentid & " and BoardName='" & BoardName & "' And ID<>" & GuestBoardID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 Response.Write ("<script>top.$.dialog.alert('对不起,版块名称已存在!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT top 1 * FROM KS_GuestBoard Where ID=" & GuestBoardID
			   RSObj.Open SqlStr, Conn, 1, 3
				 RSObj("BoardName") = BoardName
				 RSObj("Note") = Note
				 RSObj("Master") = Master
				 RSObj("OrderID") =OrderID
				 RSObj("Locked")=Locked
				 RSObj("ParentID")=ParentID
				 RSObj("BoardRules")=BoardRules
				 RSObj("Settings")=Settings
			   RSObj.Update
			   RSObj.Close
			   Set RSObj = Nothing
			   
			If Not KS.IsNul(CategoryName) Then
			   CategoryNameArr=Split(CategoryName,",")
			   categoryorder=split(Replace(categoryorder," ","")&",,,,,,,,,,,",",")
			   categoryiconArr=split(Replace(categoryicon," ","")&",,,,,,,,,,,",",")
			   categorystatusArr=split(Replace(categorystatus," ","")&",,,,,,,,,,,",",")
			   categoryIdArr=split(Replace(categoryId," ","")&",,,,,,,,,,,",",")
			   Set RS=Server.CreateObject("ADODB.RECORDSET")
			   For I=0 To Ubound(CategoryNameArr)
			      if KS.ChkClng(categoryIdArr(i))<>0 and KS.ChkClng(KS.S("categorydel"&KS.ChkClng(categoryIdArr(i))))=1 Then
				   Conn.Execute("Delete From KS_GuestCategory Where CategoryID=" & KS.ChkClng(categoryIdArr(i)))
				  Else
					  RS.Open "Select top 1 * From KS_GuestCategory Where CategoryID=" & KS.ChkClng(categoryIdArr(i)),conn,1,3
					  If RS.Eof and RS.Bof Then
					   RS.AddNew
					   RS("Status")=1
					  Else
					   RS("Status")=KS.ChkClng(KS.S("categorystatus" & categoryIdArr(i)))
					  End If
						RS("CategoryName")=trim(CategoryNameArr(i))
						RS("OrderID")=KS.ChkClng(categoryorder(i))
						RS("Ico")=trim(categoryiconArr(i))
						RS("BoardID")=GuestBoardID
					  RS.Update
					  RS.Close
				End If
               Next
		   End If
			   
			  End If
			  Application(KS.SiteSN&"_ClubBoard")=empty
			  Application(KS.SiteSN&"ClubIndex")=empty
			  Call KS.FileAssociation(1050,GuestBoardID,Settings&categoryicon,1)
			  If trim(lcase(KS.g("omaster")))<>trim(lcase(Master)) Then  UpdateMasterToUser
			  Response.Write ("<script>top.$.dialog.alert('版块修改成功!',function(){location.href='" & KS.Setting(3) & KS.Setting(89) & "club/KS.GuestBoard.asp?Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "Post.Asp?ButtonSymbol=Disabled&OpStr=论坛系统 >> <font color=red>论坛版块管理</font>';});</script>")
			End If
		  End Sub
		  
		   '更新KS_User表的版主
		  Sub UpdateMasterToUser()	
			   KS.LoadClubBoard
			   dim node,xml,master,masterarr,i
			   set xml=Application(KS.SiteSN&"_ClubBoard")
			   If IsObject(XML) Then
			     
			    for each node in xml.documentelement.selectnodes("row")
				 if node.selectsinglenode("@master").text<>"" then
					  if master="" then
					   master=node.selectsinglenode("@master").text
					  else
					   master=master& "," & node.selectsinglenode("@master").text
					  end if
				 end if
			    next
			   end if
			   dim rs,newmaster,bzgradeid,admingradeid,superbzgradeid,rsg
			   set rs=server.createobject("adodb.recordset")
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='版主'",conn,1,1
				 if not rs.eof then
				  bzgradeid=rs("gradeid")
				 else
				  bzgradeid=0
				 end if
				 rs.close
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='管理员'",conn,1,1
				 if not rs.eof then
				  admingradeid=rs(0)
				 else
				  admingradeid=0
				 end if
				 rs.close
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='超级版主'",conn,1,1
				 if not rs.eof then
				  superbzgradeid=rs(0)
				 else
				  superbzgradeid=0
				 end if
				 rs.close
			   if not ks.isnul(master) then
			     masterarr=split(master,",")
				 '先更新用户在论坛级别ID
				 rs.open "select * from ks_user where ClubSpecialPower=3",conn,1,3
				 do while not rs.eof
				      Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where TypeFlag=1 and Special=0 and ClubPostNum<=" & rs("PostNum") & " And score<=" & rs("Score") & " order by score desc,ClubPostNum Desc")
					  If Not RSG.Eof Then
						   rs("clubgradeid")=rsg(0)
					  else 
					       rsg.close
						   set rsg=conn.execute("select top 1 gradeid from KS_AskGrade where TypeFlag=1 and special=0")
						   if not rsg.eof then
						   rs("clubgradeid")=rsg(0)
						   else
					       rs("clubgradeid")=0
						   end if
					  End If
					  rs.update
					  RSG.Close
				   rs.movenext
				 loop
				 rs.close
				 
				 for i=0 to ubound(masterarr)
				  rs.open "select top 1 * from ks_user where groupid<>1 and username='" & replace(masterarr(i),"'","") & "'",conn,1,3
				  if not rs.eof then
				     if rs("ClubSpecialPower")<>2 then
					   rs("ClubSpecialPower")=3
					   rs("clubgradeid")=bzgradeid
					   rs.update
					 end if
				  end if
				  rs.close
				  if i=0 then 
				   newmaster="'" & masterarr(i) & "'"
				  else
				   newmaster=newmaster & ",'" & masterarr(i) & "'"
				  end if
				 next
				 set rs=nothing
				 conn.execute("update ks_user set ClubSpecialPower=0 where username not in(" & newmaster & ") and ClubSpecialPower<>2 and groupid<>1")
				 
			   end if
				 conn.execute("update ks_user set ClubSpecialPower=1,clubgradeid=" & admingradeid & " where groupid=1")
				 conn.execute("update ks_user set clubgradeid=" & superbzgradeid & " where ClubSpecialPower=2")
				 
          End Sub
		  
		  '删除
		  Sub GuestBoardDel()
		  		 Dim K, GuestBoardID, Page
				 Page = KS.G("Page")
				 GuestBoardID = Trim(KS.G("GuestBoardID"))
				 GuestBoardID = Split(GuestBoardID, ",")
				 For k = LBound(GuestBoardID) To UBound(GuestBoardID)
						Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
					    RS.Open "Select PostTable,id From KS_GuestBook Where BoardID=" & GuestBoardID(k),conn,1,1
						Do While Not RS.Eof
						 Conn.Execute("Delete From " & RS(0) & " Where TopicID=" & RS(1))
						 RS.MoveNext
						Loop
						RS.Close : Set RS=Nothing
					Conn.Execute ("Delete From KS_GuestBoard Where ID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestBoard Where ParentID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestCategory Where BoardID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestBook Where BoardID=" & GuestBoardID(k))
				 Next
				 Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
				Response.Write ("<script>location.href='KS.GuestBoard.asp?Page=" & Page & "';</script>")
		  End Sub
		  
		  '清空版块帖子
		  Sub DelTopic()
		        Dim GuestBoardID:GuestBoardID = KS.ChkClng(KS.G("GuestBoardID"))
		        Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select PostTable,id From KS_GuestBook Where BoardID=" & GuestBoardID,conn,1,1
				Do While Not RS.Eof
					 Conn.Execute("Delete From " & RS(0) & " Where TopicID=" & RS(1))
					 RS.MoveNext
				Loop
				Conn.Execute ("Delete From KS_GuestBook Where BoardID=" & GuestBoardID)
				Conn.Execute("Update KS_GuestBoard Set TodayNum=0,TopicNum=0,PostNum=0,LastPost='0$2010-8-20 15:18:16$无$$$$$' Where ID=" & GuestBoardID)
				RS.Close : Set RS=Nothing
				Response.Write ("<script>alert('恭喜,该版块数据已被清空!');location.href='KS.GuestBoard.asp';</script>")
		  End Sub
		  
End Class
%>
 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit
%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KSCls
Set KSCls = New Template
KSCls.Kesion()
Set KSCls = Nothing

Class Template
        Private KS
		'===========================================================================
		Private I, totalPut, TemplateSql, KS_T_RS
		Private TemplateType, ChannelID,DomainStr,MaxPerPage
		Private FileItem, CurrPath, ParentPath,InstallDir,Path
		'=============================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KMTL10006") Then                '模板管理的权限检查
			  Call KS.ReturnErr(1, "")
			  Exit Sub
			End If
			
			
			Select Case KS.G("Action")
			 Case "getcontent"
			   Call getcontent()
			   response.end
			 Case "Del"
			   Call TemplateDel()
			 Case "NewPage","Modify"
			   Call AddTemplate()
			 Case "TemplateSave"
			   Call TemplateSave()
			 Case Else
			   Call TemplateList()
			End Select
		End Sub
		
		Sub getcontent()
		 response.write Escape(KS.ReadFromFile(Replace(Replace(UnEscape(Request("TemplateFileName")),"{@TemplateDir}",KS.Setting(3) & KS.Setting(90)),"//","/")))
		End Sub
		
		Sub TemplateList()
		With Response
		InstallDir=KS.Setting(3)
        If CurrPath = "" Then
			ParentPath = ""
			CurrPath= InstallDir & KS.Setting(90)
		Else
			ParentPath = Mid(CurrPath, 1, InStrRev(CurrPath, "/") - 1)
			If ParentPath = "" Then
				ParentPath = Left(InstallDir, Len(InstallDir) - 1)
			End If
		End If
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)		
		
		.Write "<!DOCTYPE html><html>"
		.Write "<head>"
		.Write "<title>模板管理</title>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		.Write "<script language=""JavaScript"">"
		.Write "var ParentPath='" & ParentPath & "';" & vbCrLf
		.Write "var ChannelID='" & ChannelID & "';" & vbCrLf
		.Write "var TemplateType='" & TemplateType & " ';" & vbCrLf
		.Write "</script>"
		.Write "<script language=""JavaScript"" src=""../../KS_Inc/jquery.js""></script>"
		.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
		%>
		<script language="javascript">
		function CreateHtml()
		{   var ids=get_Ids(document.myform);
			if (ids!='')
			   top.openWin('发布选中的自定义页面','Include/RefreshCommonPageSave.asp?RefreshFlag=Folder&PageID='+ids,false,530,60);
			else 
				alert('请选择要发布的自定义页面!');
        }		

		function AddDIYPage()
		{
		$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("自定义单页文件 >> 添加新页面")+'&ButtonSymbol=Go';
		location.href='?Action=NewPage&flag=text';
		}		
		function EditTemplate(id)
		{
		window.parent.parent.frames['MainFrame'].location.href='KS.DIYPage.asp?Action=Modify&TemplateID='+id;
		$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("模板管理中心 >> 编辑页面")+'&ButtonSymbol=TemplateAdd';
		}
		function TextEdit(Flag)
		{
			GetSelectStatus('FolderID','SelectObjID');
		 if (SelectedFile!='')
			if (SelectedFile.indexOf(',')==-1) 
			{
			 location.href='KS.DIYPage.asp?Action=Modify&Flag='+Flag+'&id='+SelectedFile;
			 $(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape("模板管理中心 >> 编辑页面")+'&ButtonSymbol=Gosave';
			}
			else alert('一次只能编辑一个模板文件!')	 
	     else
		 alert('请选择要一个模板!');
		}
		function DelTemplate(id)
		{
		if (confirm('删除后将导致已绑定的信息找不到模板，确认操作吗?'))
		 location="KS.DIYPage.asp?Action=Del&id="+id;
		}
		

		</script>
		<%
		.Write "</head>"
		.Write "<body topmargin=""0"" leftmargin=""0"">"

        .Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick=""AddDIYPage();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>新建页面</span></li>"
		.Write "<li class='parent' onclick=""location.href='../include/refreshcommonpage.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon num'></i>发布自定义页面</span></li>"
		.Write "</ul>"	
		.Write "<div class='pageCont2'>"  
		.Write "<div class='tabTitle'>发布自定义页面</div>"
		.Write "  <form name='myform' id='myform' action='KS.DiyPage.asp' method='get'>"
		.Write "  <input type='hidden' name='action' value='Del'>"
		.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "  <tr align=""center"" class=""sort"">"
		.Write "    <td align=""center"" width=""40"">选择</td>"
		.Write "    <td height=""25"" class=""sort""> <div align=""center""><font color=""#990000"">页面名称</font></div></td>"
		.Write "    <td align=""center"">模板路径</td>"
		.Write "    <td width=""143"" align=""center"">修改时间</td>"
		.Write "    <td width=""267"" class=""sort"">操作管理</td>"
		.Write "  </tr>"
		
		call ShowContent
		
		.Write "</table>"
		
		%>
		<div class="operatingBox">
		<b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> <input type='button' onclick='CreateHtml()' class='button' value='发布选中的自定义页面'> <input type='submit' onclick="return(confirm('此操作不可逆,确定删除吗?'))" class='button' value='删除选中的自定义页面'>
		</div>
		</form>
		<%
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.Write "</div>"
		.Write "</body>"
		.Write "</html>"
		End With
		End Sub
		Sub showContent()
		  MaxPerPage=10   '每页显示条数
		With Response
           Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select * From KS_Template",conn,1,1
		   If RS.Eof And RS.Bof Then
		      .write "<tr><td class='splittd' colspan='10' align='center'>添加添加自定义页面!</td></tr>"
		   Else
		            TotalPut=rs.recordcount
					if CurrentPage < 1 then CurrentPage=1
					rs.move (CurrentPage-1)*MaxPerPage
		            dim i:i=0
			   Do While Not rs.eof
			   
			  .Write "<tr id='u" & rs("templateid") & "' onclick=""chk_iddiv('" & rs("templateid") & "')"" class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			  .Write "  <td class='splittd' align='center'><input onclick=""chk_iddiv('" & RS("templateid") & "')"" type='checkbox' name='id' id='c" & rs("templateid") & "' value='" & rs("templateid") &"'></td>"
			  .Write "  <td class='splittd'><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			  .Write "      <tr>"
			  .Write "        <td height=""20"">"
			  .Write "         <span SelectObjID=" &rs("templateid") & " onDblClick=""TextEdit('');"">"
			  .Write "         <i class=""icon num""></i>"
			  .Write "          <span style=""cursor:default"">" & rs("TemplateName") & "</span></span></td>"
			  .Write "      </tr>"
			  .Write "    </table></td>"
			  .Write "    <td class='splittd'>" & rs("templatefilename") & "</td>"
	
			
			  .Write ("<td align='center' class='splittd'>" & rs("AddDate") & " </td>")
			  .Write ("<td align='center' class='splittd'><a href=""" & ks.setting(3))
			  if ks.setting(94)<>"/" then .Write ks.setting(94)
			  .Write (rs("fsofilename") & """ target=""_blank"" class='setA'>预览页面</a>|<a href='KS.DIYPage.asp?Action=Modify&ID=" & rs("templateid") &"&Flag=text' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=" & server.URLEncode("模板管理中心 >> 编辑页面") &"&ButtonSymbol=GoSave'"" class='setA'>编辑</a>|<a href='KS.DIYPage.asp?Action=Del&id=" & rs("templateid")&"' onclick=""return(confirm('此操作不可逆，确定删除吗？'))"" class='setA'>删除</a></td>")
			  .Write "</tr>"
			  RS.MoveNext
			   i=I+1
			   if i>=MaxPerPage Then Exit Do
			 Loop
		  End If	 
			 
		  End With 
			 	   
	     End Sub
			 
			 '删除模板
		Sub TemplateDel
			Dim IDArr:IDArr=Split(Replace(KS.G("ID")," ",""),",")
			Dim I
			For i=0 to Ubound(IDarr)
			Dim FileName,CurrPath
			Call KS.DeleteFile(KS.Setting(3) & KS.Setting(94) & conn.execute("select fsofilename from ks_template where templateid=" & IDarr(i))(0))
			conn.execute("delete from ks_template where templateid=" & IDarr(i))
			Next
			'Call KS.DeleteFolder(CurrPath & "/" & FileName)
		    Response.Write "<script>window.location.href='KS.DIYPage.asp'</script>"
       End Sub
	   
	   
	   '导入模板
	  Sub AddTemplate()
		Dim Action, TemplateID, ChannelID, TemplateType, TemplateName, FsoFileName, FnameType, TemplateContent, TemplateFileName, TemplateFromFileContent,Action1,FileName
		Dim  InstallDir, TemplateDIr,PageName
		InstallDir  = KS.Setting(3)

		If KS.G("Action")="NewPage" Then
		PageName=""
		Else
		 Dim RSt:Set RSt=Server.CreateoBject("adodb.recordset")
		 rst.open "select * from KS_Template Where TemplateID=" & KS.ChkClng(KS.G("id")),Conn,1,1
		 If RSt.Eof Then
		  Call KS.Alert("参数传递出错!","")
		  exit sub
		 end if
		 PageName=rst("TemplateName")
		 FileName=rst("FsoFileName")
		 TemplateFileName=rst("TemplateFileName")
         TemplateFromFileContent=KS.ReadFromFile(Replace(Replace(TemplateFileName,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90)),"//","/"))
		End If
		
		Response.Write "<!DOCTYPE html><html><head><title>模板管理-添加模板</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		Response.Write "<script src=""../../ks_inc/jquery.js""></script>"
		Response.Write "<script type=""text/JavaScript"" src=""../../KS_Inc/common.js""></script>"
		%>
                    <script language = 'JavaScript'>
					var pos=null;
					function setPos()
					{ if (document.all){
						document.TemplateForm.Content.focus();
						pos = document.selection.createRange();
					  }else{
						pos = document.getElementById("Content").selectionStart;
					  }
					}
					
					function LoadTemplateIn()
					{ 
					    var url='KS.DIYPage.asp';
						$.get(url,{action:"getcontent",TemplateFileName:escape($("#TemplateFileName").val())},function(d){
						  $('#Content').val(unescape(d));
						})
					}	
									
		            function show_ln(txt_ln,txt_main){
			            var txt_ln  = document.getElementById(txt_ln);
			            var txt_main  = document.getElementById(txt_main);
			            txt_ln.scrollTop = txt_main.scrollTop;
			            while(txt_ln.scrollTop != txt_main.scrollTop)
			            {
				            txt_ln.value += (i++) + '\n';
				            txt_ln.scrollTop = txt_main.scrollTop;
			            }
			            return;
		            }
		            function editTab(){
			            
			         }
		            //-->
		            </script>
		<%
		Response.Write "</head>"
		Response.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		Response.Write " <div class=""topdashed sort"">修改模板</div>"
		Response.Write "<div class='pageCont2'>"
		 Response.Write "<table width='100%' height=""350"" style=""background-color:#ffffff;padding-right: 2px;padding-left: 2px;padding-bottom: 0px;"" border='0' align='center' cellpadding='0' cellspacing='1' class='ctable'>"
		 Response.Write " <form name=""TemplateForm"" method=""post"" action=""KS.DIYPage.asp?Action=TemplateSave&id=" & ks.g("id") & """ onSubmit=""return(CheckForm())"">"	
		 		
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30"" style=""text-align:left"" class='pd10 pt10'><b>单页名称：</b><input class=""textbox"" name=""PageName"" type=""text"" id=""PageName"" size=""30"" Value=""" & PageName & """> <font color=red>*</font>如，帮助中心，网站介绍等</td></tr>"
		 
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30"" style=""text-align:left"" class='pd10 pt10'><b>单页模板地址：</b><input class=""textbox"" onpropertychange=""LoadTemplateIn();"" name=""TemplateFileName"" type=""text"" id=""TemplateFileName"" size=""30"" Value=""" & TemplateFileName & """>&nbsp;"
		 	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
		Set KSCls=New ManageCls
		Response.write KSCls.Get_KS_T_C("$('#TemplateFileName')[0]")
		Set KSCls=Nothing
         Response.write  " <input type='button' class='button' value='导入' onclick='LoadTemplateIn();'/>"	 

		 Response.Write "  </td></tr>"
		 
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30"" style=""text-align:left""  class='pd10 pt10'><b>生成文件名称：</b>" & KS.Setting(3) 
		 if ks.setting(94)<>"/" then  Response.Write ks.setting(94)
		  Response.Write "<input name=""FileName"" type=""text"" id=""FileName"" class=""textbox"" size=""24"" Value=""" & FileName & """> <font color=red>*</font> 可带路径，如 ""html/help.html"",""Common/about/help.htm""等</td></tr>"

		 Response.Write "   <tr id=""toplabelarea"" class=""clefttitle"">"
		 Response.Write "	<td valign=""top"" style=""text-align:left""  class='pd10 pt10'><strong>插入标签：</strong>"
		 Response.Write "<select name=""mylabel"" id=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==选择系统函数标签==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select top 200 LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input class='button' type='button' onclick='LabelInsertCode(document.getElementById(""mylabel"").value);' value='插入标签'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;<input type=""button"" class='button' onclick=""javascript:LabelInsertCode();"" value=""选择更多标签"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		 if KS.G("Flag")="text" then
		 response.write "<input type=""button"" value=""复制代码""  class=""button"" onclick=""copy()"">"
		 end if
		 
		 Response.Write " </td>"
		 Response.Write "</tr>"
		 
		 if KS.G("Flag")="text" then
		 Response.Write "   <tr id=""codearea"">"
		 Response.Write "   <td  class='pd10 pt10'>"
		 Response.Write "     <table border='0' width='100%' cellspacing='0' cellspadding='0'>"
		 Response.Write "	  <tr>"
		 Response.Write "       <td valign=""top"" width='50'>"
		 %>
		  <textarea name="txt_ln" id="txt_ln" cols="6" style="width:50px;overflow:hidden;height:423px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;" readonly="">
<% Dim N
For N=1 To 3000
 Response.Write N & vbcrlf
Next
%>
</textarea>
		             </td>
		             <td valign="top"><textarea onclick='setPos()' onkeyup='setPos()' name="Content" rows="2" cols="30" id="Content" onscroll="show_ln('txt_ln','Content')" onKeyDown="editTab()" onChange="TemplateContent.SetContentIni();" style="height:422px;width:100%;"><%=Server.HTMLEncode(TemplateFromFileContent)%></textarea>
</td>
		             </tr>
					 </table>
					 <%
       End If
		if KS.G("Flag")<>"text" then 
		 Response.Write "   <tr id=""editorarea"">"
		 Response.Write "    <td colspan=""2"" width=""100%"" height=""510"">"
		 Response.Write "     </td>"
		 Response.Write "   </tr>"
		 end if
		 
		 Response.Write " </form>"
		 Response.Write "</table>"
		 Response.Write "</div>"
		 Response.Write "</body>"
		 Response.Write "</html>"
			 Conn.Close:Set Conn = Nothing
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "    function copy()" & vbcrlf
		Response.Write "{" & vbcrlf
		Response.Write "document.TemplateForm.Content.value=document.TemplateForm.Content.value;" & vbcrlf
        Response.Write "    document.TemplateForm.Content.select();" & vbcrlf
        Response.Write "    textRange = document.TemplateForm.Content.createTextRange();" & vbcrlf
        Response.Write "    textRange.execCommand(""Copy"");" & vbcrlf
		Response.Write "    alert('恭喜，当前代码已复制到剪贴板!');" & vbcrlf
        Response.Write "}" & vbcrlf
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{" &vbcrlf
		if KS.G("flag")="text" then 
		Response.Write " if(pos==null) {alert('请先定位插入位置!');return false;}"
		end if
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  top.openWin('插入标签','include/InsertLabel.asp?sChannelID=" & ChannelID &"&TemplateType=" & TemplateType &"',false,420,430);"&vbcrlf
		Response.Write "else"
		Response.Write "{ LabelInsert(Val); }" & vbcrlf
		Response.Write "}" & vbcrlf
		
		Response.Write "function LabelInsert(Val){" &vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ "
		Response.Write "if (document.all){ pos.text=Val; }else{" & vbcrlf
		Response.Write "  var obj=$(""#Content"");"&vbcrlf
		Response.Write "  var lstr=obj.val().substring(0,pos);"&vbcrlf
		Response.Write "	   var rstr=obj.val().substring(pos);"&vbcrlf
		Response.Write "	   obj.val(lstr+Val+rstr);			 }"&vbcrlf
		Response.Write " }" & vbcrlf
		Response.Write "}" &vbcrlf
		
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{  if ($('#PageName').val()=="""")"
		Response.Write "     {"
		Response.Write "     alert(""请输入页面名称!"");"
		Response.Write "     $('#PageName').focus();"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf
		Response.Write " if ($('#TemplateFileName').val()=="""")"
		Response.Write "     {"
		Response.Write "     alert(""请导入单页模板!"");"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf
		Response.Write "  if ($('#FileName').val()=="""")"
		Response.Write "     {"
		Response.Write "     alert(""请输入生成的文件名!"");"
		Response.Write "     $('#FileName').focus();"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf

		Response.Write "    document.TemplateForm.submit();"
		Response.Write "    return true;"
		Response.Write "}" & vbCrLf
		Response.Write "</script>" & vbCrLf
	  End Sub
	  
	  Sub TemplateSave()
	  	Dim Action, ChannelID, TemplateType, TemplateName, TemplatConTent, TemplateFileName, TemplateID, FsoFileName, TemplateContent,FileName
		Dim ObjRS, SQLStr, IsDefault, TemplateFilePath, OpStr
		 TemplateName = Trim(Request("PageName"))
		 TemplateContent = Trim(Request("Content"))
		 TemplateFileName = Request("TemplateFileName")   
		 FileName=Request("FileName")
		 If Instr(FileName,".")=0 Then
			Call KS.AlertHistory("保存的文件格式不正确!", -1)
			Set KS = Nothing:Response.End
		 Else
		   Dim FileExt:FileExt=lcase(Split(FileName,".")(1))
		   If FileExt<>"html" and FileExt<>"htm" and FileExt<>"shtml" and FileExt<>"shtm" Then
			Call KS.AlertHistory("保存的文件格式不正确,只能以html,htm,shtml或shtm为扩展名!", -1)
			Set KS = Nothing:Response.End
		   End If
		 End If
		 
		 If InStr(lcase(TemplateFileName),".asp")>0 or InStr(lcase(TemplateFileName),".asa")>0 or InStr(lcase(TemplateFileName),".php")>0 or InStr(lcase(TemplateFileName),".cer")>0 Then
			Call KS.AlertHistory("模板文件名格式不正确!", -1)
			Set KS = Nothing:Response.End
		 End If

				'检查数据正确性
				If TemplateFileName = "" Then
				  Call KS.AlertHistory("您还没有导入模板!", -1)
				  Set KS = Nothing:Response.End
				End If
				
			 TemplateContent = ReplaceBadStr(Replace(Replace(Replace(TemplateContent, "contentEditable=true", ""), KS.GetDomain, "/"), KS.Setting(2), ""))
			If (Instr(TemplateContent,"<%")<>0) or (instr(TemplateContent,"<?")<>0 and instr(TemplateContent,"?>")<>0)  or instr(lcase(TemplateContent),"createobject(""adodb.stream"")")>0 Then
				  Call KS.AlertHistory("您的模板格式不正确,请不要包含可执行代码!", -1)
				  Set KS = Nothing
				  Response.End
			 End If

			 
			  If KS.WriteTOFile(Replace(Replace(TemplateFileName,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90)),"//","/"), TemplateContent) = True Then
			   dim rs:set rs=server.createobject("adodb.recordset")
			   rs.open "select top 1 * from ks_template where templateid=" & ks.chkclng(ks.g("id")),conn,1,3
			   if rs.eof then
			    rs.addnew
			   end if
			    rs("TemplateName")=TemplateName
				rs("TemplateFileName")=TemplateFileName
				rs("fsofilename")=FileName
				rs("adddate")=now
				rs.update
				rs.close
				set rs=nothing
				'生成操作
				Dim KSRCls:Set KSRCls=New Refresh
				Call KSRCls.RefreshCommonPage(TemplateFileName,FileName)
				Set KSRCls=Nothing
			  Response.Write ("<script> top.$.dialog.alert('成功提示:模板保存成功!',function(){ location.href='System/KS.DIYPage.asp';});</script>")
			  Else
				Call KS.AlertHistory("错误提示,可能原因:1.保存失败，模板文件不存在;\n2.目录没有写入权限", -1)
				Set KS = Nothing
			  End If
		End Sub
		Function ReplaceBadStr(Content)
			Dim regEx, Matches, Match
			Set regEx = New RegExp
			regEx.Pattern = "/" & KS.Setting(89) & "([A-Z]|[a-z]|\.|\?|\=|&|;|[0-9])*#"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			ReplaceBadStr = Content
			For Each Match In Matches
				On Error Resume Next
				ReplaceBadStr = Replace(ReplaceBadStr, Match.Value, "#")
			Next
		End Function


 End Class
%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New RefreshSpecial
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshSpecial
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()

			If Not KS.ReturnPowerResult(0, "ref20001") Then                '发布专题的权限检查
				  Call KS.ReturnErr(1, "")
			End If
			If KS.Setting(78)="0" Then  
			  Response.Write "<script>alert('对不起，专题系统没有启用生成静态！');history.back();</script>"
			  Exit Sub
			End If
			With Response
			.Write "<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<title>生成专题管理</title>"
			.Write "</head>"
			.Write "<script language=""JavaScript"" src=""Common.js""></script>"
			.Write "<script>"
			.Write "function CheckTotalNumber() " & vbCrLf
			.Write "{"
			.Write "    if (document.SpecialNewForm.TotalNum.value=='') {alert('请填写专题数量');document.SpecialNewForm.TotalNum.focus();return false;}"
			.Write "    else return true;"
			.Write "}"
			.Write "</script>"
			
			.Write "<body topmargin=""0"" leftmargin=""0"" oncontextmenu=""return false;"">"
			.Write " <div class='pageCont2 pt0'><table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
			.Write "   <tr>"
			.Write "      <td colspan=2 class='tabTitle'>发布专题首页操作</td>"
			.Write "   <tr>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=Index"" method=""post"" name=""AllForm"">"
			.Write "    <tr>"
			.Write "      <td width='100' height='25' align='right' class='clefttitle'> 发布专题首页</td>"
			.Write "      <td>"
			.Write "        &nbsp;<input name=""SubmitAll"" class='button' type=""submit"" value=""发布专题首页 &gt;&gt;"" border=""0"">"
			 .Write "     </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "</table>"
			

			
			.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
			.Write "   <tr>"
			.Write "      <td colspan=2 class='tabTitle'>发布专题页操作</td>"
			.Write "   </tr>"
			.Write "    <form action=""RefreshSpecialSave.asp?Types=Special&RefreshFlag=New"" method=""post"" name=""SpecialNewForm"" onsubmit=""return(CheckTotalNumber())"">"
			.Write "    <tr>"
			.Write "      <td width='100' height='25' align='right' class='clefttitle'> 发布最新上传的</td>"
			.Write "      <td> <input class=""textbox"" name=""TotalNum"" onBlur=""CheckNumber(this,'专题数量');"" type=""text"" id=""TotalNum"" style=""width:20%"" value=""20"">"
			.Write "        个专题"
			.Write "        <input name=""Submit2"" type=""submit"" class='button' value="" 发 布 &gt;&gt;"" border=""0"">"
			.Write "      </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=Special&RefreshFlag=Folder"" method=""post"" name=""ClassForm"">"
			.Write "    <tr>"
			.Write "      <td width='100' height='25' align='right' class='clefttitle'>按专题分类发布</td>"
			.Write "      <td> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "          <tr>"
			.Write "            <td width=""39%""> <select name=""FolderID"" size=10 multiple style=""width:360px;height:146px;"" class='textbox'>"
			   Call GetSpecialClass
			.Write "             </select></td>"
			.Write "            <td width=""61%"">"
			.Write "              <input name=""Submit22"" type=""submit"" class='button' value="" 发布选中的专题 &gt;&gt;"" border=""0"">"
			.Write "              <br> <font color=""#FF0000""> 　<br>"
			.Write "              　提示：<br>"
			.Write "              　按住“CTRL”或“Shift”键可以进行多选</font></td>"
			.Write "         </tr>"
			.Write "        </table></td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=Special&RefreshFlag=All"" method=""post"" name=""AllForm"">"
			.Write "    <tr>"
			.Write "      <td width='100' height='25' align='right' class='clefttitle'>发布所有专题页</td>"
			.Write "      <td>"
			.Write "        &nbsp;&nbsp;<input name=""SubmitAll"" class='button' type=""submit"" value=""发布所有专题 &gt;&gt;"" border=""0"">"
			.Write "      </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "</table>"
			
			
			
			.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
			.Write "   <tr>"
			.Write "      <td colspan=2 class='tabTitle'>发布专题分类</td>"
			.Write "   </tr>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=ChannelSpecial&RefreshFlag=Folder"" method=""post"" name=""ChannelSpecialForm"">"
			.Write "    <tr>"
			.Write "      <td width='100' height='25' align='right' class='clefttitle'>按分类发布</td>"
			.Write "      <td> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "          <tr>"
			.Write "            <td width=""39%"">"
			.Write "            <select name=""FolderID"" size=12 multiple  style=""width:360px;height:146px;"" class='textbox'>"
			
									 Call GetSpecialClass
									 
			.Write "              </select></td>"
			.Write "            <td width=""61%"">"
			.Write "              <input name=""Submit22"" type=""submit"" class='button' value=""发布选中的专题分类页 &gt;&gt;"" border=""0"">"
			.Write "              <br> <font color=""#FF0000""> 　<br>"
			.Write "              　提示：<br>"
			.Write "              　按住“CTRL”或“Shift”键可以进行多选</font></td>"
			.Write "          </tr>"
			.Write "        </table></td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=ChannelSpecial&RefreshFlag=All"" method=""post"" name=""AllForm"">"
			.Write "    <tr class='tdbg'>"
			.Write "      <td width='100' height='25' align='right' class='clefttitle'>发布所有专题分类</td>"
			.Write "      <td style='padding-left:0;'>"
			.Write "        &nbsp;&nbsp;<input name=""SubmitAll"" class='button' type=""submit"" value=""发布所有专题分类 &gt;&gt;"" border=""0"">"
			.Write "      </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "</table>"
			
			.Write "<br><div align='center'><font color=#ff6600>友情提示：发布操作会比较占用系统资源及时间，每次发布时请尽量仅发布最新添加的信息</font></div>"
			.Write "<br><div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		.Write "<div style=""height:30px;text-align:center"">KeSion CMS X" & KS.GetVersion &", Copyright (c) 2006-" & year(now)&" <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
			.Write "</div></body>"
			.Write "</html>"
			End With
			End Sub
			Sub GetSpecialClass()
			           Dim FolderName, TempStr
					   Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
						 RS.Open "Select ClassID,ClassName From KS_SpecialClass Order BY OrderID", Conn, 1, 1
						  If Not RS.EOF Then
							Do While Not RS.EOF
								 FolderName = Trim(RS(1))
								 TempStr = TempStr & "<option value=" & RS(0) & ">" & FolderName & "</option>"
								 RS.MoveNext
							Loop
						  End If
						 RS.Close
					  Set RS = Nothing
					Response.Write TempStr
			End Sub
End Class
%> 

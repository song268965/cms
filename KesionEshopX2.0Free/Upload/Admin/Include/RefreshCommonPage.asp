<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 3.2
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New RefreshCommonPage
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshCommonPage
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		
		If Not KS.ReturnPowerResult(0, "ref20003") Then                '发布自定义页面的权限检查
			  Call KS.ReturnErr(1, "")
		End If
		
		Response.Write "<!DOCTYPE html><html>"
		Response.Write "<head>"
		Response.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>生成自定义页面管理</title>"
		Response.Write "</head>"
		Response.Write "<script language=""JavaScript"" src=""Common.js""></script>" & vbCrLf
		Response.Write "<script>" & vbCrLf
		Response.Write " function CheckForm(FormObj)" & vbCrLf
		Response.Write " {var tempstr='';" & vbCrLf
		Response.Write " for (var i=0;i<FormObj.TempPageID.length;i++){" & vbCrLf
		Response.Write "     var KM = FormObj.TempPageID[i];" & vbCrLf
		Response.Write "    if (KM.selected==true)" & vbCrLf
		Response.Write "       tempstr = tempstr + KM.value+"", """ & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    if (tempstr=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "    alert('请选择您要发布的自定义页面!');" & vbCrLf
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    FormObj.PageID.value=tempstr.substr(0,(tempstr.length-1));" & vbCrLf
		Response.Write "  return true;" & vbCrLf
		Response.Write "  }" & vbCrLf
		Response.Write "</script>" & vbCrLf
		
		Response.Write "<body topmargin=""0"" leftmargin=""0"">"
		Response.Write "<div class='pageCont2 pt0'>"
		Response.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"" bordercolor=""#efefef"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""35"" colspan=""2"" class=""tabTitle"">发布自定义页面操作</td>"
		Response.Write "  </tr>"
		Response.Write "  <form action=""RefreshCommonPageSave.asp?RefreshFlag=Folder"" onsubmit=""return(CheckForm(this))"" method=""post"" name=""ClassForm"">"
		Response.Write "    <tr>"
		Response.Write "      <td width='120' height='25' align='right' class='clefttitle'> 按系统选中页面发布</td>"
		Response.Write "      <td class='pt10'> <table width=""100%"" class='mt20' border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "         <tr>"
		Response.Write "            <td width=""39%"">"
		Response.Write "            <input type=""hidden"" name=""PageID"">"
		Response.Write "              <select name=""TempPageID"" size=12 multiple id=""TempPageID"" style=""width:300px;height:200px;"">"
					  
			  Dim TempStr, ID, FolderName
				Dim ObjRS
				
				Set ObjRS = Server.CreateObject("AdoDB.RecordSet")
				ObjRS.Open ("Select TemplateID,TemplateName From KS_Template Order By TemplateID desc"), Conn, 1, 1
				
				Do While Not ObjRS.EOF
				   ID = Trim(ObjRS(0))
				   FolderName = Trim(ObjRS(1))
				   TempStr = TempStr & "<option value='" & ID & "'>" & FolderName & " </option>"
				ObjRS.MoveNext
				Loop
				ObjRS.Close
				Set ObjRS = Nothing
			   Response.Write (TempStr)
				
		Response.Write "              </select></td>"
		Response.Write "            <td width=""61%"">"
		Response.Write "              <input name=""Submit22"" type=""submit""  class='button' value="" 发布选中的页面 &gt;&gt;"" border=""0"">"
		Response.Write "              <br> <font color=""#FF0000""> 　<br>"
		Response.Write "              　提示：<br>"
		Response.Write "              　按住“CTRL”或“Shift”键可以进行多选</font></td>"
		Response.Write "         </tr>"
		Response.Write "        </table></td>"
		Response.Write "    </tr>"
		Response.Write "  </form>"
		Response.Write "  <form action=""RefreshCommonPageSave.asp?RefreshFlag=All"" method=""post"" name=""AllForm"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""50"" align=""center""> 发布所有的页面</td>"
		Response.Write "      <td height=""50"">"
		Response.Write "        <input name=""SubmitAll""  class='button' type=""submit"" value=""发布所有自定义页面 &gt;&gt;"" border=""0"">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "  </form>"
		Response.Write "</table>"
		Response.Write "<br><br><br><div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS X" &KS.GetVersion &", Copyright (c) 2006-" & year(now)& " <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
End Class
%> 

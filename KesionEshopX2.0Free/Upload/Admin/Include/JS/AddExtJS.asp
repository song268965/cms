<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="../label/LabelFunction.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New AddExtJS
KSCls.Kesion()
Set KSCls = Nothing

Class AddExtJS
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'主体部分
		Public Sub Kesion()
		Dim JSID, JSRS, SQLStr, JSName, Descript, JSConfig, JSFlag, ParentID
		Dim Action, RSCheck, JSFileName, FolderID
		With Response
		Set JSRS = Server.CreateObject("Adodb.RecordSet")
		Action = Request.QueryString("Action")
		JSID = Request("JSId")
		FolderID = Trim(Request("FolderID"))
		If JSID = "" Then
		  Action = "Add"
		Else
		  Action = "Edit"
			Set JSRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_JSFile] Where JSID='" & JSID & "'"
			JSRS.Open SQLStr, Conn, 1, 1
			JSName = Replace(Replace(JSRS("JSName"), "{JS_", ""), "}", "")
			Descript = JSRS("Description")
			FolderID = JSRS("FolderID")
			JSConfig = Server.HTMLEncode(Trim(Replace(JSRS("JSConfig"), "GetExtJS,", "")))
			JSFileName = JSRS("JSFileName")
			JSRS.Close
		End If
		.Write "<!DOCTYPE html><html>"
		.Write "<head>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.Write "<title>新建JS</title>"
		.Write "</head>"
		.Write "<script language=""JavaScript"" src=""../../../ks_inc/jQuery.js""></script>"		
		.Write "<script language=""JavaScript"" src=""../../../ks_inc/Common.js""></script>"
		.Write "<link href=""../Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		%>
		<script>
		var pos=null;
		function setPos()
		{ 
		    if (document.all){
						document.JSForm.JSConfig.focus();
						pos = document.selection.createRange();
					  }else{
						pos = document.getElementById("JSConfig").selectionStart;
					  }
		}
		function LabelInsertCode(Val)
		{ 
		  if(pos==null) { alert('请先定位插入位置!');return false;} 
		  if (Val!=''){
		  
		   if (document.all){ 
		   pos.text=Val; 
		   }else{
		  var obj=$("#JSConfig");
		       var lstr=obj.val().substring(0,pos);
			   var rstr=obj.val().substring(pos);
			   obj.val(lstr+Val+rstr);			
			    }
		}}
		
		</script>
		<%
		

		.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "  <form name=""JSForm"" method=""post"" id=""JSForm"" action=""AddJSSave.asp"">"
		.Write "   <input type=""hidden"" name=""JSType"" value=""1"">"
		.Write "   <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.Write "   <input type=""hidden"" name=""JSID"" value=""" & JSID & """>"
		.Write "   <input type=""hidden"" name=""Page"" value=""" & Request("Page") & """>"
		.Write "  <input type=""hidden"" name=""FileUrl"" value=""AddExtJS.asp"">"
		.Write "    <tr> "
		.Write "      <td height=""123"" valign=""top"">" & ReturnJSInfo(JSID, JSName, JSFileName, FolderID, 3, Descript) & "</td>"
		.Write "    </tr>"
		.Write "    <tr><td colspan=""2"" align=""center"" height=""25"" class=""tableBorder1""><strong>自 定 义 静 态 JS 内 容</strong></td></tr>"
		
		Response.Write "   <tr class=""tableBorder1"" height=25>"
		 Response.Write "	<td  colspan=""2"">"
		 Response.Write "    &nbsp;&nbsp;&nbsp;&nbsp;"
		 Response.Write " <select name=""mylabel"" id=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==选择系统函数标签==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input type='button' onclick='LabelInsertCode($(""#mylabel"").val());' value='插入标签'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;</Td>"
		 Response.Write "      </Tr>"

		
		.Write "    <tr><td colspan=""2"" height=""50""><textarea onclick='setPos()' style=""width:100%"" type=""hidden"" ROWS='17' onkeyup='setPos();SetEditorValue();' onblur='SetEditorValue();' COLS='108' name=""JSConfig"" id=""JSConfig"">" &JSConfig & "</textarea></td></tr>"
		.Write "    <tr>"
		.Write "      <td valign=""top"">"
		.Write "</td></tr>"
		.Write "  </form>"
		.Write "</table>"
		.Write "</body>"
		.Write "</html>"
		.Write "<script language=""JavaScript"">"
		.Write "function SetEditorValue()"
		.Write "{var TempJSConfig=document.JSForm.JSConfig.value;"
		.Write "}"
		.Write "function CheckForm()"
		.Write "{ var form=document.JSForm; "
		.Write "  if (form.JSName.value=='')"
		.Write "   {"
		.Write "    alert('请输入JS名称!');"
		.Write "    form.JSName.focus();"
		.Write "    return false;"
		.Write "   }"
		.Write "      if (form.JSFileName.value=='')"
		.Write "      {"
		.Write "       alert('请输入JS文件名');"
		.Write "      form.JSFileName.focus(); "
		.Write "      return false"
		.Write "      }"
		.Write "     if (CheckEnglishStr(form.JSFileName,'JS文件名')==false) "
		.Write "       return false;"
		.Write "     if (!IsExt(form.JSFileName.value,'JS'))"
		.Write "       { alert('JS文件名的扩展名必须是.js');"
		.Write "          form.JSFileName.focus(); "
		.Write "          return false;"
		.Write "       }"
		.Write "  if (form.JSConfig.value=='')"
		.Write "  {"
		.Write "    alert('请输入JS内容!');"
		.Write "    return false;"
		.Write "   }"
		.Write "   form.JSConfig.value='GetExtJS,'+form.JSConfig.value;"
		.Write "   form.submit();"
		.Write "   return true;"
		.Write "}"
		.Write "</script>    "
		End With
		End Sub
End Class
%> 

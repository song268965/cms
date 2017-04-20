<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
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
Set KSCls = New FolderList
KSCls.Kesion()
Set KSCls = Nothing

Class FolderList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Function Kesion()
		Dim CurrPath, FsoObj, FolderObj, SubFolderObj, FileObj, I, FsoItem
		Dim ParentPath, FileExtName, AllowShowExtNameStr
		AllowShowExtNameStr = "htm,html,shtml"
		CurrPath = Request("CurrPath")
		If CurrPath = "" Then CurrPath = "/"
		Set FsoObj = KS.InitialObject(KS.Setting(99))
		Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
		Set SubFolderObj = FolderObj.SubFolders
		Set FileObj = FolderObj.Files
		
		Response.Write "<!DOCTYPE html><html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		Response.Write "<link href='Admin_Style.CSS' rel='stylesheet'>" &vbcrlf
		Response.Write "<script src=""../../ks_inc/jquery.js""></script>" &vbcrlf
		Response.Write "</head><body topmargin='0' leftmargin='0' scroll=yes>"
		Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		Response.Write "  <tr>"
		Response.Write "    <td height='20' class='sort'> <div align='center'><font color='#000000'>名称</font></div></td>"
		Response.Write "    <td height='20' class='sort'> <div align='center'><font color='#000000'>类型</font></div></td>"
		Response.Write "    <td height='20' class='sort'> <div align='center'><font color='#000000'>修改日期</font></div></td>"
		Response.Write "  </tr>"
		
		 For Each FsoItem In SubFolderObj
		 
		Response.Write "  <tr>"
		Response.Write "    <td height='20'>"
		Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
		Response.Write "          <tr title='双击鼠标进入此目录'>"
		Response.Write "            <td> &nbsp;<i class='icon folder'></i><a href=""javascript:OpenFolder('" & FsoItem.name & "');""><span class='FolderItem' Path='" & FsoItem.name & "' onDblClick=""OpenFolder('" & FsoItem.name & "');"" onClick='SelectFolder(this);'>"
		Response.Write FsoItem.name
		Response.Write "             </span></a> </td>"
		Response.Write "          </tr>"
		Response.Write "        </table>"
		Response.Write "      </div></td>"
		Response.Write "    <td height='20'>"
		Response.Write "      <div align='center'>目录</div></td>"
		Response.Write "    <td height='20'>"
		Response.Write "      <div align='center'>" & FsoItem.size & "</div></td>"
		Response.Write "  </tr>"
		  Next
		For Each FsoItem In FileObj
			FileExtName = LCase(Mid(FsoItem.name, InStrRev(FsoItem.name, ".") + 1))
			If KS.CheckFileShowOrNot(AllowShowExtNameStr, FileExtName) = True Then
		
		Response.Write "<tr title='单击选择文件'>"
		Response.Write "    <td height='20'>"
		Response.Write "      <table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		Response.Write "        <tr>"
		Response.Write "          <td>&nbsp;<img src='../../Editor/KSPlus/FileIcon/html.gif' align='absmiddle'/> <span class='FolderItem' File='" & FsoItem.name & "' onDblClick=""parent.SelectFile('" & replace(FsoItem.name,"'","\'") & "');"" onClick=""SelectFile(this,'" & replace(FsoItem.name,"'","\'") & "');"">"
		Response.Write FsoItem.name
		Response.Write "            </span></td>"
		Response.Write "        </tr>"
		Response.Write "      </table>"
		Response.Write "    </td>"
		Response.Write "    <td height='20'> <div align='center'>"
		Response.Write FsoItem.Type
		Response.Write "      </div></td>"
		Response.Write "    <td height='20'> <div align='center'>"
		Response.Write FsoItem.DateLastModified
		Response.Write "      </div></td>"
		Response.Write "  </tr>"
		
			End If
		Next
		
		Response.Write "</table></body></html>"
		
		Set FsoObj = Nothing
		Set FolderObj = Nothing
		Set FileObj = Nothing
		
		Response.Write "<script language='JavaScript'>"
		Response.Write "var CurrPath='" & CurrPath & "';"
		Response.Write "var FileName='';"
		Response.Write "function SelectFile(Obj,file)"
		Response.Write "{"
		Response.Write "    for (var i=0;i<document.all.length;i++)"
		Response.Write "    {"
		Response.Write "        if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';"
		Response.Write "    }"
		Response.Write "    Obj.className='FolderSelectItem';"
		Response.Write "    FileName=file;"
		Response.Write "}"
		Response.Write "function SelectFolder(Obj)"
		Response.Write "{   FileName='';"
		Response.Write "    for (var i=0;i<document.all.length;i++)"
		Response.Write "    {"
		Response.Write "        if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';"
		Response.Write "    }"
		Response.Write "    Obj.className='FolderSelectItem';"
		Response.Write "}"
		Response.Write "function OpenFolder(Obj)"
		Response.Write "{ "
		Response.Write "    var SubmitPath='';"
		Response.Write "    if (CurrPath=='/') SubmitPath=CurrPath+Obj;"
		Response.Write "    else SubmitPath=CurrPath+'/'+Obj;"
		Response.Write "    location.href='FolderList.asp?CurrPath='+SubmitPath;"&vbcrlf
		Response.Write "    AddFolderList(parent.document.getElementById('FolderSelectList'),SubmitPath);"&vbcrlf
		Response.Write "}"
		Response.Write "function AddFolderList(SelectObj, Label)"&vbcrlf
		Response.Write "{"&vbcrlf
		Response.Write " if (!SearchOptionExists(Label)) { "&vbcrlf
        Response.Write "           jQuery('#FolderSelectList', parent.document).append(""<option selected value='"" + Label + ""'>"" + Label + ""</option>"");  " 
        Response.Write "  } "&vbcrlf
		Response.Write "} "&vbcrlf
		Response.Write "function SearchOptionExists(SearchText) "&vbcrlf
		Response.Write "{"
        Response.Write "var b = false; "&vbcrlf
        Response.Write "    jQuery('#FolderSelectList option', parent.document).each(function() { "&vbcrlf
        Response.Write "        if (jQuery(this).text() == SearchText) { "&vbcrlf
        Response.Write "            jQuery(this).attr(""selected"", ""true""); "&vbcrlf
        Response.Write "            b = true; "&vbcrlf
        Response.Write "            return; "&vbcrlf
        Response.Write "        } "&vbcrlf
        Response.Write "    }); "&vbcrlf
        Response.Write "    return b; "&vbcrlf
		Response.Write "} "&vbcrlf
		Response.Write "</script>"
		End Function
End Class
%> 

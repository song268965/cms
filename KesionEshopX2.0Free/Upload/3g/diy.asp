<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/3GCls.asp"-->
<!--#include file="../api/cls_api.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New ListCls
KSCls.Kesion()
Set KSCls = Nothing

Class ListCls
        Private KS,F_C,KSR
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#include file="include/function.asp"-->
		<%
		Public Sub Kesion()
		 dim id:id=KS.ChkClng(request("id"))
		 if id=0 then ks.die "error"
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from KS_WapTemplate Where ID='" & ID & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   rs.close:set rs=nothing
		   KS.Die "error!"
		 End If
		 F_C = rs("TemplateContent")
		 F_C = Replace(F_C&"","{$TemplateName}",rs("TemplateName"))
		 RS.CLose
		 Set RS=Nothing
		 
		 InitialCommon
		 F_C=KSR.KSLabelReplaceAll(F_C)
		 KS.Die F_C
			 
		End Sub
End Class
%>

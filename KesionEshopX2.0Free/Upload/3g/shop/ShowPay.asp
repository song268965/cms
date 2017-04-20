<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit
response.Buffer=true
%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../include/3gCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New ShowPay
KSCls.Kesion()
Set KSCls = Nothing

Class ShowPay
        Private KS, KSRFObj
		Private F_C
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		%>
		<!--#include file="../include/function.asp"-->
		<%
		Public Sub Kesion()
		      FCls.RefreshType = "ShopPay"
		       F_C = KSRFObj.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(5,10) &"/showpay.html")
				InitialCommon
				F_C = KSRFObj.KSLabelReplaceAll(F_C)
		  Response.write F_C
	   End Sub

End Class
%>

 

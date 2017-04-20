<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit
response.Buffer=true
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
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
		Private FileContent
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      FCls.RefreshType = "ShopPay"
		        If KS.Setting(126)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				FileContent = KSRFObj.LoadTemplate(KS.Setting(126))
				FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
		  Response.write FileContent
	   End Sub

End Class
%>

 

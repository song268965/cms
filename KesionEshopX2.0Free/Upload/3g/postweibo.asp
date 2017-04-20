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
        Private KS,F_C,KSR,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
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
		 If KSUser.UserLoginChecked=FALSE Then Response.Redirect("login.asp")
		 F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/weibo/post.html")
		 InitialCommon 
		 Dim MaxLen,CheckJS
		 if KS.ChkClng(KS.SSetting(50))=0 or KS.ChkClng(KS.SSetting(34))>255 then MaxLen="255" else MaxLen=KS.ChkClng(KS.SSetting(34))
		 F_C=Replace(F_C,"{#MaxLen}",MaxLen)
		 F_C=Replace(F_C,"{#ShowSynchronizedOption}",KSUser.ShowSynchronizedOption(CheckJS))
		 F_C=Replace(F_C,"{#CheckJS}",CheckJS)
		 
		 F_C=KSR.KSLabelReplaceAll(F_C) 
		 KS.Die F_C
		End Sub
End Class
%>

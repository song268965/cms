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
		 If KS.C("UserName")<>"" and KS.C("PassWord")<>"" then response.redirect "user.asp"
		 if request("url")<>"" then
		 GCls.ComeUrl=request("url")
		 end if
		 F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/login.html")    
		 InitialCommon
		 FCls.RefreshType = "userlogin" '设置刷新类型，以便取得当前位置导航等
		 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		 
         F_C=Replace(F_C,"{$IsShowVerifyCode}",KS.ChkClng(KS.Setting(34)))
		 If KS.ChkClng(KS.Setting(34))=1 Then
         F_C=Replace(F_C,"{$ShowVerifyCode}","")
		 Else
         F_C=Replace(F_C,"{$ShowVerifyCode}"," style=""display:none""")
		 End If
		 
		 Dim LoginStr
		 If cbool(API_QQEnable) Then
		   LoginStr="<li><a title=""使用qq账号登录"" href=""" & KS.GetDomain & "api/qq/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_qq_big.png"" align=""absmiddle""/>QQ</a></li>"
		 End If
		 If cbool(API_WeiXinEnable) Then
		   LoginStr=LoginStr & "<li><a title=""使用微信账号登录"" href=""" & KS.GetDomain & "api/weixin/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_weixin_big.png"" align=""absmiddle""/>微信</a></li>"
		 End If
		 If cbool(API_SinaEnable) Then
		   LoginStr=LoginStr & "<li><a title=""使用新浪微博账号登录"" href=""" & KS.GetDomain & "api/sina/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_sina_big.png"" align=""absmiddle""/>微博</a></li>"
		 End If
		 If cbool(API_AlipayEnable) Then
		   LoginStr=LoginStr & "<li><a title=""使用支付宝登录"" href=""" & KS.GetDomain & "api/alipay/alipay_auth_authorize.asp""><img src=""" &KS.GetDomain & "images/default/icon_alipay_big.png"" align=""absmiddle""/>支付宝</a></li>"
		 End If
		 If LoginStr<>"" Then LoginStr="" & LoginStr
		 F_C=Replace(F_C,"{$ShowQuickLogin}","<ul>" & LoginStr & "</ul>")
		 
		 
		 F_C=KSR.KSLabelReplaceAll(F_C)
		 KS.Die F_C
			 
		End Sub
End Class
%>

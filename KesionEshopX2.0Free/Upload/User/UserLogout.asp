<%@language=vbscript codepage="65001" %>
<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../API/cls_api.asp"-->
<!--#include file="../api/uc_client/client.asp"-->
<%
Dim UserName,PassWord
UserName=KS.C("UserName")
If UserName<>"" And Not IsNull(UserName) Then
Conn.Execute("Update KS_User Set isonline=0 Where UserName='" & UserName & "'")
End If
If cbool(EnabledSubDomain) Then
	Response.Cookies(KS.SiteSn).domain=RootDomain					
Else
    Response.Cookies(KS.SiteSn).path = "/"
End If
Response.Cookies(KS.SiteSn)("UserName") = ""
Response.Cookies(KS.SiteSn)("Password") = ""
Response.Cookies(KS.SiteSn)("RndPassword")=""
Response.Cookies(KS.SiteSn)("PowerList")=""
Response.Cookies(KS.SiteSn)("AdminName")=""
Response.Cookies(KS.SiteSn)("AdminPass")=""
Response.Cookies(KS.SiteSn)("SuperTF")=""
Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
Response.Cookies(KS.SiteSn)("ModelPower")=""
Session(KS.SiteSN&"UserInfo")=""
session.Abandon()

'-----------------------------------------------------------------
'系统整合
'-----------------------------------------------------------------
If API_Enable Then
	response.write uc_user_synlogout() '返回javascript分别调用各个应用进行退出
End If

Dim Url
    If trim(Request.ServerVariables("http_referer"))="" Then 
	 Url="/"
    elseif instr(Lcase(Request.ServerVariables("HTTP_REFERER")),"index.asp")>0 then
	 Url="../"
	else
     Url=Request.ServerVariables("http_referer")
	end if
	If API_Enable Then
	  Response.Write "<script>setTimeout(""location.href='" & Url & "';"", 100 );</script>"
	Else
	  Response.Redirect(Url)
	End If
'-----------------------------------------------------------------

Set KS=Nothing
%> 

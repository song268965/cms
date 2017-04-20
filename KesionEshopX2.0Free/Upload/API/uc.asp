<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../plus/MD5.asp"-->
<!--#include file="cls_api.asp"-->
<!--#include file="uc_client/config.asp"-->
<%

if not API_Enable Then KS.die "error"

Dim tget, ttime, code, action
code = Request.QueryString("code")


code = uc_authcode(code,"DECODE",UC_KEY)



'测试
dim content:content=  code 
dim stm:set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以文本模式读取
		stm.mode=3
		stm.charset="utf-8"
		stm.open
		stm.WriteText content
		stm.SaveToFile server.MapPath("/123.txt"),2 
		stm.flush
		stm.Close
		set stm=nothing
		
Set tget = parse_str(code)
		
		
If Len(code) < 5 Then
    Response.write "Invalid Request"
	Response.End()
End If
ttime = tget("time")
If Not IsNumeric(ttime) Or ttime = "" Then
    Response.write "Invalid Request"
	Response.End()
End If
ttime = DateAdd("s",ttime,"1970-01-01 08:00:00")
ttime = DateDiff("s",ttime,Now())
If CInt(ttime) > 60 Then '一分钟内有效，要求服务器上的时间相差不能超过1分钟
    Response.write "Authorization has expiried"
	Response.End()
End If
'验证部分，不用修改 结束
action = tget("action")
Dim ids, uid, oldusername, newusername, username, password, orgpassword, salt
Select Case action
    Case "test"
        Response.write 1
    Case "synlogin"  '登录
		Response.Addheader "P3P","CP=""CURa ADMa DEVa PSAo PSDo OUR BUS UNI PUR INT DEM STA PRE COM NAV OTC NOI DSP COR"""
	    username=tget("username")
		if ks.isnul(username) then ks.die "error"
		dim rs:set rs=conn.execute("select top 1 [username],[password] from KS_User Where UserName='" & KS.DelSQL(username) & "'")
		if not rs.eof then
		  password=rs(1)
        end if
		rs.close
		set rs=nothing
		if not ks.isnul(password) then
		 Call DoLogin(userName,Password)
		end if
		
        Response.write 1  '最后返回 成功
	Case "synlogout" '退出
		Response.Addheader "P3P","CP=""CURa ADMa DEVa PSAo PSDo OUR BUS UNI PUR INT DEM STA PRE COM NAV OTC NOI DSP COR"""
		UserName=KS.C("UserName")
		If UserName<>"" And Not IsNull(UserName) Then
		Conn.Execute("Update KS_User Set isonline=0 Where UserName='" & KS.DelSQL(UserName) & "'")
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


        Response.write 1  '最后返回 成功
End Select


Function parse_str(str)
    Dim objData, aryData, i, aryT
    Set objData = Server.CreateObject("Scripting.Dictionary")
    aryData = Split(str,"&")
    For i = 0 To UBound(aryData)
        aryT = Split(aryData(i), "=")
        If UBound(aryT) > 0 Then
            objData.add aryT(0), aryT(1)
        Else
            objData.add aryT(0), ""
        End If
    Next
    Set parse_str = objData
    Set objData = Nothing
End Function

%>
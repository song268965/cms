<!--#include file="../../plus/md5.asp"-->
<!--#include file="config.asp"-->
<%
If EnabledSubDomain Then
	Response.Cookies(KS.SiteSn).domain=RootDomain					
Else
    Response.Cookies(KS.SiteSn).path = "/"
End If

call weixin_callback()
function weixin_callback()
    if(lcase(Request("state")) = lcase(Session("state"))) Then
	    Dim token_url,result
		
		token_url = "https://api.weixin.qq.com/sns/oauth2/access_token"
		
        result = file_get_contents(token_url,"get","appid=" & AppID &"&secret=" &AppKey &"&code=" & KS.CheckXSS(REQUEST("code")) & "&grant_type=authorization_code")
		
		
		if instr(result,"errcode")<>0 then
			dim obj:set obj = getjson(result)
			if isobject(obj) Then
			  ks.echo "<h3>error:</h3>" & obj.errcode
			  ks.echo "<h3>msg:</h3>" & obj.errmsg
			End If
			set obj=nothing
			ks.die ""
		end if
		if result<>"" then
			set obj = getjson(result)
			Response.Cookies(KS.SiteSn).Expires = Date + 365
			Response.Cookies(KS.SiteSn)("access_token") = obj.access_token
			Response.Cookies(KS.SiteSn)("weixinopenid") = obj.openid
		else
		  ks.die "error!"
		end if
    Else 
        KS.Echo "The state does not match. You may be a victim of CSRF."
    End If
End Function



response.write "<div style='margin-top:90px;color:#666;font-size:16px;text-align:center;'><img src='" & KS.GetDomain &"images/default/loadingAnimation.gif'/><br/><br/>正在登录中，请稍候！！！如果长时间没有反应请<a href=""weixinbind.asp"" target=""parent"" style='color:red'>点此跳转</a>。</div>"

if ks.isnul(ks.c("weixinopenid")) then 
    ks.die "没有返回openid!"
	set ks=nothing
	closeconn
else
  set ks=nothing
  closeconn
  response.Write "<script>top.location.href='weixinbind.asp';</script>"
  response.Redirect("weixinbind.asp")
end  if
%>
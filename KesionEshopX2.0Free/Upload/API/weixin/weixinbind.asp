<!--#include file="../../plus/md5.asp"-->
<!--#include file="config.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
dim openid,username,password,action,rs
if ks.isnul(KS.C("weixinopenid")) then ks.die "没有返回openid!"
action=KS.S("Action")
If Action="check" Then
	openid=ks.s("openid")
	username=trim(ks.r(ks.s("username")))
	password=md5(KS.R(ks.s("password")),16)
	if ks.c("weixinopenid")<>openid then
	  ks.die "请不要非法绑定!"
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open "select top 1 * from ks_user where username='" & username & "' and password='" & password & "'",conn,1,1
	if rs.eof and rs.bof then
	  rs.close:set rs=nothing
	  ks.die "<script>alert('对不起，您输入的账号不存在或是密码不正确，请重输!');history.back(-1);</script>"
	else
	    rs.close
		set rs=nothing
		'绑定到已有账号
		conn.execute("update ks_user set weixinopenid='" & openid & "' where username='" & username & "'")
		'调用登录
		Call DoLogin(username,password)
	end if
ElseIf Action="doreg" Then
        Call DoRegSave(4)
Else
    '===================绑定处理=================================
		dim msg,nickname,figureurl,sex
		dim resultxml:resultxml=get_user_info(4,ks.c("access_token"),KS.C("weixinopenid"))
		dim obj:set obj = getjson(resultxml)
		if instr(resultxml,"errcode")<>0 then
		   if isobject(obj) Then
		    ks.die obj.errmsg
		   else
		    ks.die "error!"
		   end if
		Else
			if isobject(obj) Then
			  nickname=obj.nickname
			  figureurl=obj.headimgurl
			  sex=obj.sex
			  if sex="1" then sex="男" else sex="女"
			End If
			set obj=nothing
		end if

		set rs=conn.execute("select top 1 * from ks_user where weixinopenid='" & ks.delsql(ks.c("weixinopenid")) & "'")
	    if rs.eof and rs.bof then
		     rs.close
			 set rs=nothing
		  if ks.c("username")<>"" and ks.c("password")<>"" then '如果当前会员是登录状态的，直接绑定
			 Conn.Execute("Update KS_User Set weixinopenid='" & ks.c("weixinopenid") & "' where username='" & KS.DelSQL(ks.c("username")) & "'")
			 Session(KS.SiteSN&"UserInfo")=""
			 Response.Redirect("../../user/user_bind.asp")
		  else
		   Call DoBind("用微信登录成功",nickname,figureurl,sex,ks.c("openid"))
		  end if
		Else
			 username=rs("username")
			 password=rs("password")
			 rs.close
			 set rs=nothing
			 Call DoLogin(username,password)
		end if
		
	'=============================================================
End If
set ks=nothing
closeconn
%>
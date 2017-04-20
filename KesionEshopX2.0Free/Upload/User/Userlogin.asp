<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../API/cls_api.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KSCls
Set KSCls = New UserLogin
KSCls.Kesion()
Set KSCls = Nothing

Class UserLogin
        Private KS
		Private KSUser,Action
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		 Action=KS.S("Action")
		 If Action="3g" Then
		    Get3GLogin()
			Exit Sub
		 ElseIf Action="script" Then
		  Call GetLoginByScript()
		  Exit Sub
		 ElseIf Action="checklogin" Then
		  Call CheckUserIsLogin()
		  Exit Sub
		 ElseIf Action="PoploginStr" Then
		  GetPoploginStr()
		  Exit Sub
		 End If
		%>
		<!DOCTYPE html>
<html>
<head>
<title>会员登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%If KS.S("Action")<>"Poplogin" then%>
<style type="text/css">
<!--
.textbox{BACKGROUND-COLOR: #ffffff;BORDER: #ccc 1px solid;COLOR: #999;HEIGHT: 22px;line-height:22pxborder-color: #666666 #666666 #666666 #666666; font-size: 9pt;FONT-FAMILY: verdana;}
TD{FONT-FAMILY:宋体;FONT-SIZE: 9pt;line-height: 130%;}
a{text-decoration: none;} /* 链接无下划线,有为underline */
a:link {color: #000000;} /* 未访问的链接 */
a:visited {color: #333333;} /* 已访问的链接 */
a:hover{COLOR: #AE0927;} /* 鼠标在链接上 */
a:active {color: #0000ff;} /* 点击激活链接 */
.logintitle{font-size:14px;color:#336699;font-weight:bold}
#PopLogin td{font-size:14px;line-height:180%}
#PopLogin td a{color:#336699;text-decoration:underline}
#PopLogin td span{color:#5F5C67;font-size:13px}
#PopLogin td input{margin:2px}
.btn{border-color:#3366cc;margin-right:1em;color:#fff;background:#3366cc;}
.btn{border-width:1px;cursor:pointer;padding:.1em 1em;*padding:0 1em;font-size:9pt; line-height:130%; overflow:visible;}
-->
</style>
<%end if%>
<%
If EnabledSubDomain Then
 response.write "<script>document.domain=""" & RootDomain &""";</script>" &vbcrlf
end if
%>
<script src="../ks_inc/jquery.js"></script>
<script src="../ks_inc/common.js"></script>
<script language="javascript">
function CheckForm(){
	var username=document.loginform.Username.value;
	var pass=document.loginform.Password.value;
	if (username=='')
	{
	  KesionJS.Alert('请输入用户名!','document.loginform.Username.focus();');
	  return false;
    }
	if (pass=='')
	{
	  KesionJS.Alert('请输入登录密码!','document.loginform.Password.focus();');
	  return false;
	 }
	 <% If KS.Setting(34)="1" Then%>
	 if (document.loginform.Verifycode.value==''){
	  KesionJS.Alert('请输入验证码!','document.loginform.Verifycode.focus();');
	  return false;
	 }
	 <%End If%>
	 return true;
}
var seccodefocus=0;
function getverifycode() {
    var obj=document.getElementById("vcodetext_menu");
	obj.style.top = (-parseInt(obj.style.height)) + 'px';
	obj.style.left = '0px';
	obj.style.display = '';
	document.getElementById('vcodeimg').src =document.getElementById('vcodeimg').src.split('?')[0]+'?time=' + Math.random();
	$("#vcodetext_menu").fadeOut('fast').fadeIn('fast');
}</script>
</head>
<body leftmargin="0" topmargin="0" style="background-color:transparent;<%If KS.S("Action")="Poplogin" then response.write "background:url(images/loginbg.png) repeat-x;"%>">
		<%
		If KS.S("Action")="Top" Then
		   Call Login1()
		ElseIf KS.S("Action")="Poplogin" Then
		   Call PopLogin()
		Else
		   Call Login2()
		 End If
		End Sub
		
		'script方式调用
		Sub GetLoginByScript()
           If KSUser.UserLoginChecked=false Then
		    %>
			function CheckForm(){
				if (document.loginform.username.value==''||document.loginform.username.value=='UID/用户名/Email'){$.dialog.alert('请输入登录用户名！',function(){document.loginform.username.focus();});return false;}
				if (document.loginform.password.value==''){$.dialog.alert('请输入登录密码！',function(){document.loginform.password.focus();});return false; }
				 <% If KS.Setting(34)="1" Then%>
				 if (document.loginform.Verifycode.value==''){ $.dialog.alert('请输入验证码！',function(){document.loginform.Verifycode.focus();});return false;}
				 <%End If%>
				 return true;}
			<%
		    KS.Echo "document.write('<form name=""loginform"" id=""loginform"" method=""POST"" action=""" & KS.GetDomain & "user/checkuserlogin.asp""><img src="""&ks.getdomain&"user/images/user.png"" align=""absmiddle"" />用户名 <input type=""text"" style=""padding:1px;color:#999;"" onfocus=""if(this.value==\'UID/用户名/Email\'){this.value=\'\';}"" onblur=""if(this.value==\'\'){this.value=\'UID/用户名/Email\';}"" value=""UID/用户名/Email"" tabindex=""1"" maxlength=""30"" name=""username"" id=""username"" size=""14"" class=""textbox""/>&nbsp;<img src="""&ks.getdomain&"user/images/lock.png"" align=""absmiddle"" />密码 <input tabindex=""2"" style=""padding:1px;FONT-FAMILY: verdana;"" type=""password"" maxlength=""30"" name=""password"" size=""8"" id=""password"" class=""textbox""/>&nbsp;');"
			 If KS.Setting(34)="1" Then
			  KS.Echo "document.write('<span>验证码 </span>');writeVerifyCode('" & KS.GetDomain & "');"
			 End If
			 KS.Echo "document.write('&nbsp;<input align=""absmiddle"" type=""image"" src=""" & KS.GetDomain & "images/login.gif"" onclick=""return(CheckForm())""  class=""lgbtn""/>&nbsp;<a href=""" & KS.GetDomain & "user/reg/"" target=""_self""><img src="""&ks.getdomain&"images/reg.gif"" align=""absmiddle"" /></a>&nbsp;<a href=""" & KS.GetDomain & "user/getpassword.asp""><img src="""&ks.getdomain&"user/images/info.png"" align=""absmiddle"" />找回密码</a>');"
			 if cbool(API_QQEnable) then
			  KS.Echo "document.write(' <a title=""使用qq号登录"" href=""" & KS.GetDomain & "api/qq/redirect_to_login.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_qq.png"" /></a>');"
			 End If
			 if cbool(API_WeiXinEnable) then
			  KS.Echo "document.write(' <a title=""使用微信扫码登录"" href=""" & KS.GetDomain & "api/weixin/redirect_to_login.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_weixin.png"" /></a>');"
			 End If
			 if cbool(API_SinaEnable) Then
			  KS.Echo "document.write(' <a title=""使用新浪微博账号登录"" href=""" & KS.GetDomain & "api/sina/redirect_to_login.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_sina.png"" /></a>');"
             End If
			 if cbool(API_AlipayEnable) Then
			  KS.Echo "document.write(' <a title=""使用支付宝登录"" href=""" & KS.GetDomain & "api/alipay/alipay_auth_authorize.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_alipay.png"" /></a>');"
             End If
			 KS.Echo "document.write('</form>');"
		   Else
		    if not KS.IsNul(Session("PopTips")) Then
		    %>
		   function popShowMessage(tipstr){
			 if(document.readyState=="complete"){ 
				if (tipstr==null || tipstr=='')return;
				$.dialog.tips('<div id="tipsmessage" style="font-size:14px;color:#ff6600;margin-top:12px;text-align:center">'+tipstr+'</div>',4,'face-smile.png',function(){}); 
				}else{ 
			  setTimeout(function(){popShowMessage(tipstr);},10); 
			  }
			}
		   <%
		     ks.echo "popShowMessage('" & Session("PopTips") &"');"
			 Session("PopTips")=""
			End If
		     KS.Echo "document.write('您好！<span style=""color:red"">" & KSUser.UserName & "</span>,欢迎来到会员中心!【<a href=""" & KS.GetDomain & "user/"">会员中心</a>】【<a href=""" & KS.GetDomain & "user/user_Message.asp?action=inbox"">短消息"& GetMailTips()& "</a>】【<a href=""" & KS.GetDomain & "User/UserLogout.asp"">退出</a>】');"
		   End If
		End Sub
		
		Sub CheckUserIsLogin()
		  If KSUser.UserLoginChecked=false Then
		    If KS.S("S")="1" Then
		     KS.Echo "var user={'loginstr':'<a href=""javascript:void(0)"" onclick=""ShowPopLogin()"">请登录</a><a href=""" & KS.GetDomain & "user/reg/"" target=""_blank"">注册</a>'}"
			Else
			 KS.Echo "var user={'loginstr':'<form name=""loginform"" id=""loginform"" method=""POST"" action=""" & KS.GetDomain & "user/checkuserlogin.asp"">用户名 <input type=""text"" maxlength=""30"" name=""username"" id=""username"" size=""12"" class=""textbox""/>&nbsp;密码 <input style=""FONT-FAMILY: verdana;"" type=""password"" maxlength=""30"" name=""password"" size=""12"" id=""password"" class=""textbox""/>&nbsp;"
			 If KS.Setting(34)="1" Then
			  KS.Echo "<span>验证码 </span><span style=""position: relative;""><input name=""Verifycode""  maxlength=""6"" size=""5"" class=""textbox"" onblur=""if(!seccodefocus) {document.getElementById(\'vcodetext_menu\').style.display=\'none\';}""  id=""Verifycode""  onfocus=""showverifycode(0)""  autocomplete=""off""/><div class=""verifybox""  style=""position:absolute;display:none;cursor: pointer;width: 124px; height: 44px;left:0px;top:40px;z-index:10009;padding:0;"" id=""codebox0"" onmouseout=""seccodefocus = 0"" onmouseover=""seccodefocus = 1""><img width=""145"" src=""" & KS.GetDomain & "plus/verifycode.asp?time=0.001"" id=""vcodeimg0"" title=""看不清点这里刷新"" onclick=""showverifycode(0)""/></div></span>&nbsp;"
			 End If
			 KS.Echo "<input align=""absmiddle"" type=""image"" src=""" & KS.GetDomain & "images/login.gif"" onclick=""return(CheckLoginForm())""  class=""lgbtn""/>&nbsp;<a href=""" & KS.GetDomain & "user/reg/"" target=""_self"">注册</a>&nbsp;|&nbsp;<a href=""" & KS.GetDomain & "user/getpassword.asp"">找回密码</a></form>'}"
			End If
		  Else
		    KS.Echo "var user={'loginstr':'<a title=""会员中心"" href=""" & KS.GetDomain & "user/"" class=""admin""><span class=""uname"">" & KSUser.UserName & "</span></a> <a href=""" & KS.GetDomain & "user/user_message.asp"" class=""mess""><span id=NewMessage>(" & GetMailTips & ")</span></a><a href=""javascript:;"" onclick=""$.dialog.confirm(\'确定安全退出吗？\',function(){ top.location.href=\'" & KS.GetDomain & "user/UserLogout.asp\';},function(){});"" class=""admin"">退出</a>'}"
		  End If
		End Sub
		
		
		Function GetMailTips()
		    Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
			'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)
			 IF MyMailTotal>0 Then 
			  GetMailTips="<font color=""red"">" & MyMailTotal & "</font><bgsound src=""" & KS.GetDomain & "User/images/mail.wav"" border=0>"  
			  Else
			  GetMailTips=0
			 End If
		End Function
		
		Sub PopLogin()
		If KS.C("UserName")<>"" Then KS.Die "<script>top.location.href='index.asp';</script>"
		%>
		  <style type="text/css">
<!--
body{font-family: "\5FAE\8F6F\96C5\9ED1";}
.textbox
{
padding: 9px 9px 9px 34px;
width: 220px;
line-height: 20px;
border: 1px solid #d9e1e5;
border-radius: 3px;
outline: none;

}
TD
{
FONT-SIZE: 9pt;line-height: 130%;
}
a{text-decoration: none;} /* 链接无下划线,有为underline */
a:link {color: #000000;} /* 未访问的链接 */
a:visited {color: #333333;} /* 已访问的链接 */
a:hover{COLOR: #AE0927;} /* 鼠标在链接上 */
a:active {color: #0000ff;} /* 点击激活链接 */
#PopLogin{ margin-top:30px;}
#PopLogin td{font-size:14px;line-height:45px; }
#PopLogin td a{color: #009aff}
#PopLogin td span{}
#PopLogin td input{margin:2px; text-indent:5px; color:#666}
#PopLogin .txt{ font-size:14px;}
#PopLogin .btn{ color:#FFF;cursor:pointer; border:0; width:100%; height:48px; line-height:18px;background-color: #76b600;
border: 1px solid #76b600; font-size:18px;font-family: "\5FAE\8F6F\96C5\9ED1";display: inline-block;
border-radius: 5px;}

.sanf {
text-align: center;
margin-top: 20px;
font-size: 14px;
border-top: 1px solid #d9e1e5;
padding-top: 12px;}

 .login_bottom {
text-align: center;
margin-top: 30px;
font-weight: 500;
font-size: 14px;}
 .login_bottom a{color: #009aff;}
 
 
.user_ico, .pass_ico, .yz_ico, .em_ico {
background: url(images/reg_ico.png) no-repeat;
}
 
 .user_ico {
background-position: 5px 10px;
}
.pass_ico {
background-position: 5px -20px;}
-->
</style>
		  <form action="checkuserlogin.asp" method="post" name="loginform">
			   <div id="LoginPanel" style="width:100%;">
				<table id="PopLogin" cellspacing="0" width="60%" align="center">
						<tr>
							<td>
								<input name="Username" type="text" value="" placeholder="UID/用户名/Email" maxlength="20" id="Username" class="textbox user_ico" />
								<span id="RequiredFieldValidator1" style="color:Red;display:none;"></span>
							</td>
						</tr>
				<tr>
					
					<td>
						<input name="Password" type="password" id="Password" class="textbox pass_ico" placeholder="登录密码" value="" />
						<span id="RequiredFieldValidator2" style="color:Red;display:none;"></span></td>
				</tr>
				<% If KS.Setting(34)="1" Then%>
				<tr id="vefify">
					<td>
							<input name="Verifycode" placeholder="验证码" type="text" id="Verifycode" class="textbox pass_ico" style="width:48px;" />
							<img align="absmiddle" width="125" src="../plus/verifycode.asp" style="cursor:pointer" title="看不清可以点此刷新" id="vcodeimg" onClick="this.src=this.src+'?time=0.001'"/>
				</tr>
               <%End If%>
				<tr>
				   <td>
					   <a href="<%=KS.Setting(3)%>user/getpassword.asp"  target="_blank" style=" float:right">忘记密码</a>
						 <label><span class="txt"> <input name="ExpiresDate" type="checkbox" id="ExpiresDate" value="checkbox">	永久登录</span>
							 <label><div id="ValidationSummary1" style="color:Red;display:none;">
		
			</div></label>
						 </label>  
				   
							
					</td>
							
				</tr>
				<tr>
					<td>
						  <input type="hidden" name="Action" value="PopLogin">
					 <input type="submit" name="Button1" value="登 录" onClick="return(CheckForm())" id="Button1" class="btn" />
					
					</td>
				</tr>
                </table>
                
                <div class="login_bottom">
                        没有账号？<a href="../user/reg/" target="_blank">现在注册</a>
                 </div>       
				
				<div class="sanf">
	    
		           <%
				If cbool(API_QQEnable) or cbool(API_SinaEnable) or cbool(API_AlipayEnable) or cbool(API_WeiXinEnable) then ks.echo "账号通："
				if cbool(API_QQEnable) then
			  KS.Echo " <a title=""使用qq号登录"" href=""" & KS.GetDomain & "api/qq/redirect_to_login.asp"" target=""_blank""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_qq.png"" border=""0"" align=""absmiddle""/></a>"
			 End If
			if cbool(API_WeiXinEnable) then
			  KS.Echo " <a title=""使用微信扫码登录"" href=""" & KS.GetDomain & "api/weixin/redirect_to_login.asp"" target=""_blank""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_weixin.png"" border=""0"" align=""absmiddle""/></a>"
			 End If
			 if cbool(API_SinaEnable) Then
			  KS.Echo " <a title=""使用新浪微博账号登录"" href=""" & KS.GetDomain & "api/sina/redirect_to_login.asp"" target=""_blank""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_sina.png"" border=""0"" align=""absmiddle""/></a>"
             End If
			 if cbool(API_AlipayEnable) Then
			  KS.Echo " <a title=""使用支付宝登录"" href=""" & KS.GetDomain & "api/alipay/alipay_auth_authorize.asp"" target=""_blank""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_alipay.png"" border=""0"" align=""absmiddle""/></a>"
             End If
				%>
		       </div>
</div>
			   
		 </form>

		<%
		End Sub
		
		sub Get3GLogin()
		  Dim Str
		  If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then
		   Str="<a href=""" & KS.GetDomain & "3g/login.asp"" class=""login"">登录</a>" 
		   str=str &" | <a href=""" & KS.GetDomain & "3g/reg.asp"" class=""reg"">注册</a>"
		  Else
		   Str="<a href=""" & KS.GetDomain & "3g/user.asp"" style=""color:#3087C4"">" & KS.C("UserName") & "</a><a onclick=""return(confirm(\'确定退出登录吗？\'));"" href=""" & KS.GetDomain & "3g/Logout.asp"" style="" font-size:0.6rem"">退出</a>"
		  End If
		   KS.Die ("document.write('" & str & "');")
		End Sub
		
		'给$.getScript弹出窗口用
		Sub GetPoploginStr()
		 Dim Str
		 str="<table id=""PopLogin"" style=""font-size:14px;line-height:180%"" width=""100%"" height=""184"" cellpadding=""0"" cellspacing=""0"" border=""0""><tr><td><table border=""0"" width=""95%"" align=""center""><tr><td style=""border-right:solid 1px #cccccc"">没有账号？<a href=""" & KS.GetDomain & "user/reg/"" target=""_blank"">现在注册</a><br/>密码忘了, <a href=""" & KS.GetDomain &"user/getpassword.asp""  target=""_blank"">我要找回</a> <br />"
		 
		 If cbool(API_QQEnable) or cbool(API_SinaEnable) or cbool(API_AlipayEnable) or cbool(API_WeixinEnable) then str=str & "账号通："
			 if cbool(API_QQEnable) then
			  str=str & " <a target=""_top"" title=""使用qq号登录"" href=""" & KS.GetDomain & "api/qq/redirect_to_login.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_qq.png"" border=""0"" align=""absmiddle""/></a>"
			 End If
			 if cbool(API_WeixinEnable) then
			  str=str & " <a target=""_top"" title=""使用微信扫码登录"" href=""" & KS.GetDomain & "api/weixin/redirect_to_login.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_weixin.png"" border=""0"" align=""absmiddle""/></a>"
			 End If
			 if cbool(API_SinaEnable) Then
			  str=str & " <a target=""_top"" title=""使用新浪微博账号登录"" href=""" & KS.GetDomain & "api/sina/redirect_to_login.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_sina.png"" border=""0"" align=""absmiddle""/></a>"
             End If
			 if cbool(API_AlipayEnable) Then
			  str=str & " <a target=""_top"" title=""使用支付宝登录"" href=""" & KS.GetDomain & "api/alipay/alipay_auth_authorize.asp""><img align=""absmiddle"" src=""" & KS.GetDomain &"images/default/icon_alipay.png"" border=""0"" align=""absmiddle""/></a>"
             End If
		 

		 Str=Str & "</td><td style=""text-align:left""><div style=""font-size:14px;color:#336699;font-weight:bold"" class=""logintitle"">用户登录</div><span>用户账号：</span><input type=""text"" name=""Username"" class=""textbox""><br /><span>登录密码：</span><input type=""password"" name=""Password"" class=""textbox""><br/>"
		 If KS.Setting(34)="1" Then
		 str=str & "<span>附加字符：</span><input name=""Verifycode""  maxlength=""6"" size=""5"" class=""textbox"" onblur=""if(!seccodefocus) {document.getElementById(\'codebox0\').style.display=\'none\';}""  id=""Verifycode""   autocomplete=""off""/><div class=""verifybox""  style=""position:absolute;display:none;cursor: pointer;width: 124px; height: 44px;left:0px;top:40px;z-index:10009;padding:0;"" id=""codebox0"" onmouseout=""seccodefocus = 0"" ></div><img style=""cursor:pointer"" width=""96"" src=""" & KS.GetDomain & "plus/verifycode.asp?time=0.001"" id=""vcodeimg0"" title=""看不清点这里刷新"" onclick=""showverifycode(0)""/><br/>"
		End If
		 Str=Str & "<input type=""submit"" class=""btn"" onclick=""return(CheckLoginForm())"" value="" 登 录 "" name=""submit""><input name=""ExpiresDate"" type=""checkbox"" id=""ExpiresDate"" value=""checkbox"">	<span>永久登录</span></td></tr></table></td></tr></table>"
		 %>
			<!DOCTYPE html
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<%
			If EnabledSubDomain Then
			 response.write "<script>document.domain=""" & RootDomain &""";</script>" &vbcrlf
			end if
			%>
			<script src="../ks_inc/jquery.js" type="text/javascript"></script>
			<script src="../ks_inc/common.js" type="text/javascript"></script>
			<style type="text/css">
			.textbox{BACKGROUND-COLOR: #ffffff;
			BORDER: #ccc 1px solid;
			COLOR: #999;
			HEIGHT: 22px;
			line-height:22px
			border-color: #666666 #666666 #666666 #666666; font-size: 9pt;FONT-FAMILY: verdana}
			TD
			{FONT-FAMILY:宋体;FONT-SIZE: 9pt;line-height: 130%;}
			a{text-decoration: none;} /* 链接无下划线,有为underline */
			a:link {color: #000000;} /* 未访问的链接 */
			a:visited {color: #333333;} /* 已访问的链接 */
			a:hover{COLOR: #AE0927;} /* 鼠标在链接上 */
			a:active {color: #0000ff;} /* 点击激活链接 */
			.logintitle{font-size:14px;color:#336699;font-weight:bold}
			#PopLogin td{font-size:14px;line-height:180%}
			#PopLogin td a{color:#336699;text-decoration:underline}
			#PopLogin td span{color:#5F5C67;font-size:13px}
			#PopLogin td input{margin:2px}
			.btn{border-color:#3366cc;margin-right:1em;color:#fff;background:#3366cc;}
			.btn{border-width:1px;cursor:pointer;padding:.1em 1em;*padding:0 1em;font-size:9pt; line-height:130%; overflow:visible;}
			
			</style>
			<script type="text/javascript">
			 function CheckLoginForm(){
				var username=document.loginform.Username.value;
				var pass=document.loginform.Password.value;
				if (username==''){
				  $.dialog.alert('请输入用户名',function(){
				  document.loginform.Username.focus();
				  });
				  return false;}
				if (pass==''){
				  $.dialog.alert('请输入登录密码',function(){
				  document.loginform.Password.focus();
				  });
				  return false;
				 }
				  return true;
			} 
			</script>
			</head>
			<body>
			 <form action='<%=KS.GetDomain%>user/checkuserlogin.asp' method='post' name='loginform'>
			  <input type="hidden" name="parentbox" value="1"/>
			  <input  type="hidden" name="action" value="PopLogin"/>
			   <%=str%>
			 </form>
		    </body>
		  </html>
		 <%
		End Sub
		
		Sub Login1()
			If KSUser.UserLoginChecked() = False Then
			%>

				<table cellspacing="0" cellpadding="0" width="99%" border="0">
				<form name="loginform" action="<%=KS.GetDomain%>User/CheckUserLogin.asp?Action=Top" method="post">
								<tr>
								  <td>用户名：<input class="textbox" size="10" name="Username" />  密 码：<input class="textbox" type="Password" size="10" name="Password"><%if KS.Setting(34)="1" Then%>验证码：<input name="Verifycode" type="text" class="textbox" id="Verifycode" size="6" /><%
				Response.Write " <IMG width=50 height=21 style=""cursor:pointer"" src=""" & KS.GetDomain & "plus/verifycode.asp?n=" & Timer & """ onClick=""this.src='" & KS.GetDomain & "plus/verifycode.asp?n='+ Math.random();"" align=""absmiddle"">"
				end if%> 
								    <input name="loginsubmit" type="image"  onClick="return(CheckForm())" src="<%=KS.GetDomain%>images/login.gif" align="top" />
								    &nbsp;<a href="<%=KS.GetDomain%>user/reg/" target="_blank"><img src="<%=KS.GetDomain%>images/reg.gif"  border="0" align="absmiddle" twffan="done" /></a></td>
								</tr>
							</table>
			<%Else

			%>
			<table cellspacing="0" cellpadding="0" width="99%" border="0">
				<tr>
			     <td height="22" align="center">您好!<font color=red><%=KSUser.UserName%></font>,欢迎来到会员中心!&nbsp;【<a href="<%=KS.GetDomain%>User/index.asp?User_Message.asp?action=inbox" target="_parent">收信箱 <%=GetMailTips()%></a>】&nbsp;【<a href="<%=KS.GetDomain%>User/index.asp" target="_parent">会员中心</a>】&nbsp;【<a href="<%=KS.GetDomain%>User/UserLogout.asp">退出登录</a>】</td>
				</tr>
			</table>
<%End IF
		End Sub
		Sub Login2()
			If KSUser.UserLoginChecked() = False Then
			%>
			 <form name="loginform" action="CheckUserLogin.asp" method="post">
			<table align="center" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td height="25">用户名：
				<input name="Username" style="width:110px" tabindex="1" type="text" class="textbox" id="Username" size="15"></td>
			  </tr>
			  <tr>
				<td height="25">密　码：
				<input name="Password" style="width:110px" tabindex="2" type="password" class="textbox" id="Password" size="16"></td>
			  </tr>
			  <%if KS.Setting(34)="1" Then%>
			  <tr>
				<td height="25">验证码：
				<span style="position: relative;"><input tabindex="3" value="点此输入验证码" name="Verifycode"  maxlength="6" style="width:110px" class="textbox" onBlur="if(!seccodefocus) {document.getElementById('vcodetext_menu').style.display='none';}"  id="Verifycode"  onfocus="if(this.value=='点此输入验证码')this.value='';getverifycode()"  autocomplete="off"/><div class="verifybox"  style="position:absolute;display:none;cursor: pointer;width: 124px; height: 44px;left:0px;top:40px;z-index:10009;padding:0;" id="vcodetext_menu" onMouseOut="seccodefocus = 0" onMouseOver="seccodefocus = 1"><img width="130" src="<%=KS.GetDomain%>plus/verifycode.asp?time=0.001" id="vcodeimg" title="看不清点这里刷新" onClick="getverifycode()"/></div></span> 
				</td>
			  </tr>
			  <%end if%>
			 
			  <tr>
				<td height="25"><div align="center">
				  <input type="submit" name="Submit"  onClick="return(CheckForm())" class="inputButton" value="登录">

				  <input name="ExpiresDate" type="checkbox" id="ExpiresDate" value="checkbox">
			永久登录</div></td>
			  </tr>
			   <tr>
				<td height="25"><div align="center"><img src="<%=KS.GetDomain%>images/losspass.gif" align="absmiddle"> <a href="<%=KS.GetDomain%>User/GetPassword.asp" target="_parent">忘记密码</a> <img src="<%=KS.GetDomain%>images/mas.gif" align="absmiddle"> <a href="<%=KS.GetDomain%>user/reg/" target="_parent">新会员注册</a>    </div></td>
			  </tr>
            </table>
			  </form>
			<%Else
			 dim  ChargeTypeStr
			 if KSUser.ChargeType=1 Then
			   ChargeTypeStr="扣点"
			 elseif KSUser.ChargeType=2 Then
			   ChargeTypeStr="有效期"
			 else
			   ChargeTypeStr="无限期"
			 End If
			%>
			<style>
			.userlogin tr td{ line-height:20px;}
			</style>
			<table align="center" style="margin-top:5px" width="80%" border="0" cellspacing="0" cellpadding="0" class="userlogin">
			<tr><td align="left"><font color=red><%=KSUser.UserName%></font>,
           <%
			If (Hour(Now) < 6) Then
            Response.Write "<font color=##0066FF>凌晨好!</font>"
			ElseIf (Hour(Now) < 9) Then
				Response.Write "<font color=##000099>早上好!</font>"
			ElseIf (Hour(Now) < 12) Then
				Response.Write "<font color=##FF6699>上午好!</font>"
			ElseIf (Hour(Now) < 14) Then
				Response.Write "<font color=##FF6600>中午好!</font>"
			ElseIf (Hour(Now) < 17) Then
				Response.Write "<font color=##FF00FF>下午好!</font>"
			ElseIf (Hour(Now) < 18) Then
				Response.Write "<font color=##0033FF>傍晚好!</font>"
			Else
				Response.Write "<font color=##ff0000>晚上好!</font>"
			End If
			%>&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>计费方式： <strong><%= ChargeTypeStr%></strong> </td></tr>
			<tr><td>经验积分： <strong><%=KSUser.GetUserInfo("Score")%></strong> 分</td></tr>
			<%if KSUser.ChargeType=1 or KSUser.ChargeType=2 then%>
			<% if KSUser.ChargeType=1 then%>
			<tr><td>可用点券： <strong><%=KSUser.GetUserInfo("Point")%></strong> 点</td></tr>
			<%else%>
			<tr><td>剩余天数： <strong><%=KSUser.GetEdays%></strong></td></tr>
			<%end if%>
			<%end if%>
			<tr><td>待阅短信： <strong><%=GetMailTips()%></strong> 条</td></tr>
			<tr><td>登录次数： <strong><%=KSUser.GetUserInfo("LoginTimes")%></strong> 次</td></tr>
            <tr><td nowrap="nowrap">【<a href="<%=KS.GetDomain%>User/index.asp" target="_parent">会员中心</a>】【<a href="<%=KS.GetDomain%>User/UserLogout.asp">退出登录</a>】</td></tr>
			</table>
<%End IF
  End Sub
End Class
%>

 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
Dim KS:Set KS=New Publiccls
Dim id:id=KS.ChkClng(request("id"))
if id=0 then ks.die "<script>alert('参数出错!');window.close();</script>"
dim rs:set rs=conn.execute("select top 1 * from KS_InterView Where ID=" & id)
if rs.eof and rs.bof then
 rs.close
 set rs=nothing
 ks.die "<script>alert('对不起，访谈主题不存在!');window.close();</script>"
end if
if rs("locked")="1" then 
 rs.close
 set rs=nothing
 ks.die "<script>alert('对不起，该访谈已结束!');window.close();</script>"
end if
dim guests:guests=rs("guests")
rs.close
set rs=nothing

Select Case  KS.G("Action")
 Case "LoginCheck"
  Call CheckLogin()
 Case "LoginOut"
  Call LoginOut()
 Case Else
  Call Main()
End Select
Sub Main()
%>
<!DOCTYPE html>
<html>
<head>
<title><%=KS.Setting(0)%>---在线访谈系统</title>
<script type="text/JavaScript" src="../ks_inc/jquery.js"></script>
<script type="text/JavaScript" src="../ks_inc/common.js"></script>
<script type="text/javascript" src="../ks_Inc/lhgdialog.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<style type="text/css"> 
	html{color:#000;font-family:Arial,sans-serif;font-size:12px;}
	h1, h2, h3, h4, h5, h6, h7, p, ul, ol,div,span, dl, dt, dd, li, body,em,i, form, input,i,cite, button, img, cite, strong,    em,label,fieldset,pre,code,blockquote, table, td, th ,tr{ padding:0; margin:0;outline:0 none;}
	img, table, td, th ,tr { border:0;}
	address,caption,cite,code,dfn,em,th,var{font-style:normal;font-weight:normal;}
	select,img,select{font-size:12px;vertical-align:middle;color:#666; font-family:Arial,sans-serif}
	.checkbox{vertical-align:middle;margin-right:5px;margin-top:-2px; margin-bottom:1px;}
	textarea{font-size:12px;color:#666; font-family:Arial,sans-serif}
	table{border-collapse:collapse;border-spacing:0;}
	ul, ol, li { list-style-type:none;}
	a { color:#0082cb; text-decoration:none;}
	a:hover{text-decoration:none;}
	ul:after,.clearfix:after { content: "."; display: block; height: 0; clear: both; visibility: hidden; }/* 不适合用clear时使用 */
	ul,.clearfix{ zoom:1;}
	.clear{clear:both;font-size:0px; line-height:0px;height:1px;overflow:hidden;}/*  空白占位  */
	body {margin:0 auto;font-size:12px; background:#E0F1FB;color:#666;position:relative}
	#wrap{margin-top:80px;}
	.main{width:800px;margin:0px auto;}
	.main_L{width:380px;float:right;background:url(images/linebg.png) no-repeat left center; padding-left:25px;margin-right:17px;display:inline;}
	.tabbox ul{margin-top:10px;}
	.tabbox li{padding:3px 0px 5px; position:relative;}
	.tabbox li.btn{padding-top:10px;padding-left:98px;}
	.tabbox .label{width:350px;height:38px;background:url(images/textbg.png) right -45px no-repeat;  } 
	.tabbox .label:hover{background:url(images/textbg.png) right 0px no-repeat;  } 
	.labelhover{width:350px;height:38px;background:url(images/textbg.png) right 0px no-repeat;  } 
	.tabbox label{font-size:14px;color:#666} 
	.tabbox .input,.tabbox .textinput{width:230px;height:26px;line-height:26px; padding:2px;padding-left:5px;border:0px;margin-top:3px;margin-left:10px;background-color:transparent; font-family:Verdana, Arial, Helvetica, sans-serif;}
	.tabbox .textinputhover,.tabbox .textinputhover{border:1px solid #aaa;}
	.regsubmit{width:182px;height:53px;border:0px none; background:url(images/reg_btn.jpg) 0px 0px no-repeat; cursor:pointer}
	.regsubmit:hover{background:url(images/reg_btn.jpg) 0px -53px no-repeat;}
	.main_R{width:330px;float:left;margin-left:38px;display:inline;}
	.tabbox .companyul{margin-top:20px}
	.rzm{margin-left:30px;line-height:25px;color:#999999}
	.rzm span{color:#CC0000;}
    .family{margin-top:80px; line-height:25px; font-size:14px; font-family:"微软雅黑";}
	.family h3{height:40px;text-align:center;line-height:40px;font-size:30px;font-weight:bold;color:#FD8504;}
    .family h3 span{font-size:30px;color:#666;}
	.foot{width:800px;margin:0px auto;text-align:center;padding:8px 0 0 0px;line-height:24px;}
	.foot a{color:##474747;}
	.foot a:visited{ color:#666;}
</style>
</head>
<body id="wrap">

<%
dim username,password
If not KS.IsNul(KS.C("AdminName")) and not KS.IsNul(KS.C("AdminPass")) Then
     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "select top 1 * From KS_InterView Where ID=" & ID ,conn,1,1
	 If Not RS.Eof Then
	  username=rs("HostUserID")
	  password=rs("HostUserPass")
	 End If
	 rs.close
	 set rs=nothing
end if
%>
 
<table width="809" border="0" height="418" align="center"  style="margin:0 auto;background:url(images/regbg.png);">
 <FORM ACTION="Login.asp?Action=LoginCheck" method="post" name="LoginForm" id="LoginForm" onSubmit="return(CheckForm(this))">
 <input type="hidden" name="id" value="<%=id%>"/>
<tr>
 <td><div id="step_1" class="main">
				<div class="main_L">
 
					<div class="tabbox">
						<ul id="regSpan" class="companyul">
							
							<li style="z-index:1000">
								<div class="label">
									<label for="email" style="padding-left:28px">登录账号：</label><input type="text" name="UserName" id="UserName" class="textinput" tabindex="1" autocomplete="off" value="<%=username%>"/>
								</div>
							</li>
							<li>
								<div class="label">
									<label for="password" style="padding-left:28px">登录密码：</label><input value="<%=password%>" type="password" tabindex="2" name="PWD" id="PWD" class="textinput" />
 
								</div>
								
							</li>
						  
						<li style="clear:both">
								<div>
									<label for="email" style="padding-left:28px">选择角色：</label>
									&nbsp;<select name="role">
									 <option value="主持人">主持人</option>
									 <%
									 if not ks.isnul(guests) then
									   dim i,garr:garr=split(guests,",")
									   for i=0 to ubound(garr)
									    if not ks.isnul(garr(i)) then
									    response.write "<option value='" & garr(i) &"'>" & garr(i) &"</option>"
										end if
									   next
									 end if
									 %>
									</select>
									
								</div>
							</li>
							
							
							<li class="btn" id="nextStep">
							  <input type="submit" tabindex="5" name="submitbtn" id="submitbtn"  class="regsubmit" value=" ">
							</li>
						</ul>
					</div>
				</div>
				<div class="main_R">
					<div class="family">
							<h3>在线访谈</h3>
							 <br/>主持人及嘉宾登录平台，提供主持人及嘉宾互动交流！
					</div>
				
				</div>
			</div>
 </td>
</tr>
</FORM>
</table>
<script type="text/javascript"> 
<!--
$(document).ready(function() { 
    <%if username<>"" and password<>"" then%>
	 $("#LoginForm").submit();
	<%end if%>
	$(".label").hover(function(){$(this).removeClass("label");$(this).addClass("labelhover");
	},function(){
	$(this).removeClass("labelhover");$(this).addClass("label");});
});
 
setTimeout(function(){$("#UserName").focus();},500); 
 
function CheckForm(ObjForm) {
  if(ObjForm.UserName.value == '') {
    $.dialog.alert('请输入登录账号！',function(){ObjForm.UserName.focus();});
    return false;
  }
  if(ObjForm.PWD.value == '') {
    $.dialog.alert('请输入登录密码！',function(){ObjForm.PWD.focus();});
    return false;
  }
 return true;
  
}
//-->
</script>
 
<div class="foot">
	<p>厦门科汛软件有限公司 Copyright &copy;2006-<%=year(now)%> <a href="http://www.kesion.com" target="_blank"> www.kesion.com</a>,All Rights Reserved. </p>
</div>
<br/><br/><br/>
 
</body>
</html>
<%
end sub

Sub CheckLogin()
 Dim UserName:UserName = KS.R(trim(KS.S("username")))
 Dim PassWord:PassWord=KS.R(KS.S("Pwd"))
 if username="" or password="" then KS.AlertHintScript "登录用户及密码必须输入！"
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "select top 1 * From KS_InterView Where ID=" & ID & " and HostUserID='" & UserName &"' and HostUserPass='" & PassWord &"'",conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close:Set RS=Nothing
   KS.AlertHintScript "对不起，您输入的登录账号不正确！"
 End If
 RS.Close
 Set RS=Nothing
       
	   If EnabledSubDomain Then
			Response.Cookies(KS.SiteSn).domain=RootDomain					
	  Else
            Response.Cookies(KS.SiteSn).path = "/"
	  End If		
	  Response.Cookies(KS.SiteSn)("InterViewUserName") = UserName
	  Response.Cookies(KS.SiteSn)("InterViewPass") = PassWord
	  Response.Cookies(KS.SiteSn)("InterRole") = KS.S("Role")
	  Response.Cookies(KS.SiteSn)("InterViewID") = ID
 Response.Redirect "main.asp"
End Sub
Sub LoginOut()
		 
		   If EnabledSubDomain Then
				Response.Cookies(KS.SiteSn).domain=RootDomain					
			Else
                Response.Cookies(KS.SiteSn).path = "/"
			End If
			dim id:id=KS.C("InterViewID")
			 Response.Cookies(KS.SiteSn)("InterViewUserName") =""
			  Response.Cookies(KS.SiteSn)("InterViewPass") = ""
			  Response.Cookies(KS.SiteSn)("InterRole") = ""
			  Response.Cookies(KS.SiteSn)("InterViewID") = ""
			  Response.Redirect "show.asp?id=" &id
End Sub

set ks=nothing
closeconn
%>

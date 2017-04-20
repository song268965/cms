<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="Include/3GCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New GetPassCls
KSCls.Kesion()
Set KSCls = Nothing

Class GetPassCls
        Private KS,KSR,Action,F_C,FormStr,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Set KSUser=Nothing
		 CloseConn
		End Sub
		%>
		<!--#include file="include/function.asp"-->
		<%
      Public Sub Kesion()
	     Action=KS.S("Action")
		 F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/GetPassWord.html")    
		 FCls.RefreshType = "getpassword" '设置刷新类型，以便取得当前位置导航等
		 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		  InitialCommon
		 Select Case lcase(Action)
		  Case "next" CheckTimes:GetPASSNext
		  Case "next2" CheckTimes:GetPassNext2
		  Case "next3" CheckTimes:GetPassNext3
		  Case "next4" CheckTimes:GetPassNext4
		  Case "next5" CheckTimes:GetPassNext5
		  Case "verify" GetPassVerify
		  Case "doget" DoGetPass
		  Case Else
			 GetPassWordForm
		 End Select
		
		 F_C=Replace(F_C,"{$GetPassWordForm}",FormStr)
		 F_C=KSR.KSLabelReplaceAll(F_C)
		 KS.Die F_C
      End Sub
	  
	  sub CheckTimes()
	    If KS.ChkClng(KS.Setting(123))=0 Then Exit Sub
		'删除大于10天的无用记录
		Conn.Execute("Delete From KS_UserRecord Where flag=1 and datediff(" & DataPart_D & ",adddate," & sqlnowstring &")>10")
		
		if ks.chkclng(conn.execute("select count(1) from ks_userrecord where flag=1 and datediff(" & DataPart_D & ",adddate," & sqlnowstring &")=0 and userip='" & ks.getip &"'")(0))>=KS.ChkClng(KS.Setting(123)) then
				 Response.Write("<script>alert('对不起，系统限定每天只能使用" & KS.ChkClng(KS.Setting(123)) & "次找回密码功能!');history.back();</script>")
				 Response.End
			 end if
	 end sub 
	  
	  		
	   Sub GetPassWordForm()
		  FormStr="<form name=""getpassform"" action=""getpassword.asp"" method=""post""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""action"" value=""next""/>" &vbcrlf
	      FormStr=FormStr &"<h1>请输入您注册时填写的用户名</h1>"
		  FormStr=FormStr &"<table  cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"" class=""password"">" &vbcrlf
		  FormStr=FormStr &"<tr><td>找回方式：</td><td style=""text-align:left""><select name=""gettype"" id=""gettype"" onchange=""if(this.value==1){jQuery('#showemail').show();}else{jQuery('#showemail').hide();}""><option value=""1"">邮箱找回</option><option value=""2"" selected>安全问题找回</option>"
		  If Not KS.IsNul(Split(KS.Setting(155)&"∮∮","∮")(2)) And KS.Setting(157)="1" Then
			FormStr=FormStr & "<option value=""3"" selected>手机短信验证</option>"
		 End If
		  FormStr=FormStr &"</select></td></tr>"
		  FormStr=FormStr &"<tr><td width=""80"">用 户 名：</td><td style=""text-align:left""><input type=""text"" name=""UserName"" class=""input""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"<tbody style=""display:none"" id=""showemail""><tr><td class=""tdbg"">电子邮箱：</td><td style=""text-align:left""><input type=""text"" name=""Email"" class=""input""/></td><tr><td></td><td><span style='color:#999'>填写您在本站注册时使用的Email.</span></td></tr></tbody>" &vbcrlf
		  FormStr=FormStr &"<tr><td colspan=""2"" style=""height:40px;text-align:center""/> <input class=""button"" type=""submit"" value=""下一步""  "" onclick=""return(checkmyform())""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"</table>" &vbcrlf
		  FormStr=FormStr &"</form>" &vbcrlf
	   End Sub
	   
	   Sub GetPASSNext()
	     Dim UserName:UserName=KS.S("UserName")
		 Dim Email:Email=KS.S("Email")
		 Dim RS,GetType:GetType=KS.ChkClng(KS.S("GetType"))
		 If KS.IsNul(UserName) Then
		   KS.Die "<script>alert('请输入用户名!');history.back();</script>"
		 End If
		 If GetType=2 Then
		  GetPassByQuestion UserName
		  Exit Sub
		 ElseIf GetType=3 Then
		  GetPassBySms UserName
		  Exit Sub
		 ElseIf KS.IsNul(Email) Then
		   KS.Die "<script>alert('请输入邮箱地址!');history.back();</script>"
		 End If
		 If Not KS.IsValidEmail(Email) Then
		   KS.Die "<script>alert('您输入的邮箱地址不正确!');history.back();</script>"
		 End If
		  Call KSUser.AddUserRecord(1,"找回密码操作!") '记录操作
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select top 1 * From KS_User Where UserName='" & UserName & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   KS.Die "<script>alert('对不起，您输入的用户名不存在!');history.back();</script>"
		 End If
		 RS.CLose
		 RS.Open "select top 1 * From KS_User Where UserName='" & UserName & "' and email='" & Email & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   KS.Die "<script>alert('对不起，您输入的邮箱和您绑定的邮箱地址不正确!');history.back();</script>"
		 End If
		 Dim UserID,RealName
		 UserID=RS("UserId")
		 RealName=RS("RealName")
		 If KS.IsNul(RealName) Then RealName=UserName
		 RS.Close
		 Set RS=Nothing
		 Dim CheckCode:CheckCode=KS.MakeRandom(10)
		 Conn.Execute("Update KS_User Set RndPassWord='" & CheckCode & "' where username='"& UserName & "'")
		 Dim CheckUrl:CheckUrl=KS.GetDomain &"User/GetPassWord.asp?action=Verify&UserID=" & UserId &"&CheckNum=" & CheckCode
		 Dim MailBodyStr:MailBodyStr="您好" & RealName & "!<br/>这是由["&KS.Setting(0) & "]网站用于取回用户密码发送的邮件！<br/>----------------------------------------------------------------------<br/><strong>密码重置说明</strong><br/>----------------------------------------------------------------------<br/>请点击以下链接重置您的密码：<br/><a href=""" & checkurl & """ target=""_blank"">" & checkurl & "</a><br/><span style=""color:#999999"">(如果上面不是链接形式，请将该地址手工粘贴到浏览器地址栏再访问)</span><br/>在上面的链接所打开的页面中输入新的密码后提交，您即可使用新的密码登录网站了。您可以在用户控制面板中随时修改您的密码。<br/>本请求提交者的 IP 为 " & KS.GetIP & "<br/>此致<br/>" & KS.Setting(0) & "&nbsp;&nbsp;" & KS.GetDomain

         Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "[" & KS.Setting(0) & "]取回密码说明", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
        
		 IF ReturnInfo="OK" Then
          F_C=Replace(F_C,"{$GetPassWordForm}","恭喜，取回密码的方法已通过 Email 发送到您的信箱<span style=""color:red"">" & KS.CheckXSS(Email) & "</span>中,请注意查收！")
		 Else
          F_C=Replace(F_C,"{$GetPassWordForm}","对不起，邮件发送失败，原因：" &ReturnInfo )
		 End If
	   End Sub
	   
	   
	   '==========================================按手机短信取回_begin======================================================
	   Sub GetPassBySms(UserName)
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * From KS_User Where UserName='" & UserName &"'",conn,1,1
		  If RS.Eof And RS.Bof Then
		     RS.Close : Set RS=Nothing
		     KS.Die "<script>alert('对不起，您输入的用户名[" & UserName & "]不存在!');history.back();</script>"
		  End If
		  FormStr="<form name=""getpassform"" action=""getpassword.asp"" method=""post""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""action"" value=""next4""/>" &vbcrlf
	      FormStr=FormStr &"<div style=""text-align:left""><h1>请输入您注册时的手机号</h1>"
		  FormStr=FormStr &"<table  cellspacing=""0"" width=""100%"" cellpadding=""0"" border=""0"" class=""password"">" &vbcrlf
		  FormStr=FormStr &"<tr><td width=""100"" class=""tdbg"">用 户 名：</td><td style=""text-align:left""><input type=""hidden"" value=""" & UserName &""" name=""UserName"" class=""textbox""/>" & UserName &"</td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td class=""tdbg"">手机号码：</td><td style=""text-align:left""><input type=""text"" name=""Mobile"" id=""Mobile"" class=""input""/></td></tr>"
		  FormStr=FormStr &"<tr><td class=""tdbg"">短信验证码：</td><td style=""text-align:left""><input type=""text"" name=""MobileCode"" id=""MobileCode"" maxlength=""6"" class=""input""/></td></tr>"
		  FormStr=FormStr &"<tr><td class=""tdbg"" colspan=""2"" style=""text-align:center""><input type=""button"" id=""MobileCodeBtn"" onclick=""getMobileCode(" & KS.ChkClng(split(KS.Setting(156)&"∮","∮")(1)) &",'102','Mobile','MobileCodeBtn','" & UserName & "')"" value=""免费获取手机验证码"" class=""button""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td colspan=""2"" style=""height:40px;text-align:center""/> <input class=""button"" type=""submit"" value="" 确定提交 ""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"</table>" &vbcrlf
		  FormStr=FormStr &"</div>" & vbcrlf
		  FormStr=FormStr &"</form>" &vbcrlf
		  RS.Close : Set RS=Nothing
	   End Sub
	   
	    Sub CheckMobieCode()
		   Dim UserName:UserName=KS.DelSQL(KS.S("UserName"))
	       Dim Mobile:Mobile=KS.DelSQL(KS.S("Mobile"))
	       Dim MobileCode:MobileCode=KS.DelSQL(KS.S("MobileCode"))
	       If KS.IsNul(Mobile) Then
		    KS.Die "<script>alert('请输入手机号码!');history.back();</script>"
		   End If
	       If KS.IsNul(MobileCode) Then
		    KS.Die "<script>alert('请输入手机短信验证码!');history.back();</script>"
		   End If
	      Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * From KS_User Where UserName='" & UserName & "' and Mobile='" & Mobile & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   KS.Die "<script>alert('对不起，您输入的手机号码和您绑定的手机号码不一致!');history.back();</script>"
		 End If
		  Dim RSM:Set RSM=Conn.Execute("Select top 1 * From KS_UserRecord Where flag=102 And UserName='" & Mobile &"' Order By ID Desc")
			 If RSM.Eof And RSM.Bof Then
			   RSM.Close
			   Set RSM=Nothing
		   	   Response.Write("<script>alert('对不起，您输入的手机短信验证码不正确！');history.back(-1);</script>")
		       Exit Sub
			 End If
			 Dim RightMobileCode:RightMobileCode=RSM("Note")
			 Dim RightSendDate:RightSendDate=RSM("AddDate")
			 Dim RightMobile:RightMobile=RSM("UserName")
			 RSM.Close
			 Set RSM=Nothing
			 Dim TimeAllow:TimeAllow=KS.ChkClng(split(KS.Setting(156)&"∮∮","∮")(4))
             If RightMobile<>Mobile Then
			   Response.Write("<script>alert('对不起，您输入的手机号码与接收短消息的手机号码不一致！');history.back(-1);</script>")
		       Exit Sub
			 ElseIf MobileCode<>RightMobileCode Then
			   Response.Write("<script>alert('对不起，您输入的手机短信验证码不正确！');history.back(-1);</script>")
		       Exit Sub
			 ElseIf TimeAllow>0 and DateDiff("n",RightSendDate,Now)>TimeAllow Then
			   Response.Write("<script>alert('对不起，您输入的手机短信验证码已失效！');history.back(-1);</script>")
		       Exit Sub
			 End If
	   End Sub 
	   
	   Sub GetPassNext4()
	     Dim UserName:UserName=KS.S("UserName")
		 Dim Mobile:Mobile=KS.S("Mobile")
		 Dim MobileCode:MobileCode=KS.S("MobileCode")
		 If KS.IsNul(UserName) Then
		   KS.Die "<script>alert('请输入用户名!');history.back();</script>"
		 End If
		 
		 
		 Call KSUser.AddUserRecord(1,"找回密码操作!") '记录操作
		
		 CheckMobieCode
		 
		 
		 
	      FormStr="<div style=""text-align:left"">恭喜，您已通过手机短信验证，请重新设置您的密码"
		  FormStr=FormStr &"<form name=""getpassform"" action=""getpassword.asp"" method=""post""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""action"" value=""next5""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""MobileCode"" value=""" & KS.S("MobileCode") &"""/>" &vbcrlf
		  
		  FormStr=FormStr &"<input type=""hidden"" name=""Mobile"" value=""" & KS.S("Mobile") &"""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""username"" value=""" & username &"""/>" &vbcrlf
		  FormStr=FormStr &"<table  cellspacing=""0"" cellpadding=""0"" border=""0"">" &vbcrlf
		  FormStr=FormStr &"<tr><td width=""100"" class=""tdbg"" style=""text-align:right"">用 户 名：</td><td style=""text-align:left"">" & username &"</td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td class=""tdbg"" style=""text-align:right"">新 密 码：</td><td style=""text-align:left""><input type=""password"" name=""PassWord"" id=""PassWord"" class=""textbox""/></td></tr>"
		  FormStr=FormStr &"<tr><td class=""tdbg"" style=""text-align:right"">确定密码：</td><td style=""text-align:left""><input type=""password"" name=""RePassWord"" id=""RePassWord"" class=""textbox""/></td></tr>"
		  FormStr=FormStr &"<tr><td colspan=""2"" style=""height:40px;text-align:center""/> <input class=""button"" type=""submit"" value="" 确定提交 "" onclick=""return(checkgetform())""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"</table>" &vbcrlf
		  FormStr=FormStr &"</form>" &vbcrlf
		  FormStr=FormStr &"</div>" & vbcrlf
	   End Sub
	   Sub GetPassNext5()
	     Dim UserName:UserName=KS.S("UserName")
		 Dim PassWord:PassWord=KS.S("PassWord")
		 Dim RePassWord:RePassWord=KS.S("RePassWord")
		 
		  CheckMobieCode()
		 
		 If KS.IsNul(PassWord) Or KS.IsNul(RePassWord) Then
		   KS.Die "<script>alert('请输入您的新密码!');history.back();</script>"
		 End If
		 If PassWord<>RePassWord Then
		   KS.Die "<script>alert('两次输入的密码不一致!');history.back();</script>"
		 End If
		 If KS.IsNul(UserName) Then
		   KS.Die "<script>alert('请输入用户名!');history.back();</script>"
		 End If

		 Conn.Execute("Update KS_User Set [PassWord]='" & MD5(PassWord,16) & "' where UserName='" & UserName &"'")
		 KS.Die "<script>alert('恭喜，您的新密码已生效，现在可以登录了!');location.href='login.asp';</script>"
	   End Sub
	   '==========================================按手机短信取回_end======================================================
	   
	   
	   
	   '==========================================按注册时的安全问题取回_begin==============================================
	   Sub GetPassByQuestion(UserName)
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * From KS_User Where UserName='" & UserName &"'",conn,1,1
		  If RS.Eof And RS.Bof Then
		     RS.Close : Set RS=Nothing
		     KS.Die "<script>alert('对不起，您输入的用户名[" & UserName & "]不存在!');history.back();</script>"
		  End If
	      FormStr="<div style=""text-align:left""><h1>请回答您设置的密码答案</h1>"
		  FormStr=FormStr &"<form name=""getpassform"" action=""getpassword.asp"" method=""post""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""action"" value=""next2""/>" &vbcrlf
		  FormStr=FormStr &"<table  cellspacing=""0"" width=""100%"" cellpadding=""0"" border=""0"" class=""password"">" &vbcrlf
		  If KS.IsNul(RS("Question")) And KS.IsNul(RS("Answer")) Then
		  FormStr=FormStr &"<tr>用 户 名：<input type=""hidden"" value=""" & UserName &""" name=""UserName"" class=""textbox""/>" & UserName &"</td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td>对不起，您未设置密码安全问题和答案，无法通过安全问题方式找回密码，请选择其它方式！</td></tr>"
		  FormStr=FormStr &"<tr><td style=""height:40px;text-align:center""/> <input class=""button"" type=""button"" value="" 返回 "" onclick=""history.back()""/></td></tr>" &vbcrlf
		  Else
		  FormStr=FormStr &"<tr><td width=""100"" class=""tdbg"" width=""80"">用 户 名：</td><td style=""text-align:left""><input type=""hidden"" value=""" & UserName &""" name=""UserName"" class=""textbox""/>" & UserName &"</td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td class=""tdbg"">您的问题：</td><td style=""text-align:left"">" & RS("Question") & "</td></tr>"
		  FormStr=FormStr &"<tr id=""showemail""><td class=""tdbg"">您的答案：</td><td style=""text-align:left""><input type=""text"" name=""Answer"" id=""Answer"" class=""input""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td colspan=""2"" style=""height:40px;text-align:center""/> <input class=""button"" type=""submit"" value="" 确定提交 ""/></td></tr>" &vbcrlf
		  End If
		  FormStr=FormStr &"</table>" &vbcrlf
		  FormStr=FormStr &"</form>" &vbcrlf
		  FormStr=FormStr &"</div>" & vbcrlf
		  RS.Close : Set RS=Nothing
	   End Sub
	   Sub GetPassNext2()
	     Dim UserName:UserName=KS.S("UserName")
		 Dim Answer:Answer=KS.S("Answer")
		 If KS.IsNul(UserName) Then
		   KS.Die "<script>alert('请输入用户名!');history.back();</script>"
		 End If
		 If KS.IsNul(Answer) Then
		   KS.Die "<script>alert('请输入您设置的取回密码问题答案!');history.back();</script>"
		 End If
		 Call KSUser.AddUserRecord(1,"找回密码操作!") '记录操作
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select top 1 * From KS_User Where UserName='" & UserName & "' and Answer='" & Answer &"'",conn,1,1
		 If RS.Eof And  RS.Bof Then
		    RS.Close: Set RS=Nothing
			KS.Die "<script>alert('对不起，您输入的密码答案不正确!');history.back();</script>"
		 End If
		 RS.Close:Set RS=Nothing
	      FormStr="<div style=""text-align:left""><h1>恭喜，您的密码取回答案回答正确，请设置新密码</h1>"
		  FormStr=FormStr &"<form name=""getpassform"" action=""getpassword.asp"" method=""post""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""action"" value=""next3""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""answer"" value=""" & Answer &"""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""username"" value=""" & username &"""/>" &vbcrlf
		  FormStr=FormStr &"<table  cellspacing=""0"" cellpadding=""0"" border=""0"">" &vbcrlf
		  FormStr=FormStr &"<tr><td width=""100"" class=""tdbg"">用 户 名：</td><td style=""text-align:left"">" & username &"</td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td class=""tdbg"">新 密 码：</td><td style=""text-align:left""><input type=""password"" name=""PassWord"" id=""PassWord"" class=""textbox""/></td></tr>"
		  FormStr=FormStr &"<tr><td class=""tdbg"">确定密码：</td><td style=""text-align:left""><input type=""password"" name=""RePassWord"" id=""RePassWord"" class=""input""/></td></tr>"
		  FormStr=FormStr &"<tr><td colspan=""2"" style=""height:40px;text-align:center""/> <input class=""button"" type=""submit"" value="" 确定提交 "" onclick=""return(checkgetform())""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"</table>" &vbcrlf
		  FormStr=FormStr &"</form>" &vbcrlf
		  FormStr=FormStr &"</div>" & vbcrlf
	   End Sub
	  '==========================================按注册时的安全问题取回_end==============================================

	   
	   
	   Sub GetPassNext3()
	     Dim UserName:UserName=KS.S("UserName")
		 Dim Answer:Answer=KS.S("Answer")
		 Dim PassWord:PassWord=KS.S("PassWord")
		 Dim RePassWord:RePassWord=KS.S("RePassWord")
		 If KS.IsNul(PassWord) Or KS.IsNul(RePassWord) Then
		   KS.Die "<script>alert('请输入您的新密码!');history.back();</script>"
		 End If
		 If PassWord<>RePassWord Then
		   KS.Die "<script>alert('两次输入的密码不一致!');history.back();</script>"
		 End If
		 If KS.IsNul(UserName) Then
		   KS.Die "<script>alert('请输入用户名!');history.back();</script>"
		 End If
		 If KS.IsNul(Answer) Then
		   KS.Die "<script>alert('请输入您设置的取回密码问题答案!');history.back();</script>"
		 End If
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select top 1 * From KS_User Where UserName='" & UserName & "' and Answer='" & Answer &"'",conn,1,1
		 If RS.Eof And  RS.Bof Then
		    RS.Close: Set RS=Nothing
			KS.Die "<script>alert('对不起，您输入的密码答案不正确!');history.back();</script>"
		 End If
		 RS.Close:Set RS=Nothing
		 Conn.Execute("Update KS_User Set [PassWord]='" & MD5(PassWord,16) & "' where UserName='" & UserName &"'")
		 KS.Die "<script>alert('恭喜，您的新密码已生效，现在可以登录了!');location.href='login';</script>"
	   End Sub
	   '=====================================================================================================
	   
	   Sub GetPassVerify()
	     Dim UserID:UserID=KS.ChkClng(KS.S("UserID"))
		 Dim CheckNum:CheckNum=KS.S("CheckNum")
		 If UserID=0 Or CheckNum="" Then KS.Die "error"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select top 1 * From KS_User Where UserID=" & UserID & " and RndPassWord='" & CheckNum & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   FormStr="对不起，您访问的链接地址已失败或是您非法访问！"
		 Else
	      FormStr="<div style=""text-align:left""><h1>请重置您的登录密码</h1>"
		  FormStr=FormStr &"<form name=""getpassform"" action=""getpassword.asp"" method=""post""/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""UserID"" value=""" & UserID & """/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""CheckNum"" value=""" & CheckNum & """/>" &vbcrlf
		  FormStr=FormStr &"<input type=""hidden"" name=""action"" value=""doget""/>" &vbcrlf
		  FormStr=FormStr &"<table  cellspacing=""0"" cellpadding=""0"" border=""0"">" &vbcrlf
		  FormStr=FormStr &"<tr><td width=""100"" class=""tdbg"" style=""text-align:right"">用 户 名：</td><td style=""text-align:left"">" & RS("UserName") & "</td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td class=""tdbg"" style=""text-align:right"">新 密 码：</td><td style=""text-align:left""><input type=""PassWord"" name=""PassWord"" class=""textbox""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td class=""tdbg"" style=""text-align:right"">重复密码：</td><td style=""text-align:left""><input type=""PassWord"" name=""RePassWord"" class=""textbox""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"<tr><td colspan=""2"" style=""height=""40"" text-align:center""/> <input class=""button"" type=""submit"" value=""确定重置"" onclick=""return(checkgetform())""/></td></tr>" &vbcrlf
		  FormStr=FormStr &"</table>" &vbcrlf
		  FormStr=FormStr &"</form>" &vbcrlf
		  FormStr=FormStr &"</div>" & vbcrlf
		 End If
		 RS.Close
		 Set RS=Nothing
	   End Sub
	   
	   Sub DoGetPass()
	   	 Dim UserID:UserID=KS.ChkClng(KS.S("UserID"))
		 Dim CheckNum:CheckNum=KS.S("CheckNum")
		 Dim PassWord:PassWord=KS.S("PassWord")
		 Dim RePassWord:RePassWord=KS.S("RePassWord")
		 If UserID=0 Or CheckNum="" Then KS.Die "error"
		 If KS.IsNul(PassWord) Or KS.IsNul(RePassWord) Then
		   KS.Die "<script>alert('请输入您的新密码!');history.back();</script>"
		 End If
		 If PassWord<>RePassWord Then
		   KS.Die "<script>alert('两次输入的密码不一致!');history.back();</script>"
		 End If
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select top 1 * From KS_User Where UserID=" & UserID & " and RndPassWord='" & CheckNum & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.CLose :Set RS=NOthing
		   KS.Die "<script>alert('出错了。请不要非法访问!');window.close();</script>"
		 End If
		 RS.Close:Set RS=Nothing
		 Conn.Execute("Update KS_User Set [PassWord]='" & MD5(PassWord,16) & "' where userid=" & userid)
		 KS.Die "<script>alert('恭喜，您的新密码已生效，现在可以登录了!');location.href='login.asp';</script>"
	   End Sub
	   	
       
End Class
%> 

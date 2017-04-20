<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../api/cls_api.asp"-->
<!--#include file="../api/uc_client/client.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New User_RegPost
KSCls.Kesion()
Set KSCls = Nothing

Class User_RegPost
        Private KS,KSRFObj
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		 IF KS.Setting(21)=0 Then : Response.Redirect "index.asp"  :  Response.End

		 '系统配置参数
		 Dim Locked:Locked=0
		 Dim VerificCodeTF:VerificCodeTF=KS.ChkClng(KS.Setting(27))
		 Dim EmailMultiRegTF:EmailMultiRegTF=KS.Setting(28)
		 Dim UserNameLimitChar:UserNameLimitChar=KS.ChkClng(KS.Setting(29))
		 Dim UserNameMaxChar:UserNameMaxChar=KS.ChkClng(KS.Setting(30))
		 Dim EnabledUserName:EnabledUserName=KS.Setting(31)
		 Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38) : If Not IsNumerIc(NewRegUserMoney) Then NewRegUserMoney=0
		 Dim NewRegUserScore:NewRegUserScore=KS.Setting(39) : If Not IsNumeric(NewRegUserScore) Then NewRegUserScore=0
		 Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40) : If Not IsNumeric(NewRegUserPoint) Then NewRegUserPoint=0
		 
         
		  If Request.ServerVariables("HTTP_REFERER")="" Then Call KS.Alert("请不要非法提交!","../"):Response.End
		  If Instr(Lcase(Request.ServerVariables("HTTP_REFERER")),"3g/reg.asp")=0 Then Call KS.Alert("请不要非法提交!","../") : Response.End
		 
		 '收集用户资料
		 Dim Verifycode:Verifycode=KS.S("Verifycode")
		  IF lcase(Trim(Verifycode))<>lcase(Trim(Session("Verifycode"))) And VerificCodeTF=1 then 
		   	 Response.Write("<script>alert('验证码有误，请重新输入！');history.back(-1);</script>")
		     Exit Sub
		  End IF
		  
		 Dim Mobile:Mobile=KS.S("Mobile")
		 If Instr(mobile,",")<>0 then mobile=replace(mobile,",","")
 '需要手机验证码
		  If Not KS.IsNul(Split(KS.Setting(155)&"∮","∮")(0)) and KS.Setting(157)="1"  Then
		     Dim MobileCode:MobileCode=KS.S("MobileCode")
			 If KS.IsNul(Mobile) Then
		   	   Response.Write("<script>alert('请输入手机号码！');history.back(-1);</script>")
		       Exit Sub
			 End If
			 If KS.IsNul(MobileCode) Then 
		   	   Response.Write("<script>alert('请输入手机短信验证码！');history.back(-1);</script>")
		       Exit Sub
			 End If
			 Dim RSM:Set RSM=Conn.Execute("Select top 1 * From KS_UserRecord Where flag=101 And UserName='" & Mobile &"' Order By ID Desc")
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
		  End If
		  
		  
		  '检查注册回答问题
		  Dim CanReg,n
		   If Mid(KS.Setting(161),1,1)="1" Then
		        CanReg=false
				 For N=0 To Ubound(Split(KS.GetCurrQuestion(162),vbcrlf))
				   If Trim(Request.Form("a" & MD5(n,16)))<>"" Then
					  If trim(Lcase(Request.Form("a" & MD5(n,16))))<>trim(Lcase(Split(KS.GetCurrQuestion(163),vbcrlf)(n))) Then
					   Call KS.AlertHistory("对不起,注册问题的回答不正确!",-1) : Response.End
					   CanReg=false
					  Else
					   CanReg=True
					  End If
				   End If
				 Next
			 If CanReg=false Then Call KS.AlertHistory("对不起,注册答案不能为空!",-1) : Response.End
		   End If
		 
		  
		 Dim UserName:UserName=KS.R(KS.S("UserName"))
		 If UserName = "" Or KS.strLength(UserName) > UserNameMaxChar Or KS.strLength(UserName) < UserNameLimitChar Then
		   	 Response.Write("<script>alert('请输入用户名(不能大于" & UserNameMaxChar & "小于" & UserNameLimitChar & ")');history.back();</script>")
			 Exit Sub
		 Elseif isnumeric(UserName) then
			 Response.Write("<script>alert('对不起，会员名不能是纯数字！');history.back();</script>")
			 Exit Sub
		 Elseif KS.HasChinese(username) and KS.ChkClng(KS.Setting(175))="0" then
		   	 Response.Write("<script>alert('对不起，系统设置用户名不能含有中文！');history.back();</script>")
			 Exit Sub
         ElseIF KS.FoundInArr(EnabledUserName, UserName, "|") = True Then
		   	 Response.Write("<script>alert('您输入的用户名为系统禁止注册的用户名');history.back();</script>")
			 Exit Sub
		 ElseIF InStr(UserName, "-") > 0 Or InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
             Response.Write("<script>alert('用户名中含有非法字符');history.back();</script>")
			 Exit Sub
        End If
		 
		 Dim PassWord,RePassWord,NoMD5_Pass
		 If Session("PassWord")<>"" Then
		   PassWord=Session("PassWord")
		 Else
			 PassWord=KS.R(KS.S("PassWord"))
			 If PassWord = "" Then
				 Response.Write("<script>alert('请输入登录密码!');history.back();</script>")
				 Exit Sub
			 End If
		 End If
		 NoMD5_Pass=PassWord
		 Dim RndPassword:RndPassword=NoMD5_Pass
		 Dim Question:Question=KS.S("Question")
		 Dim Answer:Answer=KS.S("Answer")
		 If KS.Setting(148)="1" Then
			 If Question = "" Then
				' Response.Write("<script>alert('密码提示问题不能为空!');history.back();<//script>")
				' Exit Sub
			 ElseIF Answer="" Then
				 'Response.Write("<script>alert('密码答案不能为空');history.back();<//script>")
				' Exit Sub
			 End If
		 End If
		 
		 Dim Email:Email=KS.S("Email")
		 if KS.IsValidEmail(Email)=false and Email<>"" then
			 Response.Write("<script>alert('请输入正确的电子邮箱!');history.back();</script>")
			 Exit Sub
		 end if
		 
		 Dim RealName:RealName=KS.S("RealName")
		 Dim Sex:Sex=KS.S("Sex") : If Sex="" Then Sex="男"
		 Dim Birthday:Birthday=KS.S("Birthday")
		 If Not IsDate(Birthday) Then Birthday=FormatDateTime(Now,2)
		 Dim IDCard:IDCard=KS.S("IDCard")
		 Dim OfficeTel:OfficeTel=KS.S("OfficeTel")
		 Dim HomeTel:HomeTel=KS.S("HomeTel")
		 Dim Fax:Fax=KS.S("Fax")
		 Dim province:province=KS.S("province")
		 Dim city:city=KS.S("city")
		 Dim Address:Address=KS.S("Address")
		 Dim ZIP:ZIP=KS.S("ZIP")
		 Dim HomePage:HomePage=KS.S("HomePage")
		 Dim UserFace:UserFace=KS.S("UserFace")
		 if userface="" then 
		   if sex="男" then userface=KS.GetDomain & "Images/Face/boy.jpg" else userface=KS.GetDomain & "Images/face/girl.jpg"	 	 
		 End If
		 If KS.Setting(129)="1" and (KS.Setting(149)="1" or Not KS.IsNul(Split(KS.Setting(155)&"∮","∮")(0))) Then
		   If Not Conn.Execute("Select top 1 userid from KS_User Where Mobile='" & Mobile & "'").eof Then
			 Response.Write("<script>alert('对不起，您输入的手机号码已被占用!');history.back();</script>")
			 Exit Sub
		   End If
		 End If
		 Dim QQ:QQ=KS.S("QQ")		 
		 Dim ICQ:ICQ=KS.S("ICQ")		 
		 Dim MSN:MSN=KS.S("MSN")		 
		 Dim UC:UC=KS.S("UC")		 
		 Dim Sign:Sign=KS.S("Sign")	
		 Dim Privacy:Privacy=KS.ChkClng(KS.S("Privacy"))
		 Dim AllianceUser:AllianceUser=KS.S("AllianceUser")
		 
		 Dim LastLoginIP:LastLoginIP = KS.GetIP()
		 Dim CheckNum:CheckNum = KS.MakeRandomChar(6)  '随机字符验证码
		 Dim CheckUrl:CheckUrl = Request.ServerVariables("HTTP_REFERER")
		 CheckUrl=KS.GetDomain &"User/ActiveCode.asp?Action=Active&UserId={$UserID}&CheckNum=" & CheckNum
		 	 
		 PassWord =MD5(KS.R(PassWord),16)
		 Dim RS,SQL,K
		 Dim GroupID:GroupID=KS.ChkClng(KS.S("GroupID")):If GroupID=0 Then GroupID=2
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where ID=" & GroupID&")")
		 If FieldsList="" Then FieldsList="0"
		 FieldsList=KS.FilterIDs(FieldsList)
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 If FieldsList<>"" and true=false Then
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=101 and (FieldID In(" & FieldsList & ") or (ParentFieldName<>'0' and ParentFieldName is not null))",conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		 End If
		 If KS.Setting(32)="2" And IsArray(SQL) Then
			 For K=0 To UBound(SQL,2)
			  If SQL(6,K)="0" Then
					   If SQL(1,K)="1" Then
						 If SQL(0,K)="Province&City" Then
						  If KS.S("Province")="" and  KS.S("City")="" Then
							 Response.Write "<script>alert('" & SQL(2,K) & "必须填写!');history.back();</script>"
							 Response.End()
						  End If
						 ElseIf KS.IsNul(KS.S(SQL(0,K))) Then
							 Response.Write "<script>alert('" & SQL(2,K) & "必须填写!');history.back();</script>"
							 Response.End()
						 End If
					   End If
					   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then 
						 Response.Write "<script>alert('" & SQL(2,K) & "必须填写数字!');history.back();</script>"
						 Response.End()
					   End If
					   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
						 Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的日期!');history.back();</script>"
						 Response.End()
					   End If
					   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
						Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的Email格式!');history.back();</script>"
						Response.End()
					   End If
			 End If 
		   Next
		End If
		RS.Open "Select top 1 ID From KS_UserGroup Where ID=" & GroupID,conn,1,1
		If RS.Eof And RS.Bof Then
		     Rs.Close:Set RS=Nothing
			 Response.Write "<script>alert('对不起,用户组类型不正确!');history.back();</script>"
			 Response.End()
		End If
		RS.Close
		RS.Open "select top 1 * from KS_User where UserName='" & UserName & "'", Conn, 1, 3
		If Not (RS.BOF And RS.EOF) Then
				 RS.Close:Set RS=Nothing
				 Response.Write("<script>alert('您注册的用户已经存在！请换一个用户名再试试！');history.back();</script>")
				 Exit Sub
		Else
			If EmailMultiRegTF=0 Then
				Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select top 1 UserID from KS_User where Email='" & Email & "'")
				If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
					Response.Write("<script>alert('您注册的Email已经存在！请更换Email再试试！');history.back();</script>")
					Exit Sub
				End If
				EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			 If KS.ChkClng(KS.Setting(26))=1 Then
			   If Not (Conn.Execute("select top 1 UserID From KS_User Where LastLoginIP='" & KS.GetIP & "'").eof) Then
					Response.Write("<script>alert('您的IP已经存在！不能再注册！');history.back();</script>")
					Exit Sub
			   End If
			 End If

		
		'-----------------------------------------------------------------
		'Ucenter系统整合
		'-----------------------------------------------------------------
		dim uid
		If API_Enable Then '检查ucenter里是否存在的用户,不存在则提交注册
			uid = uc_user_register(username,NoMD5_Pass,email)
			uid =CInt(uid)
			if(uid <= 0) Then 
				if(uid = -1) Then 
				    Call KS.Die("<script>alert('用户名不合法！');history.back();</script>") 
				 elseif(uid = -2) Then
				    Call KS.Die("<script>alert('包含要允许注册的词语！');history.back();</script>") 
				 elseif(uid = -3) Then
				    Call KS.Die("<script>alert('用户名已经存在！');history.back();</script>") 
				 elseif(uid = -4) Then
				    Call KS.Die("<script>alert('Email 格式有误！');history.back();</script>") 
				 elseif(uid = -5) Then
				    Call KS.Die("<script>alert('Email 不允许注册！');history.back();</script>") 
				 elseif(uid = -6) Then
				    Call KS.Die("<script>alert('该 Email 已经被注册！');history.back();</script>") 
				 else 
				   Call KS.Die("<script>alert('未定义错误！');history.back();</script>") 
				End if
			End If
		End If
		'-----------------------------------------------------------------
		 
		 
		 RS.AddNew
		 RS("GroupID")=GroupID
		 RS("UserName")=UserName
		 RS("PassWord")=PassWord
		 RS("Question")=Question
		 RS("Answer")=Answer
		 RS("Email")=Email
		 RS("RealName")=RealName
		 RS("Sex")=Sex
		 RS("Birthday")=Birthday
		 RS("IDCard")=IDCard
		 RS("OfficeTel")=OfficeTel
		 RS("HomeTel")=HomeTel
		 RS("Mobile")=Mobile
		 If Not KS.IsNul(Split(KS.Setting(155)&"∮","∮")(0)) Then
		  RS("IsMobileRZ")=1
		 End If
		 RS("Fax")=Fax
		 RS("Province")=Province
		 RS("City")=City
		 RS("Address")=Address
		 RS("Zip")=Zip
		 RS("HomePage")=HomePage
		 RS("QQ")=QQ
		 RS("ICQ")=ICQ
		 RS("MSN")=MSN
		 RS("UC")=UC
		 RS("UserFace")=UserFace
		 RS("Sign")=Sign
		 RS("Privacy")=Privacy
		 RS("RegDate")=Now
		 RS("BeginDate")=Now '开始计算时间
		 RS("LastLoginIP")=LastLoginIP
		 RS("JoinDate")=Now
		 RS("LastLoginTime")=Now
		 RS("CheckNum")=CheckNum
		 RS("RndPassword")=RndPassword
		 RS("LoginTimes")=1
		 RS("PostNum")=0
		 
		 '自定义字段
		 If IsArray(SQL) Then
			 Dim UpFiles
			 For K=0 To UBound(SQL,2)
			  If left(Lcase(SQL(0,K)),3)="ks_" Then
			   RS(SQL(0,K))=KS.S(SQL(0,K))
			   If SQL(3,K)="9" or SQL(3,K)="10" Then
			   UpFiles=KS.S(SQL(0,K))
			   End If
			  End If
			  If SQL(4,K)="1" Then
			   RS(SQL(0,K)&"_Unit")=KS.S(SQL(0,K)&"_Unit")
			  End If
			 Next
		 End If
		 
		 RS("AllianceUser")=AllianceUser

		 '新会员注册，更新相应的数据
		 RS("ChargeType")=1
		 RS("Money")=0
		 RS("Score")=0
		 
		 If KS.ChkClng(KS.U_G(GroupID,"chargetype"))=1 Then
		  NewRegUserPoint=KS.ChkClng(KS.U_G(GroupID,"grouppoint"))
		 End If
		 If KS.ChkClng(KS.U_G(GroupID,"showonreg"))=0 Then  KS.Die "<script>alert('对不起，该用户组不允许注册!');</script>"
		 
		 RS("Point")=0
		 RS("Locked")=Locked
		 RS.Update
		 
		 RS.MoveLast
		 Dim UserID:UserID=RS("UserID")
		 RS.Close
		 If Not KS.IsNul(UpFiles) Then
		  Call KS.FileAssociation(1023,UserID,UpFiles,0)
		 End If
		 CheckUrl=replace(CheckUrl,"{$UserID}",UserID)
		 
		'更新论坛级别
		If KS.ChkClng(KS.Setting(56))=1 Then
			Dim RSG:Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where Special=0 and typeflag=1 and ClubPostNum<=0 And score<=" & NewRegUserScore & " order by score desc,ClubPostNum Desc")
			If Not RSG.Eof Then
				 Conn.Execute("Update KS_User Set ClubGradeID=" & RSG(0) & " WHERE UserName='" & UserName & "'")
			End If
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
            doc.documentElement.attributes.getNamedItem("totalusernum").text=Conn.Execute("Select count(1) From KS_User")(0)
			doc.documentElement.attributes.getNamedItem("newreguser").text=UserName
			doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		End If
		Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where Special=0 and typeflag=0 And score<=" & NewRegUserScore & " order by score desc,gradeid Desc")
		If Not RSG.Eof Then
			 Conn.Execute("Update KS_User Set GradeID=" & RSG(0) & ",GradeTitle='" & RSG(1) & "' WHERE UserName='" & UserName & "'")
		End If
		RSG.Close:Set RSG=Nothing


		 If NewRegUserPoint<>0 Then
		   Call KS.PointInOrOut(0,0,UserName,1,NewRegUserPoint,"系统","注册新会员,赠送!",0)
		 End If
		 IF NewRegUserScore<>0 Then
		   Call KS.ScoreInOrOut(UserName,1,NewRegUserScore,"系统","注册新会员,赠送!",0,0)
		 End If
		 If NewRegUserMoney<>0 Then 
		  Call KS.MoneyInOrOut(UserName,UserName,NewRegUserMoney,4,1,now,0,"System","注册新会员,赠送!",0,0,0)
		 End If
		 
		 RS.Open "Select top 1 * From KS_UserGroup Where ID=" & GroupID,conn,1,1
		 If RS.Eof Then RS.Close : Set RS=Nothing :Response.Write "<script>location.href='../../';</script>"
		 
		 Dim EmailCheckTF:EmailCheckTF=RS("ValidType")
		 Dim UserType:UserType= RS("UserType")
		 Conn.Execute("Update KS_User Set ChargeType=" & RS("ChargeType") & ",EDays=" & RS("ValidDays") & ",UserType=" & UserType &" Where UserID=" & UserID)
		 
		 Dim MailBodyStr
		 If EmailCheckTF = 1 Then
			MailBodyStr = Replace(RS("ValidEmail"), "{$CheckNum}", CheckNum)
			MailBodyStr = Replace(MailBodyStr, "{$CheckUrl}", CheckUrl)
			MailBodyStr = Replace(MailBodyStr, "{$UserName}", UserName)
			MailBodyStr = Replace(MailBodyStr, "{$PassWord}", NoMD5_Pass)
	       Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "新用户注册激活信", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
			  IF ReturnInfo="OK" Then
			     Conn.Execute("Update KS_User Set Locked=3 where userid=" & UserID)        '设置待激活
				 ReturnInfo="<li class=""okbox""><span id=""blue"">恭喜您注册成功!</span> 欢迎您入住{$GetSiteName}。<br />注册验证码已发送到您的信箱<font color='#ff6600>'>" &Email &"</font>，激活后正式成为本站会员!</li>"
			  Else
				ReturnInfo="<li>信件发送失败!失败原因:" & ReturnInfo & "，请联系网站管理员!</li>"
			  End if
		ElseIF EmailCheckTF=2 Then
		    Conn.Execute("Update KS_User Set Locked=2 where userid=" & UserID)        '设置需要后台认证
		    ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & UserName & "</font>,您需要通过管理员的认证才能成为正式会员!</li>"
		Else
		    ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & UserName & "</font>,您已成为了本站的正式会员!<br><div align=center></li>"
			
			'====================推荐计划======================================
			If AllianceUser<>"" and AllianceUser<>UserName  Then
			 If Not Conn.Execute("Select Top 1 UserID From KS_User Where UserName='" & AllianceUser & "'").eof Then
			 
			   '判断有没有恶意推荐注册,恶意注册的不给积分
			   If Conn.Execute("Select top 1 * From KS_PromotedPlan Where UserIP='" & KS.GetIP & "' And DateDiff(" & DataPart_D & ",AddDate," & SqlNowString & ")<1 And UserName='" & AllianceUser & "'").eof Then
			   
			   If KS.ChkClng(KS.Setting(144))>0 Then
			          If KS.Setting(145)="0" Then
					    Call KS.ScoreInOrOut(AllianceUser,1,KS.ChkClng(KS.Setting(144)),"系统","成功推荐一个注册用户:" & UserName & "!",0,0)
					  ElseIf KS.Setting(145)="1" Then
					    Call KS.PointInOrOut(0,0,AllianceUser,1,KS.Setting(144),"系统","成功推荐一个注册用户:" & UserName & "!",0)
					  Else
					    Call KS.MoneyInOrOut(AllianceUser,AllianceUser,KS.Setting(144),4,1,now,0,"System","成功推荐一个注册用户:" & UserName & "!",0,0,0)
					  End If
			    
			   End If
			   
			   
			   Conn.Execute("Insert InTo KS_PromotedPlan(UserName,UserIP,AddDate,ComeUrl,Score,AllianceUser) values('" & AllianceUser & "','" & KS.GetIP & "'," & SqlNowString & ",'" & KS.URLDecode(Request.ServerVariables("HTTP_REFERER")) & "'," & KS.ChkClng(KS.Setting(144)) & ",'" & UserName & "')")
			   End If
			   
			  '=================判断是不是好友邮件推荐的==================
				Dim f:f=KS.S("F")
				if f="r" Then
				 Conn.Execute("insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&AllianceUser&"','"& UserName &"',"&SqlNowString&",1,'',1)")
				End If
			  '============================================================

			 End If
			End If
			'====================推广计划结束=================================

        End IF
		
		RS.Close
			
			'==================注册成功发送手机短信_begin===============================
			Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮","∮")(1)
			If Not KS.IsNul(SmsContent) And Not KS.IsNul(Mobile) Then
			   SmsContent=Replace(SmsContent,"{$username}",UserName)
			   SmsContent=Replace(SmsContent,"{$possword}",NoMD5_Pass)
			   SmsContent=Replace(SmsContent,"{$email}",Email)
			   Call KS.SendMobileMsg(Mobile,SmsContent)
			End If
			'==================注册成功发送手机短信_end===============================
			
			
			'==================注册成功发邮件通知======================
			 If KS.Setting(146)="1" and Not KS.IsNul(KS.Setting(147)) And EmailCheckTF<>1 Then
				MailBodyStr = Replace(KS.Setting(147), "{$UserName}", UserName)
				MailBodyStr = Replace(MailBodyStr, "{$PassWord}", NoMD5_Pass)
				MailBodyStr = Replace(MailBodyStr, "{$SiteName}", KS.Setting(0))
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-会员注册成功", Email,UserName, MailBodyStr,KS.Setting(11))
				IF ReturnInfo="OK" Then
				  ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & UserName & "</font>,已将用户名和密码发到您的信箱!</li>"
				End If
			 End If
			'==========================================================
			
			
			'===============================判断商城系统有没有启用自动发放优惠券==============================
			If KS.ChkCLng(KS.C_S(5,21))=1 Then
			  RS.Open "Select * From KS_ShopCoupon Where CouponType=3 and status=1 order by  ID",conn,1,1
			   Do While Not RS.Eof
			      Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
				  RSC.Open "Select top 1 * From KS_ShopCouponUser Where 1=0",conn,1,3
				  RSC.AddNew
				  RSC("CouponID")=RS("ID")
				  RSC("CouponNum")=GetCouponNum()
				  RSC("UserName")=UserName
				  RSC("OrderID")=""
				  RSC("UseFlag")=0
				  RSC("AddDate")=Now
				  RSC("AvailableMoney")=RS("FaceValue")
				  RSC.Update
				  RSC.Close:Set RSC=Nothing
			  RS.MoveNext
			  Loop
			  RS.Close
			End If
			'====================================================================================================
			
			
		    '===================写入个人空间================
			if KS.SSetting(0)=1 And KS.SSetting(1)=1 then
			 Dim SpaceStatus:SpaceStatus=0
			 If UserType=1 then
			  if KS.ChkClng(KS.SSetting(2))=0 then
			   SpaceStatus=1
			  Else
			   SpaceStatus=0
			  End If
			 Else
			   if KS.ChkClng(KS.SSetting(2))=2 then
			    SpaceStatus=0
			   Else
			    SpaceStatus=1
			   End If
			 End If
			 RS.Open "Select top 1 * From KS_Blog Where 1=0",conn,1,3
			 RS.AddNew
			  RS("AddDate")=Now
			  RS("UserID")=UserID
			  RS("UserName")=UserName
			  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
			  If UserType=1 Then
			  RS("BlogName")=UserName & "的企业空间"
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=4 and IsDefault='true'")(0))
			  Else
			  RS("BlogName")=UserName & "的个人空间"
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  End If
			  RS("Announce")="暂无公告!"
			  RS("ContentLen")=500
			  RS("Recommend")=0
			  RS("Status")=SpaceStatus
			 RS.Update
			 RS.Close
			 '判断是企业会员，自动开通企业空间
				 On Error Resume Next
				 If UserType=11111 then  '===============默认不开通===============
				   Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
				   RS.Open "Select top 1 * From KS_EnterPrise Where 1=0",conn,1,3
				   RS.AddNew
				   
	   			     RS("UserName")=UserName
					 RS("CompanyName")=KS.S("KS_company")
					 RS("Province")=Province
					 RS("City")=City
					 RS("Address")=Address
					 RS("ZipCode")=Zip
					 RS("ContactMan")=RealName
					 RS("TelPhone")=OfficeTel
					 RS("Fax")=Fax
					 RS("AddDate")=Now
					 RS("Recommend")=0
					 RS("ClassID")=0
					 RS("SmallClassID")=0
					 RS("Status")=SpaceStatus
					  
				    If IsObject(FieldsXml) Then
						 on error resume next
						 Dim objNode,i,j,objAtr
						 Set objNode=FieldsXml.documentElement 
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								Execute("RS(""" & objAtr.Attributes.item(0).Text & """)=KS.S(""" & objAtr.Attributes.item(1).Text &""")")
						 Next
				
					   End If
				   RS.Update 
				   RS.Close
				 End If
			 End If
		    '==================================
			 Set RS=Nothing
		    If EmailCheckTF=0 Then
			If EnabledSubDomain Then
				Response.Cookies(KS.SiteSn).domain=RootDomain					
			Else
               Response.Cookies(KS.SiteSn).path = "/"
			End If
			Response.Cookies(KS.SiteSn)("UserID") = UserID
			Response.Cookies(KS.SiteSn)("UserName") = UserName
			Response.Cookies(KS.SiteSn)("PassWord") = PassWord
			Response.Cookies(KS.SiteSn)("RndPassword") = RndPassword
			End If
			Session(KS.SiteSN&"UserInfo")=""
			'-----------------------------------------------------------------
			'Ucenter系统整合
			'-----------------------------------------------------------------
			If API_Enable and uid>0 and API_ReguserUrl<>"" Then
			
			         'dim mysqlconnstr:mysqlconnstr="driver={mysql odbc 3.51 driver};database=data;server=(localhost);USER=root;password=123456"
			         dim mysqlconnstr:mysqlconnstr=API_ReguserUrl
                     dim mysqlconn:set mysqlconn = server.createobject("adodb.connection") 
                     mysqlconn.open mysqlconnstr

                     dim mysqlrs:set mysqlrs=mysqlconn.execute("select password from pre_ucenter_members where uid="&uid )

                     mysqlconn.execute("insert into pre_common_member (uid,email,username,password) values ("&uid&",'"&Email&"','"&username&"','"&mysqlrs("password")&"')" ) 
                     mysqlrs.close
                     set mysqlrs=nothing
                     set mysqlconn=nothing

               
                     response.write uc_user_synlogin(uid) '返回javascript分别调用各个应用进行登陆

				
			End If
			'-----------------------------------------------------------------
			Call ShowRegResult(ReturnInfo)
    End If	 
          
		End Sub
		
		Function GetCouponNum()
		   Do While True
			 GetCouponNum = "C" & KS.MakeRandom(10)
			 If Conn.Execute("Select CouponNum from KS_ShopCouponUser Where CouponNum='" & GetCouponNum & "'").Eof Then Exit Do
		   Loop
		End Function
		
		Sub ShowRegResult(ReturnInfo)
		   dim toUrl:toUrl="user.asp"
		   if GCls.ComeUrl<>"" then
		    toUrl=GCls.ComeUrl
		   end if
		   Gcls.ComeUrl=""
		    KS.Die "<script>alert('恭喜，会员注册成功!');location.href='" & toUrl & "';</script>"
		End Sub
		
		
End Class
%>

 

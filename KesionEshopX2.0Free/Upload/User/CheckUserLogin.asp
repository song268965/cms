﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../API/cls_api.asp"-->
<!--#include file="../api/uc_client/client.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New UserLogin
KSCls.Kesion()
Set KSCls = Nothing

Class UserLogin
        Private KS
		Private KSUser,UserID
		Private UserName,PassWord,Verifycode,ExpiresDate,RndPassword
		Private LoginVerificCodeTF
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
		    KS.Echo "<!DOCTYPE html><html>"
			If EnabledSubDomain Then
			 KS.Echo "<script>document.domain=""" & RootDomain &""";</script>" &vbcrlf
			end if
			KS.Echo "<script src=""" & KS.GetDomain & "ks_inc/jquery.js""></script>"
			KS.Echo "<script src=""" & KS.GetDomain & "ks_inc/common.js""></script>"
			KS.Echo "<body>"
			UserName=Trim(KS.R(KS.S("UserName")))
			PassWord=KS.R(KS.S("PassWord"))
			ExpiresDate=KS.R(KS.S("ExpiresDate"))
			Verifycode=KS.R(KS.S("Verifycode"))
			LoginVerificCodeTF=KS.ChkClng(KS.Setting(34))
			RndPassword=KS.R(KS.MakeRandomChar(20))
			Dim parentbox:parentbox=KS.ChkClng(Request("parentbox"))
			If parentbox=1 Then parentbox="frameElement.api.opener." Else parentbox=""
			IF UserName="" Then
		   	 KS.Die "<script>" & parentbox & "$.dialog.tips('用户名不能为空，请输入！',1,'error.gif',function(){history.back();});</script>"
			End IF
		    IF PassWord="" Then
		   	 KS.Die "<script>" & parentbox & "$.dialog.tips('登录密码不能为空，请输入！',1,'error.gif',function(){history.back();});</script>"
			End IF
			IF lcase(Trim(Verifycode))<>lcase(Trim(Session("Verifycode"))) And LoginVerificCodeTF=1 then 
		   	 KS.Die "<script>" & parentbox & "$.dialog.tips('验证码有误，请重新输入！',1,'error.gif',function(){history.back();});</script>"
			End IF
            UserName=lcase(UserName)
			
			'-----------------------------------------------------------------
			'Ucenter系统整合
			'-----------------------------------------------------------------
			If API_Enable Then
				Dim XML_login:XML_login = uc_user_login(username,password) ' 需要xml转化提取数组出来
				Dim arr_login:arr_login =  xml2array(XML_login)

				Dim uid:uid = arr_login(0)
		
				If (uid > "0") Then 
				    if conn.execute("select top 1 userid from KS_User Where UserName='" & UserName & "'").eof then
					  Call SaveReg(0,UserName,PassWord,2,"男",arr_login(3),"/images/face/boy.jpg")	
					else
					  Conn.Execute("update KS_User Set [PassWord]='" & MD5(PassWord,16) &"' where username='" & UserName &"'")
					end if
					'生成同步登录的代码
					 KS.Echo uc_user_synlogin(uid) '返回javascript分别调用各个应用进行登陆
				 elseif(uid = "-1") Then 
					'response.write "用户不存在,或者被删除"
				 elseif(uid = "-2") Then 
					'response.write "密码错"
				 else 
					'response.write "未定义"
				End If 
				
				
		    End If
			
		  '-----------------------------------------------------------------
			
			
			PassWord=MD5(PassWord,16)
			Dim Param:Param=" Where PassWord='" & PassWord & "'"
			If InStr(UserName,"@")<>0 Then
			 Param=Param & " and Email='"& UserName & "'"
			ElseIf Len(UserName)<10 and IsNumerIc(UserName) Then
			 Param=Param & " and (UserId=" & KS.ChkClng(UserName) & " or username='" & UserName &"')"
			Else
			 Param=Param & " and (UserName='" &UserName & "' or mobile='" & UserName &"')"
			End If
			 Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
			 UserRS.Open "Select top 1 * From KS_User" & Param,Conn,1,1
			 If UserRS.Eof And UserRS.BOf Then
				  UserRS.Close:Set UserRS=Nothing
				  KS.Die "<script>" & parentbox & "$.dialog.tips('你输入的用户名或密码有误，请重新输入！',1,'error.gif',function(){history.back();});</script>"
			 ElseIf UserRS("Locked")=1 Then
			   UserRS.Close:Set UserRS=Nothing
			   KS.Die "<script>" & parentbox & "$.dialog.tips('您的账号已被管理员锁定，请与管理员联系！',1,'error.gif',function(){history.back();});</script>"
			 ElseIF UserRS("Locked")=3 Then
			   UserRS.Close:Set UserRS=Nothing
			   KS.Die "<script>" & parentbox & "$.dialog.tips('您的账号还没有激活，请注意查收您的邮箱并进行激活！',1,'error.gif',function(){history.back();});</script>"
			 ElseIF UserRS("Locked")=2 Then
			   UserRS.Close:Set UserRS=Nothing
			   KS.Die "<script>$.dialog.tips('您的账号还没有通过认证！',1,'error.gif',function(){history.back();});</script>"
			 Else
			            UserName=UserRS("UserName")
			        	
			            '登录成功，更新用户相应的数据
						Dim UpdateField
						Dim ScoreTF:ScoreTF=False
						If KS.ChkClng(KS.U_S(UserRS("GroupID"),8))>0 and KS.ChkClng(KS.U_S(UserRS("GroupID"),9))>0 And datediff("n",UserRS("LastLoginTime"),now)>=KS.ChkClng(KS.U_S(UserRS("GroupID"),8)) then '判断时间
						ScoreTF=true
						End if
						UpdateField="LastLoginIP='" & KS.GetIp &"',LastLoginTime=" & SQLNowString &",LoginTimes=LoginTimes+1,RndPassword='" & RndPassword&"',IsOnline=1"
						
						'判断上一次是不是通过充值卡充值
						If UserRS("UserCardID")<>0 Then
						  Dim RSCard,ValidUnit,ExpireGroupID
						  Set RSCard=Conn.Execute("Select top 1 * From KS_UserCard Where ID=" & UserRS("UserCardID"))
						  If Not RSCard.Eof Then
						     ValidUnit=RSCard("ValidUnit")
							 ExpireGroupID=RSCard("ExpireGroupID")
							 If ValidUnit=1 Then                      '点券
							   If UserRS("Point")<=0 And ExpireGroupID<>0 Then
							     UpdateField=UpdateField & ",GroupID=" & ExpireGroupID & ",UserCardID=0,ChargeType=" & KS.ChkClng(KS.U_G(ExpireGroupID,"chargetype"))
							   End If
							 ElseIf ValidUnit=2 Then                   '有效天数
							   If UserRS("Edays")-DateDiff("D",UserRS("BeginDate"),now())<=0 And ExpireGroupID<>0 Then
							     UpdateField=UpdateField & ",GroupID=" & ExpireGroupID & ",UserCardID=0,ChargeType=" & KS.ChkClng(KS.U_G(ExpireGroupID,"chargetype"))
							   End If 
							 ElseIf ValidUnit=3 Then                  '资金
							   If UserRS("Money")<=0 And ExpireGroupID<>0 Then
							     UpdateField=UpdateField & ",GroupID=" & ExpireGroupID & ",UserCardID=0,ChargeType=" & KS.ChkClng(KS.U_G(ExpireGroupID,"chargetype"))
							   End If
							 End If
						  End If
						  RSCard.Close : Set RSCard=Nothing
						End If
						
						'签到扣分
						Call KSUser.qiandao_core(UserName,UserRs("RegDate"))
						
						Dim CurrScore:CurrScore=KS.ChkClng(UserRS("Score"))
						Dim GroupID:GroupID=KS.ChkClng(UserRS("GroupID"))
						Dim PostNum:PostNum=KS.ChkClng(UserRS("PostNum"))
						Dim ClubSpecialPower:ClubSpecialPower=KS.ChkClng(UserRS("ClubSpecialPower"))
						UserID=UserRS("UserID")
						UserRS.Close
						Set UserRS=Nothing
						
						If UpdateField<>"" Then '登录成功更新数据
						   Conn.Execute("Update KS_User Set " & UpdateField & " WHERE UserID=" & UserID)
						End If
						
						'更新论坛等级
						If ClubSpecialPower=0 Then
						  Dim RSG:Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where TypeFlag=1 and Special=0 and ClubPostNum<=" & PostNum & " And score<=" & CurrScore & " order by score desc,ClubPostNum Desc")
						  If Not RSG.Eof Then
						  Conn.Execute("Update KS_User Set ClubGradeID=" & RSG(0) & " WHERE GroupID<>1 and UserName='" & UserName & "'")
						  End If
						End If
						'更新问答等级
						Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where typeflag=0 and Special=0 and score<=" & CurrScore & " order by score desc,gradeid Desc")
						If Not RSG.Eof Then
						  Conn.Execute("Update KS_User Set GradeID=" & RSG(0) & ",GradeTitle='" & RSG(1) & "' WHERE UserName='" & UserName & "' and gradeid>5")
						End If
						'
						
						RSG.Close:Set RSG=Nothing
						
						
						
						
						If ScoreTF then 
						 Session("PopTips")=KS.U_S(GroupID,8) & "分钟后重新登录，奖励积分 +" & KS.U_S(GroupID,9) & "分！"     '用于在论坛里显示
						 Call KS.ScoreInOrOut(UserName,1,KS.ChkClng(KS.U_S(GroupID,9)),"系统",KS.ChkClng(KS.U_S(GroupID,8)) & "分钟后,重新登录奖励获得",0,0)
						End if
						
						'更新购物车的ID号
						If Not KS.IsNul(KS.C("CartID")) Then
						 Conn.Execute("Update KS_ShopPackageSelect Set UserName='" & UserName & "' where username='" & KS.C("CartID") & "'")
						 Conn.Execute("Update KS_ShoppingCart Set UserName='" & UserName & "' where username='" & KS.C("CartID") & "'")
						End If
						
							
							'自动升级用户组
							Call KSUser.UserAutoUpdateGroup(UserName)

						
							If EnabledSubDomain Then
							 Response.Cookies(KS.SiteSn).domain=RootDomain					
							Else
                             Response.Cookies(KS.SiteSn).path = "/"
							End If
						    If ExpiresDate<>"" Then Response.Cookies(KS.SiteSn).Expires = Date + 365
							Response.Cookies(KS.SiteSn)("UserID") = UserID
							Response.Cookies(KS.SiteSn)("UserName") = UserName
							Response.Cookies(KS.SiteSn)("Password") = Password
							Response.Cookies(KS.SiteSN)("RndPassword")= RndPassword
							Response.Cookies(KS.SiteSN)("GroupID")= GroupID
		
				
								
								
								If Request.Form("Action")="PopLogin" Then
								 response.write "<script>window.parent.location.reload(); </script>"
								Else
									Dim ToUrl
									 if instr(Request.ServerVariables("HTTP_REFERER"),"userlogin.asp")>0 then
									  response.write "<script>location.href='userlogin.asp';</script>"
									 elseIf InStr(lcase(Request.ServerVariables("HTTP_REFERER")), "/login") > 0 Then 
									     ToUrl="index.asp"
									ElseIf InStr(lcase(Request.ServerVariables("HTTP_REFERER")), "login") > 0 Then 
										 ToUrl= KS.GetDomain & "User/userlogin.asp?action=" & KS.S("Action")
									else
										 ToUrl= Request.ServerVariables("HTTP_REFERER")
									end if
									if GCls.ComeUrl<>"" then 
									 ToUrl=GCls.ComeUrl
									 GCls.ComeUrl=""
									 response.write "<script>location.href='" & ToUrl & "';</script>"
									Else
									 response.write "<script>location.href='" & ToUrl & "';</script>"
									End If
								End If
			 End IF
			
        End Sub
End Class
%>

 

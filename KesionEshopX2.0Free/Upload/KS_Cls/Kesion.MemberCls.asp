<!--#include file="Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
'-----------------------------------------------------------------------------------------------
Class UserCls
			Private KS,I,Node
			Public Api_QQappid,API_SinaId,API_SinaEnable,API_QQEnable,API_AlipayEnable,API_WeiXinEnable
			'---------定义会员全局变量开始---------------
			Public ID,GroupID,UserName,PassWord,RndPassword,ChargeType
			'---------定义会员全局变量结束---------------
			
			Private Sub Class_Initialize()
			  Set KS=New PublicCls
			End Sub
			Private Sub Class_Terminate()
			 Set KS=Nothing
			End Sub
            %>
			<!--#include file="WeiBoAPI.asp"-->
		    <%

			
		   '**************************************************
			'函数名：UserLoginChecked
			'作  用：判断用户是否登录
			'返回值：true或false
			'**************************************************
			Public Function UserLoginChecked()
                'on error resume next
				Dim sUserName:sUserName = KS.R(Trim(KS.C("UserName")))
				PassWord= KS.R(Trim(KS.C("Password")))
				RndPassword=KS.R(Trim(KS.C("RndPassword")))
				IF KS.IsNul(sUserName) Then
				   UserLoginChecked=false
				   Exit Function
				ElseIf IsObject(Session(KS.SiteSN&"UserInfo")) Then
				   UserLoginChecked=True
				Else
					Dim UserRs
					If DataBaseType=1 Then
						Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
						Set Cmd.ActiveConnection=conn
						Cmd.CommandText="KS_UserSQL"
						Cmd.CommandType=4	
						CMD.Prepared = true 	
						Cmd.Parameters.Append cmd.CreateParameter("@username",200,1,50,susername)
						Cmd.Parameters.Append cmd.CreateParameter("@password",200,1,50,password)
						Set UserRs=Cmd.execute
						Set CMD=Nothing
				   Else
					   Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
					   UserRS.Open "Select top 1 a.*,b.SpaceSize From KS_User a inner join KS_UserGroup b on a.groupid=b.id Where UserName='" & sUserName & "' And PassWord='" & PassWord & "'",Conn,1,1
				   End If
					IF UserRS.Eof And UserRS.Bof Then
					  UserLoginChecked=false
					  Exit Function
					Else
					  If KS.ChkClng(KS.Setting(35))=1 And trim(RndPassword)<>trim(UserRS("RndPassword")) Then
				         UserLoginChecked=false
						 Response.Write ("<script>alert('发现有人正在使用你的账号，你被迫退出！');parent.location.href='" & KS.GetDomain & "User/UserLogout.asp';</script>")
					     Response.end
					  End If
						
						if KS.ChkClng(UserRS("qiandao"))>0 then
							call qiandao_core(UserRS("UserName"),UserRS("RegDate"))'签到扣分
						end if
						 
						 
						  '更新活动时间及在线状态
						  If Not KS.IsNul(session("setonlinestatus")) Then
						   Conn.Execute("Update KS_User Set LastLoginTime=" & SQLNowString & " Where UserName='" & UserRS("UserName") & "'")
						  Else
						   Conn.Execute("Update KS_User Set LastLoginTime=" & SQLNowString & ",IsOnline=1 Where UserName='" & UserRS("UserName") & "'")
						  End If
						  
						  '更新其它会员的在线情况
						  If KS.IsNUL(Application("LastUpdateTime")) or (isDate(Application("LastUpdateTime")) and DateDiff("n",Application("LastUpdateTime"),Now)>CLng(KS.Setting(8))) Then
							 Application("LastUpdateTime")=Now
							 Conn.Execute("Update KS_User set isonline=0 WHERE DateDIff(" & DataPart_S &",LastLoginTime," & SQLNowString & ") > "& CLng(KS.Setting(8)) &" * 60")
						  End If
						  
						  Set Session(KS.SiteSN&"UserInfo")=KS.RsToXml(UserRS,"row","")  '写入session
						  
						  UserLoginChecked=true
					End if
					UserRS.Close:Set UserRS=Nothing
			   End IF
			   If IsObject(Session(KS.SiteSN&"UserInfo")) Then
				   Set Node=Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row")
				   UserName=Node.SelectSingleNode("@username").text
				   GroupID=Node.SelectSingleNode("@groupid").text
				   ChargeType=Node.SelectSingleNode("@chargetype").text
			   End If
			End Function
			
			Function qiandao_core(UserName,regdate)
			 	'===========签到扣分＝＝＝＝＝＝＝＝＝＝＝
						dim RSkc,dayks,dayqds,i,RS,lxscore,day_end
						i=0
						if KS.ChkClng(Ks.Setting(201))=1 and isDate(Ks.Setting(206)) then
							lxscore=ks.chkclng(Ks.Setting(205))
							dayks=CDate(Ks.Setting(206))
							dayks=datediff("d",dayks,now())
							dayqds=ks.chkclng(conn.execute("select count(id) from KS_qiandao where username='" & username &"'")(0))
							if dayks-dayqds>0 then
								if dayqds=0 then'日期
									day_end=CDate(Ks.Setting(206))
								else
									day_end=conn.execute("Select top 1 adddate From KS_qiandao  where username='" & username &"' order by adddate desc")(0)
								end if
								
								if regdate>day_end then day_end=regdate
								dayks=datediff("d",day_end,now())

								for i=1 to dayks-dayqds
									if dayqds=0 and  i=1 then 
										day_end=CDate(Ks.Setting(206))
										if regdate>day_end then day_end=regdate
									else
										day_end=dateadd("d",1,day_end) 
									end if
									Set RS=Server.CreateObject("ADODB.RECORDSET")
									RS.Open "Select top 1 * From KS_qiandao",conn,1,3
									RS.AddNew
									rs("Content")="未签到"
									rs("qdxq")=0
									rs("AddDate")=day_end
									rs("username")=username
									rs("UserIP")=ks.getip
									rs("Status")=1										
									RS.Update
									RS.close
									Set RS = Nothing
									Call KS.ScoreInOrOut(UserName,2,lxscore,"system",day_end&"未签到扣分！",0,0)
								next
							end if
						end if
					'=====================================================
			End Function
			
			Function GetUserInfo(ByVal FieldName)
			   If KS.IsNul(UserName) Or KS.IsNul(PassWord) Then Exit Function
			   If Not IsObject(Session(KS.SiteSN&"UserInfo")) Then UserLoginChecked
			   If IsObject(Session(KS.SiteSN&"UserInfo")) Then
				   Set Node=Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@" & lcase(FieldName))
				   If Not Node Is Nothing Then  GetUserInfo=Node.Text Else GetUserInfo=""
			   End If
			End Function


			Public Property Get GetEdays()
					GetEdays = GetUserInfo("Edays")-DateDiff("D",GetUserInfo("BeginDate"),now())
			End Property

			'可用积分
			Function GetScore()
			 if KS.ChkClng(GetUserInfo("score"))-KS.ChkClng(GetUserInfo("scorehasuse"))>0 then
			  GetScore=KS.ChkClng(GetUserInfo("score"))-KS.ChkClng(GetUserInfo("scorehasuse"))
			 Else
			  GetScore=0
			 End If
			End Function
			
			'总充值金额
			Function GetConsumMoney(UserName)
			    dim m:m=0
			    Dim rs:set rs=server.CreateObject("adodb.recordset")
				rs.open "select Money from ks_logmoney where IncomeOrPayOut=1 and username='" &username &"'",conn,1,1
				 do while not rs.eof
				    m=m+rs(0)
				 rs.movenext
				 loop
				 rs.close:set rs=nothing
				GetConsumMoney=m
			End Function
			
			'判断自动升级用户组
			Function UserAutoUpdateGroup(UserName)
			  If KS.IsNul(UserName) Then Exit Function
			  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select top 1 groupid,BeginDate,Edays,score from ks_user where username='" & username & "'",conn,1,1
			  If RS.Eof And RS.Bof Then
			    RS.Close:Set RS=Nothing
				Exit Function
			  End If
			  Dim UpdateGroupID,UpdateMoney,AutoUpdatePostNum,AutoUpdateScore,MyGroupID:MyGroupID=RS("GroupID")
			  Dim BeginDate:BeginDate=rs("BeginDate")
			  Dim ValidDays:ValidDays=rs("Edays")
			  Dim Score:Score=RS("Score")
			  Dim MyPostNum:MyPostNum=KS.ChkClng(Conn.Execute("Select count(1) From KS_ItemInfo Where Inputer='" & UserName &"' and Verific=1")(0))
			  RS.Close
			  if MyGroupID=1 then Exit Function:Set RS=Nothing '管理员不能升级
			  rs.open "select id,groupname,Descript,AutoUpdatePostNum,AutoUpdateScore,AutoUpdateXH from ks_usergroup where [type]=1 and id<>1 and AutoUpdateTF=1 order by id desc",conn,1,1
			  UpdateGroupID=0
			  Do While Not RS.Eof
			   UpdateMoney=RS("AutoUpdateXH"):If Not IsNumeric(UpdateMoney) Then UpdateMoney=0
			   AutoUpdatePostNum=KS.ChkClng(RS("AutoUpdatePostNum"))
			   AutoUpdateScore=KS.ChkClng(RS("AutoUpdateScore"))
			 
			    If Score>=AutoUpdateScore And MyPostNum>=AutoUpdatePostNum And Round(GetConsumMoney(UserName))>=Round(UpdateMoney) and MyGroupID<RS("ID") Then UpdateGroupID=RS("ID")
			  
			  RS.MoveNext
			  Loop
			  RS.CLose
			  If UpdateGroupID<>0 Then
			        If EnabledSubDomain Then
							 Response.Cookies(KS.SiteSn).domain=RootDomain					
					Else
                             Response.Cookies(KS.SiteSn).path = "/"
					End If
					Response.Cookies(KS.SiteSN)("GroupID")= UpdateGroupID
							
			    rs.open "select top 1 * from ks_usergroup where id=" & UpdateGroupID,conn,1,1
				if not rs.eof then
				   if RS("ChargeType")=2 then
					dim tmpDays:tmpDays=ValidDays-DateDiff("D",BeginDate,now())
					if tmpDays>0 then
						conn.execute("update ks_user set GroupID=" & UpdateGroupID & " where username='" & username & "'")
					else
					    Conn.Execute("Update KS_User Set ChargeType=" & RS("ChargeType") & ",EDays=" & RS("ValidDays") & ",UserType=" & RS("UserType") &",GroupID=" & UpdateGroupID & " Where UserName='" & UserName &"'")
					end if
				   else
					Conn.Execute("Update KS_User Set ChargeType=" & RS("ChargeType") & ",UserType=" & RS("UserType") &",GroupID=" & UpdateGroupID & " Where UserName='" & UserName &"'")
				   end if
					'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
					Dim MailContent:MailContent="尊敬的会员" & UserName&"!<br/>&nbsp;&nbsp;&nbsp;&nbsp;恭喜您升级为“" & KS.GetUserGroupName(UpdateGroupID) & "”会员级别。您可以<a href=User_EditInfo.asp?Action=permissions><u>点此</u></a>查看相关会员级别权利！"
					Call KS.SendInfo(username,"系统","恭喜，您的会员等级升级了！",MailContent)
					Session(KS.SiteSN&"UserInfo")=""
				end if
				rs.close
				set rs=nothing
			  End If
			End Function


			Sub InnerLocation(msg)
				KS.Echo "<script type=""text/javascript"">$('#locationid').html(""" & msg & """);</script>"
			End Sub
		    
			'取得权限
            Function CheckPower(OpType)
			  CheckPower=KS.FoundInArr(KS.U_G(GroupID,"powerlist"),OpType,",")
			End Function
			Sub CheckPowerAndDie(OpType)
			   If UserLoginChecked=false Then
			      If KS.S("From3G")="true" Then
				   KS.Die "<script>top.location.href='" & KS.Setting(3) & KS.WSetting(4) & "/login.asp';</script>"
				  Else
			       KS.Die "<script>top.location.href='Login';</script>"
				  End If
			   End If
			   If CheckPower(OpType)=false Then
			    KS.ShowError("对不起,你没有此项操作的权限!")
			   End If
			End Sub
			
			'用户上传目录
			Function GetUserFolder(UserName)
			   If KS.HasChinese(UserName) Then
			   Dim Ce:Set Ce=new CtoeCls
			   UserName="[" & Ce.CTOE(KS.R(UserName)) & "]"
			   Set Ce=Nothing
			   End If
			   GetUserFolder=KS.Setting(3)&KS.Setting(91)&"User/" & username & "/"
			End Function
			
			'保存远程图片
			Function SaveBeyoundFile(Str)
			    If KS.ChkClng(KS.Setting(92))=1 And Not KS.IsNul(Str) Then
				  Dim FormPath:FormPath =KS.Setting(3) & KS.Setting(91)& "user/" & GetUserInfo("userid") & "/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/" 
				 KS.CreateListFolder (FormPath)
				 SaveBeyoundFile = KS.ReplaceBeyondUrl(Str, FormPath) 
				Else
				 SaveBeyoundFile = Str
				End If
			End Function
			
           Sub CheckMoney(ChannelID)
		     
			 '检查每个模型每天最多发信息量
			 If KS.ChkCLng(KS.U_S(GroupID,2))<>-1 Then
			   If KS.ChkClng(Conn.Execute("Select Count(*) From " & KS.C_S(ChannelID,2) &" Where inputer='" & username & "' and Year(AddDate)=" & Year(Now) & " and Month(AddDate)=" & Month(Now) & " and Day(AddDate)=" & Day(Now))(0))>=KS.ChkCLng(KS.U_S(GroupID,2)) Then
			   	KS.Die "<script>$.dialog.tips('对不起,本频道限定每个会员每天只能发布<span style=""color=red"">" & KS.ChkCLng(KS.U_S(GroupID,2)) & "</span>条信息!',3,'error.gif',function(){history.back();});</script>"
			   End If
			 End If
			 
		     If datediff("n",GetUserInfo("RegDate"),now)<KS.ChkClng(KS.C_S(ChannelID,49)) and KS.ChkClng(KS.C_S(ChannelID,49))<>0 Then
			 KS.Die "<script>$.dialog.tips('本频道要求新注册会员" & KS.ChkClng(KS.C_S(ChannelID,49)) & "分钟后才可以发表!',3,'error.gif',function(){history.back();});</script>"
			 End If
		     If cdbl(KS.C_S(ChannelID,18))<0 And cdbl(GetUserInfo("Money"))<cdbl(abs(KS.C_S(ChannelID,18))) Then
			  ks.die "<script>$.dialog.tips('在本频道发布信息最少需要消费资金" & abs(KS.C_S(ChannelID,18)) & "元,您当前可用资金为" & GetUserInfo("Money") & "元,请充值续费!',3,'error.gif',function(){history.back();});</script>"
		     End If
		     If cdbl(KS.C_S(ChannelID,19))<0 And cdbl(GetUserInfo("Point"))<cdbl(abs(KS.C_S(ChannelID,19))) Then
			 KS.Die "<script>$.dialog.tips('在本频道发布信息最少需要消费" & KS.Setting(45) & abs(KS.C_S(ChannelID,19)) & KS.Setting(46) & ",您当前可用" & KS.Setting(45) & "为" & GetUserInfo("Point") & KS.Setting(46) &"!',3,'error.gif',function(){history.back();});</script>"
		     End If
		     If KS.ChkClng(KS.C_S(ChannelID,20))<0 And KS.ChkClng(GetUserInfo("Score"))<abs(KS.C_S(ChannelID,20)) Then
			  	KS.Die "<script>$.dialog.tips('在本频道发布信息最少需要消费积分" & abs(KS.C_S(ChannelID,20)) & "分,您当前可用积分" & GetUserInfo("Score") & "分,请充值续费!',3,'error.gif',function(){history.back();});</script>"
		     End If
			 
			 '检查未审核信息以判断积分是否够用
			 Dim RS,XML,Node
			 Set RS=Conn.Execute("Select channelid From KS_ItemInfo Where Inputer='"& UserName&"' and verific=0")
			 If Not RS.Eof Then
			    Set XML=KS.RsToXml(rs,"row","")
			 End If
			 RS.Close:Set RS=Nothing
			 If IsObject(XML) Then
			     Dim TotalScore:TotalScore=0
				 Dim TotalPoint:TotalPoint=0
				 Dim TotalMoney:TotalMoney=0
				 Dim Num:Num=0
			    For Each Node In XML.DocumentElement.SelectNodes("row")
				  Dim ModelID:ModelID=Node.SelectSingleNode("@channelid").text
				  Dim Scores:Scores=Cint(KS.C_S(ModelID,20))
				  Dim Points:Points=Cint(KS.C_S(ModelID,19))
				  Dim Moneys:Moneys=Cint(KS.C_S(ModelID,18))
				  If Scores<0 Then
				   TotalScore=TotalScore+Scores
				  End If
				  If Points<0 Then
				   TotalPoint=TotalPoint+Points
				  End If
				  If Moneys<0 Then
				   TotalMoney=TotalMoney+Moneys
				  End If
				  Num=Num+1
				Next
				
				If TotalMoney<0 Then
				  If cint(Abs(TotalMoney)+abs(KS.C_S(ChannelID,18)))>cint(GetUserInfo("Money"))  and KS.C_S(Channelid,18)<0 Then
				   	ks.die "<script>$.dialog.tips('在本频道发布信息最少需要消费资金"& abs(KS.C_S(ChannelID,18)) & "元,您的可用资金<font color=#ff6600>" & GetUserInfo("Money") & "</font>元,因在所有投稿栏目中您有<font color=red>" & Num & "</font>篇文档未审核,导致可用资金不足!',3,'error.gif',function(){history.back();});</script>"
				  End If
				End If
				
				If TotalPoint<0 Then
				  If cint(Abs(TotalPoint)+abs(KS.C_S(ChannelID,19)))>cint(GetUserInfo("Point")) and KS.C_S(Channelid,19)<0 Then
		           	ks.die "<script>$.dialog.tips('在本频道发布信息最少需要消费"& KS.Setting(45) & abs(KS.C_S(ChannelID,19)) & KS.Setting(46) & ",您的可用" & KS.Setting(45) & "<font color=#ff6600>" & GetUserInfo("Point") & "</font>" & KS.Setting(46) & ",因在所有投稿栏目中您有<font color=red>" & Num & "</font>篇文档未审核,导致可用" & KS.Setting(45) & "不足!',3,'error.gif',function(){history.back();});</script>"
				  End If
				End If
				
				If TotalScore<0 Then
				  If cint(Abs(TotalScore)+abs(KS.C_S(Channelid,20)))>cint(GetUserInfo("Score")) and KS.C_S(Channelid,20)<0 Then
		           	ks.die "<script>$.dialog.tips('在本频道发布信息最少需要消费积分" & abs(KS.C_S(ChannelID,20)) & "分,您的可用积分<font color=#ff6600>" & GetUserInfo("Score") & "</font>分,因在所有投稿栏目中您有<font color=red>" & Num & "</font>篇文档未审核,导致可用积分不足!!',3,'error.gif',function(){history.back();});</script>"
				  End If
				End If
			 End If
		   End Sub	
		   
		   Function GetModelCharge(ChannelID)
		    Dim ChargeStr,ModelChargeType:ModelChargeType=KS.ChkClng(KS.C_S(ChannelID,34))
			 If ModelChargeType=0 Then 
				   ChargeStr=KS.Setting(45)
			 ElseIf ModelChargeType=1 Then
				  ChargeStr="资金"
			 Else
				   ChargeStr="积分"
			 End If
			GetModelCharge=ChargeStr
		  End Function
		   
		   '用户使用明细
		   Sub UseLogConsum(BasicType,ChannelID,InfoID,Title)
		     Dim Num:Num=KS.ChkClng(KS.U_S(GroupID,11))
		     If Num<>0 Then
				If KS.ChkClng(Conn.Execute("Select Count(1) From KS_LogConsum Where " & InfoID & " not in(select infoid from ks_logconsum Where year(AddDate)=" & year(Now) & " and month(AddDate)=" & month(now) & " and day(AddDate)=" & day(now) &") and year(AddDate)=" & year(Now) & " and month(AddDate)=" & month(now) & " and day(AddDate)=" & day(now) &" and UserName='" &UserName & "' and BasicType=" & BasicType)(0))>=Num Then
				 Select Case BasicType
				   Case 3 KS.Die "<script>alert('系统限制您所在的用户组级别,每人每天只能下载" & Num & "个!');window.close();</script>"
				   Case 4,7 KS.Die "<script>alert('系统限制您所在的用户组级别,每人每天只能观看" & Num & KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3) &"!');window.close();</script>"
				   Case Else
				    KS.Die "<script>alert('系统限制您所在的用户组级别,每人每天只能查看" & Num & KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3)&"!');window.close();</script>"

				 End SELECT
				End If
			 End If
		     dim rs:set rs=server.createobject("adodb.recordset")
			 rs.open "select top 1  * from KS_LogConsum where channelid=" & channelid &" and infoid=" & infoid & " and username='" & username & "'",conn,1,3
			 if rs.eof and rs.bof then
			   rs.addnew
			   rs("basictype")=basictype
			   rs("channelid")=channelid
			   rs("infoid")=infoid
			   rs("title")=title
			   rs("username")=username
			   rs("adddate")=now
			   rs("times")=1
			 else
			   rs("times")=rs("times")+1
			   rs("adddate")=now
			 end if
			  rs.update
			  rs.close:set rs=nothing
		   End Sub
		   
		   '刷新添加时间
		   Sub RefreshInfo(TableName)
		   If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))="0" Then
			   KS.Die "<script language=JavaScript>$.dialog.alert('对不起，本频道没有开通此功能!',function(){location.replace('" & Request.ServerVariables("HTTP_REFERER") &"');});</script>"
		   End If
		 If KS.ChkClng(KS.U_S(GroupID,12))=0 Then
		 	  KS.Die "<script language=JavaScript>$.dialog.alert('对不起，您没有使用此功能的权限，请联系本站管理员!',function(){location.replace('" & Request.ServerVariables("HTTP_REFERER") &"');});</script>"
		 End If
		   Dim rsf:set rsf=server.createobject("adodb.recordset")
			   rsf.open "select top 1 adddate from [" & TableName & "] where inputer='" & UserName & "' and id=" & ks.chkclng(ks.g("id")),conn,1,3
			   if rsf.eof then
			     rsf.close:set rsf=nothing
				 KS.Die "<script language=JavaScript>$.dialog.alert('参数传递出错！',function(){location.replace('" & Request.ServerVariables("HTTP_REFERER") &"');});</script>"
			   end if
			   Dim refreshtime:refreshtime=rsf(0)
			   Dim NextTime:NextTime=DateAdd("n",KS.U_S(GroupID,12),refreshtime)
			   if datediff("s",NextTime,now)<1 then
			    rsf.close:set rsf=nothing
				  KS.Die "<script language=JavaScript>$.dialog.alert('对不起，每次刷新间隔" & KS.U_S(GroupID,12) & "分钟，本条信息下次的刷新时间为" & NextTime & "以后!',function(){location.replace('" & Request.ServerVariables("HTTP_REFERER") &"');});</script>"
			   else
			     rsf(0)=now
				 rsf.update
			   end if
			   rsf.close:set rsf=nothing
			   KS.Die "<script language=JavaScript>$.dialog.alert('恭喜，刷新成功!',function(){location.replace('" & Request.ServerVariables("HTTP_REFERER") &"');});</script>"
		End Sub
		
		'投稿无需审核，生成HTML
		Sub RefreshHtml(RSObj,ChannelID)
		   If KS.C_S(ChannelID,17)=2  and (KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2) Then
		       Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set RefreshTF=1 Where ID=" & RSObj("ID"))
			  Dim KSRObj:Set KSRObj=New Refresh
			  Dim DocXML:Set DocXML=KS.RsToXml(RSObj,"row","root")
			  Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
			   KSRObj.ModelID=ChannelID
			   KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
			   Call KSRObj.RefreshContent()
			   If KS.C_S(Channelid,7)=1 Then  '生成栏目HTML
			     Dim R_RS,I,FolderIDArr:FolderIDArr=Split(left(KS.C_C(KSRObj.Node.SelectSingleNode("@tid").text,8),Len(KS.C_C(KSRObj.Node.SelectSingleNode("@tid").text,8))-1),",")
				 For I=0 To Ubound(FolderIDArr)
				   Set R_RS=Conn.Execute("Select Top 1 * From KS_Class Where ID='" & FolderIDArr(i) & "'") 
			       Call KSRObj.RefreshFolder(ChannelID,R_RS)  '调用栏目刷新函数
				   R_RS.Close
				 Next
				 Set R_RS=Nothing
			   End If
			   
			   If Split(KS.Setting(5),".")(1)="asp" Then
			   Else
				FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
				FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				Dim SaveFilePath:SaveFilePath = KS.Setting(3) & KS.Setting(5)
				Dim FileContent:FileContent = KSRObj.LoadTemplate(KS.Setting(110))
				FileContent = KSRObj.KSLabelReplaceAll(FileContent) '替换函数标签
				Call KSRObj.FSOSaveFile(FileContent, SaveFilePath)
			   End If
			   Set KSRobj=Nothing
		   End If
		   
		End Sub
		
		   
		   '删除模型信息数据
		   Sub DelItemInfo(ChannelID,ComeUrl)
		        Dim ID:ID=KS.S("ID")
				ID=KS.FilterIDs(ID)
				If ID="" Then Call KS.Alert("你没有选中要删除的" & KS.C_S(ChannelID,3) & "!",ComeUrl):Response.End
				Dim RS,DelIDS,DownField
				'判断是不是下载模型
				If KS.C_S(ChannelID,6)=3 Then DownField=",DownUrls"
				
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				If KS.ChkClng(KS.U_S(GroupID,1))=1 Then
				RS.Open "Select id " & DownField &"  From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' And ID In(" & ID & ")",conn,1,3
				Else
				RS.Open "Select id " & DownField &" From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' and Verific<>1 And ID In(" & ID & ")",conn,1,3
				End If
				
				Do While Not RS.Eof
				  If DelIds="" Then DelIDs=RS(0)   Else DelIds=DelIds & "," & RS(0)
				  '=======================================删除附件=========================
				  Dim RSD:Set RSD=Server.CreateObject("ADODB.RECORDSET")
				  RSD.Open "Select FileName From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & ID & ")",Conn,1,1
				  Do While Not RSD.Eof
				   if conn.execute("select top 1 filename From KS_UploadFiles Where InfoID not in(" & ID & ") and FileName like '%" & RSD(0) & "%'").eof Then
				    Call KS.DeleteFile(RSD(0))
				   End If
				   RSD.MoveNext
				  Loop
				  RSD.Close
				  conn.Execute ("Delete From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & rs(0) & ")")
				  
				  '下载系统删除下载文件
				  If KS.C_S(ChannelID,6)=3 Then
				    Dim DownUrls:DownUrls=RS(1)
					Dim DownArr,K,DownItemArr,DownUrl
					If Not KS.IsNul(DownUrls) Then
						DownArr=Split(DownUrls,"|||")
						For K=0 To Ubound(DownArr)
						  DownItemArr = Split(DownArr(k),"|")
						  DownUrl = Replace(DownItemArr(2),KS.Setting(2),"")
						  if conn.execute("select top 1 filename From KS_UploadFiles Where InfoID not in(" & ID & ") and FileName like '%" & DownUrl & "%'").eof Then
						   if instr(lcase(DownUrl),".asp")=0  and Instr(lcase(DownUrl),"user/" & GetUserInfo("userid") &"/")>0 then
						    Call KS.DeleteFile(DownUrl)  '删除
						   end if
						  end if
						Next
					End If
				  End If
				  '============================================================================================================
				  RS.MoveNext
				Loop
				RS.Close
				
				Dim ContentPageArr,Tids
				  RS.Open "Select * FROM " & KS.C_S(ChannelID,2) &" Where ID in(" & ID & ")", conn, 1, 1
				  Do While Not RS.EOF 
					 Dim FolderID:FolderID = Trim(RS("Tid"))
					 If Tids="" Then Tids=RS("Tid") Else Tids=Tids & "," & RS("Tid")
					 If KS.C_S(ChannelID,6)=1 Then
					  ContentPageArr = Split(RS("ArticleContent")&" ", "[NextPage]")
					 ElseIf KS.C_S(ChannelID,6)=2 Then
					  ContentPageArr = Split(RS("PicUrls"), "|||")
					 End If
					 Call DelInfoFile(ChannelID,FolderID,ContentPageArr,RS("Fname"),RS("ID"),RS("AddDate"))
					 
				 RS.MoveNext
				Loop
				  RS.Close
				Set RS = Nothing
				
				
				If KS.ChkClng(KS.U_S(GroupID,1))=1 Then
				  Conn.Execute("Delete From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' And ID In(" & ID & ")")
				  Conn.Execute("Delete From KS_ItemInfo Where Inputer='" & UserName & "' and InfoID in(" & ID & ") and channelid=" & ChannelID)

				Else
				  Conn.Execute("Delete From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' and Verific<>1 And ID In(" & ID & ")")
				  Conn.Execute("Delete From KS_ItemInfo Where Inputer='" & UserName & "' and Verific<>1 and InfoID in(" & ID & ") and channelid=" & ChannelID)
				End If
				
				If Tids<>"" Then  '重新生成HTML
				  Dim KSRObj:Set KSRObj=New Refresh
				   If KS.C_S(Channelid,7)=1 Then  '生成栏目HTML
					 Dim R_RS,I,FolderIDArr:FolderIDArr=Split(Tids,",")
					 For I=0 To Ubound(FolderIDArr)
					   Set R_RS=Conn.Execute("Select Top 1 * From KS_Class Where ID='" & FolderIDArr(i) & "'") 
					   Call KSRObj.RefreshFolder(ChannelID,R_RS)  '调用栏目刷新函数
					   R_RS.Close
					 Next
					 Set R_RS=Nothing
				   End If
				   
				   If Split(KS.Setting(5),".")(1)="asp" Then
				   Else
					FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
					FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
					Dim SaveFilePath:SaveFilePath = KS.Setting(3) & KS.Setting(5)
					Dim FileContent:FileContent = KSRObj.LoadTemplate(KS.Setting(110))
					FileContent = KSRObj.KSLabelReplaceAll(FileContent) '替换函数标签
					Call KSRObj.FSOSaveFile(FileContent, SaveFilePath)
				   End If
				   Set KSRobj=Nothing
				End If

           
				if ComeUrl="" then
				Response.Redirect("../index.asp")
				else
				Response.Redirect ComeUrl
				end if
		   End Sub
		   
		   '参数:ChannelID-模型id,FolderID-栏目ID,ContentPageArr-分页数组，FileName-文件名,AddDate-添加时间
		Sub DelInfoFile(ChannelID,FolderID,ContentPageArr,FileName,InfoID,AddDate)
		        on error resume next
		 		Dim CurrPath,FExt,Fname,TotalPage,I,CurrPathAndName
				CurrPath = KS.LoadFsoContentRule(ChannelID,FolderID,InfoID,AddDate)		 
				FExt = Mid(Trim(FileName), InStrRev(Trim(FileName), "."))    '分离出扩展名
				Fname = Replace(Trim(FileName), FExt, "")                    '分离出文件名 如 2005/9-10/1254ddd
				  		 
	    		  If IsArray(ContentPageArr) Then TotalPage = UBound(ContentPageArr) + 1 Else TotalPage=1
				  If TotalPage > 1 and  KS.C_S(ChannelID,6)<=2 Then
					For I = LBound(ContentPageArr) To UBound(ContentPageArr)
					 If I = 0 Then
					  CurrPathAndName = CurrPath & FileName
					 Else
					  CurrPathAndName = CurrPath & Fname & "_" & (I + 1) & FExt
					 End If
					 Call KS.DeleteFile(CurrPathAndName)
					Next
				  Else
				   CurrPathAndName = CurrPath & FileName
				   Call KS.DeleteFile(CurrPathAndName)
				  End If
		End Sub
		   		
			'返回专栏选择框
		  Function UserClassOption(TypeID,Sel)
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' And TypeID="&TypeID,Conn,1,1
			Do While Not RS.Eof
			  If Sel=RS(0) Then
			  UserClassOption=UserClassOption & "<option value=""" & RS(0) & """ selected>" & RS(1) & "</option>"
			  Else
			  UserClassOption=UserClassOption & "<option value=""" & RS(0) & """>" & RS(1) & "</option>"
			  End iF
			  RS.MoveNext
			Loop
			RS.Close:Set RS=Nothing
		  End Function
		  
		 '从xml中加载模型字段
	     Sub LoadModelField(ChannelID,ByRef FieldXML,ByRef FieldNode)
		    set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
			
			if FieldXML.parseError.errorCode<>0 Then
				 Call KS.CreateFieldXML(ChannelID,"")
				 FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
			End If
			
			if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
			 Set FieldNode=FieldXML.DocumentElement.SelectNodes("fielditem[showonuserform=1 && fieldtype!=13]")
			end if
	     End Sub
		  
		  
		  '取得会员中心信息添加时的自定义字段
		  Function GetDiyField(ChannelID,FieldXML,Node,FieldDictionary)
		      Dim I,K,F_Arr,O_Arr,F_Value,UnitValue,V_Arr
			  Dim O_Text,O_Value,BRStr,O_Len,F_V,fieldname,fieldtype,XTitle,XWidth,XHeight,XMaxlength
              If Node.SelectSingleNode("parentfieldname").text="0" Or KS.IsNul(Node.SelectSingleNode("parentfieldname").text) Then
				    fieldname = Node.SelectSingleNode("@fieldname").text
					fieldtype = Node.SelectSingleNode("fieldtype").text
				    XTitle    = Node.SelectSingleNode("title").text
					XWidth    = Node.SelectSingleNode("width").text
					XHeight   = Node.SelectSingleNode("height").text
					XMaxlength= Node.SelectSingleNode("maxlength").text
				    GetDiyField=GetDiyField & "<dd><div>" & XTitle & "：</div>"
					If Isobject(FieldDictionary) Then
					    F_Value=FieldDictionary.item(lcase(fieldname))
					    If Node.SelectSingleNode("showunit").text="1" Then
					    UnitValue=FieldDictionary.item(lcase(fieldname) &"_unit")
						End If
					 Else
					   if lcase(Node.SelectSingleNode("defaultvalue").text)="now" then
					   F_Value=now
					   elseif lcase(Node.SelectSingleNode("defaultvalue").text)="date" then
					   F_Value=date
					   else
					   F_Value=Node.SelectSingleNode("defaultvalue").text
					   end if
					  If Instr(F_Value,"|")<>0 Then
					    F_Value=LFCls.GetSingleFieldValue("select top 1 " & Split(F_Value,"|")(1) & " from " & Split(F_Value,"|")(0) & " where username='" & UserName & "'") 
					   End If
					 End If

				   Select Case fieldtype
				   
				    Case 14  '绑定其它模型
						  Dim Mtid:Mtid=0
						  Dim MChannelID:MChannelID=KS.ChkClng(Node.SelectSingleNode("defaultvalue").text)
						  if KS.ChkClng(KS.S("ID"))<>0 then
							  Dim RST:Set RST=Conn.Execute("select tid from " & KS.C_S(MChannelID,2) & " Where ID=" & KS.ChkClng(F_Value))
							  If NOT RST.Eof Then
							   MTid=RST(0)
							  End If
							  RST.Close  : Set RST=Nothing
						  End If
						  GetDiyField=GetDiyField & " <select class=""select"" onchange=""LoadItemInfo('" & fieldname & "',this.value," & MChannelID & ",0)"" name='" & fieldname & "_tid' id='" & fieldname & "_tid' style='width:250px'>"
				          GetDiyField=GetDiyField & " <option value='0'>--请选择" & KS.GetClassName(MChannelID) &"--</option>"
				          GetDiyField=GetDiyField & Replace(KS.LoadClassOption(MChannelID,true),"value='" & Mtid & "'","value='" & Mtid &"' selected") & " </select>"&vbcrlf
				          GetDiyField=GetDiyField & "<select class=""select"" name='" & fieldname & "' id='" & fieldname & "' style='width:300px'>"
				          GetDiyField=GetDiyField & " <option value='0'>--请选择" & KS.C_S(MChannelID,3) &"--</option>"
				          GetDiyField=GetDiyField & "</select>" &vbcrlf
						  if KS.ChkClng(KS.S("ID"))<>0 then
						  GetDiyField=GetDiyField & "<script> $(document).ready(function(){  LoadItemInfo('" & fieldname & "',$('#" & fieldname & "_tid').val()," & MChannelID & "," & KS.ChkClng(F_Value)& ")  });</script>" &vbcrlf
						  End If
						  
				     Case 2
				       GetDiyField=GetDiyField & "<textarea style=""width:" & XWidth & "px;height:" & XHeight & "px"" rows=""5"" class=""textbox"" name=""" & FieldName & """ id=""" & FieldName &""">" & F_Value & "</textarea>"
					 Case 3,11
					  If Instr(F_Value,"[#")<>0 then 
					   GetDiyField=GetDiyField & Replace(F_Value,"]","|select]")
					  Else
					   GetDiyField = GetDiyField & GetSelectOption(ChannelID,FieldDictionary,FieldXML,fieldtype,fieldname,XWidth,Node.SelectSingleNode("options").text,F_Value)
					  End If
					 Case 6
					    If Instr(F_Value,"[#")<>0 then 
					     GetDiyField=GetDiyField & Replace(F_Value,"]","|radio]")
					    Else
					     GetDiyField=GetDiyField & GetRadioOption(fieldname,Node.SelectSingleNode("options").text,F_Value)
						End If
					 Case 7
					 If Instr(F_Value,"[#")<>0 then 
					   GetDiyField=GetDiyField & Replace(F_Value,"]","|checkbox]")
					  Else
					   GetDiyField = GetDiyField & GetCheckBoxOption(fieldname,Node.SelectSingleNode("options").text,F_Value)
					  End If
					 Case 9
					 Case 10
					    If KS.IsNul(F_Value) Then F_Value=""
						if XWidth<500 then XWidth=500
						If GetEditorType()="CKEditor" Then
							 GetDiyField=GetDiyField & "<table><tr><td><textarea id=""" & fieldname &""" name=""" & fieldname &""">"& Server.HTMLEncode(F_Value) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('" & fieldname &"', {width:""" & XWidth &""",height:""" & Xheight & """,toolbar:""" & Node.SelectSingleNode("editortype").text & """,filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script></td></tr></table>"
						Else	
						 GetDiyField=GetDiyField & "<script id=""" & fieldname &""" name=""" & fieldname &""" type=""text/plain"" style=""width:" & XWidth &"px;height:" & Xheight & "px;"">" &F_Value & "</script>"
						 if XMaxlength<>0 then
							 GetDiyField=GetDiyField & "<script>setTimeout(""editor" & fieldname &" = " & GetEditorTag() &".getEditor('" & fieldname &"',{toolbars:[" & Replace(GetEditorToolBar(Node.SelectSingleNode("editortype").text),"'source',","") &"],maximumWords:" &XMaxlength & ",elementPathEnabled:false});"",10);</script>"
						 Else
							 GetDiyField=GetDiyField & "<script>setTimeout(""editor" & fieldname &" = " & GetEditorTag() &".getEditor('" & fieldname &"',{toolbars:[" & Replace(GetEditorToolBar(Node.SelectSingleNode("editortype").text),"'source',","") &"],wordCount:false,elementPathEnabled:false});"",10);</script>"
						 End If
					   End If	

					 Case Else
					   Dim MaxLength:MaxLength=XMaxlength
					   If Not IsNumerIc(MaxLength)  Or MaxLength="0" Then MaxLength=255
					   GetDiyField=GetDiyField & "<input type=""text"" maxlength=""" & MaxLength &""" class=""textbox"" style=""width:" & XWidth & "px"" name=""" & FieldName & """ id=""" & FieldName & """ value=""" & F_Value & """>"
				   End Select
				   
				   If Node.SelectSingleNode("showunit").text="1" Then 
				      If Instr(F_Value,"[#")<>0 then 
					   GetDiyField=GetDiyField & Replace(F_Value,"]","|unit]")
					  Else
					   GetDiyField=GetDiyField & GetUnitOption(fieldname,Node.SelectSingleNode("unitoptions").text,UnitValue)
					 End If
				   End If
				   if FieldType=9 Then 
				     GetDiyField=GetDiyField & "<table cellspacing=""0"" cellpadding=""0""><tr><td><input type=""text"" maxlength=""" & MaxLength &""" class=""textbox"" style=""width:" & XWidth & "px"" name=""" & FieldName & """ id=""" & FieldName & """ value=""" & F_Value & """>&nbsp;</td><td width=""170""><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?FieldName=" & FieldName & "&AllowFileExt=" & server.URLEncode(Node.SelectSingleNode("allowfileext").text) & "&MaxFileSize=" & Node.SelectSingleNode("maxfilesize").text & "&Type=Field&FieldID=" & Node.SelectSingleNode("@id").text & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='170' height='30'></iframe></td><td>"
					 If Node.SelectSingleNode("mustfilltf").text="1" Then GetDiyField=GetDiyField & "<font color=red> * </font>"
					 GetDiyField=GetDiyField & "<span style=""color:blue;margin-top:5px"">" &  Node.SelectSingleNode("tips").text & "</span></td></tr></table>"
                  Else
				   If Node.SelectSingleNode("mustfilltf").text="1" Then GetDiyField=GetDiyField & "<font color=red> * </font>"
				   GetDiyField=GetDiyField & " <span style=""color:blue;margin-top:5px"">" &  Node.SelectSingleNode("tips").text & "</span>"
				  End If
				   GetDiyField=GetDiyField & "</dd>"
				 End If
			  
		   End Function
		  
		  
		   
		   '单选
		   Function GetRadioOption(FieldName,OptionValue,SelectValue)
		      Dim O_Arr,K,O_Len,F_V,O_Value,O_Text,Str
		      O_Arr=Split(OptionValue,"\n"): O_Len=Ubound(O_Arr)
			  For K=0 To O_Len
				 F_V=Split(O_Arr(K),"|")
				 If O_Arr(K)<>"" Then
					If Ubound(F_V)=1 Then
					  O_Value=F_V(0):O_Text=F_V(1)
					Else
					  O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If trim(SelectValue)=trim(O_Value) Then
						Str=Str & "<label><input type=""radio"" name=""" & FieldName & """ value=""" & O_Value& """ checked>" & O_Text&"</label>"
					Else
						Str=Str & "<label><input type=""radio"" name=""" & FieldName & """ value=""" & O_Value& """>" & O_Text&"</label>"
				    End If
				End If
			 Next
			 GetRadioOption=Str
		   End Function
		   '多选
		   Function GetCheckBoxOption(FieldName,OptionValue,SelectValue)
		    Dim O_Arr,K,O_Len,F_V,O_Value,O_Text,Str
		     O_Arr=Split(OptionValue,"\n"): O_Len=Ubound(O_Arr)
			 For K=0 To O_Len
				 F_V=Split(O_Arr(K),"|")
				 If O_Arr(K)<>"" Then
					 If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					 Else
						O_Value=F_V(0):O_Text=F_V(0)
					 End If						   
				     If KS.FoundInArr(trim(SelectValue),trim(O_Value),",")=true Then
						 str=str & "<label><input type=""checkbox"" name=""" &FieldName& """ value=""" & O_Value& """ checked>" & O_Text & "</label>"
					 Else
						 str=str & "<label><input type=""checkbox"" name=""" &FieldName& """ value=""" &O_Value& """>" & O_Text &"</label>"
					 End If
				 End If
			Next
			GetCheckBoxOption=str
		   End Function
		   
		   '单位
		   Function GetUnitOption(FieldName,UnitOption,UnitValue)
		      dim str,K
		      str = " <select name=""" & FieldName & "_Unit"" id=""" & FieldName & "_Unit"">"
			  If Not KS.IsNul(UnitOption) Then
				  Dim UnitOptionsArr:UnitOptionsArr=Split(UnitOption,"\n")
				  For K=0 To Ubound(UnitOptionsArr)
					if trim(UnitValue)=trim(UnitOptionsArr(k)) then
					 str=str & "<option value='" & UnitOptionsArr(k) & "' selected>" & UnitOptionsArr(k) & "</option>"
					else
					 str=str & "<option value='" & UnitOptionsArr(k) & "'>" & UnitOptionsArr(k) & "</option>"                 
					end if
				  Next
			 End If
			 str=str & "</select>"
			 GetUnitOption = str
		   End Function
		   '取得下拉及联动选项
		   'Function GetSelectOption(ChannelID,UserDefineFieldValueStr,F_Arr,SelectType,FieldName,Width,OptionValue,SelectValue)
		   Function GetSelectOption(ChannelID,FieldDictionary,FieldXML,SelectType,FieldName,Width,OptionValue,SelectValue)
		     Dim Str,O_Arr,O_Len,K,F_V,O_Value,O_Text
		       If SelectType=11 Then
					str="<span id='box_" &FieldName & "'><select class='select' modified=""false"" style=""width:" & Width & "px"" id=""" & FieldName &""" name=""" &FieldName & """ onchange=""fill" & FieldName &"(this.value)""><option value=''>---请选择---</option>"
	
				Else
				 str= "<span id='box_" &FieldName & "'><select class=""select"" style=""width:" & Width & """ id=""" &FieldName &""" name="""& FieldName & """>"
				End If
				O_Arr=Split(OptionValue,"\n"): O_Len=Ubound(O_Arr)
				For K=0 To O_Len
				  F_V=Split(O_Arr(K),"|")
				  If O_Arr(K)<>"" Then
					   If Ubound(F_V)=1 Then
				 	    O_Value=F_V(0):O_Text=F_V(1)
					   Else
						O_Value=F_V(0):O_Text=F_V(0)
					   End If						   
					   If trim(SelectValue)=trim(O_Value) Then
						  str=str & "<option value=""" &O_Value& """ selected>" & O_Text & "</option>"
					   Else
						  str=str & "<option value=""" & O_Value& """>" &O_Text & "</option>"
					   End If
				   End If
			  Next
			  str=str & "</select></span>"
			  '联动菜单
			  If SelectType=11  Then
				Dim JSStr
				str=str & GetLinkAgeMenuStr(ChannelID,FieldXML,FieldDictionary,FieldName,JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
			  End If
			  GetSelectOption=str
		   End Function
		   
		   '取得联动菜单
		   Function GetLinkAgeMenuStr(ChannelID,FieldXML,FieldDictionary,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str,Node,FieldName
			 If ParentFieldName="0" Or ParentFieldName="" Then Exit Function
			 Dim PNode:Set PNode=FieldXML.DocumentElement.selectsinglenode("fielditem[parentfieldname='" & ParentFieldName &"']")
			 If not pnode is nothing Then 
			     FieldName=pnode.selectsinglenode("@fieldname").text
			     Str=Str & " <select class='select' modified=""false"" name='" & FieldName & "' id='" & FieldName & "' onchange='fill" & FieldName & "(this.value)' style='width:" & pnode.selectsinglenode("width").text & "px'><option value=''>--请选择--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();" &vbcrlf
				  Options=pnode.selectsinglenode("options").text
				  OArr=Split(Options,"\n")
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=Varr(0)
					 F=Varr(0)
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>" &vbcrlf
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& FieldName&"').empty();" &vbcrlf &_
							   "$('#"& FieldName&"').append('<option value="""">--请选择--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & FieldName & "').options[document.getElementById('" & FieldName & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}" &vbcrlf
				 Dim DefaultVAL
				 If IsObject(FieldDictionary) Then DefaultVAL=FieldDictionary.item(lcase(fieldName))
				 If Not KS.IsNul(DefaultVAL) Then
				  str=str & "<script>$(document).ready(function(){fill"&ParentFieldName&"($('select[name=" &ParentFieldName&"] option:selected').val()); $('#"& FieldName&"').val('" & DefaultVAL & "');})</script>" &vbcrlf
				 End If
				 GetLinkAgeMenuStr=str & GetLinkAgeMenuStr(ChannelID,FieldXML,FieldDictionary,FieldName,JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
		   End Function
		   
	   
		   
		   '根据用户组返回对应模型的可用栏目
			Sub GetClassByGroupID(ByVal ChannelID,ByVal ClassID,FieldName,Selbutton,oid)
				Dim SQL,K,Node,ClassStr,Pstr,TJ,SpaceStr,Xml
				KS.LoadClassConfig()
				If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
				Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
				If Xml.length=1 Then
				    For Each Node In Xml
If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Then
					  KS.Echo ("<script>alert('对不起,您没有本" & KS.GetClassName(ChannelID) &"投稿的权限!');history.back();</script>")  
					Else				   
					  KS.Echo "<font color=red><b>" & Node.SelectSingleNode("@ks1").text & "</b></font>"
				      KS.Echo "<input type='hidden' value='" & Node.SelectSingleNode("@ks0").text & "' name='" & FieldName &"' id='" & FieldName &"'>"
					End If
				  Next
				  Exit Sub
				End If
				
			    If KS.C_S(ChannelID,41)="3" Then	
				   KS.Echo "<script src=""" & KS.Setting(3) &"user/showclass.asp?FieldName=" & FieldName &"&channelid=" & ChannelID &"&classid=" & ClassID & """></script>"
				  Exit Sub
				End If

					
					KS.Echo "<select class=""select"" onchange=""if ($('#ClassID option:selected').attr('pubtf')==0){alert('系统设置不能在此" & KS.GetClassName(ChannelID) &"下发表!');return;}"
					if lcase(FieldName)="otid" then KS.Echo"LoadItemInfo('oid',this.value," & ChannelID & ",0);"
					If ChannelId=5 and lcase(FieldName)<>"otid" Then KS.Echo "getBrandList();"
					KS.Echo """ name='" & FieldName &"' id='" & FieldName &"' style='width:250px'>"
					KS.Echo "<option value='0'>--选择" & KS.GetClassName(channelid) &"--</option>" &vbcrlf
					For Each Node In Xml
					  If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or (Node.SelectSingleNode("@ks20").text="0" and Node.SelectSingleNode("@ks19").text="0") Then
					  Else
							SpaceStr=""
							TJ=Node.SelectSingleNode("@ks10").text
							If TJ>1 Then
							 For k = 1 To TJ - 1
								SpaceStr = SpaceStr & "──"
							 Next
							End If
							
							If ClassID=Node.SelectSingleNode("@ks0").text Then
								KS.Echo "<option pubtf='" & Node.SelectSingleNode("@ks20").text & "' value='" & Node.SelectSingleNode("@ks0").text & "' selected>" & SpaceStr& Node.SelectSingleNode("@ks1").text & "</option>"
							Else
								KS.Echo "<option pubtf='" & Node.SelectSingleNode("@ks20").text & "' &  value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & "</option>"
							End If
					  End If
					Next
					KS.Echo "</select>"
					
					if lcase(FieldName)="otid" then   '附属栏目
					  Call EchoOTidInfo(channelid,ClassID,Oid)	
					end if
			End Sub
			
		Sub EchoOTidInfo(OtherModel,OTid,Oid)	
				Response.Write " <select name='oid' id='oid' class=""select"" style='width:300px'>"
				Response.Write " <option value='0'>--请选择" & KS.C_S(OtherModel,3) &"--</option>"
				Response.Write "</select>"
				%>
				  <script>
				    <%
					IF KS.ChkClng(KS.S("ID"))<>0 THEN
					  response.write "$(document).ready(function(){  LoadItemInfo(""oid"",'" & OTid & "'," & OtherModel &"," & Oid &"); });"
					END IF
					%>
				</script>
		 <%End Sub

		'检查录入
		Sub CheckDiyField(FieldXML,showback)
		   Dim Node,FieldName,FieldType,XTitle,Str
		     If showback=true Then str="history.back();"
			 If FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[showonuserform=1&&fieldtype!=0&&fieldtype!=13]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
						     FieldName = Node.SelectSingleNode("@fieldname").text
							 FieldType = KS.ChkClng(Node.SelectSingleNode("fieldtype").text)
							 XTitle    = Node.SelectSingleNode("title").text
							 If Node.SelectSingleNode("mustfilltf").text="1" And KS.IsNul(KS.G(FieldName)) Then KS.Die "<script>alert('" & XTitle & "必须填写!');" & str & "</script>"
				             If (FieldType=4 or FieldType=12) And Not KS.IsNul(KS.G(FieldName)) And Not Isnumeric(KS.G(FieldName)) Then KS.Die "<script>alert('" & XTitle & "必须填写数字!');" & str & "</script>"
				             If FieldType=5 And Not KS.IsNul(KS.G(FieldName)) And Not IsDate(KS.G(FieldName)) Then KS.Die "<script>alert('" & XTitle & "必须填写正确的日期!');" & str & "</script>"
				             If FieldType=8 And Not KS.IsValidEmail(KS.G(FieldName)) and Node.SelectSingleNode("mustfilltf").text="1" Then KS.Die "<script>alert('" & XTitle & "必须填写正确的Email!');" & str & "</script>"
						Next
				  End If
			 End If
		End Sub	
		'更新自定义字段的值
		Sub AddDiyFieldValue(ByRef RS,FieldXML)
		      Dim Node,FieldName,FieldType
			  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[showonuserform=1&&fieldtype!=0]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
							 FieldName = Node.SelectSingleNode("@fieldname").text
							 FieldType = Node.SelectSingleNode("fieldtype").text
							  If (Not KS.IsNul(KS.G(FieldName)) And (FieldType="4" Or FieldType="12")) or  (FieldType<>"4" and FieldType<>"12") Then
								If FieldType="10"  Then   '支持HTML时
								 RS("" & FieldName & "")=Request.Form(FieldName)
								elseIf FieldType="5" and not isdate(KS.G(FieldName)) Then
								ElseIf FieldType<>"13" Then
								 RS("" & FieldName & "")=KS.G(FieldName)
								end if
								If Node.SelectSingleNode("showunit").text="1"  Then
								RS("" & FieldName & "_Unit")=KS.G(FieldName&"_Unit")
								End If
								If FieldType="14" Then
								 RS("" & FieldName & "_ChannelID")=KS.ChkClng(Node.SelectSingleNode("defaultvalue").text)
								End If
							 End If
						Next
				 End If
			 End If
		 End Sub
			
		 '**************************************************
		'函数名：ShowClassTree
		'作  用：返回允许投稿的目录树。
		'参  数：FolderID ----选择项ID, ChannelID-----返回频道目录树
		'返回值：允许投稿的整棵树
		'**************************************************
		Public Function ShowClassTree(ChannelID,GroupID)
				Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
				KS.LoadClassConfig()
				If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
				
				TreeStr="<table style=""margin:8px"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
				  SpaceStr=""
				      TreeStr=TreeStr & "<tr ParentID='" & Node.SelectSingleNode("@ks13").text &"'><td>" & vbcrlf
					  TJ=Node.SelectSingleNode("@ks10").text
					  If TJ>1 Then
						 For k = 1 To TJ - 1
							SpaceStr = SpaceStr & "──"
						 Next
						If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or Node.SelectSingleNode("@ks20").text="0"  Then
						 TreeStr=TreeStr& SpaceStr & "<img src='../user/images/doc.gif'><span disabled TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & " <font color=red>[X]</font></a></span>"
						Else
						  TreeStr = TreeStr & SpaceStr & "<img src='../user/images/doc.gif'><span TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & "</a></span><input type='checkbox' id='selclassid' name='selclassid' onclick=""set(this,this.value,'" & Node.SelectSingleNode("@ks1").text & "');"" value='" & Node.SelectSingleNode("@ks0").text & "'>"
						End If
					  Else
					   If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or Node.SelectSingleNode("@ks20").text="0" Then
						 TreeStr=TreeStr & "<img src='../user/images/m_list_22.gif'><span disabled TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & " <font color=red>[X]</font></a></span>"
					   Else
						 TreeStr = TreeStr & "<img src='../user/images/m_list_22.gif'><span TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & "</a></span><input type='checkbox' id='selclassid' name='selclassid' onclick=""set(this,this.value,'" & Node.SelectSingleNode("@ks1").text & "');"" value='" & Node.SelectSingleNode("@ks0").text & "'>"
						End If
					  End If
						TreeStr=TreeStr & vbcrlf & "</td>"&vbcrlf
						TreeStr=TreeStr & "</tr>" & vbcrlf
				Next
		       TreeStr=TreeStr &"</table>"
		       ShowClassTree=TreeStr
		End Function

		
		
		
		'增加动态到微博
		'参数 username 用户 note 备注 fromtype 广播来源 如1 论坛，2博文，3投稿等
		Sub AddToWeibo(username,note,fromtype)
		   Dim UserID:UserID=GetUserInfo("userid")
		  If KS.IsNul(UserID) Then UserID=KS.C("UserID")
		  dim CopyFrom
		  Dim Wbtb:Wbtb=KS.SSetting(56)&"00000000000000000000000000000000000000"
		  If mid(wbtb,KS.ChkClng(fromtype),1)<>"1" then Exit Sub
		  select case fromtype
		    case "1" CopyFrom="论坛主题"
		    case "2" CopyFrom="空间博文"
		    case "3" CopyFrom="空间相册"
		    case "4" CopyFrom="空间圈子"
		    case "5" CopyFrom="内容投稿"
		    case "6" CopyFrom="会员中心"
		    case "7" CopyFrom="企业新闻"
		    case "8" CopyFrom="企业证书"
		    case "9" CopyFrom="招聘频道"
		  end select
		  
		  Call SaveWeiBo(UserName,UserID,0,Left(note,255),CopyFrom)
		End Sub
		
		
		'发布一条微博
		'参数: UserName 发布人 UserID 发布人用户ID，TransID 转播的ID,Content 广播内容 CopyFrom 来源
		Function AddWeiBo(UserName,UserID,TransID,Content,CopyFrom)
		 Dim MaxLen:MaxLen=KS.ChkClng(KS.SSetting(34))
		 If MaxLen=0 Or MaxLen>255 Then MaxLen=255
		 if KS.ChkClng(KS.SSetting(33))<>0 And Len(Content)<KS.ChkClng(KS.SSetting(33)) Then AddWeiBo="系统限定最少要输入" & KS.SSetting(33) & "个字符，多说几个字吧！":Exit Function
		 if Len(Content)>MaxLen Then AddWeiBo="系统限定最多只能输入" & MaxLen & "个字符，少说几个字吧！":Exit Function
		 If KS.IsNul(Content) Then AddWeiBo="请输入内容！":Exit Function
		 '防发帖机
         dim kk,sarr
         sarr=split(KS.WordFilter,",")
         for kk=0 to ubound(sarr)
               if instr(content,sarr(kk))<>0 then 
                  AddWeiBo="含有非常关键词:" & sarr(kk) &",请不要非法提交恶意信息！"
				  Exit Function
               end if
         next

	 	  Call SaveWeiBo(UserName,UserID,TransID,Content,CopyFrom)
		  AddWeiBo="success"
		End Function
		
       '微博广播写入数据库
		Sub SaveWeiBo(UserName,UserID,TransID,Content,CopyFrom)
		
		  Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
		 If TransID=0 Then '不是转发
		  Dim InfoId:InfoID=0
		  If CopyFrom="论坛主题" Then
		    InfoID=Conn.Execute("Select max(id) From KS_GuestBook")(0)
		  ElseIf CopyFrom="空间博文" Then 
		    InfoID=Conn.Execute("Select max(id) From KS_BlogInfo")(0)
		  End If
		  
		  RSObj.Open "Select top 1 * From KS_UserLog where 1=0",Conn,1,3
		   RSObj.AddNew
		     RSObj("UserName")=UserName
			 RSobj("UserID")=UserID
			 RSObj("Note")=Content
			 RSObj("AddDate")=Now
			 RSObj("TransNum")=0
			 RSObj("CmtNum")=0
			 RSObj("CopyFrom")=CopyFrom
			 RSObj("InfoID")=InfoID
		   RSObj.Update
		   RSObj.MoveLast
		   Dim NewId:NewID=RSObj("ID")
		  RSObj.Close 
		Else
		   NewID=TransID
		   Conn.Execute("Update KS_UserLog Set TransNum=TransNum+1 Where ID	=" & NewID)
		End If
		
		  RSObj.Open "select top 1 * From KS_UserLogR Where 1=0" ,Conn,1,3
		  RSObj.AddNew
		    RSObj("MsgId")=NewID
			RSObj("UserId")=UserID
			RSObj("UserName")=UserName
			If TransID=0 Then '不是转发
			 RSObj("Type")=0
			 RSObj("Msg")=""
			Else
			 RSObj("Type")=1
			 RSObj("Msg")=Content
			End If
			 RSObj("Status")=1
			RSObj("transtime")=Now
		  RSObj.Update
		  RSObj.Close
		   Set RSObj=Nothing
		 '增加广播数
		  Conn.Execute("Update KS_User Set MsgNum=MsgNum+1,LastPostWeiBoTime=" & SQLNowString & ",LastPostWeiBoID=" & NewID &" Where UserName='" & UserName & "'")
		  Session(KS.SiteSN & "UserInfo")=""
		End Sub	
		
		
		
		'增加操作记录
		Sub AddUserRecord(flag,note)
		  Dim UserID:UserID=GetUserInfo("userid")
		  If KS.IsNul(UserID) Then UserID=KS.C("UserID")
		  if username="" then username="游客"
		  Conn.Execute("Insert Into KS_UserRecord([userid],[username],[flag],[note],[adddate],[userip]) values(" & KS.ChkClng(UserID) & ",'" & UserName & "'," & flag & ",'" & KS.FilterIllegalChar(replace(note,"'","""")) & "'," & SqlNowString & ",'" & KS.GetIP() & "')")
		End Sub
		
		'签到
		Sub QianDao()
		%>
		<script type="text/javascript">
			function  fqiandao(){
				 var qdxq=$("input:radio[name=qdxq]:checked").val()
				 var qdContent=escape($("#qdContent").val())
				 $.get("<%=KS.Setting(3)%>user/User_qiandao.asp",{qdxq:qdxq,qdContent:qdContent,action:"qiandao",anticache:Math.floor(Math.random()*1000)},function(d){//读取列表
				  $("#myqiandao").html("已签到").attr("onclick","").click(function(){
				   $.dialog.alert('您今天已经签到过了!');
				  });;
				  });
				 $(".qiandao_form").hide(200);
			}
		</script>
		<%
		if not conn.execute("select top 1 username from KS_qiandao where username='" & username &"' and year(AddDate)=year(" & SqlNowString & ") and month(AddDate)=month(" & SqlNowString &") and day(AddDate)=day(" & SqlNowString & ") ").eof then
			dim qiandao_d:qiandao_d="ok"
		end if
		%>
		<div class="qd_daym">
			<div  class="qd_day">
				<ul>
				<li>
				<%
				dim Weekday_Name,month_str,day_str
				Weekday_Name=WeekdayName(Weekday(now()))
				Weekday_Name=Replace(Weekday_Name,"星期","周")
				if KS.ChkClng(month(now()))<10 then month_str="0"& month(now()) else month_str=month(now()) 
				if KS.ChkClng(day(now()))<10 then day_str="0"& day(now()) else  day_str=day(now())
				Response.Write(  month_str &"."& day_str )
				%>
				</li>
				<li><%=Weekday_Name%></li>
				</ul>
			</div>
			
			<div style="position:relative;float:left;z-index:1px;">
				<div align="center" onMouseOver="this.className='qiandao qiandao_hvr'" <%if qiandao_d<>"ok" then%>onclick="$('.qiandao_form').toggle(200);"<%else%>onclick="$.dialog.alert('您今天已经签到过了！');"<%end if%>  onMouseOut="this.className='qiandao'" class="qiandao" id="myqiandao"> 
					<%if qiandao_d="ok" then Response.Write"已签到" else Response.Write"签到"%>
					
				</div>
				<div class="qiandao_form">
				<div class="qd">你今天签到了吗？</div>
				<div class="qiandao_formx" >
					<ul>
					<%
					dim ii
					for ii=0 to 19
						%>
						<li><input name="qdxq"  value="<%=ii%>" type="radio"><img src="/images/emot/<%=ii%>.gif"  style="width:18px;height:18px;" align="absmiddle"></li>
						<%
					next
					%>
					</ul>
					<div><textarea name="qdContent" class="textbox" id="qdContent" cols="" rows="" style="width:270px; height:50px;"></textarea></div>
					<div align="center" style="margin:10px;">
						<input class="button" value="开始签到" type="submit" onclick="fqiandao();">  
					</div>
				</div>
				</div>
			</div>
			
			<div  class="qd_order" onMouseOver="this.className='qd_order qiandao_hvr'" onMouseOut="this.className='qd_order'"  >
				<a href="user_qiandao.asp" >记录</a> 
			</div>
			
		</div>
		<%
		End Sub

           '头部
		 Sub Head()
		   %>
			<div  class="notice" style="height:30px;line-height:30px;padding-left:6px"><a href="<%=KS.GetDomain%>" target="_parent">网站首页</a> >> <a href="<%=KS.GetDomain%>user/index.asp">会员中心</a> >> <span id="locationid">操作导航</span>  </div>
		   <%
		End Sub
		   '门户头部
		Sub SpaceHead()
		   %>
			<div  class="notice" style="height:30px;line-height:30px;padding-left:6px"><a href="<%=KS.GetDomain%>" target="_parent">网站首页</a> >> <a href="<%=KS.GetDomain%>user/space.asp">空间门户</a> >> <span id="locationid">操作导航</span>  </div>
		   <%
		End Sub
End Class
%> 

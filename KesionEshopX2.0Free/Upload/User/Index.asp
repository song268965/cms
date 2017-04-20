<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New UserList
KSCls.Kesion()
Set KSCls = Nothing

Class UserList
        Private KS,KSUser,LoginTF,TopDir,KSR
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<!--#include file="../KS_Cls/UbbFunction.asp"-->
		<%
       Public Sub loadMain()
		'Call KSUser.Head()
		'Call KSUser.InnerLocation("会员首页")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		TopDir=KSUser.GetUserFolder(ksuser.getuserinfo("userid"))
		
		'==========================设置在线状态================================
        If Request.QueryString("action")="offline" then
		 session("setonlinestatus")="true"
		 Conn.Execute("Update KS_User Set isonline=0 where username='" & KSUser.UserName &"'")
		 If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@isonline").Text=0
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		ElseIf Request.QueryString("action")="setonline" Then
		 session("setonlinestatus")="true"
		 Conn.Execute("Update KS_User Set isonline=1 where username='" & KSUser.UserName &"'")
		 If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@isonline").Text=1
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End If
		'===================================================================

		%>
	  
		     <h2><span><a href="User_EditInfo.asp?action=PassInfo">修改密码</a> | <a href="User_EditInfo.asp">编辑资料</a> | <a href="User_EarnScore.asp" target="_self" class="toptitle" style="color:red;font-weight:bold">我要赚积分</a></span>我的信息</h2>
			 
			 <%
			  Dim UserFaceSrc:UserFaceSrc=KSUser.GetUserInfo("UserFace")
			  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
			%>
			 <div class="ar_r_t"><a href="User_EditInfo.asp?Action=face" title="修改头像"><img width="120" height="120" src="<%=UserFaceSrc%>" onerror="this.onerror=null;this.src='images/noavatar_middle.gif'" alt="修改头像"></a><br /><a href="User_EditInfo.asp?Action=face">[修改头像]</a> </div>			
			
			 <div class="userrightdetail">
			   <li>
			   <span class="uname">您的账号：
			  <%If Not KS.IsNul(KSUser.GetUserInfo("realname")) Then
			     response.write KS.CheckXSS(KSUser.GetUserInfo("realname") &"(" & KSUser.UserName &")")
			    Else
				 response.write KSUser.UserName
				End If		 
			  %>
			  </span> </li>
			  <li> <span class="uid">您的ID号：<%=KSUser.GetUserInfo("UserID")%> </span>
			  
			   </li>
			   <li>您所在用户组：<%=KS.U_G(KSUser.GroupID,"groupname")%></li>
			   <li>注册时间：<%=KSUser.GetUserInfo("regdate")%></li>
               <li>登录次数：<%=KSUser.GetUserInfo("logintimes")%> 次</li>
               <li>最后活动时间：<%=KS.GetTimeFormat(KSUser.GetUserInfo("LastLoginTime"))%></li>
               <li>最后登录IP：<%=KSUser.GetUserInfo("lastloginip")%></li>
			   <li class="spacelimit full">您的空间上限容量为<font color=red><%=round(KSUser.GetUserInfo("SpaceSize")/1024,2)%>M</font>
			 已使用<font color=green><%dim sy:sy=Round(KS.GetFolderSize(TopDir)/1024/1024,2)
						if sy<1 then response.write "0" & sy else response.write sy%>M</font>的空间容量 <a class="userview" href="User_Files.asp">查看</a>
			 </li>
			 <li class="message full">我的消息：
			 <a href="user_message.asp">短消息(<%=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)%>)条</a> <%If KS.SSetting(0)="1" Then%>
				 | <a href="User_Message.asp?action=Message">空间留言(<%=Conn.Execute("Select Count(ID) From KS_BlogMessage Where username='" &KSUser.UserName &"' And readtf=0")(0)%>)条</a> | <a href="user_message.asp?action=friendrequest">好友请求(<%=conn.execute("select count(id) from ks_friend where friend='" & ksuser.username & "' and accepted=0")(0)%>)条</a> | <a href="user_message.asp?action=Comment">日志评论(<%=Conn.Execute("Select Count(ID) From KS_BlogComment Where username='" &KSUser.UserName &"' And readtf=0")(0)%>)条</a> | <a href="user_message.asp?action=photoComment">照片评论(<%=Conn.Execute("Select Count(ID) From KS_photoComment Where username='" &KSUser.UserName &"' And readtf=0")(0)%>)条</a>
				 <%end if%>
			 </li>
			 
			 </div>
			
			
			 
			 <div class="clear"></div>
			 
			 
			 
			 <h2><span><a href="user_payonline.asp" target="_self"><font color="red">我要充值</font></a></span>我的财富</h2>
			 
			 <div class="jffs">
			   您的计费方式为<%if KSUser.ChargeType=1 Then 
										  Response.Write "<font color='#ff6600'>扣点数</font>"
										  ElseIf KSUser.ChargeType=2 Then
										   Response.Write "有效期,到期时间：" & cdate(KSUser.GetUserInfo("BeginDate"))+KSUser.GetUserInfo("Edays") 
										  Else
										   Response.Write "<font color='#ff6600'>永不过期</font>"
										  End If
										  %>
										  
										  <%
									   if KS.ChkClng(KSUser.GetUserInfo("UserCardID"))<>0 then
									      Dim RSCard,ValidUnit,ExpireGroupID,ExpireTips
										  Set RSCard=Conn.Execute("Select top 1 * From KS_UserCard Where ID=" & KSUser.GetUserInfo("UserCardID"))
										  If Not RSCard.Eof Then
											 ValidUnit=RSCard("ValidUnit")
											 ExpireGroupID=RSCard("ExpireGroupID")
											 If ValidUnit=1 Then                      '点券
											   If KSUser.GetUserInfo("Point")<=10 And ExpireGroupID<>0 Then
											    ExpireTips="您的" & KS.Setting(45) & "快使用完毕了"
											   End If
											 ElseIf ValidUnit=2 Then                   '有效天数
											   If KSUser.GetUserInfo("Edays")<=7 And ExpireGroupID<>0 Then
											    ExpireTips="您还有" & KSUser.GetUserInfo("Edays") & "天就过期了"
											   End If 
											 ElseIf ValidUnit=3 Then                  '资金
											   If KSUser.GetUserInfo("Money")<=10 And ExpireGroupID<>0 Then
												 ExpireTips="您的账户资金快使用完毕了"
											   End If
											 End If
										  End If
										  RSCard.Close : Set RSCard=Nothing
										  If ExpireTips<>"" and ExpireGroupID<>0  then
										  response.write "<br/><span style='color:red'>温馨提示：您上一次使用充值卡充值，" & ExpireTips & "，<br/>过期后您将自动转为<font color='blue'>"  & KS.U_G(ExpireGroupID,"groupname") & "</font>，为了更好的服务请尽快充值！</span>"
										  end if
									   end if
									  %> 
										  
			 </div>
			 
			  <div class="mymoney" >
					  
					
									
						
										<li>
										<span><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1)%>元</span>
										<p>可用资金</p>
										</li>
										
										<li>
										 <span><%=formatnumber(KSUser.GetUserInfo("Point"),0,-1) & "" & KS.Setting(46)%></span>
										 <p>可用<%=KS.Setting(45)%></p>
										</li>
										
										<li>
										  <span style="color:green;"><%=KSUser.GetUserInfo("score")%>分</span>
										  <p>总积分</p>
										</li>
										<li>
										  <span style="color:red;"><%=KSUser.GetScore%>分</span>
										  <p>可用积分</p>
										</li>
										<li>
										  <span><%=KS.ChkClng(KSUser.GetUserInfo("scorehasuse"))%>分</span>
										  <p>已消费积分</p>
										</li>
										
			                          
					  
			                             
			 </div>
			 
				  <div class="clear"></div>

				<%
				Call KSUser.initialOpenId '初始化API接口数据
				if cbool(KSUser.API_QQEnable)=true or cbool(KSUser.API_SinaEnable)=true or cbool(KSUser.API_AlipayEnable)=true or cbool(KSUser.API_WeiXinEnable)=true then%>
				 <h2>账号通：</h2><div class="zht"><%
				 
				  if cbool(KSUser.API_QQEnable)=true then
				    if not ks.isnul(ksuser.getuserinfo("qqopenid")) then
					  response.write "<span title='qq登录已绑定'><img src='../images/default/icon_qq.png' align='absmiddle' alt='qq登录已绑定'/>QQ</span>&nbsp;&nbsp;"
					else
					  response.write "<span title='qq登录未绑定'><img src='../images/default/icon_qq_no.png' align='absmiddle' alt='qq登录未绑定'/>QQ</span>&nbsp;&nbsp;"
					end if
				  end if

				  if cbool(KSUser.API_WeiXinEnable)=true then
				    if not ks.isnul(ksuser.getuserinfo("weixinopenid")) then
					  response.write "<span title='微信登录已绑定'><img src='../images/default/icon_weixin.png' align='absmiddle' alt='微信登录已绑定'/>微信登录</span>&nbsp;&nbsp;"
					else
					  response.write "<span title='微信登录未绑定'><img src='../images/default/icon_weixin_no.png' align='absmiddle' alt='微信登录未绑定'/>微信登录</span>&nbsp;&nbsp;"
					end if
				  end if

				  
				  if cbool(KSUser.API_SinaEnable)=true then
				    if not ks.isnul(ksuser.getuserinfo("sinaid")) then
					  response.write "<span title='新浪微博已绑定'><img src='../images/default/icon_sina.png' align='absmiddle' title='新浪微博已绑定'/>  新浪微博</span>&nbsp;&nbsp;"
					else
					  response.write "<span title='新浪微博未绑定'><img src='../images/default/icon_sina_no.png' align='absmiddle' title='新浪微博未绑定'/>  新浪微博</span>&nbsp;&nbsp;"
					end if
				  end if
				  if cbool(KSUser.API_AlipayEnable)=true then
				    if not ks.isnul(ksuser.getuserinfo("alipayid")) then
					  response.write "<span title='支付宝已绑定'><img src='../images/default/icon_alipay.png' align='absmiddle' title='支付宝已绑定'/> 支付宝</span>"
					else
					  response.write "<span title='支付宝未绑定'><img src='../images/default/icon_alipay_no.png' align='absmiddle' title='支付宝未绑定'/> 支付宝</span>"
					end if
				  end if
				 %>
				 (<a href="user_bind.asp">绑定管理</a>)
				 </div>
			<%end if%>
		<div class="clear"></div>
		<%If not conn.execute("select top 1 id from KS_MallScore where status=1").eof and lcase(KS.GetAppStatus("mallscore"))="true" then%>
		<h2><span><a href="User_ScoreExchange.asp">更多&raquo;</a></span>兑换礼品：</h2>
		<div class="zfdf">
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
		   <tr class=titlename>
		     <td>礼品图片</td>
			 <td>礼品名称</td>
			 <td>数量</td>
			 <td>截止日期</td>
			 <td>兑换方式</td>
			 <td>操作</td>
		   </tr>
		   <%
		   dim rs:set rs=server.CreateObject("adodb.recordset")
		   RS.open "select top 5 * from KS_MallScore where status=1 order by id desc",conn,1,1
		   do while not rs.eof 
		   %>
			<tr class=tdbg>
			  <td  class="splittd"><a href='User_ScoreExchange.asp?action=showdetail&id=<%=rs("id")%>'><img src="<%=rs("photourl")%>" /></a></td>
			  <td  class="splittd"><%=rs("productname")%> <%if rs("recommend")=1 then response.write "&nbsp;<font color=red>荐</font>"%></td>
			  <td  class="splittd"><%=rs("Quantity")%> 件</td>
			  <td  class="splittd"><%=formatdatetime(rs("enddate"),2)%></td>
			  <td  class="splittd" >
			  <%if rs("chargeType")=0 Then%>
	积分<font color=red><%=Rs("Score")%></font>分
	<%else%>
	<%=KS.Setting(45)%><font color=red><%=Rs("Score")%></font><%=KS.Setting(46)%>
	<%end if%>
			  </td>
			  <td  class="splittd" ><a href='User_ScoreExchange.asp?action=showdetail&id=<%=rs("id")%>'>查看</a> <a href="user_scoreexchange.asp?action=exchange&id=<%=rs("id")%>">兑换</a></td>
			</tr>
		  <% rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  %>
		 </table>
		</div>
	<% end if
  End Sub
  
	  
End Class
%> 

﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
Set KSCls = New User_LogScore
KSCls.Kesion()
Set KSCls = Nothing

Class User_LogScore
        Private KS,KSUser
		Private CurrentPage,totalPut,TotalPages,SQL
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
		Public Sub loadMain()	
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		
		Dim TGUnit:TGUnit="积分"
		Dim TGUnit1:TGUnit1="分"
			If KS.Setting(145)="1" Then 
			 TGUnit=KS.Setting(45) : TGUnit1=KS.Setting(46)
			ElseIf KS.Setting(145)="2" Then 
			 TGUnit="人民币": TGUnit1="元"
		    End If
		Call KSUser.InnerLocation("我要赚" & TGUnit)
	  %>

	 <style type="text/css">
	  .splittd{height:100px;font-size:16px;padding-left:10px;}
	  .red{color:red;}
	 </style>
	  
	<table border="0" align="center" class="border">
        <%if KS.Setting(140)="1" Then%>
				<tr>
				  <td class="splittd">
				      <table>
					    <tr>
						  <td><strong>任务名称：</strong></td>
						  <td><span class="red">将本站推荐给朋友将获得<%=TGUnit%></span></td>
						</tr>
						<tr>
						 <td><strong>任务介绍：</strong></td>
						 <td>成功推荐一个访问者,您就可以增加 <font color=red><%=KS.Setting(141)%></font> <%=TGUnit1%><%=TGUnit%>。赶快行动吧！</td>
						</tr>
						<tr>
						 <td valign="top"><strong>复制代码：</strong></td>
						 <td>
						  <div id="copytext" class="copytext"><%=Replace(Replace(Server.HTMLEncode(KS.Setting(142)),"{$UID}",KSUser.UserName),"{$GetSiteUrl}",KS.GetDomain)%></div>
						  <br/><button class="pn" type="button" onClick="copyToClipboard(document.getElementById('copytext').innerHTML);"><strong>复制代码</strong></button>
						 </td>
						 </tr>
						 </table>
													
					</td>
				 </tr>
		 <%end if%>
	  <%if KS.Setting(143)="1" Then%>
		   <tr>
				<td class="splittd"><br/>
				  <table>
				   <tr>
				    <td><strong>任务名称：</strong></td>
					<td><span class="red">引导朋友注册将获得<%=TGUnit%></span></td>
				   </tr>
				   <tr>
				    <td><strong>任务介绍：</strong></td>
					<td>成功推荐一个用户注册,您就可以增加 <font color=red><%=KS.Setting(144)%></font> <%=TGUnit1%><%=TGUnit%>,同一天内推荐同一个IP的用户注册，只计一次！</td>
				   </tr>
				   <tr>
				    <td valign="top"><strong>复制代码： </strong>
					</td>
					<td>
					 <div  class="copytext" id="copytext1"><%=KS.GetDomain%>user/reg/?uid=<%=KSUser.GetUserInfo("UserId")%></div>
									<br/>
									<button class="pn" name="button2" type="button" onClick="copyToClipboard($('#copytext1').text());"><strong>复制链接</strong></button>	</td>
						</tr>
						</table>
					 </td>
				   </tr>
			 <%end if%>
			 
				<tr>
				  <td class="splittd"><br />
				      <table>
					    <tr>
						  <td><strong>任务名称：</strong></td>
						  <td><span class="red">邮件邀请好友注册</span></td>
						</tr>
						<tr>
						 <td><strong>任务介绍：</strong></td>
						 <td>给好友发送邀请邮件，好友通过收到的邮件里的链接成功注册为本站会员，您就可以增加 <font color=red><%=KS.Setting(144)%></font> <%=TGUnit1%><%=TGUnit%>，同一天内推荐同一个IP的朋友注册，只计一次！</td>
						</tr>
						<tr>
						 <td valign="top"></td>
						 <td><button class="pn" type="button" onClick="location.href='User_friend.asp?Action=mail'"><strong>我要参加</strong></button>
						 </td>
						 </tr>
						 </table>
													
					</td>
				 </tr>
			 
			 
		 </table>
				
		  <%
  End Sub
    
  
End Class
%> 

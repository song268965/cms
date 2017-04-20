<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_UserGroup
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserGroup
        Private KS
		Private MaxPerPage,CurrentPage,TotalPut
		Private RS,Sql
		Private ComeUrl

		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	        Response.Write "<!DOCTYPE html><html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script src=""../../KS_Inc/jquery.js""></script>"
			Response.Write"<script src=""../../KS_Inc/common.js""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write "<div class='tabTitle mt20'>用户实名认证管理</div>"
            If Not KS.ReturnPowerResult(0, "KMUA10016") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If

			Select Case Trim(request("Action"))
			Case "yyrzrz" call yyrzrz()
			case "yyrzsave" call yyrzsave()
			case "sfzrz" call sfzrz()
			case "sfzrzsave" call sfzrzsave()
			case "mobilerz" call mobilerz()
			case "mobilerzsave" Call mobilerzsave()
			case "emailrz" call emailrz()
			case "emailrzsave" call emailrzsave()
			case "delrz" call delrz()
			Case else
				call main()
			End Select
			
	
		End Sub
		
		sub main()
		    dim param:param="Where IsRz<>0 or issfzrz<>0 or ismobilerz<>0"
			if request("keyword")<>"" then
			  param=param & " and username like '%" & ks.g("keyword") & "%'"
			end if
			Set rs=Server.CreateObject("Adodb.RecordSet")
			sql="select * from KS_User " & param & " order by UserID"
			rs.Open sql,Conn,1,1
		    CurrentPage=KS.ChkClng(KS.S("Page"))
			If CurrentPage < 1 Then	CurrentPage = 1
		%>
   <div class="pageCont2">
		<table border="0" align="center" width="100%" cellpadding="0" cellspacing="0">
		  <tr align="center" class="sort">
			<td  width="45">ID号</td>
			<td width="168">用户名</td>
			<td>用户类型</td>
			<td>营业执照认证</td>
			<td>身份证认证</td>
			<td>手机认证</td>
			<td>邮箱认证</td>
			<td>管理</td>
		  </tr>
		  <%
		  If RS.Eof And RS.BOf Then
		    Response.Write "<tr class='list'><td colspan=9 class='splittd' style='text-align:center'>没有人申请实名认证!</td></tr>"
		  Else
			totalPut = conn.execute("select count(1) from ks_user " & param)(0)
			If CurrentPage >1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
				RS.Move (CurrentPage - 1) * MaxPerPage
			End If
		    Dim i:I=0
			  dim isqy:isqy=false
			  do while not rs.EOF
				isqy=not conn.execute("select top 1 username from ks_enterprise where username='" & RS("UserName") & "'").eof
			  %>
			  <tr height="40" align="center" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				<td class="splittd" width="45"><%=rs("UserID")%></td>
				<td class="splittd"><%=rs("UserName")%></td>
				<td class="splittd" align="center"><%
				if isqy then
					Response.Write "<font color=blue>企业用户</font>"
				else
					Response.Write "<font color=green>个人用户</font>"
				end if
				%> </td>
				<td class="splittd" style="line-height:22px" align="center">
				<%If isqy then
				   response.write "<a href=""?action=yyrzrz&username=" & rs("username") & """>"
				   dim yyrzrz:yyrzrz=conn.execute("select top 1 isrz from ks_enterprise where username='" & rs("username") & "'")(0)
				   if yyrzrz=1 then
					response.write "<i class='icon no'></i> <font color=#999999>已认证</font>"
				   elseif yyrzrz=2 then
					response.write "<font color=red>已提交,点此审核</font>"
				   elseif yyrzrz=3 then
					 response.write "<font color=green>认证不通过，退回</font>"
				   else
					response.write "未提交"
				   end if
				   response.write "</a>"
				else
				   response.write "---"
				end if
			   %>
			   </td>
			   <td class="splittd" align="center">
				<%
				 response.write "&nbsp;&nbsp;<a href='?action=sfzrz&username=" & rs("username") & "'>"
				if rs("issfzrz")=1 then
				 response.write "<i class='icon no'></i> <font color=#999999>已认证</font>"
				elseif rs("issfzrz")=2 then
				 response.write "<font color=red>已提交,点此审核</font>"
				elseif rs("issfzrz")=3 then
				 response.write "<font color=green>认证不通过，退回</font>"
				else
				 response.write "未提交"
				end if
				 response.write "</a>"
			   %>
			   </td>
			   <td class="splittd" align="center">
				<%
				 response.write "<a href='?action=mobilerz&username=" & rs("username") & "'>"
				if rs("ismobilerz")=1 then
				 response.write "<i class='icon no'></i> <font color=#999999>已认证</font>"
				elseif rs("ismobilerz")=2 then
				 response.write "<font color=red>已提交,点此审核</font>"
				elseif rs("ismobilerz")=3 then
				 response.write "<font color=green>认证不通过，退回</font>"
				else
				 response.write "未提交"
				end if
				 response.write "</a>"
				 
			   %>
			   </td>
			   <td class="splittd" align="center">
				<%
				 
				 response.write "<a href='?action=emailrz&username=" & rs("username") & "'>"
				if rs("isemailrz")=1 then
				 response.write "<i class='icon no'></i> <font color=#999999>已认证</font>"
				elseif rs("isemailrz")=2 then
				 response.write "<font color=red>已提交,点此审核</font>"
				else
				 response.write "未提交"
				end if
				 response.write "</a>"
				%>
				 
				</td>
				<td class="splittd">
				<a href='../../space/company/rz.asp?userid=<%=rs("userid")%>' target="_blank" class="setA">查看</a>|
				<a href='?action=delrz&userid=<%=rs("userid")%>' onclick="return(confirm('确定删除该用户的所有认证信息吗？'))" class="setA">删除认证</a>
				</td>
				
			  </tr>
			  <%
			  i=i+1
			 if i>=maxperpage then exit do
			rs.MoveNext
		loop
	End If	  %>
	
	  <tr>
            <td height="30" align='right' colspan=9>
				<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
			   </td>
      </tr>
	  <form name="sform" action="KS.UserRZ.asp" method="post">
	  <tr>
            <td height="30" colspan=9>
				<strong>快速查找=></strong>  用户名：<input type="text" name="keyword" class="textbox" size="17"/> <input type="submit" value=" 搜 索 " class="button"/>
			</td>
      </tr>
	  </form>
	</table>  
		<%
			rs.Close:set rs=Nothing
		
		end sub
		
		'营业执照认证审核
		sub yyrzrz()
		  if ks.g("username")="" then
		    ks.die "<script>alert('对不起，用户名出错!');history.back();</script>"
		  end if
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_enterprise where username='" & ks.g("username") & "'",conn,1,1
		 if rs.eof and rs.bof then
		   rs.close:set rs=nothing
		   ks.die "出错啦!"
		 end if
		 %>
		 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#CompanyName").val()==''){
			  alert('请输入公司名称!');
			  $("#CompanyName").focus();
			  return false;
			}
			if ($("#BusinessLicense").val()==''){
			  alert('请输入注册号!');
			  $("#BusinessLicense").focus();
			  return false;
			}
			if ($("#LegalPeople").val()==''){
			  alert('请输入企业法人!');
			  $("#LegalPeople").focus();
			  return false;
			}
			if ($("#Address").val()==''){
			  alert('请输入公司地址!');
			  $("#Address").focus();
			  return false;
			}
			if ($("#RegisteredCapital").val()==''){
			  alert('请输入注册资金!');
			  $("#RegisteredCapital").focus();
			  return false;
			}
			if ($("#Business").val()==''){
			  alert('请输入经营范围!');
			  $("#Business").focus();
			  return false;
			}
			return true;
		  }
		 </script>
		 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
			<form method="post" id="myform" action="KS.UserRZ.asp" name="myform" onSubmit="return CheckForm();">
             <input type="hidden" name="action" value="yyrzsave" />   
             <input type="hidden" name="username" value="<%=rs("username")%>" />   
			<tr class="sort"> 
			  <td height="25" colspan="2" align="center"><font size="2"><strong>审核营业执照认证</strong></font></td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>公司名称：</strong></div></td>
			  <td height="30">  <input name="CompanyName"  id="CompanyName" type="text" size=30 value="<%=RS("CompanyName")%>">		      </td>
			</tr>
			 <tr class="tdbg">
                            <td class="clefttitle" align="right">注 册 号：</td>
                            <td><input name="BusinessLicense" class="textbox" type="text" id="BusinessLicense" value="<%=rs("BusinessLicense")%>" size="30" maxlength="50" /></td>
             </tr>
              <tr class="tdbg">
                            <td class="clefttitle" align="right">企业法人：</td>
                            <td><input name="LegalPeople" class="textbox" type="text" id="LegalPeople" value="<%=rs("LegalPeople")%>" size="30" maxlength="50" />
                            </td>
              </tr>
				 <tr class="tdbg">
                            <td class="clefttitle" align="right">公司地址：</td>
                            <td><input name="Address" class="textbox" type="text" id="Address" value="<%=rs("Address")%>" size="30" maxlength="50" /></td>
                </tr>
				 <tr class="tdbg">
                            <td class="clefttitle" align="right">注册资金：</td>
                            <td>
							<input type="text" name="RegisteredCapital" value="<%=rs("RegisteredCapital")%>" class="textbox"/>
							</td>
                   </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">经营范围：</td>
                            <td>
							<textarea name="Business" id="Business" class="textbox" style="width:300px;height:60px"><%=rs("Business")%></textarea></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">成立日期：</td>
                            <td>
							<input type="text" name="Foundation" id="Foundation" value="<%=rs("Foundation")%>" class="textbox"/>
							</td>
                          </tr>
						 <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照复印件：</td>
						  <td><input type="text" class="textbox" size="30" name="photourl" value="<%=rs("photourl")%>"/>
						  </td>
						</tr>
						<%if rs("photourl")<>"" then%>
                       <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照预览：</td>
						  <td><a href="<%=rs("photourl")%>" target="_blank"><img src="<%=rs("photourl")%>" width="200" border="0" style="border:1px solid #cccccc;padding:1px"/></a>
						  </td>
						</tr>
					<%end if%>
					<tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">审核标志：</td>
						  <td>
						   <input type="radio" value="1" name="isrz"<%if rs("isrz")="1" then response.write " checked"%> />审核通过<br/>
						   <input type="radio" value="2" name="isrz"<%if rs("isrz")="2" then response.write " checked"%> />已提交未审核<br/>
						   <input type="radio" value="3" name="isrz"<%if rs("isrz")="3" then response.write " checked"%> />审核不通过，退回要求重填<br/>
						  </td>
					</tr>
					<tr class="tdbg">
						  <td height="60"></td>
						  <td>
						   <input type="submit" value="确定保存" class="button"/>
						   <input type="button" onclick="history.back(-1);" value="取消返回" class="button"/>
						  </td>
					</tr>
		 </form>
		 </table>
		 <%
		end sub
		
		sub yyrzsave()
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 1 * from ks_enterprise where username='" & ks.g("username") & "'",conn,1,3
		  if rs.eof and rs.bof then
		   rs.close
		   set rs=nothing
		   ks.die "<script>alert('出错啦!');history.back();</script>"
		  end if
		  rs("companyname")=ks.g("companyname")
		  rs("BusinessLicense")=ks.g("BusinessLicense")
		  rs("LegalPeople")=ks.g("LegalPeople")
		  rs("Address")=ks.g("Address")
		  rs("RegisteredCapital")=ks.g("RegisteredCapital")
		  rs("Business")=ks.g("Business")
		  rs("Foundation")=ks.g("Foundation")
		  rs("isrz")=ks.chkclng(ks.g("isrz"))
		  rs("photourl")=ks.g("photourl")
		  rs("rzsj")=now
		  rs.update
		 rs.close
		 set rs=nothing
		 conn.execute("update ks_user set isrz=" & ks.chkclng(ks.g("isrz")) & " where username='" & ks.g("username") & "'")
		 ks.die "<script>alert('恭喜，您的操作已保存!');location.href='KS.UserRZ.ASP';</script>"
		end sub
		
		'身份证认证审核
		sub sfzrz()
		  if ks.g("username")="" then
		    ks.die "<script>alert('对不起，用户名出错!');history.back();</script>"
		  end if
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_user where username='" & ks.g("username") & "'",conn,1,1
		 if rs.eof and rs.bof then
		   rs.close:set rs=nothing
		   ks.die "出错啦!"
		 end if
		 %>
		 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#RealName").val()==''){
			  alert('请输入姓名!');
			  $("#RealName").focus();
			  return false;
			}
			if ($("#IDCard").val()==''){
			  alert('请输入身份证号!');
			  $("#IDCard").focus();
			  return false;
			}
			return true;
		  }
		 </script>
		 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
			<form method="post" id="myform" action="KS.UserRZ.asp" name="myform" onSubmit="return CheckForm();">
             <input type="hidden" name="action" value="sfzrzsave" />   
             <input type="hidden" name="username" value="<%=rs("username")%>" />   
			<tr class="sort"> 
			  <td height="25" colspan="2" align="center"><font size="2"><strong>审核身份证认证</strong></font></td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>真实姓名：</strong></div></td>
			  <td height="30">  <input name="RealName"  id="RealName" type="text" size=30 value="<%=RS("RealName")%>">		      </td>
			</tr>
			 <tr class="tdbg">
                            <td class="clefttitle" align="right">身份证号码：</td>
                            <td><input name="IDCard" class="textbox" type="text" id="IDCard" value="<%=rs("IDCard")%>" size="30" maxlength="50" /></td>
             </tr>
             
						  
						 <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">身份证复印件：</td>
						  <td><input type="text" class="textbox" size="30" name="photourl" value="<%=rs("sfzphotourl")%>"/>
						  </td>
						</tr>
						<%if rs("sfzphotourl")<>"" then%>
                       <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照预览：</td>
						  <td><a href="<%=rs("sfzphotourl")%>" target="_blank"><img src="<%=rs("sfzphotourl")%>" width="200" border="0" style="border:1px solid #cccccc;padding:1px"/></a>
						  </td>
						</tr>
					<%end if%>
					<tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">审核标志：</td>
						  <td>
						   <input type="radio" value="1" name="issfzrz"<%if rs("issfzrz")="1" then response.write " checked"%> />审核通过<br/>
						   <input type="radio" value="2" name="issfzrz"<%if rs("issfzrz")="2" then response.write " checked"%> />已提交未审核<br/>
						   <input type="radio" value="3" name="issfzrz"<%if rs("issfzrz")="3" then response.write " checked"%> />审核不通过，退回要求重填<br/>
						  </td>
					</tr>
					<tr class="tdbg">
						  <td height="60"></td>
						  <td>
						   <input type="submit" value="确定保存" class="button"/>
						   <input type="button" onclick="history.back(-1);" value="取消返回" class="button"/>
						  </td>
					</tr>
		 </form>
		 </table>
		 <%
		end sub
		
		sub sfzrzsave()
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 1 * from ks_user where username='" & ks.g("username") & "'",conn,1,3
		  if rs.eof and rs.bof then
		   rs.close
		   set rs=nothing
		   ks.die "<script>alert('出错啦!');history.back();</script>"
		  end if
		  rs("realname")=ks.g("realname")
		  rs("idcard")=ks.g("idcard")
		  rs("issfzrz")=ks.chkclng(ks.g("issfzrz"))
		  rs("sfzphotourl")=ks.g("photourl")
		  rs.update
		 rs.close
		 set rs=nothing
		 ks.die "<script>alert('恭喜，您的操作已保存!');location.href='KS.UserRZ.ASP';</script>"
		end sub
		
		Sub mobilerz()
		 if ks.g("username")="" then
		    ks.die "<script>alert('对不起，用户名出错!');history.back();</script>"
		  end if
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_user where username='" & ks.g("username") & "'",conn,1,1
		 if rs.eof and rs.bof then
		   rs.close:set rs=nothing
		   ks.die "出错啦!"
		 end if
		 %>
		 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#Mobile").val()==''){
			  alert('请输入手机号码!');
			  $("#Mobile").focus();
			  return false;
			}
			return true;
		  }
		 </script>
		 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
			<form method="post" id="myform" action="KS.UserRZ.asp" name="myform" onSubmit="return CheckForm();">
             <input type="hidden" name="action" value="mobilerzsave" />   
             <input type="hidden" name="username" value="<%=rs("username")%>" />   
			<tr class="sort"> 
			  <td height="25" colspan="2" align="center"><font size="2"><strong>手机认证</strong></font></td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>客户姓名：</strong></div></td>
			  <td height="30">  <%=rs("realname")%> </td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>手机号码：</strong></div></td>
			  <td height="30">  <input name="Mobile"  id="Mobile" type="text" size=30 value="<%=RS("Mobile")%>">		      </td>
			</tr>
			 
					<tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">审核标志：</td>
						  <td>
						   <input type="radio" value="1" name="ismobilerz"<%if rs("ismobilerz")="1" then response.write " checked"%> />审核通过<br/>
						   <input type="radio" value="2" name="ismobilerz"<%if rs("ismobilerz")="2" then response.write " checked"%> />已提交未审核<br/>
						   <input type="radio" value="3" name="ismobilerz"<%if rs("ismobilerz")="3" then response.write " checked"%> />审核不通过，退回要求重填<br/>
						  </td>
					</tr>
					<tr class="tdbg">
						  <td height="60"></td>
						  <td>
						   <input type="submit" value="确定保存" class="button"/>
						   <input type="button" onclick="history.back(-1);" value="取消返回" class="button"/>
						  </td>
					</tr>
		 </form>
		 </table>
		 <br/>
		 <strong><font color=red>说明：目前采用的是管理员人工认证方式，请手工发一条消息让您的客户确定。</font></strong>
		 <%
		End Sub
		
		Sub mobilerzsave()
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 1 * from ks_user where username='" & ks.g("username") & "'",conn,1,3
		  if rs.eof and rs.bof then
		   rs.close
		   set rs=nothing
		   ks.die "<script>alert('出错啦!');history.back();</script>"
		  end if
		  rs("mobile")=ks.g("mobile")
		  rs("ismobilerz")=ks.chkclng(ks.g("ismobilerz"))
		  rs.update
		 rs.close
		 set rs=nothing
		 ks.die "<script>alert('恭喜，您的操作已保存!');location.href='KS.UserRZ.ASP';</script>"
		End Sub
		
		sub emailrz()
           if ks.g("username")="" then
		    ks.die "<script>alert('对不起，用户名出错!');history.back();</script>"
		  end if
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_user where username='" & ks.g("username") & "'",conn,1,1
		 if rs.eof and rs.bof then
		   rs.close:set rs=nothing
		   ks.die "出错啦!"
		 end if
		 %>
		 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#Email").val()==''){
			  alert('请输入手机号码!');
			  $("#Email").focus();
			  return false;
			}
			return true;
		  }
		 </script>
		 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
			<form method="post" id="myform" action="KS.UserRZ.asp" name="myform" onSubmit="return CheckForm();">
             <input type="hidden" name="action" value="emailrzsave" />   
             <input type="hidden" name="username" value="<%=rs("username")%>" />   
			<tr class="sort"> 
			  <td height="25" colspan="2" align="center"><font size="2"><strong>邮箱认证</strong></font></td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>客户姓名：</strong></div></td>
			  <td height="30">  <%=rs("realname")%> </td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>电子邮箱：</strong></div></td>
			  <td height="30">  <input name="Email"  id="Email" type="text" size=30 value="<%=RS("Email")%>">		      </td>
			</tr>
			 
					<tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">审核标志：</td>
						  <td>
						   <input type="radio" value="1" name="isemailrz"<%if rs("isemailrz")="1" then response.write " checked"%> />审核通过<br/>
						   <input type="radio" value="2" name="isemailrz"<%if rs("isemailrz")="2" then response.write " checked"%> />已提交未审核<br/>
						  </td>
					</tr>
					<tr class="tdbg">
						  <td height="60"></td>
						  <td>
						   <input type="submit" value="确定保存" class="button"/>
						   <input type="button" onclick="history.back(-1);" value="取消返回" class="button"/>
						  </td>
					</tr>
		 </form>
		 </table>
		 <br/>
		 <strong><font color=red>说明：邮箱认证用户可以自助完成，一般不需要后台管理员手工审核。</font></strong>
         </div>
		 <%
		end sub
		
		sub emailrzsave()
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 1 * from ks_user where username='" & ks.g("username") & "'",conn,1,3
		  if rs.eof and rs.bof then
		   rs.close
		   set rs=nothing
		   ks.die "<script>alert('出错啦!');history.back();</script>"
		  end if
		  rs("email")=ks.g("email")
		  rs("isemailrz")=ks.chkclng(ks.g("isemailrz"))
		  rs.update
		 rs.close
		 set rs=nothing
		 ks.die "<script>alert('恭喜，您的操作已保存!');location.href='KS.UserRZ.ASP';</script>"
		end sub
		
		sub delrz()
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_user where userid=" & KS.ChkClng(KS.G("UserID")),conn,1,3
		 If rs.eof and rs.bof then
		   rs.close
		   set rs=nothing
		   ks.die "<script>alert('出错啦!');history.back();</script>"
		 end if
		 rs("isrz")=0
		 rs("IsSFZRZ")=0
		 rs("IsMobileRZ")=0
		 rs("IsEmailRZ")=0
		 rs("SFZPhotoUrl")=""
		 rs.update
		 rs.close
		 set rs=nothing
		 KS.Die "<script>alert('恭喜，删除该用户认证信息成功!');location.href='KS.UserRZ.ASP';</script>"
		end sub
		
End Class
		%>
 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../plus/md5.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
'response.buffer=false
Dim KSCls
Set KSCls = New Admin_UserMessage
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserMessage
        Private KS,KSR
		Private Action,RSObj,MaxPerPage,TotalPut
		Private Title, Message, Numc

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR=New Refresh
		   MaxPerPage = 20
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSR=Nothing
		End Sub
		Public Sub Kesion()
		    '删除指定日期的消息
			 Conn.Execute("Delete From KS_Message Where AutoDelDays>0 And DateDiff(" & DataPart_D &",sendtime," & SQLNowString &")>=AutoDelDays")
	        Response.Write "<!DOCTYPE html><html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script src=""../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			Response.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			Response.Write EchoUeditorHead()
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			
            If Not KS.ReturnPowerResult(0, "KMUA10003") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If	
					
			Response.Write"<div class=""topdashed menu_top_fixed quickLink"" style=""text-align:left"">用户短信管理：<a href='?action=new'>发送站内短信</a> <a href='?action=sms'>发送手机短信</a></div>"
			Response.Write "<div class=""menu_top_fixed_height""></div>"
		Action=Trim(Request("Action"))
		Select Case Action
		Case "new","edit"
		    call SendMsg()
		Case "add"
			call savemsg()
	    Case "sms"
		    Call SendSMS() 
		Case "saveSendSms"
		    Call saveSendSms()
		Case "saveedit"
		   If Not KS.ReturnPowerResult(0, "KMUA1000312") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
		    Call editsavemsg()
		Case "delall"
		     If Not KS.ReturnPowerResult(0, "KMUA1000311") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
			call delall()
		Case "delchk"
			call delchk()
		Case "del"
		    If Not KS.ReturnPowerResult(0, "KMUA1000311") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
		    call delbyid()
		Case else
			call main()
		end Select
		Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS X" & GetVersion &", Copyright (c) 2006-" & year(now)&" <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"%>
		</body>
		</html>
		<%
		End Sub
		
		Sub Main()

         %>
        <div class="pageCont2">  
        <div class="tabTitle">用户短信管理</div>
		 <form name="myform" method="Post" action="KS.UserMessage.asp">
		<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
				  <tr class='sort'>
					<td height="22" width="50" align="center">选中</td>
					<td align="center">标题</td>
					<td align="center">发送者</td>
					<td align="center" width="80">接收者</td>
					<td align="center">发送时间</td>
					<td width="40" align="center">状态</td>
					<td width="120" align="center">操作</td>
				  </tr>
			<%
		           Set RSObj = Server.CreateObject("ADODB.RecordSet")
				   Dim Param:Param=" where 1=1"
				   If KS.S("KeyWord")<>"" Then
				     select case KS.ChkClng(KS.S("condition"))
					   case 1
					    Param=Param & " and title like '%" & KS.S("KeyWord") & "%'"
					   case 2
					    Param=Param & " and Sender like '%" & KS.S("KeyWord") & "%'"
					   case 3
					    Param=Param & " and Incept like '%" & KS.S("KeyWord") & "%'"
					 end select 
				   End If
				   RSObj.Open "SELECT * FROM KS_Message " & Param & " order by id Desc", Conn, 1, 1
				 If RSObj.EOF Then
				    Response.Write "<tr><td colspan=8 height='30' align='center'>找不到任何短消息！</td></tr>"
				 Else
					totalPut = RSObj.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
			End If
				 %>	
		<tr class='list' onMouseOver="this.className='list'" onMouseOut="this.className='list'">
			<td colspan=8 height="30" class="operatingBox">
			<input type="hidden" value="del" name="action">
			<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">选中本页显示的所有记录&nbsp;<input type="submit" value="删除选中的记录" onclick="return(confirm('确定删除选中的记录吗？'))" class="button">
			&nbsp
			<input type="button" value="发送短信" onclick="location.href='?action=new';" class="button">
					 </td>
		  </tr> 
		  <%
		  Response.Write "<tr><td colspan='7' align='right'>"
		  			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)

			Response.Write "</td></tr>"

		  %> 
		</table>
		</form>
		<div>
		<form action="KS.UserMessage.asp" name="myform" method="post">
		   <div>
			  &nbsp;<strong>快速搜索=></strong>
			 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;条件:
			 <select name="condition" class="textbox">
			  <option value=1>短信标题</option>
			  <option value=2>发送用户</option>
			  <option value=3>接收用户</option>
			 </select>
			  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
			  </div>
		</form>
		</div>
		
		<table width="100%" border="0" align=center cellpadding="3" cellspacing="1" class="ctable">
		  <tr align="center" class="sort"> 
			<td height="25" colspan="2">短消息管理(批量删除)</td>
		  </tr>
		  <form action="KS.UserMessage.asp?action=del" method=post>
		  </form>
		  <form action="KS.UserMessage.asp?action=delall" method=post>
			<tr> 
			  <td colspan="2" bgcolor="#FFFFFF" class=tdbg> 批量删除用户指定日期内短消息（默认为删除已读信息）：<br>
				<select name="delDate" size=1>
				  <option value=7>一个星期前</option>
				  <option value=30>一个月前</option>
				  <option value=60>两个月前</option>
				  <option value=180>半年前</option>
				  <option value="all">所有信息</option>
				</select>
				&nbsp; 
				<input type="checkbox" name="isread" value="yes">
				包括未读信息 
				<input type="submit" name="Submit" class="button" value="提 交">
			  </td>
			</tr>
		  </form>
		  <form action="KS.UserMessage.asp?action=delchk" method=post>
			<tr> 
			  <td colspan="2" bgcolor="#FFFFFF" class=tdbg> 批量删除含有某关键字短信（注意：本操作将删除所有已读和未读信息）：<br>
				关键字： 
				<input class="textbox" type="text" name="keyword" size=30>
				&nbsp;在 
				<select name="selaction" size=1>
				  <option value=1>标题中</option>
				  <option value=2>内容中</option>
				</select>
				&nbsp; 
				<input type="submit" name="Submit" value="提 交" class='button'>
			  </td>
			</tr>
		  </form>
		</table>
		<%
		End Sub
		
		Sub ShowContent()
		 Dim i:i=1
		 Do While Not RSObj.Eof
		 %>
		  <tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" id='u<%=RSObj("ID")%>' onclick="chk_iddiv('<%=RSObj("ID")%>')">
		    <td class="splittd" style="text-align:center"><input name="id" type="checkbox" onClick="unselectall()" id='c<%=RSObj("ID")%>'value="<%=RSObj("ID")%>"></td>
			<td class="splittd"><img src="../images/bg30.png" align="absmiddle"><a href="?action=edit&id=<%=rsobj("id")%>"><%=KS.Gottopic(rsobj("title"),35)%></a></td>
			<td class="splittd" style="text-align:center"><%=rsobj("sender")%></td>
			<td class="splittd" style="text-align:center"><%=rsobj("Incept")%></td>
			<td class="splittd" style="text-align:center"><%=rsobj("sendtime")%></td>
			<td class="splittd" align="center">
			<%if rsobj("flag")=0 then
			   response.write "<font color=red>未读</font>"
			  else
			   response.write "<font color=blue>已读</font>"
			  end if
			 %>
			</td>
			<td class="splittd" align="center"><a href="?action=edit&id=<%=rsobj("id")%>" class="setA">修改</a>|<a onclick="return(confirm('删除后不可恢复，确定删除吗?'))" href="?action=del&id=<%=rsobj("id")%>"  class="setA">删除</a></td>
		  </tr>
		 <%if i>=maxperpage then exit do
		   i=I+1
		  RSObj.MoveNext
		 Loop
		End Sub
		
		'站内消息
		Sub SendMsg()
		  dim flag,display,Incept,title,content,sendtime,AutoDelDays
		  If KS.S("Action")="edit" then
		    flag="saveedit"
			display=" style='display:none'"
			dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select top 1 * from ks_message where id="& ks.chkclng(ks.s("id")),conn,1,1
			if not rs.eof then
			 Incept=rs("Incept")
			 title=rs("title")
			 content=rs("content")
			 sendtime=rs("sendtime")
			 AutoDelDays=rs("AutoDelDays")
			end if
			rs.close:set rs=nothing
		  else
		    flag="add"
			 AutoDelDays=0
			 dim userid:userid=KS.FilterIds(replace(request("userid")," ",""))
			 dim usernamelist
			 if userid<>"" then
				 set rs=KS.InitialObject("adodb.recordset")
				 rs.open "select userid,username from ks_user where userid in("& userid & ")",conn,1,1
				 do while not rs.eof
				  if usernamelist="" then
				   usernamelist=rs(1)
				  else
				   usernamelist=usernamelist &"," & rs(1)
				  end if
				  rs.movenext
				 loop
				 rs.close:set rs=nothing
			 end if
		  end if
		  
		  
		if action="edit" and not ks.isnul(content) then
		  %>
		  
		  <div id="showdiv" class="pageCont pt10" style="border-radius:5px 5px 0 0">
			  <style>
				  .annex{width:600px;margin :15px; border : 1px dashed #999; background : #f9f9f9; line-height : normal;}
				  .annex td{padding-top:10px;padding-left:10px;padding-bottom:5px;}
			 </style>
			<table width="100%" border="0" align=center cellpadding="3" cellspacing="1" class="ctable">
				<tr class="sort">
				  <td height="25" colspan="2" align="center">查看短消息内容 
				  <input type="button" class="button" value="修改" onclick="$('#showdiv').hide();$('#modifydiv').show();$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ChannelID=1&OpStr=<%=Server.URLEncode("用户系统 >> 修改短消息内容")%>&ButtonSymbol=GoSave';"/>
				  </td>
				</tr>
				<tr>
				 <td>
				  <%
				  response.write KS.ClearBadChr(KS.HtmlCode(content))
				  %>
			   </td>
			  </tr>
			 </table>
		 </div>
		 
		  <%
		end if
		 %>
	<div class="pageCont2">
		<div id="modifydiv"<%if action="edit" then response.write " style='display:none'"%>>
		  <form action="KS.UserMessage.asp?action=<%=flag%>" method=post name="myform" id="myform">
		   <input type="hidden" value="<%=KS.S("id")%>" name="id">
		<table width="100%" border="0" style="margin-top:3px" align=center cellpadding="3" cellspacing="1" class="ctable">
			<tr class="sort">
			  <td height="25" colspan="2" align="center"><%if action="edit" then response.write "修改" else response.write "发送"%>短消息</td>
		    </tr>
			<tr class="tdbg"<%=display%>>
				<td height="25" align="right" class="clefttitle" width="179">用户类别：</td>
				<td>
				<Input type="radio" name="UserType" value="1" checked onclick="UType(this.value)">用户名单
				<Input type="radio" name="UserType" value="2" onclick="UType(this.value)">用户组
                <%If KS.C("SuperTF")="1" Then%>
				<Input type="radio" name="UserType" value="0" onclick="UType(this.value)">所有用户
                <%end if%>
                </td>
			</tr>
			<%if ks.s("action")="edit" then%>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">接收用户：</td>
				<td> 
				<%=Incept%>
				</td>
			</tr>
            <tr class="tdbg">
			   <td height="25" align="right" class="clefttitle">发送时间：</td>
			   <td><input type="text" name="sendtime" class="textbox" value="<%=sendtime%>"> <font color=red>格式：0000-00-00 00:00</font></td>
			</tr>
			<%else%>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">用 户 名：</td>
				<td> <INPUT class="textbox" TYPE="text" value="<%=usernamelist%>" NAME="UserName" size="80"><br>
				请输入用户名：(多个用户名请以英文逗号“,”分隔,注意区分大小写)</td>
			</tr>
			<%end if%>
			<tr class="tdbg" id="ToGroupID" style="display:none;">
				<td height="25" align="right" class="clefttitle">用 户 组：</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
					<tr><td>
					<%
					IF KS.C("SuperTF")=1 Then
					 response.write KS.GetUserGroup_CheckBox("GroupID",0,4)
					Else
					    Dim RSG:Set RSG=Conn.Execute("select top 1 AllowGroupID From KS_UserGroup Where ID=" & KS.ChkClng(KS.C("Groupid")))
						If Not RSG.Eof Then
						       dim AllowGroupID:AllowGroupID=RSG("AllowGroupID")
							   dim rowNum:RowNuM=3
							   dim OptionName:OptionName="GroupID"
								Dim n:n=0
							   IF RowNum<=0 Then RowNum=3
							   
							   
								KS.LoadUserGroup()
								Dim Node,str,i,DocNode
								Set DocNode=Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
								
								
								str="<select name=""" & OptionName &""" multiple style=""width:380px;height:200px"">"
								For I=0 To DocNode.Length-1
									Set Node=DocNode.Item(i)
									Dim TJ,SpaceStr,k
									SpaceStr=""
									TJ=Node.SelectSingleNode("@tj").text
									For k = 1 To TJ - 1
									 SpaceStr = SpaceStr & "──"
									Next
									Dim Nstr
									 If TJ=1 Then
									  Nstr="+ " & Node.SelectSingleNode("@groupname").text &""
									 Else
									  Nstr=Node.SelectSingleNode("@groupname").text 
									 End If
									 NStr=SpaceStr & Nstr
									
									 If KS.FoundInArr(AllowGroupID,Node.SelectSingleNode("@id").text,",") Then
									 str=str & "<option style='color:green' value='" &Node.SelectSingleNode("@id").text &"'>" & Nstr & "</option>"&vbcrlf
									 Else
									 str=str & "<OPTGROUP label=""" & Nstr&"""> </OPTGROUP>"&vbcrlf
									 End If
									
								Next
								Str=Str &"</select><div style='color:green'>您只能选择，绿色的用户组，按Ctrl或是Shift可以多选。</div>"
							
								RESPONSE.Write str
						End If
						RSG.Close
						Set RSG=Nothing
					End If
					%>
					</td></tr>
					<tr><td height=20><input type="button" value="打开高级设置" class="button" name="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
					<tr><td height=20 ID="UpSetting" style="display:NONE">
						<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
						<tr><td height=20 colspan="4">符合条件设置(以下条件将对选择的用户组生效)</td></tr>
						<tr>
							<td width="15%">最后登陆时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="LoginTime" onkeyup="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="LoginTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginTimeType" value="1">少于							</td>
							<td width="15%">注册时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="RegTime" onkeyup="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="RegTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="RegTimeType" value="1">少于							</td>
						</tr>
						<tr>
							<td>登陆次数：</td>
							<td><input class="textbox" type="text" name="Logins" size=6 onkeyup="CheckNumber(this,'次数')">次 &nbsp;<INPUT TYPE="radio" NAME="LoginsType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginsType" value="1">少于							</td>
							<td>发表文章：</td>
							<td><input class="textbox" type="text" name="UserArticle" size=6 onkeyup="CheckNumber(this,'篇数')">篇 &nbsp;<INPUT TYPE="radio" NAME="UserArticleType" checked value="0">多于 <INPUT TYPE="radio" NAME="UserArticleType" value="1">少于</td>
						</tr></table>
					</td></tr></table>				</td>
			</tr>
			<tr class=tdbg> 
			  <td width="179" height="25" align="right" class="clefttitle">消息标题：</td>
			  <td width="1106"> 
			  <input class="textbox" type="text"  value="<%=title%>" name="title" size="80">			  </td>
			</tr>
			<tr class=tdbg> 
			  <td width="179" height="25" align="right" class="clefttitle">消息内容：</td>
			  <td width="1106"> 
			    <%
				 Response.Write "<script id=""message"" name=""message"" type=""text/plain"" style=""width:90%;height:220px;"">" &KS.ClearBadChr(content)&"</script>"
	             Response.Write "<script>setTimeout(""var editor = " & GetEditorTag() &".getEditor('message',{toolbars:[" & GetEditorToolBar("newstool") &"],wordCount:false,autoHeightEnabled:false});"",10);</script>"
				%>
				<%IF Action="new" Then%> 
	            <br/>可用标签： {$UserName}-用户名,{$RealName} -用户姓名, {$GetSiteName}-网站名称,{$GetSiteUrl} -网站URL
				<%End If%>
			  </td>
			</tr>
            <tr class=tdbg> 
			  <td width="179" height="25" align="right" class="clefttitle">超过：</td>
			  <td width="1106"> 
			  <input class="textbox" type="text"  value="<%=AutoDelDays%>" name="AutoDelDays" size="10" style="width:40px;text-align:center">天,自动删除该短消息。设置“0”不自动删除。			  </td>
			</tr>
			<tr class=tdbg> 
			  <td height="25" colspan="2" style="text-align:center"> 
				  <input type="button" name="Submit" value="<%if action="edit" then response.write "修改" else response.write "发送"%>消息" class='button' onclick="return(CheckForm())">
				  <input type="reset" name="Submit2" value="重新填写" class='button'>			  </td>
		    </tr>
		</table>
		  </form>
		</div>  
	 </div>
		<script>
		 function CheckForm()
		 {
		   if (document.myform.title.value==''){
			 top.$.dialog.alert('站内短信标题不能为空！',function(){
			 document.myform.title.focus();
			 });
			 return false;
		  }
		  if (editor.hasContents()==false)
			{
			  top.$.dialog.alert("站内短信内容不能为空！",function(){
			  editor.focus();
			  });
			  return false;
		   } 
           $("#myform").submit();
		 }
		</script>
		<br>
		
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function openset(v,s){
			if (v.value=='打开高级设置'){
				document.getElementById(s).style.display = "";
				v.value="关闭高级设置";
			}
			else{
				v.value="打开高级设置";
				document.getElementById(s).style.display = "none";
			}
		}
		function UType(n){
			if (n==1){
				document.getElementById("ToUserName").style.display = "";
				document.getElementById("ToGroupID").style.display = "none";
			}
			else if(n==2){
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "";
			}
			else{
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "none";
			}
		}
		//-->
		</SCRIPT>
		<%
		
		end sub
		
		sub editsavemsg()
		   dim id:id=ks.chkclng(ks.s("id"))
		   dim title:title=ks.g("title")
		   dim content:content=Request.Form("message")
		   dim sendtime:sendtime=ks.s("sendtime")
		   if not isdate(sendtime) then
		    Response.Write "<script>alert('时间格式不正确!');history.back();</script>"
			Exit Sub
			end if
			dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select  top 1 * from ks_message where id=" &id,conn,1,3
			if not rs.eof then
			  rs("title")=title
			  rs("content")=content
			  rs("sendtime")=sendtime
			  rs("AutoDelDays")=KS.ChkClng(Request("AutoDelDays"))
			  rs.update
			end if
			rs.close
			set rs=nothing
			response.write "<script>top.$.dialog.alert('恭喜，修改成功!',function(){ location.href='user/ks.usermessage.asp';}); </script>"
		   response.end
		end sub
		
		Sub delbyid()
		  If Ks.G("id")="" Then
				Response.Write("<script>alert('参数传递出错!');history.back();</script>")
				Exit Sub
			end if
		    Conn.Execute("delete from ks_message where id in(" & KS.FilterIDs(KS.G("id")) &")")
			Response.Write Response.Write("<script>alert('恭喜，删除操作成功！');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
		End Sub
		
		Sub del()
			if KS.G("username")="" then
				Response.Write("<script>alert('请输入要批量删除的用户名!');history.back();</script>")
				Exit Sub
			end if
			sql="delete from KS_Message where sender='"&KS.G("username")&"'"

			Conn.Execute(sql)
			
			Response.Write Response.Write("<script>alert('操作成功！请继续别的操作!');</script>")
		End Sub
		
		sub delall()
			dim selflag,sql
			if request("isread")="yes" then
			selflag=""
			else
			selflag=" and flag=1"
			end if
				select case request("delDate")
				case "all"
				sql="delete from KS_Message where id>0 "&selflag
				case 7
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>7 "&selflag
				case 30
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>30 "&selflag
				case 60
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>60 "&selflag
				case 180
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>180 "&selflag
				end select
				Conn.Execute(sql)
                Call KS.Die("<script>top.$.dialog.alert('操作成功！请继续别的操作。',function(){location.href='" & KS.Setting(3) & KS.Setting(89) &"user/KS.UserMessage.asp';});</script>")
		end Sub
		
		Sub delchk()
			if request.form("keyword")="" then
				KS.ShowError("请输入关键字！")
				Exit sub
			end if
			if request.form("selaction")=1 then
					conn.Execute("delete from KS_Message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
			elseif request.form("selaction")=2 then
				
					conn.Execute("delete from KS_Message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
			else
				KS.ShowError("未指定相关参数！")
			end if
                Call KS.Die("<script>top.$.dialog.alert('操作成功！请继续别的操作。',function(){location.href='" & KS.Setting(3) & KS.Setting(89) &"user/KS.UserMessage.asp';});</script>")
		End Sub
		
		Sub SaveMsg()
			Server.ScriptTimeout=99999
			Dim UserType
			UserType = Trim(Request.Form("UserType"))
			Title	 = Trim(Request.Form("title"))
			Message  = Request.Form("message")
			If Title="" or Message="" Then
				KS.Showerror("请填写消息的标题和内容!")
				Exit Sub
			End If
			If Len(Message) > KS.Setting(48) Then
				KS.Showerror("消息内容不能多于" & KS.Setting(48) & "字节")
				Exit Sub
			End If 
 
			Select Case UserType
			Case "0" : SaveMsg_0()	'按所有用户
			Case "1" : SaveMsg_1()	'按指定用户
			Case "2" : SaveMsg_2()	'按指定用户组
			Case Else
				KS.Showerror("请输入收信的用户!") : Exit Sub
			End Select
		  Call KS.Die("<script>top.$.dialog.alert('操作成功！本次发送"&Numc+1&"个用户。请继续别的操作。',function(){location.href='" & KS.Setting(3) & KS.Setting(89) &"user/KS.UserMessage.asp';});</script>")

		End Sub
		
		Function ReplaceLabel(ByVal Msg,UserName,RealName)
		   Msg=Replace(Msg,"'","''")
		   Msg=Replace(Msg,"{$UserName}",UserName)
		   If RealName="" Then RealName=UserName
		   Msg=Replace(Msg,"{$RealName}",realname)
		   Msg=Replace(Msg,"{$GetSiteName}",KS.Setting(0))
		   Msg=Replace(Msg,"{$GetSiteUrl}",KS.GetDomain)
		   Msg=Replace(Msg,"{$SendDate}",Now)
		   ReplaceLabel=Msg
		End Function
		
		'按所有用户发送
		Sub SaveMsg_0()
			Dim Rs,Sql,i
			Sql = "Select UserName,RealName From KS_User Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				Numc= Ubound(SQL,2)
				For i=0 To Numc
				   if cbool(KS.SendInfo(SQL(0,i),KS.C("AdminName"),Replace(Title,"'","''"),ReplaceLabel(Message,SQL(0,i),SQL(1,i))))=false then
				    KS.Die "<script>alert('用户" & SQL(0,I) & "不存在或是该用户邮箱已满，信件发送失败！');history.back();</script>"
				   end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'按指定用户
		Sub SaveMsg_1()
			Dim ToUserName,Rs,Sql,i
			ToUserName = Trim(Request.Form("UserName"))
			If ToUserName = "" Then
				KS.Showerror("请填写目标用户名，注意区分大小写。")
				Exit Sub
			End If
			ToUserName = Replace(ToUserName,"'","")
			ToUserName = Split(ToUserName,",")
			Numc= Ubound(ToUserName)
			For i=0 To Numc
				if cbool(KS.SendInfo(ToUserName(i),KS.C("AdminName"),Title,ReplaceLabel(Message,ToUserName(i),ToUserName(i))))=false then
				  KS.Die "<script>alert('用户" & ToUserName(i) & "不存在或是该用户邮箱已满，信件发送失败！');history.back();</script>"
				end if
			Next
		End Sub
		'按指定用户组及条件发送
		Sub SaveMsg_2()
			Dim GroupID,ErrMsg,i
			Dim SearchStr,TempValue,DayStr
			GroupID = Replace(Request.Form("GroupID"),chr(32),"")
			If GroupID="" Then
			    ErrMsg = "请正确选取相应的用户组。"
			ElseIf GroupID<>"" and Not Isnumeric(Replace(GroupID,",","")) Then
				ErrMsg = "请正确选取相应的用户组。"
			Else
				GroupID = KS.R(GroupID)
			End If
			DayStr = "'d'"
			If Instr(GroupID,",")>0 Then
				SearchStr = "GroupID in ("&GroupID&")"
			Else
				SearchStr = "GroupID = "&KS.R(GroupID)
			End If
			'登陆次数
			TempValue = Request.Form("Logins")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginsType"),"LoginTimes")
			End If
			'发表文章
			TempValue = Request.Form("UserArticle")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserArticleType"),"(select count(id) from ks_iteminfo where inputer=ks_user.username)")
			End If
			'最后登陆时间
			TempValue = Request.Form("LoginTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginTimeType"),"Datediff("&DayStr&",LastLoginTime,"&SqlNowString&")")
			End If
			'注册时间
			TempValue = Request.Form("RegTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("RegTimeType"),"Datediff("&DayStr&",JoinDate,"&SqlNowString&")")
			End If
			If SearchStr="" Then
				ErrMsg = "请填写发送的条件选项。"
			End If
			If ErrMsg<>"" Then KS.Showerror(ErrMsg) : Exit Sub
			Dim Rs,Sql
			Sql = "Select UserName,RealName From KS_User Where "& SearchStr & " Order By UserID Desc"
			
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				Numc= Ubound(SQL,2)
				For i=0 To Numc
				    IF Cbool(KS.SendInfo(SQL(0,i),KS.C("AdminName"),Replace(Title,"'","''"),ReplaceLabel(Message,SQL(0,i),SQL(1,i))))=false then
					 KS.Die "<script>alert('用户" & SQL(0,I) & "不存在或是该用户邮箱已满，信件发送失败！');history.back();</script>"
					end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		
		Function GetSearchString(Get_Value,Get_SearchStr,UpType,UpColumn)
			Get_Value = Clng(Get_Value)
			If Get_SearchStr<>"" Then Get_SearchStr = Get_SearchStr & " and " 
			If UpType="1" Then
				Get_SearchStr = Get_SearchStr & UpColumn &" <= "&Get_Value
			Else
				Get_SearchStr = Get_SearchStr & UpColumn &" >= "&Get_Value
			End If
			GetSearchString = Get_SearchStr
		End Function
		
		
		'手机短信
		Sub SendSMS()
		  dim flag,display,Incept,title,content,sendtime,AutoDelDays

			 dim userid:userid=KS.FilterIds(replace(request("userid")," ",""))
			 dim usernamelist,mobileList
			 if userid<>"" then
				 dim rs:set rs=KS.InitialObject("adodb.recordset")
				 rs.open "select userid,username,mobile from ks_user where userid in("& userid & ")",conn,1,1
				 do while not rs.eof
				   if rs("mobile")<>"" then
				      if mobileList="" then
					    mobileList=rs("mobile")
					  else
					    mobileList=mobileList &"," & rs("mobile")
					  end if
					  if usernamelist="" then
					   usernamelist=rs(1)
					  else
					   usernamelist=usernamelist &"," & rs(1)
					  end if
				  End If
				  rs.movenext
				 loop
				 rs.close:set rs=nothing
			 end if
		  
		%>
		<script>
			function dogetbalance() {
               jQuery("#mybalance").html("<img src='../images/loading.gif' />查询中...");
               jQuery.get("../system/KS.Setting.asp", { action: "balance",rnd:Math.random()}, function(val) {
                   jQuery("#mybalance").html("余额："+val+"条");
               });
           }
			</script>
	<div class="pageCont2">  
		  <form action="KS.UserMessage.asp?action=saveSendSms" method=post name="myform" id="myform">
	
		<table width="100%" border="0" style="margin-top:3px" align=center cellpadding="3" cellspacing="1" class="ctable">
			<tr class="sort">
			  <td height="25" colspan="2" align="center">发送手机短信</td>
		    </tr>
			<tr class="tdbg">
				<td height="25" align="right" class="clefttitle" width="179">账户余额：</td>
				<td><input type="button" class="button" value="查询短信余额" onclick="dogetbalance();"/>
				  <span id="mybalance"></span>
                </td>
			</tr>
			<tr class="tdbg">
				<td height="25" align="right" class="clefttitle" width="179">用户类别：</td>
				<td>
				<Input type="radio" name="UserType" value="1" checked onclick="UType(this.value)">手机名单
				<Input type="radio" name="UserType" value="2" onclick="UType(this.value)">用户组
                <%If KS.C("SuperTF")="1" Then%>
				<Input type="radio" name="UserType" value="0" onclick="UType(this.value)">所有用户
                <%end if%>
                </td>
			</tr>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">手机号码：</td>
				<td> 
				<INPUT class="textbox" TYPE="text" value="<%=KS.FilterRepeatInArray(mobileList,",")%>" NAME="Mobile" size="80"><br>
				<span class="tips">请输入手机号码：(多个手机号请以英文逗号“,”分隔,注意区分大小写)。</span></td>
			</tr>
			<tr class="tdbg" id="ToGroupID" style="display:none;">
				<td height="25" align="right" class="clefttitle">用 户 组：</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
					<tr><td>
					<%
					IF KS.C("SuperTF")=1 Then
					 response.write KS.GetUserGroup_CheckBox("GroupID",0,4)
					Else
					    Dim RSG:Set RSG=Conn.Execute("select top 1 AllowGroupID From KS_UserGroup Where ID=" & KS.ChkClng(KS.C("Groupid")))
						If Not RSG.Eof Then
						       dim AllowGroupID:AllowGroupID=RSG("AllowGroupID")
							   dim rowNum:RowNuM=3
							   dim OptionName:OptionName="GroupID"
								Dim n:n=0
							   IF RowNum<=0 Then RowNum=3
							   
							   
								KS.LoadUserGroup()
								Dim Node,str,i,DocNode
								Set DocNode=Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
								
								
								str="<select name=""" & OptionName &""" multiple style=""width:380px;height:200px"">"
								For I=0 To DocNode.Length-1
									Set Node=DocNode.Item(i)
									Dim TJ,SpaceStr,k
									SpaceStr=""
									TJ=Node.SelectSingleNode("@tj").text
									For k = 1 To TJ - 1
									 SpaceStr = SpaceStr & "──"
									Next
									Dim Nstr
									 If TJ=1 Then
									  Nstr="+ " & Node.SelectSingleNode("@groupname").text &""
									 Else
									  Nstr=Node.SelectSingleNode("@groupname").text 
									 End If
									 NStr=SpaceStr & Nstr
									
									 If KS.FoundInArr(AllowGroupID,Node.SelectSingleNode("@id").text,",") Then
									 str=str & "<option style='color:green' value='" &Node.SelectSingleNode("@id").text &"'>" & Nstr & "</option>"&vbcrlf
									 Else
									 str=str & "<OPTGROUP label=""" & Nstr&"""> </OPTGROUP>"&vbcrlf
									 End If
									
								Next
								Str=Str &"</select><div style='color:green'>您只能选择，绿色的用户组，按Ctrl或是Shift可以多选。</div>"
							
								RESPONSE.Write str
						End If
						RSG.Close
						Set RSG=Nothing
					End If
					%>
					</td></tr>
					<tr><td height=20><input type="button" value="打开高级设置" class="button" name="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
					<tr><td height=20 ID="UpSetting" style="display:NONE">
						<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
						<tr><td height=20 colspan="4">符合条件设置(以下条件将对选择的用户组生效)</td></tr>
						<tr>
							<td width="15%">最后登陆时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="LoginTime" onkeyup="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="LoginTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginTimeType" value="1">少于							</td>
							<td width="15%">注册时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="RegTime" onkeyup="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="RegTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="RegTimeType" value="1">少于							</td>
						</tr>
						<tr>
							<td>登陆次数：</td>
							<td><input class="textbox" type="text" name="Logins" size=6 onkeyup="CheckNumber(this,'次数')">次 &nbsp;<INPUT TYPE="radio" NAME="LoginsType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginsType" value="1">少于							</td>
							<td>发表文章：</td>
							<td><input class="textbox" type="text" name="UserArticle" size=6 onkeyup="CheckNumber(this,'篇数')">篇 &nbsp;<INPUT TYPE="radio" NAME="UserArticleType" checked value="0">多于 <INPUT TYPE="radio" NAME="UserArticleType" value="1">少于</td>
						</tr></table>
					</td></tr></table>				</td>
			</tr>

			<tr class=tdbg> 
			  <td width="179" height="25" align="right" class="clefttitle">发送内容：</td>
			  <td width="1106"> 
			    <textarea name="content" id="content" class="texbox" style="width:550px;height:200px">您好{$UserName},祝您节日快乐！</textarea>
				<br/>
	            <br/>可用标签： {$UserName}-用户名,{$RealName} -用户姓名, {$GetSiteName}-网站名称,{$GetSiteUrl} -网站URL
			  </td>
			</tr>
            
			<tr class=tdbg> 
			  <td height="25"></td>
			  <td>
				  <input type="button" name="Submit" value="确定发送" class='button' onclick="return(CheckForm())">
				  <input type="reset" name="Submit2" value="重新填写" class='button'>			  </td>
		    </tr>
		</table>
		  </form>
		 </div> 
		<script>
		 function CheckForm()
		 {
		   if (document.myform.content.value==''){
			 top.$.dialog.alert('短信内容不能为空！',function(){
			 document.myform.content.focus();
			 });
			 return false;
		  }
		  
           $("#myform").submit();
		 }
		</script>
		<br>
		
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function openset(v,s){
			if (v.value=='打开高级设置'){
				document.getElementById(s).style.display = "";
				v.value="关闭高级设置";
			}
			else{
				v.value="打开高级设置";
				document.getElementById(s).style.display = "none";
			}
		}
		function UType(n){
			if (n==1){
				document.getElementById("ToUserName").style.display = "";
				document.getElementById("ToGroupID").style.display = "none";
			}
			else if(n==2){
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "";
			}
			else{
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "none";
			}
		}
		//-->
		</SCRIPT>
		<%
		
		end sub
		
		
		Sub SaveSendSms()
			Server.ScriptTimeout=99999
			Dim UserType
			UserType = Trim(Request.Form("UserType"))
			Message  = Request.Form("content")
			If  Message="" Then
				KS.Showerror("请填写要发送的短信内容!")
				Exit Sub
			End If

			Select Case UserType
			Case "0" : SaveSendSms_0()	'按所有用户
			Case "1" : SaveSendSms_1()	'按指定用户
			Case "2" : SaveSendSms_2()	'按指定用户组
			Case Else
				KS.Showerror("请输入收信的用户!") : Exit Sub
			End Select
		  Call KS.Die("<script>top.$.dialog.alert('操作成功！本次发送"&Numc+1&"个用户。请继续别的操作。',function(){location.href='" & KS.Setting(3) & KS.Setting(89) &"user/KS.UserMessage.asp';});</script>")
		End Sub
		
		'按所有用户发送
		Sub SaveSendSms_0()
			Call ToSend("Select UserName,RealName,Mobile From KS_User Order By UserID Desc")
		End Sub
		
		Sub ToSend(SQLStr)
		   response.write "<br/><div style=""padding:10px;background:#fff;margin:10px;border:1px solid #ccc"">"
		    Dim Rstr,MobileList
			Numc=0
		    Dim RS:Set RS=Conn.Execute(SQLStr)
			Do While Not RS.Eof 
			   If rs("Mobile")<> "" and Instr(MobileList,RS("Mobile"))=0 Then
			      Response.Write "<li>正在发送，手机号：" & rs("mobile") &""
				  
				 Rstr=KS.SendMobileMsg(rs("Mobile"),ReplaceLabel(Message,rs(0),rs(1)))
				 If Isnumeric(Rstr) and KS.ChkClng(Rstr)>0 Then
				   Numc=Numc+1
				   Response.Write "，<font color=green>成功</font>"
				 Else
				   Response.Write "，<font color=red>失败</font>"
				 End If
				 Response.Write "</li>"
				 Response.Flush()
				 MobileList=MobileList & "," & RS("Mobile")
				 
			   End If
			RS.MoveNext
			Loop
			RS.Close
			Set RS=Nothing
			Response.Write "<li><Br/><strong>所有发送完毕,成功发送<font color=red>" & Numc &"</font>条消息!</strong></li>"
			Response.Write "<li><Br/><input type=""button"" class=""button"" value=""返回"" onclick=""history.back();""/></li>"
            KS.Die "</div>"
		
		End Sub
		
		'按指定用户
		Sub SaveSendSms_1()
			Dim ToUserName,Sql,i,MobileList,Rstr,Numc
			ToUserName = Trim(Request.Form("Mobile"))
			If ToUserName = "" Then
				KS.Showerror("请填写要发送的手机号码。")
				Exit Sub
			End If
			Dim Arr:Arr=Split(ToUserName,",")
			Numc=0
			response.write "<br/><div style=""padding:10px;background:#fff;margin:10px;border:1px solid #ccc"">"
			For I=0 To Ubound(Arr)
			  If Arr(i)<> "" and Instr(MobileList,Arr(i))=0 Then
			     Response.Write "<li>正在发送，手机号：" &Arr(i) &""
				 
				 dim rs:set rs=conn.execute("select top 1 UserName,RealName,Mobile From KS_User Where Mobile='" & arr(i) &"'")
				 dim username,realname
				 if not rs.eof then
				    username=rs(0)
					realname=rs(1)
				 else
				    username=arr(i)
					realname=username
				 end if
				 rs.close
				 set rs=nothing
				  
				 Rstr=KS.SendMobileMsg(Arr(i),ReplaceLabel(Message,username,realname))
				 If Isnumeric(Rstr) and KS.ChkClng(Rstr)>0 Then
				   Numc=Numc+1
				   Response.Write "，<font color=green>成功</font>"
				 Else
				   Response.Write "，<font color=red>失败</font>"
				 End If
				 Response.Write "</li>"
				 Response.Flush()
				 MobileList=MobileList & "," & Arr(i)
				 
			   End If
			Next
			
			Response.Write "<li><Br/><strong>所有发送完毕,成功发送<font color=red>" & Numc &"</font>条消息!</strong></li>"
			Response.Write "<li><Br/><input type=""button"" class=""button"" value=""返回"" onclick=""history.back();""/></li>"
            KS.Die "</div>"
		End Sub
		'按指定用户组及条件发送
		Sub SaveSendSms_2()
			Dim GroupID,ErrMsg,i
			Dim SearchStr,TempValue,DayStr
			GroupID = Replace(Request.Form("GroupID"),chr(32),"")
			If GroupID="" Then
			    ErrMsg = "请正确选取相应的用户组。"
			ElseIf GroupID<>"" and Not Isnumeric(Replace(GroupID,",","")) Then
				ErrMsg = "请正确选取相应的用户组。"
			Else
				GroupID = KS.R(GroupID)
			End If
			DayStr = "'d'"
			If Instr(GroupID,",")>0 Then
				SearchStr = "GroupID in ("&GroupID&")"
			Else
				SearchStr = "GroupID = "&KS.R(GroupID)
			End If
			'登陆次数
			TempValue = Request.Form("Logins")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginsType"),"LoginTimes")
			End If
			'发表文章
			TempValue = Request.Form("UserArticle")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserArticleType"),"(select count(id) from ks_iteminfo where inputer=ks_user.username)")
			End If
			'最后登陆时间
			TempValue = Request.Form("LoginTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginTimeType"),"Datediff("&DayStr&",LastLoginTime,"&SqlNowString&")")
			End If
			'注册时间
			TempValue = Request.Form("RegTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("RegTimeType"),"Datediff("&DayStr&",JoinDate,"&SqlNowString&")")
			End If
			If SearchStr="" Then
				ErrMsg = "请填写发送的条件选项。"
			End If
			If ErrMsg<>"" Then KS.Showerror(ErrMsg) : Exit Sub
			Call ToSend("Select UserName,RealName,Mobile From KS_User Where "& SearchStr & " Order By UserID Desc")
		End Sub
		
		
End Class
%> 

﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_UserProgress
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserProgress
        Private KS,Action,Page,KSCls,flag,Str
		Private I, totalPut, MaxPerPage, SqlStr,ChannelID,ItemName,ItemName1,RS
		Private ch_rs,ch_sql,ModelEname,Inputer
		
		Private Sub Class_Initialize()
		  MaxPerPage = 10
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With Response
                Action=KS.G("Action")
				If Not KS.ReturnPowerResult(0, "KMUA10011") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF
				flag=KS.G("flag")
			If flag<>"excel" then
            %>
			<!DOCTYPE html>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<link href="../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
			<script src="../../KS_Inc/jquery.js"></script>
			</head>
			<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
			<script>
			 function ShowDetail(user,param){  
			   top.$.dialog.open("User/KS.UserProgress.asp?Action=ShowDetail&username="+escape(user)+"&"+param,{title:"查看用户搞件详细记录",width:860,height:500}); }
			</script>
			<body>
			<%If action<>"ShowDetail" then%>
			<div class='topdashed'><a href="?">用户组稿件统计</a> | <a href="?action=ShowUser">会员稿件统计</a> | <a href="?action=ShowAdmin">管理员稿件统计</a></div>
            <div class="pageCont2">
            <div class="tabTitle">稿件统计</div>
			<%
			end if
		    Else
			    Response.AddHeader "Content-Disposition", "attachment;filename=" & year(now) &"-" & month(now) &"-" & day(now) &".xls" 
				Response.ContentType = "application/vnd.ms-excel" 
				Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			End If
			 Page=KS.G("Page")
			
			 Select Case Action
			  Case "ShowDetail"  Call ShowDetail()
			  Case "ShowAdmin","ShowUser" Call ShowAdmin()
			  Case Else   Call MainList()
			 End Select
			 .Write "</div>"
			.Write "</body>"
			.Write "</html>"
			End With
		End Sub
		Sub MainList()
			  Dim Str,K,RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,ItemUnit From KS_Channel Where ChannelStatus=1 and channelid<>6 and BasicType<9 Order By ChannelID")
		      Dim SQL:SQL=RSC.GetRows(-1)
			  		str="<table width='100%' align='center' cellpadding='0' cellspacing='0'"
					if flag="excel" then str=str & " border='1'" Else str=str & " border='0'"
					str=str & ">   <tr class='sort'>"
					str=str & "    <td width='50' align='center'>序号</td>"
					str=str & "    <td align='center'>用户组</td>"
					str=str & "    <td width='120' align='center'>模块</td>"
					str=str & "    <td width='120' align='center'>今日</td>"
					str=str & "    <td width='120' align='center'>本周</td>"
					str=str & "    <td width='120' align='center'>本月</td>"
					str=str & "    <td width='120' align='center'>今年</td>"
					str=str & "    <td width='120' align='center'>所有</td>"
					str=str & "  </tr>"

				Call KS.LoadUserGroup()
				Dim Node,Param,GroupID
				For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
				
				str=str & "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			    str=str & "<td width='50' class='splittd' style='text-align:center'>" & i+1 & "</td>"
			    str=str & "<td class='splittd' align='center'>"
			    str=str & "<strong>" & Node.SelectSingleNode("@groupname").text & "</strong>"
			    str=str & "</td><td colspan='7' class='splittd' style='border-left:1px solid #ede7e7'>"
				str=str & "<table border='0' width='100%' cellspacing='0' cellpadding='0'>" &vbcrlf
		    For K=0 to Ubound(SQL,2)
			   GroupID=Node.SelectSingleNode("@id").text
			   Param=" a inner join KS_User b on a.Inputer=b.username where b.GroupID="&GroupID
			 str=str & "<tr>" &vbcrlf
			 str=str & "<td height='22' width='110' style='padding-left:10px'>" & SQL(1,k) & "</td>" & vbcrlf
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail("""",""groupid=" & groupid &"&ChannelID=" & SQL(0,K)&"&Flag=today"");' title='点击查看详情！'><font color=red>" & Conn.Execute("select count(1) from " & SQL(2,K) & Param &" And datediff("&DataPart_D&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail("""",""groupid=" & GroupID &"&ChannelID=" & SQL(0,K)&"&Flag=week"");' title='点击查看详情！'><font color=green>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_W&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail("""",""groupid=" & GroupID &"&ChannelID=" & SQL(0,K)&"&Flag=month"");' title='点击查看详情！'><font color=#ff6600>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_M&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail("""",""groupid=" & GroupID &"&ChannelID=" & SQL(0,K)&"&Flag=year"");' title='点击查看详情！'><font color=blue>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_Y&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
			 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail("""",""groupid=" & GroupID &"&ChannelID=" & SQL(0,K)&"&Flag=all"");' title='点击查看详情！'><font color=red>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param)(0) & "</font> " & SQL(4,K) & "</a></td>"
			 str=str & "</tr>" & vbcrlf
		
			Next
			str=str & "</table>"
				str=str & "</td></tr>"
				i=i+1
				Next
			str=str & "</table>"	
			
			If flag<>"excel" then
			 response.write str
			%>
			  <br/>
			  <div style="text-align:center">
			   <input type="button" value=" 打印本页 " class="button" onClick="window.print()"/>
			   <input type="button" value=" 导出Excel " class="button" onClick="location.href='?flag=excel';"/>
			  </div>
              
<br/><br/>
	<%         Else
	             Response.Write KS.ScriptHtml(str,"a",3)
	           End If
		End Sub
		
		
		
		Sub ShowAdmin()
		if flag<>"excel" then %>
		<div>
<form name='myform' action='KS.UserProgress.asp' method='post'>
<input type='hidden' value='<%=channelid%>' name='channelid'>
<input type='hidden' value='<%=ks.g("action")%>' name='action'>
搜索指定用户的稿件情况:<input type='text' class="textbox" name='username'>&nbsp;<input type='submit' class='button' value='搜索用户'>
</form>
</div>
		<%end if
		 str="<br><table width='100%' align='center' border='0' cellpadding='0' cellspacing='0'"
		 if flag="excel" then str=str & " border='1'" Else str=str & " border='0'"
		 str=str & ">   <tr class='sort'>"
		 str=str & "  <td width='50' align='center'>序号</td>"
		 str=str & "  <td align='center'>管理员</td>"
		str=str & "    <td width='130'align='center'>模块</td>"
		str=str & "    <td width='120' align='center'>今日</td>"
		str=str & "    <td width='120' align='center'>本周</td>"
		str=str & "    <td width='120' align='center'>本月</td>"
		str=str & "    <td width='120' align='center'>今年</td>"
		str=str & "    <td width='120' align='center'>所有</td>"
		str=str & "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		   Dim Param:Param=" where 1=1 "
		   if request("username")<>"" then param=param & " and a.username='" & ks.s("username") &"'"
		   If KS.G("Action")="ShowAdmin" Then
		   SqlStr = "SELECT a.UserName,a.RealName,a.Sex,b.userface,b.UserId FROM [KS_Admin] a inner join ks_user b on a.PrUserName=b.username " & Param & " order by AdminID"
		   Else
		   param=param & " and groupid<>1"
		   SqlStr = "SELECT UserName,RealName,Sex,userface,UserId FROM [KS_User] a  " & Param & " order by userid"
		   End If
			  RS.Open SqlStr, conn, 1, 1
			  If RS.Bof And RS.EOF Then
			   str=str & "<tr><td height=""30px"" style=""text-align:center"" colspan=15>没有找到对应的用户！</td></tr>"
			  Else
				     If KS.G("Action")="ShowAdmin" Then
					  totalPut = RS.RecordCount
					 Else
					  totalPut = Conn.Execute("select count(1) From KS_User a " & Param)(0)
					 End If
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
						    End If
					Call showContent
			End If
		  str=str & "  </td>"
		  str=str & "</tr>"

		 str=str & "</table>"
		 If flag<>"excel" then
			 Response.write str
			 Response.Write ("<div style='text-align:center'>")
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			 %>
			 <br/>
			  <div style="clear:both;text-align:center">
			   <input type="button" value=" 打印本页 " class="button" onClick="window.print()"/>
			   <input type="button" value=" 本页导出Excel " class="button" onClick="location.href='?action=<%=Action%>&flag=excel';"/>
			  </div>
              
<br/><br/>
			 <%
			 Response.Write ("</div><br/><br/>")
		 Else
		      Response.Write KS.ScriptHtml(str,"a",3)
		 End If
		End Sub
		Sub showContent()
		  Dim userface,Param,I,K,RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,ItemUnit From KS_Channel Where ChannelStatus=1 and channelid<>6 and BasicType<9 Order By ChannelID")
		  Dim SQL:SQL=RSC.GetRows(-1)
		  RSC.Close:Set RSC=Nothing
		
		  Do While Not RS.EOF
		   str=str & "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		   str=str & "<td width='50' class='splittd' style='text-align:center'>" & (currentpage-1)*maxperpage+i+1 & "</td>"
		   str=str & "<td  class='splittd' align='center'>"
		   userface=rs(3)
		   if ks.isnul(userface) then
		     if rs(2)="男" then 
			  userface=KS.Setting(2) & KS.Setting(3) &"images/face/boy.jpg"
			 else
			  userface=KS.Setting(2) & KS.Setting(3) &"images/face/girl.jpg"
			 end if
		   end if
		   if left(lcase(userface),4)<>"http" then userface=KS.Setting(2) & KS.Setting(3) & UserFace
		   str=str & "<a href='" & KS.GetSpaceUrl(rs(4)) & "' target='_blank'><img  onerror=""this.src='../images/face/boy.jpg';"" src='" & Userface & "' width='50' border='0'></a>"
		   str=str & "<div>" & RS(0) 
		   if not ks.isnul(rs(1)) then
		   str=str & "<br/><span class='tips'>(" & RS(1) & ")</span>"
		   end if
		   str=str & "</div></td><td colspan='7' height='22' class='splittd' style='border-left:1px solid #ede7e7'>"
		   
		  str=str & "<table border='0' width='100%' cellspacing='0' cellpadding='0'>" &vbcrlf
		    For K=0 to Ubound(SQL,2)
			   Param=" Where Inputer='" & RS(0) & "'"
			 str=str & "<tr>" &vbcrlf
			 str=str & "<td height='22' style='padding-left:10px'>" & SQL(1,k) & "</td>" & vbcrlf
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail(""" & RS(0) &""",""ChannelID=" & SQL(0,K)&"&Flag=today"");' title='点击查看详情！'><font color=red>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_D&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail(""" & RS(0) &""",""ChannelID=" & SQL(0,K)&"&Flag=week"");' title='点击查看详情！'><font color=green>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_W&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail(""" & RS(0) &""",""ChannelID=" & SQL(0,K)&"&Flag=month"");' title='点击查看详情！'><font color=#ff6600>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_M&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail(""" & RS(0) &""",""ChannelID=" & SQL(0,K)&"&Flag=year"");' title='点击查看详情！'><font color=blue>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff("&DataPart_Y&",AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
			 str=str & "<td width='120' align='center'><a href='javascript:ShowDetail(""" & RS(0) &""",""ChannelID=" & SQL(0,K)&"&Flag=all"");' title='点击查看详情！'><font color=red>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param)(0) & "</font> " & SQL(4,K) & "</a></td>"
			 str=str & "</tr>" & vbcrlf
		    Next
		   str=str & "</table>"
			
			
		   str=str & "</td>"
		   str=str & "</tr>"
		    I = I + 1
		    If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		   Loop
		   RS.Close
	End Sub
		 

		 
	 Sub ShowDetail()
		    With Response	
            End WIth
			Dim UserName:UserName=KS.G("UserName")
			Dim GroupID:GroupID=KS.ChkClng(Request("GroupID"))
			Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
			Dim Flag:Flag=KS.G("Flag")
			Dim SQLStr,Param
			 MaxPerPage = 15
			 If GroupID<>0 Then UserName="用户组【" & KS.U_G(GroupID,"groupname") & "】"
			Response.Write "<div style='height:35px;line-height:35px;text-align:center'>"
			Select Case Flag
			 Case "today"
			  Response.Write "查看<font color=red>" & UserName & "</font> 今天添加的" &KS.C_S(ChannelID,3) &""
			   Param=" And datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")=0"
			 Case "week"
			  Response.Write "查看<font color=red>" & UserName & "</font> 本周添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff(" & DataPart_W & ",AddDate," & SqlNowString & ")=0"
			 Case "month"
			  Response.Write "查看<font color=red>" & UserName & "</font> 本月添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff(" & DataPart_M & ",AddDate," & SqlNowString & ")=0"
			 Case "year"
			  Response.Write "查看<font color=red>" & UserName & "</font> 今年添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff(" & DataPart_Y & ",AddDate," & SqlNowString & ")=0"
			 Case "all"
			  Response.Write "查看<font color=red>" & UserName & "</font> 所有添加的" &KS.C_S(ChannelID,3)
			End Select
			
			if groupid=0 then
			     param=" where Inputer='" & UserName & "'" & param
				 SQLStr="Select id,title,Inputer,adddate from " & KS.C_S(ChannelID,2) & " a "
			else 
			     param=" inner join ks_user b on a.inputer=b.username Where  b.groupid=" & groupid & param
				 SQLStr="Select a.id,a.title,a.Inputer,a.adddate from " & KS.C_S(ChannelID,2) & " a "
			end if
		
			SQLStr=SQLStr & Param & " Order By a.ID Desc"
			
			Response.Write ",共计 <span id='total' style='color:brown'>0</span> 条数据</div>"
			Response.Write "<table width='95%' align='center' border='0' cellpadding='0' cellspacing='0'>"
			Response.Write "    <tr class='sort'>"
			Response.Write "    <td width='100' align='center'>ID</td>"
			Response.Write "    <td align='center'>名称</td>"
			Response.Write "    <td  align='center'>录入员</td>"
			Response.Write "    <td  align='center'>录入时间</td>"
			Response.Write "    <td width='100' align='center'>查看详情</td>"
			Response.Write "  </tr>"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
             RS.Open SqlStr, conn, 1, 1
				 If Not RS.EOF Then
					  
					  totalPut = conn.execute("select count(1) from " & KS.C_S(ChannelID,2) & " a " &  param)(0)
					  response.write "<script>$('#total').html(" & totalput &");</script>"
						If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
						End If
								Call showDetailContent(ChannelID)
			End If
		 Response.Write "  </td>"
		 Response.Write "</tr>"

		 Response.Write "</table>"
		 Response.Write ("<div style='display:block;text-align:center'>")
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	     Response.Write ("</div")
		 Response.Write "</table><Br/><br/>"
		 End Sub
		 
		 Sub showDetailContent(ChannelID)
		  Dim I:I=0
		  Do While Not RS.Eof
		   Response.Write "<tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		   Response.Write "<td class='splittd' style='height:20px;' align='center'>" & RS(0) & "</td>"
		   Response.Write "<td class='splittd'>" & KS.Gottopic(RS(1),50) & "</td>"
		   Response.Write "<td class='splittd' align='center'>" & RS(2) & "</td>"
		   Response.Write "<td class='splittd' align='center'>" & RS(3) & "</td>"
		   Response.Write "<td class='splittd' align='center'><a href='../../item/show.asp?d=" & RS(0) &"&m=" & channelid & "' target='_blank'>查看内容</a></td>"
		   Response.Write "</tr>"
		  	I = I + 1
		    If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		   Loop
		   RS.Close
		 End Sub
End Class
%> 

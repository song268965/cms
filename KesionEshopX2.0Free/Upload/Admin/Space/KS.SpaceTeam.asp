<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
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
Set KSCls = New SpaceTeam
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceTeam
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10004") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write EchoUeditorHead()
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeam.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>圈子管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?action=topic';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon num'></i>帖子管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeamSkin.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon merge'></i>模板管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon folder'></i>圈子分类</span></li>"
			  .Write "</ul>"
			End With
		
		
		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.End
		End If
		CurrentPage = KS.ChkClng(request("page"))
		If CInt(CurrentPage) <= 0 Then CurrentPage = 1
		Select Case KS.G("action")
		 Case "Del"
		  Call TeamDel()
		 Case "lock"
		  Call TeamLock()
		 Case "unlock"
		  Call TeamUnLock()
		 Case "verific"
		  Call TeamVerific()
		 Case "recommend"
		  Call Blogrecommend()
		 Case "Cancelrecommend"
		  Call BlogCancelrecommend()
		 case "topic" topicshow
		 case "topicdel" topicdel
		 case "class" classshow
		 case "modify" modify
		 case "modifysave" modifysave
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<div class="pageCont2">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>圈子名称</th>
	<td nowrap>创建者</th>
	<td nowrap>创建时间</th>
	<td nowrap>圈子状态</th>
	<td nowrap>浏览权限</th>
	<td nowrap>管理操作</th>
</tr>
<%
		totalPut = Conn.Execute("Select Count(ID) from KS_Team")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Team order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>没有用户申请圈子！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="KS.SpaceTeam.asp">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="../../space/group.asp?id=<%=rs("id")%>" target="_blank"><%=Rs("Teamname")%></a>
	<%If rs("recommend")=1 then response.write "<font color=red>荐</font>"%>
	</td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("verific")
	 case 0
	  response.write "<font color=red>未审</font>"
	 case 1
	  response.write "<font color=green>已审</font>"
	 case 2
	  response.write "<font color=blue>锁定</font>"
	end select
	%></td>
	<td class="splittd" align="center"><%
	select case rs("viewtf")
	 case 0
	  response.write "<font color=#999999>不限</font>"
	 case 1
	  response.write "<font color=green>加入成员</font>"
	 case 2
	  response.write "<font color=blue>注册会员</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../../space/group.asp?id=<%=rs("id")%>" target="_blank" class="setA">浏览</a>| 
	<a href="?action=modify&id=<%=rs("id")%>" onclick="window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('空间门户管理 >> <font color=red>修改圈子信息</font>')+'&ButtonSymbol=GOSave';" class="setA">编辑</a>|
	
	<a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('删除圈子将删除圈子下的所有信息，确定删除吗？'));" class="setA">删除</a>|&nbsp;<%if rs("verific")=0 then%><a href="?Action=verific&id=<%=rs("id")%>" class="setA">审核</a>|<%elseif rs("verific")=1 then%><a href="?Action=lock&id=<%=rs("id")%>" class="setA">锁定</a>|<%elseif rs("verific")=2 then%><a href="?Action=unlock&id=<%=rs("id")%>" class="setA">解锁</a>|<%end if%>
	<%if rs("recommend")="0" then%>
	<a href="?Action=recommend&id=<%=rs("id")%>" class="setA">设为推荐</a>
	<%else%>
	<a href="?Action=Cancelrecommend&id=<%=rs("id")%>" class="setA"><font color=red>取消推荐</font></a>
	<%end if%>
	
	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='operatingBox' onMouseOver="this.className='operatingBox'" onMouseOut="this.className='operatingBox'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input type="hidden" name="action" value="Del">
	<input class="button" type="submit" name="Submit2" value="批量删除" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){document.getElementById('action').value='Del';this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" value="批量审核" onclick="document.getElementById('action').value='verific';">
	<input class="button" type="submit" value="批量锁定" onclick="document.getElementById('action').value='lock';">
	<input class="button" type="submit" value="批量解锁" onclick="document.getElementById('action').value='unlock';">
	</td>
</tr>
</form>
<tr>
	<td  class='tdbg' onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'" colspan=7 align=right>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
</div>
<%
End Sub


sub modify()
Dim TeamName,username,PhotoUrl,ClassID,Announce,Note,Verific,ViewTF,JoinTF,Recommend
 Dim ID:ID=KS.ChkClng(Request("id"))
 If Id<>0 Then
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "select * from ks_team where id=" & id,conn,1,1
	 If RS.Eof AND RS.Bof Then
	   RS.Close
	   Set RS=Nothing
	   KS.AlertHintScript "对不起，找不到记录！"
	 End If
	 TeamName = RS("TeamName")
	 username = RS("username")
	 PhotoUrl = RS("PhotoUrl")
	 ClassID  = RS("ClassID")
	 Announce = RS("Announce")
	 Note     = RS("Note")
	 Verific  = RS("Verific")
	 ViewTF   = RS("ViewTF")
	 JoinTF   = RS("JoinTF")
	 Recommend   = RS("Recommend")
 Else 
     KS.AlertHintScript "对不起，找不到记录！"
 End If
 %>
 <script type="text/javascript">
 function CheckForm()
 {
 <%if request("action")="add" then%>
   if ($("input[name=username]").val()=='')
   {
     alert('用户名称必须输入！');
	 $("input[name=username]").focus();
	 return false;
   }
 <%end if%>
   if ($("input[name=TeamName]").val()=='')
   {
     alert('空间名称必须输入！');
	 $("input[name=TeamName]").focus();
	 return false;
   }
   $("#myform").submit();
 }
 </script>
 <div class="pageCont2">
 <form name="myform" id="myform" action="?action=modifysave" method="post">
   <input type="hidden" value="<%=ID%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
 <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="otable mt0">
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>创建人：</strong></td>
           <td height='28'><%=username%></td>
          </tr>
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>圈子名称：</strong></td>
           <td height='28'><input type='text' name='TeamName' class="textbox" value='<%=TeamName%>' size="40"> <font color=red>*</font></td>
          </tr> 
		  
         
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>圈子图片：</strong></td>
           <td height='28'><input type='text' class="textbox" name='PhotoUrl' value='<%=PhotoUrl%>' size="40"></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>圈子分类：</strong></td>
           <td height='28'><select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-请选择类别-</option>
                    <% Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
							  RSC.Open "Select * From KS_TeamClass order by orderid",conn,1,1
							  If Not RSC.EOF Then
							   Do While Not RSC.Eof 
							   If Trim(ClassID)=trim(RSC("ClassID")) Then
								  Response.Write "<option value=""" & RSC("ClassID") & """ selected>" & RSC("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RSC("ClassID") & """>" & RSC("ClassName") & "</option>"
							   End iF
								 RSC.MoveNext
							   Loop
							  End If
							  RSC.Close:Set RSC=Nothing
							  %>
                  </select>   </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>圈子加入说明：</strong></td>
           <td height='28'><%
		   Response.Write EchoEditor("Note",Note,"Basic","96%","200px")
			%> </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间公告：</strong></td>
           <td height='28'><%
		   Response.Write EchoEditor("Announce",Announce,"Basic","96%","200px")
			%>
		   </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>圈子加入条件：</strong></td>
           <td height='28'><input type="radio" value="1" name="JoinTF" <%if JoinTF="1" then response.write " checked"%>>任意加入
                       <input type="radio" value="2" name="JoinTF"<%if JoinTF="2" then response.write " checked"%>>申请加入
                       <input type="radio" value="3" name="JoinTF"<%if JoinTF="3" then response.write " checked"%>>仅可邀请

          </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>圈子浏览权限：</strong></td>
           <td height='28'><input type="radio" value="0" name="ViewTF"<%if ViewTF="0" then response.write " checked"%>>无任何限制
                       <input type="radio" value="1" name="ViewTF"<%if ViewTF="1" then response.write " checked"%>>仅加入本圈子的成员
                       <input type="radio" value="2" name="ViewTF"<%if ViewTF="2" then response.write " checked"%>>注册会员
          </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>是否推荐：</strong></td>
           <td height='28'><input name="recommend" type="radio" value="1"<%if recommend=1 then response.write " checked"%> /> 是 <input name="recommend" type="radio" value="0" <%if recommend=0 then response.write " checked"%>/> 否</td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbg'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>状态：</strong></td>
           <td height='28'><input name="verific" type="radio" value="1"<%if verific=1 then response.write " checked"%> /> 已审核 <input name="verific" type="radio" value="0" <%if verific=0 then response.write " checked"%>/> 未审核<input name="verific" type="radio" value="2" <%if verific=2 then response.write " checked"%>/> 锁定</td>
          </tr>  
         
 </table>
   </form>
 </div>
 <%
End Sub

Sub ModifySave()

 Dim ID:ID=KS.ChkClng(Request("id"))
 Dim UserID,UserName:UserName=KS.G("UserName")
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 If ID=0 Then  
   If KS.IsNul(UserName) Then
    KS.Die "<script>alert('请输入圈子创始人的用户名!');history.back();</script>"
   End If
   RS.Open "select top 1 userid from KS_User Where UserName='" & UserName & "'",conn,1,1
   If RS.Eof And RS.Bof Then
     RS.Close
	 Set RS=Nothing
	 KS.Die "<script>alert('对不起，您输入的用户名不存在!');history.back();</script>"
   End If
   UserID=RS(0)
   RS.Close
 End If
 RS.Open "select top 1 * from ks_Team where id=" & id,conn,1,3
 If RS.Eof AND RS.Bof Then
   RS.addnew
   RS("UserName")=UserName
   RS("AddDate")=Now
 End If
 RS("TeamName")=KS.G("TeamName")
 RS("PhotoUrl")=KS.G("PhotoUrl")
 RS("ClassID")=KS.ChkClng(KS.G("ClassID"))
 RS("Note")=Request.Form("Note")
 RS("Announce")=Request.Form("Announce")
 RS("verific")=KS.ChkClng(KS.G("verific"))
 RS("viewTF")=KS.ChkClng(KS.G("viewTF"))
 RS("JoinTF")=KS.ChkClng(KS.G("JoinTF"))
 RS("recommend")=KS.ChkClNG(KS.G("recommend"))
 RS.Update
 RS.Close
 Set RS=Nothing
  if id=0 then
 Response.Write "<script>alert('恭喜，圈子添加成功！');$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>圈子管理</font>');location.href='KS.team.asp';</script>"
  else
 Response.Write "<script>alert('恭喜，圈修改成功！');$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>圈子管理</font>');location.href='"& Request.Form("ComeUrl") & "';</script>"
 end if
End Sub





	'删除
	Sub TeamDel()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 Dim tid
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
	 RS.Open "Select * from KS_Team Where ID in(" & id & ")",conn,1,1
	 do while not rs.eof
	  tid=rs("id")
  	Conn.execute("delete from ks_uploadfiles where channelid=1030 and infoid in(" & tid& ")")
  	Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where teamid in(" & tid& "))")
	  Conn.execute("Delete From KS_TeamTopic Where teamid=" & tid)
	  Conn.Execute("Delete From KS_TeamUsers Where teamid=" & tid)
	  rs.movenext
	 loop
	 rs.close:set rs=nothing
	 Conn.execute("Delete From KS_Team Where ID In("& id & ")")
	 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	'锁定
	Sub TeamLock()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set verific=2 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	'解锁
	Sub TeamUnLock()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set verific=1 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	'审核
	Sub TeamVerific
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set verific=1 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	sub Blogrecommend()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set recommend=1 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	end sub
	
	sub BlogCancelrecommend()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set recommend=0 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	end sub
     
	'帖子管理
    Sub topicshow()
%>
<div class="pageCont2">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>主题</th>
	<td nowrap>用 户 名</th>
	<td nowrap>发 表 时 间</th>
	<td nowrap>状 态</th>
	<td nowrap>管 理 操 作</th>
</tr>
<%
		totalPut = Conn.Execute("Select Count(ID) from KS_TeamTopic")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_TeamTopic order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7 class='splittd'>没有人发表圈子主题！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=topicdel>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd">
	<%if rs("parentid")=0 then
	 response.write "<font color=red>[主]</font>"
	 end if
	 %>
	<a href="../../space/group.asp?action=showtopic&id=<%=rs("teamid")%>&tid=<%=rs("id")%>" target="_blank"><%=Rs("title")%></a><% if rs("isbest")="1" then response.write "<img src=""../images/jh.gif"" align=""absmiddle"">"%></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "未审"
	 case 1
	  response.write "<font color=red>正常</font>"
	 case else
	  response.write "屏蔽"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../../space/group.asp?action=showtopic&id=<%=rs("teamid")%>&tid=<%=rs("id")%>" target="_blank">浏览</a> <a href="?Action=topicdel&ID=<%=RS("ID")%>" onclick="return(confirm('确定删除该帖子吗？'));">删除</a> </td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='operatingBox' onMouseOver="this.className='operatingBox'" onMouseOut="this.className='operatingBox'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的主题 " onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=7 align=right>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
</div>
<%
End Sub

'删除帖子
Sub topicdel()
 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
 If ID="" Then KS.Die "<script>alert('对不起，您没有选择!');history.back();</script>"

		 dim rst:set rst=server.createobject("adodb.recordset")
		 rst.open "select * from ks_teamtopic where id in(" & id & ")",conn,1,1
		 if not rst.eof then
		   do while not rst.eof
			 Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(" & rst("id")& ")")
			 Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where parentid=" & rst("id")& ")")
		     conn.execute("delete from ks_teamtopic where parentid=" & rst("id"))
			 rst.movenext
		   loop
		 end if
		 rst.close:set rst=nothing
		 conn.execute("delete from ks_teamtopic where id in(" & id & ")")
		 response.write "<script>alert('删除成功');location.href='"& request.servervariables("http_referer") & "';</script>"
End Sub

'分类管理
Sub ClassShow()
%>		
		<div class="pageCont2">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>编号</strong></td>
			<td width="217"><strong>类型名称</strong></td>
			<td width="197"><strong>排序</strong></td>
			<td width="196"><strong>管理操作</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_TeamClass order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">还没有添加任何的圈子分类!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?action=class&x=a">
				<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" align="center"><%=rs("ClassID")%> <input name="ClassID" type="hidden" id="ClassID" value="<%=rs("ClassID")%>"></td>
				  <td class="splittd" width="217" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" value="<%=rs("ClassName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="Submit" class="button" type="submit"value=" 修改 ">&nbsp;
				  <a onclick="return(confirm('确定删除吗?'))" href="?action=class&x=c&classid=<%=rs("classid")%>">删除</a></td>
				</tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?action=class&x=b" method="post" name="myform" id="form">
		    <tr>
		      <td class="spltitd" colspan="4" height="25">&nbsp;&nbsp;<strong>&gt;&gt;新增圈子分类<<</strong></td>
		    </tr>
			<tr valign="middle" class="list"> 
			  <td class="spltitd" height="25"></td>
			  <td class="spltitd" height="25" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" size="25"></td>
			  <td class="spltitd" height="25" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td class="spltitd" height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
		</form>
</table>
</div>  
		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_TeamClass set ClassName='" & KS.G("ClassName") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "恭喜,分类修改成功!"
		   case "b"
		       If KS.G("ClassName")="" Then Response.Write "<script>alert('请输入类型名称!');history.back();</script>":response.end
			   
				conn.execute("Insert into KS_TeamClass(ClassName,orderid)values('" & KS.G("ClassName") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				KS.AlertHintScript "恭喜,分类添加成功!"
		   case "c"
				conn.execute("Delete from KS_TeamClass where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "恭喜,分类删除成功!"
		End Select
		%></body>
		</html>
<%End Sub

End Class
%> 

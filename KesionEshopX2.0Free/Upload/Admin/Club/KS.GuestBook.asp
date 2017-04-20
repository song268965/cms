<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
Response.Buffer=true
Response.CharSet="utf-8"
Server.ScriptTimeout=9999999

'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Guest_Manage
KSCls.Kesion()
Set KSCls = Nothing

Class Guest_Manage
        Private KS,Action,KSCls
	    Private MaxPerPage, TotalPut , CurrPage, TotalPage, i, j, Loopno
	    Private KeyWord, SearchType,SqlStr,RS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
	Public Sub Kesion()
	 If Not KS.ReturnPowerResult(0, "KSMB10000") Then                  '权限检查
		Call KS.ReturnErr(1, "")   
		Response.End()
	 End iF
	KeyWord = KS.R(Trim(Request("keyword")))
	SearchType = KS.R(Trim(Request("SearchType")))
	CurrPage = KS.ChkClng(Request("Page")) : If CurrPage<=0 Then CurrPage=1
	Action=KS.G("Action")
	Select Case Action
	 Case "Comment" Comment
	 Case "Main"  GuestMain
	 Case "Del"   GuestDel
	 Case "Revert" Revert
	 Case "DelRecycle" DelRecycle
	 Case "DelRecycleAll" DelRecycleAll
	 Case "DoVerifyReply" DoVerifyReply
	 Case "DelReply" DelReply
	 Case "Delete" DeleteTJ
	 Case "doDelete" doDeleteTJ
	 Case Else  GuestMain
	 End Select
	End Sub
Sub Nav()
%>
<script>
$(function(){
			var myurl = ""+window.location+"";
			myurl=myurl.split('Club/')[1];
			$(".tabs li").each(function() {
				var url = $(this).find("a").attr("href");	
				if(url==myurl){
					$(this).addClass("active");
				};
			});
			
		});

</script>
<div class='tabs_header'>
<ul class="tabs"><li><a href="KS.GuestBook.asp"><span>主题管理</span></a></li><li><a href="KS.GuestBook.asp?Action=Recycle"><span>帖子回收站</span></a></li><li><a href="KS.GuestBook.asp?Action=Comment"><span>帖子点评管理</span></a></li><li><a href="KS.GuestBook.asp?Action=VerifyReply"><span>帖子回复审核</span></a></li><li><a href="KS.GuestBook.asp?action=Delete"><span>按条件删除操作</span></a></li></ul>
</div>
<!--<div class="menu_top_fixed_height"></div>-->
<%
End Sub


Sub DeleteTJ()
%>

<!DOCTYPE html>
<html>
<head>
<title>点评管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="../Include/admin_Style.css" type="text/css">
<script src="../../ks_inc/common.js"></script>
<script src="../../ks_inc/jquery.js"></script>
</head>
<body>
<%Nav%>
<div class="pageCont">
<form name="myform" action="?action=doDelete" method="post">
<table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
		<tr class="sort">
			<td colspan="2">按条件删除操作</td>
		</tr>
		<tr><td class="pt10" colspan="2"></td></tr>
		<tr class="list">
		  <td height="30" align="right" width="200"><strong>用户名：</strong></td>
		  <td>
		  <input size="20" type="text" name="key" class="textbox"/>
		  <span class="tips">用户名如果不输入，将删除符合条件的所有记录！</span>
		  </td>
		  </td>
		</tr>
		<tr class="list">
		  <td height="40" align="right" width="200"><strong>时间限制：</strong></td><td>
		  <select name="timelimit" class="textbox">
		   <option value="0">---删除所有数据---</option>
		   <option value="1">---删除1天内的数据---</option>
		   <option value="3">---删除3天内的数据---</option>
		   <option value="5">---删除5天内的数据---</option>
		   <option value="15">---删除15天内的数据---</option>
		   <option value="30">---删除30天内的数据---</option>
		   <option value="60">---删除2个月内的数据---</option>
		   <option value="90">---删除3个月内的数据---</option>
		   <option value="180">---删除6个月内的数据---</option>
		   <option value="365">---删除1年内的数据---</option>
		  </select>
		  </td>
		  </td>
		</tr>
		<tr class="list">
		  <td height="40" align="right" width="200"><strong>删除数据表：</strong></td><td>
		  <select name="delTable" class="textbox">
		   <option value="0">---不限(符合条件的主题和回复都删除)---</option>
		   <option value="KS_GuestBook">主题表(KS_GuestBook)</option>
		   <%
		   Dim Node,TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			TableXML.async = false
			TableXML.setProperty "ServerHTTPRequest", true 
			TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
			Dim Url,N:N=0
			For Each Node In TableXML.DocumentElement.SelectNodes("item")
			  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"'>回复表(" & Node.SelectSingleNode("tablename").text&")</option>"
			Next
		   %>
		  </select>
		  </td>
		  </td>
		</tr>
		<tr class="list">
		  <td height="40" align="right" width="200"><strong>版面限制：</strong></td><td>
		      <%With Response
			    .Write " <select name='boardid' id='boardid' class='textbox'>"
				.Write "  <option value='0'>---版面不限制---</option>"
				 Call KS.LoadClubBoard()
			     Dim Xml,Str
			     Set Xml=Application(KS.SiteSN&"_ClubBoard")
			     for each node in xml.documentelement.selectnodes("row[@parentid=0]")
				      .Write ("<OPTGROUP label=&nbsp;+" & node.selectsinglenode("@boardname").text & " </OPTGROUP>")
					for each n in xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
					  .Write ("<option value='" & N.SelectSingleNode("@id").text & "'>&nbsp;|-" & n.selectsinglenode("@boardname").text &"</option>")
					next
				next
				
				.Write "</select>"
			  End With
			  %>
		  </td>
		  </td>
		</tr>
        <tr class="list">
		  <td height="30" align="right" width="200"><strong>内容含有关键字：</strong></td>
		  <td>
		  <input size="20" type="text" name="content" class="textbox"/>
		  <span class="tips">可以留空，否则只删除含有这里关键词的记录！</span>
		  </td>
		  </td>
		</tr>
		<tr class="list">
		  <td height="40" align="right" width="200"><strong>删除选项：</strong></td><td>
		   <input type="radio" name="DelType" value="0" checked="checked">放入回收站
		   <input type="radio" name="DelType" value="1">直接删除
		  </td>
		  </td>
		</tr>

		<tr class="list">
		  <td height="40" align="right" width="200"></td><td>
		   <input type="submit" value="执行删除操作" onClick="return(confirm('提示：执行本操作我们强烈建议您做好数据库备份工作！您现在确认执行本操作吗？'))" class="button"/>
		  </td>
		  </td>
		</tr>
</table>
</form>	
</div>
<div class="footerTable"></div>	
</body>
</html>
<%
End Sub

Sub doDeleteTJ()
  Dim key:key=KS.S("key")
  Dim timelimit:timelimit=KS.ChkClng(KS.S("timelimit"))
  Dim delTable:delTable=Lcase(KS.S("delTable"))
  Dim boardid:boardid=KS.ChkClng(KS.S("boardid"))
  Dim DelType:DelType=KS.ChkClng(KS.S("DelType"))
  Dim content:content=KS.S("content")
  IF KS.IsNul(Key) And timelimit=0 Then
    KS.Die "<script>top.$.dialog.alert('用户名没有输入的情况，删除的时间范围不能没有限制!',function(){ history.back() })</script>"
  End If
  Dim TopicIDs,RS
  
    Dim Param:Param=" Where DelTF=0"
	If Not KS.IsNul(Key) Then Param=Param & " and UserName='" & Key & "'"
	If Not KS.IsNul(content) Then Param=Param & " and content like '%" & content &"%'"
	
	If delTable="0" Then   '删除符合条件的所有数据
	    
		Dim Param1:Param1=Param
	    If timelimit<>0 Then '时间限制
		  Param=Param & " and datediff(" & DataPart_D & ",AddTime," & SQLNowString &")<=" & timelimit
		  Param1=Param1 & " and datediff(" & DataPart_D & ",ReplayTime," & SQLNowString &")<=" & timelimit
		End If
		
		If boardId<>0 Then
		  Param=Param & " and boardid=" & BoardID 
		  Param1=Param1 & " and topicid in(select id from ks_GuestBook Where boardid=" & BoardID &")"
		End If
	
	    Dim TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		TableXML.async = false
		TableXML.setProperty "ServerHTTPRequest", true 
		TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
		Dim Url,Node,N:N=0
		
		For Each Node In TableXML.DocumentElement.SelectNodes("item")
		  Call DelReplyByTable(Node.SelectSingleNode("tablename").text,Param1,DelType)
		Next
	    
		'删除主题
		If DelType=0 Then '仅放入回收站
			   Conn.Execute("Update KS_GuestBook Set DelTF=1" & Param)
		 Else
		        TopicIds=""
				  Set RS=Conn.Execute("Select ID,ChannelID,InfoID From KS_GuestBook " & Param)
				  Do While Not RS.Eof 
					If TopicIDs="" Then
					 TopicIDs=RS(0)
					Else
					 TopicIDs=TopicIDs & "," & RS(0)
					End If
				  RS.MoveNext
				  Loop
				  RS.Close : Set RS=Nothing
				  If TopicIDs<>"" Then doDelete(TopicIds)
		  End If
	      
	Else
	
	    If timelimit<>0 Then '时间限制
		 If delTable="ks_guestbook" then
		  Param=Param & " and datediff(" & DataPart_D & ",AddTime," & SQLNowString &")<=" & timelimit
		 else
		  Param=Param & " and datediff(" & DataPart_D & ",ReplayTime," & SQLNowString &")<=" & timelimit
		 end if
		End If
		
		If boardId<>0 Then
		  If delTable="ks_guestbook" then 
		   Param=Param & " and boardid=" & BoardID
		  Else
		   Param=Param & " and topicid in(select id from ks_GuestBook Where boardid=" & BoardID &")"
		  End If
		End If
		
		
		  if deltable="ks_guestbook" then
		    If DelType=0 Then '仅放入回收站
			   Conn.Execute("Update " & delTable & " Set DelTF=1" & Param)
			Else
		       TopicIds=""
			   if lcase(delTable)="ks_guestbook" then  '彻底删除
				  Set RS=Conn.Execute("Select ID,ChannelID,InfoID From KS_GuestBook " & Param)
				  Do While Not RS.Eof 
					If TopicIDs="" Then
					 TopicIDs=RS(0)
					Else
					 TopicIDs=TopicIDs & "," & RS(0)
					End If
				  RS.MoveNext
				  Loop
				  RS.Close : Set RS=Nothing
				  If TopicIDs<>"" Then doDelete(TopicIds)
			 End If
		   End If
		  Else
		    Call DelReplyByTable(DelTable,Param,DelType)
		  End If

    End If
	KS.AlertHintScript "恭喜，按设定的条件删除成功！"
End Sub

'删除回复表
Sub DelReplyByTable(DelTable,Param,DelType)
  Dim TopicIds:TopicIds=""
  Dim RS:Set RS=Conn.Execute("Select TopicId From " & DelTable & Param)
  Do While Not RS.Eof 
				If TopicIDs="" Then
				 TopicIDs=RS(0)
				Else
				 TopicIDs=TopicIDs & "," & RS(0)
				End If
			  RS.MoveNext
	 Loop
	  RS.Close
	  If DelType=0 Then '回收站
	     Conn.Execute("Update " & DelTable & " Set Deltf=1 " & Param)
	  Else
		 Conn.Execute("Delete From " & DelTable & Param)
	  End If
	  If Not KS.IsNul(TopicIDs) Then '更新主题回复数
		    RS.Open "select id From KS_GuestBook Where id IN (" & TopicIDs & ") group by id",CONN,1,1
	    Do While NOT RS.Eof
			     Dim RNum:Rnum=Conn.Execute("select count(1) From " & DelTable &" WHERE DelTF=0 and TopicID="& RS("ID"))(0)
			     Conn.Execute("Update KS_GuestBook Set TotalReplay=" & rNum & " Where ID=" & RS(0))
		   RS.MoveNext
	    Loop
		  RS.Close
	   End If
   Set RS=Nothing
End Sub

Sub Comment()
 If Request.QueryString("flag")="del" then
   Dim ID:ID=KS.FilterIds(KS.G("ID"))
   If Id="" Then KS.AlertHintScript "对不起，您没有选择记录!" : Exit Sub
   Conn.Execute("Delete From KS_GuestComment Where ID in (" & ID & ")")
   KS.AlertHintScript "恭喜，删除成功！"
 End If
%>
<!DOCTYPE html>
<html>
<head>
<title>点评管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="../Include/admin_Style.css" type="text/css">
<script src="../../ks_inc/common.js"></script>
<script src="../../ks_inc/jquery.js"></script>
<body>
<%Nav%>
<div class="pageCont">
<form name="myform" action="KS.GuestBook.asp?Action=Comment&flag=del" method=post>
<table border="0" width="100%" align="center"  cellpadding="0" cellspacing="0">
		<tr class="sort">
					<td>&nbsp;</td>
					<td>主题</td>
					<td>评论内容</td>
					<td>用户</td>
					<td>时间</td>
					<td>威望</td>
					<td>管理操作</td>
		</tr>
		 <%
		MaxPerPage=20
		SQLStr=KS.GetPageSQL("KS_GuestComment","id",MaxPerPage,CurrPage,1,"","*")
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,Conn,1,1 
		If RS.Eof or RS.Bof Then 
			Response.Write "<tr class='list'><td colspan='10' align='center' height='30' class='splittd'>暂时还没有任何记录！</td></tr>"
		Else
			totalPut = Conn.Execute("Select count(id) from [KS_GuestComment]")(0)
			i = 0
			Do While Not RS.Eof 
			%>
			<tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=RS("ID")%>' onClick="chk_iddiv('<%=RS("ID")%>')">
			  <td  height="30" class='splittd' align="center" valign="middle"><input onClick="chk_iddiv('<%=RS("ID")%>')" type="checkbox" id='c<%=Trim(RS("ID"))%>' name="id" value="<%=Trim(RS("ID"))%>"></td>
			  <td class='splittd'>
			  <%
			  Dim RST:Set RST=Conn.Execute("Select top 1 Subject From KS_GuestBook Where id=" & rs("tid"))
			  If Not RST.Eof Then
			    Response.Write "<a href='" & KS.GetClubShowUrl(rs("tid")) & "' target='_blank'>" & RST(0) & "</a>"
			  End if
			  RST.Close :Set RST=Nothing
			  %>
			  </td>
			  <td class='splittd'><%=KS.Gottopic(rs("comment"),50)%></td>
			  <td class='splittd'>&nbsp;<%=rs("username")%></td>
			  <td class='splittd' align="center"><%=rs("adddate")%></td>
			  <td class='splittd' align="center"><%=rs("Prestige")%></td>
			  <td class='splittd' align="center"><a href="?action=Comment&flag=del&id=<%=rs("id")%>" onClick="return(confirm('确定删除该点评吗？'))">删除</a></td>
			</tr>
			<%
			RS.MoveNext
			Loop
		End If
		RS.Close
		 %>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" class='operatingBox'>
          <tr>
		    <td ><label><input type="checkbox"  onClick="if (this.checked){Select(0)}else{Select(1)}">全部选中</label>
              <input name="delbtn" value="批量删除"  class="button" type="submit" onClick="return(confirm('确定删除吗？'));">
			</td>

          </tr>	

</table>
</form>
</div>
<div class="footerTable pt10"> <%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,true,true)
%>
</div>
</body>
</html>
<%
End Sub

Sub GuestMain()
%>
<!DOCTYPE html>
<html>
<head>
<title>内容管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="../Include/admin_Style.css" type="text/css">
<script src="../../ks_inc/common.js"></script>
<script src="../../ks_inc/jquery.js"></script>
<script language="JavaScript">
<!--
function cdel()
{   if(get_Ids(document.myform)==''){alert('您没有选择记录!');return false;};
	if (confirm("你真的要删除这条主题吗？不可恢复！")){
		document.myform.Flag.value = "del";
		document.myform.submit();
	}
}
function ccheck()
{
	if(get_Ids(document.myform)==''){alert('您没有选择记录!');return false;};
	if (confirm("你确定要审核选中主题吗？")){
		document.myform.Flag.value = "check";
		document.myform.submit();
	}
}
function cuncheck()
{
	if(get_Ids(document.myform)==''){alert('您没有选择记录!');return false;};
	if (confirm("你确定要取消选中主题的审核吗？")){
		document.myform.Flag.value = "uncheck";
		document.myform.submit();
	}
}
//-->
</script>

<%Nav%>

<%if request("action")="Recycle" Then
    Call Recycle() : Exit Sub
  elseif request("action")="VerifyReply" then
    Call VerifyReply() : Exit Sub
  end If
%>
<div class="tableTop noborder">
	<form action="KS.GuestBook.asp" method="post" name="search" id="search">
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="TableBar">
		<tr>
		  <td height="25"><strong class="mr0">快速搜索 --&gt;&gt;</strong><span class="tiaoJian"> 关键词：</span>
			    <input type="text" name="keyword" class="textbox" size="35" value="<%=KeyWord%>" onMouseOver="this.focus()" onFocus="this.select()">
				<select name="SearchType" size="1" class="inputlist">
				  <option value="content" <%If SearchType = "content" Then Response.Write "selected"%>>帖子主题</option>
				  <option value="author" <%If SearchType = "author" Then Response.Write "selected"%>>发帖者</option>
				</select>
				<input type="submit" class="button" name="imageField" value=" 搜 索 ">
		  </td>
		</tr>
	</table>
	</form>
</div>
<div class="pageCont2 mt20">

	<form name="myform" action="KS.GuestBook.asp?Action=Del" method=post>
	<input name="Flag" type="hidden" value="" id="Flag">
<table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
		<tr class="sort">
					<td><input type="checkbox"  name='selectCheck' onClick="if (this.checked){Select(0)}else{Select(1)}"></td>
					<td width="60%" style="text-align:left;padding-left:35px;">主题</td>
					<td>发帖者</td>
					<td>回复/查看</td>
					<td>最后发表</td>
					<td>状态</td>
		</tr>
	<%
	Dim Param:Param=" Deltf=0"
	If Not KS.IsNul( KeyWord) Then
		If SearchType = "content" Then
			Param=param & " and Subject LIKE '%"& KeyWord &"%'"  
		Else
			Param=param & " and UserName LIKE '%"& KeyWord &"%'" 
		End If
	ENd If
	MaxPerPage=20
	CurrPage = KS.ChkClng(Request("Page")) : If CurrPage<=0 Then CurrPage=1
	SQLStr=KS.GetPageSQL("KS_GuestBook","id",MaxPerPage,CurrPage,1,Param,"*")
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open SqlStr,Conn,1,1 
	If RS.Eof or RS.Bof Then 
		Response.Write "<tr class='list'><td colspan='10' align='center' height='30'>暂时还没有任何记录！</td></tr>"
	Else
	    If Param<>"" Then Param=" Where " & Param
		totalPut = Conn.Execute("Select count(id) from [KS_GuestBook] " & Param)(0)
		i = 0
		Do While Not RS.Eof 
%>
        <tr class='list' onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=RS("ID")%>' onClick="chk_iddiv('<%=RS("ID")%>')">
          <td class='splittd' align="center" valign="middle"><input onClick="chk_iddiv('<%=RS("ID")%>')" type="checkbox" id='c<%=Trim(RS("ID"))%>' name="id" value="<%=Trim(RS("ID"))%>"></td>
		 <td class='splittd'><img src="../../<%=KS.Setting(66)%>/images/common.gif" align="absmiddle">
		  
		 <% on error resume next
		   response.write "[<a href='" & KS.GetClubListUrl(rs("boardid")) & "' target='_blank'>" & conn.execute("select boardname from ks_guestboard where id=" & rs("boardid"))(0) & "</a>]"
		  %>
		 <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank">
		<%=rs("subject")%></a>
		 <%if not ks.isnul(rs("annexext")) then%>
		 <img src="../../editor/ksplus/fileicon/<%=rs("annexext")%>.gif" alt="<%=rs("annexext")%>附件" align="absmiddle">
		 <%end if%>
		 <%if rs("ispic")="1" then%>
		 <img src="../../editor/ksplus/fileicon/gif.gif" alt="gif图片附件" align="absmiddle">
		 <%elseif rs("ispic")="2" then%>
		 <img src="../../editor/ksplus/fileicon/jpg.gif" alt="jpg图片附件" align="absmiddle">
		 <%end if%>
		 <%if rs("isslide")="1" then%>
		  <font color=red>幻</font>
		 <%end if%>
		 
		 <span class="noshow">
		 <%
		  If rs("isslide")="1" then
		   response.write "<a href='?Action=Del&flag=unslide&ID=" & rs("id") & "'>取消幻灯</a>"
		  else
		   if rs("ispic")<>"0" then
		   response.write "<a href='?Action=Del&flag=slide&ID=" & rs("id") & "'>设置幻灯</a>"
		   end if
		  end if
		 %>
		 <a href="KS.GuestBook.asp?Action=Del&flag=del&ID=<%=rs("id")%>" onClick="return(confirm('所有该主题下的回复也将被删除，确定吗？'))">删除</a> | <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank">查看</a> 
		 </span>
		 
		 
		 </td>
		 <td class='splittd' style="text-align:center">
		 <%
		 if ks.isnul(rs("username")) then 
		  response.write "游客"
		 else
		  response.write rs("username")
		 end if
		 %>
		 </td>
		 <td class='splittd'  style="text-align:center">
		 <%
		  response.write RS("TotalReplay") & "/" & rs("hits")
		 %>
		 </td>
		 <td class='splittd'  style="text-align:center">
		 <%
		 if ks.isnul(RS("LastReplayUser")) then 
		  response.write "游客"
		 else
		  response.write RS("LastReplayUser")
		 end if
		 %>
		 </td>
		 <td class='splittd'  style="text-align:center">
		 <%
		  If rs("verific")=1 then
		   response.write "<a href='?Action=Del&flag=uncheck&ID=" & rs("id") & "'><font color=blue>已审</font></a>"
		  else
		   response.write "<a href='?Action=Del&flag=check&ID=" & rs("id") & "'><font color=red>未审</font></a>"
		  end if
		 %>
		 </td>

		</tr>
        <%
		i=i+1
		if i>=maxperpage then exit do
	RS.MoveNext
	Loop
	%>
	</table>
</form>

	<%
End if
RS.Close
Set RS=Nothing
%><div class="mt10"></div>
        <table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" >
          <tr>
		    <td>&nbsp;<label><input type="checkbox"  name='selectCheck' onClick="if (this.checked){Select(0)}else{Select(1)}">全部选中</label>
              <input name="delbtn" value="删除"  class="button" type="button" onClick="cdel();">
			  <input name="delbtn" value="审核" class="button" type="button" onClick="ccheck();">
	          <input name="delbtn" value="取消审核" class="button" type="button" onClick="cuncheck();">
			</td>

          </tr>
      </table>
 <%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,true,true)
%>
<br style="clear:both">

<div class="attention">
<strong>特别提醒：</strong>
只有上传图片附件的帖子才可以设置幻灯属性,建议只设置jpg格式附件的帖子为幻灯,否则可能调用不出来。
</div>
</div>
<%
 End Sub
 
 '审核回复
 Sub VerifyReply()
     Dim Table:Table=KS.G("Table")
    If KS.IsNul(Table) Then Table="KS_GuestReply"
   %>
   <div class="pageCont"><div class="pt10 pd10"><label><input type="checkbox"  name='selectCheck' onClick="if(this.checked){Select(0)}else{Select(1)}">全部选中</label>
<strong>选择数据表：</strong><select name="table" onChange="location.href='?action=VerifyReply&table='+this.value" class="h30">
   <%
 
    Dim Node,TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	Dim Url,N:N=0
    For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  If KS.S("Table")=Node.SelectSingleNode("tablename").text Then
	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"' selected>回复表(" & Node.SelectSingleNode("tablename").text&" 共" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &" where verific=0 and deltf=0")(0) &"条)</option>"
	  Else
	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"'>回复表(" & Node.SelectSingleNode("tablename").text&" 共" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &" where verific=0 and deltf=0")(0) &"条)</option>"
	  End If
	Next
	
	Dim param:Param=" verific=0 and deltf=0"
	MaxPerPage=20
	CurrPage = KS.ChkClng(Request("Page")) : If CurrPage<=0 Then CurrPage=1
	SQLStr=KS.GetPageSQL(Table,"id",MaxPerPage,CurrPage,1,Param,"*")
	If Param<>"" Then Param=" Where " & Param
	totalPut = Conn.Execute("Select count(id) from [" & Table & "] " & Param)(0)
 %>
   </select>
   
   当前正在管理的数据表：<font color=blue><%=Table%></font>,共有 <font color=red><%=totalput%></font> 条需要审核</div>
    <table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
	<form name="KS_GuestBook" action="KS.GuestBook.asp" method="post">
	<input type="hidden" name="action" id="action" value=""/>
	<input type="hidden" name="table" id="table" value="<%=table%>"/>
		<tr class="sort">
					<td>&nbsp;</td>
					<td>回复内容</td>
					<td>作者</td>
					<td>发表时间</td>
					<td>状态</td>
					<td>管理操作</td>
		</tr>
		<%
        on error resume next
		Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open SqlStr,conn,1,1
		If RS.Eof or RS.Bof Then 
		Response.Write "<tr class='list'><td colspan='10' class='splittd' align='center' height='30'>回收站中没有记录！</td></tr>"
	    Else
			i = 0
			Do While Not RS.Eof 
			%>
        <tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=RS("ID")%>' onClick="chk_iddiv('<%=RS("ID")%>')">
             <td  height="30" class='splittd' align="center" valign="middle"><input onClick="chk_iddiv('<%=RS("ID")%>')" type="checkbox" id='c<%=Trim(RS("ID"))%>' name="id" value="<%=Trim(RS("ID"))%>"></td>

		     <td class='splittd'><img src="../../<%=KS.Setting(66)%>/images/common.gif" align="absmiddle">
		      <a href="<%=KS.GetClubShowUrl(rs("topicid"))%>" target="_blank"><%=ks.gottopic(rs("content"),80)%></a>
			 </td>
		     <td class='splittd'><%=rs("username")%></td>
		     <td class='splittd'><%=formatdatetime(rs("ReplayTime"),2)%></td>
		     <td class='splittd' style="text-align:center"><%
			 if rs("verific")="1" then
			   response.write "<font color=green>已审</font>"
			 else
			   response.write "<font color=red>未审</font>"
			 end if
			 %></td>
			 <td class="splittd" nowrap style="text-align:center"><a href="?table=<%=table%>&action=DoVerifyReply&id=<%=rs("id")%>">审核</a> <a href="?action=DelReply&table=<%=table%>&id=<%=rs("id")%>" onClick="return(confirm('此操作不可逆，确定执行删除吗？'));">删除</a></td>
		    </tr>
			<%
			RS.MoveNext
			Loop
	    End If
		RS.Close:Set RS=nothing
		%>
  </table>
<table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" class="operatingBox">
          <tr>
		    <td ><label><input type="checkbox"  name='selectCheck' onClick="if(this.checked){Select(0)}else{Select(1)}">全部选中</label>

              <input name="delbtn" value=" 删除到回收站 "  class="button" type="submit" onClick="if (confirm('此操作不可逆，确定删除选中的记录到回收站吗？')){$('#action').val('DelReply');}else{return false;}">
	          <input name="delbtn" value=" 批量审核 " class="button" type="submit" onClick="$('#action').val('DoVerifyReply');">
			</td>

          </tr>
      </table>
	 </form>
	</div>
<div class="footerTable pt10">
  <%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,true,true)
%>
<div style="clear:both"></div>
</div>
  <%
 End Sub
 
 Sub DoVerifyReply()
   Dim Table:Table=KS.S("Table")
   Dim ID:ID=KS.FilterIds(KS.S("ID"))
   If ID="" Then KS.AlertHintScript "对不起，没有选择记录!"
   Conn.Execute("Update " & Table & " Set Verific=1 Where ID in(" & ID &")")
   KS.AlertHintScript "恭喜，选中的回复审核成功!"
 End Sub
 
 '删除回复
 Sub DelReply()
   Dim Table:Table=KS.S("Table")
   Dim ID:ID=KS.FilterIds(KS.S("ID"))
   If ID="" Then KS.AlertHintScript "对不起，没有选择记录!"
   Conn.Execute("Update " & Table & " Set deltf=1 Where ID in(" & ID &")")
   KS.AlertHintScript "恭喜，选中的回复删除到回收站成功!"
 End Sub
 
 Sub Recycle()
    Dim Table:Table=KS.G("Table")
    If KS.IsNul(Table) Then Table="KS_GuestBook"
   %>
   <div class="pageCont"><div class="pt10 pd10"><strong>选择数据表：</strong><select name="table" onChange="location.href='?action=Recycle&table='+this.value"  class="h30">
   <option value="KS_GuestBook">主题表(KS_GuestBook 共<%=conn.execute("select count(1) from KS_GuestBook where deltf=1")(0)%>条)</option>
   <%
 
    Dim Node,TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	Dim Url,N:N=0
    For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  If KS.S("Table")=Node.SelectSingleNode("tablename").text Then
	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"' selected>回复表(" & Node.SelectSingleNode("tablename").text&" 共" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &" where deltf=1")(0) &"条)</option>"
	  Else
	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"'>回复表(" & Node.SelectSingleNode("tablename").text&" 共" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &" where deltf=1")(0) &"条)</option>"
	  End If
	Next
	
	Dim param:Param=" DelTF=1"
	MaxPerPage=20
	CurrPage = KS.ChkClng(Request("Page")) : If CurrPage<=0 Then CurrPage=1
	SQLStr=KS.GetPageSQL(Table,"id",MaxPerPage,CurrPage,1,Param,"*")
	If Param<>"" Then Param=" Where " & Param
	totalPut = Conn.Execute("Select count(id) from [" & Table & "] " & Param)(0)
 %>
   </select>
   
   当前正在管理的数据表：<font color=blue><%=Table%></font>,共有 <font color=red><%=totalput%></font> 条</div>
 	<form name="KS_GuestBook" action="KS.GuestBook.asp" method="post">
<div class="mt10"></div>
 <table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" >
          <tr>
		    <td ><label><input type="checkbox"  name='selectCheck' onClick="if(this.checked){Select(0)}else{Select(1)}">全部选中</label>
              <input name="delbtn" value="彻底删除"  class="button" type="submit" onClick="if (confirm('此操作不可逆，确定彻底删除选中的记录吗？')){$('#action').val('DelRecycle');}else{return false;}">
              <input name="delbtn" value="一键清空"  class="button" type="submit" onClick="if (confirm('此操作不可逆，确定彻底一键清空记录吗？')){$('#action').val('DelRecycleAll');}else{return false;}">
	          <input name="delbtn" value="批量还原" class="button" type="submit" onClick="$('#action').val('Revert');">
			</td>

          </tr>
      </table>
	<input type="hidden" name="action" id="action" value=""/>
	<input type="hidden" name="table" id="table" value="<%=table%>"/>
	
 <table border="0" width="100%" align="center" cellpadding="0" cellspacing="0" class="mt10">
		<tr class="sort">
					<td>&nbsp;</td>
					<%if lcase(table)<>"ks_guestbook" Then%>
					<td>回复内容</td>
					<td>作者</td>
					<td>发表时间</td>
					<%else%>
					<td>标题</td>
					<td>版面</td>
					<td>作者</td>
					<td>最后发表</td>
				    <%end if%>
					<td>管理操作</td>
		</tr>
		<%
        on error resume next
		Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open SqlStr,conn,1,1
		If RS.Eof or RS.Bof Then 
		Response.Write "<tr class='list'><td colspan='10' align='center' height='30'>回收站中没有记录！</td></tr>"
	    Else
			i = 0
			Do While Not RS.Eof 
			%>
        <tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=RS("ID")%>' onClick="chk_iddiv('<%=RS("ID")%>')">
             <td  height="30" class='splittd' align="center" valign="middle"><input onClick="chk_iddiv('<%=RS("ID")%>')" type="checkbox" id='c<%=Trim(RS("ID"))%>' name="id" value="<%=Trim(RS("ID"))%>"></td>
			 <%if lcase(table)<>"ks_guestbook" Then%>
		     <td class='splittd'><img src="../../<%=KS.Setting(66)%>/images/common.gif" align="absmiddle">
		      <a href="<%=KS.GetClubShowUrl(rs("topicid"))%>" target="_blank"><%=ks.gottopic(rs("content"),80)%></a>
			 </td>
		     <td class='splittd'><%=rs("username")%></td>
		     <td class='splittd'><%=formatdatetime(rs("ReplayTime"),2)%></td>
			 <%else%>
		     <td class='splittd'><img src="../../<%=KS.Setting(66)%>/images/common.gif" align="absmiddle">
		      <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank"><%=KS.Gottopic(rs("subject"),38)%></a> (跟贴<font color=red> <%=rs("TotalReplay")%></font> 条)
			 </td>
			 <td class="splittd"><%response.write "<a href='" & KS.GetClubListUrl(rs("boardid")) & "' target='_blank'>" & conn.execute("select top 1 boardname from ks_guestboard where id=" & rs("boardid"))(0) & "</a>"
             %></td>
			 <td class="splittd"><a href="<%=KS.GetSpaceUrl(rs("userid"))%>" target="_blank"><%=rs("username")%></a></td>
			 <td class="splittd" style="text-align:center"><%=Formatdatetime(rs("LastReplayTime"),2)%></td>
			 <%end if%>
			 <td class="splittd" nowrap style="text-align:center"><a href="?table=<%=table%>&action=Revert&id=<%=rs("id")%>" class="setA">还原</a>|<a href="?action=DelRecycle&table=<%=table%>&id=<%=rs("id")%>" onClick="return(confirm('此操作不可逆，确定执行删除吗？'));" class="setA">删除</a></td>
		    </tr>
			<%
			RS.MoveNext
			Loop
	    End If
		RS.Close:Set RS=nothing
		%>
  </table>
  <div class="pt10"></div>
<table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" >
          <tr>
		    <td ><label><input type="checkbox"  name='selectCheck' onClick="if(this.checked){Select(0)}else{Select(1)}">全部选中</label>

              <input name="delbtn" value="彻底删除"  class="button" type="submit" onClick="if (confirm('此操作不可逆，确定彻底删除选中的记录吗？')){$('#action').val('DelRecycle');}else{return false;}">
              <input name="delbtn" value="一键清空"  class="button" type="submit" onClick="if (confirm('此操作不可逆，确定彻底一键清空记录吗？')){$('#action').val('DelRecycleAll');}else{return false;}">
	          <input name="delbtn" value="批量还原" class="button" type="submit" onClick="$('#action').val('Revert');">
			</td>

          </tr>
      </table>
	 </form>
	 </div>
<div class="footerTable pt10">
  <%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,true,true)
%>

<div style="clear:both"></div>
<div class="attention">
<strong>特别提醒：</strong>
彻底删除后，将不能恢复，慎重操作！
</div>
</div>

  <%
 End Sub
 
 '还原
 Sub Revert()
  Dim ID:ID=KS.FilterIds(KS.S("ID"))
  Dim Table:Table=KS.G("Table")
  If KS.IsNul(ID) Or Table="" Then KS.AlertHintScript "没有选择要还原的记录!"
  if Lcase(table)<>"ks_guestbook" Then
    Dim RS:Set RS=Conn.Execute("Select TopicID From " & Table &" Where id In ( "& ID & ")")
	Do While Not RS.Eof
	  Conn.Execute("Update KS_GuestBook Set TotalReplay=TotalReplay+1 Where id=" & rs(0))
	 RS.MoveNext
	Loop
	RS.Close
	Set RS=Nothing
  End If
  Conn.Execute("Update " & Table & " Set DelTF=0 Where ID In(" & ID &")")
  KS.AlertHintScript "恭喜，还原成功!"
 End Sub
 
 '一键清空
 Sub DelRecycleAll()
 Dim RS,Table:Table=KS.G("Table")
  if Lcase(table)<>"ks_guestbook" Then  '删除回复
	   Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select ID,TopicID From " & Table & " Where DelTF=1",conn,1,1
	   Do While Not RS.Eof 
		 Conn.Execute("Delete From KS_GuestComment Where Tid=" & rs(1) & " and pid=" & rs(0))
	   RS.MoveNext
	   Loop
	   RS.CLOSE:Set RS=Nothing
    Conn.Execute("Delete From " &Table & " Where DelTF=1")
	KS.AlertHintScript "恭喜，一键清除数据表" & Table & "回收站的数据成功!"
  Else
	  Dim TopicIds
	  Set RS=Conn.Execute("Select Id From KS_GuestBook Where DelTF=1")
	  Do While Not RS.Eof 
		   If TopicIDs="" Then
			 TopicIDs=RS(0)
			Else
			TopicIDs=TopicIDs & "," & RS(0)
			End If
		  RS.MoveNext
		  Loop
	   RS.Close : Set RS=Nothing
	   If TopicIds<>"" Then
		Call DoDelete(TopicIds)
	   Else
		KS.AlertHintScript "数据表" & Table & "回收站中没有记录!"
	   End If
  End If
 End Sub
 
 '彻底删除
 Sub DelRecycle()
  Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
  Dim Table:Table=KS.G("Table")
  If KS.IsNul(TopicIds) Or Table="" Then KS.AlertHintScript "没有选择要删除的记录!"
  if Lcase(table)<>"ks_guestbook" Then  '删除回复
   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select ID,TopicID From " & Table & " Where ID in("& TopicIds&")",conn,1,1
   Do While Not RS.Eof 
     Conn.Execute("Delete From KS_GuestComment Where Tid=" & rs(1) & " and pid=" & rs(0))
   RS.MoveNext
   Loop
   RS.CLOSE:Set RS=Nothing
   Conn.Execute("Delete From " &Table & " Where ID in("& TopicIds&")")
	KS.AlertHintScript "恭喜，清除数据表" & Table & "回收站的选中的数据成功!"
  Else
   Call DoDelete(TopicIds)
  End If
 End Sub
 
 Sub doDelete(TopicIds)
  Dim TodayNum:TodayNum=0
  dim boardid,postTable,userName,id,BSetting,ChannelID,InfoId
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select UserName,boardid,subject,AddTime,PostTable,ID,ChannelID,InfoID From KS_GuestBook Where ID in(" & TopicIds &")",conn,1,1
			If Not RS.Eof Then
			 Do While Not RS.Eof
				  id=RS("ID"): boardid=rs(1): postTable=rs(4):userName=rs(0)
				  ChannelID=rs("channelid"):infoid=rs("infoid")
				  If DateDiff("d",rs(3),Now)=0 Then
				   TodayNum=TodayNum+1
				  End If
				  If boardid<>0 then 
					 KS.LoadClubBoard()
					 On Error Resume Next
					 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					 Dim LastPost,LastPostArr:LastPostArr=Split(Node.SelectSingleNode("@lastpost").text,"$")
					 
					 '更新首页的最新主题
					 If KS.ChkClng(LastPostArr(0))=ID Then
					   Dim RSNew:Set RSNew=Conn.Execute("Select top 1 ID,BoardID,Subject,AddTime From KS_GuestBook Where BoardID=" & boardid & " and verific=1 and id<>" & id & " order by id desc")
					   If Not RSNew.Eof Then
						 LastPost=RSNew(0) & "$" & RSNew(3) & "$" & Replace(left(RSNew(2),200),"$","") & "$$$$$$$$"
					   Else
						 LastPost="无$无$无$$$$$$$$"
					   End If
					   Conn.Execute("Update KS_GuestBoard Set LastPost='" & LastPost & "' Where ID=" & BoardID)
					   Node.SelectSingleNode("@lastpost").text=LastPost
					 End If
				  end if
				  
				  if not KS.ISNul(rs(0)) then
				     On Error Resume Next
					 BSetting=Node.SelectSingleNode("@settings").text
					 If Not KS.IsNul(BSetting) Then
						 If KS.ChkClng(Split(BSetting,"$")(34))<>0 Then
						  Conn.Execute("Update KS_User Set Prestige=Prestige-" & KS.ChkClng(Split(BSetting,"$")(34)) & " Where UserName='" & rs(0) &"' and Prestige>0")
						 End If
					 
					   If KS.ChkClng(Split(BSetting,"$")(7))>0 Then
						Call KS.ScoreInOrOut(rs(0),2,KS.ChkClng(Split(BSetting,"$")(7)),"系统","在论坛您发表的主题[" & rs(2) & "]被删除!",0,0)
					   End If
					 End If
				  end if
				  
				  Dim Num,replyNum:replyNum=Conn.Execute("Select count(id) from " & PostTable & " where topicid=" & id)(0)
				  TodayNum=TodayNum+Conn.Execute("Select count(id) from " & PostTable & " where topicid=" & id &" and datediff(" & DataPart_D & ",ReplayTime," & SqlNowString&")=0")(0)
				  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				  Doc.async = false
				  Doc.setProperty "ServerHTTPRequest", true 
				  Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
				  Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)-TodayNum
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("todaynum").text=Num
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("postnum").text)-replyNum
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("postnum").text=Num
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("topicnum").text)-1
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("topicnum").text=Num
				  
				  Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum-" & TodayNum & " where id=" &boardid &" and todaynum>=" & TodayNum)
				  Conn.Execute("Update KS_GuestBoard Set PostNum=PostNum-" & replyNum -1& " where id=" &boardid &" and PostNum>=" & replyNum-1)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@postnum").text=Conn.Execute("Select PostNum From KS_GuestBoard Where id=" & boardid)(0)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@todaynum").text=Conn.Execute("Select TodayNum From KS_GuestBoard Where id=" & boardid)(0)
		
				  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					
				  If KS.ChkClng(ChannelID)<>0 And KS.ChkClng(InfoID)<>0 Then  '删除绑定模型数据
				    Conn.Execute("Delete From " & KS.C_S(ChannelID,2) & " Where PostID=" & ID)
				    Conn.Execute("Delete From KS_ItemInfo Where ChannelID=" & ChannelID & " and InfoID=" & InfoID)
				  End If
					
					Conn.Execute("update KS_User set postNum=postNum-1 where userName='" & UserName & "' and postNum>0")
					Conn.Execute("delete from KS_Guestbook where id=" & ID)
					Conn.Execute("Delete From KS_GuestComment Where tid=" & ID)
					Conn.Execute("delete from " & PostTable & " where TopicID=" & ID)
					Conn.Execute("delete from KS_UploadFiles where ID=" & ID & " and channelid=9994")
			  RS.MoveNext
			Loop 
			End If
			rs.close:set rs=nothing
			
			
    KS.AlertHintScript "恭喜，删除成功!"

 End Sub
 
 
 '删除帖子
 Sub GuestDel()
			Dim strIdList,arrIdList,iId,i,Flag,SqlStr
			strIdList = Trim(KS.G("ID"))
			Flag = Trim(KS.G("Flag"))
			Select Case Flag
			Case "del"
				If Not IsEmpty(strIdList) Then
				    strIdList=KS.FilterIds(strIdList)
					If strIdList<>"" Then
					    Call KS.delweibo("论坛主题",strIdList)
						Conn.Execute ("Update KS_GuestBook Set DelTF=1 WHERE ID in (" & strIdList & ")")
					End If
					Call KS.AlertHintScript("信息删除成功，确认返回！")
				Else
					Call KS.AlertHistory("请至少选择一条信息记录！",-1)
				End If
			Case "check"
				If Not IsEmpty(KS.FilterIds(strIdList)) Then
				    Dim RS,ChannelID,InfoID
					Set RS=Conn.Execute("Select * From KS_GuestBook Where ID in(" & KS.FilterIds(strIdList) & ")")
					Do While Not RS.Eof
					    ChannelID=RS("ChannelID"): InfoID=RS("InfoID")
						If ChannelID<>0 And InfoID<>0 Then
						  Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set verific=1 Where ID=" & InfoID)
						  Conn.Execute("Update KS_ItemInfo Set verific=1 Where ChannelID=" & ChannelID & " And ID=" & InfoID)
						End If
						Conn.Execute("update " & RS("PostTable") &" set verific=1 where TopicID=" & RS("ID"))
					RS.MoveNext
					Loop
					RS.Close :Set RS=Nothing
					Conn.Execute("UPDATE KS_GuestBook SET Verific = 1 WHERE ID in(" & KS.FilterIds(strIdList) & ")")
					Call KS.Alert("信息审核成功，确认返回！",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("请至少选择一条信息记录！",-1)
				End If
			Case "uncheck"
					If Not IsEmpty(KS.FilterIds(strIdList)) Then
						Set RS=Conn.Execute("Select * From KS_GuestBook Where ID in(" & KS.FilterIds(strIdList) & ")")
						Do While Not RS.Eof
							ChannelID=RS("ChannelID"): InfoID=RS("InfoID")
							If ChannelID<>0 And InfoID<>0 Then
							  Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set verific=0 Where ID=" & InfoID)
							  Conn.Execute("Update KS_ItemInfo Set verific=0 Where ChannelID=" & ChannelID & " And ID=" & InfoID)
							End If
							Conn.Execute("update " & RS("PostTable") &" set verific=0 where TopicID=" & RS("ID"))
						RS.MoveNext
						Loop
						RS.Close :Set RS=Nothing
						Conn.Execute("UPDATE KS_GuestBook SET Verific = 0 WHERE ID in(" & KS.FilterIds(strIdList) & ")")
						Call KS.Alert("取消审核成功，确认返回！",Request.ServerVariables("HTTP_REFERER"))
					Else
						Call KS.AlertHistory("请至少选择一条信息记录！",-1)
					End If
		  case "slide"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						Conn.Execute("UPDATE KS_GuestBook SET isslide = 1 WHERE ID="&iId&"")			
					Next
					Call KS.Alert("设置幻灯属性成功，确认返回！",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("请至少选择一条信息记录！",-1)
				End If
		  case "unslide"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						Conn.Execute("UPDATE KS_GuestBook SET isslide = 0 WHERE ID="&iId&"")			
					Next
					Call KS.Alert("取消幻灯属性成功，确认返回！",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("请至少选择一条信息记录！",-1)
				End If
		End Select
	End Sub
 
End Class
%>
 

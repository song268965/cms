﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New JSFreeView
KSCls.Kesion()
Set KSCls = Nothing

Class JSFreeView
        Private KS
		Private ArticleSql,ArticleRS,JSID
		Private i,totalPut,CurrentPage,MaxPerPage
		Private Sub Class_Initialize()
		   MaxPerPage=15
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
			JSID=Request.QueryString("JSID")
			if Not isempty(request("page")) and request("page")<>"" then
				  currentPage=Cint(request("page"))
			else
				  currentPage=1
			end if
			If KS.G("Action")="ArticleMoveOut" Then Call ArticleMoveOut()
			%>
			<!DOCTYPE html><html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<title>Article列表</title>
			</head>
			<link href="Admin_Style.CSS" rel="stylesheet">
			<script language="JavaScript" src="../../ks_inc/jquery.js"></script>
			<script language="JavaScript" src="../../ks_inc/common.js"></script>
			<script language="JavaScript">
			var Page='<%=CurrentPage%>';
			var JSID='<%=JSID%>';
			function ArticleMoveOut(ID)
			{
			 location.href="JSFreeView.asp?Action=ArticleMoveOut&Page="+Page+"&NewsID="+ID+"&JSID="+JSID;
			}
			function Delete(op)
			{
				 var ids=get_Ids(document.myform);

				if (ids!='')
				 {  
				  if (confirm("确定要执行删除操作吗？"))
				  ArticleMoveOut(ids);
				 }
				else 
				 alert('请选择要删除的文章');
			}
			function GetKeyDown()
			{
			if (event.ctrlKey && event.keyCode==68)
			  Delete('');
			else	
			 if (event.keyCode==46)Delete('');
			  event.cancelBubble=true;
			}
			</script>
			<body topmargin="0" leftmargin="0"  onkeydown="GetKeyDown();" onselectstart="return false;">

				  <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
				  <tr align="center" class="sort"> 
					  <td width="482" height="25"> <div align="center">文章标题</div></td>
					  <td >更新时间</td>
			  </tr>
			  <form name="myform" id="myform" action="JSFreeView.asp" method="post">
			  <%
			   ArticleSql="select * from [KS_Article] where JSID Like '%" & JSId & "%' and DelTF=0 order by AddDate desc"
			SET ArticleRS=Server.CreateObject("AdoDb.RecordSet")
			 ArticleRS.Open ArticleSql,conn,1,1 
					 IF ArticleRS.eof and ArticleRS.bof THEN
					  Response.Write("<tr colspan=3><td align=center><strong>此自由JS还没加入文章!</font></td></tr>")
					 ELSE
					        	totalPut=ArticleRS.recordcount
								if currentPage>1 and (currentPage-1)*MaxPerPage<totalPut then
										ArticleRS.move  (currentPage-1)*MaxPerPage
								end if
								Call showContent
				END IF
		%>
			</form>
			</table>
				  <table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
					  <td align="center"><hr></td>
					</tr>
					<tr>
					  <td align="right">
				        <input type="button" class="button" name="Submit" value="预览" onClick="location.href='JS_Main.asp?Action=JSView&CanView=1&JSID=<%=JSID%>';">
						<input type="button" class="button" name="Submit2" value="删除" onClick="Delete()"> 
						<input type="button" class="button" name="Submit" value="关闭" onClick="top.box.close()">
					  </td>
					</tr>
				  </table></td>
			  </tr>
			</table>
			<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)%>
			</body>
			</html>
	<%		Set ArticleRS = Nothing
			End Sub	
			
			sub showContent
				 do while not ArticleRS.eof
	%>
			 <tr class='list' id='u<%=ArticleRS("ID")%>' onclick="chk_iddiv('<%=ArticleRS("ID")%>')" onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'""> 
					<td height="20" class="splittd"> 
					 <input ltype='<%=ArticleRS("ID")%>' type='checkbox' onclick="chk_iddiv('<%=ArticleRS("ID")%>')" id='c<%=ArticleRS("ID")%>'   name='id' value='<%=ArticleRS("ID")%>'/>
							  <%IF Cint(ArticleRS("PicNews"))=1 THEN 
								 Response.Write("<img src=../Images/Folder/TheSmallPicNews1.gif border=0 align=absmiddle>") 
							   ELSE 
							   Response.Write("<img src=../Images/Folder/TheSmallWordNews1.gif border=0 align=absmiddle>")
							   END IF
							   %> 
							 <span style="cursor:default"><%=Left(ArticleRS("Title"),28)%></span> 
					 </td>
					  <td width="220" align="center" class="splittd">
						<%IF YEAR(NOW())&MONTH(NOW())&DAY(NOW())=YEAR(ArticleRS("AddDate"))&MONTH(ArticleRS("AddDate"))&DAY(ArticleRS("AddDate")) THEN Response.Write("<Font color=red>"&ArticleRS("AddDate")&"</font>") ELSE Response.Write(ArticleRS("AddDate")) END IF%>
					  </td>      
			  </tr>
			  <%i=i+1
									if i>=MaxPerPage then Exit Do
								   ArticleRS.movenext
								   loop
									ArticleRS.close
							   %>
				  <%
				  End Sub
				  
				  '移除文章
				  Sub ArticleMoveOut()
				  		Dim K, JSID, Page
						Dim ArticleRS
						Dim NewsID, FolderID
						Dim KSRObj:Set KSRObj=New Refresh
						Set ArticleRS=Server.CreateObject("ADODB.Recordset")
						Page = Trim(KS.G("Page"))
						NewsID = Split(KS.G("NewsID"), ",") '获得要移出文章的ID集合
						JSID = KS.G("JSID")
						
						For K = LBound(NewsID) To UBound(NewsID)
							 '从文章中删除此JSID
							  ArticleRS.Open "Select JSID From KS_Article Where ID=" & NewsID(K), Conn, 1, 3
								 If Not ArticleRS.EOF Then
								   ArticleRS(0) = Replace(Replace(ArticleRS(0), JSID & ",", ""),","&JSID,"")
								   ArticleRS.Update
								End If
							ArticleRS.Close
						Next
								   '刷新JS
								   Dim JSRS
								   Set JSRS = Conn.Execute("Select top 1 JSName From KS_JSFile Where JSID='" & JSID & "'")
								   Dim JSName
								   JSName = Trim(JSRS(0))
								   KSRObj.RefreshJS (JSName)
								   JSRS.Close
								   Set JSRS = Nothing
								   
						Set ArticleRS = Nothing
						Response.Redirect "JSFreeView.asp?Page=" & Page & "&JSID=" & JSID

				  End Sub
End Class
			%> 

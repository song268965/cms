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
Set KSCls = New User_Favorite
KSCls.Kesion()
Set KSCls = Nothing

Class User_Favorite
        Private KS,KSUser,action
		Private currpage,totalPut
		Private RS,MaxPerPage
		Private ChannelID,i,Param
		Private TempStr,SqlStr
		Private InfoIDArr,InfoID
		Private Sub Class_Initialize()
			MaxPerPage =10
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
		  Call KSUser.Head()
		  Call KSUser.InnerLocation("我的收藏夹")
	  	  KSUser.CheckPowerAndDie("s16")
		  action=KS.S("action")
		  If KS.S("page") <> "" Then
			currpage = CInt(KS.S("page"))
		  Else
			currpage = 1
		  End If
		  Param=" Where UserName='"& KSUser.UserName &"'"
		%>
		 <div class="tabs">						  
			<ul>
				<li<%If action="" or action="Add" then KS.Echo " class='puton'"%>><a href="User_Favorite.asp">我收藏的信息(<%=Conn.Execute("Select count(id) from KS_Favorite" & Param & " and channelid<>6")(0)%>)</a></li>
				<%if KS.Setting(56)="1" then%>
				<li<%If action="bbs" then KS.Echo " class='puton'"%>><a href="?action=bbs">我发表的帖子</a></li>
				<li<%If action="cy" Then KS.Echo " class='puton'"%>><a href="?action=cy">参与的帖子</a></li>
				<li<%If action="fav" Then KS.Echo " class='puton'"%>><a href="?action=fav">我收藏的帖子</a></li>
				<li<%If action="medal" Then KS.Echo " class='puton'"%>><a href="?action=medal">论坛勋章中心</a></li>
			   <%end if%>
			</ul>					   
			 </div>		

				<%
				Select Case action
				  Case "Add" Call AddFav()
				  Case "Cancel" call CanCel()
				  case "bbscancel"  Call FavCancel() : KS.Die ""
				  case "applyMedal" applyMedal
				  case "medal" Medal
				  case "bbs","cy" bbsinfo
				  case "fav" Fav
				  case else myFav
				End Select
			 

  End Sub
  
  
  Sub AddFav()
                   Dim RSAdd
				   InfoID=KS.ChkClng(KS.S("InfoID"))
				   ChannelID=KS.ChkClng(KS.S("ChannelID"))
				   If InfoID=0 Or Channelid=0 Then
				     KS.Die "<script>KesionJS.Alert('您没有选择要收藏的信息！');</script>"
				   End If

				   Set RSAdd=Server.CreateObject("Adodb.Recordset")
				   RSADD.Open "Select top 1 * From KS_Favorite Where ChannelID=" & ChannelID & " And InfoID=" & InfoID & " And UserName='" & KSUser.UserName & "'",Conn,1,3
				   IF RSADD.Eof And RSADD.Bof Then
				      RSADD.AddNew
					    RSAdd(1)=KSUser.UserName
						RSAdd(2)=ChannelID
						RSAdd(3)=InfoID
						RSAdd(4)=Now
					  RSAdd.Update
				   End IF
				   RSADD.Close:SET RSADD=Nothing
				   KS.Die "<script>KesionJS.Alert('恭喜，收藏成功!');</script>"
  End Sub
  
  Sub CanCel()
        InfoID=KS.S("InfoID")
		InfoID=Replace(InfoID," ","")
		InfoID=KS.FilterIDs(InfoID)
		If InfoID="" Then
					   Response.Write "<script>alert('您没有选择要取消收藏的信息！');history.back();</script>"
					   Response.End
		End If
		 Conn.Execute("Delete From KS_Favorite Where ID In(" & InfoID & ") And UserName='" & KSUser.UserName & "'")
		KS.Die "<script>alert('恭喜，取消收藏成功!');location.href='User_Favorite.asp';</script>"
  End Sub
  
  Sub MyFav()
      If ChannelID="" or not isnumeric(ChannelID) Then ChannelID=0
      IF ChannelID<>0 Then  Param= Param & " and ChannelID=" & ChannelID
	  %>
						    
				<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
				<tr height="28" class="title">
					<td height="25" align="center">选中</td>
					<td align="left">名称</td>
					<td align="center">操作</td>
				</tr>

					<%
						Set RS=Server.CreateObject("AdodB.Recordset")
						 SqlStr="Select ID,ChannelID,InfoID,AddDate From KS_Favorite "& Param &" and  Channelid<>6 order by id desc"
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td height=25 align=center colspan=5 valign=top>您的收藏夹没有内容!</td></tr>"
						 Else
									totalPut = RS.RecordCount
									If currpage < 1 Then	currpage = 1
			
								If currpage >1 and  (currpage - 1) * MaxPerPage < totalPut Then
										RS.Move (currpage - 1) * MaxPerPage
								End If
								Dim I,SQL,K
			Response.Write "<FORM Action=""User_Favorite.asp?Action=Cancel&ChannelID=" & ChannelID& "&Page=" & currpage & """ name=""myform"" method=""post"">"
			SQL=RS.GetRows(MaxPerPage)
			For K=0 To Ubound(SQL,2)
		%>
                <tr>
				   <td  class="splittd" style="height:70px;text-align:center">
				      <input id="InfoID" type="checkbox" value="<%=SQL(0,K)%>"  name="InfoID">
				     </td>
                     <td  class="splittd">
						<%
						Select Case KS.C_S(SQL(1,K),6)
						   Case 1 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 2 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 3 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 4 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 5 SqlStr="Select top 1 ID,Title,Tid,0,0,Fname,0,AddDate,hits From KS_Product Where ID=" & SQL(2,K)
						   Case 7 SqlStr="Select top 1 ID,Title,Tid,0,0,Fname,0,AddDate,hits From KS_Movie Where ID=" & SQL(2,K)
						   Case 8 SqlStr="Select top 1 ID,Title,Tid,0,0,Fname,0,AddDate,hits From KS_GQ Where ID=" & SQL(2,K)
                           Case 9 SqlStr="Select top 1 ID,Title,0,0,0,0,0,date,hits From KS_SJ Where ID=" & SQL(2,K)
						   Case else SqlStr="Select top 1 ID From KS_Article Where 1=0"
						  End Select
						  
						  Dim Url,RSF:Set RSF=Conn.Execute(SqlStr)
						  If Not RSF.Eof Then
						   If SQL(1,K)=9 then
						    url="../html/sj/" & RSF(0) & ".htm"
						   else
						    url=KS.GetItemUrl(SQL(1,K),RSF(2),RSF(0),RSF(5),RSF(7))
						   end if
						   Response.Write "<div class=""ContentTitle""> <img src=""images/fav.gif""><a href=""" & url & """ target=""_blank"">" & RSF(1) & " </a></div>"
						   Response.Write "<div class=""Contenttips"">"
						   Response.Write "<span>类型：" & KS.C_S(SQL(1,K),3) & " 收藏时间:" & KS.GetTimeFormat(SQL(3,K)) & " 信息最后更新：" & KS.GetTimeFormat(RSF(7)) & " 人气：" & RSF(8)
						  End If
											
											
											%>
                                            </span> 
											</div> 
											</td>
											
                                            <td class="splittd" align="center">
											<a class="box" href="User_Favorite.asp?Action=Cancel&Page=<%=currpage%>&InfoID=<%=SQL(0,K)%>" onclick = "return (confirm('确定取消该<%=KS.C_S(SQL(1,K),3)%>的收藏吗?'))">取消收藏</a>
											</td>
                                          </tr>

                                      <%
	  Next
			
%>
								<tr>
								  <td height="30" style="text-align:center">
								  <INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">
								  </td>
								  <td colspan="2">
								  <INPUT  class="button" onClick="return(confirm('确定取消选定的收藏吗?'));" type=submit value="取消选定的收藏" name=submit1>
								  </td>
								  
								  </FORM>
								</tr>
                                <tr>
                                  <td height="30" align='right' colspan=3>
										<%Call KS.ShowPage(totalput, MaxPerPage, "", currpage,false,true)%>
							      </td>
                                </tr>
			<%
				End If
   %>
					
          </table>
		  </td>
		  </tr>
</table>
<%
  End Sub


 sub bbsinfo()
		Call KSUser.InnerLocation("我发表的主题")
%>
		<div  class="writeblog">
		   <form action="user_favorite.asp" method="post" name="searchform">
		   <input type="hidden" name="action" value="<%=request("action")%>"/>
					主题搜索：</strong>  关键字 <input type="text" name="KeyWord" onfocus="if (this.value=='关键字'){this.value=''}" class="textbox" value="<%if request("keyword")<>"" then response.write ks.s("keyword") else response.write "关键字"%>" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" 搜 索 ">
			</form>
        </div>

		
   <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
			<tr height="28" class="title">
				<td height="25">主题</td>
				<td align="center">版块</td>
				<td align="center">回复</td>
				<td align="center">最后发表</td>
			</tr>
		<% 
		   dim 	sql
		
			dim param:param=" where username='" & ksuser.username &"'"
			if not ks.isnul(ks.s("keyword")) then param=param & " and subject like '%" & ks.s("keyword") & "%'"

		
			'取帖子存放数据表
			if request("action")="cy" then
				Dim Nodes,Doc,TableName
				set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				Doc.async = false
				Doc.setProperty "ServerHTTPRequest", true 
				Doc.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
				Set Nodes=Doc.DocumentElement.SelectSingleNode("item[@isdefault='1']")
				TableName=nodes.selectsinglenode("tablename").text
				Set Doc=Nothing
				sql="select * from KS_Guestbook where id in(select top 200 topicid from " & TableName & param &") order by LastReplayTime desc"
			else
			    sql="select * from KS_Guestbook " & param & " order by id desc"
			end if
		
			set rs=server.createobject("adodb.recordset")
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=4 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有发表过任何主题！</td>
			</tr>
		<%else
		          totalPut = RS.RecordCount
			      If CurrPage > 1  and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				  End If
				  i=0
		      do while not rs.eof
			    if i mod 2=0 then
				%>
				<tr class='tdbg'>
				<%
				else
				%>
				<tr class='tdbg trbg'>
				<%
				end if
				Dim PhotoUrl:PhotoUrl=RS("face")
		        If KS.IsNul(PhotoUrl) Then PhotoUrl=KSUser.GetUserInfo("UserFace")
				%>
					<td height="25" class="splittd">
							<div class="ContentTitle">
							<a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank" style="float:left; margin-right:8px;"><img src="<%=PhotoUrl%>" style="margin-right:3px;border:1px solid #ccc;padding:2px" onerror="this.src='../images/face/boy.jpg';" width="52" height="52" align="left"/></a>
							 <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank"><%=rs("subject")%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>发表时间:[<%=KS.GetTimeFormat1(rs("addtime"),false)%>]
							  状态:[<%if rs("verific")="1" then response.write "已审核" else response.write "未审核"%>]
							 </span>
							 </div>
					</td>
                   <td class="splittd" align="center">
							<%
							Dim Node
							KS.LoadClubBoard
			               Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & rs("boardid") &"]")
						   if not node is nothing then
						     KS.Echo "<a href='" & KS.GetClubListUrl(rs("boardid")) &"' target='_blank'>" & Node.SelectSingleNode("@boardname").text & "</a>"
						   else
						     KS.Echo "---"
						   end if
						   Set Node=Nothing
							%>
							</td>
							<td class="splittd" align=center>
							<%=RS("TotalReplay")%>
							</td>
							<td class="splittd" align=center>
							<a href='<%=KS.GetSpaceUrl(RS("LastReplayUserID"))%>' target='_blank'><%=RS("LastReplayUser")%></a>
							<div class="Contenttips"><%=KS.GetTimeFormat1(RS("LastReplayTime"),True)%></div>
							</td>
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			loop
			end if
			rs.close
			set rs=Nothing
		%>
</table>
<%
    Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,true)
   %>
			 
		<%
			if request("action")="cy" then
	  ks.echo "<div style='color:red;padding:20px 25px; font-size:14px;'><strong>说明：</strong>我参与的主题最多列出当前数据表的200条记录。</div>"
	end if

	end sub
	

	
	
	Sub Fav()
	%>
	<div class="writeblog">
		   <form action="user_favorite.asp" method="post" name="searchform">
		   <input type="hidden" name="action" value="<%=request("action")%>"/>
					主题搜索：</strong>  关键字 <input type="text" name="KeyWord" onfocus="if (this.value=='关键字'){this.value=''}" class="textbox" value="<%if request("keyword")<>"" then response.write ks.s("keyword") else response.write "关键字"%>" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" 搜 索 ">
			</form>
        </div>
			<form name="myform" action="?action=bbscancel" method="post">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
			<tr height="28" class="title">
				<td height="25" align="center">选中</td>
				<td height="25" align="center">主题</td>
				<td height="25" align="center">版块</td>
				<td width="10%" align="center">回复</td>
				<td width="15%" align="center">最后发表</td>
			</tr>
		<% 
			dim param:param=" where f.Username='"&KSUser.UserName&"'"
			if not ks.isnul(ks.s("keyword")) then param=param & " and a.subject like '%" & ks.s("keyword") & "%'"

		
			set rs=server.createobject("adodb.recordset")
			dim sql:sql="select a.*,f.favorid from KS_Guestbook a inner join KS_AskFavorite f on a.id=f.topicid " & param &" order by LastReplayTime desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有收藏问题！</td>
			</tr>
		<%else
		            totalPut = RS.RecordCount
					If CurrPage > 1  and (CurrPage - 1) * MaxPerPage < totalPut Then RS.Move (CurrPage - 1) * MaxPerPage
					i=0
		      do while not rs.eof
				if i mod 2=0 then
						%>
						<tr class='tdbg'>
						<%
						else
						%>
						<tr class='tdbg trbg'>
						<%
						end if
				%>
				            <td height="25" class="splittd" style="text-align:center"><input type="checkbox" name="favorid" value="<%=rs("favorid")%>"></td>
							<td class="splittd">
							<div class="ContentTitle">
							·<a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank"><%=rs("subject")%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>发表时间:[<%=KS.GetTimeFormat1(rs("addtime"),false)%>]
							  状态:[<%if rs("verific")="1" then response.write "已审核" else response.write "未审核"%>]
							 </span>
							 </div>
							</td>
                            <td class="splittd" align="center">
							<%
							Dim Node
							KS.LoadClubBoard
			               Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & rs("boardid") &"]")
						   if not node is nothing then
						     KS.Echo "<a href='" & KS.GetClubListUrl(rs("boardid")) &"' target='_blank'>" & Node.SelectSingleNode("@boardname").text & "</a>"
						   else
						     KS.Echo "---"
						   end if
						   Set Node=Nothing
							%>
							</td>
							<td class="splittd" align=center>
							<%=RS("TotalReplay")%>
							</td>
							<td class="splittd" align=center>
							<a href='<%=KS.GetSpaceUrl(RS("LastReplayUserID"))%>' target='_blank'><%=RS("LastReplayUser")%></a>
							<div class="Contenttips"><%=KS.GetTimeFormat1(RS("LastReplayTime"),True)%></div>
							</td>
						</tr>	
						
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			
			loop
			end if
			rs.close
			set rs=Nothing
		%>
		<tr>
		<td height="30" style="height:35px;text-align:center">
			 <INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">
		</td>
		 <td colspan="3"><input type="submit" value="取消收藏" class="button" onClick="return(confirm('确定取消收藏吗?'))"></td>
		</tr>
	 </table>
		</form>
	 <%
	     Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,true)

	 
	End Sub
		
	Sub FavCancel()
		 Dim FavorID:Favorid=KS.FilterIDS(KS.S("favorid"))
		 if FavorID="" Then KS.AlertHintScript "对不起,您没有选择记录!"
		 Conn.Execute("Delete From KS_AskFavorite Where Favorid in(" & Favorid & ") and username='" & KSUser.UserName & "'")
		 KS.AlertHintScript "恭喜，取消帖子收藏成功！"
	End Sub	
	
	Sub applyMedal()
	 dim i,mstr,medalArr,MedalID,Expression
	 medalID=KS.ChkClng(KS.G("MedalID"))
	 If MedalID=0 Then KS.AlertHintScript "出错啦！"
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "Select top 1 * From KS_GuestMedal Where MedalID=" & MedalID,conn,1,1
	 If RS.Eof And RS.Bof Then
	   RS.Close : Set RS=Nothing
	   KS.AlertHIntScript "对不起，传递参数有误！"
	 End If
	 Dim LQFs,GradeID,medalname
	 Lqfs=rs("Lqfs")
	 GradeID=rs("GradeID")
	 medalname=rs("medalname")
	 Expression=split(rs("Expression")&",0,0,0,0,0,0,0,0,0,",",")
	 mstr=rs("medalid") &"|" & rs("medalname") & "|" & rs("ico")
	 RS.Close :Set RS=Nothing
	 If Lqfs="1" Then
		 If Not KS.IsNul(GradeID) Then
		   If KS.FoundInArr(gradeid,KSUser.GetUserInfo("gradeid"),",")=false Then
			 KS.AlertHintScript "对不起，您所以的论坛级别不够，申请失败！"
		   end if
		 End If
		 If KS.ChkClng(Expression(0))>0 And KS.ChkClng(KSUser.GetUserInfo("PostNum"))<KS.ChkClng(Expression(0)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少要求发帖量大于等于" & Expression(0) &"帖,您当前发了" & KSUser.GetUserInfo("PostNum") & "帖！"
		 End If
		 If KS.ChkClng(Expression(1))>0 And KS.ChkClng(KSUser.GetUserInfo("BestTopicNum"))<KS.ChkClng(Expression(1)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少要求精华帖大于等于" & Expression(1) &"帖,您当前精华帖子" & KSUser.GetUserInfo("BestTopicNum") & "帖！"
		 End If
		 If KS.ChkClng(Expression(2))>0 And KS.ChkClng(conn.execute("select count(1) from ks_guestbook where username='" & ksuser.username &"'")(0))<KS.ChkClng(Expression(2)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少要求主题帖大于等于" & Expression(2) &"帖,您当前主题帖子" & KS.ChkClng(conn.execute("select count(1) from ks_guestbook where username='" & ksuser.username &"'")(0)) & "帖！"
		 End If
		 If KS.ChkClng(Expression(3))>0 And KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(Expression(3)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少积分大于等于" & Expression(3) &"分,您当前积分" & KSUser.GetUserInfo("score") & "分！"
		 End If
		 If KS.ChkClng(Expression(4))>0 And KS.ChkClng(KSUser.GetUserInfo("Prestige"))<KS.ChkClng(Expression(4)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少威望大于等于" & Expression(4) &"分,您当前威望" & KSUser.GetUserInfo("Prestige") & "分！"
		 End If
		 If KS.ChkClng(Expression(5))>0 And KS.ChkClng(KSUser.GetUserInfo("money"))<KS.ChkClng(Expression(5)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少资金大于等于" & Expression(4) &"元,您当前资金" & KSUser.GetUserInfo("Money") & "元！"
		 End If
		 If KS.ChkClng(Expression(6))>0 And KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(Expression(6)) Then
			 KS.AlertHintScript "对不起，申请该勋章至少点券大于等于" & Expression(6) &"点,您当前点券" & KSUser.GetUserInfo("Money") & "点！"
		 End If
	 ElseIf Lqfs="2" Then '积分购买
	   If KS.ChkClng(Expression(7))>0 Then
	     If KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(Expression(7)) Then
		    KS.AlertHintScript "对不起，您的积分不够，本枚勋章需要花 " & KS.ChkClng(Expression(7)) & " 分积分，您当前可用积分为 " & KS.ChkClng(KSUser.GetUserInfo("score")) & " 分!"
		 Else
		    Session("ScoreHasUse")="+" '设置只累计消费积分
		 	Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(Expression(7)),"系统","购买论坛勋章[" & medalname & "]消费!",0,0)

		 End If
	   Else
	    KS.AlertHIntScript "停止购买！"
	   End If
	 Else 
	   KS.AlertHIntScript "出错！"
	 End If
	 
	 Dim newMedalStr,MyMedal:MyMedal=KSUser.GetUserInfo("medal")
	 If Not KS.IsNul(MyMedal) Then
	   medalArr=split(MyMedal,"@@@")
	   for i=0 to ubound(medalArr)
	     if split(medalArr(i),"|")(0)<>medalid then
		   if newMedalStr="" then
		   newMedalStr=medalArr(i)
		   else
		    newmedalStr=newmedalStr & "@@@" & medalArr(i)
		   end if
		 end if
	   next
	 End If
	 if newmedalStr="" then
	   newmedalStr=mstr
	 else
	   newmedalStr=newmedalStr & "@@@" & mstr
	 end if
	If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@medal").Text=newmedalStr
	 Conn.Execute("Update KS_User Set Medal='" & newmedalStr & "' where username='" & KSUser.UserName &"'")
	 If Lqfs="1" Then
	  KS.AlertHintScript "恭喜，勋章申请成功！！！"
	 Else
	  KS.AlertHintScript "恭喜，勋章购买成功！！！"
	 End If
	End Sub
	
	Sub Medal()
	 Call KSUser.InnerLocation("勋章中心")
	 Dim i,medalArr,MyMedal,MedalIds
	 MyMedal=KSUser.GetUserInfo("medal")
	%>
	<style type="text/css">
	 .medallist{margin:20px 0;}
	 .medallist li{width:25%;float:left;text-align:center; font-size:14px; margin-bottom:20px;}
	 .medallist .h{height:130px}
	 .normal{color:#999;font-weight:normal}
	</style>
	<script src="../ks_inc/jquery.imagePreview.1.0.js"></script>
	 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
		<tr height="28" class="title">
				<td height="25">我的勋章</td>
				<td class="normal"><%if KS.IsNul(myMedal) Then
		     response.write "您拥有 0 枚勋章!"
			else
			  medalArr=split(mymedal,"@@@")
			 response.write "您拥有 <font color=#ff6600>" & ubound(medalArr)+1 & "</font> 枚勋章!"
			end if
			
		  %></td>
	    </tr>
		<tr>
		 <td class="splittd" colspan="2">
		  <div class="medallist">
		   <ul>
		  <%if isArray(medalArr) Then
		    for i=0 to ubound(medalArr)
			  MedalIds=MedalIds & split(medalArr(i),"|")(0) & ","
			  response.write "<li><img src='../" & KS.Setting(66) & "/images/medal/" & split(medalArr(i),"|")(2) &"'><br/>" & split(medalArr(i),"|")(1) & "</li>"
			next
			else
			  response.write "<li>您没有勋章!</li>"
			end if
		  %>
		  </ul>
		  </div>
		 </td>
		</tr>
		<tr height="28" class="title">
				<td height="25">全部勋章</td>
				<td class="normal">以下列出本站的全部勋章，带申请的勋章您可以申请拥有。</td>
	    </tr>
		<tr>
		 <td class="splittd" colspan="2">
		  <div class="medallist">
		   <ul>
		  <%
		  dim rs:set rs=conn.execute("select medalid,medalname,ico,descript,LQFS,Expression From KS_GuestMedal Where status=1 order by medalid")
		  Do While Not RS.Eof
			  response.write "<li class=""h""><a target='_blank' title='" & rs("descript") & "' href='../" & KS.Setting(66) & "/images/medal/" & rs("ico") &"' class='preview'><img width='30' src='../" & KS.Setting(66) & "/images/medal/" & rs("ico") &"'></a><br/><strong>" & rs("medalname") & "</strong><br/><br>"
			 if KS.FoundInArr(MedalIds,rs("medalid"),",") Then
			    response.write "<div><input type='button' value='已拥有√' disabled></div>"
			 Else
			  if rs("lqfs")="1" then
			    response.write "<div><form action='?' method='post'><input type='hidden' name='medalid' value='" & rs("medalid") & "'/><input type='hidden' name='action' value='applyMedal'/><input type='submit' value=' 申 请 ' class='button'></form></div>"
			  elseif rs("lqfs")="2" then
			    response.write "<div><form action='?' method='post'><input type='hidden' name='medalid' value='" & rs("medalid") & "'/><input type='hidden' name='action' value='applyMedal'/><input type='submit' value=' 购买（花" & split(rs("Expression"),",")(7) &" 积分） ' class='button'></form></div>"
			  else
			    response.write "<div><input type='button' value='人工授予'  class='button' disabled></div>"
			  end if
			 End If
			  response.write "</li>"
			  rs.movenext
		  Loop
		  RS.Close
		  Set RS=Nothing
		  %>
		  </ul>
		  </div>
		 </td>
		</tr>
	</table>
	<%
	End Sub
End Class
%> 

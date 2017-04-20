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
Set KSCls = New User_userform
KSCls.Kesion()
Set KSCls = Nothing

Class User_userform
        Private KS,KSUser
		Private totalPut,TotalPages,SQL
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
		Call KSUser.InnerLocation("查看表单")
       if ks.s("action")="delitem" then
	     delitem
	   end if
		%>
		<script type="text/javascript">
				function shownote(txt){ 
					var notestr=$("#" + txt).html()
					$.dialog({id: 'dshow',title:'查看回复',content: notestr});
				}
				function winclose(){
					var list = $.dialog.list;
					for( var i in list ){
						list[i].close();
					}
				}
		 </script>
		<%
		dim str,i,ii,rsj,form_id
		i=1:form_id=0
		Set RS=Server.CreateObject("ADODB.Recordset")
		RS.Open "Select FormName,PostByStep,TableName,Template,Templ_url,id From KS_Form where status=1 and AllowShowOnUser=1 ",conn,1,1
		if rs.eof and rs.bof then
		 rs.close:set rs=nothing
		  Response.Write "<div class=""tabs"">没有可查看的表单项！</div>"
		else
		Response.Write "<div class=""tabs""><ul>"
		while not rs.eof
		    if ks.chkclng(ks.g("id"))=rs("id") or (i=1 and request("id")="") then
			form_id=KS.ChkClng(rs("id"))	
			Response.Write "<li class=""puton"" style=""display:block; cursor:pointer;"" ><a href=""?id="& rs("id") &"""> "&rs(0)&"</a></li>"
			else
			Response.Write "<li style=""display:block; cursor:pointer;"" ><a href=""?id="& rs("id") &"""> "&rs(0)&"</a></li>"
			end if
			rs.movenext
			i=i+1
		wend
		Response.Write "</ul></div>"
		rs.close
		end if
		
		if form_id>0 then
		RS.Open "Select top 1  FormName,PostByStep,TableName,Template,Templ_url,ID From KS_Form where id="&form_id,conn,1,1
		i=1
		while not rs.eof
			
			%>
			<div class="writeblog"><input name="" class="button" type="button" onclick="$('.titlename').find('td').show();$('.tdbg').find('td').show();" value="显示所有字段" />
			</div>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1"  class="border"  id="c0<%=i%>" <%if i>=2 then Response.Write("style='display:none'")%> >
				 
				 <tr class="titlename">
					<%
					Set RSj=Server.CreateObject("ADODB.Recordset")
					rsj.open "Select Title,FieldName From KS_FormField Where ShowOnManage=1 And ItemID=" & rs("ID") & " Order By OrderID,FieldID",Conn,1,1
					If Not rsj.Eof Then SQL=rsj.GetRows(-1)
					ii=1
					If IsArray(SQL) Then
						For ii=0 To Ubound(SQL,2)
						 if ii<=4 then
						  	Response.Write "<td height=""25""    width=""100"" align='center'>"& SQL(0,II) &"</td>"
						  else
						 	Response.Write "<td height=""25""   width=""100""  align='center' style=""display:none"">"& SQL(0,II) &"</td>"
						  end if
						Next
					End If
					rsj.close
					
					%>	
				   <td width="70" height="25" align="center"> <strong>状态</strong></td>
				   <td width="70"align="center"> <strong>回复</strong></td>
				   <td align="center"> <strong>操作</strong></td>
				   </tr>
				  
				   
				   <%
					dim CurrentPage:CurrentPage=KS.ChkClng(request("page"))
					if CurrentPage<=0 then CurrentPage=1
					Set RSj=Server.CreateObject("ADODB.Recordset")
					rsj.open "select * from " & rs("TableName") & " where username='" &KSUser.UserName & "' order by adddate desc" ,conn,1,1
					ii=1
					if not rsj.eof then
					                totalPut = rsj.RecordCount
									If CurrentPage < 1 Then	CurrentPage = 1
									If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
											rsj.Move (CurrentPage - 1) * MaxPerPage
									End If
					do while not rsj.eof
					%>  
					<tr class='tdbg'>
						
						<%dim fi:fi=0
						 If IsArray(SQL) Then
							For I=0 To Ubound(SQL,2)
								if i<=4 then
									response.write "<td class='splittd'  width=""100"">&nbsp;" & rsj(trim(sql(1,i))) & "</td>"
								else
									response.write "<td class='splittd'  width=""100"" style=""display:none"">&nbsp;" & rsj(trim(sql(1,i))) & "</td>"
								end if
							Next
						End If
						%>
						
						<td class='splittd' height="25" align="center">
						<%If rsj("status")=0 then%>
						<font color="red">未审核</font>
						<%else%>
						<font color=green>已审核</font>
						<%end if%>
						</td>
						
						<td class='splittd' align="center">
						<%if ks.isnul(rsj("note")) then%>
						<font color="red">未回复</font>
						<%else%>
						<a target="_blank" href="../Plus/form/content.asp?FormID=<%=rs("ID")%>&id=<%=rsj("id")%>">
						<font color=green>有回复</font>
						</a>
						<!--<a href="#" onclick="shownote('note_<%=ii%>');"><font color="#0099CC">查看回复</font></a>-->
						<%end if%>
						</td>
						<td class='splittd' align="center">
						<a target="_blank" href="../Plus/form/content.asp?FormID=<%=rs("ID")%>&id=<%=rsj("id")%>">查看详情</a>
						<a href="?action=delitem&FormID=<%=form_id%>&id=<%=rsj("id")%>" onclick="return(confirm('此操作不可逆，确定删除吗？'));">删除</a></td>
						
					</tr>
					<%
					 if ii>= maxperpage then exit do
				    	ii=ii+1
					rsj.movenext
					loop
					else
						Response.Write "<table width=""99%""><tr class='tdbg'><td height=""25""   align=""center""><strong>没有记录</strong></td></tr></table>"
					end if
					rsj.close
					%>	
				  
			</table>
			<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
			<%
			rs.movenext
			i=i+1
		wend
		rs.close
	 end if
  End Sub
    
	sub delitem()
	 dim formid:formid=ks.chkclng(request("formid"))
	 dim id:id=ks.chkclng(request("id"))
	 if id=0 or formid=0 then ks.alerthintscript "参数出错!"
	 dim rs:Set RS=Server.CreateObject("ADODB.Recordset")
	 RS.Open "Select TableName,delform From KS_Form where id=" & formid,conn,1,1
	 if rs.eof and rs.bof then
	   rs.close
	   set rs=nothing
	   ks.alerthintscript "参数出错!"
	 end if
	 dim tablename:tablename=rs(0)
	 dim delform:delform=KS.ChkClng(rs(1))
	 rs.close
	 set rs=nothing
	 if delform=0 then
	  %>
	 <script language=JavaScript>
$.dialog.alert('对不起，系统设置不能删除!',function(){location.replace('<%=Request.ServerVariables("HTTP_REFERER")%>');
});</script>
     <%
	 else
	 conn.execute("delete from " & tablename & " where id=" & id & " and username='" & ksuser.username& "'")
	 end if
	 %>
	 <script language=JavaScript>
$.dialog.alert('恭喜，删除成功!',function(){location.replace('<%=Request.ServerVariables("HTTP_REFERER")%>');
});</script>
     <%
	
	end sub

  
End Class
%> 

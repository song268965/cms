<!--#include file="Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
'-----------------------------------------------------------------------------------------------
'科汛网站管理系统,通用函数类
'开发:林文仲 版本 v9.0
'-----------------------------------------------------------------------------------------------
Class MediaCls
	  Private Sub Class_Initialize()
      End Sub
	 Private Sub Class_Terminate()
	 End Sub
	 '歌曲搜索
	 Sub GetSearchForm()
	 %>
	       <table border="0" width="250" cellpadding="0" cellspacing="0" align="center">
           <form name="form1"  method="get" action="KS.MusicSearch.asp" target="main">
		   <tr>
              <td><select name="stype" class="textbox" >  
              <option value="Music" selected>歌曲名称</option>
              <option value="Special">专辑名称</option>
              <option value="Singer">歌手姓名</option> 
            </select>
              </td>
              <td>
            <input size=12 type="text" class="textbox" value="输入关键字" onClick="Javascript:this.value=''" onfocus=this.select() onmouseover=this.focus() name=keyword  maxlength="30"> 
              </td>
              <td>
              <input type="SUBMIT" VALUE=" 查 询 " class="button" name="SUBMIT1">
			  </td>
            </tr>
			</form>
          </table>
	 <%
	 End Sub
	 '分类导航
	Sub  GetSongTypeList(Action,ClassID,SclassID,NclassID)
		  Dim RS,Sql
		  %> 一级栏目: <select name="classid" size="1" onChange="location.href='?<%=Action%>classid='+this.options[this.selectedIndex].value;">
				<option value="" <%if ClassID="" then%> selected<%end if%>>选择栏目</option>
			<%
			set rs=server.createobject("adodb.recordset")
			sql="select * from KS_MSClass"
			rs.open sql,conn,1,1
			do while not rs.eof
			%>
							<option<%if cstr(ClassID)=cstr(rs("classid")) and ClassID<>"" then%> selected<%end if%> value="<%=CStr(rs("classID"))%>" name=classid><%=rs("class")%></option>
			<%
			rs.movenext
			loop
			rs.close
			%>
						  </select>
						  二级栏目:
			<%if ClassID<>"" then%>
						  <select name="sclassid" size="1" onChange="location.href='?<%=Action%>classid=<%=ClassID%>&sclassid='+this.options[this.selectedIndex].value;">
							<option value="" <%if SclassID="" then%> selected<%end if%>>选择栏目</option>
			<%
				sql="select * from KS_MSSClass where classid="&ClassID
				rs.open sql,conn,1,1
				Do while not rs.eof
			%>
							<option<%if cstr(SclassID)=cstr(rs("sclassid")) and SclassID<>"" then%> selected<%end if%> value="<%=CStr(rs("sclassid"))%>" name=sclassid><%=rs("sclass")%></option>
			<%
				rs.MoveNext
				Loop
				rs.close
			%>
			<%else%>
						  <select name="sclassid" size="1">
							<option value="" selected>选择栏目</option>
			<%end if%>
						  </select>
						  三级栏目:
			<%if SclassID<>"" then%>
						  <select name="Nclassid" size="1" onChange="location.href='?<%=Action%>classid=<%=ClassID%>&SClassid=<%=SclassID%>&nclassid='+this.options[this.selectedIndex].value;">
							<option value="" <%if NclassID="" then%> selected<%end if%>>选择栏目</option>
			<%
				sql="select * From KS_MSSinger where Sclassid="&SclassID
				rs.open sql,conn,1,1
				Do while not rs.eof
			%>
							<option<%if cstr(NclassID)=cstr(rs("Nclassid")) and NclassID<>"" then%> selected<%end if%> value="<%=CStr(rs("Nclassid"))%>" name=Nclassid><%=rs("Nclass")%></option>
			
			<%
				rs.MoveNext
				Loop
				rs.close
			%>
			<%else%>
						  <select name="Nclassid" size="1">
							<option value="" selected>选择栏目</option>
			<%end if%>
						  </select>
	<%					 
		End Sub
	   '媒体服务器
	   '参数 TypeID--服务类型 1音乐，2影视 ,SelID--已选ID
	   Sub GetMediaServer(Typeid,SelID)
	     IF TypeID="" Or Not ISnumeric(TypeID) Then Exit Sub
	     Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
		 Dim SqlStr
		  IF SelID="" Or Not ISNUMERIC(SelID) Then SelID=0
		   SqlStr="Select ID,MC From KS_MediaServer Where TypeID=" & TypeID
		 
		 RS.Open SqlStr,Conn,1,1
		   Response.Write "<Select Name=""ServerID"">"
		   Response.Write "<option value=""0"">-不使用服务器地址-</option>"
		   If SelID=9999 Then
		   Response.Write "<option value=""9999"" style='color:red' selected>-外部服务器-</option>"
		   Else
		   Response.Write "<option value=""9999"" style='color:red'>-外部服务器-</option>"
		   End If
		 IF Not RS.EOF Then
		   Do While Not RS.Eof 
		     IF Cint(SelID)=RS(0) Then
			   Response.Write "<option value=""" & RS(0) & """ selected>" & rs(1) & "</option>"
			 Else
			   Response.Write "<option value=""" & RS(0) & """>" & rs(1) & "</option>"
			 End IF
		    RS.MoveNext
		   Loop
		 End IF
		  Response.Write "</select>"
		  RS.Close
		Set RS=Nothing
	   End Sub
	   
End Class
%> 

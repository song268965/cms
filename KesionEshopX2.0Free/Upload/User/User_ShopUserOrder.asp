<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New User_ItemSign
KSCls.Kesion()
Set KSCls = Nothing

Class User_ItemSign
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr
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
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		  
		
		dim Action,Rs,Content,qdxq
			Action=KS.G("Action")
			select case Action
			case "Add","Edit"
				call S_AddEdit()
			case "Save"
				call S_save()
			case "Del"
				call S_Del()	
			case else
				call qdMain()
			end select
		
		
  End Sub
      sub Main_menu()
	  	%>
		<div class="tabs">	
			<ul>
				<li><a href="user_order.asp">我的订单</a></li>
				<li><a href="user_order.asp?action=coupon">我的优惠券</a></li>
				<li class="puton" ><a href="User_ShopUserOrder.asp">我的收货信息</a></li>
			</ul>
        </div>
		<%
	  end sub
	  sub qdMain()
		 Call KSUser.Head()
		 Call KSUser.InnerLocation("我的收货信息")
		 If KS.ChkClng(KS.S("page")) <> 0 Then
				CurrentPage = CInt(KS.S("page"))
		 Else
				CurrentPage = 1
		 End If
		 Call Main_menu
	    %>
		
		<div class="writeblog"><img src="images/icon1.png" align="absmiddle"><a href="User_ShopUserOrder.asp?Action=Add">增加收货信息</a></div>
			
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
					<tr class="title" align=middle>
					  <td width=100 height="25">收货人姓名</td>
					   <td width=300>收货人地址</td>
					  <td width=100>收货人邮编</td>
					  <td width=150>手机号码</td>
					  <td>操作</td>
					</tr>
					<%  
						 SqlStr="Select * From KS_ShopUserOrder where username='"& ksuser.username &"' order By id desc"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1
						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>找不到您要的记录!</td></tr>"
								 Else
								 totalPut = RS.RecordCount
								If CurrentPage < 1 Then
										CurrentPage = 1
								End If
								
			
								If CurrentPage = 1 Then
									Call ShowContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowContent
									Else
										CurrentPage = 1
										Call ShowContent
									End If
								End If
				End If

						
						 %>
					
          </table>
		  </td>
		  </tr>
</table>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		<%
	  end sub
	  
	 Sub ShowContent()
		 Dim I,intotalmoney,outtotalmoney,Page_s,qdxq,RSkc,ContactMan,Address,ZipCode,Mobile
		 Page_s=(CurrentPage-1)* MaxPerPage
		 Do While Not rs.eof 
		%>
		<tr class="tdbg">
		  
		  <td  class="splittd" align=middle><%=rs("ContactMan")%></td>
		  <td  class="splittd" align=middle ><%=rs("Address")%></td>
		  <td  class="splittd" align=middle ><%=rs("ZipCode")%></td>
		  <td  class="splittd" align=middle ><%=rs("Mobile")%></td>
		  <td class="splittd" align=middle>
			<a href="User_ShopUserOrder.asp?Action=Edit&id=<%=rs("id")%>&page=<%=KS.S("page")%>" >修改</a>
			<a href="javascript:;" onclick = "$.dialog.confirm('确定删除信息吗?',function(){location.href='User_ShopUserOrder.asp?Action=Del&id=<%=rs("id")%>&page=<%=KS.S("page")%>';},function(){})" >删除</a>
		   </td>
		</tr>
		<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

		 loop
		%>
  
		<%
		End Sub
		
		 Sub S_AddEdit()
		 	Call Main_menu
			dim RS,ContactMan,Address,ZipCode,Mobile,Phone,Email,QQ
			if KS.G("action")="Edit" then
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select top 1 * From KS_ShopUserOrder where  username='" & KSUser.UserName &"' and id="& KS.ChkClng(KS.G("id")) ,conn,1,1
				If not RS.EOF And  not RS.BOF Then
					ContactMan=rs("ContactMan")
					Address=rs("Address")
					ZipCode=rs("ZipCode")
					Mobile=rs("Mobile")
					Email=rs("Email")
					QQ=rs("QQ")
					Phone=rs("Phone")
				end if
				Rs.close: Set RS = Nothing
			end if
			
			%>
			<script type="text/javascript">
			  function CheckForm() 
				{ 
				  if ($("#ContactMan").val()==""){
				   alert('请输入收货人姓名!');
				   $("#ContactMan").focus();
				   return false;}
				  if ($("#Address").val()==""){
				   alert('请输入收货人地址!');
				   $("#Address").focus();
				   return false; }
				  if ($("#ZipCode").val()==""){
				   alert('请输入收货人邮编!');
				   $("#ZipCode").focus();
				   return false; }
				  if ($("#Phone").val()==""&&$("#Mobile").val()==""){
				   alert('收货人手机或电话至少要填一个!');
				   $("#Mobile").focus();
				   return false;}
				  if ($("#Email").val()==""){
				   alert('请输入收货人邮箱!');
				   $("#Email").focus();
				   return false;}
				  if ($("#mustyf").val()==1){
					if ($("#tocity").val()==''){
				   alert('请选择送货地区!');
				   return false;}
				  }
				  $("#myform").submit();
				}
			</script>
			  
			  <FORM name="myform" id="myform"  action="User_ShopUserOrder.asp?action=Save" method="post">
			  <input type="hidden" name="AddEdit" value="<%=KS.G("action")%>" />
			  <input type="hidden" name="id" value="<%=KS.G("id")%>" />
			  <input type="hidden" name="page" value="<%=KS.S("page")%>" />
			  <table cellSpacing=1 cellPadding=3 class="border" align=center border=0 style="margin:10px;">
                          <tr>
                              <td colSpan=2 class="titleinput">请填写收货信息</td>
                          </tr> 
                          <tr class=tdbg>
                              <td align=right width=100>收货人姓名：</td>
                              <td><INPUT class="textbox"  maxLength=50 value="<%=ContactMan%>" name="ContactMan" id="ContactMan">* </td>
                          </tr>
                          <tr class=tdbg>
                             <td align=right width=100>收货人地址：</td>
                             <td><INPUT class="textbox" maxLength=255 size=60 value="<%=Address%>" name="Address" id="Address">*</td>
                           </tr>
                           <tr class=tdbg>
                             <td align=right width=100>收货人邮编：</td>
                             <td height=20><INPUT class="textbox" maxLength=6 value="<%=ZipCode%>" name="ZipCode" id="ZipCode">* </td>
                           </tr>
                           <tr class=tdbg>
                             <td align=right width=100>手机号码：</td>
                             <td><INPUT class="textbox" maxLength=50 size=20 value="<%=Mobile%>" name="Mobile" id="Mobile">
                                                或固定电话<INPUT maxLength=50 class="textbox" size=20 value="<%=Phone%>" name="Phone" id="Phone">*<span class='tips'>两个至少填一个</span> </td>
                            </tr>
                            <tr class=tdbg>
                              <td align=right width=100>收货人邮箱：</td>
                              <td height=20><INPUT class="textbox" maxLength=100 size=30 value="<%=Email%>" name=Email id="Email">*</td>
                            </tr>
                            <tr class=tdbg>
                               <td align=right>收货人QQ：</td>
                               <td><INPUT class="textbox" maxLength=50 size=30 value="<%=QQ%>" name="QQ" id="QQ"></td>
                            </tr>
				 </table>
				 <div align="center"><button class="pn" type="button" onClick="return CheckForm();" name="Submit"><strong> OK,保存信息 </strong></button></div>
				 </FORM>
			<%
		 End Sub
		 Sub S_save ()
			dim RS
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_ShopUserOrder where username='" & KSUser.UserName &"' and id="& KS.ChkClng(KS.G("id")) ,conn,1,3
			If RS.EOF And  RS.BOF Then rs.addnew
			rs("username")=ksuser.username
			rs("ContactMan")=KS.G("ContactMan")
			rs("Address")=KS.G("Address")
			rs("ZipCode")=KS.G("ZipCode")
			rs("Mobile")=KS.G("Mobile")
			rs("Email")=KS.G("Email")
			rs("QQ")=KS.G("QQ")
			rs("Phone")=KS.G("Phone")
			Rs.update
			Rs.close: Set RS = Nothing
			Response.Write "<script>$.dialog.tips('恭喜，信息修改成功！',1,'success.gif',function(){location.href='User_ShopUserOrder.asp?page="& KS.S("page") &"';});</script>"
		 End Sub
		 Sub S_Del()
		 	if KS.ChkClng(KS.G("id"))<>0 then
				Conn.Execute("Delete from KS_ShopUserOrder where username='" & KSUser.UserName &"' and id ="&KS.ChkClng(KS.G("id")) )
				Response.Write "<script>location.href='User_ShopUserOrder.asp?page="& KS.S("page") &"'</script>"
			end if
		 End sub
End Class
%> 

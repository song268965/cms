<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Brand
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Brand
        Private KS,Action,ComeUrl,Page,ItemName,Table,ClassID
		Private I,totalPut,CurrentPage,KeySql,RS,MaxPerPage,KSCls
		Private Sub Class_Initialize()
		  MaxPerPage =18
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
			.Write "<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
			.Write "<title>收货人信息管理</title>"
			.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>" & vbCrLf
	        .Write "<script language=""JavaScript"" src=""../../KS_Inc/Jquery.js""></script>" & vbCrLf
             Action=KS.G("Action")
			   If Not KS.ReturnPowerResult(5, "M520075") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End if
			
			CurrentPage = KS.ChkClng(KS.G("page"))
			If CurrentPage<1 Then CurrentPage = 1

			 ItemName=KS.C_S(5,3)
			 Page=KS.G("Page")
			 ClassID=KS.G("ClassID")
			 
			 Select Case Action
			  case "Add","Edit"
					call S_AddEdit()
			  case "Save"
					call S_save()
			  case "Del"
					call S_Del()	
			  Case Else
			   Call S_Main()
			 End Select
			.Write "</body>"
			.Write "</html>"
		  End With
		End Sub

		Sub S_Main()			
			With Response
			.Write "<body topmargin='0' leftmargin='0'  onkeydown='GetKeyDown();' onselectstart='return false;'>"
			%>


			<div class='tableTop mt20' style='text-align:left'>
			<form action="?" name="myform" method="post" >
			   <table>
				   <tr>
					   <td><strong>按用户搜索=></strong><span class="tiaoJian">用户名:</span><input type="text" class='textbox' name="keyword">
						  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
					   </td>
				   </tr>
			   </table>
			</form>
			</div>
			<%
			If Not KS.IsNul(Request("keyword")) Then
			 .write "<div style='text-align:left;height:25px;line-height:25px'>搜索关键词“<font color=red>" & KS.G("keyword") & "</font>”搜索结果:</div>"
			End If
			.Write "<div class='pageCont2 mt20'><table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.Write ("<form name='myform' method='Post' action='?'>")
	        .Write ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
			.Write "        <tr class='sort'>"
			.Write "          <td class='sort' align='center' width='120'>用户名</td>"
			.Write "          <td class='sort' align='center' width='120'>收货人姓名</td>"
			.Write "          <td class='sort' align='center' width='450'>收货人地址</td>"
			.Write "          <td class='sort' align='center' width='120'>收货人邮编</td>"
			.Write "          <td align='center' class='sort' width='120'>手机号码</td>"
			.Write "          <td align='center' class='sort' width='120'>操作</td>"
			.Write "  </tr>"
			  
			  Set RS = Server.CreateObject("ADODB.RecordSet")
			  Dim Param:Param=" Where 1=1"
			  If Request("keyword")<>"" Then
			    Param=Param & " and username='"& KS.G("keyword") &"' "
			  End If
			  KeySql="Select * From KS_ShopUserOrder " & Param 
			  RS.Open KeySql, conn, 1, 1
			  If Not RS.EOF Then
						totalPut = RS.RecordCount
						If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
						End If
						Call showContent
			  Else
			   .Write "<tr><td colspan='8' class='splittd' style='text-align:center'>没有记录！</td></tr>"
			  End If
			.Write "</table>"
			.Write ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .Write ("<tr>")
	        .Write ("</form><td align='right'>")
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .Write ("</td></tr></table></div>")
            End With
			End Sub
			
			Sub showContent()
			   With Response
					Do While Not RS.EOF
					  .Write "<td class='splittd' height='20' align='center'><i class='icon manage'></i>"
					  .Write "  <span style='cursor:default;'>" & RS("username") & "</span></td>"
					   .Write "  <td class='splittd' align='center'>" &RS("ContactMan") &"</td>"
					  .Write "  <td class='splittd' align='center'>" &rs("Address") &"</td>"
					  .Write "  <td class='splittd' align='center'> "& rs("ZipCode") &"</td>"
					  .Write "  <td class='splittd' align='center'> "& rs("Mobile") &"</td>"
					.Write "  <td class='splittd' align='center'> "
					 %>
					 	<a href="KS.ShopUserOrder.asp?Action=Edit&id=<%=rs("id")%>&page=<%=CurrentPage%>" class="setA">修改</a>|
			<a href="javascript:;" onclick = "top.$.dialog.confirm('确定删除信息吗?',function(){location.href='shop/KS.ShopUserOrder.asp?Action=Del&id=<%=rs("id")%>&page=<%=CurrentPage%>';},function(){})" class="setA">删除</a>
					 <%
					 .Write "  </td> "
					  .Write "</tr>"
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RS.MoveNext
					Loop
					  RS.Close
				End With
			End Sub
			
			Sub S_AddEdit()
			dim RS,ContactMan,Address,ZipCode,Mobile,Phone,Email,QQ
			if KS.G("action")="Edit" then
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select top 1 * From KS_ShopUserOrder where  id="& KS.ChkClng(KS.G("id")) ,conn,1,1
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
				   top.$.dialog.alert('请输入收货人姓名!',function(){
				   $("#ContactMan").focus();});
				   return false;}
				  if ($("#Address").val()==""){
				   top.$.dialog.alert('请输入收货人地址!',function(){
				   $("#Address").focus();});
				   return false; }
				  if ($("#ZipCode").val()==""){
				   top.$.dialog.alert('请输入收货人邮编!',function(){
				   $("#ZipCode").focus();});
				   return false; }
				  if ($("#Phone").val()==""&&$("#Mobile").val()==""){
				   top.$.dialog.alert('收货人手机或电话至少要填一个!',function(){
				   $("#Mobile").focus();});
				   return false;}
				  if ($("#Email").val()==""){
				   top.$.dialog.alert('请输入收货人邮箱!',function(){
				   $("#Email").focus();});
				   return false;}
				  if ($("#mustyf").val()==1){
					if ($("#tocity").val()==''){
				   top.$.dialog.alert('请选择送货地区!');
				   return false;}
				  }
				  $("#myform").submit();
				}
			</script>
			  <div class="topdashed sort">请填写收货信息</div>
              <div class="pageCont2">
			  <FORM name="myform" id="myform"  action="KS.ShopUserOrder.asp?action=Save" method="post">
			  <input type="hidden" name="AddEdit" value="<%=KS.G("action")%>" />
			  <input type="hidden" name="id" value="<%=KS.G("id")%>" />
			  <input type="hidden" name="page" value="<%=KS.S("page")%>" />
			  <dl class="dtable">
                          <dd><div>收货人姓名：</div>
                           <INPUT class="textbox"  maxLength=50 value="<%=ContactMan%>" name="ContactMan" id="ContactMan">* 
                          </dd>
                          <dd><div>收货人地址：</div>
                          <INPUT class="textbox" maxLength=255 size=60 value="<%=Address%>" name="Address" id="Address">*
                           </dd>
                           <dd><div>收货人邮编：</div>
                            <INPUT class="textbox" maxLength=6 value="<%=ZipCode%>" name="ZipCode" id="ZipCode">* 
                           </dd>
                           <dd><div>手机号码：</div>
                           <INPUT class="textbox" maxLength=50 size=20 value="<%=Mobile%>" name="Mobile" id="Mobile">
                                                或固定电话<INPUT maxLength=50 class="textbox" size=20 value="<%=Phone%>" name="Phone" id="Phone">*<span class='tips'>两个至少填一个</span>
                            </dd>
                            <dd><div>收货人邮箱：</div>
                              <INPUT class="textbox" maxLength=100 size=30 value="<%=Email%>" name=Email id="Email">*
                            </dd>
                            <dd><div>收货人QQ：</div>
                              <INPUT class="textbox" maxLength=50 size=30 value="<%=QQ%>" name="QQ" id="QQ">
                            </dd>
				 </dl>
				 <div align="center"><button class="button" type="button" onClick="return CheckForm();" name="Submit"><strong> OK,保存信息 </strong></button></div>
				 </FORM>
				 </div>
			<%
		 End Sub
		 Sub S_save ()
			dim RS
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_ShopUserOrder where id="& KS.ChkClng(KS.G("id")) ,conn,1,3
			rs("ContactMan")=KS.G("ContactMan")
			rs("Address")=KS.G("Address")
			rs("ZipCode")=KS.G("ZipCode")
			rs("Mobile")=KS.G("Mobile")
			rs("Email")=KS.G("Email")
			rs("QQ")=KS.G("QQ")
			rs("Phone")=KS.G("Phone")
			Rs.update
			Rs.close: Set RS = Nothing
			Response.Write "<script>top.$.dialog.alert('恭喜，信息修改成功！',function(){location.href='shop/KS.ShopUserOrder.asp?page="& KS.S("page") &"';});</script>"
		 End Sub
		 Sub S_Del()
		 	if KS.ChkClng(KS.G("id"))<>0 then
				Conn.Execute("Delete  from KS_ShopUserOrder where id ="&KS.ChkClng(KS.G("id")) )
				Response.Write "<script>location.href='KS.ShopUserOrder.asp?page="& KS.S("page") &"'</script>"
			end if
		 End sub
			
			
			
			

End Class
%>
 

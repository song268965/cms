<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_LogDeliver
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_LogDeliver
        Private KS,KSCls
		Private totalPut,rs, MaxPerPage,DomainStr,SearchType,SQLParam
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
		If Not KS.ReturnPowerResult(5, "M510015") Then  Call KS.ReturnErr(1, ""):Exit Sub   
		SearchType=KS.ChkClng(KS.G("SearchType"))
		%>
<!DOCTYPE html><html>
<head><title>发退货查询</title>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<link href="../include/Admin_Style.css" type=text/css rel=stylesheet>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">
  <div class="tableTop mt20">
<FORM name=form1 action=KS.LogDeliver.asp method=get>
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td><strong>发退货查询：</strong></td>
      <td>快速查询： 
<Select onchange=javascript:submit() size=1 name=SearchType> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>所有发退货记录</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>最近10天内的新记录</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>最近一月内的新记录</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>所有发货记录</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>所有退货记录</Option>
  <Option value=6<%If SearchType="6" Then Response.write " selected"%>>所有申请退货记录</Option>
      </Select>&nbsp;&nbsp;&nbsp;&nbsp;<a href="KS.LogDeliver.asp">发退货记录首页</a></td></FORM>
<FORM name=form2 action=KS.LogDeliver.asp method=post>
      <td>高级查询： 
<Select id=Field name=Field > 
  <Option value=1 selected>客户姓名</Option> 
  <Option value=2>用户名</Option> 
  <Option value=3>发退货日期</Option> 
  <Option value=4>经手人</Option> 
  <Option value=5>快递公司</Option> 
  <Option value=6>快递单号</Option> 
  <Option value=7>订单号</Option>
</Select> 
  <Input id=Keyword class='textbox' maxLength=30 name=Keyword> 
  <Input class='button' type=submit value=" 查 询 " name=Submit2> 
        <Input id=SearchType type=hidden value=5 name=SearchType> </td>
    </tr> 
	 </table></FORM>
  </div>
  <div class="pageCont2 mt20">

 <table width="100%" cellSpacing=0 cellPadding=0  border="0" >
    <tr>
      <td align=left height="28"><i class='icon mainer'></i>您现在的位置：<a href="KS.LogDeliver.asp">发退货记录管理</a>&nbsp;&gt;&gt;&nbsp;
	  <%Dim SearchTypeStr
	    Dim KeyWord:KeyWord=KS.G("KeyWord")
	  Select Case SearchType
	     Case 0 :SearchTypeStr="所有记录"
		 Case 1 :SearchTypeStr="最近10天内的新记录"
		 Case 2 :SearchTypeStr="最近一月内的新记录"
		 Case 3 :SearchTypeStr="所有发货记录"
		 Case 4 :SearchTypeStr="所有退货记录"
		 Case 6 :SearchTypeStr="所有申请退货记录"
		 Case 5 
		    Select Case KS.ChkClng(KS.G("Field"))
			  Case 1:SearchTypeStr="客户姓名含有<font color=red>""" & KeyWord & """</font>"
			  Case 2:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
			  Case 3:SearchTypeStr="发退货日期含有<font color=red>""" & KeyWord & """</font>"
			  Case 4:SearchTypeStr="经手人含有<font color=red>""" & KeyWord & """</font>"
			  Case 5:SearchTypeStr="快递公司含有<font color=red>""" & KeyWord & """</font>"
			  Case 6:SearchTypeStr="快递单号含有<font color=red>""" & KeyWord & """</font>"
			  Case 7:SearchTypeStr="订单号含有<font color=red>""" & KeyWord & """</font>"
			End Select
	  End Select
	  Response.Write SearchTypeStr%></td>
    </tr>
  </table>
  <table cellSpacing="0" cellPadding="0" width="100%" border="0">
    <tr class=sort align=middle>
      <td width=70>日期</td>
	  <td>订单编号</td>
	  <td>用户名</td>
      <td>方向</td>
      <td>客户姓名</td>
      <td>快递公司</td>
      <td>快递单号</td>
      <td>经手人</td>
      <td>状态</td>
      <td>处理</td>
    </tr>
	<%
			MaxPerPage=20
			SqlParam="1=1"
            If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1
					SqlParam=SqlParam &" And datediff(" & DataPart_D & ",DeliverDate," & SqlNowString & ")<=10"
			   Case 2
					SqlParam=SqlParam &" And datediff(" & DataPart_D & ",DeliverDate," & SqlNowString & ")<=30"
			  Case 3 : SqlParam = SqlParam & "And status=1"
			  Case 4 : SqlParam = SqlParam & "And DeliverType=2"
			  Case 6 : SqlParam = SqlParam & "And DeliverType=3"
			  Case 5
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1:SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2:SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 3:SqlParam=SqlParam &" And DeliverDate Like '%" & Keyword & "%'"
				   Case 4:SqlParam=SqlParam &" And HandlerName Like '%" & Keyword & "%'"
				   Case 5:SqlParam=SqlParam &" And ExpressCompany Like '%" & Keyword & "%'"
				   Case 6:SqlParam=SqlParam &" And ExpressNumber Like '%" & Keyword & "%'"
				   Case 7:SqlParam=SqlParam &" And OrderID Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_LogDeliver Where " & SqlParam & " Order By ID Desc",Conn,1,1
	If RS.Eof AND RS.Bof Then
	 Response.WRITE "<tr class=list onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'""><td colspan=12 style='text-align:center' height='25'>找不到" & SearchTypeStr & "!</td></tr>"
   Else
              totalPut = RS.RecordCount
			  If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
				RS.Move (CurrentPage - 1) * MaxPerPage
			  End If
			  Call showContent()
   End If
   RS.Close:Set RS=Nothing
   %>
     <tr>
	   <td colspan="10" style="text-align:right">
         <%
		   	  '显示分页信息
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		   %>
    </td></tr>
	  </table>

</div>
</body>
</html>
   <%
   End Sub
  
  Sub ShowContent()
     Dim I,intotalDeliver,outtotalDeliver
     Do While Not rs.eof 
	%>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td height="30" align=middle><%=formatdatetime(rs("DeliverDate"),2)%></td>
      <td align=middle><%=rs("orderid")%></td>
	  <td align=center><%=rs("username")%></td>
	  <td align=middle>
	  <%
	  If rs("DeliverType")=1 Then
	   response.write "发货"
	  ElseIf rs("DeliverType")=3 Then
	   Response.Write "<font color=red>申请退货</font>"
	  ElseIf rs("DeliverType")=4 Then
	   Response.Write "<font color=blue>妥协交易</font>"
	  Else
	   Response.Write "<font color=green>成功退货</font>"
	  End If
	  %></td>
      <td align=middle><%=rs("ClientName")%></td>
      <td align=middle><%=rs("ExpressCompany")%></td>
      <td align=center><%=rs("ExpressNumber")%></td>
      <td align=center><%=rs("HandlerName")%></td>
      <td align=center>
	  <% If rs("status")=1 and rs("DeliverType")=1 Then
	      Response.Write "<font color=green>已签收</font>"
		 ElseIf rs("DeliverType")=3 and rs("status")=0 Then
	      Response.Write "<font color=red>待处理</font>"
		 ElseIf rs("status")=1 Then
	      Response.Write "<font color=green>已完结</font>"
		 Else
		  response.write "---"
		 End If
		 %></td>
      <td align="center" nowrap="nowrap">
	  <%If rs("DeliverType")=3 Then%>
	   <input type='button' value='同意退货' class="button" onClick="location.href='KS.ShopOrder.asp?Action=BankRefund&ID=<%=rs("orderid")%>';"/>
	   <input type='button' value='已妥协' class="button" onClick="if (confirm('确定将订单状态恢复正常吗？')){location.href='KS.ShopOrder.asp?Action=BankRefundOK&ID=<%=rs("orderid")%>';}"/>
	  <%else%>
	  ---
	  <%end if%>
	  </td>
    </tr>	
	<tr>
	 <td class="splittd" style="border-top:#f1f1f1 1px solid;height:50px" align="right">
	 <%If rs("DeliverType")=3 Then%>
	  <strong>退货原因：</strong>
	 <%else%>
	  备注说明：
	 <%end if%>
	 </td><td class="splittd" style="border-top:#f1f1f1 1px solid;color:#888;" colspan="10"><%=replace(rs("Remark")&"",chr(13),"<br/>")%>
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
End Class
%> 

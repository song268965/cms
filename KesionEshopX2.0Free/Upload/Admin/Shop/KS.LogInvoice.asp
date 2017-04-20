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
Set KSCls = New Admin_LogInvoice
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_LogInvoice
        Private KS,KSCls
		Private totalPut,rs, CurrentPage, MaxPerPage,DomainStr,SearchType,SQLParam
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
		If Not KS.ReturnPowerResult(5, "M510016") Then  Call KS.ReturnErr(1, ""):Exit Sub   
		%>
		<!DOCTYPE html><html>
<head><title>开发票查询</title>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<link href="../include/Admin_Style.css" type=text/css rel=stylesheet>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">

  <div class="tableTop mt20">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td><strong class="mr0">发票记录查询：</strong></td>
<FORM name=form1 action=KS.LogInvoice.asp method=get>
      <td valign="top"><span>快速查询：</span> 
<Select onchange=javascript:submit() size=1 name=SearchType> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>所有发票记录</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>最近10天内的新记录</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>最近一月内的新记录</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>所有地税普通发票</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>所有国税普通发票</Option>
  <Option value=4<%If SearchType="5" Then Response.write " selected"%>>所有增值税发票</Option>
      </Select></td></FORM>
<FORM name=form2 action=KS.LogInvoice.asp method=post>
      <td><span>高级查询：</span> 
<Select id=Field name=Field> 
  <Option value=1 selected>客户姓名</Option> 
  <Option value=2>用户名</Option> 
  <Option value=3>开发票日期</Option> 
  <Option value=4>经手人</Option> 
  <Option value=5>发票号码</Option> 
  <Option value=6>发票抬头</Option> 
  <Option value=7>订单号</Option>
</Select> 
  <Input id=Keyword class='textbox' maxLength=30 name=Keyword> 
  <Input class='button' type=submit value=" 查 询 " name=Submit2> 
        <Input id=SearchType type=hidden value=6 name=SearchType> </td></FORM>
    </tr>  
 </table>
  </div>


		<%
		 If KS.G("Action")="ShowInvoice" Then
		   Call ShowInvoice()
		 Else
		   Call LogList()
		 End If
		End Sub
		Sub LogList
		SearchType=KS.ChkClng(KS.G("SearchType"))
		%>
 <table width="100%" cellSpacing=0 cellPadding=0  border="0" class="hide">
    <tr>
      <td align=left height="28"><i class="icon mainer"></i> 您现在的位置：<a href="KS.LogInvoice.asp">开发票记录管理</a>&nbsp;&gt;&gt;&nbsp;
	  <%Dim SearchTypeStr
	    Dim KeyWord:KeyWord=KS.G("KeyWord")
	  Select Case SearchType
	     Case 0 :SearchTypeStr="所有发票记录"
		 Case 1 :SearchTypeStr="最近10天内的新记录"
		 Case 2 :SearchTypeStr="最近一月内的新记录"
		 Case 3 :SearchTypeStr="所有地税普通发票"
		 Case 4 :SearchTypeStr="所有国税普通发票"
		 Case 5 :SearchTypeStr="所有增值税发票"
		 Case 6 
		    Select Case KS.ChkClng(KS.G("Field"))
			  Case 1:SearchTypeStr="客户姓名含有<font color=red>""" & KeyWord & """</font>"
			  Case 2:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
			  Case 3:SearchTypeStr="开发票日期含有<font color=red>""" & KeyWord & """</font>"
			  Case 4:SearchTypeStr="经手人含有<font color=red>""" & KeyWord & """</font>"
			  Case 5:SearchTypeStr="发票号码含有<font color=red>""" & KeyWord & """</font>"
			  Case 6:SearchTypeStr="发票抬头含有<font color=red>""" & KeyWord & """</font>"
			  Case 7:SearchTypeStr="订单号含有<font color=red>""" & KeyWord & """</font>"
			End Select
	  End Select
	  Response.Write SearchTypeStr%></td>
    </tr>
  </table>
  <div class="pageCont2 mt20">
  <table  cellSpacing=0 cellPadding=0 width="100%" border=0>
    <tr class=sort align=middle>
        <td>发票抬头</td>
      <td width=80>日期</td>
      <td width=100>客户名称</td>
      <td width=60>用户名</td>
      <td width=100>订单编号</td>
      <td width=80>发票类型</td>
      <td width=80>发票号码</td>
      <td width=100>发票金额</td>
      <td width=60>开票人</td>
    </tr>
	<%
			MaxPerPage=20
			If KS.G("page") <> "" Then
				  CurrentPage = KS.ChkClng(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			SqlParam="1=1"
            If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1: SqlParam=SqlParam &" And datediff("& DataPart_D & ",InvoiceDate," & SqlNowString & ")<=10"
			   Case 2: SqlParam=SqlParam &" And datediff("& DataPart_D & ",InvoiceDate," & SqlNowString & ")<=30"
			  Case 3 : SqlParam = SqlParam & "And InvoiceType='地税普通发票'"
			  Case 4 : SqlParam = SqlParam & "And InvoiceType='国税普通发票'"
			  Case 5 : SqlParam = SqlParam & "And InvoiceType='增值税发票'"
			  Case 6
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1:SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2:SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 3:SqlParam=SqlParam &" And InvoiceDate Like '%" & Keyword & "%'"
				   Case 4:SqlParam=SqlParam &" And HandlerName Like '%" & Keyword & "%'"
				   Case 5:SqlParam=SqlParam &" And InvoiceNum Like '%" & Keyword & "%'"
				   Case 6:SqlParam=SqlParam &" And InvoiceTitle Like '%" & Keyword & "%'"
				   Case 7:SqlParam=SqlParam &" And OrderID Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_LogInvoice Where " & SqlParam & " Order By ID Desc",Conn,1,1
	If RS.Eof AND RS.Bof Then
	 Response.WRITE "<tr class=list onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'""><td class='splittd' style='text-align:center' colspan=9 align=center height='25'>找不到" & SearchTypeStr & "!</td></tr>"
   Else
                          totalPut = RS.RecordCount
							If CurrentPage < 1 Then	CurrentPage = 1

							If CurrentPage > 1  And (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
									
							End If
									Call showContent()
   End If
   RS.Close:Set RS=Nothing
   %>  </table>

     <div align=right>
         <%
		   	  '显示分页信息
			  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		   %>
    </div>
	</div>
</body>
</html>
   <%
   End Sub
  
  Sub ShowContent()
     Dim I,intotalInvoice,outtotalInvoice
     Do While Not rs.eof 
	%>
	<tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td><a href="KS.LogInvoice.asp?Action=ShowInvoice&ID=<%=rs("id")%>"><%=rs("InvoiceTitle")%></a></td>
      <td align=middle height="45" width=80><%=rs("InvoiceDate")%></td>
      <td align=middle width=60><%=rs("ClientName")%></td>
      <td align=middle width=60><%=rs("username")%></td>
      <td align=middle width=100><%=rs("orderid")%></td>
      <td align=middle width=80><%=rs("InvoiceType")%></td>
      <td align=middle width=80><%=rs("Invoicenum")%></td>
      <td align=middle width=60><%=rs("MoneyTotal")%></td>
      <td align=middle width=60><%=rs("HandlerName")%></td>
    </tr>
	<tr><td height="1" colspan="10" style="border-bottom:#f1f1f1 1px solid"></td></tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
		End Sub
		
		'查看发票信息
		Sub ShowInvoice()
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "SELECT * FROM KS_LogInvoice Where ID=" &ID,conn,1,1
		 IF RS.Eof Then
		    rs.close:set rs=nothing
		   Response.Write "<script>alert('参数出错!');history.back();</script>"
		   exit sub
		 End IF
		 %><div class="pageCont2 mt20">
  <table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=22><B>发 票 信 息</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">客户名称：</td>
      <td><%=rs("ClientName")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">订单编号：</td>
      <td><%=rs("orderid")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">订单金额：</td>
      <td><%=rs("MoneyTotal")%>元</td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">开票日期：</td>
      <td><%=rs("invoicedate")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票类型：</td>
      <td><%=rs("invoicetype")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票号码：</td>
      <td><%=rs("invoicenum")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票抬头：</td>
      <td><%=rs("InvoiceTitle")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票内容：</td>
      <td><%=rs("invoicecontent")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票金额：</td>
      <td><%=rs("MoneyTotal")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">开 票 人：</td>
      <td><%=rs("HandlerName")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">录 入 员：</td>
      <td><%=rs("inputer")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">录入时间：</td>
      <td><%=rs("inputtime")%></td>
    </tr>
  </table>
  <br>
  <div align=center><input type='button' value=' 返 回 ' onclick='javascript:history.back();' class='button'></div></div>
		 <%
		 rs.close
		 set rs=nothing
		 
		End Sub
End Class
%> 

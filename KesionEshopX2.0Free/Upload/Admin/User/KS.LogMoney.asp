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
Set KSCls = New Admin_LogMoney
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_LogMoney
        Private KS,KSCls,SQL,i
		Private totalPut,rs, Page, MaxPerPage,DomainStr,SearchType,SQLParam
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
	    If Not (KS.ReturnPowerResult(0, "KMUA10007") or KS.ReturnPowerResult(5,"M510014"))Then
			  Call KS.ReturnErr(1, "")
			End If
		SearchType=KS.ChkClng(KS.G("SearchType"))
		 dim begindate:begindate=request("begindate")
		 dim enddate:enddate=request("enddate")

		%>
<!DOCTYPE html><html><head><title>资金明细查询</title>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<link href="../include/Admin_Style.css" type=text/css rel=stylesheet>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">
<div class="quickLink">
    <table width="100%">
        <tr>
          <td align=left>您现在的位置：<a href="KS.LogMoney.asp">资金明细记录管理</a>&nbsp;&gt;&gt;&nbsp;
          <%Dim SearchTypeStr,SQLStr,TotalPages
            Dim KeyWord:KeyWord=KS.G("KeyWord")
          Select Case SearchType
             Case 0 :SearchTypeStr="所有资金明细记录"
             Case 1 :SearchTypeStr="最近10天内的新资金明细记录"
             Case 2 :SearchTypeStr="最近一月内的新资金明细记录"
             Case 3 :SearchTypeStr="所有收入记录"
             Case 4 :SearchTypeStr="所有支出记录"
             Case 5 
                Select Case KS.ChkClng(KS.G("Field"))
                  Case 1:SearchTypeStr="客户姓名含有<font color=red>""" & KeyWord & """</font>"
                  Case 2:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
                  Case 3:SearchTypeStr="交易时间含有<font color=red>""" & KeyWord & """</font>"
                End Select
            Case 100:SearchTypeStr="时间段查询结果"
          End Select
          Response.Write SearchTypeStr%></td>
        </tr>
      </table>
</div>  
  <div class="tableTop">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>资金明细查询：</td>

<FORM name=form2 action=KS.LogMoney.asp method=post>
      <td >高级查询： 
<Select id=Field name=Field> 
  <Option value=1 selected>客户姓名</Option> 
  <Option value=2>用户名</Option> 
  <Option value=3>交易时间</Option> 
</Select> 
  <Input id=Keyword class='textbox' maxLength=30 name=Keyword> 
  <Input class='button' type=submit value=" 查 询 " name=Submit2> 
        <Input id=SearchType type=hidden value=5 name=SearchType> </td></FORM>
    </tr>
  </table>
  </div>

<div class="tableTop mt20">  
   <table width="100%" border="0">
<form action="?action=search&SearchType=100" method=post name="myform">
   
   <tr>
     <td width="5%"><strong style="margin-right:0;">按时间段查询</strong></td>
     <td width="48%">
       <table width="100%"  align="center" border=0 cellPadding=0 cellSpacing=0>
         <tr>
           <td nowrap="nowrap" class=form-left>起始日期：
             <%if isdate(begindate) then%>
             <input type="text" name="begindate" value="<%=begindate%>" size="12" class="textbox">
             <%else%>
             <input type="text" name="begindate" value="<%=year(now)&"-"&month(now)&"-1"%>" size="12" class="textbox">
             <%end if%>
             <br>
             <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;如：2008-1-1</font>
			</td>    
		    <td  class=form-left nowrap="nowrap">终止日期：  
		      <%if isdate(enddate) then%>            
		      <input type="text" name="enddate" value="<%=enddate%>" size="12" class="textbox">
		      <%else%>
		      <input type="text" name="enddate" value="<%=formatdatetime(now,2)%>" size="12" class="textbox">
		      <%end if%>
		      <br>
		      <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;如：2008-1-31</font>
			  <br/><br/>
			</td>    
		  </tr>
        </table>
	 </td>
	</tr>
	<tr>
     <td>标志：</td><td width="43%"><input type="radio" name="direction" value="0"<%if request("direction")="0" or request("direction")="" then response.write " checked"%>>不限&nbsp;&nbsp;&nbsp;<input type="radio" name="direction" value="1"<%if request("direction")="1" then response.write " checked"%>>收入&nbsp;&nbsp;&nbsp;<input type="radio" name="direction" value="2"<%if request("direction")="2" then response.write " checked"%>>支出	 
	    <%
		 Call PaymentPlat(0)
		 If IsArray(SQL) Then
		   response.write "<br/><br/><strong class='mr0'>交易平台：</strong><select style='width:80px' name='PaymentID'>"
		   response.write "<option value='0'>-不限-</option>"
		   for i=0 to ubound(SQL,2)
		     If SQL(2,I)=1 Then
			  if trim(KS.S("PaymentID"))=trim(sql(0,i)) Then
		     response.write "<option value='" & SQL(0,I) & "' selected>" & Sql(1,i) &"</option>"
			  else
		     response.write "<option value='" & SQL(0,I) & "'>" & Sql(1,i) &"</option>"
			  end if
			 End If
		   next
		   response.write "</select>"
		 End If
		%>
	  关键字:
      <input type="text" name="keyword" class="textbox" size="10" value="<%=request("keyword")%>"/>
      <input name="submit2" type="submit" class="button" value="快速查找" />
      </td>
    </tr>
</form>
 </table>
</div> 
<table width="100%">
    <tr>
      <td align=left>
	  <%
	  if begindate<>"" and enddate<>"" then
	   response.write "<br><div align=center style='font-size:14px'>"
	 response.write "查询时间段 <font color=red>" & begindate & "</font> 至 <font color=red>" & enddate & "</font><br></div>"
	  end if
	  %></td>
    </tr>
  </table>

<div class="tabTitle mt20">会员资金明细</div>
<div class="tabs_header mt0">
    <ul class="tabs">
    <li<%if SearchType=0 then response.write " class='active'"%>><a href="KS.LogMoney.asp?<%=KS.QueryParam("SearchType")%>"><span>所有资金明细记录</span></a></li>
    <li<%if SearchType=1 then response.write " class='active'"%>><a href="KS.LogMoney.asp?SearchType=1&<%=KS.QueryParam("SearchType")%>"><span style='color:red'>最近10天内记录</span></a></li>
    <li<%if SearchType=2 then response.write " class='active'"%>><a href="KS.LogMoney.asp?SearchType=2&<%=KS.QueryParam("SearchType")%>"><span>最近一月内记录</span></a></li>
    <li<%if SearchType=3 then response.write " class='active'"%>><a href="KS.LogMoney.asp?SearchType=3&<%=KS.QueryParam("SearchType")%>"><span>所有收入记录</span></a></li>
    <li<%if SearchType=4 then response.write " class='active'"%>><a href="KS.LogMoney.asp?SearchType=4&<%=KS.QueryParam("SearchType")%>"><span>所有支出记录</span></a></li>

    </ul>
</div>
<div class="pageCont">
  <table cellSpacing="0" cellPadding="0" width="100%" border=0>
    <tr class=sort align=middle>
      <td width=120>交易时间</td>
      <td width=80>用户名</td>
      <td width=80>客户姓名</td>
      <td width=60>交易方式</td>
      <td width=50>币种</td>
      <td width=80>收入金额</td>
      <td width=80>支出金额</td>
      <td width=40>摘要</td>
      <td width=40>余额</td>
      <td>备注/说明</td>
    </tr>
	<%
			MaxPerPage=20
			Page = KS.ChkClng(KS.G("page"))
			If Page<=0 Then  Page = 1
			SqlParam="1=1"
            If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1:SqlParam=SqlParam &" And datediff("& DataPart_D & ",Logtime," & SqlNowString & ")<=10"
			   Case 2:SqlParam=SqlParam &" And datediff("& DataPart_D & ",Logtime," & SqlNowString & ")<=30"
			  Case 3 : SqlParam = SqlParam & "And IncomeOrPayOut=1"
			  Case 4 : SqlParam = SqlParam & "And IncomeOrPayOut=2"
			  Case 5
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1
				     SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2
				     SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 3
				     SqlParam=SqlParam &" And logtime Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
			If CInt(DataBaseType) = 1 Then         'Sql
				if isdate(begindate) then SqlParam=SqlParam & " and logtime>='" & begindate & "'"
				if isdate(enddate) then enddate=DateAdd("d", 1,EndDate):SqlParam=SqlParam & " and logtime<='" & enddate & "'"
			else
				if isdate(begindate) then SqlParam=SqlParam & " and logtime>=#" & begindate & "#"
				if isdate(enddate) then enddate=DateAdd("d", 1,EndDate):SqlParam=SqlParam & " and logtime<=#" & enddate & "#"
			end if
			if KS.ChkClng(KS.G("direction"))<>0 Then SqlParam=SqlParam & " and IncomeOrPayOut=" & KS.ChkClng(KS.G("Direction"))
			If KS.ChkClng(KS.G("PaymentID"))<>0 Then SqlParam=SqlParam & " and PaymentID=" & KS.ChkClng(KS.G("PaymentID"))


    If DataBaseType=1 Then
					Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
					Set Cmd.ActiveConnection=conn
					Cmd.CommandText="KS_GetPageRecords"
					Cmd.CommandType=4	
					CMD.Prepared = true 
					Cmd.Parameters.Append cmd.CreateParameter("@tblName",202,1,200)
					Cmd.Parameters.Append cmd.CreateParameter("@fldName",202,1,200)
					Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
					Cmd.Parameters.Append cmd.CreateParameter("@pageindex",3)
					Cmd.Parameters.Append cmd.CreateParameter("@ordertype",3)
					Cmd.Parameters.Append cmd.CreateParameter("@strWhere",202,1,1000)
					Cmd.Parameters.Append cmd.CreateParameter("@fieldIds",202,1,1000)
					Cmd("@tblName")="KS_LogMoney"
					Cmd("@fldName")= "ID"
					Cmd("@pagesize")=MaxPerPage
					Cmd("@pageindex")=page
					Cmd("@ordertype")=1
					Cmd("@strWhere")=SqlParam
					Cmd("@fieldIds")="*"
					Set Rs=Cmd.Execute
					Set Cmd=Nothing
	 Else
			SQLStr=KS.GetPageSQL("KS_LogMoney","ID",MaxPerPage,Page,1,SqlParam,"*")
			Set RS = Server.CreateObject("AdoDb.RecordSet")
			RS.Open SQLStr, conn, 1, 1
	 End If
	If RS.Eof AND RS.Bof Then
	 Response.WRITE "<tr class=list onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'""><td colspan=10 style='text-align:center' height='25'>找不到" & SearchTypeStr & "!</td></tr>"
   Else
        totalPut = Conn.Execute("Select Count(1) From KS_LogMoney Where "& SQLParam)(0)
		Call showContent()
   End If
   RS.Close:Set RS=Nothing
   %>
     <div align=right class="footerTable pt10">
         <%
		   	  '显示分页信息
  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
		   %>
    </div>
	
	<div class="attention" style="clear:both">
     <font color=red>说明：为免引起不必要的纠纷，资金明细仅提供查询功能，不能删除操作！</font>
	</div>
</body>
</html>
   <%
   End Sub
   
   Function PaymentPlat(PayMentID)
     Dim RS,I,str
     If Not IsArray(SQL) Then
		 Set RS=Conn.Execute("Select ID,PlatName,IsDisabled From KS_PaymentPlat Order By Id")
		 If Not RS.Eof Then
		  SQL=RS.GetRows(-1)
		 End If
		 RS.Close:Set RS=Nothing
	 End If
	 If IsArray(SQL) Then
	   For I=0 To Ubound(SQL,2)
	     If PayMentID=SQL(0,I) Then str=SQL(1,I) : Exit For
	   Next
	 End If
	 If KS.IsNul(str) Then Str="---"
	 PaymentPlat=str
   End Function
  
  Sub ShowContent()
     on error resume next
     Dim I,intotalmoney,outtotalmoney
     Do While Not rs.eof 
	%>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td class="splittd" align=middle width=120><%=rs("LogTime")%></td>
      <td class="splittd" align=middle width=80><%=rs("username")%></td>
	  <td class="splittd" align=middle width=80><%=rs("clientname")%></td>
      <td class="splittd" align=middle width=60>
	  <% 
	    Response.Write PaymentPlat(rs("PaymentID"))
	 %>
	  </td>
      <td class="splittd" align=middle width=50>人民币</td>
      <td class="splittd" width=80 align=right> &nbsp;
	  <%If rs("IncomeOrPayOut")=1 Then
			if rs("money")<1 and rs("money")>0 then
			 Response.Write "0" &formatnumber(rs("money"),2)
			else
			 Response.Write formatnumber(rs("money"),2)
			end if
		 intotalmoney=intotalmoney+rs("money")
	    End If
		%></td>
      <td class="splittd" align=right width=80>&nbsp;
	  <%If rs("IncomeOrPayOut")=2 Then
			if rs("money")<1 and rs("money")>0 then
			 Response.Write "0" &formatnumber(rs("money"),2)
			else
			 Response.Write formatnumber(rs("money"),2)
			end if
		 outtotalmoney=outtotalmoney+rs("money")
	    End If
		%></td>
      <td class="splittd" align=center width=40>
	  <% If rs("IncomeOrPayOut")=1 Then
	      Response.Write "<font color=red>收入</font>"
		 Else
		  Response.Write "<font color=green>支出</font>"
		 End If
		 %></td>
      <td class="splittd" align=middle><%=formatnumber(rs("currmoney"),2)%></td>
      <td class="splittd" align=middle><%=rs("Remark")%></td>
    </tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td class="splittd" align=right colSpan=5>本页合计：</td>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2)%></td>
      <td class="splittd" colSpan=3>&nbsp;</td>
    </tr>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td class="splittd" align=right colSpan=5>总计金额：</td>
	  <%intotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where "& SqlParam & " And IncomeOrPayOut=1")(0)
	    outtotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where "& SqlParam & " And IncomeOrPayOut=2")(0)
	    if not isnumeric(intotalmoney) then intotalmoney=0
		if not isnumeric(outtotalmoney) then outtotalmoney=0
	  %>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2)%></td>
      <td class="splittd" align=middle colSpan=3>资金余额：<%=formatnumber(intotalmoney-outtotalmoney,2)%></td>
    </tr>
  </table>
</div>  
		<%
		End Sub
End Class
%> 

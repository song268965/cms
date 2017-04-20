<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New Admin_EnterPrise
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrise
        Private KS,Param,KSCls
		Private Action,i,strClass,RS,SQL
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		
		 If Request("Action")= "export" Then
		   DoExport
		   Exit Sub
		 End If
		
		 With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class='tabs_header'><ul class='tabs'><li"
			  If KS.G("Action")="Custom" Then .Write " class='active'"
			  .Write "><a href='?action=Custom'><span>客户统计</span></a></li><li"
			  If KS.G("Action")="Order" Then .Write " class='active'"
			  .Write "><a href='?action=Order'><span>订单统计</span></a></li><li"
			  If KS.G("Action")="Sale" Then .Write " class='active'"
			  .Write "><a href='?action=Sale'><span>销售概况</span></a></li><li"
			  If KS.G("Action")="SalePM" Then .Write " class='active''"
			  .Write "><a href='?action=SalePM'><span>销售排名</span></a></li><li"
			  If KS.G("Action")="ProSaleCount" Then .Write " class='active''"
			  .Write "><a href='?action=ProSaleCount'><span>各商品销售统计</span></a></li> "
			  .Write "</ul></div>"
		End With
	   	If Not KS.ReturnPowerResult(5, "M510017") Then  Call KS.ReturnErr(1, ""):Exit Sub   
		Select Case KS.G("action")
		 Case "SalePM"
		  SalePM
		 Case "Custom"
		  Custom
		 Case "Order" 
		  Order
		 Case "Sale" 
		  Sale
		 Case "ProSaleCount" ProSaleCount 
		 Case Else
		End Select
		
		
End Sub

Sub DoExport()
   Response.AddHeader "Content-Disposition", "attachment;filename=orderitem.xls" 
   Response.ContentType = "application/vnd.ms-excel" 
   Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
   Response.Write Request.Form("content")
End Sub

Sub ProSaleCount()
 Dim MaxPerPage,CurrentPage,TotalPut,flag
 MaxPerPage=50
 CurrentPage=KS.ChkClng(KS.G("Page"))
 If CurrentPage<=0 Then CurrentPage=1
 flag=ks.chkclng(request("flag"))
 if request("flag")="" then flag=1
%>
<script type="text/javascript">
function ShowSale(id,title)
 {  top.openWin("查看商品销售详情","shop/KS.ShopProSale.asp?proid="+id+"&title="+escape(title),false);}
</script>
<div class="pageCont">
<div class="allshop">
<a href="?action=ProSaleCount"<%if flag=0 then response.write " style='color:red'"%>>所有商品</a> | <a href="?action=ProSaleCount&flag=1"<%if flag=1 then response.write " style='color:red'"%>>仅显示有销售记录的商品</a> | <a href="?action=ProSaleCount&flag=2"<%if flag=2 then response.write " style='color:red'"%>>仅显示有有付款记录的商品</a>
</div>
<%
dim param:param=" where verific=1 and deltf=0"
if flag=1 then
  param=param & " and id in(select proid from ks_orderitem)"
elseif flag=2 then
  param=param & " and id in(select i.proid from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.MoneyReceipt>0)"
end if
if request("key")<>"" then
  param=param & " and title like '%" & KS.G("Key") & "%'"
end if

 Dim str,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "select * From KS_Product " & param & " Order By Id Desc",conn,1,1
  str= "<table width=""99%"" border=""{$border}"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""otable"">"
  str=str & "<tr class=""sort"">"
  str=str & "<td class='splittd'><strong>商品名称</strong> </td>"
  str=str & "<td class='splittd'><strong>销售量</strong> </td>"
  str=str & "<td class='splittd'><strong>已付款</strong> </td>"
  str=str & "<td class='splittd'><strong>未付款</strong> </td>"
  str=str & "<td class='splittd'><strong>销售情况</strong> </td>"
  str=str & "</tr>"
  If RS.Eof And RS.Bof Then
		str=str & "<tr><td colspan=5 class='splittd' style='text-align:center'>找不到任何销售的商品!</td></tr>"
  Else
		      totalPut = rs.recordcount
			  If CurrentPage < 1 Then	CurrentPage = 1
			  If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrentPage - 1) * MaxPerPage
			  Else
				  	CurrentPage = 1
			  End If
		      i=0
			  do while not rs.eof 
				 str=str & "<tr><td class='splittd'><img src='../Images/ico/doc0.gif' align='absmiddle'/><a href=""javascript:ShowSale(" & rs("id") & ",'" & rs("title")  & "');"">" & rs("title") & "</a></td>"
				 str=str & "<td class='splittd' style='text-align:center'>" & KS.ChkClng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where i.proid=" & rs("id"))(0)) & " " & rs("unit") &"</td>"
				 str=str & "<td class='splittd' style='text-align:center'><font color=blue>" & ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.MoneyReceipt>0 and i.proid=" & rs("id"))(0)) & "</font> " & rs("unit") &"</td>"
				 str=str & "<td class='splittd' style='text-align:center'><font color=red>" & ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.MoneyReceipt<=0 and i.proid=" & rs("id"))(0)) & "</font> " & rs("unit") &"</td>"
				 str=str & "<td class='splittd' style='text-align:center'><a href=""javascript:ShowSale(" & rs("id") & ",'" & rs("title")  & "');"">查看销售详情</a></td>"
				 str=str & "</tr>"
				 i=i+1
				 If I>=MaxPerPage Then Exit Do
			  rs.movenext
			  loop
		  End If
			  rs.close
			  set rs=nothing
		str=str & "</table>"
		response.write replace(str,"{$border}",0)
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
 %>
 <div style="clear:both"></div>
 <br/>
   <table width="99%" style="margin-top:10px" border="0" align="center" cellpadding="1" cellspacing="1">
					  <form name="myform" action="?action=export" method="post">
					    <textarea name="content" style="display:none"><%=replace(str,"{$border}",1)%></textarea>
					  <tr>
						<td height="30" style="text-align:left"> <input class="button" type="submit" value=" 将以上结果导出到excel ">
					   </td>
					  </tr>
					   </form>
   <table width="99%" style="margin-top:10px" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
					  <form name="myform" action="?action=ProSaleCount" method="post">
					  <tr class="tdbg" >
						<td height="30" class='clefttitle' style="text-align:left"><strong>商品名称</strong> <input name='key' type='text' id='key' size='30'  class='textbox'>       <input class="button" type="submit" value=" 开 始 搜 索 ">
					   </td>
					  </tr>
					   </form>

</table></div><div class="footerTable"></div>
 
 <%
End Sub

Sub SalePM()
%>
<script src="../../KS_Inc/DatePicker/WdatePicker.js"></script>
<div class="pageCont">
<div class="allshop"><a href='?action=SalePM'<%If request("d")="" then response.write " style='color:red'"%>>所有时间</a>|<a href='?action=SalePM&d=1'<%If request("d")="1" then response.write " style='color:red'"%>>今日排名</a>|<a href='?action=SalePM&d=2'<%If request("d")="2" then response.write " style='color:red'"%>>本周排名</a>|<a href='?action=SalePM&d=3'<%If request("d")="3" then response.write " style='color:red'"%>>本月排名</a>|<a href='?action=SalePM&d=6'<%If request("d")="6" then response.write " style='color:red'"%>>上个月排名</a>|<a href='?action=SalePM&d=4'<%If request("d")="4" then response.write " style='color:red'"%>>本年度排名</a></div>
<table width="99%" style="margin-top:10px" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left"><strong>按时间段统计</strong> </td>
					  </tr>
					  <form name="myform" action="?action=SalePM" method="post">
					  <tr class="tdbg">
						<td height="30">&nbsp;&nbsp;开始日期<input name='BeginDate' type='text' id='BeginDate' value='<%=Year(Now) & "-" & Month(NOW) & "-1"%>' size='20' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  class='textbox'>
						&nbsp;&nbsp;
						结束日期<input name='EndDate' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  type='text' id='EndDate' value='<%=formatdatetime(Now,2)%>' size='20'  class='textbox'>    </td>
						<td>
					    <input type="hidden" name="d" value="5"/>
					    <input class="button" type="submit" value=" 开 始 统 计 ">
					   </td>
					  </tr>
					   </form>

</table>
<br/>
<%
 dim rs,sql,MaxPerPage,CurrentPage,i,TotalPut,Param,str
 MaxPerPage=50
 CurrentPage=KS.ChkClng(KS.G("Page"))
 If CurrentPage<=0 Then CurrentPage=1

  'Param="where o.status<>0 and o.status<>3"
  Param="where o.MoneyReceipt>0"
		  
 str="<div style=""padding:5px;color:green;font-weight:bold"">"
If request("d")<>"" then
 
   select case KS.ChkClng(Request("d"))
     case 1 
	  str=str & "今日销售排行"
	  Param=Param & " and year(inputtime)=" & year(now) & " and month(inputtime)=" & month(now) & " and day(inputtime)=" & day(now)
	 case 2
	  str=str & "本周销售排行"
	  Param=Param & " and datediff(" & DataPart_W & ",[inputtime]," & SQLNowString & ")=0"
	 case 3
	  str=str & "本月销售排行"
	  Param=Param & " and datediff(" & DataPart_M & ",[inputtime]," & SQLNowString & ")=0"
	 case 6
	  str=str & "上个月销售排行"
	  Param=Param & " and datediff(" & DataPart_M & ",[inputtime]," & SQLNowString & ")=1"
	 case 4
	  str=str & "本年度销售排行"
	  Param=Param & " and datediff(" & DataPart_Y & ",[inputtime]," & SQLNowString & ")=0"
	 case 5
	  str=str &  "时间段 <font color=red>" & KS.S("BeginDate") & " 至 " & KS.S("EndDate") & "</font> 的销售排行"
			If KS.S("BeginDate")<>"" and IsDate(KS.S("BeginDate")) Then 
			  If DataBaseType=1 Then
			   Param=Param & " and inputtime>='" & KS.S("BeginDate") & "'"
			  Else
			   Param=Param & " and inputtime>=#" & KS.S("BeginDate") & "#"
			  End if
			End If
			If KS.S("EndDate")<>"" and IsDate(KS.S("EndDate")) Then 
			  Dim EndDate:EndDate = DateAdd("d", 1, Request("EndDate"))
			  If DataBaseType=1 Then
			   Param=Param & " and O.InputTime<='" & DateAdd("d", 1,EndDate) & "'"
			  Else
			   Param=Param & " and O.InputTime<=#" & DateAdd("d", 1,EndDate) & "#"
			  End if
			End If
   end select 
     
else
 str=str & "所有时间的商品销售排行"
end if
str=str & "</div>"

         str=str & "<table width=""99%"" border=""{$border}"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""otable"">"
		 str=str &"<tr class=""sort"">"
		 str=str & "<th class='splittd' height='26' style=""text-align:center;font-size:12px""><strong>商品名称</strong> </th>"
		 str=str &"	<th class='splittd' style=""text-align:center;font-size:12px""><strong>销售量</strong> </th>"
		 str=str &"	<th class='splittd' style=""text-align:center;font-size:12px""><strong>总销售额(元)</strong> </th>"
		 str=str &"</tr>"
		  sql="select p.title,p.unit,a.proid,sum(a.Amount) as SaleNum,sum(TotalPrice) as SaleMoney from (ks_orderitem a inner join ks_product p on a.proid=p.id) inner join ks_order o on a.orderid=o.orderid " & Param & " group by a.proid,p.title,p.unit order by sum(Amount) desc"
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open sql,conn,1,1
		  If RS.Eof And RS.Bof Then
		    str=str & "<tr><td colspan=5 class='splittd' style='text-align:center'>找不到任何销售的商品!</td></tr>"
		  Else
		      totalPut = conn.execute("select count(1) from (ks_orderitem a inner join ks_product p on a.proid=p.id) inner join ks_order o on a.orderid=o.orderid " & Param)(0)
			  If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrentPage - 1) * MaxPerPage
			  Else
				  	CurrentPage = 1
			  End If
		      i=0
			  do while not rs.eof 
				 str=str & "<tr><td class='splittd'><a href='../../item/show.asp?m=5&d=" & rs("proid") & "' target='_blank'>" & rs(0) & "</a></td>"
				 str=str & "<td class='splittd' style='text-align:center'>" & rs("salenum") & " " & rs(1) & "</td>"
				 str=str & "<td class='splittd' style='text-align:center'>￥" & formatnumber(rs("salemoney"),2,-1,-1) & "</td>"
				 str=str & "</tr>"
				 i=i+1
				 If I>=MaxPerPage Then Exit Do
			  rs.movenext
			  loop
		  End If
			  rs.close
			  set rs=nothing
		str=str &"</table>"
		response.write replace(str,"{$border}",0)
		
Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
%>
<br/><br/><br/>
<form name="form1" action="KS.ShopStats.asp" method="post">
 <textarea name="content" style="display:none"><%=replace(str,"{$border}",1)%></textarea>
 <input type="hidden" name="action" value="export" />
 <input type="submit" class="button" value="将以上数据导出到Excel" />
</form></div><div class="footerTable pt10">
<div class="attention">
<font color=red><strong>说明：</strong><br/>
1、这里统计到的数据不包括非确认及无效订单;<br/>
2、您可以选择按时间段统计，这里每页显示50条记录;
</font>
</div></div>
<%
End Sub

Sub Custom()
%>			  
				 <div class="pageCont pd10">
				 <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="otable md10 otableBor">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left;text-align: left;background:#f3f6f7;padding-left: 15px;font-size: 14px;"><strong>会员购买率</strong> （会员购买率 = 会员有效订单数 ÷ 会员订单总数） </td>
					  </tr>
					  <tr class="tdbg">
						<td height="30" style="text-align:center">会员总数</td>
						<td height="30" style="text-align:center">会员有效订单数</td>
						<td height="30" style="text-align:center">会员订单总数</td>
						<td height="30" style="text-align:center">会员购买率 </td>
					  </tr>
					  <tr>
						<td height="30" style="text-align:center"><%
						Dim TotalUser:TotalUser=Conn.Execute("Select count(*) From KS_User")(0)
						Response.Write TotalUser
						%></td>
						<td height="30" style="text-align:center"><%
						Dim HasOrderUserTotal:HasOrderUserTotal=Conn.Execute("Select count(*) From KS_User Where UserName in(select username from ks_order where status<>0 and Status<>3)")(0)
						Response.Write HasOrderUserTotal%></td>
						<td height="30" style="text-align:center"><%
						Dim UserTotalOrder:UserTotalOrder=Conn.Execute("Select count(*) From KS_order Where UserName<>'游客'")(0)
						response.write UserTotalOrder
						%></td>
						<td height="30" style="text-align:center">
						<%
						 if HasOrderUserTotal<>0 and UserTotalOrder<>0 Then
						  Response.Write formatpercent(HasOrderUserTotal/UserTotalOrder,2)
						 else
						  response.write "0%"
						 end if
						%>
						</td>
					  </tr>
					</table>
					
				 <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="otable md10 otableBor">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left;background:#f3f6f7;padding-left: 15px;font-size: 14px;"><strong>每会员平均订单数及购物额</strong> （每会员订单数 = 会员订单总数 ÷ 会员总数） （每会员购物额 = 会员购物总额 ÷ 会员总数） </td>
					  </tr>
					  <tr class="tdbg">
						<td height="30" style="text-align:center">会员购物总额</td>
						<td height="30" style="text-align:center">每会员订单数</td>
						<td height="30" style="text-align:center">每会员购物额</td>
					  </tr>
					  <tr>
						<td height="30" align="center">￥<%
						Dim AllTotalMoney:AllTotalMoney=Conn.Execute("Select Sum(NoUseCouponMoney) From KS_Order where username<>'游客' and status<>0 and Status<>3")(0)
						if AllTotalMoney="" Or IsNull(AllTotalMoney) Then
						 response.write "0"
						Else
						 Response.Write AllTotalMoney
						End If
						%> 元</td>
						<td height="30" align="center">
						<%
						Dim AllUserOrder:AllUserOrder=Conn.Execute("Select count(*) From KS_Order where username<>'游客' and status<>0 and Status<>3")(0)
						If TotalUser<>0 and AllUserOrder<>0 Then
						 Response.Write AllUserOrder/TotalUser
						else
						Response.Write "0"
						end if%></td>
						<td height="30" align="center">
						￥<%
						 if AllTotalMoney<>0 and not isnull(AllTotalMoney) and TotalUser<>0 and not isnull(TotalUser) Then
						  response.write AllTotalMoney/TotalUser
						 else
						  response.write "0"
						 end if
						%>元</td>
						
					  </tr>
					</table>
					
					 <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="otable md10 otableBor">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left;background:#f3f6f7;padding-left: 15px;font-size: 14px;"><strong>匿名会员平均订单额及购物总额</strong> （匿名会员平均订单额 = 匿名会员购物总额 ÷ 匿名会员订单总数）  </td>
					  </tr>
					  <tr class="tdbg">
						<td height="30" style="text-align:center">匿名会员购物总额</td>
						<td height="30" style="text-align:center">匿名会员订单总数</td>
						<td height="30" style="text-align:center">匿名会员平均订单额</td>
					  </tr>
					  <tr>
						<td height="30" align="center">￥<%
						Dim AllNMTotalMoney:AllNMTotalMoney=Conn.Execute("Select Sum(NoUseCouponMoney) From KS_Order where username='游客' and status<>0 and Status<>3")(0)
						if AllNMTotalMoney="" Or IsNull(AllNMTotalMoney) Then
						 response.write "0"
						Else
						 Response.Write AllNMTotalMoney
						End If
						%> 元</td>
						<td height="30" align="center">
						<%
						Dim AllNMOrder:AllNMOrder=Conn.Execute("Select count(*) From KS_Order where username='游客' and status<>0 and Status<>3")(0)
						If AllNMOrder<>0 Then
						 Response.Write AllNMOrder
						else
						Response.Write "0"
						end if%></td>
						<td height="30" align="center">
						￥<%
						 if AllNMTotalMoney<>0 and not isnull(AllNMTotalMoney) and AllNMOrder<>0 and not isnull(AllNMOrder) Then
						  response.write AllNMTotalMoney/AllNMOrder
						 else
						  response.write "0"
						 end if
						%>元</td>
						
					  </tr>
					</table>
					
					<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="otable md10 otableBor">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left;background: #f3f6f7;padding-left: 15px;font-size: 14px;"><strong>综合统计</strong>  </td>
					  </tr>
					  <tr class="tdbg">
						<td height="30" style="text-align:center">总订单数 ／ 总购买金额</td>
						<td height="30" style="text-align:center">会员订单% ／非会员订单%　</td>
						<td height="30" style="text-align:center">会员购买金额% ／非会员购买金额%</td>
					  </tr>
					  <tr>
						<td height="30" align="center">
						<%
						Dim TotalOrder:TotalOrder=Conn.Execute("Select count(*) From KS_Order Where Status<>0 and Status<>3")(0)
						Dim TotalMoney:TotalMoney=Conn.Execute("Select sum(NoUseCouponMoney) From KS_Order Where Status<>0 and Status<>3")(0)
						Response.Write TotalOrder
						%>／￥<%=TotalMoney%> 元</td>
						<td height="30" align="center">
						<%
						 if AllUserOrder<>0 and TotalOrder<>0 then
						 response.write formatpercent(AllUserOrder/TotalOrder,2)
						 else
						 response.write "0%"
						 end if
						%>／
						<%
						if TotalOrder<>0 and AllNMOrder<>0 then
						 response.write formatpercent(AllNMOrder/TotalOrder,2)
						else
						 response.write "0%"
						end if
						%>
						</td>
						<td height="30" align="center">
						<%
						 if AllTotalMoney<>0 and TotalMoney<>0 Then
						   response.write formatpercent(AllTotalMoney/TotalMoney,2)
						 Else
						   response.write "0%"
						 end if
						 %>
						／
						 <%
						 if AllNMTotalMoney<>0 and TotalMoney<>0 Then
						   response.write formatpercent(AllNMTotalMoney/TotalMoney,2)
						 Else
						   response.write "0%"
						 end if
						
						%></td>
						
					  </tr>
					</table></div><div class="footerTable"></div>
					
				<%
End Sub

Sub Order()
%><script src="../../KS_Inc/DatePicker/WdatePicker.js"></script>
<div class="pageCont">
<form name="myform" action="?action=Order" method="post">
         <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr class="tdbg">
			<td height='25' class="clefttitle" colspan='4' style="text-align:left;background:#f3f6f7;padding-left: 15px;font-size: 14px;"><strong>按时间段统计</strong> </td>
		  </tr>
		  <tr class="tdbg">
			<td height="50">&nbsp;&nbsp;开始日期<input name='BeginDate' type='text' id='BeginDate' value='<%=Year(Now) & "-" & Month(NOW) & "-1"%>' size='20' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  class='textbox'>   
			&nbsp;&nbsp;
			结束日期<input name='EndDate' type='text' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" id='EndDate' value='<%=formatdatetime(Now,2)%>' size='20'  class='textbox'>                
			<input class="button" type="submit" value=" 开 始 统 计 ">
		   </td>
		  </tr>
		</table>
</form>
<%
Dim Param:Param=" Where 1=1"
If KS.S("BeginDate")<>"" and IsDate(KS.S("BeginDate")) Then 
  If DataBaseType=1 Then
   Param=Param & " and InputTime>='" & KS.S("BeginDate") & "'"
  Else
   Param=Param & " and InputTime>=#" & KS.S("BeginDate") & "#"
  End if
End If
If KS.S("EndDate")<>"" and IsDate(KS.S("EndDate")) Then 
  If DataBaseType=1 Then
   Param=Param & " and InputTime<='" & DateAdd("d", 1,KS.S("EndDate")) & "'"
  Else
   Param=Param & " and InputTime<=#" & DateAdd("d", 1,KS.S("EndDate")) & "#"
  End if
End If

Dim WQR:WQR=KS.ChkCLng(Conn.Execute("Select count(*) From KS_Order " &Param & " and Status=0")(0))
Dim YQR:YQR=KS.ChkCLng(Conn.Execute("Select count(*) From KS_Order " &Param & " and Status=1")(0))
Dim YCJ:YCJ=KS.ChkCLng(Conn.Execute("Select count(*) From KS_Order " &Param & " and Status=2")(0))
Dim wx:wx=KS.ChkCLng(Conn.Execute("Select count(*) From KS_Order " &Param & " and Status=3")(0))
%>
<script src="../../ks_inc/flotr2.min.js" type="text/javascript"></script>

 <table width="100%" cellspacing="0" cellpadding="3" id="order_circs-table">
      <tr>
        <td align="center">
		
		<div id="container" style="margin:0 auto;width:980px;height:400px"></div>
	
    <script type="text/javascript" src="flotr2.min.js"></script>
    <script type="text/javascript">
   (function basic_pie(container) {
  var graph;
  graph = Flotr.draw(container, [
    { data : [[0, <%=YQR%>]], label : '已确认' },
    { data : [[0, <%=YCJ%>]], label : '已成交',
      pie : {
        explode :50
      } },
    { data : [[0, <%=WQR%>]], label : '未确认' },
    { data : [[0, <%=Wx%>]], label : '无效订单' },
  ],
   {
    HtmlText : true,
    grid : {
      verticalLines : false,
      horizontalLines : false
    },
    xaxis : { showLabels : false },
    yaxis : { showLabels : false },
    pie : {
      show : true, 
      explode : 6
    },
    mouse : { track : true },
    legend : {
      position : 'se',
      backgroundColor : '#D2E8FF'
    }
  });
})(document.getElementById("container"));
</script>
		
		
		
		
		
                </td>
      </tr>
    </table></div><div class="footerTable"></div>
<%
End Sub

Sub Sale
Dim K,Y,OrderTotal,OrderXmlStr,SalesTotal,SalesXmlStr,UnitStr
Dim StartYear:StartYear=KS.G("StartYear")
Dim EndYear:EndYear=KS.G("EndYear")
Dim StartMonth:StartMonth=KS.G("StartMonth")
Dim EndMonth:EndMonth=KS.G("EndMonth")
If Not IsNumeric(StartYear) Then StartYear=Year(Now)
If Not IsNumeric(EndYear) Then EndYear=Year(Now)
If Not Isnumeric(StartMonth) Then StartMonth=1
If Not IsNumeric(EndMonth) Then EndMonth=Month(Now)

If KS.G("saletype")="year" Then
	For Y=StartYear To EndYear
		 OrderTotal=Conn.Execute("Select Count(*) From KS_Order Where Year(inputtime)=" & Y)(0)
		' OrderXmlStr=OrderXmlStr & "<set label='" & Y & "年' value='" & OrderTotal & "' />"
		  OrderXmlStr=OrderXmlStr & "[[" & Y &", " & OrderTotal &"]],"
		 OrderTotal=Conn.Execute("Select Sum(NoUseCouponMoney) From KS_Order Where status<>0 and Status<>3 and Year(inputtime)=" & Y)(0)
		 SalesXmlStr=SalesXmlStr & "[[" & Y &", " & OrderTotal &"]],"
		 UnitStr=UnitStr &"[" & Y  &",""" & Y &"年""],"
		' SalesXmlStr=SalesXmlStr & "<set label='" & Y & "年' value='" & OrderTotal & "' />"
	Next
Else
	For Y=StartYear To EndYear
		For K=StartMonth To 12
		 OrderTotal=KS.ChkClng(Conn.Execute("Select Count(*) From KS_Order Where Year(inputtime)=" & Y & " and Month(inputtime)=" & K)(0))
		 OrderXmlStr=OrderXmlStr & "[[" & K &", " & OrderTotal &"]],"

		 OrderTotal=Conn.Execute("Select Sum(NoUseCouponMoney) From KS_Order Where status<>0 and Status<>3 and Year(inputtime)=" & Y & " and Month(inputtime)=" & K)(0)
		 SalesXmlStr=SalesXmlStr & "[[" & K &", " & OrderTotal &"]],"
		 UnitStr=UnitStr &"[" & K  &",""" & K &"月份""],"
		 If Cint(Y)=Cint(EndYear) And Cint(K)=Cint(EndMonth) Then Exit For
		Next
		If Y=EndYear And K=EndMonth Then Exit For
	Next
End If
%>
<div class="pageCont">
 <table width="99%" border="0" align="center" cellpadding="1" cellspacing="1" class="otable">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left"><strong>按月查看走势</strong> </td>
					  </tr>
					  <form name="myform" action="?action=Sale" method="post">
					  <tr class="tdbg">
						<td height="30">&nbsp;&nbsp;
						<select name="StartYear">
						 <%
						  for k=year(now)-2 to year(Now)
						   if k=cint(startyear) then
						   response.write "<option  value=" & k & " selected>" & k & "</option>"
						   else
						   response.write "<option  value=" & k & ">" & k & "</option>"
						   end if
						  next
						 %>
						</select>年
						<select name="StartMonth">
						 <%
						  for k=1 to 12
						   if k=cint(startmonth) then
						   response.write "<option  value=" & k & " selected>" & k & "</option>"
						   else
						   response.write "<option  value=" & k & ">" & k & "</option>"
						   end if
						  next
						 %>
						</select>月
						至
						<select name="EndYear">
						 <%
						  for k=year(now)-2 to year(Now)
						   if k=cint(endyear) then
						   response.write "<option  value=" & k & " selected>" & k & "</option>"
						   else
						   response.write "<option  value=" & k & ">" & k & "</option>"
						   end if
						  next
						 %>
						</select>年
						<select name="EndMonth">
						 <%
						  for k=1 to 12
						   if k=cint(endmonth) then
						   response.write "<option  value=" & k & " selected>" & k & "</option>"
						   else
						   response.write "<option  value=" & k & ">" & k & "</option>"
						   end if
						  next
						 %>
						</select>月
						
						<input class="button" type="submit" value=" 查看走势 ">
						</td>
					  </tr>
					   </form>
		</table>
 <table width="99%" style="margin-top:10px" border="0" align="center" cellpadding="1" cellspacing="1" class="otable">
					  <tr class="tdbg">
						<td height='25' class='clefttitle' colspan='4' style="text-align:left"><strong>按年查看走势</strong> </td>
					  </tr>
					  <form name="myform" action="?action=Sale" method="post">
					  <input type="hidden" name="saletype" value="year">
					  <tr class="tdbg">
						<td height="30">&nbsp;&nbsp;
						<select name="StartYear">
						 <%
						  for k=year(now)-5 to year(Now)
						   if k=cint(startyear) then
						   response.write "<option  value=" & k & " selected>" & k & "</option>"
						   else
						   response.write "<option  value=" & k & ">" & k & "</option>"
						   end if
						  next
						 %>
						</select>年
						至
						<select name="EndYear">
						 <%
						  for k=year(now)-5 to year(Now)
						   if k=cint(endyear) then
						   response.write "<option  value=" & k & " selected>" & k & "</option>"
						   else
						   response.write "<option  value=" & k & ">" & k & "</option>"
						   end if
						  next
						 %>
						</select>年
						
						
						<input class="button" type="submit" value=" 查看走势 ">
						</td>
					  </tr>
					   </form>
		</table>


   <div style="margin:10px;font-weight:Bold;color:#ff6600;font-size:14px">订单走势</div>
   <div id="tabbody-div">
      <!-- 订单数量 -->
      <table width="90%" id="order-table">
        <tr><td align="center">
		
		
		<div id="container1" style="margin:0 auto;width:800px;height:400px"></div>
    <script type="text/javascript" src="../../ks_inc/flotr2.min.js"></script>
    <script type="text/javascript">
   (function basic_bars(container, horizontal) {

  var
    horizontal = (horizontal ? true : false), // Show horizontal bars
    point,                                    // Data point variable declaration
    i;
  
  // Draw the graph
  Flotr.draw(
    container,
    [ <%=OrderXmlStr%> ],
    {
      bars : {
        show : true,
        horizontal : horizontal,
        shadowSize : 0,
        barWidth : 1
      },
	  xaxis: {
              ticks:[<%=UnitStr%>], // 自定义X轴
               minorTicks: null,
			  <%If KS.G("saletype")="year" then%>
			   title:'年份',   
			   <%else%>
			   title:'月份',   
			   <%end if%>
               showLabels:true                             // 是否显示X轴刻度
	  },
      mouse : {
        track : true,
        relative : true
      },
      yaxis : {
        min : 0,
		title:"订单数(单位:个)",
        autoscaleMargin : 1
      }
    }
  );
})(document.getElementById("container1"));

    </script>
		

        </td></tr>
      </table>

      <!-- 营业额 -->
	  <div style="margin:10px;font-weight:Bold;color:#ff6600;font-size:14px">销售额走势</div>
      <table width="90%" id="turnover-table">
        <tr><td align="center">
		
		<div id="container" style="margin:0 auto;width:800px;height:400px"></div>
    <script type="text/javascript" src="../../ks_inc/flotr2.min.js"></script>
    <script type="text/javascript">
   (function basic_bars(container, horizontal) {

  var
    horizontal = (horizontal ? true : false), // Show horizontal bars
    point,                                    // Data point variable declaration
    i;
  
  // Draw the graph
  Flotr.draw(
    container,
    [ <%=SalesXmlStr%> ],
    {
      bars : {
        show : true,
        horizontal : horizontal,
        shadowSize : 0,
        barWidth : 1
      },
	  xaxis: {
            ticks:[<%=UnitStr%>], // 自定义X轴
               minorTicks: null,
			   <%If KS.G("saletype")="year" then%>
			   title:'年份',   
			   <%else%>
			   title:'月份',   
			   <%end if%>
               showLabels:true                             // 是否显示X轴刻度
	  },
      mouse : {
        track : true,
        relative : true
      },
      yaxis : {
        min : 0,
		title:"营业额(单位:元)",
        autoscaleMargin : 1
      }
    }
  );
})(document.getElementById("container"));

    </script>
		
		
		
		  
        </td></tr>
      </table>
    </div></div><div class="footerTable"></div>
</div>
<%
End Sub

End Class
%> 

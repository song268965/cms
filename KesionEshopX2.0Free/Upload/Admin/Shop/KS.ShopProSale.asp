<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<% 

Dim KSCls
Set KSCls = New Admin_ShopProSale
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ShopProSale
        Private KS,Param,KSCls,id,page,totalput,maxperpage
		Private Action,i,strClass,RS,SQL
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		id=KS.ChkClng(ks.g("proid"))
		if request("action")="export" then 
		  export
		  response.end
		end if
		%>
		<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><link href="../Include/Admin_Style.CSS" rel="stylesheet" type="text/css"><script src="../../KS_Inc/common.js" language="JavaScript"></script><script language='JavaScript' src='../../KS_Inc/jquery.js'></script>
</head><body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
 <% if request("title")<>"" then%>
    <div style="margin:10px;">商品<font color=red>“<%=KS.CheckXSS(request("title"))%>”</font>的销售情况[总销售量：<font color=blue><%=KS.ChkClng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where i.proid=" & id)(0))%></font> 件，其中已付款：<font color=green><%=ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.MoneyReceipt>0 and i.proid=" & id)(0))%> </font>件，未付款：<font color=red><%=ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.MoneyReceipt<=0 and i.proid=" & id)(0))%></font> 件]以下是详细记录</div>
 <%end if%>
 <table width='100%'>
            <%
			dim ordertype:ordertype=0
			Page= KS.ChkClng(request("page"))
            If Page<=0 Then Page=1
            MaxPerPage=20
			 dim rso:set rso=server.createobject("adodb.recordset")
			 rso.open "select i.*,o.username,o.orderid,o.InputTime,o.MoneyReceipt,o.MoneyTotal,o.ordertype from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where i.proid=" & id & " order by i.id desc",conn,1,1
			 if rso.eof and rso.bof then
			   response.write "<tr><td style='text-align:center'>此商品还没有销售记录！</td></tr>"
			 else
				 TotalPut = rso.recordcount
				 If Page >1 and (Page - 1) * MaxPerPage < totalPut Then
						RSo.Move (Page - 1) * MaxPerPage
				 Else
						Page = 1
				 End If
			      
			      response.write "<tr  class='sort'><td>购买人</td><td>购买时间</td><td style='text-align:center'>数量</td><td style='text-align:center'>金额</td><td style='text-align:center'>付款状态</td></tr>"
				  dim ii:ii=0
				 do while not rso.eof
				   ordertype=rso("ordertype")
				   response.write "<tr>"
				   response.write " <td class='splittd'>" &rso("username") &"</td>"
				   response.write " <td class='splittd' align='center'>" &rso("inputtime") &"</td>"
				   response.write " <td class='splittd' style='text-align:center'>" &rso("amount") &"</td>"
				   response.write " <td class='splittd' style='text-align:center'><font color=blue>" &formatnumber(rso("MoneyTotal"),2,-1,-1) &"元</font></td>"
				   response.write "<td class='splittd'  style='text-align:center'>"
				    if rso("MoneyReceipt")>=rso("MoneyTotal") then
					  response.write "<font color=blue>已付款</font>"
					elseif rso("MoneyReceipt")=0 then
					  response.write "<font color=red>未付款</font>"
					else
					 response.write "已付定金"
					end if
				   response.write "</td>"
				   response.write "</tr>"
				   ii=ii+1
				   if ii>=maxperpage then exit do
				 rso.movenext
				 loop
			 end if
			 rso.close
			 set rso=nothing
	    response.write "<tr><td colspan=10>"
		Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
		response.write "</td></tr>"
		response.Write "</table>"
		
		response.write "<div style='text-align:center'><input type='button' class='button' value='导出到Excel' onclick=""window.open('KS.ShopProSale.asp?ordertype=" & ordertype & "&action=export&proid=" & id & "');""/></div>"
		
         End Sub

         sub export()
		    Response.AddHeader "Content-Disposition", "attachment;filename=sale.xls" 
			Response.ContentType = "application/vnd.ms-excel" 
			Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			 
			  dim title
			  if request("ordertype")="1" then
				  if conn.execute("select top 1 subject from ks_groupbuy where id=" & id).eof then
					 ks.die "商品已不存在！"
				  else
					 title=conn.execute("select top 1 subject from ks_groupbuy where id=" & id)(0)
				  end if
			  else
				  if conn.execute("select top 1 title from ks_product where id=" & id).eof then
					 ks.die "商品已不存在！"
				  else
					 title=conn.execute("select top 1 title from ks_product where id=" & id)(0)
				  end if
			  end if
			  response.write "<table width=""100%"" border=""1"">"
			 dim rso:set rso=server.createobject("adodb.recordset")
			 rso.open "select top 5000 i.*,o.username,o.orderid,o.InputTime,o.MoneyReceipt,o.MoneyTotal from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where i.proid=" & id & " order by i.id desc",conn,1,1
			 if rso.eof and rso.bof then
			   response.write "<tr><td style='text-align:center'>此商品还没有销售记录！</td></tr>"
			 else
				  dim ii:ii=0
			      response.write "<tr  class='sort'><th>序号</th><th>商品名称</th><th>购买人</th><th>购买时间</th><th style='text-align:center'>数量</th><th style='text-align:center'>金额</th><th style='text-align:center'>付款状态</th></tr>"
				 do while not rso.eof
				   ii=ii+1
				   response.write "<tr>"
				   response.write " <td class='splittd' style='text-align:center'>" & ii &"、</td>"
				   response.write " <td class='splittd'>" & title &"</td>"
				   response.write " <td class='splittd'>" &rso("username") &"</td>"
				   response.write " <td class='splittd' style='text-align:center'>" &rso("inputtime") &"</td>"
				   response.write " <td class='splittd' style='text-align:center'>" &rso("amount") &"</td>"
				   response.write " <td class='splittd' style='text-align:center'><font color=blue>" &formatnumber(rso("MoneyTotal"),2,-1,-1) &"元</font></td>"
				   response.write "<td class='splittd'  style='text-align:center'>"
				    if rso("MoneyReceipt")>=rso("MoneyTotal") then
					  response.write "<font color=blue>已付款</font>"
					elseif rso("MoneyReceipt")=0 then
					  response.write "<font color=red>未付款</font>"
					else
					 response.write "已付定金"
					end if
				   response.write "</td>"
				   response.write "</tr>"
				 rso.movenext
				 loop
			 end if
			 rso.close
			 set rso=nothing
		response.Write "</table>"
			
		 end sub
End Class
%> 

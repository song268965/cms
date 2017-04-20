<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
If Request.ServerVariables("HTTP_REFERER")="" Then Response.End
Dim KS,Action
Set KS=New PublicCls
Dim ProID,Stock,RS,SqlStr
ProID = KS.ChkClng(KS.S("ID"))
If ProID = 0 Then
        Stock = 0
Else
      action=KS.S("Action")
	  If Action="GroupBuyHasSold" Then
	    SqlStr = "select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.ProID=" & ProID
	  ElseIf Action="HasSold" Then
	    SqlStr = "select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where (o.status=1 or o.status=2) and i.ProID=" & ProID
	  Else
        SqlStr = "SELECT top 1 TotalNum FROM KS_Product Where ID=" & ProID
	  End If
	  
        Set RS = Server.CreateObject("ADODB.Recordset")
        RS.Open SqlStr, conn, 1, 1
        If RS.bof And RS.EOF Then
            Stock = 0
        Else
            Stock = rs(0)
        End If
        rs.Close
        Set rs = Nothing
		if action="GroupBuyHasSold" then
		 Stock=KS.ChkClng(Stock)+KS.ChkClng(conn.execute("select top 1 HasBuyNum from ks_groupbuy where id=" & proid)(0))
		end if
End If
Response.Write "document.write('" & KS.ChkClng(Stock) & "');"
Call CloseConn()
Set KS=Nothing	
%> 

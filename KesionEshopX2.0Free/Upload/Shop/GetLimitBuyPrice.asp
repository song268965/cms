<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KS
Set KS=New PublicCls
Dim ProID,Tips,RS,SqlStr
ProID = KS.ChkClng(KS.S("ProID"))
 If ProID = 0 Then
 Else
        SqlStr = "SELECT top 1 I.IsLimitbuy,I.LimitBuyPrice,I.LimitBuyAmount,I.Unit,L.LimitBuyBeginTime,L.LimitBuyEndTime FROM KS_Product I Inner Join KS_ShopLimitBuy L ON I.LimitBuyTaskID=L.id Where I.IsLimitbuy<>0 and I.ID=" & ProID
        Set RS = Server.CreateObject("ADODB.Recordset")
        RS.Open SqlStr, conn, 1, 1
        If Not RS.bof  Then
		   If RS(0)="1" Then
            Tips = "限时抢购价：￥ <font color=""red"">" & KS.GetPrice(rs(1)) & "</font> 元<br/>"
		   Else
            Tips = "限量抢购价：￥ <font color=""red"">" & KS.GetPrice(rs(1) & "</font> 元"
		   End If
		   
		    Tips=Tips & "<span style=""font-size:12px;color:#ff3300"">"
		   If RS(0)="1" And  (Now>RS(5)) Then
		    Tips=Tips & "(抢购时间已过了"
		   Else
		    if rs(2)>0 then
				Tips = Tips & " (还剩:" & RS(2) & RS(3) 
				If RS(0)="1" Then
				 Tips=Tips & " 活动时间:" & RS(4) & "至" & RS(5)
				End If
			else
		    Tips = Tips & " (抢购结束" 
			end if
		   End If
			Tips = Tips & ")</span>"
        End If
        rs.Close
        Set rs = Nothing
End If
Response.Write "document.write('" & Tips & "');"

Call CloseConn()
Set KS=Nothing	
%>

<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS
Set KS=New PublicCls
Dim ID,Tips,RS,SqlStr,str,url,Num,TaskType,PhotoUrl,limitbuyBegintime,limitbuyendtime
Num=KS.ChkCLng(KS.S("Num"))
If Num=0 Then Num=2
ID = KS.ChkClng(KS.S("ID"))
        '====================================检测没有到达最后支付期限,没有支付的订单============================
		sqlstr="Select i.proid,i.amount,i.OrderID,o.paytime,L.LimitBuyPayTime,O.status,O.InputTime From (KS_OrderItem i inner join ks_order o on i.orderid=o.orderid) inner join KS_ShopLimitBuy L on I.LimitBuyTaskID=L.ID Where (L.LimitBuyPayTime)<>0 and i.IsLimitBuy<>0 and o.status=0 and o.MoneyReceipt=0"
		
		
        Set RS = Server.CreateObject("ADODB.Recordset")
		RS.Open SQLStr,conn,1,3
		Do While Not RS.Eof
		  if KS.ChkClng(rs("limitbuypaytime"))<>0 And Now>DateAdd("h",rs("limitbuypaytime"),RS("InputTime")) then
		   RS("status")=3   '设成无效定单
		   RS.Update
		   Conn.Execute("Update KS_Product Set LimitBuyAmount=LimitBuyAmount+" & rs("amount") & ",TotalNum=TotalNum+" & rs("amount") &" where id=" & rs("proid"))  '增加可抢购数
		  end if
		 RS.MoveNext
		Loop
		RS.Close
		
		'==========================================================================================================
        '抢购结束的商品恢复状态
		 conn.execute("update ks_product set islimitbuy=0,LimitBuyTaskID=0 where LimitBuyTaskID in(select id from KS_ShopLimitBuy where datediff(" & DataPart_S &",LimitBuyEndTime," & SqlNowString&")>0)")
		 
 
        SqlStr = "SELECT  top " & Num & " I.ID,I.Title,I.Fname,i.photourl,I.Tid,I.Price,I.Price_member,i.unit,I.IsLimitbuy,I.LimitBuyPrice,I.LimitBuyAmount,I.Unit,L.LimitBuyBeginTime,L.LimitBuyEndTime,L.TaskType,I.AddDate FROM KS_Product I Inner Join KS_ShopLimitBuy L ON I.LimitBuyTaskID=L.id Where L.Status=1"
		If Id<>0 Then SqlStr=SqlStr & " and L.id=" & ID Else SqlStr=SqlStr & " order by l.id desc"
        RS.Open SqlStr, conn, 1, 1
        If Not RS.bof  Then 
		    ' str="<ul>"
		    do while not RS.Eof
			 url=KS.GetItemURL(5,RS("Tid"),rs("ID"),RS("Fname"),RS("AddDate"))
			 photourl=rs("photourl")
			 if ks.isnul(photourl) then photourl=KS.GetDomain & "images/nopic.gif"
			 limitbuyBegintime=rs("limitbuyBegintime")
			 limitbuyendtime=rs("limitbuyendtime")
			 TaskType=rs("TaskType")
			 dim sysl:sysl=rs("limitbuyamount")
			 dim strtips
			 if sysl<=0 then 
			   strtips=" href=""#"" onclick=""alert('抢购结束!');return false"" style=""cursor:pointer"""
			 else
			   strtips=" href=""" & url &""""
			 end if
	        str=str & "<ul><img class=""pho"" onerror=""this.src='" & KS.GetDomain & "images/nopic.gif';"" src=""" & PhotoUrl & """ /><img class=""qiang"" src=""" & KS.GetDomain & "images/icon_qiang.png""><br/><p><a target=""_blank""" & strtips & " class=""limitbuytitle"" title=""" & rs("title") & """>" & rs("title") & "</a></p><p class=""mery"">￥" & FormatNumber(rs("LimitBuyPrice"),2,-1) & "</p><p>原价:<s>" & FormatNumber(rs("price"),2,-1) & "</s></p><div class=""qianggou""><a target=""_blank""" & strtips & ">立即抢购</a></div></ul>"

						 
			 
			RS.MoveNext
			Loop
			'str=str & "</ul>"
			
			'response.write "var t" & id & "=new calculagraph();t"&id&"._tasktype=" & tasktype&";t" & id & "._id='time" & id & "';t" & id & "._sT='" & year(limitbuyBegintime)&"-" & month(limitbuyBegintime) & "-" & day(limitbuyBegintime) & " " & hour(limitbuyBegintime) & ":" & minute(limitbuyBegintime) & ":" & second(limitbuyBegintime) & "';  t" & id & "._cT='" & year(now)&"-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & "';  t" & id & "._eT='" & year(limitbuyendtime)&"-" & month(limitbuyendtime) & "-" & day(limitbuyendtime) & " " & hour(limitbuyendtime) & ":" & minute(limitbuyendtime) & ":" & second(limitbuyendtime) &"';   t" & id & "._interval();"
			
			 response.write "show_date_time('" & LimitBuyBeginTime &"','" & LimitBuyEndTime &"','time" &id &"',1);"
			if ks.s("from")="script" then '脚本调用，可以跨域调用
			response.write "document.write('" & replace(str,"'","\'") & "');" &vbcrlf
			response.write "document.write('<div class=""timeBox"" id=""time" & id & """>正在加载…</div>');"&vbcrlf
            else
			 response.write "|"
			 response.write escape(str)
			end if
        End If
        rs.Close
        Set rs = Nothing
Call CloseConn()
Set KS=Nothing	
%>


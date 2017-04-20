<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.StaticCls.asp"-->
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
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing


Class SiteIndex
        Private KS, KSR,ID,Template,RS,Action,KeyWord
		Private MaxPerPage,Page,totalPut,PageNum,XML,Node
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   Action=KS.S("Action")
		   Template = KSR.LoadTemplate(KS.Setting(173))
		   FCls.RefreshType = "ordersearch" '设置刷新类型，以便取得当前位置导航等
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   If Action="dosearch" Then
		     dosearch
		   End If
		   Template=KSR.KSLabelReplaceAll(Template)
		   Response.Write Template  
		End Sub
		
		Sub DoSearch()
		  KeyWord=KS.CheckXSS(KS.R(KS.S("OrderID")))
		  If KS.IsNul(KeyWord)  Then KS.AlertHintScript "请输入订单号或是客户姓名！"
		  Dim SqlStr,RS,Str,Total,n
		  SQLStr="select TOP 10 * From KS_Order Where OrderID='" & KeyWord & "' or ContactMan='" & KeyWord & "' Order By ID Desc"
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open (SQLStr),conn,1,1
		  If RS.Eof And RS.Bof Then
		   Str="<div class=""searchtips"">对不起，您输入的关键字<span style='color:red'>""" & keyword & """</span>没有找到符合要求的订单！</div>"
		  Else
		     str=str &"<div class=""searchtips"">订单号<span style='color:red'>""" & keyword & """</span>，查询结果如下：</div>"
			Total= rs.recordcount
			n=1
		    Do While Not RS.Eof
			  if Total>1 then str=str & "<strong>第 <font color=red>" & n & "</font> 条订单：</strong>"
			  Str=Str &  OrderDetailStr(RS) & "<br/>"
			 
			  str=str &"<table id='payment'><tr><td width='280' nowrap>"
			 If RS("Status")=3 Then
			   Str=Str & "本订单在指定时间内没有付款,已作废!"
			 Else 
				 If RS("MoneyReceipt")<RS("MoneyTotal") Then
				 str=str &"" & PayMentStr(rs("id"),(RS("MoneyTotal")-RS("MoneyReceipt"))) &"</td><td>"
				 str=str & "<input class=""btn"" type='button' name='Submit' value='从余额中扣款支付' onClick=""window.location.href='../user/User_Order.asp?Action=AddPayment&ID=" & rs("id") & "0'"">"
				end if
				if rs("DeliverStatus")=1 Then
				 str=str & "<td><input class=""btn"" type='button' name='Submit' value='签收商品' onClick=""window.location.href='../user/User_Order.asp?Action=SignUp&ID=" & RS("ID") & "'""></td>"
				end if
		
		   end if

		   str=str &"<input class=""btn"" type='button' name='Submit' value='打印订单' onClick=""window.print();""></td></tr></table>"
		   str=str &"<br/>"
			  n=n+1
			RS.MoveNext
			Loop
		  End If
		  RS.Close:Set RS=Nothing
		  Template=Replace(Template,"{$ShowOrderSearchResult}",str)
		End Sub
		
		Function PayMentStr(OrderAutoID,RealMoneyTotal)
		   Dim SQL,K,Param,PayStr
		  Set RS=Server.CreateOBject("ADODB.RECORDSET")
					   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 Order By OrderID",conn,1,1
					   If Not RS.Eof Then SQL=RS.GetRows(-1)
					   RS.Close:Set RS=Nothing
					   If Not IsArray(SQL) Then
						PayStr=""
					   Else
					     PayStr="<form name=""payform" & OrderAutoID &""" method=""get"" action=""payonline.asp"" target=""_blank"">"
						 PayStr=PayStr & "<input type=""hidden"" name=""id"" value=""" & OrderAutoID & """/>"
						 PayStr=PayStr & "<strong>选择支付平台：</strong><select name=""PaymentPlat"" id=""PaymentPlat""><option value=''>---请选择支付平台---</option>"
						 For K=0 To Ubound(SQL,2)
						   PayStr=PayStr & "<option value='" & SQL(0,K) & "' name='PaymentPlat'"
						   If trim(SQL(3,K))="1" Then PayStr=PayStr &  " selected"
						   PayStr=PayStr &  ">"& SQL(1,K) & "</option>"
						   'PayStr=PayStr &  ">"& SQL(1,K) & "(" & SQL(2,K) &")</option>"
						 Next
					   End If
					   PayStr=PayStr & "</select><div style='padding-left:80px;margin:10px;'>"
					  
					   Dim PArr:Parr=Split(KS.Setting(82)&"||||||||","|")
					  If Parr(0)="1" Then
					   PayStr=PayStr & "<input class=""btn"" type=""submit"" onclick=""if(document.payform" & OrderAutoID &".PaymentPlat.value==''){alert('请选择支付平台!');return false;}"" style=""padding:2px"" value="" 进入支付平台在线支付 "" />"
					  ElseIf Parr(0)="2" Then
					   PayStr=PayStr & "<td><input class=""btn"" type=""submit"" onclick=""if(document.payform" & OrderAutoID &".PaymentPlat.value==''){alert('请选择支付平台!');return false;}""  value="" 在线支付" & Parr(1) & "元定金 "" />"
					  Else 
					   PayStr=PayStr & "支付金额:<input type='text' size='8' name='money' value='" & RealMoneyTotal & "'/> 元<br/><input type=""submit"" style=""padding:2px"" class=""btn"" value="" 确认在线支付 "" onclick=""if(document.payform" & OrderAutoID &".PaymentPlat.value==''){alert('请选择支付平台!');return false;}""/>"
					  End If 
					  
					   PayStr=PayStr & "</div>"
					   PayStr=PayStr & "</form>"
					   PayMentStr=PayStr
		End Function
		
		'返回订单详细信息
		Function  OrderDetailStr(RS)
		 OrderDetailStr="<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr align='center' class='title'>    <td height='22'><b>订 单 信 息</b>（订单编号：" & RS("ORDERID") & "）</td>  </tr>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr>" & vbcrlf
		 OrderDetailStr=OrderDetailStr & " <td height='25'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "  <table width='100%'  border='0' cellpadding='2' cellspacing='0'> "   & vbcrlf
		 OrderDetailStr=OrderDetailStr & "    <tr class='tdbg'>"
		 OrderDetailStr=OrderDetailStr & "	         <td width='18%'>客户姓名：<font color='red'>" & RS("Contactman") & "</td>      "
		 OrderDetailStr=OrderDetailStr & "			 <td width='20%'>用 户 名：<font color='red'>" & rs("username") & "</td> " &vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='20%'>代 理 商：</td>"
		OrderDetailStr=OrderDetailStr & "			 <td width='18%'>购买日期：<font color='red'>" & formatdatetime(rs("inputtime"),2) & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='24%'>下单时间：<font color='red'>" & formatdatetime(rs("inputtime"),2) & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "	</tr>"
		OrderDetailStr=OrderDetailStr & "	<tr class='tdbg'> "      
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>需要发票："
			    If RS("NeedInvoice")=1 Then
				  OrderDetailStr=OrderDetailStr & "<Font color=red>√</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "		 </td> "       
		OrderDetailStr=OrderDetailStr & "	 <td width='20%'>已开发票："	
				  If RS("Invoiced")=1 Then
				   OrderDetailStr=OrderDetailStr & "<font color=green>√</font>"
				  Else
				   OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td> "
		OrderDetailStr=OrderDetailStr & "	<td width='20%'>订单状态："	
			if RS("Status")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>等待确认</font>"
				  ElseIf RS("Status")=1 Then
				 OrderDetailStr=OrderDetailStr & "<font color=green>已经确认</font>"
				  ElseIf RS("Status")=2 Then
				 OrderDetailStr=OrderDetailStr & "<font color=#a7a7a7>已结清</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td>"
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>付款情况："	
			     If RS("MoneyReceipt")<=0 Then
				   OrderDetailStr=OrderDetailStr & "<font color=red>等待汇款</font>"
				  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>已收定金</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=green>已经付清</font>"
				  End If

       OrderDetailStr=OrderDetailStr & "</td>"
	   OrderDetailStr=OrderDetailStr & "        <td width='24%'>物流状态："
				if RS("DeliverStatus")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>未发货</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已发货</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已签收</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  OrderDetailStr=OrderDetailStr & "<font color=#ff6600>退货</font>"
				 End If
	OrderDetailStr=OrderDetailStr & "		</td></tr>    </table> "
    OrderDetailStr=OrderDetailStr & " </td>  </tr> " 
	OrderDetailStr=OrderDetailStr & "   <tr align='center'>"
	OrderDetailStr=OrderDetailStr & "       <td height='25' style='text-align:left'>"
	OrderDetailStr=OrderDetailStr & "	   <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
	OrderDetailStr=OrderDetailStr & "	           <tr class='tdbg'>"
	OrderDetailStr=OrderDetailStr & "			             <td width='12%' align='right'>收货人姓名：</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("contactman") & "</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='12%' align='right'>联系电话：</td> "      
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("phone") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>"
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>"
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>收货人地址：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("address") & "</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>邮政编码：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" &rs("zipcode") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>  "      
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>收货人邮箱：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("email") & "</td> "         
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>收货人手机：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("mobile") & "</td>       "
	OrderDetailStr=OrderDetailStr & "			   </tr>"        
	OrderDetailStr=OrderDetailStr & "			   <tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "			              <td width='12%' align='right'>付款方式：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & KS.ReturnPayMent(rs("PaymentType"),0) & "</td>"   
OrderDetailStr=OrderDetailStr & "						  <td colspan='2' width='38%'>快递公司：" 
	
	  dim rst,foundexpress
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
	    OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
	  If DataBaseType=1 Then
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
	  Else
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
	  End If
	  if rst.eof then
	  else
	   OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	  end if
	 rst.close : set rst=nothing
	 End If
	
	
	OrderDetailStr=OrderDetailStr & " 发往<span style='color:red'>" & rs("tocity") & "</span></td>"
	OrderDetailStr=OrderDetailStr & "				</tr> "       
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>  "        
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>发票信息：</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>"
	 If RS("Invoiced")=1 Then OrderDetailStr=OrderDetailStr & rs("InvoiceContent") &"</td>"
    OrderDetailStr=OrderDetailStr & "						 <td width='12%' align='right'>备注/留言：</td>"          
	OrderDetailStr=OrderDetailStr & "							<td width='38%'>" & rs("Remark") & "</td>       "
	OrderDetailStr=OrderDetailStr & "				 </tr>  "  
	OrderDetailStr=OrderDetailStr & "				 </table>"
	OrderDetailStr=OrderDetailStr & "			</td>  "
	OrderDetailStr=OrderDetailStr & "		</tr>  "
	OrderDetailStr=OrderDetailStr & "		<tr><td>"
	OrderDetailStr=OrderDetailStr & "		<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "
	OrderDetailStr=OrderDetailStr & "		  <tr align='center' class='title' height='25'>  "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>商 品 名 称</b></td> "   
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>单位</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='55'><b>数量</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>原价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>实价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>指定价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='85'><b>金 额</b></td>   " 
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>服务期限</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>备注</b></td>  "
	OrderDetailStr=OrderDetailStr & "		  </tr> "
			 Dim TotalPrice,attributecart,RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			   RSI.Open "Select * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id",conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
				' Response.Write "<script>alert('找不到相关商品');history.back();<//script>"
			  Else
			  Do While Not RSI.Eof
			  attributecart=rsi("attributecart")
			  if not ks.isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"  
		If rs("OrderType")=1 Then
		OrderDetailStr=OrderDetailStr & "	   <td width='*'><a href='" & KS.Setting(3) & "shop/groupbuyshow.asp?id=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 subject from ks_groupbuy where id=" & rsi("proid"))(0) 
		Else  
		OrderDetailStr=OrderDetailStr & "	   <td width='*'><a href='" & KS.Setting(3) & "item/show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 title from ks_product where id=" & rsi("proid"))(0) 
		End If
		
		If RSI("IsChangedBuy")="1" Then OrderDetailStr=OrderDetailStr & "(换购)"
		
		
			  Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  RSP.Open "Select top 1 I.Title,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID"),conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(换购)"
				  else
				    if isdate(RSP("LimitBuyPayTime")) then
					   If LimitBuyPayTime="" Then
					   LimitBuyPayTime=RSP("LimitBuyPayTime")
					   ElseIf LimitBuyPayTime>RSP("LimitBuyPayTime") Then
						LimitBuyPayTime=RSP("LimitBuyPayTime")
					   End If
					end if
				  end  if
				  If RSI("IsLimitBuy")="1" Then OrderDetailStr=OrderDetailStr & "<span style='color:green'>(限时抢购)</span>"
				  If RSI("IsLimitBuy")="2" Then OrderDetailStr=OrderDetailStr & "<span style='color:blue'>(限量抢购)</span>"
			  End If
			  RSP.Close:Set RSP=Nothing
		
		OrderDetailStr=OrderDetailStr & "</a>" & attributecart & "</td>    "
		If RS("OrderType")="1" Then
		OrderDetailStr=OrderDetailStr & "	   <td width='45' align=center>件</td>"
		Else
		OrderDetailStr=OrderDetailStr & "	   <td width='45' align=center>"& Conn.execute("select unit from ks_product where id=" & rsi("proid"))(0) & "</td>"
	    End If
		OrderDetailStr=OrderDetailStr & "<td width='55' align='center'>" & rsi("amount") &"</td>"
		
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("price_original"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("price"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("realprice"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='85' align='right'>" & formatnumber(rsi("realprice")*rsi("amount"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align=center>" & rsi("ServiceTerm") & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td align=center width='45'>" & rsi("Remark") & "</td>  "
		OrderDetailStr=OrderDetailStr & "	   </tr> " 
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  '取得捆绑销售商品
		
		
			  TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
		End If
		
		OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '超值礼包
		
		
		OrderDetailStr=OrderDetailStr & "	   <tr class='tdbg' height='30' > "   
		OrderDetailStr=OrderDetailStr & "	    <td colspan='6' align='right'><b>合计：</b></td> "   
		OrderDetailStr=OrderDetailStr & "		<td align='right'><b>" & formatnumber(totalprice,2) & "</b></td>    "
		OrderDetailStr=OrderDetailStr & "		<td colspan='3'> </td>  "
		OrderDetailStr=OrderDetailStr & "	  </tr>    "
		OrderDetailStr=OrderDetailStr & "	  <tr class='tdbg'>"
       OrderDetailStr=OrderDetailStr & "         <td colspan='4'>付款方式折扣率：" & rs("Discount_Payment") & "%&nbsp;&nbsp;" 
	   If RS("Weight")>0 Then
	   OrderDetailStr=OrderDetailStr & "重量：" & rs("weight") & " KG"
	   End If
	   OrderDetailStr=OrderDetailStr & "&nbsp;&nbsp;运费：" & rs("Charge_Deliver")&" 元&nbsp;&nbsp;&nbsp;&nbsp;税率：" & KS.Setting(65) &"%&nbsp;&nbsp;&nbsp;&nbsp;价格含税："
				IF KS.Setting(64)=1 Then 
				   OrderDetailStr=OrderDetailStr & "是"
				  Else
				   OrderDetailStr=OrderDetailStr & "不含税"
				  End If
				  Dim TaxMoney
				  Dim TaxRate:TaxRate=KS.Setting(65)
				 If KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+TaxRate/100

				OrderDetailStr=OrderDetailStr & "<br>实际金额：(" & rs("MoneyGoods") & "×" & rs("Discount_Payment") & "%＋"&rs("Charge_Deliver") & ")×"
				if KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then OrderDetailStr=OrderDetailStr & "100%" Else OrderDetailStr=OrderDetailStr & "(1＋" & TaxRate & "%)" 
				OrderDetailStr=OrderDetailStr & "＝" & formatnumber(rs("NoUseCouponMoney"),2) & "元  </td>"
    OrderDetailStr=OrderDetailStr & "<td  colspan='3' align=right><b>订单金额：</b> ￥" & formatnumber(rs("NoUseCouponMoney"),2) & " 元<br>"
	If KS.ChkClng(RS("CouponUserID"))<>0 and RS("UseCouponMoney")>0 Then
	OrderDetailStr=OrderDetailStr & "<b>使用优惠券：</b> <font color=#ff6600>￥" & formatnumber(RS("UseCouponMoney"),2) & " 元</font><br>"
	End If
	OrderDetailStr=OrderDetailStr & "<b>应付金额：</b> ￥" & formatnumber(rs("MoneyTotal"),2) & "  元</td>"
    OrderDetailStr=OrderDetailStr & "<td colspan='3' align='left'><b>已付款：</b>￥<font color=red>" & formatnumber(rs("MoneyReceipt"),2) & "</font></b>"
	If RS("MoneyReceipt")<RS("MoneyTotal") Then
	OrderDetailStr=OrderDetailStr & "<br><B>尚欠款：￥<font color=blue>" & formatnumber(RS("MoneyTotal")-RS("MoneyReceipt"),2) &"</B>"
	End If
	OrderDetailStr=OrderDetailStr & "</td></tr></table></td>  "
	OrderDetailStr=OrderDetailStr & "</tr>"  
	OrderDetailStr=OrderDetailStr & "     <tr><td><br><b>注：</b>“<font color='blue'>原价</font>”指商品的原始零售价，“<font color='green'>实价</font>”指系统自动计算出来的商品最终价格，“<font color='red'>指定价</font>”指管理员根据不同会员组手动指定的最终价格。商品的最终销售价格以“指定价”为准。“订单金额”指系统自动算出来的价格，本订单的最终价格以“<font color=#ff6600>应付金额</font>”为准。<br>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	
	If not conn.execute("select top 1 * from ks_orderitem where orderid='" & RS("OrderID") &"' and islimitbuy<>0").eof Then
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:red;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单是限时/限量抢购订单,限制下单后" & LimitBuyPayTime & "小时之内必须付款,即如果您在[" & DateAdd("h",LimitBuyPayTime,RS("InputTime")) & "]之前用户没有付款,本订单自动作废。</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	If RS("DeliverStatus")=1 Then
	 Dim RSD,DeliverStr
	 Set RSD=Conn.Execute("Select Top 1 * From KS_LogDeliver Where DeliverType=1 And OrderID='" & RS("OrderID") & "'")
	 If Not RSD.Eof Then
	  DeliverStr="快递公司:" & RSD("ExpressCompany") & " 物流单号:" & RSD("ExpressNumber") & " 发货日期:" & RSD("DeliverDate") & " 发货经手人:" & RSD("HandlerName")
	 End If
	 RSD.Close : Set RSD=Nothing
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:blue;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单已发货。" & DeliverStr & "</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	
	OrderDetailStr=OrderDetailStr & "	</table>"
	  End Function
	  
'取得捆绑销售商品
Function GetBundleSalePro(ByRef TotalPrice,ProID,OrderID)
  Dim Str,RS,XML,Node
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "Select I.Title,I.Unit,O.* From KS_OrderItem O inner join KS_Product I On O.ProID=I.ID Where O.SaleType=6 and BundleSaleProID=" & ProID & " and OrderID='" & OrderID & "' order by O.id",conn,1,1
  If Not RS.Eof Then
    Set XML=KS.RsToXml(rs,"row","")
  End If
  RS.Close:Set RS=Nothing
  If IsObject(XML) Then
	     str=str & "<tr height=""25"" align=""left""><td colspan=9 style=""color:green"">&nbsp;&nbsp;选购捆绑促销:</td></tr>"
       For Each Node In Xml.DocumentElement.SelectNodes("row")
         str=str & "<tr>"
		 str=str &" <td style='color:#999999'>&nbsp;" & Node.SelectSingleNode("@title").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@unit").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@amount").text &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@price_original").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td align='right'>" & formatnumber(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@serviceterm").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@remark").text &"</td>"
		 str=str & "</tr>"
		 TotalPrice=TotalPrice +round(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2) 
       Next
  End If
  GetBundleSalePro=str
End Function
	  
	  
 '得到超值礼包
 Function GetPackage(ByRef TotalPrice,OrderID)
	    If KS.IsNul(OrderID) Then Exit Function
		Dim RS,RSB,GXML,GNode,str,n,Price
		Set RS=Conn.Execute("select packid,OrderID from KS_OrderItem Where SaleType=5 and OrderID='" & OrderID & "' group by packid,OrderID")
		If Not RS.Eof Then
		 Set GXML=KS.RsToXml(Rs,"row","")
		End If
		RS.Close : Set RS=Nothing
		If IsOBJECT(GXml) Then
		   FOR 	Each GNode In GXML.DocumentElement.SelectNodes("row")
		     Set RSB=Conn.Execute("Select top 1 * From KS_ShopPackAge Where ID=" & GNode.SelectSingleNode("@packid").text)
			 If Not RSB.Eof Then
					  
						Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
						RSS.Open "Select a.title,a.Price_Member,a.Price,b.* From KS_Product A inner join KS_OrderItem b on a.id=b.proid Where b.SaleType=5 and b.packid=" & GNode.SelectSingleNode("@packid").text & " and  b.orderid='" & OrderID & "'",Conn,1,1
						  str=str & "<tr class='tdbg' height=""25"" align=""center""><td colspan=2><strong><a href='" & KS.Setting(3) & "shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a></strong></td>"
						  n=1
						  Dim TotalPackPrice,tempstr,i
						  TotalPackPrice=0 : tempstr=""
						Do While Not RSS.Eof
						 
						  For I=1 To RSS("Amount") 
							  '得到单件品价格 
							  IF KS.C("UserName")<>"" Then
								  Price=RSS("Price_Member")
							  Else
								  Price=RSS("Price")
							  End If
							
							   TotalPackPrice=TotalPackPrice+Price
							  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
							  n=n+1
						  Next
						  RSS.MoveNext
						Loop
						
						str=str &"<td>1</td><td>￥" & TotalPackPrice & "</td><td>" & rsb("discount") & "折</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>---</td><td>---</td>"
					   
						str=str & "</tr><tr><td align='left' colspan=9>您选择的套装详细如下:<br/>" & tempstr & "</td></tr>" 
						
						TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1))   '将礼包金额加入总价
						
						RSS.Close
						Set RSS=Nothing
					
			End If
			RSB.Close
		   Next
			
	    End If
		GetPackage=str
		
End Function
		
End Class
%>

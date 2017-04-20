<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New GroupBuyCart
KSCls.Kesion()
Set KSCls = Nothing

Class GroupBuyCart
        Private KS, KSR,KSUser,Product,LoginTf,TotalPrice,totalweight,MustPayOnline
		Private GroupBuy,K,ID,FileContent,IsSuccess,OrderID,OrderAutoID,RealMoneyTotal,DeliveryStr,TypeID,ToCity,DeliveryMoney
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		  IsSuccess=false              '标志订单提交成功，模板切换
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		  LoginTF=Cbool(KSUser.UserLoginChecked)
		  Call showmain()
		End Sub
		%>
		<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
		<%
		Sub ShowMain()
		     ID=KS.ChkClng(Request("ID"))
			 If Id<>0 Then Call AddToCart()
			 FileContent = KSR.LoadTemplate(KS.Setting(120)) 
			 FCls.RefreshType = "groupbycart" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			 LoadCartList
			 MustPayOnline=true
			 If IsArray(Product) Then
			     TotalPrice=0:totalweight=0
			     For K=0 To Ubound(Product,2)
				  TotalPrice=TotalPrice+Product(4,k)*Product(2,k)
				  totalweight=totalweight+Product(5,k)*Product(2,k)
				 Next 
				 
				if Product(12,0)="1" and Product(14,0)="0" then FileContent=Replace(FileContent,"{$ShowDelivery}" ," style='display:none'")
			    if (Product(13,0)="0" and  Product(12,0)="1") or (ubound(product,2)=0 and Product(13,0)="0") then 
				 FileContent=Replace(FileContent,"{$ShowPaymentType}" ," style='display:none'")
				 MustPayOnline=false
				End If
			End If

			 'FileContent=RexHtml_IF(FileContent)
			 Select Case KS.S("Action") 
			  case "order"  Call OrderSave()
			  case "del" Call Del()
			  case else
			    DeliveryStr=GetDeliveryTypeStr
			 End Select
			 
			 Immediate=false
			 Scan FileContent
			 Templates=KSR.KSLabelReplaceAll(Templates)
			 Response.write RexHtml_IF(Templates)
		End Sub
		
		
		Sub ParseArea(sTokenName, sTemplate)
			Select Case lcase(sTokenName)
			 case "cart"
			  If IsArray(Product) Then
			     TotalPrice=0:totalweight=0
			     For K=0 To Ubound(Product,2)
				  Scan sTemplate
				  TotalPrice=TotalPrice+Product(4,k)*Product(2,k)
				  totalweight=totalweight+Product(5,k)*Product(2,k)
				 Next 
			  End If
			End Select 
        End Sub 
		
        Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			   case "groupbuy"
			         Select case lcase(sTokenName)
					  case "todaygroupbuylink"  
					   If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/" Else Echo KS.GetDomain & "shop/groupbuy.asp"
					  case "historygroupbuylink"  
					   If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/history/" Else Echo KS.GetDomain & "shop/groupbuy.asp?flag=history"
					 End Select
			   case "product"
			       Select  case lcase(sTokenName)
				      case "linkurl" If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/show-" &Product(1,k) & ".html"  Else Echo KS.GetDomain & "shop/groupbuyshow.asp?id=" &Product(1,k)
					  case "cartid" echo Product(0,k)
					  case "proid" echo Product(1,k)
					  case "name" echo Product(3,k)
					  case "amount" echo Product(2,k)
					  case "cartamount"
					   if Product(7,k)="1" then
					    echo Product(2,k)
						echo "<input type=""hidden"" name=""amount" & Product(0,k) & """ value=""1""/>"
					   else
					    echo "<input class=""textbox"" onchange=""changenum(" & Product(0,k) &",this.value);"" type=""text"" name=""amount" & Product(0,k) & """ id=""amount" & Product(0,k) & """ value=""" & Product(2,k) & """ size=""5"" style=""text-align:center""/>"
					   end if
					  case "price" echo Product(4,k)
					  case "weight" echo Product(5,k)
					  case "totalprice" echo Product(4,k)*Product(2,k)
					 end select
			  case "cart"
			      Select  case lcase(sTokenName)
				    case "userid" echo GetUserID
				    case "totalprice" echo FormatNumber(TotalPrice+DeliveryMoney,2,-1,-1)
					case "deliverytype" echo deliverystr
					case "paymenttype" echo GetPayTypeStr
					case "contactman" echo KSUser.GetUserInfo("realname")
					case "address" echo KSUser.GetUserInfo("address")
					case "zipcode" echo KSUser.GetUserInfo("zip")
					case "mobile" echo KSUser.GetUserInfo("mobile")
					case "orderid" echo orderid
					case "orderautoid" echo OrderAutoID
					case "paytotalmoney" echo RealMoneyTotal
					case "paymentplat" echo ks.s("paymentplat")
					case "freight"  echo DeliveryMoney
					case "orderdetail" echo OrderDetail	 
				  End Select
		    End Select 
        End Sub 
		
		'显示订单详细信息
		Function OrderDetail()
		  dim str
		  if MustPayOnline=true then
		    str="订单金额：<font color=""#FF0000""><strong>￥" & RealMoneyTotal & "</strong></font>元<br />付款成功后，才能完成本次交易，请尽快付款 ！<br />"
			Dim PArr:Parr=Split(KS.Setting(82)&"||||||||","|")
			 If Parr(0)="1" Then
				 str=str &"<input type=""submit"" style=""padding:2px"" value="" 进入支付平台在线支付 "" />"
			 ElseIf Parr(0)="2" Then
				  str=str &"<input type=""submit"" style=""padding:2px"" value="" 在线支付" & Parr(1) & "%的定金 "" />"
			 Else 
				str=str &"支付金额:<input type='text' size='8' style=""text-align:center"" name='money' value='" & RealMoneyTotal & "'/> 元<br/>"
			    str=str &"<input type=""submit"" value=""进入在线支付平台支付""/>"
			 End If 
		  else
		    str="凭订单号消费，请保留好订单号。<br/><input onclick=""print_ele_f('myorderid')"" type=""button"" value=""打印订单号""/>"
		  end if
		  OrderDetail=str
		End Function
		
		'用户名
		Function GetUserID()
			 If KS.IsNul(KS.C("UserName")) Then
					If Not KS.IsNul(KS.C("CartID")) Then
					 GetUserID=KS.C("CartID")
					Else
					 Response.Cookies(KS.SiteSn)("CartID")=KS.MakeRandom(15)
					 GetUserID=KS.C("CartID")
					End If
			 Else
					GetUserID=KS.C("UserName")
			 End If
		End Function  
		
		Sub AddToCart()
		    '删除大于3天的购物车记录
			Conn.Execute("Delete From KS_ShoppingCart Where flag=1 and datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>3")
			
		    Dim amount:amount=KS.Chkclng(Request("num"))
			If amount=0 then amount=1
		    dim rs:set rs=server.CreateObject("adodb.recordset")
			rs.open "select top 1 * from KS_GroupBuy Where id=" & id,conn,1,1
			If RS.Eof And RS.Bof Then
			 RS.Close : Set RS=Nothing  
			 KS.Die "<script>alert('参数出错，该商品已不存在!');history.back();</script>"
			End If
			if rs("endtf")=1 then
				   RS.Close :Set RS=Nothing
                   Call KS.Alert("对不起，该团购已结束!",Request.ServerVariables("HTTP_REFERER"))
				   Exit SUb
			End If
			if rs("locked")=1 then
				   RS.Close :Set RS=Nothing
                   Call KS.Alert("对不起，该团购已锁定!",Request.ServerVariables("HTTP_REFERER"))
				   Exit SUb
			End If
			
			If DateDiff("s",now,RS("addDate"))>0 Then
				   RS.Close :Set RS=Nothing
                   Call KS.Alert("对不起，该团购还未开始!",Request.ServerVariables("HTTP_REFERER"))
				   Exit SUb
			End If
			
			If DateDiff("s",now,RS("ActiveDate"))<0 Then
				   RS.Close :Set RS=Nothing
                   Call KS.Alert("对不起，该团购已结束!",Request.ServerVariables("HTTP_REFERER"))
				   Exit SUb
			End If
            If (RS("AllowBMFlag")<>"0" And LoginTF=false) or (RS("AllowBMFlag")="2" and KS.FoundInArr(RS("AllowArrGroupID"),KSUser.GroupID,",")=false) Then
				   RS.Close :Set RS=Nothing
                   Call KS.Alert("对不起，您没有参加该团购的权限!",Request.ServerVariables("HTTP_REFERER"))
				   Exit SUb
			End iF
			if KS.ChkClng(RS("CleanCart"))=1 Then
			 Conn.Execute("Delete From KS_ShoppingCart where flag=1 AND username='" & GetUserID&"'")
			End If
				
			rs.close
			rs.open "select top 1 * from KS_ShoppingCart where flag=1 and username='" & GetUserID & "' And proid=" & id,conn,1,3
			if rs.eof and rs.bof then
			   rs.addnew
			   rs("flag")=1
			   rs("proid")=id
			   rs("username")=GetUserID
			   rs("attr")=""
			   rs("adddate")=now
			   rs("amount")=amount
			   rs.update
			end if
			rs.close
			set rs=nothing
		End Sub 
		
		Sub Del()
		 Dim CartID:CartID=KS.ChkClng(Request("cartid"))
		 If CartID=0 Then Exit Sub
		 Conn.Execute("Delete From KS_ShoppingCart Where cartID=" & CartID)
		 If KS.ChkClng(KS.Setting(179))=1 Then
		 Response.Redirect KS.GetDomain & "groupbuy/cart.html"
		 Else
		 Response.Redirect KS.GetDomain & "shop/GroupBuycart.asp"
		 End If
		End Sub
		
		Sub LoadCartList()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select c.cartid,c.proid,c.amount,g.subject,g.price,g.weight,g.Price_Original,g.LimitBuyNum,g.AllowBMFlag,g.AllowArrGroupID,g.minnum,g.IsSuccess,g.CleanCart,g.MustPayOnline,g.showdelivery from KS_ShoppingCart c  inner join KS_GroupBuy g on c.proid=g.id where c.flag=1 and c.username='" & GetUserid & "' Order By c.cartid",conn,1,1
		  If Not RS.Eof Then
		     Product=RS.GetRows(-1)
		  End If
		  RS.Close:Set RS=Nothing
		  If Not IsArray(Product) Then
		     If KS.ChkClng(KS.Setting(179))=1 Then
		     KS.Die "<script>alert('对不起，购物车中没有商品!');location.href='" & KS.GetDomain & "groupbuy/';</script>"
			 Else
		     KS.Die "<script>alert('对不起，购物车中没有商品!');location.href='" & KS.GetDomain & "shop/groupbuy.asp';</script>"
			 End If
		  End If
		End Sub
		
	 '发货方式
	  Function GetDeliveryTypeStr()
	   if Product(12,0)="1" and Product(14,0)="0" then exit function
	   Dim j,rss,rsss
	   Dim DiscountStr,SQL,I,RS,defaultcity,defaulttips
		set rsss=conn.execute("select top 1 * from KS_Delivery where isDefault=1 order by orderid")
		if not rsss.eof then 
			 typeid=rsss("expressid"):defaultcity=rsss("defaultcity")
			 defaulttips="首重：<span>"& formatnumber(round(rsss("fweight"),2),2,-1) & "</span> kg 首重价格：<span>" & formatnumber(rsss("carriage"),2,-1) & "</span> 元 续重价格：<span>" & formatnumber(round(rsss("C_fee")/rsss("W_fee"),2),2,-1)&" </span>元/kg"
		end if
		rsss.close:set rsss=nothing
		
	   if KS.IsNul(defaultcity) Then 
	    defaulttips="选择送货路线确定运费！":defaultcity="选择送货地点" 
		DeliveryMoney=0
	   Else 
	    tocity=defaultcity
		DeliveryMoney=KS.GetFreight(TypeID,ToCity,totalweight,"")
		If DeliveryMoney=-1 Then DeliveryMoney=0
		
	   End if
	   If totalweight=-1 Then DeliveryMoney=0  '包邮
	   Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "select TypeID,TypeName,IsDefault from KS_DeliveryType order by orderid,TypeID",conn,1,1
	   If Not RS.Eof Then
		 SQL=RS.GetRows(-1)
	   End IF
	   RS.Close:Set RS=Nothing
	   GetDeliveryTypeStr="<strong>快递公司：</strong><select onchange=""ajshowdata($('#ccity').html())"" name='DeliverType' id='DeliverType'>"
	   For I=0 To UBound(SQL,2)
		 If trim(typeid)=trim(sql(0,i)) Then
		GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & "</option>"
		 Else
		GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & "</option>"
		End If
	   Next
	   GetDeliveryTypeStr=GetDeliveryTypeStr & "</select>"
	  
	   GetDeliveryTypeStr=GetDeliveryTypeStr & " <input type=""hidden"" value=""" & tocity & """ name=""tocity"" id=""tocity""/> <span style='position:relative'><input class=""tocity"" style='text-align;left' name='' id='choosecity' type='button' value='" & defaultcity & "'  onclick=""$('#showprovn').show();if(this.getBoundingClientRect().top>300){showprovn.style.top=(this.offsetHeight-showprovn.offsetHeight)}else{showprovn.style.top='0'};""><span style='display:none' id='showprovn' onclick=""this.style.display='none'"" class='showcity'>"&_
				 "<table width='92%' align='center' border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
						dim pxml,node,pnode
						set rss=conn.execute("select id,City,parentid from KS_Province order by orderid asc,id")
						if not rss.eof then
						  set pxml=KS.RsToXml(rss,"row","")
						end if
						rss.close  : Set RSS=Nothing
						If IsObject(Pxml) Then
						 For Each Node In pxml.DocumentElement.SelectNodes("row[@parentid=0]")
							GetDeliveryTypeStr=GetDeliveryTypeStr&"<tr><td colspan='5' class='provincename'><strong>" & Node.SelectSingleNode("@city").text &"</td></tr>"
							j=1
							For Each pnode in Pxml.DocumentElement.SelectNodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
							IF (j MOD 5) = 1 THEN GetDeliveryTypeStr=GetDeliveryTypeStr&"<tr>"&vbcrlf
							GetDeliveryTypeStr=GetDeliveryTypeStr&"<td id='ccity' onclick=""$('#choosecity').val(this.innerHTML);ajshowdata(this.innerHTML);getyf(this.innerHTML);"" style='cursor:hand' onmouseover=""this.style.color='red'"" onmouseout=""this.style.color=''"">"&pnode.selectsinglenode("@city").text&"</td>"&vbcrlf
							if (j mod 5)=0 then GetDeliveryTypeStr=GetDeliveryTypeStr&"</tr>"&vbcrlf
							j=j+1
							Next
							
						 Next
						End If
	 
				 GetDeliveryTypeStr=GetDeliveryTypeStr&"</table>"&vbcrlf&_
				"</span></span>"&_
			"<br/><strong>价格信息：</strong><span id='jgxx' class='jgxx'>" & defaulttips & "</span>"&vbcrlf
			
			GetDeliveryTypeStr=GetDeliveryTypeStr&"<script>$(""#jgxx"").html('" & DeliveryMoney &"元');</script>"
	  End Function
	  
	  '付款方式
	  Function GetPayTypeStr()
	     if Product(13,0)="0" and  Product(12,0)="1" then exit function
		 if ubound(product,2)=0 and Product(13,0)="0" then exit function
		  Dim SQL,K,Param,PayStr,RS
		  Set RS=Server.CreateOBject("ADODB.RECORDSET")
		  RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 Order By OrderID",conn,1,1
						   If Not RS.Eof Then SQL=RS.GetRows(-1)
						   RS.Close:Set RS=Nothing
						   If Not IsArray(SQL) Then
							PayStr=""
						   Else
							 For K=0 To Ubound(SQL,2)
							   PayStr=PayStr & "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
							   If trim(SQL(3,K))="1" Then PayStr=PayStr &  " checked"
							   PayStr=PayStr &  ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
							 Next
						   End If
		   GetPayTypeStr=PayStr
	  End Function
  
	  Function  OrderSave()
	    Dim RealName:RealName=KS.S("ContactMan")
		Dim Address:Address=KS.S("Address")
		Dim ZipCode:ZipCode=KS.S("ZipCode")
		Dim Mobile:Mobile=KS.S("Mobile")
		If KS.IsNul(RealName) Then
		  KS.AlertHintScript "对不起，收货人必须输入！"
		  KS.Die ""
		End If
		If KS.IsNul(Address) Then
		  KS.AlertHintScript "对不起，收货地址必须输入！"
		  KS.Die ""
		End If
		If KS.IsNul(ZipCode) Then
		  KS.AlertHintScript "对不起，邮政编码必须输入！"
		  KS.Die ""
		End If
		If KS.IsNul(Mobile) Then
		  KS.AlertHintScript "对不起，联系手机必须输入！"
		  KS.Die ""
		End If
		
		Dim DeliverType:DeliverType=KS.ChkClng(KS.S("DeliverType"))
		Dim ToCity:ToCity=KS.R(KS.S("ToCity"))
		If KS.IsNul(ToCity) and Product(12,0)="1" and Product(14,0)="1"  Then
		  KS.AlertHintScript "对不起，请选择送往城市！"
		  KS.Die ""
		End If
		Dim TotalWeight:TotalWeight=0
		TotalPrice=0
		For K=0 To Ubound(Product,2)
			TotalWeight=TotalWeight+Product(5,k)*Product(2,k)
			TotalPrice=TotalPrice+Product(4,k)*Product(2,k)
		Next
		Dim ExpressCompany,RSA
		Dim DeliveryMoney:DeliveryMoney=KS.GetFreight(DeliverType,ToCity,totalweight,ExpressCompany)
		 If DeliveryMoney=-1 Then		
		  KS.AlertHintScript "对不起，您选择的线路还没有开通此送货方式！"
		  KS.Die ""
		 End If
		 If totalweight=0 Then DeliveryMoney=0  '包邮
		 
		 OrderID=KS.Setting(71) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&KS.MakeRandom(8)
		 For K=0 To Ubound(Product,2)
		   Set RSA=Server.CreateObject("ADODB.RECORDSET")
		   RSA.Open "select top 1 * from ks_orderitem where 1=0",conn,1,3
		   RSA.AddNew
		       RSA("UserIP")=KS.GetIP
			 If KS.C("UserName")<>"" And KS.C("PassWord")<>"" Then
			   RSA("IsMember")=1
			 Else
			   RSA("IsMember")=0
			 End If
			  RSA("OrderID")=OrderID
			  RSA("ProID")=Product(1,k)
			  RSA("SaleType")=1
			  RSA("PackID")=0
			  RSA("Price_Original")=Product(6,k)
			  RSA("Price")=Product(4,k)
			  RSA("IsChangedBuy")=0
			  RSA("RealPrice")=Product(4,k)
			  RSA("Amount")=Product(2,k)
			  RSA("AttributeCart")=""
			  RSA("TotalPrice")=Round(Product(4,k)*Product(2,k),2)
			  RSA("BeginDate")=Now
			  RSA("ServiceTerm")=0
			  RSA("BundleSaleProID")=0
		   RSA.Update
		   RSA.Close:Set RSA=Nothing
		 Next
		 
		 
		RealMoneyTotal=TotalPrice+DeliveryMoney
		 Dim UserName,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order",Conn,1,3
			RS.AddNew
			     RS("OrderID")=OrderID
				 If Cbool(KSUser.UserLoginChecked)=true Then
				  UserName= KSUser.UserName
				 Else
				  UserName = "游客"
				 End If
				RS("UserName") = UserName
				RS("MoneyTotal")=RealMoneyTotal
				RS("MoneyGoods")=TotalPrice
				RS("NoUseCouponMoney")=RealMoneyTotal
				RS("NeedInvoice")=0
				RS("InvoiceContent")=""
				RS("Remark")=KS.S("Remark")
				RS("InputTime")=Now
				RS("ContactMan")=RealName
				RS("Address")=Address
				RS("ZipCode")=ZipCode
                RS("Mobile")=Mobile
				RS("Phone")=Mobile
				RS("QQ")=""
				RS("Email")=""
				RS("PaymentType")=0
				RS("DeliverType")=KS.ChkClng(DeliverType)
                RS("Discount_Payment")=100
				RS("Charge_Deliver")=DeliveryMoney     '运费
				RS("ToCity")=ToCity '送达城市
				RS("Weight")=totalweight
				RS("OrderType")=1
				RS("CouponUserID")=0              '优惠券使用人ID
				RS("UseCouponMoney")=0           '使用优惠券的抵扣金额
				RS("PayTime")="2000-1-1"   '表示未付款
				If MustPayOnline=true Then
				RS("PayStatus")=0
				RS("Status")=0         '订单状态
				Else
				RS("PayStatus")=100   '不需要在线付款的订单。
				RS("Status")=1         '订单状态
				End If
				

				'相关初始值
				RS("Invoiced")=0       '发票未开
				RS("MoneyReceipt")=0   '已收款
				RS("BeginDate")=Now    '开始服务日期
				RS("DeliverStatus")=0  '送货状态
				RS("PresentMoney")=0       '返回客户现金
				RS("PresentPoint")=0       '返回客户点券
				RS("PresentScore")=0       '返回客户积分
			  RS.Update
			  RS.MoveLast
			  OrderAutoID=RS("id")
			  RS.Close
		      Set RS=Nothing
			  IsSuccess=true
			  Conn.Execute("Delete From KS_ShoppingCart Where flag=1 and UserName='" & GetUserID & "'")
			  
			  For K=0 To Ubound(Product,2)
			   if KS.ChkClng(conn.execute("select sum(i.amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.proid=" & Product(1,k))(0))>=KS.ChkClng(Product(10,k)) and KS.ChkClng(Product(11,k))=0 then
				 conn.execute("update ks_groupbuy set IsSuccess=1,minnumtime=" & SQLNowString &" where id=" & Product(1,k))
			   end if
			  Next
			  
	  End Function      

End Class
%>

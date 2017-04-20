<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************


Dim KSCls
Set KSCls = New ShoppingCart
KSCls.Kesion()
Set KSCls = Nothing

Class ShoppingCart
        Private KS, KSRFObj,KSUser,DomainStr
		Private ProductList,TotalPrice,TotalWeight,Price_Member,CurrWeight,usezf,UseScoreisshop,isscore_str
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		  Set KSUser = New UserCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="PackCart.asp"-->
		<%
		Public Sub Kesion()
		  Dim FileContent,Products,i,RS,strsql,CartStr,OrderAutoID
		      ProductList = KS.FilterIDs(Session("ProductList"))
			  usezf=KS.ChkClng(ks.g("usezf"))
			  If KS.S("Action")<>"Confirm" Then  Response.Write "<script>alert('非法提交!');window.close();</script>":response.end
			  TotalWeight=0
			  
			 IF KS.ChkClng(KS.Setting(180))=1 and KS.S("ToCity")="" then ks.die "<script>alert('请选择送货地点!');history.back();</script>"			   
				 '生成订单号
				 Dim OrderID:OrderID=KS.Setting(71) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&KS.MakeRandom(8)
			     Dim ContactMan:ContactMan=KS.S("ContactMan")
				 Dim Address:Address=KS.S("Address")
				 Dim ZipCode:ZipCode=KS.S("ZipCode")
				 Dim Phone:Phone=KS.S("Phone")
				 Dim Email:Email=KS.S("Email")
				 Dim Mobile:Mobile=KS.S("Mobile")
				 Dim QQ:QQ=KS.S("QQ")
				 Dim PaymentType:PaymentType=KS.ChkClng(KS.S("PaymentType"))
				 Dim DeliverType:DeliverType=KS.ChkClng(KS.S("DeliverType"))
				  If DeliverType=0 then DeliverType=KS.ChkClng(KS.S("expressid"))
				 Dim InvoiceContent:InvoiceContent=KS.S("InvoiceContent")
				 Dim NeedInvoice:NeedInvoice=KS.ChkClng(KS.S("NeedInvoice"))
				 Dim Remark:Remark=KS.S("Remark")
                 Dim ProScore,RSA,RealPrice,MoneyGoods,TotalScore,RealMoneyTotal,CouponID,CouponUserID,CouponNum,FaceValue,MaxDiscount,availablemoney
				
				  AvailableMoney=0 : MaxDiscount=0
				  
				If ContactMan="" Then Response.Write "<script>alert('请填写收货信息!');history.back();</script>":response.end
				
				 CouponNum=trim(KS.S("CouponNum"))
				 Set RS=Server.CreateObject("ADODB.RecordSet") 
				 If CouponNum<>"" and Cbool(KSUser.UserLoginChecked)=true Then
				  RS.Open "Select top 1 * From KS_ShopCouponUser Where CouponNum='" & CouponNum & "'",conn,1,3
				  If RS.Eof And RS.Bof Then
				    RS.Close:Set RS=Nothing
					Response.Write "<script>alert('对不起,您输入的优惠券不可用!');history.back();</script>"
					Exit Sub
				  Else
				     If RS("UseFlag")=1 and Round(RS("AvailableMoney"),2)<=0 Then
					   RS.Close:Set RS=Nothing
					   Response.Write "<script>alert('对不起,该优惠券的可抵扣金额已用完!');history.back();</script>"
					   Exit Sub
					 ElseIf Cbool(KSUser.UserLoginChecked)=false Then
					   RS.Close:Set RS=Nothing
					   Response.Write "<script>alert('对不起,必须登录后才可以使用优惠券!');history.back();</script>"
					   Exit Sub
					 Else  
					   'RS("UserName")=KSUser.UserName
					   'RS.Update
					   CouponID=KS.ChkClng(RS("CouponID"))
					 End If
				  End If
				   RS.Close 
				 End If
				
				 
				 If CouponID="" Or CouponID=0 Then
				   CouponID=KS.ChkClng(KS.S("Couponid"))
				 End If
				 
				
				 
				 If CouponID<>0 and Cbool(KSUser.UserLoginChecked)=true Then
				  If KS.ChkClng(KS.S("Couponid"))=0 And CouponNum<>"" Then
				  RS.Open "Select A.*,b.id as CouponUserID,b.AvailableMoney From KS_ShopCoupon A Inner Join KS_ShopCouponUser B ON A.ID=B.CouponID Where B.CouponNum='" & CouponNum & "'",conn,1,1
				  Else
				  RS.Open "Select A.*,b.id as CouponUserID,b.AvailableMoney From KS_ShopCoupon A Inner Join KS_ShopCouponUser B ON A.ID=B.CouponID Where B.ID=" &KS.ChkClng(KS.S("Couponid")) ,conn,1,1
				  End If
				  If Not RS.Eof Then
				     If DateDiff("s",RS("BeginDate"),Now)<0 Then
					  RS.Close:Set RS=Nothing
		              Response.Write "对不起,您输入的优惠券需要" & RS("BeginDate") & "后才能使用!"
					  Exit Sub
				     ElseIf DateDiff("s",RS("EndDate"),Now)>0 Then
					  RS.Close:Set RS=Nothing
					  Response.Write "<script>alert('对不起,您输入的优惠券已过期!');history.back();</script>"
					  Exit Sub
					 ElseIf RS("Status")=0 Then
					  RS.Close:Set RS=Nothing
					  Response.Write "<script>alert('对不起,您输入的优惠券已锁定!');history.back();</script>"
					  Exit Sub
					 ElseIf round(KS.S("TRealTotalPrice"),2)<round(RS("MinAmount"),2) Then
					  Response.Write "<script>alert('对不起,该优惠券要求订单金额必须大于等于 ￥" & RS("MinAmount") & " 元才可以抵用!');history.back();</script>"
					  RS.Close:Set RS=Nothing
					  Exit Sub
					 ElseIf Round(RS("AvailableMoney"),2)<=0 Then
					  RS.Close:Set RS=Nothing
					  Response.Write "<script>alert('对不起,该优惠券已全部抵用完了,不能再使用!');history.back();</script>"
					  Exit Sub
					 Else
					   MaxDiscount=RS("maxdiscount")
					   AvailableMoney=RS("AvailableMoney")
					   FaceValue=RS("FaceValue") 
					   CouponUserID=RS("CouponUserID")
					   CouponID=RS("ID")
					 End If
				  End If
				  RS.Close
				 End If
				 
				Dim yhlx:yhlx=KS.ChkClng(Request("yhlx"))
				Dim MyScore:MyScore=KS.ChkClng(Request("myscore"))
				KSUser.UserLoginChecked
				Dim UseScoreMoney:UseScoreMoney=0
				Dim NowMyScore:NowMyScore=KS.ChkClng(KSUser.GetScore())
				Dim ScoreRate:ScoreRate=KS.Setting(182)
				If Not IsNumeric(ScoreRate) Then ScoreRate=0
				If yhlx=1 and MyScore>0 and KS.ChkClng(request("usezf"))=1 and KS.ChkClng(ScoreRate)>0 Then
					 Dim LimitTotalMoney:LimitTotalMoney=KS.Setting(183)
					 Dim LimitPer:LimitPer=KS.Setting(184)
					 If Not IsNumeric(LimitTotalMoney) Then LimitTotalMoney=0
					 If Not IsNumeric(LimitPer) Then LimitPer=0

				     If MyScore>NowMyScore Then
					  Response.Write "<script>alert('对不起,您的可用积分只有" & NowMyScore & "分!');history.back();</script>"
					  Exit Sub
					 ElseIf round(KS.S("TRealTotalPrice"),2)<round(LimitTotalMoney,2) and round(LimitTotalMoney,2)>0 Then
					  Response.Write "<script>alert('对不起,系统限定只有订单金额达到" & LimitTotalMoney & "元时才可以使用积分抵用!');history.back();</script>"
					  Exit Sub
					 End If
					 UseScoreMoney=MyScore/ScoreRate
					 If Round(LimitPer,2)>0 Then
					    dim allowscoremoney:allowscoremoney=round(KS.S("TRealTotalPrice"),2)*Round(LimitPer,2)/100
					   If Round(UseScoreMoney,2)> round(allowscoremoney,2) Then
					    dim allowscore:allowscore=allowscoremoney * ScoreRate
					    Response.Write "<script>alert('对不起,系统限定积分抵扣金额不能超过订单总金额的" & LimitPer & "%,您最多可以用" & allowscore & "积分抵扣" & allowscoremoney & "元!');history.back();</script>"
					    Exit Sub
					   End If
					 End If
				End If
				

				 
				 
			' If Not KS.IsNul(ProductList) Then 
				'   RS.Open "select I.* ,L.LimitBuyBeginTime,L.LimitBuyEndTime,L.ID as TaskID,L.TaskType from KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id where I.ID in ("&ProductList&") order by I.ID",Conn,1,1
					
					 RS.Open "select I.* ,L.LimitBuyBeginTime,L.LimitBuyEndTime,L.ID as TaskID,L.TaskType,C.Attr,C.Amount,C.AttrID from (KS_Product I Inner join KS_ShoppingCart c on i.id=c.proid) Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id where c.flag=0 and c.ispay=1 and c.username='" & GetUserID & "' order by I.ID",Conn,1,1
				   If RS.Eof And RS.Bof Then
				     if conn.execute("Select top 1 a.title,a.Price_Member,a.Price,b.*,0 as attrid From KS_Product A inner join KS_ShopPackageSelect b on a.id=b.proid Where b.UserName='" & GetUserID & "'").eof then
						 RS.CLose:Set RS=Nothing
						 KS.Die "<script>alert('对不起，购物车中没有商品!');history.back();</script>"
					 end if
				   End If
				  
					 Do While Not RS.eof
					 
					 If RS("AttrID")<>0 Then 
						  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RS("AttrID"))
						  If Not RSAttr.Eof Then
							Price_Member=RSAttr("Price")
							CurrWeight=RSAttr("Weight")
						  Else
							Price_Member=RS("Price_Member")
							CurrWeight=RS("Weight")
						  End If
						  RSAttr.CLose:Set RSAttr=Nothing
					 Else
						Price_Member=RS("Price_Member")
						CurrWeight=RS("Weight")
					 End If
					  Conn.Execute("Update KS_Product Set SaleNum=SaleNum+" & KS.ChkClng(RS("Amount")) & " Where id=" & RS("ID")) '记录销售量
					  
					   Set RSA=Server.CreateObject("ADODB.RecordSet")
					   RSA.Open "select top 1 * from KS_OrderItem where ID is null",Conn,1,3
					   RSA.AddNew
					     RSA("AttrID")=RS("AttrID")
					     RSA("UserIP")=KS.GetIP
						 If KS.C("UserName")<>"" And KS.C("PassWord")<>"" Then RSA("IsMember")=1 Else RSA("IsMember")=0
						 RSA("OrderID")=OrderID
						 RSA("ProID")=RS("ID")
						 RSA("SaleType")=RS("isdiscount")
						 RSA("Price_Original")=RS("Price")
						 RSA("Price")=Price_Member
						 If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then RSA("IsChangedBuy")=1 Else RSA("IsChangedBuy")=0
					     If RS("IsLimitBuy")="1" And Now>RS("LimitBuyBeginTime") And RS("LimitBuyEndTime")>Now And RS("LimitBuyAmount")>0 Then
						 RSA("LimitBuyTaskID")=rs("taskid")
						 RSA("IsLimitBuy")=1
						 ElseIf RS("IsLimitBuy")="2" And RS("LimitBuyAmount")>0 Then
						 RSA("LimitBuyTaskID")=rs("taskid")
						 RSA("IsLimitBuy")=2
						 Else
						 RSA("LimitBuyTaskID")=0
						 RSA("IsLimitBuy")=0
						 End If
						 
						 
						If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
						   RealPrice=Session("ChangeBuyPrice")	
						   ProScore=0									
					   ElseIf RS("IsLimitBuy")="1" And Now>RS("LimitBuyBeginTime") And RS("LimitBuyEndTime")>Now And RS("LimitBuyAmount")>0 Then
						   RealPrice=RS("LimitBuyPrice")
						   ProScore=0
					   ElseIf RS("IsLimitBuy")="2" And RS("LimitBuyAmount")>0 Then
						   RealPrice=RS("LimitBuyPrice")
                            ProScore=0
					   ElseIF Cbool(KSUser.UserLoginChecked)=true Then
						  Dim Discount:Discount=KS.U_S(KSUser.GroupID,17)
						  Dim JFDiscount:JFDiscount=KS.U_S(KSUser.GroupID,18)
						   If Not IsNumeric(Discount) Then Discount=0
						   If Not IsNumeric(JFDiscount) Then JFDiscount=0
						  If KS.U_S(KSUser.GroupID,21)="1" and rs("vipprice")<>"0" then
							RealPrice=RS("VipPrice")
						  ElseIf KS.ChkClng(RS("isdiscount"))=0 or Discount=0 Then
						    RealPrice=Price_Member
						  Else
							RealPrice=FormatNumber(Price_Member*discount/10,2,-1)
						  End If
						  If JFDiscount=0 Then
							ProScore=0
						  ElseIf JFDiscount=1 Then
							ProScore=KS.ChkClng(RealPrice)*RS("Amount")
						  Else
							ProScore=RealPrice*JFDiscount*RS("Amount")'bug改动
						  End If
						Else
						  RealPrice=Price_Member
						End If
						 
						'积分购买
						if usezf=1 then 
						     UseScoreisshop=KS.ChkClng(KS.S("myscore"))
							 UseScoreisshop= UseScoreisshop+(RS("Score")*RS("Amount"))
							 if KSUser.GetScore() >= UseScoreisshop then
								isscore_str=""& UseScoreisshop &"积分+"
							 else
							 	KS.Die "<script>alert('对不起，您输入的支付积分超过了您的积分余额!');history.back();</script>"	
								Response.end()
							 end if	
						end if	
						if ks.chkclng(rs("WholesaleNum"))<>0 and KS.ChkClng(RS("Amount"))>=ks.chkclng(rs("WholesaleNum")) then realPrice=rs("WholesalePrice")
						 RSA("Score")=KS.ChkClng(ProScore)
						 RSA("RealPrice")=RealPrice
						 RSA("Amount")=KS.ChkClng(RS("Amount"))
						 RSA("AttributeCart")=RS("Attr")
						 RSA("TotalPrice")=Round(RealPrice*RS("Amount"),2)
						 RSA("BeginDate")=Now
						 RSA("ServiceTerm")=RS("ServiceTerm")
						 RSA("PackID")=0
						 RSA("BundleSaleProID")=0
					   RSA.Update
					   RSA.Close:Set RSA=Nothing
					   MoneyGoods=MoneyGoods+Round(RealPrice*RS("Amount"),2)
					   TotalScore=TotalScore+ProScore
					

					   IF RS("Amount")>=KS.ChkClng(rs("FreeShipping")) And KS.ChkClng(rs("FreeShipping")) <>0 Then 
						 TotalWeight=-1
					   End If
					   if TotalWeight<>-1 then
					   TotalWeight=TotalWeight+CurrWeight*RS("Amount")
					   end if

                       If RS("TaskType")=2 Or (RS("TaskType")=1 and Now>RS("LimitBuyBeginTime") And RS("LimitBuyEndTime")>Now) Then
						 '扣除供抢购数
						 Conn.Execute("Update KS_Product Set LimitBuyAmount=LimitBuyAmount-" & RS("Amount")& " Where id=" & RS("ID"))
						 Conn.Execute("Update KS_Product Set LimitBuyAmount=0 Where id=" & RS("ID") & " and LimitBuyAmount<0")
					   End If
						 
						 '将捆绑促销的抢购商品加入KS_OrderItem表
						 Dim RSK:Set RSK=Conn.Execute("Select I.ID,I.Title,I.ServiceTerm,I.Price,b.Price as realprice,b.amount,b.AttributeCart,i.Weight,b.id as kbid From KS_Product I inner Join KS_ShopBundleSelect b on i.id=b.pid Where B.ProID=" & RS("ID") & " and b.username='" & GetUserID & "' order by I.id")
						 Do While Not RSK.Eof
						       Conn.Execute("Update KS_Product Set SaleNum=SaleNum+" & KS.ChkClng(RSK("Amount")) & " Where id=" & RSK("ID")) '记录销售量
						   	   Set RSA=Server.CreateObject("ADODB.RecordSet")
							   RSA.Open "select top 1 * from KS_OrderItem where ID is null",Conn,1,3
							   RSA.AddNew
							     RSA("AttrID")=0
							     RSA("UserIP")=KS.GetIP
								 If KS.C("UserName")<>"" And KS.C("PassWord")<>"" Then RSA("IsMember")=1 Else RSA("IsMember")=0
								 RSA("OrderID")=OrderID
								 RSA("ProID")=RSK("ID")
								 RSA("SaleType")=6       '捆绑销售的商品
								 RSA("Price_Original")=RSK("Price")
								 RSA("Price")=RSK("realprice")
								 RSA("IsChangedBuy")=0
								 RSA("LimitBuyTaskID")=0
								 RSA("IsLimitBuy")=0
								 RSA("RealPrice")=RSK("RealPrice")
								 RSA("Amount")=RSK("Amount")
								 RSA("AttributeCart")=RSK("AttributeCart")
								 RSA("TotalPrice")=Round(RSK("RealPrice")*RSK("Amount"),2)
								 RSA("BeginDate")=Now
								 RSA("ServiceTerm")=RSK("ServiceTerm")
								 RSA("PackID")=0
								 RSA("BundleSaleProID")=RS("ID")
							   RSA.Update

								 moneyGoods=MoneyGoods + Round(RSK("RealPrice")*RSK("Amount"),2)
                                 if TotalWeight<>-1 then
                                 TotalWeight=TotalWeight+RSK("Weight")*RSK("Amount")
								 end if
						       RSA.Close:Set RSA=Nothing
							   	'删除捆绑促销订购表数
				                Conn.Execute("Delete From KS_ShopBundleSelect Where id=" & rsk("kbid"))

						 RSK.MoveNext
						 Loop
						 RSK.Close:Set RSK=Nothing
						Session("Amount"&RS("ID"))=""

					  RS.MoveNext
					 Loop
				  RS.Close:Set RS=Nothing
			' End If 								
			  
			 '====================================将礼包内的商品加入到orderitem表===========================================
			 Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select a.title,a.weight,a.Price,a.Price_Member,b.* From KS_Product A inner join KS_ShopPackageSelect b on a.id=b.proid Where b.UserName='" & GetUserID & "'",Conn,1,1
				 Do While Not RS.Eof
				      Conn.Execute("Update KS_Product Set SaleNum=SaleNum+" & RS("Amount") & " Where id=" & RS("ProID")) '记录销售量
				 	  Set RSA=Server.CreateObject("ADODB.RecordSet")
					   RSA.Open "select top 1 * from KS_OrderItem where ID is null",Conn,1,3
					   RSA.AddNew
					     RSA("AttrID")=RS("AttrID")
						 If RS("AttrID")<>0 Then 
							 Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RS("AttrID"))
							  If Not RSAttr.Eof Then
								Price_Member=RSAttr("Price")
								CurrWeight=RSAttr("Weight")
							  Else
								Price_Member=RS("Price_Member")
								CurrWeight=RS("Weight")
							  End If
							  RSAttr.CLose:Set RSAttr=Nothing
						 Else
							Price_Member=RS("Price_Member")
							CurrWeight=RS("Weight")
						 End If
						 
						 if TotalWeight<>-1 then
						 TotalWeight=TotalWeight+CurrWeight*RS("Amount")
						 end if
					     RSA("UserIP")=KS.GetIP
						 If KS.C("UserName")<>"" And KS.C("PassWord")<>"" Then RSA("IsMember")=1 Else RSA("IsMember")=0
						 RSA("OrderID")=OrderID
						 RSA("ProID")=RS("ProID")
						 RSA("SaleType")=5
						 RSA("PackID")=rs("packid")
						 RSA("Price_Original")=RS("Price")
						 RSA("Price")=Price_Member
						 RSA("IsChangedBuy")=0
						 
					
						 RealPrice=Price_Member
						 RealPrice=RealPrice * Conn.Execute("Select top 1 discount From KS_ShopPackage Where ID=" & RS("PackID"))(0)/10
						 RSA("RealPrice")=RealPrice
						 RSA("Amount")=RS("Amount")
						 RSA("AttributeCart")=RS("AttributeCart")
						 RSA("TotalPrice")=Round(RealPrice*RS("Amount"),2)
						 RSA("BeginDate")=Now
						 RSA("ServiceTerm")=0
						 RSA("BundleSaleProID")=0
					   RSA.Update
					   RSA.Close:Set RSA=Nothing
					   MoneyGoods=MoneyGoods+Round(RealPrice*RS("Amount"),2)

				  RS.MoveNext
				 Loop
				 RS.Close
				 
				 '删除礼包订购表数据
				 Conn.Execute("Delete From KS_ShopPackageSelect Where UserName='" & GetUserID &"'")
				 '删除捆绑促销订购表数
				 Conn.Execute("Delete From KS_ShopBundleSelect Where UserName='" & GetUserID &"'")
				
	             '实际支付金额。
				 Dim ToCity:ToCity=KS.S("ToCity")
				 Dim PaymentDiscount:PayMentDiscount=KS.ReturnPayment(PaymentType,1)
				 Dim DeliveryMoney
				 
				 dim freeDelivery:freeDelivery=0     '满足一定金额免邮
				 if isnumeric(ks.setting(207)) then freeDelivery=round(ks.setting(207),2) else freeDelivery=0
				 if MoneyGoods>=freeDelivery and freeDelivery<>0 then
					DeliveryMoney=0
				 else
				    DeliveryMoney=KS.GetFreight(DeliverType,ToCity,TotalWeight,"") 
				 End If
				 
				 If Not IsNumeric(DeliveryMoney) Then DeliveryMoney=0
				 Dim TaxRate:TaxRate=KS.Setting(65)
				 Dim IncludeTax:IncludeTax=KS.Setting(64)
				 Dim TaxMoney,UserName,NoUseCouponMoney
				 If IncludeTax=1 Or NeedInvoice=0 Then TaxMoney=1 Else TaxMoney=1+Taxrate/100
				 '总金额 = (总价*付费方式折扣+运费)*(1+税率)
				 RealMoneyTotal=Round((MoneyGoods*PayMentDiscount/100+DeliveryMoney)*TaxMoney,2)
				 
				 
				 NoUseCouponMoney=RealMoneyTotal
				 Dim JKCouponMoney:JKCouponMoney=0
				 
				If AvailableMoney>0 Then
				    JKCouponMoney=AvailableMoney
				    If MaxDiscount>0 Then
					 dim allowmoney:allowmoney=round(RealMoneyTotal,2)* (maxdiscount/100) '按百分比得可抵扣金额
					  if (allowmoney>availablemoney) then
						 allowmoney=availablemoney
					  end if
					  JKCouponMoney=allowmoney
					End If
					if round(RealMoneyTotal,2)<round(JKCouponMoney,2) then 
					 JKCouponMoney= round(RealMoneyTotal,2)
					 RealMoneyTotal=0
					Else
				     RealMoneyTotal=Round(RealMoneyTotal-JKCouponMoney,2)
					End if
				ElseIf UseScoreMoney>0 Then  '使用积分抵扣
				    RealMoneyTotal=Round(RealMoneyTotal-UseScoreMoney,2)
				End If

			   If Cbool(KSUser.UserLoginChecked)=true Then
				  UserName= KSUser.UserName
				 Else
				  UserName = "游客"
				 End If
			   
		       If KS.Setting(210)="1" and KS.Setting(211)<>"" Then
			  		dim ReturnInfo,MailBodyStr
					MailBodyStr=MailBodyStr&"订单号:"&OrderID&"<br/>"
					MailBodyStr=MailBodyStr&"用户:"&UserName&"<br/>"
					MailBodyStr=MailBodyStr&"总金额:"&RealMoneyTotal&"<br/>"
					MailBodyStr=MailBodyStr&"收货人姓名:"&ContactMan&"<br/>"
					MailBodyStr=MailBodyStr&"收货人手机:"&Mobile&"<br/>"
					MailBodyStr=MailBodyStr&"收货人电话:"&Phone&"<br/>"
					MailBodyStr=MailBodyStr&"收货人QQ:"&QQ&"<br/>"
					MailBodyStr=MailBodyStr&"收货人Email:"&Email&"<br/>"
					'ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-订单",KS.Setting(211),UserName, MailBodyStr,KS.Setting(11))
			   End If
			   RS.Open "Select top 1 * From KS_Order",Conn,1,3
			   RS.AddNew
			    RS("OrderID")=OrderID
				RS("UserName") = UserName
				RS("MoneyTotal")=RealMoneyTotal
				RS("MoneyGoods")=MoneyGoods
				RS("NoUseCouponMoney")=NoUseCouponMoney
				RS("NeedInvoice")=NeedInvoice
				RS("InvoiceContent")=InvoiceContent
				RS("InvoiceCode")=KS.S("InvoiceCode")
				RS("InvoiceAddress")=KS.S("InvoiceAddress")
				RS("InvoiceTel")=KS.S("InvoiceTel")
				RS("InvoiceBank")=KS.S("InvoiceBank")
				RS("InvoiceBankCard")=KS.S("InvoiceBankCard")
				RS("Remark")=Remark
				RS("InputTime")=Now
				If Cbool(KSUser.UserLoginChecked)=true Then
					dim RSkj:Set RSkj=Server.CreateObject("ADODB.RECORDSET")
					RSkj.Open "Select top 1 * From KS_ShopUserOrder where username='"& UserName &"' and ContactMan='"& ContactMan &"'",conn,1,3
					If RSkj.EOF And RSkj.BOF Then 
						RSkj.AddNew
						RSkj("AddDate")=now()
					end if
					RSkj("username")=username
					RSkj("ContactMan")=ContactMan
					RSkj("Address")=Address
					RSkj("ZipCode")=ZipCode
					RSkj("Mobile")=Mobile
					RSkj("Phone")=Phone
					RSkj("QQ")=QQ
					RSkj("Email")=Email									
					RSkj.Update
					RSkj.close:Set RSkj = Nothing	
				end if
				RS("ContactMan")=ContactMan
				RS("Address")=Address
				RS("ZipCode")=ZipCode
                RS("Mobile")=Mobile
				RS("Phone")=Phone
				RS("QQ")=QQ
				RS("Email")=Email
				RS("PaymentType")=KS.ChkClng(PaymentType)
				RS("DeliverType")=KS.ChkClng(DeliverType)
                RS("Discount_Payment")=PaymentDiscount   '付款方式折扣率
				RS("Charge_Deliver")=DeliveryMoney     '运费
				RS("ToCity")=ToCity '送达城市
				If IsNumeric(TotalWeight) Then
				 RS("Weight")=TotalWeight
				Else
				 RS("Weight")=0
				End If
				RS("OrderType")=0
			
			    if JKCouponMoney>0 or UseScoreMoney>0 then  totalscore=0      '使用优惠券时，不送积分
				RS("UseScoreMoney")=UseScoreMoney
				if UseScoreMoney>0 then
				RS("UseScore")=MyScore
				else
				RS("UseScore")=0
				end if
				RS("TotalScore")=KS.ChkClng(TotalScore)
				RS("scoretf")=0
				RS("DeliveryDate")="2000-1-1"   '表示未发货
				
				RS("CouponUserID")=KS.ChkClng(CouponUserID)               '优惠券使用人ID
				RS("UseCouponMoney")=JKCouponMoney          '使用优惠券的抵扣金额
				RS("PayTime")="2000-1-1"   '表示未付款
				
				if (Round(RealMoneyTotal)<=0) Then
				  RS("PayStatus")=1 '使用优惠券，则设置为已付款
				  RS("Status")=1
				End If

				'相关初始值
				RS("Invoiced")=0       '发票未开
				if (JKCouponMoney>0) then
				 RS("MoneyReceipt")=JKCouponMoney   '已收款
				 RS("Status")=1         '订单状态
				else
				 RS("MoneyReceipt")=0   '已收款
				 RS("Status")=0         '订单状态
				end if
				RS("BeginDate")=Now    '开始服务日期
				RS("DeliverStatus")=0  '送货状态
				RS("PresentMoney")=0       '返回客户现金
				RS("PresentPoint")=0       '返回客户点券
				RS("PresentScore")=0       '返回客户积分
				RS("UseScoreisshop")=UseScoreisshop '应付积分
			  RS.Update
			  RS.MoveLast
			  OrderAutoID=RS("id")
			  RS.Close
			  
			  If KS.ChkClng(CouponUserID)<>0 and JKCouponMoney>0 Then
			   RS.Open "Select top 1 * From KS_ShopCouponUser WHERE ID=" & KS.ChkClng(CouponUserID),conn,1,3
			   If Not RS.Eof Then
			       RS("UseFlag")=1
				   RS("UserName")=UserName
				   RS("UseTime")=Now
				   RS("OrderID")=OrderID
				   If KS.IsNul(RS("Note")) Then
				     RS("Note")="于[" & Now & "]在订单[" & OrderID & "]中抵扣了" & KS.GetPrice(Round(JKCouponMoney,2)) & "元;"
				   Else
				     RS("Note")=rs("Note") & "<br/>于[" & Now & "]在订单[" & OrderID & "]中抵扣了" & KS.GetPrice(Round(JKCouponMoney,2)) & "元;"
				   End If
				   RS("AvailableMoney")=RS("AvailableMoney")-JKCouponMoney
				   If Round(RS("AvailableMoney"),2)<0 Then RS("AvailableMoney")=0
				   RS.Update 
			   End If
			   RS.Close
			   'Conn.Execute("Update KS_ShopCouponUser Set AvailableMoney=AvailableMoney-" & facevalue &",UseFlag=1,UserName='" & UserName & "',UseTime=" & SQLNowString & ",OrderID='" & OrderID & "' Where id=" & KS.ChkClng(CouponUserID)) 
			  End If
			  Set RS=Nothing
			  
			  '更新用户积分
			 if usezf=1 Then
			  If UseScoreMoney>0 And cbool(KSUser.UserLoginChecked)=true Then
			   Session("ScoreHasUse")="+"   '设置只累计消费积分
			   Call KS.ScoreInOrOut(KSUser.UserName,2,MyScore,"系统","扣减商城购物金额<font color=red>"& UseScoreMoney &"</font>元,订单号：" & OrderID & "!",0,0)
			  End If
			 End If
			  
			  Session("ProductList")=""  '交易成功！置购物车参数为空
			  Conn.Execute("Delete From KS_ShoppingCart Where flag=0 and ispay=1 and UserName='" & GetUserID & "'")
			  
		           If KS.Setting(124)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSRFObj.LoadTemplate(KS.Setting(124))
				   FCls.RefreshType = "ShoppingSuccess" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0"        '设置当前刷新目录ID 为"0" 以取得通用标签
				 If Trim(FileContent) = "" Then FileContent = "商城订单提交成功页模板不存在!"
				 
				 '得到支付平台
				 If Instr(FileContent,"{$PayMentList}")<>0 Then
					   Dim SQL,K,Param,PayStr
					   Set RS=Server.CreateOBject("ADODB.RECORDSET")
					   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsShow=1 and IsDisabled=1 Order By OrderID",conn,1,1
					   If Not RS.Eof Then SQL=RS.GetRows(-1)
					   RS.Close:Set RS=Nothing
					   If Not IsArray(SQL) Then
						PayStr=""
					   Else
					     PayStr="<form name=""myform"" id=""myform"" method=""get"" action=""payonline.asp"" target=""_blank"">"
						 PayStr=PayStr & "<input type=""hidden"" name=""id"" value=""" & OrderAutoID & """/>"
						 For K=0 To Ubound(SQL,2)
						   PayStr=PayStr & "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
						   If trim(SQL(3,K))="1" Then PayStr=PayStr &  " checked"
						   PayStr=PayStr &  ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
						 Next
					   End If
					   Dim PArr:Parr=Split(KS.Setting(82)&"||||||||","|")
					  
					  If Parr(0)="1" Then
					   PayStr=PayStr & "<input type=""submit"" style=""padding:2px"" value="" 进入支付平台在线支付 "" />"
					  ElseIf Parr(0)="2" Then
					   PayStr=PayStr & "<input type=""submit"" style=""padding:2px"" value="" 在线支付" & Parr(1) & "%的定金 "" />"
					  Else 
					   PayStr=PayStr & "支付金额:<input type='text' size='8' name='money' value='" & RealMoneyTotal & "'/> 元<br/><input type=""submit"" class=""queren"" value="" 确认在线支付 "" />"
					  End If 
					   PayStr=PayStr & "</form>"
					  FileContent=Replace(FileContent,"{$PayMentList}",PayStr) 
				 End If
                 FileContent=Replace(FileContent,"{$ShowOrderID}",OrderID)
			     FileContent=Replace(FileContent,"{$ShowOrderAutoID}",OrderAutoID)
			     FileContent=Replace(FileContent,"{$ShowTotalMoney}",KS.GetPrice(RealMoneyTotal)) 
				 FileContent=Replace(FileContent,"{$ShowTotalscore}",isscore_str)  
			     FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
                 KS.Echo FileContent
       End Sub
	   
End Class
%>

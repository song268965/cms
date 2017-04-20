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
Set KSCls = New PayMentCls
KSCls.Kesion()
Set KSCls = Nothing

Class PayMentCls
        Private KS, KSRFObj,KSUser,DomainStr
		Private TotalPrice,TotalScore,RealPrice,Price_Original,Discount,Amount,TotalWeight
		Private ProductList,DeliverType,DefaultCity,DeliveryStr,DeliveryMoney,istype,isscore_str,isshop_str,isshop_n,IsScore_z,IsScore_no
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
		  Dim FileContent,Products,i,RS,strsql,CartStr	
		  TotalWeight=0
		  Dim FreeShipping,NoFreeShipping
		  Dim cartid:cartid=KS.FilterIDs(Request("ID"))
		  If cartid="" Then KS.Die "<script>alert('对不起，您没有选择要去结算的商品！');location.href='shoppingcart.asp';</script>"
		  istype=KS.ChkClng(ks.g("istype")):IsScore_z=0:IsScore_no=0

			
		If KS.ChkClng(KS.Setting(63))=0 And Cbool(KSUser.UserLoginChecked)=false Then Response.Write "<script>alert('本商城设置注册用户才可购买，请先登录!');location.href='" &  DomainStr  & "user/login/';</script>"

		           If KS.Setting(122)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSRFObj.LoadTemplate(KS.Setting(122))
				   FCls.RefreshType = "ShoppingPayMent" '设置刷新类型，以便取得当前位置导航等
				   Fcls.RefreshFolderID = "0"        '设置当前刷新目录ID 为"0" 以取得通用标签
				   If Trim(FileContent) = "" Then FileContent = "商城购物车模板不存在!"
				   If KS.ChkClng(KS.Setting(180))<>1 Then FileContent=Replace(FileContent,"{$ShowDelivery}" ," style='display:none'")
	
	Dim packstr:packstr=GetPackage(false)          '礼包
	
	DeliveryStr = GetDeliveryTypeStr   '送货方式	
				   
    Set RS=Server.CreateObject("ADODB.RecordSet") 
	'If ProductList<>"" Then
	Dim ProBuyAttr
	'删除没有在当前购物车内的捆绑商品
	 conn.execute("delete from KS_ShopBundleSelect where username='" & getuserid & "' and proid not in(select proid from KS_ShoppingCart where flag=0 and username='" & getuserid & "')")
	 
	 '设置当前要去付款的商品
	 conn.execute("update KS_ShoppingCart set ispay=1 where cartid in (" & cartid &")")
	 
	 

	'strsql="select I.visitornum,I.MemberNum,I.ID,I.Title,I.ProductType,I.Price_Original,I.Price,I.Price_Member,I.Discount,I.TotalNum,I.GroupPrice,I.photourl,I.unit,I.IsLimitBuy,I.LimitBuyPrice,I.LimitBuyAmount,L.LimitBuyBeginTime,L.LimitBuyEndTime,i.tid,i.fname from KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  where I.ID in ("&ProductList&") order by I.IsChangedBuy,I.ID"
		strsql="select I.visitornum,I.MemberNum,I.ID,I.FreeShipping,I.Title,I.WholesaleNum,I.WholesalePrice,i.weight,I.Price,I.Price_Member,I.VipPrice,I.IsDiscount,I.TotalNum,I.PhotoUrl,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,I.LimitBuyAmount,i.tid,i.fname,i.adddate,i.Istype,i.Score,L.LimitBuyBeginTime,L.LimitBuyEndTime,C.Attr,C.Amount,C.CartID,C.AttrID from (KS_Product I Inner join KS_ShoppingCart c on i.id=c.proid) Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id   where c.flag=0 and c.cartid in (" & cartid &") and c.username='" & GetUserID & "' order by I.IsChangedBuy,I.ID"

	Call SetBundleSaleAmount()  '设置捆绑促销购买数量
	'ElseIf packstr="" Then
	'Response.Write "<script>alert('您的购物车中没有商品!');history.back();<//script>":response.end
	'End If

	CartStr="      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" width=""100%"" class=""border"">" & vbcrlf
	CartStr=CartStr&"          <tr class='title' height=""25"">" & vbcrlf 
	CartStr=CartStr&"            <td width=""6%"" align=""center"">编号</td>" &vbcrlf
	CartStr=CartStr&"            <td align=""center"">商品名称</td>" & vbcrlf
	CartStr=CartStr&"            <td width=""6%"" align=""center"">数量</td>" & vbcrlf
	CartStr=CartStr&"			 <td walign=""center"">商城价</td>" & vbcrlf
	CartStr=CartStr&"			 <td width=""10%"" align=""center"">折扣</td>"&vbcrlf
	CartStr=CartStr&"			 <td align=""center"">您的价格</td>" & vbcrlf
	CartStr=CartStr&"			 <td align=""center"">总计</td>" & vbcrlf
	CartStr=CartStr&"			 <td width=""8%"" align=""center"">赠送积分</td>" & vbcrlf
	CartStr=CartStr&"          </tr>"&vbcrlf
	
	Dim TotalNum:TotalNum=0	
    Dim Price_Member:Price_Member=0
    Dim CurrWeight:CurrWeight=0
	dim istype_str
	If strSql<>"" Then
			RS.open strsql,conn,1,1
			If RS.Eof And RS.Bof Then
			  If packstr="" Then  RS.CLose:Set RS=Nothing : KS.AlertHIntScript "对不起，您的购物车中没有商品!"
			End If
			Amount = 1 : isshop_str="": isshop_n=0
			Do While Not RS.EOF
				 Amount = KS.ChkClng(KS.S( "Q_" & RS("CartID")))
				 If Amount <= 0 Then 
					Amount = KS.ChkClng(RS("Amount"))
					If Amount <= 0 Then Amount = 1
				 End If
				 
			If RS("AttrID")<>0 Then 
			  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RS("AttrID"))
			  If Not RSAttr.Eof Then
				TotalNum = RSAttr("amount")
				Price_Member=RSAttr("Price")
				CurrWeight=RSAttr("Weight")
			  Else
				TotalNum = RS("TotalNum")
				Price_Member=RS("Price_Member")
				CurrWeight=RS("Weight")
			  End If
			  RSAttr.CLose:Set RSAttr=Nothing
			Else
				TotalNum = RS("TotalNum")
				Price_Member=RS("Price_Member")
				CurrWeight=RS("Weight")
			End If	 
				 
				 
			
			IF TotalNum < Amount Then
				Amount = 1
				response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]暂时库存不足，最多只能购买" & TotalNum & RS("unit") & "！');history.back(-1);</script>" 
				response.End()
			End IF
			Dim ProDiscount:ProDiscount=""
			Dim ProScore:ProScore=0
			IF RS("IsLimitBuy")<>"0" and RS("LimitBuyAmount") < Amount Then
				Amount = 1
				response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]还剩" & RS("LimitBuyAmount") & RS("unit") & "供抢购！');history.back(-1);</script>" 
				rs.close:set rs=nothing
				response.End()
			End If
			
			Conn.Execute("Update KS_ShoppingCart Set Amount=" & Amount & " where cartid=" & rs("cartid"))

			Call CheckProductNum(RS)
			ProBuyAttr=""

			If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
			   RealPrice=Session("ChangeBuyPrice")
			ElseIf RS("IsLimitBuy")="1" And Now>RS("LimitBuyBeginTime") And RS("LimitBuyEndTime")>Now And RS("LimitBuyAmount")>0 Then
			   RealPrice=RS("LimitBuyPrice")
			   ProBuyAttr="<span style='color:green'>(限时抢购)</span>"
			   ProDiscount="---"
			ElseIf RS("IsLimitBuy")="2" And RS("LimitBuyAmount")>0 Then
			   RealPrice=RS("LimitBuyPrice")
			   ProBuyAttr="<span style='color:blue'>(限量抢购)</span>"
			   ProDiscount="---"
			ElseIf Not Conn.Execute("Select price From KS_ShopBundleSelect Where username='" & GetUserID & "' and Pid=" & RS("ID") &" And ProID in(select proid from KS_ShoppingCart where flag=0 and username='" & getuserid & "')").EOF Then
			   RealPrice=Conn.Execute("Select price From KS_ShopBundleSelect Where username='" & GetUserID & "' and Pid=" & RS("ID") &" And ProID in(select proid from KS_ShoppingCart where flag=0 and username='" & getuserid & "')")(0)
			   ProBuyAttr="<span style='color:blue'>(捆绑促销)</span>"
			   ProDiscount="---"
			   
			ElseIF Cbool(KSUser.UserLoginChecked)=true Then
			  Dim Discount:Discount=KS.U_S(KSUser.GroupID,17)
			  Dim JFDiscount:JFDiscount=KS.U_S(KSUser.GroupID,18)
			   If Not IsNumeric(Discount) Then Discount=0
			   If Not IsNumeric(JFDiscount) Then JFDiscount=0
			   If KS.U_S(KSUser.GroupID,21)="1" and rs("vipprice")<>"0" then
				RealPrice=RS("VipPrice")
				ProDiscount="VIP价"
			  ElseIf KS.ChkClng(RS("isdiscount"))=0 or Discount=0 Then
			    RealPrice=Price_Member
				ProDiscount="无"
			  Else
			   RealPrice=KS.GetPrice(Price_Member*discount/10)
			   ProDiscount=Discount & "折"
			  End If
			  If JFDiscount=0 Then
				ProScore=0
			  ElseIf JFDiscount=1 Then
				ProScore=KS.ChkClng(RealPrice) * Amount
			  Else
				ProScore=KS.ChkClng(RealPrice*JFDiscount * Amount)
			  End If
			    if JFDiscount<>0 and JFDiscount<>1 and KS.ChkClng(rs("isdiscount"))=1 then ProDiscount=ProDiscount & " <font color=green>" & JFDiscount & "</font>倍积分"

			Else
			  RealPrice=Price_Member
			End If
			
			
			
			IF Amount>=KS.ChkClng(rs("FreeShipping")) And KS.ChkClng(rs("FreeShipping")) <>0 Then 
			 If FreeShipping="" Then FreeShipping=RS("Title") Else FreeShipping=FreeShipping & "," & RS("Title")
			 TotalWeight=-1
			 ProBuyAttr=ProBuyAttr &"<span style='color:red'>(免邮)</span>"
			Else
			 If NOFreeShipping="" Then NOFreeShipping=RS("Title") Else NOFreeShipping=NOFreeShipping & "," & RS("Title")
			End If
			if TotalWeight<>-1 then
			If IsNumerIc(RS("Weight")) Then TotalWeight=TotalWeight+CurrWeight*amount
			end if
			if ks.chkclng(rs("WholesaleNum"))<>0 and amount>=ks.chkclng(rs("WholesaleNum")) then realPrice=rs("WholesalePrice")
		
			
			dim Score:Score=KS.ChkClng(rs("Score"))
			  if KS.ChkClng(rs("istype"))=1 then
			  	if Score<>0 then
					if KSUser.GetScore() >= Score*Amount  then
						isscore_str="<b><font color=""#006600"">"& Score*Amount &"积分"
						if RealPrice*Amount > 0 then
							isscore_str=isscore_str&"+￥"& KS.GetPrice(RealPrice*Amount) &"元</font></b>"
						end if
						IsScore_z=IsScore_z + (Score*Amount) 
					else
						isscore_str="<b><font color=""#FF0000"">(积分不够)</font></b>" :IsScore_no=1
					end if
					if isshop_str="" then isshop_str=RS("Title") else isshop_str=isshop_str&","&RS("Title")
				else
					isshop_n=isshop_n+1	
				end if
			  else
			  	istype=0	
			  end if
			  istype_str=istype_str&KS.ChkClng(rs("istype"))
			TotalPrice=TotalPrice+Round(RealPrice*Amount,2)
			TotalScore=TotalScore+ProScore
			Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
			If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
			
			
			
					  CartStr=CartStr&"<tr class='tdbg' height=""25"" align=""center""> " & vbcrlf
					  CartStr=CartStr&"  <td>" & RS("CartID") & "</td>" & vbcrlf
					  CartStr=CartStr&"  <td align=""left""><table cellspacing=""0"" cellpadding=""0""><tr><td width=""60""><img src=""" & PhotoUrl & """ alt=""" & RS("Title") & """ width=""50"" height=""50"" align=""left"" style=""border:1px solid #ccc;padding:2px""/></td><td style='word-break:break-all;height:20px'><a href=""" & KS.GetItemUrl(5,rs("tid"),rs("id"),rs("fname"),rs("adddate")) & """ target=""_blank"" title='" & rs("title") & "'>" & RS("Title") 
					  If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
					  CartStr=CartStr& "(换购)"
					  Else
					  CartStr=CartStr& ProBuyAttr
					  End If
					  CartStr=CartStr&"</a>" & RS("Attr")  & "</td></tr></table></td>" & vbcrlf
					  CartStr=CartStr&"  <td>" & Amount & "</td>" & vbcrlf
					  CartStr=CartStr&"	<td>￥" & KS.GetPrice(Price_Member) & "</td>" & vbcrlf
					  If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
					  CartStr=CartStr&"	<td align='center'>换购</td>" & vbcrlf
					  ElseIf ProDiscount<>"" Then
					  CartStr=CartStr&"	<td align='center'>" & ProDiscount & "</td>" & vbcrlf
					  Else
					  CartStr=CartStr&"	<td align='center'>无折扣</td>" & vbcrlf
					  End If
					  
					 
					   if isscore_str="" then
					   		  CartStr=CartStr&"	<td style='color:#ff3300;font-size:14px;font-weight:bold'>￥" & KS.GetPrice(RealPrice) & "</td>" & vbcrlf
							 CartStr=CartStr&"	<td> "&isscore_str&"￥" & KS.GetPrice(Round(RealPrice*Amount,2))  & "</td>" & vbcrlf
					   else
					   
					         CartStr=CartStr&"	<td style='color:#ff3300;font-size:14px;font-weight:bold'><span>" & Score & "</span>积分"
							   if RealPrice > 0 then
									CartStr=CartStr& " + ￥"& KS.GetPrice(RealPrice) &"元"
							   end if
							   CartStr=CartStr& "</td>" & vbcrlf
					   
							 CartStr=CartStr&"	<td> "&isscore_str&"</td>" & vbcrlf
					   end if
					 
					  
		              CartStr=CartStr&"	<td>" &ProScore & " 分</td>" & vbcrlf
					  CartStr=CartStr&"</tr>" & vbcrlf
		              CartStr=CartStr & GetBundleSalePro(RS("ID"),false)   '获得捆绑促销的商品
					  
				 RS.MoveNext
				 Loop
			RS.close:set RS=nothing
    End If
	if IsScore_no=1 then
		 KS.Die "<script>alert('积分不够不能购买商品，请重新选择！');location.href='shoppingcart.asp?istype=1';</script>"
	end if
	if KSUser.GetScore()<IsScore_z then
		KS.Die "<script>alert('积分不够不能购买商品，请重新选择！');location.href='shoppingcart.asp?istype=1';</script>"
	end if 
	if InStr(istype_str,"01")>0 or InStr(istype_str,"10") and isshop_n>0  then
		 KS.Die "<script>alert('温馨提示：\n\n由于积分购买商品，不能和非积分购买商品合并付款，请重新选择！');location.href='shoppingcart.asp?istype=1';</script>"
	end if
	 if FreeShipping<>"" and NoFreeShipping<>"" Then
	   KS.Die "<script>alert('温馨提示：\n\n由于商品“" & FreeShipping & "”免邮费，不能和非免邮商品合并付款，请重新选择！');location.href='shoppingcart.asp';</script>"
	 End If
	

     Dim LMStr:LMStr=packstr
	   
     CartStr=CartStr & LMStr          '礼包

	CartStr=CartStr&"<tr class='tdbg'> " & vbcrlf
	CartStr=CartStr&" <td colspan=""7"" style='padding-left:20px;font-size:14px'><strong>结算信息：</strong>" & vbcrlf
	
	'===========================计算运费===============================
	 Dim ExpressCompany,money
	 dim freeDelivery  '满足一定金额免邮
	 if isnumeric(ks.setting(207)) then freeDelivery=round(ks.setting(207),2) else freeDelivery=0
	 if TotalPrice>=freeDelivery and freeDelivery<>0 then
	    totalweight=-1
	    DeliveryMoney=0
		CartStr=CartStr&"<script>$('#jgxx').html('免邮');</script>"
	 else
		 DeliveryMoney=KS.GetFreight(KS.ChkClng(DeliverType),defaultcity,totalweight,ExpressCompany)
		 If DeliveryMoney=-1 Then	DeliveryMoney=0	
		 if (deliveryMoney=0) then 
		  CartStr=CartStr&"<script>$('#jgxx').html('免邮');</script>"
		 else
		  CartStr=CartStr&"<script>$('#jgxx').html('" & deliveryMoney & "元');</script>"
		 end if
	 end if
	 
	

	'====================================================================        
    CartStr=CartStr&" <div style='margin-left:20px'>商品金额：<span id='ordergoodsmoney'>"
	CartStr=CartStr& FormatNumber(TotalPrice,2,-1) & "</span>&nbsp;元 "
	CartStr=CartStr&"+ 运费：<span id='orderyf'>" &DeliveryMoney &"</span>元 + 税费：<span id='ordertax'>0</span>元 <span  id=""ORDER_COSTS""></span>&nbsp;&nbsp;&nbsp;&nbsp;<strong>应付总额：</strong>"
	CartStr=CartStr&"<font color=red>￥<span id='ordertotalmoney'>" & KS.GetPrice(TotalPrice+DeliveryMoney) & "</span></font>&nbsp;元"
	if IsScore_z>0 then
		CartStr=CartStr&" + <font color=red>"& IsScore_z & "</font>&nbsp;积分 "
	end if
	CartStr=CartStr&"&nbsp;&nbsp;&nbsp;赠送积分：<span id=""ORDER_SCORE"" style=""color:green"">" & KS.ChkClng(TotalScore) & "</span> 分<input type='hidden' name='TRealTotalPrice' id='TRealTotalPrice' value='" & TotalPrice+DeliveryMoney & "'><input type='hidden' name='RealTotalPrice' id='RealTotalPrice' value='" & TotalPrice+DeliveryMoney & "'><input type='hidden' name='usezf' id='usezf' value='0'></div>" & vbcrlf
    CartStr=CartStr&"  </td>" &vbcrlf
    CartStr=CartStr&"</tr>" & vbcrlf
	
	'使用优惠券
	IF Cbool(KSUser.UserLoginChecked)=true  and LMStr="" and instr(ProBuyAttr,"抢购")=0 Then
		CartStr=CartStr & "<tr class='tdbg'><td colspan='10'><table border=0 cellpadding=4 cellspacing=2 width='90%' align='center'><tr><td style='text-align:left;width:170px'>" & vbcrlf
	
	 If KS.Setting(181)="1" Then
		Dim MyScore:MyScore=KSUser.GetScore()
		CartStr=CartStr & " <div><strong><label><input onclick=""$('#sss').show();$('#ccc').hide();"" type='radio' name='yhlx' value='1' checked>使用积分抵扣订单金额</label></strong></div>" & vbcrlf  
		CartStr=CartStr & " <div><strong><label><input type='radio' onclick=""$('#sss').hide();$('#ccc').show();"" name='yhlx' value='2'>使用优惠券抵扣订单金额</label></strong></div>"
	Else
	    CartStr=CartStr & " <div><strong><label><input type='hidden' name='yhlx' value='2'>使用优惠券抵扣订单金额</label></strong></div>"
	End If  
	    CartStr=CartStr & "</td><td style='text-align:left'>"
	If KS.Setting(181)="1" Then
		CartStr=CartStr & "<div id='sss'>您当前可用积分 <font color=green>" & MyScore & "</font> 分,花费<input onkeyup=""this.value=this.value.replace(/\D/g,'');"" onafterpaste=""this.value=this.value.replace(/\D/g,'')"" type='text' name='myscore' id='myscore' value='" & MyScore & "' size='6' style='text-align:center'>积分用于抵扣订单费用 <input type='button' value='使用' onclick=""userscore(" & KS.ChkClng(TotalScore)&","&MyScore&")"" class='button'/></div>"
		CartStr=CartStr & "<div id='ccc' style='display:none'>"
	Else
		CartStr=CartStr & "<div id='ccc'"
	End If
	
 
	
	CartStr=CartStr & "选择已有的优惠券" & GetCouponOptionList & "或者输入优惠券号<input type=""text"" id=""couponnum"" name=""couponnum"" size=""10""> <input type=""button"" value=""验证"" onclick=""validateCoupon()""></div>"
	
	CartStr=CartStr & "</td></tr></table></td>"
	CartStr=CartStr & "</tr>"
	End If



    CartStr=CartStr&"</table>" & vbcrlf
			 	 
		   FileContent = Replace(FileContent,"{$ShowShoppingCart}",CartStr)
		   If Cbool(KSUser.UserLoginChecked)=False Then 
		   FileContent = Replace(FileContent,"{$ShowLoginTips}","<strong><font color=ff6600>温馨提示：您还没有注册或登录。享受更多会员优惠，请先<a href=""../user/login/"">登录</a>或<a href=""../?dp=reg"" target=""_blank"">注册</a>成为商城会员！</font></strong>")
           Else
		   FileContent = Replace(FileContent,"{$ShowLoginTips}","亲爱的" & KSUser.UserName &"! 您的个人信息-> 用户组："&KS.GetUserGroupName(KSUser.GroupID)&"&nbsp;可用资金：&nbsp;<font color=""green"">" & KS.GetPrice(KSUser.GetUserInfo("Money")) & "</font>&nbsp;元 " & KS.Setting(45) & "：&nbsp;<font color=green>" & KSUser.GetUserInfo("Point") & "</font>&nbsp;" & KS.Setting(46)&" 积分：&nbsp;<font color=""green"">" & KSUser.GetUserInfo("Score") & "</font>&nbsp;分")
		   End If
		   FileContent = ReplaceUserInfo(FileContent)
		   FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent 
		    
End Sub

Function GetCouponOptionList()
    Dim RS,SQL,I,Str,Param,SQLStr
	Param=Param & " and b.status=1 and datediff(" & DataPart_S & ",b.enddate," & SqlNowString & ")<0 And datediff(" & DataPart_S & ",b.begindate," & SqlNowString & ")>0"
	SQLStr="Select a.id,b.title,b.facevalue,a.AvailableMoney From KS_ShopCouponUser A inner Join KS_ShopCoupon B on a.couponid=b.id where AvailableMoney>0 and a.username='" & KSUser.UserName & "'" & Param	   
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open SQLStr,conn,1,1
	If Not RS.Eof Then
	  SQL=RS.GetRows(-1)
	End If
	RS.Close:Set RS=Nothing
	Str="<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
	
	Str=Str & "<select name=""couponid"" onchange=""changeCoupon(this.value)"" style=""width:250px"">"
	Str=Str & "<option value='0'>---请选择---</option>"
	If IsArray(SQL) Then
	 For I=0 To  Ubound(SQL,2)
	  Str=Str & "<option value='" & SQL(0,I) & "'>" & SQL(1,I) & "[面值￥" & round(SQL(2,I),2) &"元 余额￥" & KS.GetPrice(SQL(3,I)) & "元]</option>"
	 Next
	End If
	Str=Str & "</select>"
	GetCouponOptionList=Str
  End Function



  Function  ReplaceUserInfo(FileContent)
	 FileContent=Replace(FileContent,"{$Contactup}",GetContactup)
     FileContent=Replace(FileContent,"{$ContactMan}",KSUser.GetUserInfo("RealName"))
	 FileContent=Replace(FileContent,"{$ContactMan}",KSUser.GetUserInfo("RealName"))
     FileContent=Replace(FileContent,"{$Address}",KSUser.GetUserInfo("Address"))
     FileContent=Replace(FileContent,"{$ZipCode}",KSUser.GetUserInfo("Zip"))
     FileContent=Replace(FileContent,"{$Phone}",KSUser.GetUserInfo("OfficeTel"))
     FileContent=Replace(FileContent,"{$Email}",KSUser.GetUserInfo("Email"))
     FileContent=Replace(FileContent,"{$Mobile}",KSUser.GetUserInfo("Mobile"))
     FileContent=Replace(FileContent,"{$QQ}",KSUser.GetUserInfo("QQ"))
	 FileContent=Replace(FileContent,"{$TotalWeight}",TotalWeight)
	 FileContent=Replace(FileContent,"{$TaxRate}",KS.Setting(65))
	 FileContent=Replace(FileContent,"{$IncludeTax}",KS.Setting(64))
	 FileContent=Replace(FileContent,"{$PaymentType}",GetPaymentTypeStr)
	 DeliveryStr=DeliveryStr & "<input type='hidden' id='mustyf' value='" & KS.Setting(180) & "'>"
	 FileContent=Replace(FileContent,"{$DeliveryType}",DeliveryStr)
	 ReplaceUserInfo=FileContent
  End Function
  Function GetContactup()
    GetContactup=GetContactup & "<input type=""hidden"" name=""istype"" value="""& istype &""">"&vbcrlf
 	If Cbool(KSUser.UserLoginChecked)=true Then
		'GetContactup=GetContactup &"<input name=""up-input"" class=""up-input"" onclick=""getupinput()"" value=""上次收货信息记录"" type=""button"">"
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 50 * From KS_ShopUserOrder where UserName='"& KSUser.UserName &"' order by id desc",conn,1,1
		If Not RS.EOF And  Not RS.BOF Then
			Do While Not RS.Eof
			     
				GetContactup=GetContactup & "<input name=""radio_cds"" onclick=""getupinput(this)"" type=""radio"" value="""& rs("id") &""" />" & rs("ContactMan") & "&nbsp;" & rs("Address")&vbcrlf
				GetContactup=GetContactup & "<span><input name=""up_ContactMan"" type=""hidden"" value="""& rs("ContactMan") &""">"&vbcrlf
				GetContactup=GetContactup & "<input name=""up_Address"" type=""hidden"" value="""& rs("Address") &""">" &vbcrlf
				GetContactup=GetContactup & "<input name=""up_ZipCode"" type=""hidden"" value="""& rs("ZipCode") &""">" &vbcrlf
				GetContactup=GetContactup & "<input name=""up_Mobile"" type=""hidden"" value="""& rs("Mobile") &""">" &vbcrlf
				GetContactup=GetContactup & "<input name=""up_Phone"" type=""hidden"" value="""& rs("Phone") &""">" &vbcrlf
				GetContactup=GetContactup & "<input name=""up_QQ"" type=""hidden"" value="""& rs("QQ") &""">" &vbcrlf
				GetContactup=GetContactup & "<input name=""up_Email"" type=""hidden"" value="""& rs("Email") &"""></span><br/>" &vbcrlf
			RS.MoveNext
			Loop
		End IF
		RS.Close:Set RS=Nothing
	end if
  End Function 
  '付款方式
  Function GetPaymentTypeStr()
   Dim DiscountStr,SQL,I,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select TypeID,TypeName,IsDefault,Discount from KS_PaymentType order by orderid",conn,1,1
   If Not RS.Eof Then
     SQL=RS.GetRows(-1)
   End IF
   RS.Close:Set RS=Nothing
   GetPaymentTypeStr="<select name='PaymentType'>"
   If IsArray(SQL) Then
	   For I=0 To UBound(SQL,2)
		 If SQL(3,I)<>100 Then
		  DiscountStr="折扣率 " & SQL(3,I) & "%"
		 Else
		  DiscountStr=""
		 End iF
		 If SQL(2,I)=1 Then
		GetPaymentTypeStr=GetPaymentTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & " " & DiscountStr & "</option>"
		 Else
		GetPaymentTypeStr=GetPaymentTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & " " & DiscountStr & "</option>"
		End If
	   Next
  End If
   GetPaymentTypeStr=GetPaymentTypeStr & "</select>"
  End Function
  
  '发货方式
  Function GetDeliveryTypeStr()
	   Dim j,rss,rsss
	   Dim DiscountStr,SQL,I,RS,defaulttips,tocity
		set rsss=conn.execute("select top 1 * from KS_Delivery where isDefault=1 order by orderid")
		if not rsss.eof then 
			 DeliverType=rsss("expressid"):defaultcity=rsss("defaultcity")
			 if rsss("W_fee")<>"0" then
			 defaulttips="首重：<span>"& KS.GetPrice(round(rsss("fweight"),2)) & "</span> kg 首重价格：<span>" & KS.GetPrice(rsss("carriage")) & "</span> 元 续重价格：<span>" & KS.GetPrice(round(rsss("C_fee")/rsss("W_fee"),2))&" </span>元/kg"
			end if
		end if
		rsss.close:set rsss=nothing
	   if KS.IsNul(defaultcity) Then defaulttips="选择送货路线确定运费！":defaultcity="选择送货地点" Else tocity=defaultcity
	   Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "select TypeID,TypeName,IsDefault from KS_DeliveryType where typeid in(select ExpressID from KS_Delivery) order by orderid,TypeID",conn,1,1
	   If Not RS.Eof Then
		 SQL=RS.GetRows(-1)
	   End IF
	   RS.Close:Set RS=Nothing
	   GetDeliveryTypeStr="<strong>快递公司：</strong><select onchange=""ajshowdata($('#choosecity').val())"" name='DeliverType' id='DeliverType'>"
	   If IsArray(SQL) Then
	   For I=0 To UBound(SQL,2)
		 If trim(DeliverType)=trim(sql(0,i)) Then
		GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & "</option>"
		 Else
		GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & "</option>"
		End If
	   Next
	  End If
	   GetDeliveryTypeStr=GetDeliveryTypeStr & "</select>"
	  
	   GetDeliveryTypeStr=GetDeliveryTypeStr & " <input type=""hidden"" value=""" & tocity & """ name=""tocity"" id=""tocity""/> <span style='position:relative'><input class=""tocity"" style='text-align;left' name='' id='choosecity' type='button' value='" & defaultcity & "'  onclick=""$('#showprovn').show();if(this.getBoundingClientRect().top>300){showprovn.style.top=(this.offsetHeight-showprovn.offsetHeight)}else{showprovn.style.top='0'};""><span id='showprovn' onclick=""this.style.display='none'"" class='showcity'>"&_
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
	  End Function
  


End Class
%>

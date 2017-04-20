<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
select case request("Action")
 case "changegroupcartnum" Call changegroupcartnum()
 case "changegroupcartdelivery" Call changegroupcartdelivery()
 case "Shop_ValidateCoupon" Call Shop_ValidateCoupon()
 case "Shop_GetCoupon" Call Shop_GetCoupon()
 case "checkscore" call CheckScore()
 case "Shop_BrandOption" call Shop_BrandOption()
 case "Shop_CheckProID" call Shop_CheckProID()
 case "deleteproitem" Call deleteproitem()
 case "getSpecification" call getSpecification()
 case "getstock" call getstock()
 case "getcartstock" call getcartstock()
 case "Shop_LimitBuyTask" call Shop_LimitBuyTask()
 case "Shop_SearchProduct" call Shop_SearchProduct()
 case "Shop_ShowPrice" call Shop_ShowPrice()
 case else call delivery()
end select
Set KS=Nothing


sub Shop_LimitBuyTask()
    Dim LimitBuyTaskID:LimitBuyTaskID=KS.ChkClng(Request("LimitBuyTaskID"))
	Dim TaskType:TaskType=KS.ChkClng(Request("TaskType"))
    Dim RST:Set RST=Conn.Execute("Select ID,taskname from KS_ShopLimitBuy Where TaskType=" & TaskType &" and Status=1 Order by id desc")
	Do While NOt RST.Eof
		If LimitBuyTaskID=RST(0) Then
				Response.Write "<option value='" & RST(0) & "' selected>" & RST(1) & "</option>"
		Else
				Response.Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
		End If
	   RST.MoveNext
	Loop
	RST.CLose 
	Set RST=Nothing
	KS.Die ("")
end sub

sub Shop_SearchProduct()
  dim proids:proids=KS.CheckXSS(KS.S("proids"))
  dim title:title=UnEscape(request("Title"))
   Dim RST:Set RST=Conn.Execute("Select top 100 ID,Title,Price_Member from KS_Product Where deltf=0 and verific=1 and (proid='" & proids &"' or title like '%" & KS.DelSQL(title) &"%') Order by id desc")
   dim i:i=0
	Do While NOt RST.Eof
	    if i>0 then response.write "§"
		Response.Write  rst(0) &"◇" & rst(1) & "◇" & rst(2)
		i=i+1
	   RST.MoveNext
	Loop
	RST.CLose 
	Set RST=Nothing
	KS.Die ("")
end sub

sub Shop_ShowPrice()

    '抢购结束的商品恢复状态
	conn.execute("update ks_product set islimitbuy=0,LimitBuyTaskID=0 where LimitBuyTaskID in(select id from KS_ShopLimitBuy where datediff(" & DataPart_S &",LimitBuyEndTime," & SqlNowString&")>0)")

   dim ProID:ProID=KS.ChkClng(Request("ID"))
   if ProID=0 Then KS.Die ""
   dim str:str=""
   dim LimitBuyBeginTime:LimitBuyBeginTime=""
   dim LimitBuyEndTime:LimitBuyEndTime=""
   dim SqlStr:SqlStr = "SELECT Top 1 Price_Member,isdiscount,islimitbuy,limitbuyprice,istype,score,VipPrice,LimitBuyBeginTime,LimitBuyEndTime FROM KS_Product a Left Join KS_ShopLimitBuy b ON A.LimitBuyTaskID=b.id Where a.ID=" & ProID
   dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
        RS.Open SqlStr, conn, 1, 1
        If Not RS.bof  Then
		  dim vipprice:vipprice=RS("VipPrice")
		  if KS.IsNul(vipprice) Then VipPrice=RS("Price_Member")
		  dim price_member:price_member=KS.GetPrice(RS("Price_Member"))
		  if price_member<1 then	price_member=price_member
		     if KS.ChkClng(rs("istype"))=0 then
			 	price_member="￥" &price_member &"元"
			 else
				if KS.ChkClng(price_member)=0 then
					price_member=rs("score") &"分 "	 
				else
					price_member=rs("score") &"分+" & price_member&"元"	 	
				end if
				
			 end if
		   VipPrice=KS.GetPrice(VipPrice)

		  If RS("IsLimitBuy")<>0 Then
		    LimitBuyBeginTime=RS("LimitBuyBeginTime")
		    LimitBuyEndTime=RS("LimitBuyEndTime")
		    str="促 销 价：<span class=""vipprice"">￥" & RS("LimitBuyPrice") & "元</span>  <span style=""color:#888"">剩余:<span id=""time" & ProID &"""></span></span><br/>"
		    str=str & "VIP会员价：<span class=""vipprice"">￥" & VipPrice & "元</span>"
		    str=str & "<br/>普通会员价：<span class=""price"" style=""text-decoration:line-through"">" & price_member & "</span>"
		  Else
		    str=str & "VIP会员价：<span class=""vipprice"">￥" & VipPrice & "元</span>"
		    str=str & "<br/>普通会员价：<span class=""price"">" & price_member& "</span>"
		  End If
			
		End If
   RS.Close
   Set RS=Nothing
   Response.Write "document.write('" & str & "');"
   If LimitBuyBeginTime<>"" Then
     KS.Die "show_date_time('" & LimitBuyBeginTime &"','" & LimitBuyEndTime &"','time" & ProID &"',1);"
   End If
   KS.Die ""
end sub




sub getstock()
  dim rs:set rs=conn.execute("select top 1 amount from KS_ShopSpecificationPrice where id=" & KS.ChkClng(KS.S("Attrid")))
  if not rs.eof then
    Response.Write "var data={'amount':'" & rs(0) & "'}"
  else
    Response.Write "var data={'amount':'0'}"
  end if
  rs.close:set rs=nothing
end sub

sub getcartstock()
  dim proid:proid=KS.ChkClng(Request("proid"))
  dim attrid:attrid=KS.ChkClng(Request("attrid"))
  dim f:f=KS.ChkClng(Request("f"))
  dim cartid:cartid=KS.ChkClng(Request("cartid"))
  dim buynum:buynum=KS.ChkClng(Request("buynum"))
  dim typeflag:typeflag=KS.ChkClng(Request("type"))
  dim rs,stock,currnum
  if attrid<>0 then
   set rs=conn.execute("select top 1 amount from KS_ShopSpecificationPrice where id=" & attrid)
  else
   set rs=conn.execute("select top 1 totalnum from KS_Product where id=" & proid)
  end if
  if not rs.eof then stock=rs(0) else stock=0
  rs.close
  
  if buynum<>0 then
	  if (buynum<=0) then ks.die escape("err|购买数量不能少于1")
	  if (buynum>stock) then ks.die escape("err|对不起，购买数量最多为" & stock)
	    if typeflag=1 then
		 conn.execute("update KS_ShopBundleSelect set amount=" & buynum & " where id=" & cartid)
		else
		 conn.execute("update KS_ShoppingCart set amount=" & buynum & " where cartid=" & cartid)
		end if
		 ks.die "succ"
  else
      if typeflag=1 then
	   rs.open "select top 1 * from KS_ShopBundleSelect where id=" & cartid,conn,1,1
	  else
	   rs.open "select top 1 * from KS_ShoppingCart where cartid=" & cartid,conn,1,1
	  end if
	  if not rs.eof then currnum=rs("amount") else currnum=0
	  rs.close
	  set rs=nothing
	  if (f=0 and currnum<=1) then ks.die escape("err|购买数量不能少于1")
	  if (f=0 and currnum>0) then
	     if typeflag=1 then
		 conn.execute("update KS_ShopBundleSelect set amount=amount-1 where id=" & cartid)
		 else
		 conn.execute("update KS_ShoppingCart set amount=amount-1 where cartid=" & cartid)
		 end if
		 ks.die "succ"
	  end if
	  if f=1 and (currnum+1)>stock then
		  ks.die escape("err|对不起，购买数量最多为" & stock)
	  else
	     if typeflag=1 then
		 conn.execute("update KS_ShopBundleSelect set amount=amount+1 where id=" & cartid)
		 else
		 conn.execute("update KS_ShoppingCart set amount=amount+1 where cartid=" & cartid)
		 end if
		 ks.die "succ"
	  end if
 end if
end sub

sub delivery()
  If KS.IsNul(Request.ServerVariables("HTTP_REFERER")) Then Exit Sub
  dim expressid:expressid=KS.ChkClng(KS.S("expressid"))
  Dim City:City=KS.DelSQL(UnEscape(Request("City")))
  If expressid=0 Or KS.IsNul(City) Then KS.Die Escape("error|参数出错！|0")
  
  '=====================================计算订单运费==========================================
  Dim ExpressCompany,totalweight:totalweight=KS.S("totalweight")
  if Not IsNumeric(totalweight) Then totalweight=0
  if KS.ChkClng(totalweight)=-1 then KS.Die Escape("success|免邮|0")
  Dim DeliveryMoney:DeliveryMoney=KS.GetFreight(expressid,city,totalweight,ExpressCompany) 
  If DeliveryMoney=-1 Then DeliveryMoney=0
  '============================================================================================
  
  If DataBaseType=1 Then
  sql="select top 1 * From KS_Delivery Where ExpressID="& expressid &" and (convert(varchar(200),tocity)='全国统一运费' or convert(varchar(200),tocity)='' or tocity is null)"
  Else
  sql="select top 1 * From KS_Delivery Where ExpressID="& expressid &" and (tocity='全国统一运费' or tocity='')"
  End If
  set rs = Server.CreateObject("ADODB.recordset")
  rs.open sql,conn,1,1
  if not rs.eof then
    KS.Die Escape("success|首重：<span>"& formatnumber(round(rs("fweight"),2),2,-1) & "</span> kg 首重价格：<span>" & formatnumber(rs("carriage"),2,-1) & "</span> 元 续重价格：<span>" & formatnumber(round(rs("C_fee")/rs("W_fee"),2),2,-1)&" </span>元/kg" & "|" & DeliveryMoney &"")
	RS.Close : Set RS=Nothing
  end if
  rs.close
  sql="select Top 1 * from KS_Delivery where ExpressID="& expressid &" and tocity like '%"&city&"%'"
  rs.Open sql, conn,1,1
  if not rs.EOF then
    KS.Die Escape("success|首重：<span>"& formatnumber(round(rs("fweight"),2),2,-1) & "</span> kg 首重价格：<span>" & formatnumber(rs("carriage"),2,-1) & "</span> 元 续重价格：<span>" & formatnumber(round(rs("C_fee")/rs("W_fee"),2),2,-1)&" </span>元/kg" & "|" & DeliveryMoney &"")
	RS.Close : Set RS=Nothing
  End If
  If DataBaseType=1 Then
   sql="select Top 1 * from KS_Delivery where ExpressID="& expressid &" and (convert(varchar(200),tocity)='' or tocity is null)"
  Else
   sql="select Top 1 * from KS_Delivery where ExpressID="& expressid &" and (tocity='' or tocity is null)"
  End If
  set rs=conn.execute(sql)
  If RS.Eof Then
   response.write "error|你选择的路线暂时还没开通快递业务，请重新选择！"
  else  '全国统一运费
    KS.Echo Escape("success|首重：<span>"& formatnumber(round(rs("fweight"),2),2,-1) & "</span> kg 首重价格：<span>" & formatnumber(rs("carriage"),2,-1) & "</span> 元")
	If rs("W_fee")<>0 and rs("C_fee")<>0 THEN
	KS.Echo Escape(" 续重价格：<span>" & formatnumber(round(rs("C_fee")/rs("W_fee"),2),2,-1)&" </span>元/kg" & "")
	End If
	KS.Echo "|" & DeliveryMoney
  end if
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
end sub

Sub changegroupcartnum()
  Dim CartID:CartID=KS.ChkClng(KS.S("CartID"))
  Dim Num:Num=KS.ChkClng(KS.S("Num"))
  Dim TotalPrice
  Dim DeliverType:DeliverType=KS.ChkClng(KS.S("DeliverType"))
  Dim tocity:tocity=KS.DelSQL(Unescape(Request("tocity")))
  If CartID=0 or Num=0 Then KS.Die Escape("error|参数出错啦!")
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "select top 1 g.subject,g.id,g.LimitBuyNum,g.price,c.amount from KS_ShoppingCart c inner join KS_GroupBuy g on c.proid=g.id where c.cartid=" & cartid,conn,1,1
  if RS.Eof And RS.Bof Then
   RS.Close:Set RS=Nothing
   KS.Die Escape("error|参数出错啦,找不到记录!")
  End If
  Dim LimitBuyNum:LimitBuyNum=KS.ChkClng(RS("LimitBuyNum"))
  If LimitBuyNum<>0 Then
   If Num>LimitBuyNum Then
     RS.Close:Set RS=Nothing
     KS.Die Escape("error|对不起，本商品最多只能购买“ " &LimitBuyNum & "件 ”!")
   End If
  End If
  Dim Price:Price=RS("Price")
  Dim OriginPrice:OriginPrice=RS("Amount")*Price
  RS.CLOSE
  Conn.Execute("Update KS_ShoppingCart Set amount=" & Num & " where cartid=" & CartID)
  totalweight=0
  TotalPrice=0
  RS.Open "select c.amount,g.weight,g.price from KS_ShoppingCart c  inner join KS_GroupBuy g on c.proid=g.id where c.flag=1 and c.username='" & KS.R(KS.S("UserID")) & "' Order By c.cartid",conn,1,1
  Do While Not RS.Eof
   TotalPrice=TotalPrice+RS(0)*RS(2)
   totalweight=totalweight+RS(0)*RS(1)
   RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
  Dim DeliveryMoney
  If tocity<>"" Then
  DeliveryMoney=KS.GetFreight(DeliverType,ToCity,totalweight,"")
  iF DeliveryMoney=-1 Then DeliveryMoney=0
  Else
  DeliveryMoney=0
  End If
  If totalweight=0 Then DeliveryMoney=0
  KS.Die "success|" & (Num * Price) & "|" & FormatNumber(TotalPrice+DeliveryMoney,2,-1,-1) & "|" & DeliveryMoney
End Sub

Sub changegroupcartdelivery()
  Dim DeliverType:DeliverType=KS.ChkClng(KS.S("DeliverType"))
  Dim tocity:tocity=KS.DelSQL(Unescape(Request("tocity")))
  If DeliverType=0 or tocity="" Then KS.Die Escape("error|出错啦，请选择发往城市!")
  Dim totalweight:totalweight=0
  Dim TotalPrice:TotalPrice=0
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "select c.amount,g.weight,g.price from KS_ShoppingCart c  inner join KS_GroupBuy g on c.proid=g.id where c.flag=1 and c.username='" & KS.R(KS.S("UserID")) & "' Order By c.cartid",conn,1,1
  Do While Not RS.Eof
   TotalPrice=TotalPrice+RS(0)*RS(2)
   totalweight=totalweight+RS(0)*RS(1)
   RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
  Dim DeliveryMoney
  DeliveryMoney=KS.GetFreight(DeliverType,ToCity,totalweight,"")
  iF DeliveryMoney=-1 or totalweight=0 Then DeliveryMoney=0
  KS.Die "success|" & FormatNumber(TotalPrice+DeliveryMoney,2,-1,-1) & "|" & DeliveryMoney
End Sub

'验证优惠券
Sub Shop_ValidateCoupon()
  Dim CouponNum:CouponNum=Trim(KS.S("CouponNum"))
  If CouponNum="" Then Exit Sub
  Dim RS:Set RS=Conn.Execute("SELECT Top 1 A.FaceValue,A.MinAmount,A.MaxDiscount,b.AvailableMoney,a.begindate,a.EndDate,a.status FROM KS_ShopCoupon A Inner Join KS_ShopCouponUser B On A.ID=B.CouponID Where B.CouponNum='" & CouponNum & "'")
  If Not RS.Eof Then
       If DateDiff("s",RS("BeginDate"),Now)<0 Then
		Response.Write escape("对不起,您输入的优惠券需要" & RS("BeginDate") & "后才能使用!")
	   ElseIf DateDiff("s",RS("EndDate"),Now)>0 Then
		Response.Write escape("对不起,您输入的优惠券已过使用期限!")
	   ElseIf RS("Status")=0 Then
		Response.Write escape("对不起,您输入的优惠券已被锁定!")
	   ElseIf RS("AvailableMoney")<=0 Then
		Response.Write escape("对不起,您输入的优惠券已用完!")
	   Else
		Response.Write RS(0) & "|" & RS(1) & "|" & RS(2)&"|"&RS(3)
	   End If
  End If
  RS.Close:Set RS=Nothing
End Sub

Sub Shop_GetCoupon()
  Dim CouponUserID:CouponUserID=KS.ChkClng(KS.S("CouponID"))
  If CouponUserID=0 Then Exit Sub
  Dim RS:Set RS=Conn.Execute("SELECT Top 1 FaceValue,MinAmount,MaxDiscount,b.AvailableMoney,a.EndDate,a.status FROM KS_ShopCoupon A Inner Join KS_ShopCouponUser B ON A.ID=B.CouponID Where b.id=" & CouponUserID)
  If Not RS.Eof Then
   If DateDiff("s",RS("EndDate"),Now)>0 Then
    Response.Write escape("对不起,您输入的优惠券已过使用期限!")
   ElseIf RS("Status")=0 Then
    Response.Write escape("对不起,您输入的优惠券已被锁定!")
   ElseIf RS("AvailableMoney")<=0 Then
	Response.Write escape("对不起,您输入的优惠券已用完!")
   Else
    Response.Write RS(0) & "|" & RS(1) & "|" & RS(2)&"|"&RS(3)
   End If
  End If
  RS.Close:Set RS=Nothing
End Sub

Sub CheckScore()
	Dim MyScore:MyScore=KS.ChkClng(Request("myscore"))
	Dim UseScoreMoney:UseScoreMoney=0
	Dim NowMyScore:NowMyScore=KS.ChkClng(Request("score"))
	Dim Money:Money=KS.S("money")
	If Not IsNumeric(Money) Then Money=0
	Dim ScoreRate:ScoreRate=KS.Setting(182)
	If Not IsNumeric(ScoreRate) Then ScoreRate=0
	If KS.ChkClng(ScoreRate)>0 Then
		 Dim LimitTotalMoney:LimitTotalMoney=KS.Setting(183)
		 Dim LimitPer:LimitPer=KS.Setting(184)
		 If Not IsNumeric(LimitTotalMoney) Then LimitTotalMoney=0
		 If Not IsNumeric(LimitPer) Then LimitPer=0
                If MyScore>NowMyScore Then
					Response.Write escape("对不起,您的可用积分只有" & NowMyScore & "分!")
					Exit Sub
				ElseIf round(Money)<round(LimitTotalMoney) and round(LimitTotalMoney)>0 Then
					  Response.Write escape("对不起,系统限定只有订单金额达到" & LimitTotalMoney & "元时才可以使用积分抵用!")
					  Exit Sub
				End If
					 UseScoreMoney=MyScore/ScoreRate
					 If Round(LimitPer)>0 Then
					    dim allowscoremoney:allowscoremoney=round(Money)*Round(LimitPer)/100
					   If Round(UseScoreMoney)> round(allowscoremoney) Then
					    dim allowscore:allowscore=allowscoremoney * ScoreRate
					    Response.Write escape("对不起,系统限定积分抵扣金额不能超过订单总金额的" & LimitPer & "%,您最多可以用" & allowscore & "积分抵扣" & allowscoremoney & "元!")
					    Exit Sub
					   End If
					 End If
			ks.die "succ|"&UseScoreMoney
	End If
	ks.die "succ|"&UseScoreMoney
End Sub



'根据栏目ID得品牌列表
Sub Shop_BrandOption()
  Dim SQL,K,From:From=KS.G("From")
  Dim ClassID:ClassID=KS.G("ClassID")
  If (ClassID="" Or ClassID="0") Then Response.Write Escape("请先选择栏目!"):Response.End
  Dim Str:Str=GetBrandByClassID(ClassID,0)
  If Str="Null" Then
     If From<>"User" Then
     Response.Write Escape("&nbsp;<font color=blue>该栏目下没有添加品牌，请先</font><a onclick=""window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=管理中心 >> 品牌管理 >> <font color=red>新增品牌</font>&ButtonSymbol=GO'"" href='KS.ShopBrand.asp?action=Add&classid=" & classid & "'><font color=red>添加</font></a>")
	 End If
  Else
     If From="User" Then Response.Write Escape("所属品牌：")
	  Response.Write Escape(Str)
  End If
End Sub
		
Function GetBrandByClassID(ClassID,BrandID)
         Dim SXML:set SXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		 SXML.async = false
		 SXML.setProperty "ServerHTTPRequest", true 
		 SXML.load(Server.MapPath(KS.Setting(3)& "config/shopbrand.xml"))
		 if SXML.parseError.errorCode<>0 Then
			Call KS.CreateBrandCache()
		 End If
		 Dim NodeS,Node
		 Set Nodes=SXML.DocumentElement.SelectNodes("item[@classid='" & ClassID &"']")
		 If Nodes.Length>0 Then
		     GetBrandByClassID = "<select name='brandid'>"
			 GetBrandByClassID = GetBrandByClassID & "<option value='0'>-请选择品牌-</option>"
			 For Each Node In Nodes
			  If KS.ChkClng(BrandID)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & Node.SelectSingleNode("@id").text & "' selected>" & Node.SelectSingleNode("brandname").text & "</option>"
			  Else
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & Node.SelectSingleNode("@id").text & "'>" & Node.SelectSingleNode("brandname").text & "</option>"
			  End If
			 Next
			 GetBrandByClassID = GetBrandByClassID &  "</select>"
		 Else
		  GetBrandByClassID="Null" 
		 End If
End Function
'检查商品ID是否可用
Sub Shop_CheckProID()
 Dim proid:proid=KS.DelSQL(UnEscape(Request("proid")))
 Dim ID:ID=KS.ChkClng(KS.S("ID"))
 Dim SQLStr
 If ProID="" Then 
   Response.Write Escape("你没有输入商品编号!")
 Else
   If Id=0 Then
    SqlStr="Select ProID From KS_Product Where ProID='" & ProID & "'"
   Else
    SqlStr="Select ProID From KS_Product Where ID<>" & ID & " and ProID='" & ProID & "'"
   End IF
   If Conn.Execute(SqlStr).Eof Then
    Response.Write Escape("恭喜,该商品编号可用!")
   Else
    Response.Write Escape("对不起,该商品编号已存在!")
   End If
 End If
End Sub

'删除选中的货号
Sub deleteproitem()
 	 Dim UserName : UserName=KS.C("AdminName")
	 Dim Pass : Pass=KS.C("PassWord")
	  if KS.IsNul(UserName) Or KS.IsNul(Pass) Then
	   KS.Die Escape("error|对不起，您没有权限!")
	  End If
	  If Conn.Execute("Select top 1 * From KS_Admin Where UserName='" & UserName & "' and PassWord='" & Pass & "'").eof Then
	    KS.Die Escape("error|对不起，您没有权限!")
	  End If
     Dim ID:ID=KS.ChkClng(KS.S("id"))
	 If Id=0 Then	KS.Die Escape("error|对不起，参数传递出错!")
	 Conn.Execute("Delete From KS_ShopSpecificationPrice Where ID=" & ID)
     KS.Die "success|"
End Sub

Sub getSpecification()
 	 Dim UserName : UserName=KS.C("AdminName")
	 Dim Pass : Pass=KS.C("PassWord")
	  if KS.IsNul(UserName) Or KS.IsNul(Pass) Then
	   KS.Die Escape("error|对不起，您没有权限!")
	  End If
	
    Dim classid:classid=KS.S("ClassID")
	set rst=conn.execute("select title,showtype,svalue from KS_ShopSpecification s inner join KS_ShopSpecificationR r on s.id=r.sid where r.classid='" & classid & "' order by s.orderid,s.id")
			if not rst.eof then
			 dim i,ii,itemarr,tempstr,rsql:rsql=rst.getrows(-1)		 
			    tempstr="<TABLE width='98%' align='left' style=""border:2px solid #efefef"" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
				      tempstr=tempstr & "<tbody>"
					for i=0 to ubound(rsql,2)
					  dim showtype:showtype=rsql(1,i)
				      tempstr=tempstr & "<tr><td style='width:62px' align='right'><strong>" & rsql(0,i) &":</strong></td><td class='atcs' width='1000'><input type='hidden' name='attrtitle" &i & "' id='attrtitle" & i &"' value='" & rsql(0,i) &"'/><input type='hidden' name='ashowtype" &i & "' id='ashowtype" & i &"' value='" & showtype &"'/>"
					    dim sv:sv=split(rsql(2,i),",")
						for ii=0 to ubound(sv)
						  dim tname,showstr
						  if showtype="2" then
						    tname=split(sv(ii),"|")(0):showstr="<input value='" & split(sv(ii),"|")(1) & "' size='5' type='hidden'  name='timg" &i& ii &"' id='timg" &i& ii &"'/><img align='baseline' title='" & tname &"' src='" & split(sv(ii),"|")(1) & "' id='i" & i &ii&"' width='25' height='25'/><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?ChannelID=5&get=min&UpType=Pic&imgname=i" & i &ii&"&FieldName=timg" & I &Ii &"' frameborder=0 scrolling=no width='50' height='26'></iframe>"
						  else
						    tname=sv(ii): showstr=tname
						  end if
						  tempstr=tempstr &"<li><label><input onclick=""getlist(" & ubound(sv) &")"" type='checkbox' name='cc" & i &"' value='" & tname & "'/>" & showstr &"</label></li>"
						next
					  tempstr=tempstr &"</td></tr>"
					 next
					 tempstr=tempstr &"</tbody>"
					 
					 
					 tempstr=tempstr & "<tr class='tdbg'>"
					 tempstr=tempstr & " <td colspan='2' style='padding-left:10px'>"
					
					 
					 tempstr=tempstr & "      <table class='ctable' border='0' cellspacing='1' cellpadding='1'>"
				     tempstr=tempstr & "       <tr class='sort' style='text-align:center;'>"
				     tempstr=tempstr & "       <td  width='100'>货号</td>"
						 redim itemarr(ubound(rsql,2))
						 for i=0 to ubound(rsql,2)
						  itemarr(i)=rsql(2,i)
						  tempstr=tempstr & "       <td width='150' style='display:none' id='tt" & i &"'>" & rsql(0,i) &"</td>"
						 next
				      tempstr=tempstr & "       <td  width='100'>销售价</td>"
				      tempstr=tempstr & "       <td  width='100'>库存</td>"
				      tempstr=tempstr & "       <td  width='100'>重量</td>"
				      tempstr=tempstr & "       <td  width='100'>操作</td>"
					  tempstr=tempstr & "       </tr>"
					  tempstr=tempstr & "       <tbody id='alist'></tbody>"
					 tempstr=tempstr & "      </table>"
					 tempstr=tempstr & " </td>"
					 tempstr=tempstr & "</tr></TABLE>"
					 response.write tempstr
		end if
End Sub
%>
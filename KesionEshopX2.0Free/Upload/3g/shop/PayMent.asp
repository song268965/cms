<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../include/3gCls.asp"-->
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
        Private KS, KSRFObj,KSUser,DomainStr,F_C
		Private TotalPrice,TotalScore,RealPrice,Price_Original,Discount,Amount,TotalWeight
		Private ProductList,DeliverType,DefaultCity,DeliveryStr,DeliveryMoney
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
		<!--#include file="../include/function.asp"-->
		<%
		Public Sub Kesion()
		  Dim Products,i,RS,strsql,CartStr	
		  TotalWeight=0

			
		If KS.ChkClng(KS.Setting(63))=0 And Cbool(KSUser.UserLoginChecked)=false Then Response.Write "<script>alert('本商城设置注册用户才可购买，请先登录!');location.href='../login.asp';</script>"

		          F_C = KSRFObj.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(5,10) &"/payment.html")
				   InitialCommon
				   FCls.RefreshType = "ShoppingPayMent" '设置刷新类型，以便取得当前位置导航等
				   Fcls.RefreshFolderID = "0"        '设置当前刷新目录ID 为"0" 以取得通用标签
				   If KS.ChkClng(KS.Setting(180))<>1 Then F_C=Replace(F_C,"{$ShowDelivery}" ," style='display:none'")
	
	Dim packstr:packstr=GetPackage(false)          '礼包
	
	DeliveryStr = GetDeliveryTypeStr   '送货方式	
				   
    Set RS=Server.CreateObject("ADODB.RecordSet") 
	'If ProductList<>"" Then
	Dim ProBuyAttr
	'删除没有在当前购物车内的捆绑商品
	 conn.execute("delete from KS_ShopBundleSelect where username='" & getuserid & "' and proid not in(select proid from KS_ShoppingCart where flag=0 and username='" & getuserid & "')")

	'strsql="select I.visitornum,I.MemberNum,I.ID,I.Title,I.ProductType,I.Price_Original,I.Price,I.Price_Member,I.Discount,I.TotalNum,I.GroupPrice,I.photourl,I.unit,I.IsLimitBuy,I.LimitBuyPrice,I.LimitBuyAmount,L.LimitBuyBeginTime,L.LimitBuyEndTime,i.tid,i.fname from KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  where I.ID in ("&ProductList&") order by I.IsChangedBuy,I.ID"
		strsql="select I.visitornum,I.MemberNum,I.ID,I.Title,I.AddDate,i.weight,I.Price,I.Price_Member,I.IsDiscount,I.TotalNum,I.PhotoUrl,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,I.LimitBuyAmount,i.tid,i.fname,L.LimitBuyBeginTime,L.LimitBuyEndTime,C.Attr,C.Amount,C.CartID,C.AttrID from (KS_Product I Inner join KS_ShoppingCart c on i.id=c.proid) Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id   where c.flag=0 and c.username='" & GetUserID & "' order by I.IsChangedBuy,I.ID"

	Call SetBundleSaleAmount()  '设置捆绑促销购买数量
	'ElseIf packstr="" Then
	'Response.Write "<script>alert('您的购物车中没有商品!');history.back();<//script>":response.end
	'End If

	CartStr="<form action=""ShoppingCart.asp"" method=""POST"" name=""check"">"&vbcrlf
	CartStr=CartStr&"      <table border=""0"" cellspacing=""1"" cellpadding=""1"" align=""center"" width=""100%"" class=""border"">" & vbcrlf

	Dim TotalNum:TotalNum=0	
    Dim Price_Member:Price_Member=0
    Dim CurrWeight:CurrWeight=0
	If strSql<>"" Then
			RS.open strsql,conn,1,1
			If RS.Eof And RS.Bof Then
			  If packstr="" Then  RS.CLose:Set RS=Nothing : KS.AlertHIntScript "对不起，您的购物车中没有商品!"
			End If
			Amount = 1
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
			  If KS.ChkClng(RS("isdiscount"))=0 or Discount=0 Then
			    RealPrice=Price_Member
				ProDiscount="无"
			  Else
			   RealPrice=FormatNumber(Price_Member*discount/10,2,-1)
			   ProDiscount=Discount & "折"
			  End If
			  If JFDiscount=0 Then
				ProScore=0
			  ElseIf JFDiscount=1 or KS.ChkClng(rs("isdiscount"))=0 Then
				ProScore=KS.ChkClng(RealPrice) * Amount
			  Else
				ProScore=KS.ChkClng(RealPrice*JFDiscount) * Amount
			  End If
			    if JFDiscount<>0 and JFDiscount<>1 and KS.ChkClng(rs("isdiscount"))=1 then ProDiscount=ProDiscount & " <font color=green>" & JFDiscount & "</font>倍积分"

			Else
			  RealPrice=Price_Member
			End If
			If IsNumerIc(RS("Weight")) Then TotalWeight=TotalWeight+CurrWeight*amount
			
			TotalPrice=TotalPrice+Round(RealPrice*Amount,2)
			TotalScore=TotalScore+ProScore
			Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
			If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
					  CartStr=CartStr&"<table border=""0"" style=""margin-top:5px;margin-bottom:5px;border-bottom:1px solid #ccc""  cellspacing=""1"" cellpadding=""1"" align=""center"" width=""100%"" class=""border""><tr class='tdbg' height=""25"" align=""center""> " & vbcrlf
					  CartStr=CartStr&"  <td align=""left"" class=""carimgbox""><img src=""" & PhotoUrl & """ alt=""" & RS("Title") & """ width=""50"" height=""50"" align=""left"" style=""border:1px solid #ccc;padding:2px""/></td><td align=""left""><div ><a href=""" & KS.GetItemUrl(5,rs("tid"),rs("id"),rs("fname"),rs("adddate")) & """ target=""_blank"">" & RS("Title") 
					  If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
					  CartStr=CartStr& "(换购)"
					  Else
					  CartStr=CartStr& ProBuyAttr
					  End If
					  CartStr=CartStr&"</a> " & RS("Attr")  & "</div>" & vbcrlf
					  CartStr=CartStr&" 数量：" & Amount & "<br/>" & vbcrlf
					  CartStr=CartStr&"	价格：￥" & FormatNumber(Round(RealPrice*Amount,2),2,-1)  & vbcrlf
					  CartStr=CartStr&"</td></tr></table>" & vbcrlf
		              CartStr=CartStr & GetBundleSalePro(RS("ID"),false)   '获得捆绑促销的商品
					  
				 RS.MoveNext
				 Loop
			RS.close:set RS=nothing
    End If

     Dim LMStr:LMStr=packstr
	   
     CartStr=CartStr & LMStr          '礼包

	CartStr=CartStr&"<tr class='tdbg'> " & vbcrlf
	CartStr=CartStr&" <td colspan=""7"" style='padding-left:20px;font-size:14px'><strong>结算信息：</strong>" & vbcrlf
	
	'===========================计算运费===============================
	 Dim ExpressCompany
	 DeliveryMoney=KS.GetFreight(KS.ChkClng(DeliverType),defaultcity,totalweight,ExpressCompany)
	 If DeliveryMoney=-1 Then	DeliveryMoney=0	
	'====================================================================
	
	             
    CartStr=CartStr&" <div style='margin-left:20px'>商品金额：<span id='ordergoodsmoney'>" & Round(TotalPrice,2) & "</span>&nbsp;元 + 运费：<span id='orderyf'>" &DeliveryMoney &"</span>元 + 税费：<span id='ordertax'>0</span>元 <span  id=""ORDER_COSTS""></span>&nbsp;&nbsp;&nbsp;&nbsp;<strong>应付总额：</strong><font color=red>￥<span id='ordertotalmoney'>" & (TotalPrice+DeliveryMoney) & "</span></font>&nbsp;元&nbsp;&nbsp;&nbsp;赠送积分：<span id=""ORDER_SCORE"" style=""color:green"">" & KS.ChkClng(TotalScore) & "</span> 分<input type='hidden' name='TRealTotalPrice' id='TRealTotalPrice' value='" & TotalPrice+DeliveryMoney & "'><input type='hidden' name='RealTotalPrice' id='RealTotalPrice' value='" & TotalPrice+DeliveryMoney & "'><input type='hidden' name='usezf' id='usezf' value='0'></div>" & vbcrlf
    CartStr=CartStr&"  </td>" &vbcrlf
    CartStr=CartStr&"</tr>" & vbcrlf
	
	'使用优惠券
	IF Cbool(KSUser.UserLoginChecked)=true  and LMStr="" and instr(ProBuyAttr,"抢购")=0 Then
		CartStr=CartStr & "<tr class='tdbg'><td colspan='10'><table border=0 cellpadding=4 cellspacing=2 width='90%' align='center'><tr><td nowrap style='text-align:left;width:170px'>" & vbcrlf
	
	 If KS.Setting(181)="1" Then
		Dim MyScore:MyScore=KSUser.GetScore()
		CartStr=CartStr & " <div><strong><label><input onclick=""$('#sss').show();$('#ccc').hide();"" type='radio' name='yhlx' value='1' checked>使用积分抵扣订单金额</label></strong></div>" & vbcrlf  
		CartStr=CartStr & " <div><strong><label><input type='radio' onclick=""$('#sss').hide();$('#ccc').show();"" name='yhlx' value='2'>使用优惠券抵扣订单金额</label></strong></div>"
	Else
	    CartStr=CartStr & " <div><strong><label><input type='hidden' name='yhlx' value='2'>使用优惠券抵扣订单金额</label></strong></div>"
	End If  
	    CartStr=CartStr & "</td></tr><tr><td style='text-align:left;'>"
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
			 	 
		   F_C = Replace(F_C,"{$ShowShoppingCart}",CartStr)
		   If Cbool(KSUser.UserLoginChecked)=False Then 
		   F_C = Replace(F_C,"{$ShowLoginTips}","<strong><font color=ff6600>温馨提示：您还没有注册或登录。享受更多会员优惠，请先<a href=""../login.asp"">登录</a>或<a href=""../../user/reg"" target=""_blank"">注册</a>成为商城会员！</font></strong>")
           Else
		   F_C = Replace(F_C,"{$ShowLoginTips}","亲爱的" & KSUser.UserName &"! 级别："&KS.GetUserGroupName(KSUser.GroupID)&"&nbsp;可用资金：&nbsp;<font color=""green"">" & FormatNumber(KSUser.GetUserInfo("Money"),2,-1,0,0) & "</font>&nbsp;元 " & KS.Setting(45) & "：&nbsp;<font color=green>" & KSUser.GetUserInfo("Point") & "</font>&nbsp;" & KS.Setting(46)&" 积分：&nbsp;<font color=""green"">" & KSUser.GetUserInfo("Score") & "</font>&nbsp;分")
		   End If
		   F_C = ReplaceUserInfo(F_C)
		   F_C=KSRFObj.KSLabelReplaceAll(F_C)
		   Response.Write F_C 
		    
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
	
	Str=Str & "<select name=""couponid"" onchange=""changeCoupon(this.value)"">"
	Str=Str & "<option value='0'>---请选择---</option>"
	If IsArray(SQL) Then
	 For I=0 To  Ubound(SQL,2)
	  Str=Str & "<option value='" & SQL(0,I) & "'>" & SQL(1,I) & "[面值￥" & round(SQL(2,I),2) &"元 余额￥" & formatnumber(SQL(3,I),2,-1) & "元]</option>"
	 Next
	End If
	Str=Str & "</select>"
	GetCouponOptionList=Str
  End Function


  Function  ReplaceUserInfo(F_C)
     F_C=Replace(F_C,"{$ContactMan}",KSUser.GetUserInfo("RealName"))
     F_C=Replace(F_C,"{$Address}",KSUser.GetUserInfo("Address"))
     F_C=Replace(F_C,"{$ZipCode}",KSUser.GetUserInfo("Zip"))
     F_C=Replace(F_C,"{$Phone}",KSUser.GetUserInfo("OfficeTel"))
     F_C=Replace(F_C,"{$Email}",KSUser.GetUserInfo("Email"))
     F_C=Replace(F_C,"{$Mobile}",KSUser.GetUserInfo("Mobile"))
     F_C=Replace(F_C,"{$QQ}",KSUser.GetUserInfo("QQ"))
	 F_C=Replace(F_C,"{$TotalWeight}",TotalWeight)
	 F_C=Replace(F_C,"{$TaxRate}",KS.Setting(65))
	 F_C=Replace(F_C,"{$IncludeTax}",KS.Setting(64))
	 F_C=Replace(F_C,"{$PaymentType}",GetPaymentTypeStr)
	 DeliveryStr=DeliveryStr & "<input type='hidden' id='mustyf' value='" & KS.Setting(180) & "'>"
	 F_C=Replace(F_C,"{$DeliveryType}",DeliveryStr)
	 ReplaceUserInfo=F_C
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
			 defaulttips="首重：<span>"& formatnumber(round(rsss("fweight"),2),2,-1) & "</span> kg 首重价格：<span>" & formatnumber(rsss("carriage"),2,-1) & "</span> 元 续重价格：<span>" & formatnumber(round(rsss("C_fee")/rsss("W_fee"),2),2,-1)&" </span>元/kg"
			end if
		end if
		rsss.close:set rsss=nothing
	   if KS.IsNul(defaultcity) Then defaulttips="选择路线！":defaultcity="选择送货地点" Else tocity=defaultcity
	   Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "select TypeID,TypeName,IsDefault from KS_DeliveryType order by orderid,TypeID",conn,1,1
	   If Not RS.Eof Then
		 SQL=RS.GetRows(-1)
	   End IF
	   RS.Close:Set RS=Nothing
	   GetDeliveryTypeStr="<strong>快递公司：</strong><select onchange=""ajshowdata($('#ccity').innerHTML)"" name='DeliverType' id='DeliverType'>"
	   For I=0 To UBound(SQL,2)
		 If trim(DeliverType)=trim(sql(0,i)) Then
		GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & "</option>"
		 Else
		GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & "</option>"
		End If
	   Next
	   GetDeliveryTypeStr=GetDeliveryTypeStr & "</select>"
	  
	   GetDeliveryTypeStr=GetDeliveryTypeStr & " <br/><input type=""hidden"" value=""" & tocity & """ name=""tocity"" id=""tocity""/> <span><input class=""tocity"" style='text-align;left' name='' id='choosecity' type='button' value='" & defaultcity & "'  onclick=""$('#showprovn').show();if(this.getBoundingClientRect().top>300){showprovn.style.top=(this.offsetHeight-showprovn.offsetHeight)}else{showprovn.style.top='0'};""><span id='showprovn' onclick=""this.style.display='none'"" class='showcity'>"&_
				 "<table width='100%' align='center' border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
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

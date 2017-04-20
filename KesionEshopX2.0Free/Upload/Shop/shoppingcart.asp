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
Set KSCls = New CartCls
KSCls.Kesion()
Set KSCls = Nothing

Class CartCls
        Private KS, KSRFObj,KSUser,DomainStr
		Private ProductList,LoginTF,TotalWeight
		Private TotalPrice,TotalScore,RealPrice,Price_Original,Discount,Amount,attrid,istype,isscore_str,IsScore_z
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
		    istype=KS.ChkClng(ks.g("istype")):IsScore_z=0
		    LoginTF=KSUser.UserLoginChecked
		    ProductList = Session("ProductList")
			GCls.ComeUrl=GCls.GetUrl()
			attrid=KS.ChkClng(Request("attrid"))
		    Dim FileContent,Products,i,RS,strsql,CarListStr	
		 	Products = Split(Replace(KS.S("id")&""," ",""), ",")
		   
		    If Replace(KS.S("cartid")," ","")="" And KS.S("action")="set" Then 
			   ProductList=""
			   Conn.Execute("delete from ks_shoppingcart where flag=0 and username='" & GetUserID & "'")
			ElseIf KS.S("Action")="set" And KS.FilterIds(KS.S("CartID"))<>"" Then
			  Conn.Execute("delete from ks_shoppingcart where flag=0 and username='" & GetUserID & "' and cartid not in(" &KS.FilterIds(KS.S("CartID")) & ")")
			  
			  Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
			  RSS.Open "Select * from KS_ShoppingCart Where flag=0 and UserName='" & GetUserID & "'",conn,1,3
			  Do While Not RSS.Eof
			    Dim RSK
				If KS.ChkClng(RSS("Attrid"))<>0 Then
				 Set RSK=Conn.Execute("Select top 1 a.title,b.amount as TotalNum,a.Unit From KS_Product a inner join KS_ShopSpecificationPrice b ON a.id=B.Proid Where b.ID=" & RSS("Attrid"))
				Else
				 Set RSK=Conn.Execute("Select top 1 title,TotalNum,Unit From KS_Product Where ID=" & RSS("ProID"))
				End If
				If Not RSK.Eof Then
				   If KS.ChkClng(KS.S("Q_"&RSS("CartID")))>RSK(1) Then
					RSS("Amount")=RSK("TotalNum")
					RSS.Update
	                response.write "<script language=javascript>alert('对不起，["&RSK("Title")&"]暂时库存不足，最多只能购买" & RSK("TotalNum") & RSK("unit") & "！');location.href='shoppingcart.asp';</script>" 
				   Else
					RSS.Update
				    RSS("Amount")=KS.ChkClng(KS.S("Q_"&RSS("CartID")))
				   End If
				End IF
				RSK.Close:Set RSK=Nothing
			  RSS.MoveNext
			  Loop
			  RSS.Close:Set RSS=Nothing
			Else 
				'删除大于3天的购物车记录
			    Conn.Execute("Delete From KS_ShoppingCart Where flag=0 and datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>3")
				For I=0 To UBound(Products)
				   If IsNumeric(Products(I)) Then
				   PutToShopBag Products(I), ProductList,I
				   End If
				Next
			End If
			
			'从购物车中删除已在捆绑销售的商品
			Conn.Execute("Delete From KS_ShoppingCart Where Proid in (select pid from KS_ShopBundleSelect where username='" & GetUserId& "') and username='" & GetUserId & "'")
			
			'设置商品没有提交去结算
			Conn.Execute("Update KS_Shoppingcart Set IsPay=0 Where UserName='"& GetUserID &"'")
			
			
			If KS.S("Action")="Del" Then 
			  Call DelProduct(KS.S("ID"))
			ElseIf KS.S("Action")="present" Then
			  AddPresent()
			ElseIf KS.S("Action")="delpack" Then
			  DelPack()    '删除礼包
			ElseIf KS.S("Action")="addBundleSale" Then
		      call addBundleSale()
			ElseIf KS.S("Action")="BundleSale" Then
			  call delBundleSale()
			End If
			
			If Not KS.IsNul(KS.FilterIds(KS.S("Bundid"))) And KS.ChkClng(KS.S("id"))<>0 Then  '判断内容页是否有选择捆绑商品
			  Call addBundleSaleFromNR()
			End If
			
		           If KS.Setting(121)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSRFObj.LoadTemplate(KS.Setting(121))
				   FCls.RefreshType = "ShoppingCart" '设置刷新类型，以便取得当前位置导航等
				    Fcls.RefreshFolderID = "0"        '设置当前刷新目录ID 为"0" 以取得通用标签
					If Trim(FileContent) = "" Then FileContent = "商城购物车模板不存在!"
Set RS=Server.CreateObject("ADODB.RecordSet") 

		 Dim ProBuyAttr
		  '删除没有在当前购物车内的捆绑商品
		   conn.execute("delete from KS_ShopBundleSelect where username='" & getuserid & "' and proid not in(select proid from KS_ShoppingCart where flag=0 and username='" & getuserid & "')")
		   strsql="select I.ID,I.Title,I.FreeShipping,I.WholesaleNum,I.WholesalePrice,I.Price,I.Price_Member,I.VipPrice,I.IsDiscount,I.TotalNum,I.PhotoUrl,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,I.LimitBuyAmount,i.istype,L.LimitBuyBeginTime,L.LimitBuyEndTime,I.MemberNum,I.VisitorNum,I.ArrGroupID,I.Tid,I.Fname,I.AddDate,C.Attr,C.Amount,C.CartID,C.AttrID,C.ProID,i.Score from (KS_Product I Inner join KS_ShoppingCart c on i.id=c.proid) Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id where i.deltf=0 and c.flag=0 and c.username='" & GetUserID & "' order by I.IsChangedBuy,I.ID"

		 If KS.S("Action")="set" then Call SetBundleSaleAmount()  '设置捆绑促销购买数量
		 
RS.open strsql,conn,1,1
	CarListStr="<script>var dir='" & KS.GetDomain &"';</script><script src='../shop/js/shop.detail.js'></script>"&vbcrlf
	If EnabledSubDomain Then
	 CarListStr=CarListStr&"<script>document.domain=""" & RootDomain &""";</script>" &vbcrlf
	end if
	CarListStr=CarListStr&"<form action=""payment.asp"" method=""POST"" name=""check"">"&vbcrlf
	CarListStr=CarListStr&"<input type=""hidden"" name=""istype"" value="""& istype &""">"
	CarListStr=CarListStr&"      <table border=""0"" id=""shoppingtable"" cellspacing=""0"" cellpadding=""0"" align=""center"" width=""100%"" class=""border"">" & vbcrlf
	CarListStr=CarListStr&"            <tr class='title' height=""25"">" & vbcrlf 
	CarListStr=CarListStr&"            <td width=""6%"" align=""center"">编 号</td>" &vbcrlf
	CarListStr=CarListStr&"            <td align=""center"">商 品 名 称</td>" & vbcrlf
	CarListStr=CarListStr&"            <td align=""center"">数 量</td>" & vbcrlf
	CarListStr=CarListStr&"			   <td align=""center"">商城价</td>" & vbcrlf
	CarListStr=CarListStr&"			   <td align=""center"">优惠</td>"&vbcrlf
	CarListStr=CarListStr&"			   <td align=""center"">您的价格</td>" & vbcrlf
	CarListStr=CarListStr&"			   <td align=""center"">总 计</td>" & vbcrlf
	CarListStr=CarListStr&"			   <td style=""display:none"" align=""center"">赠送积分</td>" & vbcrlf
	CarListStr=CarListStr&"			   <td align=""center"">操作</td>" & vbcrlf
	CarListStr=CarListStr&"          </tr>"&vbcrlf

Dim TotalNum:TotalNum=0	
Dim Price_Member:Price_Member=0
If Not RS.Eof Then
Amount = 1
Do While Not RS.EOF
	 if KS.ChkClng(RS("Istype"))<>KS.ChkClng(ks.g("istype")) then 
	    if KS.IsNul(Request.ServerVariables("HTTP_REFERER")) Then
	 	 response.write "<script language=javascript>alert('对不起，普通商品不能和积分商品放在一起');location.href='"& Request.ServerVariables("HTTP_REFERER") &"';</script>"
		Else
	 	 response.write "<script language=javascript>alert('对不起，普通商品不能和积分商品放在一起');location.href='index.asp';</script>"
		End If
	 end if 
	 If Not KS.IsNul(RS("ArrGroupID")) Then
	   If KS.FoundInArr(RS("ArrGroupID"),KSUser.GetUserInfo("GroupID"),",")=false Then
	     Conn.Execute("Delete From KS_ShoppingCart Where flag=0 and Proid=" & RS("id") & " And username='" & GetUserID & "'")
	     response.write "<script language=javascript>alert('对不起，您的用户级别不能购买商品“"&RS("Title")&"”！');history.back(-1);</script>" 
	     response.End()
	   End If
	 End If
	 If RS("AttrID")<>0 Then 
	  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RS("AttrID"))
	  If Not RSAttr.Eof Then
	    TotalNum = RSAttr("amount")
		Price_Member=RSAttr("Price")
	  Else
	    TotalNum = RS("TotalNum")
		Price_Member=RS("Price_Member")
	  End If
	  RSAttr.CLose:Set RSAttr=Nothing
	 Else
	    TotalNum = RS("TotalNum")
		Price_Member=RS("Price_Member")
	 End If
     Amount = rs("amount")
     If Amount <= 0  Then Amount = 1
IF KS.ChkCLng(TotalNum) < KS.ChkClng(Amount) Then
	Amount = 1
	response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]暂时库存不足，最多只能购买" & TotalNum & RS("unit") & "！');history.back(-1);</script>" 
	response.End()
End IF

Call CheckProductNum(RS)

Dim ProDiscount:ProDiscount=""
Dim ProScore:ProScore=""
Dim SingleScore:SingleScore=0
IF RS("IsLimitBuy")<>"0" and RS("LimitBuyAmount") < Amount Then
	Amount = 1
	Session("Amount"&RS("ID")) = 1
	If RS("LimitBuyAmount")=0 Then
	Conn.Execute("Update KS_Product Set IsLimitBuy=0 Where ID=" & RS("ID"))
	response.write "<script language=javascript>alert('对不起,["&RS("Title")&"]已被抢购完,价格已恢复!');location.href='ShoppingCart.asp';</script>" 
	Else
	response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]还剩" & RS("LimitBuyAmount") & RS("unit") & "供抢购！');history.back(-1);</script>" 
	rs.close:set rs=nothing
	response.End()
	End If
End If
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
ElseIF Cbool(LoginTF)=true Then
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
   RealPrice=Price_Member*discount/10
   ProDiscount=Discount & "折"
  End If
  If JFDiscount=0 Then
	SingleScore=0
  ElseIf JFDiscount=1 Then
    SingleScore=KS.ChkClng(RealPrice)
  Else
    SingleScore=RealPrice*JFDiscount'bug改动
  End If
	ProScore=SingleScore * Amount
  if JFDiscount<>0 and JFDiscount<>1 and KS.ChkClng(rs("isdiscount"))=1 then ProDiscount=ProDiscount & " <font color=green>" & JFDiscount & "</font>倍积分"
Else
  RealPrice=Price_Member
End If

if ks.chkclng(rs("WholesaleNum"))<>0 and amount>=ks.chkclng(rs("WholesaleNum")) then realPrice=rs("WholesalePrice")
			
			dim isscore_str1,isscore_is,Score:Score=KS.ChkClng(rs("Score"))
			dim totalscore_str:totalscore_str="" '小计
			isscore_str="":isscore_is=0
			  'if istype=1 then
			  	if Score<>0 then
					if KSUser.GetScore() >= Score*Amount  then
						isscore_str="<b><font color=""#006600""><span  id='realscore"&rs("cartid") &"'>" & Score & "</span>积分 "
						totalscore_str="<b><font color=""#006600""><span  id='realscore"&rs("cartid") &"'>" & Score*Amount & "</span>积分 "
						if RealPrice > 0 then
							isscore_str=isscore_str & " + ￥<span  id='realmember"&rs("cartid") &"'>"& KS.GetPrice(RealPrice) &"</span>元</font></b>"
							totalscore_str=totalscore_str & " + ￥<span name='totalmyprice' id='realmember"&rs("cartid") &"'>"& KS.GetPrice(RealPrice*Amount) &"</span>元</font></b>"
						end if
						'RealPrice=0
						isscore_is=1
						IsScore_z=IsScore_z+(Score*Amount)
					else
						isscore_str="<b><font color=""#FF0000"">(积分不够)</font></b>"	
						isscore_str1= "<script>alert('积分不够!');</script>"
						isscore_is=0
					end if
				end if
			  'else
			  '	istype=0	
			  'end if


TotalPrice=TotalPrice+Round(RealPrice*Amount,2)
TotalScore=TotalScore+KS.ChkClng(ProScore)
Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
				
          CarListStr=CarListStr&isscore_str1&"<tr class='tdbg' height=""25"" align=""center""> " & vbcrlf
          CarListStr=CarListStr&"  <td><input type='checkbox' value='" & RS("cartid") & "' name='id' checked/></td>" & vbcrlf
          CarListStr=CarListStr&"  <td align=""left"" valign=""top""><img src=""" & PhotoUrl & """ alt=""" & RS("Title") & """ width=""50"" height=""50"" align=""left"" style=""border:1px solid #ccc;padding:2px;margin-right:4px""/><a href=""" & KS.GetItemUrl(5,rs("tid"),rs("id"),rs("fname"),rs("AddDate")) & """ target=""_blank"">" & RS("Title")
		  If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
		  CarListStr=CarListStr& "<span style='color:#ff6600'>(换购)</span>"
		  Else
		  CarListStr=CarListStr& ProBuyAttr
		  End If
		  if (amount>=ks.chkclng(rs("FreeShipping")) and ks.chkclng(rs("FreeShipping"))<>0) then CarListStr=CarListStr &" <span style='color:red'>(免邮)</span>"
		  
		  
		  CarListStr=CarListStr&"</a><br/>"
		  CarListStr=CarListStr& rs("attr")
		  CarListStr=CarListStr & "</td>" & vbcrlf
		  If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
          CarListStr=CarListStr&"  <td><input type=""hidden""   name=""Q_" & RS("cartid") & """ value=""1"" size=""5"" style=""text-align:center"" class=""textbox"" readonly> 1</td>" & vbcrlf
		  Else
          CarListStr=CarListStr&"  <td><a href='javascript:;' onclick='shop.buynums(0," & rs("proid") & "," & rs("attrid") & "," & RS("cartid") & ",0,0,"&rs("Score")&")'><img border='0' src='../images/default/ico_close.gif'/></a> <input onchange=""shop.buynums(0," & rs("proid") & "," & rs("attrid") & "," & RS("cartid") & ",this.value,-1)"" type=""Text"" name=""Q_" & RS("cartid") & """ id=""Q_" & RS("cartid") & """ value=""" & Amount & """ size=""4"" style=""text-align:center"" class=""textbox""> <a href='javascript:;' onclick='shop.buynums(0," & rs("proid") & "," & rs("attrid") & "," & RS("cartid") & ",0,1)'><img border='0' src='../images/default/ico_open.gif'/></a> </td>" & vbcrlf
		  End If
	   	  CarListStr=CarListStr&"	<td>￥" & KS.GetPrice(Price_Member) & "<span id='hidmyprice_Member" & rs("cartid") &"' style='display:none'>"& RealPrice  & "</span></td>" & vbcrlf
		  If Trim(RS("ID"))=trim(Session("ChangeBuyID")) Then
		  CarListStr=CarListStr&"	<td align='center'>换购</td>" & vbcrlf
		  ElseIf ProDiscount<>"" Then
		  CarListStr=CarListStr&"	<td align='center'>" & ProDiscount & "</td>" & vbcrlf
		  Else
		  CarListStr=CarListStr&"	<td align='center'>无优惠</td>" & vbcrlf
		  End If
		  
		   if isscore_str="" then
		   		CarListStr=CarListStr&"	<td style='color:#ff3300;font-size:14px;font-weight:bold'>￥<span id='realPrice"&rs("cartid") &"'>" &KS.GetPrice(RealPrice) & "</span></td>" & vbcrlf
		   else
			   CarListStr=CarListStr&"	<td style='color:#ff3300;font-size:14px;font-weight:bold'><span>" & Score & "</span>积分"
			   if RealPrice > 0 then
					CarListStr=CarListStr& " + ￥"& KS.GetPrice(RealPrice) &"元"
			   end if
			   CarListStr=CarListStr& "</td>" & vbcrlf
		   end if
		   
		 CarListStr=CarListStr&"<td><input type='hidden' id='hidWholesalePrice" & rs("cartid") &"' value='" & rs("WholesalePrice") & "'/><input type='hidden' id='hidWholesaleNum" & rs("cartid") &"' value='" & rs("WholesaleNum") & "'/><span style='display:none' id='hidisscore" & rs("cartid") &"'>" & Score & "</span>" 
		 if isscore_str="" or isscore_is=0 then
			 CarListStr=CarListStr & isscore_str &"￥<span name='totalmyprice' id='myprice" & rs("cartid") &"'>"& KS.GetPrice(RealPrice*Amount) & "</span><span id='hidmyprice" & rs("cartid") &"' style='display:none'>"& KS.GetPrice(RealPrice)  & "</span>"
		else
		 	 CarListStr=CarListStr & totalscore_str
		end if
		 
		CarListStr=CarListStr& "</td>" & vbcrlf
		  CarListStr=CarListStr&"	<td style=""display:none""><span style='display:none' id='hidmyscore" & rs("cartid") &"'>" & SingleScore & "</span><span name='totalmyscore' id='myscore" & rs("cartid") &"'>" & KS.ChkClng(ProScore) & "</span> 分</td>" & vbcrlf
		  CarListStr=CarListStr&"	<td><a href=""?Action=Del&ID=" & RS("CartID") & ""
		  if istype=1 then
		  	CarListStr=CarListStr&"&istype=1"
		  end if 
		  CarListStr=CarListStr&""">删除</a> <a href=""../User/index.asp?User_Favorite.asp?Action=Add&ChannelID=5&InfoID=" & RS("ID") & """ target=""_blank"">收藏</a></td>" & vbcrlf
          CarListStr=CarListStr&"</tr>" & vbcrlf
		  CarListStr=CarListStr & GetBundleSalePro(RS("ID"),true)   '获得捆绑促销的商品
     RS.MoveNext
     Loop
	 
	    CarListStr=CarListStr& GetPackage(true)         '礼包
	Else 
	    dim packstr:packstr=GetPackage(true)           '礼包
		if packstr="" then
		  CarListStr=CarListStr&"	<tr class='tdbg'><td colspan=8>您的购物车没有商品!</td></tr>" & vbcrlf
		else
		  CarListStr=CarListStr& packstr
		end if
	End If
	RS.close
	
	

	CarListStr=CarListStr&"<tr class='tdbg'> " & vbcrlf
	CarListStr=CarListStr&" <td colspan=""6"" style='text-align:right' nowrap>"
If KS.ChkClng(KS.Setting(63))=0 And Cbool(LoginTF)=false Then 
     CarListStr=CarListStr&" <img src=""images/shop_btn2.gif"" onclick=""ShowLogin();"" / style=""cursor:hand"">&nbsp; " & vbcrlf
    Else
     CarListStr=CarListStr&" <input type=""image"" name=""payment"" src=""images/shop_btn2.gif"" onclick=""this.form.action='PayMent.asp';this.form.target='_parent';"">&nbsp; " & vbcrlf
	End If
	Dim ComeUrl:ComeUrl=request.ServerVariables("http_referer")
	If ComeUrl="" or instr(lcase(comeurl&""),"shoppingcart.asp")<>0 Then ComeUrl="../index.asp"
    CarListStr=CarListStr&" <a href='" &comeurl & "'><img  src=""images/shop_btn3.gif"" border=0></a>" & vbcrlf
    CarListStr=CarListStr&"  </td>" &vbcrlf	
	CarListStr=CarListStr&"<td colspan=""3"" height=""25"">合计：<font color=""#FF6600"">￥<span id='totalprice'>" & FormatNumber(TotalPrice,2,-1) & "</span></font>&nbsp;元,可得积分：<font color=green><span id='totalscore'>" & KS.ChkClng(TotalScore) & "</span></font> 分</td>" & vbcrlf
    CarListStr=CarListStr&"    </tr>" & vbcrlf
    CarListStr=CarListStr&"</table>" & vbcrlf
    CarListStr=CarListStr&"</form>"&vbcrlf
	'检查换购品合法性
	Call CheckChangeBuy(TotalPrice)
			 	 
	Dim ShowChangeBuy:ShowChangeBuy=true
	  RS.Open "Select * From KS_ShopBundleSale Where ProID in(select proid from KS_ShoppingCart where flag=0 and username='" & getuserid & "')",conn,1,1
	  If Not RS.Eof Then
	   ShowChangeBuy=false
	  End If
	  RS.Close
	
	If ShowChangeBuy=false Then       '不允许显示换购时,显示捆绑销售的商品
	  Dim ProID:ProID=KS.ChkClng(Request("id"))
	  If ProID=0 Then  Proid=Conn.Execute("Select top 1 ProID From KS_ShopBundleSale Where ProID in(select proid from ks_shoppingcart where username='" & GetUserID &"')")(0)  '如果没有传商品ID过来,随机找一条有捆绑销售的产品
	     RS.Open "Select I.Tid,I.Fname,I.AddDate,I.ID,I.Title,I.Price,I.PhotoUrl,B.KBPrice,B.Proid From KS_Product I Inner Join KS_ShopBundleSale b on i.id=b.kbproid Where b.kbproid not in(select pid from KS_ShopBundleSelect where username='" & getuserid &"') and B.proid=" & proid,conn,1,1
		 If Not RS.Eof Then
		    Set GXML=KS.RsToXml(RS,"row","")
		 End If
		 RS.Close
		 If IsObject(GXML) Then
		    CarListStr=CarListStr&"<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF"" class=""carttr"">" &vbcrlf
		    CarListStr=CarListStr&"<tr><td style='border-bottom:1px solid #cccccc;padding-bottom:4px;color:green;font-size:14px'><strong>&nbsp;<img src='../images/default/arrow_w.gif' align='absmiddle' /> 您可能还需要以下商品(捆绑促销)</strong></td></tr>" &vbcrlf
		    CarListStr=CarListStr&"<tr><td>" &vbcrlf
			CarListStr=CarListStr&"<div class='kblist'><ul>"
			 For Each Node In GXML.DocumentElement.SelectNodes("row")
					PhotoUrl=Node.SelectSingleNode("@photourl").text
					If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
			        CarListStr=CarListStr&"<li><table border='0'><tr><td style='width:90px'><a href='" & KS.GetItemUrl(5,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text,Node.SelectSingleNode("@adddate").text) & "' target='_blank'><img src='" & PhotoUrl & "' border='0' width='80' height='80' title='" & Node.SelectSingleNode("@title").text & "'></a></td><td style='line-height:26px'><a href=""" & KS.GetItemUrl(5,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text,Node.SelectSingleNode("@adddate").text) & """ target=""_blank"" class=""t"">" & Node.SelectSingleNode("@title").text & "</a><br/><span style='color:#ff6600'>仅需￥" & Node.SelectSingleNode("@kbprice").text & "元</span><br/><a href='?action=addBundleSale&pid=" & Node.SelectSingleNode("@id").text & "&proid=" & node.selectsinglenode("@proid").text & "&id=" & ks.chkclng(ks.s("id")) & "'><img src='images/addcart.gif' border='0'></a></td></tr></table></li>"
			 Next
			CarListStr=CarListStr&"</ul></div>" &vbcrlf
		    CarListStr=CarListStr&"</td></tr>" &vbcrlf
		    CarListStr=CarListStr&"</table>" &vbcrlf
		 End If
	  
	Else
			'换购商品
			Dim XML,Node,Param,GXML,GNode
			RS.Open "select ChangeBuyNeedPrice,ChangeBuyPresentPrice from ks_product  where IsChangedBuy=1 group by ChangeBuyPresentPrice,ChangeBuyNeedPrice",Conn,1,1
			 If Not RS.Eof Then
			  Set GXML=KS.RsToXml(RS,"row","")
			 End If
			 RS.Close
			 If IsObject(GXML) Then
			  Dim oldchangebuypresentprice,clickStr
			  CarListStr=CarListStr&"<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF"" class=""carttr"">" &vbcrlf
			  For Each GNode In GXML.DocumentElement.SelectNodes("row")
					  CarListStr=CarListStr&"<tr><td align=""left"">" & vbcrlf
					  If GNode.SelectSingleNode("@changebuypresentprice").text=0 Then 
							CarListStr=CarListStr&"<div style=""width:100%; text-align:left; background:#eeeeee""><div style=""background:#000000; float:left; height:25px; padding:0 20px; line-height:25px"">　<b style=""font-size:14px; color:#FFFFFF"">免费领取</b></div><div style=""background:#eeeeee; float:left; height:25px; padding:0 20px; line-height:25px""><font color=""#000000""><b>满<font color=""#ff0000"">￥" & GNode.SelectSingleNode("@changebuyneedprice").text & "</font>元,可免费领取以下礼品</b>"&vbcrlf
						   If round(TotalPrice,2)<round(GNode.SelectSingleNode("@changebuyneedprice").text,2) Then
						   CarListStr=CarListStr&"[您只需要再购买<font color=red>￥" &GNode.SelectSingleNode("@changebuyneedprice").text-TotalPrice & "</font>元的商品,就可免费领取赠品]"
						   End If
					   Else
							CarListStr=CarListStr&"<div style=""width:100%; text-align:left; background:#eeeeee""><div style=""background:#000000; float:left; height:25px; padding:0 20px; line-height:25px"">　<b style=""font-size:14px; color:#FFFFFF"">换购商品</b></div><div style=""background:#eeeeee; float:left; height:25px; padding:0 20px; line-height:25px""><font color=""#000000""><b>满<font color=""#ff0000"">￥" & GNode.SelectSingleNode("@changebuyneedprice").text & "</font>元,加<font color=""#ff0000"">￥" & GNode.SelectSingleNode("@changebuypresentprice").text & "</font>元,可换购以下商品</b>"&vbcrlf
							
						   If round(TotalPrice,2)<round(GNode.SelectSingleNode("@changebuyneedprice").text,2) Then
						   CarListStr=CarListStr&"[您只需要再购买<font color=red>￥" &GNode.SelectSingleNode("@changebuyneedprice").text-TotalPrice & "</font>元的商品,就可选购本级别的赠品]"
						   End If
			
					   End If
					 CarListStr=CarListStr&"</font></div></td></tr>" &vbcrlf  
					 CarListStr=CarListStr&"<tr><td>" &vbcrlf 		   
					 CarListStr=CarListStr&"<div class='prolist'><ul>" &vbcrlf 	
					 
					'If Session("ProductList")<>"" Then
					' Param=" and id not in(" & Session("ProductList") & ")"
					'End If
					RS.Open "Select top 50 ID,Title,ChangeBuyNeedPrice,ChangeBuyPresentPrice,Price,PhotoUrl From KS_Product Where IsChangedBuy=1 And verific=1 and ChangeBuyPresentPrice=" & GNode.SelectSingleNode("@changebuypresentprice").text & " and ChangeBuyNeedPrice=" & GNode.SelectSingleNode("@changebuyneedprice").text &" and  deltf=0 " & Param & " Order By ChangeBuyPresentPrice,Id Desc"
					If Not RS.Eof Then 
					 Set Xml=KS.RsToXml(RS,"row","")
					End If
					RS.Close
					If IsObject(XML) Then
					  For Each Node In XML.DocumentElement.SelectNodes("row")
							PhotoUrl=Node.SelectSingleNode("@photourl").text
							If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
							CarListStr=CarListStr&"<li><a onclick=""return(shop.checkchangebuy(" & GNode.SelectSingleNode("@changebuyneedprice").text&"))"" href='?action=present&id=" & Node.SelectSingleNode("@id").text & "'><img src='" & PhotoUrl & "' border='0' width='60' height='60' title='" & Node.SelectSingleNode("@title").text & "'></a><br/><a href=""../item/show.asp?m=5&d=" & Node.SelectSingleNode("@id").text & """ target=""_blank"" class=""t"">" & Node.SelectSingleNode("@title").text & "</a><a href='?action=present&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(shop.checkchangebuy(" & GNode.SelectSingleNode("@changebuyneedprice").text&"))""><img src='images/addcart.gif' border='0'/></a></li>"
					  Next
					
					End If
					CarListStr=CarListStr&"</ul></div>"  &vbcrlf 
					CarListStr=CarListStr&"</td></tr>"  &vbcrlf
			  Next	 
			 
			 End If
	 
   End If

	
	 set RS=nothing
	 
	
	  CarListStr=CarListStr&"</table>" & vbcrlf
	
	
		   FileContent = Replace(FileContent,"{$ShowShoppingCart}",CarListStr)
		   If Cbool(KSUser.UserLoginChecked)=False Then 
		   FileContent = Replace(FileContent,"{$ShowLoginTips}","<strong><font color=ff6600>温馨提示：您还没有注册或登录。享受更多会员优惠，请先<a href=""../user/login/"">登录</a>或<a href=""../user/reg"" target=""_blank"">注册</a>成为商城会员！</font></strong>")
           Else
		   FileContent = Replace(FileContent,"{$ShowLoginTips}","亲爱的" & KSUser.UserName &"! 您的个人信息-> 用户组："&KS.GetUserGroupName(KSUser.GroupID)&"&nbsp;可用资金：&nbsp;<font color=""green"">" & KSUser.GetUserInfo("Money") & "</font>&nbsp;元 " & KS.Setting(45) & "：&nbsp;<font color=green>" & KSUser.GetUserInfo("Point") & "</font>&nbsp;" & KS.Setting(46)&" 积分：&nbsp;<font color=""green"">" & KSUser.GetUserInfo("Score") & "</font>&nbsp;分")
		   End If
		   FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		Sub PutToShopBag( Prodid, ProductList ,I)
		   if KS.S("Action")="Del" or KS.S("Action")="addBundleSale" then exit sub
		    dim attr:attr=KS.S("AttributeCart")
			
			if istype=1 then 
			  conn.execute("delete from KS_ShoppingCart where flag=0 and istype<>1 and username='" & GetUserID & "'")
			else
			  conn.execute("delete from KS_ShoppingCart where flag=0 and istype<>0 and username='" & GetUserID & "'")
			end if
			
			
			dim rs:set rs=server.CreateObject("adodb.recordset")
			
			rs.open "select top 1 * from KS_ShoppingCart where flag=0 and attrid=" & attrid & " and username='" & GetUserID & "' And proid=" & KS.ChkClng(Prodid),conn,1,3
			if rs.eof and rs.bof then
			   rs.addnew
			   rs("flag")=0
			   rs("proid")=Prodid
			   rs("attrid")=attrid
			   rs("username")=GetUserID
			   rs("attr")=attr
			   rs("adddate")=now
			   rs("amount")=KS.ChkClng(KS.S( "Q_" & Prodid))
			   if istype=1 then rs("istype")=istype
			   rs.update
			end if
			rs.close
			set rs=nothing
			 
		   if KS.S("action")="set" then
		       Conn.Execute("Delete From KS_ShoppingCart Where cartid=" & KS.ChkCLng(ks.s("cartid")) & " and username='" & GetUserID & "'")
			   If i = 0 Then
				  ProductList =Prodid
			   ElseIf KS.FoundInArr( ProductList, Prodid,",")=false Then
				  ProductList = ProductList&", "&Prodid &""
			   End If
		   else
			   If Len(ProductList) = 0 Then
				  ProductList =Prodid
			   ElseIf KS.FoundInArr( ProductList, Prodid,",")=false Then
				  ProductList = ProductList&", "&Prodid &""
			   End If
		  end if
		  
		  If KS.S("Action")="present" and Session("ChangeBuyID")<>KS.S("ID") Then
		   Call DelProduct(Session("ChangeBuyID"))
		  End If
      End Sub
	  Sub DelProduct(DelID)
	  If DelID<>"" Then
	   	 Conn.Execute("Delete From KS_ShoppingCart where cartid=" & KS.ChkClng(DelID))
	     Conn.Execute("Delete From KS_ShopBundleSelect Where UserName='" & getuserid & "' and pid=" & DelID)
	   End If
	   Dim i,Parr:Parr=Split(ProductList,",")
	   Dim NewPList
	   For i=0 To Ubound(Parr)
	    If trim(Parr(i))<>trim(DelID) Then
		 If NewPlist="" Then
		  NewPlist=Parr(i)
		 Else
		  NewPlist=NewPlist & "," & Parr(I)
		 End If
		End If
	   Next
	   ProductList=NewPlist
	  End Sub
	  
	  Sub AddPresent()
	   If KS.S("ID")="" Then KS.AlertHintScript "对不起,您没有选择换购品!"
	   If Session("ChangeBuyID")<>"" Then
	     If Session("ChangeBuyID")<>KS.S("ID") and ks.s("f")="" Then
		   KS.Die "<script>if (confirm('您已选过换购品了,是否替换?')){location.href='?f=ok&action=present&id=" & KS.S("id") &"'}else{location.href='shoppingcart.asp'}</script>"
		 End If
	   End If
	   
	   Dim RS:Set RS=Conn.Execute("Select top 1 ChangeBuyNeedPrice,ChangeBuyPresentPrice From KS_Product Where ID=" & KS.ChkClng(KS.S("ID")))
	   IF Not RS.Eof Then
	     Session("ChangeBuyID")=KS.FilterIds(KS.S("ID"))
		 Session("ChangeBuyNeedPrice")=RS(0)
		 Session("ChangeBuyPrice")=RS(1)
	   End If
	   RS.Close:Set RS=Nothing
	  End Sub
	  
	  '检查换购品合法性
	  Sub  CheckChangeBuy(TotalPrice)
       If KS.IsNul(Session("ChangeBuyID")) Or KS.IsNul(Session("ChangeBuyNeedPrice")) Then Exit Sub

	      If Round(TotalPrice,2)>=Round(Session("ChangeBuyNeedPrice"),2) Then
			Exit Sub
		  Else
		    Call DelProduct(Session("ChangeBuyID"))
			Session("ChangeBuyID")=""
			Session("ProductList")=KS.FilterIds(ProductList)
			if istype=1 then Response.Redirect "ShoppingCart.asp?istype=1" else Response.Redirect "ShoppingCart.asp"
			
			'Call KS.Alert("对不起,您的订单金额不够换购此商品!","ShoppingCart.asp")
		  End If
	  End Sub
	  
	  
	  '删除礼包
	  Sub DelPack()
	    Dim ID:ID=KS.ChkClng(Request("id"))
		If ID<>0 Then
		Conn.Execute("Delete From KS_ShopPackageSelect Where PackID=" & ID & " and username='" & GetUserID &"'")
		End If
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
	  End Sub
		
	  Sub addBundleSale()   '购物车页添加捆绑
	     Dim Pid:Pid=KS.ChkCLng(Request("pid"))        '选购的商品ID
		 Dim ProID:ProID=KS.ChkCLng(request("proid"))  '绑定的商品ID
		 If ProID<>0 and pid<>0 and ProductList<>"" then
			Call  addBundle(Pid,ProID)
		 End If
	  End Sub
	  Sub addBundleSaleFromNR() '内容页添加捆绑
	     Dim Bundid:Bundid=KS.FilterIds(KS.S("Bundid"))
		 Dim I,BundidArr
		 If Not KS.IsNul(Bundid) Then
		    BundidArr=Split(Bundid,",")
			For I=0 To Ubound(BundidArr)
			 Call addBundle(BundidArr(i),KS.ChkClng(KS.S("id")))
			Next
		 End If
	  End Sub
	  Sub addBundle(Pid,ProID)
	  	  '删除超过5天的记录
		  Conn.Execute("Delete From KS_ShopBundleSelect Where datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>5")
		  Dim Price
		  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		  RS.Open "Select top 1 KBPrice From KS_ShopBundleSale Where ProID=" & ProID & " And KBProID=" & Pid,conn,1,1
			 If RS.Eof Then
			    RS.Close :Set RS=Nothing
				Exit Sub
			 Else
			    Price=RS(0)
			 End If
		 rs.close
         RS.Open "Select top 1 * From KS_ShopBundleSelect where username='" & GetUserID & "' and pid=" & pid & " and proid=" & proid,conn,1,3
			 If RS.Eof Then
				RS.AddNew
				RS("UserName")=GetUserID
				RS("Pid")=Pid
				RS("ProID")=ProID
				RS("Amount")=1
				RS("AddDate")=Now
				RS("Price")=Price
				RS.Update
			 End If
			 RS.Close : Set RS=Nothing
	  End Sub
	  
	  Sub delBundleSale()
	    Dim SelID:SelID=KS.ChkClng(Request("SelID"))
		If SelID=0 Then Exit Sub
		Conn.Execute("Delete From KS_ShopBundleSelect Where ID=" & SelID)
	  End Sub
	  
	  
End Class
%>

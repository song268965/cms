<%
 Dim ProIDs
 '得到超值礼包
 '参数showdel  true---显示删除  false ---不显示
 Function GetPackage(showdel)
	    If KS.IsNul(GetUserID) Then Exit Function
		Dim RS,RSB,GXML,GNode,str,n,Price,CurrWeight
		Set RS=Conn.Execute("select packid,username from KS_ShopPackageSelect Where UserName='" & GetUserID & "' group by packid,username")
		If Not RS.Eof Then
		 Set GXML=KS.RsToXml(Rs,"row","")
		End If
		RS.Close : Set RS=Nothing
		If IsOBJECT(GXml) Then
		   FOR 	Each GNode In GXML.DocumentElement.SelectNodes("row")
		     Set RSB=Conn.Execute("Select top 1 * From KS_ShopPackAge Where ID=" & GNode.SelectSingleNode("@packid").text)
			 If Not RSB.Eof Then
					  If Conn.Execute("Select Sum(Amount) From KS_ShopPackageSelect Where Packid=" & GNode.SelectSingleNode("@packid").text & " and username='" & GetUserID & "'")(0)=RSB("num") or rsb("PackType")=1 Then   '商品件数是否一致
					  Dim PhotoUrl:PhotoUrl=RSB("PhotoUrl")
                      If KS.IsNul(PhotoUrl) Then PhotoUrl="../../images/nopic.gif"

						Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
						RSS.Open "Select a.title,a.weight,a.Price_Member,a.Price,b.* From KS_Product A inner join KS_ShopPackageSelect b on a.id=b.proid Where b.packid=" & GNode.SelectSingleNode("@packid").text & " and  b.UserName='" & GetUserID & "'",Conn,1,1
						  str=str & "<table border=""0"" style=""margin-top:5px;margin-bottom:5px;border-bottom:1px solid #ccc""  cellspacing=""1"" cellpadding=""1"" align=""center"" width=""100%"" class=""border""><tr class='tdbg' height=""25"" style='text-align:center'><td align=""center"" width=""80"" valign=""top""><img src=""" & PhotoUrl & """  width=""50"" height=""50"" style=""border:1px solid #ccc;padding:2px;margin-right:4px""/>"
						  if showdel then
						  str=str & "<br/><a href='?action=delpack&id=" & rsb("id") & "' onclick=""return(confirm('确定删除该礼包吗?'))"">删除</a>"
						  end if
						  str=str &" </td><td style='text-align:left'><strong>礼包【<a href='../../shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a>】您选择的套装详细如下:</strong>"
						  n=1
						  Dim TotalPackPrice,tempstr,i
						  TotalPackPrice=0 : tempstr=""
						Do While Not RSS.Eof
						 
						  For I=1 To RSS("Amount") 
							  '得到单件品价格
							  If RSS("AttrID")<>0 Then 
							  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RSS("AttrID"))
							  If Not RSAttr.Eof Then
								Price=RSAttr("Price")
								CurrWeight=RSAttr("Weight")
							  Else
								Price=RSS("Price_member")
								CurrWeight=RSS("Weight")
							  End If
							  RSAttr.CLose:Set RSAttr=Nothing
							 Else
								Price=RSS("Price_member")
								CurrWeight=RSS("Weight")
							 End If
							   	TotalWeight=TotalWeight+CurrWeight*RSS("Amount") 
							   TotalPackPrice=TotalPackPrice+Price
							  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
							  n=n+1
						  Next
						  RSS.MoveNext
						Loop
						str=str & "<br/>" & tempstr & "" 
						str=str &"数量：1<br/>折扣：" & rsb("discount") & "折<br/>价格：<span style='color:#ff3300;font-size:14px;font-weight:bold'>￥<span name='totalmyprice'>" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</span>"
						str=str & "</td></tr></table>" 
						
						TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1),2)   '将礼包金额加入总价
						RSS.Close
						Set RSS=Nothing
					
					End If
			End If
			RSB.Close
		   Next
			
	    End If
		
		
		
		GetPackage=str
		
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

Function GetBundleSalePro(ProID,ShowAction)   '获得捆绑促销的商品
   If KS.FoundInArr(ProIDs,ProID,",")=true Then Exit Function
  ProIDs=ProIDS &"," & ProID
  Dim Str,RS,XML,Node
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "Select I.ID,I.Title,i.weight,b.Price,b.amount,b.id as selid From KS_Product I inner Join KS_ShopBundleSelect b on i.id=b.pid Where B.ProID=" & ProID & " and b.username='" & GetUserID & "' order by I.id",conn,1,1
  If Not RS.Eof Then
    Set XML=KS.RsToXml(rs,"row","")
  End If
  RS.Close:Set RS=Nothing
  If IsObject(XML) Then
	  Str=Str&"<tr class='tdbg' height=""15"" align=""center""> " & vbcrlf
	  Str=Str&"<td></td><td colspan=10 style='text-align:left;color:green;'>选购捆绑促销商品:</td> " & vbcrlf
	  Str=Str&"</tr> " & vbcrlf
     For Each Node In Xml.DocumentElement.SelectNodes("row")
	  Str=Str&"<tr class='tdbg'  align=""center""> " & vbcrlf
	  Str=Str&"<td height=""10""></td>" &vbcrlf
	  str=str &"<td style='text-align:left'><img src='../images/default/arrow_w.gif' align='absmiddle' /> " & Node.SelectSingleNode("@title").text &"</td> " & vbcrlf
	  str=str &"<td><input type='hidden' name='SelID' value='" & Node.SelectSingleNode("@selid").text & "'>"
	  If ShowAction=true Then
	  str=str & "<a href='javascript:;' onclick='shop.buynums(1," & proid & ",0," & Node.SelectSingleNode("@selid").text & ",0,0)'><img border='0' src='../images/default/ico_close.gif'/></a> <input type=""Text"" name=""Q_" & Node.SelectSingleNode("@selid").text & """ id=""Q_" & Node.SelectSingleNode("@selid").text & """ value=""" & Node.SelectSingleNode("@amount").text & """ size=""4"" style=""text-align:center"" class=""textbox""> <a href='javascript:;' onclick='shop.buynums(1," & proid & ",0," & Node.SelectSingleNode("@selid").text & ",0,1)'><img border='0' src='../images/default/ico_open.gif'/></a>"
	  else
	   str=str & Node.SelectSingleNode("@amount").text
	  end if
	  str=str & "</td> " & vbcrlf
	  str=str &"<td>---</td> " & vbcrlf
	  str=str &"<td>---</td> " & vbcrlf
	  str=str &"<td style='color:#ff3300;font-size:14px;font-weight:bold'>￥" & formatnumber(Node.SelectSingleNode("@price").text,2,-1) &"</span></td> " & vbcrlf
	  
	  str=str&"	<td>￥<span name='totalmyprice' id='myprice" & Node.SelectSingleNode("@selid").text &"'>" & formatnumber(Node.SelectSingleNode("@price").text*Node.SelectSingleNode("@amount").text,2,-1)  & "</span><span id='hidmyprice" & Node.SelectSingleNode("@selid").text &"' style='display:none'>" & Node.SelectSingleNode("@price").text  & "</span></td>" & vbcrlf

	  str=str &"<td><span style='display:none' id='hidmyscore" & Node.SelectSingleNode("@selid").text &"'>0</span><span name='totalmyscore' id='myscore" & Node.SelectSingleNode("@selid").text &"'>0</span></td> " & vbcrlf
	  If ShowAction=true Then
	  str=str &"<td><a href=""?Action=BundleSale&SelID=" & Node.SelectSingleNode("@selid").text & """>删除</a> <a href=""../User/index.asp?User_Favorite.asp?Action=Add&ChannelID=5&InfoID=" & Node.SelectSingleNode("@id").text & """ target=""_blank"">收藏</a></td> " & vbcrlf
	  End If
	  Str=Str&"</tr> " & vbcrlf
	  TotalWeight=TotalWeight+Node.SelectSingleNode("@weight").text*Node.SelectSingleNode("@amount").text 
	  TotalPrice=TotalPrice+round(Node.SelectSingleNode("@price").text*Node.SelectSingleNode("@amount").text,2)   '将礼包金额加入总价

	 Next
  End If
  
  GetBundleSalePro=Str
End Function

Sub SetBundleSaleAmount()
	   Dim SelID,IdArr,I,Num
	   SelID=KS.S("SelID")
	   If SelID<>"" Then
	     SelID=Replace(SelID," ","")
		 IdArr=Split(SelID,",")
		 For I=0 To Ubound(IDArr) 
		   Num=KS.ChkClng(Request("Q_"&idArr(i)))
		   If Num>0 Then
		     Conn.Execute("Update KS_ShopBundleSelect Set Amount=" & Num & " Where ID=" & IDArr(i))
		   End If
		 Next
	   End If
End Sub
Sub CheckProductNum(RS)
Dim HasMemberBuyNum,HasVisitorBuyNum,MemberNum,visitornum,NowNum
MemberNum=KS.ChkClng(RS("MemberNum"))
visitornum=KS.ChkClng(RS("VisitorNum"))
If MemberNum<>0 Then
   HasMemberBuyNum=KS.ChkClng(Conn.Execute("select sum(Amount) From KS_OrderItem Where ProID=" & RS("ID") & " AND IsMember=1 and datediff(" & DataPart_D & ",begindate," & SQLNowString & ")<=0")(0))
End If
If visitornum<>0 Then
   HasVisitorBuyNum=KS.ChkClng(Conn.Execute("select sum(Amount) From KS_OrderItem Where ProID=" & RS("ID") & " AND IsMember=0 and datediff(" & DataPart_D & ",begindate," & SQLNowString & ")<=0")(0))
End If
If KS.C("UserName")<>"" And MemberNum<>0 Then
    If HasVisitorBuyNum<>0 Then 
	 KS.Die "<script>alert('对不起，您已用游客身份购买过["&RS("Title")&"]了，请明天再购买!');history.back(-1)</script>"
	End If
	If HasMemberBuyNum>=MemberNum Then
		Call DelProduct(RS("ID"))
		Session("ProductList")=ProductList
		response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]限制每位会员每天只能购买" & MemberNum & RS("unit") & "！');history.back(-1);</script>" 
		response.End()
	ElseIf HasMemberBuyNum+KS.ChkClng(Session("Amount"&RS("ID")))>MemberNum Then
		Session("Amount"&RS("ID")) = MemberNum-HasMemberBuyNum
		response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]限制每位会员每天只能购买" & MemberNum & RS("unit") & ",您只能再购买" & Session("Amount"&RS("ID")) &  RS("UNIT") &"！');history.back(-1);</script>" 
		response.end
	End If
End If

If visitornum<>0 And KS.C("UserName")="" And KS.C("PassWord")="" Then
    If HasMemberBuyNum<>0 Then 
	 KS.Die "<script>alert('对不起，您已用会员身份购买过["&RS("Title")&"]了，请明天再购买!');history.back(-1)</script>"
	End If
	If HasVisitorBuyNum>=visitornum Then
		Call DelProduct(RS("ID"))
		Session("ProductList")=ProductList
		response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]限制每位游客每天只能购买" & visitornum & RS("unit") & "！');history.back(-1);</script>" 
		response.End()
	 ElseIf HasVisitorBuyNum+KS.ChkClng(Session("Amount"&RS("ID")))>visitornum Then
		Session("Amount"&RS("ID")) = visitornum-HasVisitorBuyNum
		response.write "<script language=javascript>alert('对不起，["&RS("Title")&"]限制每位游客每天只能购买" & MemberNum & RS("unit") & ",您只能再购买" & Session("Amount"&RS("ID")) &  RS("UNIT") &"！');history.back(-1);</script>" 
		response.end
	 End If
End If
End Sub
%>
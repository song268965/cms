<%
'================================商城系统开始================================
		   case "getproductname" echo GetNodeText("title")
		   case "getproductintro" echo KS.ReplaceInnerLink(GetNodeText("prointro"))
		   case "getproductid" echo GetNodeText("proid")
		   case "getproducturl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"),GetNodeText("adddate"))
		   case "getproductmodel" echo GetNodeText("promodel")
		    case "getbrandid" echo GetNodeText("brandid")
		   case "getbrandname" if GetNodeText("brandid")="0" then echo "---" else echo "<a href=""" & KS.GetDomain & "shop/brand.asp?brandid=" &GetNodeText("brandid") & """ target=""_blank"">"  &KS.C_B(KS.ChkClng(GetNodeText("brandid")),"brandname") & "</a>"
		   case "getbrandename" echo KS.C_B(KS.ChkClng(GetNodeText("brandid")),"brandename")
		  case "getbrandphoto" 
		  Dim BrandPhoto:BrandPhoto=KS.C_B(KS.ChkClng(GetNodeText("brandid")),"photourl")
		  If KS.IsNul(BrandPhoto) Then   BrandPhoto="/Images/nopic.gif"
		  echo BrandPhoto
		   case "getproductspecificat" echo GetNodeText("prospecificat")
		   case "gettrademarkname" echo GetNodeText("trademarkname")
		   case "getserviceterm" echo GetNodeText("serviceterm")
		   case "getproducername" echo GetNodeText("producername")
		   case "getproducttype" echo GetProductType(GetNodeText("producttype"))
		   case "gettotalnum"    echo "<script src='" & DomainStr & "shop/getstock.asp?id=" & GetNodeText("id") &"'></script>"
		   case "getproductunit" echo GetNodeText("unit")
		   case "gethassold" echo "<script src='" & DomainStr & "shop/getstock.asp?action=HasSold&id=" & GetNodeText("id") &"'></script>"
		   'conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where (o.status=1 or o.status=2)")(0)
		   case "getprice_market" echo KS.GetPrice(GetNodeText("price"))
		   case "getvipprice" echo KS.GetPrice(GetNodeText("vipprice"))
		   case "getprice_shop" 
		     dim price_member:price_member=KS.GetPrice(GetNodeText("price_member"))
			if price_member<1 then	price_member="0"&price_member
			 
		     if KS.ChkClng(GetNodeText("istype"))=0 then
			 	echo "￥" &price_member &"元"
			 else
				if KS.ChkClng(price_member)=0 then
					echo  GetNodeText("score") &"分 "	 
				else
					echo  GetNodeText("score") &"分+" & price_member&"元"	 	
				end if
				
			 end if
		  
			
		   case "getprice" 
		   if ModelID=5 Then 
		     Echo KS.GetPrice(GetNodeText("price"))
		   Else 
		   	Echo GetNodeText("price")
		   end if	
		   
		   	case "getprice_member" 
			 Echo KS.GetPrice(GetNodeText("price_member"))
		   case "getgroupprice" echo "<span id='vipprice'><script src=""" & DomainStr &"shop/GetGroupPrice.asp?t=p&ProID=" & ItemID & """></script></span>"
		  case "getscore"   echo "<script src=""" & DomainStr &"shop/GetGroupPrice.asp?t=s&ProID=" & ItemID & """ type=""text/javascript""></script>"
		   case "getaddcarattribute"
		        dim cartaction:cartaction=DomainStr & "shop/shoppingcart.asp"
		        if fcls.callfrom3g="true" then cartaction="shop/shoppingcart.asp"
		        echo "<div class=""carbox""  style=""position:relative"">"
				echo "<form name=""cartform"" id=""cartform"" method=""post"" action=""" & cartaction &""">"
				echo "<input type=""hidden"" name=""action"" value=""Add"">"
				dim attr,attrarr,attrvalue,attrname,varr,k,vlen,alen
				attr=GetNodeText("attributecart") : alen=0 : vlen=0
				If KS.IsNul(attr) Then
				  echo "<script>var dir='" & DomainStr &"';</script>"
				Else
				   
				    dim rss:set rss=conn.execute("select top 500 * from KS_ShopSpecificationPrice where proid=" & GetNodeText("id") & " order by id")
					dim attr1str,attr11str,new1str,new2str,attr22str,new3str,attr33str
					dim minprice:minprice=0
					dim maxprice:maxprice=0
					dim jsstr:jsstr="<script type=""text/javascript"">var itemattr=new Array();"
					do while not rss.eof
					   attr1str=attr1str & rss("attr1")&"^" & rss("id")&",,,"
					   if ks.foundinarr(attr11str,rss("attr1"),",,,")=false then
					    attr11str=attr11str & rss("attr1")&",,,"
					   end if
					   if rss("price")<minprice or minprice=0 then minprice=rss("price")
					   if rss("price")>maxprice or maxprice=0 then maxprice=rss("price")
					   
					   if not ks.isnul(rss("attr2")) then
						   if ks.foundinarr(attr22str,rss("attr2"),",,,")=false then
							attr22str=attr22str & rss("attr2")&",,,"
						   end if
					   end if
					   
					   if not ks.isnul(rss("attr3")) then
						   if ks.foundinarr(attr33str,rss("attr3"),",,,")=false then
							attr33str=attr33str & rss("attr3")&",,,"
						   end if
					   end if
					   
					   jsstr=jsstr & "itemattr[" & rss("id") &"]=new Array('" & rss("attr1") & "','"& rss("attr2") & "','" & rss("attr3") &"'," & rss("price") & "," & rss("amount") & ");" 
					rss.movenext
					loop
					attrarr=split(attr,",")
					jsstr=jsstr &"var myitemname=new Array();"
					for i=0 to ubound(attrarr)
					jsstr=jsstr &"myitemname[" & i& "]='" &attrarr(i) & "';" 
					next
					if  minprice=maxprice then
				   	 jsstr=jsstr & "$('#vipprice').html('￥" & minprice & "元');"
					else
				   	 jsstr=jsstr & "$('#vipprice').html('￥" & minprice & "元~￥" & maxprice &"元');"
					end if
					jsstr=jsstr &"var dir='" & DomainStr &"';"
					jsstr=jsstr & "</script>"
					rss.close
				    echo jsstr
					
					attr11str=split(attr11str,",,,")
					attr1str=split(attr1str,",,,")
					for i=0 to ubound(attr11str)-1
					  dim ids:ids=""
					  for k=0 to ubound(attr1str)-1
					    if attr11str(i)=split(attr1str(k),"^")(0) then
						  if ids="" then
						   ids=split(attr1str(k),"^")(1)
						  else
						   ids=ids & "," &split(attr1str(k),"^")(1)
						  end if
						end if
					  next
					  if i=0 then
					   new1str=attr11str(i) & "^" & ids
					  else
					   new1str=new1str & ",,," & attr11str(i) & "^" & ids
					  end if
					next
				  	echo "<input type='hidden' id='attrid' name='attrid'>"
				  alen=ubound(attrarr)+1
				  for i=1 to alen
				   if Not KS.IsNul(attrarr(i-1)) Then
					   echo "<input type='hidden' id='attr"&i&"' name='attr" & i & "'>"
					   echo "<div id=""showattr" & i & """><span id='attrname" & i & "'>" & attrarr(i-1) & "：</span>"
				       
					    if i=1 then
						  new1str=split(new1str,",,,")
						  vlen=ubound(new1str)
						  dim itemvalue,iiarr
						  for k=0 to vlen
							iiarr=split(split(new1str(k),"^")(0),"|")
							if iiarr(1)<>"" then
							 itemvalue="<img src='" & iiarr(1) &"' width='25' height='25' title='" & iiarr(0) & "'/>"
							else
							 itemvalue=iiarr(0)
							end if
							echo "<span id=""att" & i & k & """ class=""txt"" onclick=""shop.getAttr(this," & i & "," & alen & ","&vlen&",2,'" & split(new1str(k),"^")(1) & "')"">" & itemvalue & "<i></i></span> "

						  next
						elseif i=2 and not ks.isnul(attr22str) then
						 new2str=split(attr22str,",,,")
						 vlen=ubound(new2str)-1
						 for k=0 to vlen
						    iiarr=split(split(new2str(k),"^")(0),"|")
							if iiarr(1)<>"" then
							 itemvalue="<img src='" & iiarr(1) &"' width='25' height='25' title='" & iiarr(0) & "'/>"
							else
							 itemvalue=iiarr(0)
							end if
							echo "<span id=""att" & i & k & """ class=""txt"" onclick=""shop.getAttr(this," & i & "," & alen & ","&vlen&")"">" & itemvalue & "<i></i></span> "
						 next
						elseif i=3 and not ks.isnul(attr33str) then
						 new3str=split(attr33str,",,,")
						 vlen=ubound(new3str)-1
						 for k=0 to vlen
						    iiarr=split(split(new3str(k),"^")(0),"|")
							if iiarr(1)<>"" then
							 itemvalue="<img src='" & iiarr(1) &"' width='25' height='25' title='" & iiarr(0) & "'/>"
							else
							 itemvalue=iiarr(0)
							end if
							echo "<span id=""att" & i & k & """ class=""txt"" onclick=""shop.getAttr(this," & i & "," & alen & ","&vlen&")"">" & itemvalue & "<i></i></span> "
						 next
						end if
				      echo "</div>"
				   
				   
				      
					End If
				  next
				End If
			
				echo "<div style='height:40px;line-height:40px;'>我要买：<a href='javascript:shop.buynum(0)' style='position:relative;top:2px;padding:5px;border:1px solid #ccc; background:#EBEBEB;font-size:20px; font-weight:bold;'>-</a> <input type=""text"" onkeyup=""this.value=this.value.replace(/\D/,'');"" name=""Q_" & GetNodeText("id") & """ id=""num"" size=""4"" value=""1"" style=""text-align:center; height:32px;""> <a href='javascript:shop.buynum(1)' style='position:relative;top:2px; padding:5px;border:1px solid #ccc; background:#EBEBEB;font-size:20px; font-weight:bold;'>+</a> " & GetNodeText("unit") & "<label style=""color:#999"">(库存<label id='stock'><script src='" & DomainStr & "shop/getstock.asp?id=" & GetNodeText("id") &"'></script></label><font id='unit'>"& GetNodeText("unit") & "</font>)</label></div><div id=""buyselect""></div><div><input type='hidden' name='AttributeCart' id='AttributeCart' value=''><input type='hidden' name='ID' id='ID' value='" & GetNodeText("id") &"'>"
				echo "<input type=""button"" onclick=""shop.gobuy(" & alen & ")"" id=""buybtn""> "
				echo "<input name=""istype""  type=""hidden"" id=""istype""  value=""" & KS.ChkClng(GetNodeText("istype"))&""">"
				echo "<input type=""button"" style=""position:relative;"" id=""carbtn"" onclick=""shop.addCart(event," & GetNodeText("id") & "," & alen & ")""></div>"
				
				'=========================显示捆绑商品================================
				Dim BNode,RS,GXML:Set RS=Conn.Execute("Select top 5 I.Tid,I.Fname,I.ID,I.Title,I.Price,I.PhotoUrl,B.KBPrice,B.Proid From KS_Product I Inner Join KS_ShopBundleSale b on i.id=b.kbproid Where B.proid=" & itemid)
				 If Not RS.Eof Then
					Set GXML=KS.RsToXml(RS,"row","")
				 End If
				 RS.Close:Set RS=Nothing
				 If IsObject(GXML) Then
					 echo "<div style=""font-weight:bold;margin:3px 0px 2px 2px"">您可以同时选购以下促销商品：</div>"
					 echo "<div><ul>"
					 For Each BNode In GXML.DocumentElement.SelectNodes("row")
						 echo "<li style='text-align:left'><input type='checkbox' value='" & BNode.SelectSingleNode("@id").text & "' name='Bundid'/><a href=""" & KS.GetItemUrl(5,BNode.SelectSingleNode("@tid").text,BNode.SelectSingleNode("@id").text,BNode.SelectSingleNode("@fname").text,BNode.SelectSingleNode("@adddate").text) & """ target=""_blank"" title=""" & BNode.SelectSingleNode("@title").text & """>" & KS.Gottopic(BNode.SelectSingleNode("@title").text,35) & "</a> <font style='color:#ff6600'>加￥" & formatnumber(BNode.SelectSingleNode("@kbprice").text,2,-1) & "元</font></li>"
					 Next
					 echo "</ul></div>" &vbcrlf
				 End If
				'============================================================================
				
				echo "</form></div>"
		   case "getaddcar"  echo "<a href=""" & DomainStr & "shop/ShoppingCart.asp?Action=Add&ID=" & ItemID & """ target=""_blank""><img src=""" & DomainStr & "Images/car.gif"" border=""0"" alt=""购物车""/></a>"
		   case "getaddfav"  echo "<a href=""" & DomainStr & "User/User_Favorite.asp?Action=Add&ChannelID=5&InfoID=" & ItemID & """ target=""_blank""><img src=""" & DomainStr & "Images/fav.gif"" border=""0"" alt=""加入收藏""/></a>"
		   case "getproductkeyword" echo Replace(GetNodeText("keywords"), "|", ",")
           case "getproductphotourl" echo GetNodeText("bigphoto")
		   case "getproductdate" echo LFCls.Get_Date_Field(GetNodeText("adddate"), "YYYY年MM月DD日")
		   case "getproductinput"   echo "<a href=""" & DomainStr & "Space/?" & GetNodeText("inputer") &""" target=""_blank"">" & GetNodeText("inputer") & "</a>"
		   case "getproductproperty"
		     If GetNodeText("recommend") = "1" Then Echo "<span title=""推荐"" style=""cursor:default;color:green"">荐</span> "
			 If GetNodeText("popular") = "1" Then  echo "<span title=""热门"" style=""cursor:default;color:red"">热</span> "
			 If GetNodeText("strip")="1" Then echo "<span title=""今日头条"" style=""cursor:default;color:#0000ff"">头</span> "
			 If GetNodeText("rolls") = "1" Then echo "<span title=""滚动"" style=""cursor:default;color:#F709F7"">滚</span> "
			 If GetNodeText("slide") = "1" Then echo "<span title=""幻灯片"" style=""cursor:default;color:black"">幻</span>"
		   case "getgroupbuyindexurl"
		     If KS.ChkClng(KS.Setting(179))=1 Then echo KS.GetDomain & "groupbuy/" Else Echo KS.GetDomain & "shop/groupbuy.asp"
		   case "getgroupbuyhistoryurl"
		     If KS.ChkClng(KS.Setting(179))=1 Then echo KS.GetDomain & "groupbuy/history/" Else Echo KS.GetDomain & "shop/groupbuy.asp?flag=history"
 '================================商城系统结束================================
%>
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
Set KSCls = New PackCls
KSCls.Kesion()
Set KSCls = Nothing


Class PackCls
        Private KS, KSR,ID,Template,RS,Action,PackID
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
		   PackID=KS.ChkClng(KS.S("id"))
		   'If PackID=0 Then KS.Die "非法参数!"
		   Action=KS.S("Action")
		   Select Case  Action
		         case "productlist" ProductList : Exit Sub 
				 case "showdetail" showdetail : Exit Sub
				 case "addChoseCart" 
		          InitialPack
				  addChoseCart 
				  RS.Close : Set RS=Nothing
				  Exit Sub
				 case "checksubmit"
		          InitialPack
				  checksubmit
				  RS.Close : Set RS=Nothing
				  Exit Sub
				 case "checkthsubmit"
		          InitialPack
				  checkthsubmit
				  RS.Close : Set RS=Nothing
				  Exit Sub
				 case "removeproduct" removeproduct : Exit Sub
				 case "choselist"  
				   InitialPack
				   choselist 
				   RS.Close :Set RS=Nothing
				   Exit Sub
				 case else
				   InitialPack
				   Template = KSR.LoadTemplate(RS("TemplateID"))
				   FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   
				   ReplacePackLabel
				   Template=KSR.KSLabelReplaceAll(Template)
				   Response.Write Template  
		           RS.Close : Set RS=Nothing
		   End Select
		   
		   
		End Sub
		
		'从数据库中读取礼包数据
		Sub InitialPack()
		   Set RS=Server.CreateObject("ADODB.RECORDSET")
		   If PackID=0 Then
		   RS.Open "Select top 1 * from KS_ShopPackage Where status=1 order by id desc",conn,1,1
		   Else
		   RS.Open "Select top 1 * from KS_ShopPackage Where status=1 and ID=" & PackID,conn,1,1
		   End If
		   If RS.Eof And RS.Bof Then 
		      RS.Close : Set RS=Nothing
			  KS.Die "找不到礼包或已关闭!"
		   End If
		   PackID=RS("ID")
		End Sub
		
		Sub ReplacePackLabel()
		   Template=Replace(Template,"{$GetPackID}",PackID)
		   Template=Replace(Template,"{$GetPackName}",rs("packname"))
		   Template=Replace(Template,"{$GetPackIntro}",rs("content"))
		   Template=Replace(Template,"{$GetPackNum}",rs("num"))
		   Template=Replace(Template,"{$GetPackDiscount}",rs("discount"))
		   Dim PhotoUrl,BigPhoto,RSL,str
		   PhotoUrl=RS("PhotoUrl")
		   If KS.IsNul(PhotoUrl) Then PhotoURl="/images/nopic.gif"
		   Template=Replace(Template,"{$GetPackPhotoUrl}",PhotoUrl)
		   PhotoUrl=RS("BigPhoto")
		   If KS.IsNul(PhotoUrl) Then PhotoURl="/images/nopic.gif"
		   Template=Replace(Template,"{$GetPackBigPhotoUrl}",PhotoUrl)
		   
		   '顶部自选包
		   if instr(Template,"{$GetZXPackage}")<>0 Then
			 Set RSL=Conn.Execute("Select top 10 ID,PackName From KS_ShopPackage Where PackType=0 and Status=1 Order by id desc")
			 Do While Not RSL.Eof 
			  str=str & "<li><a href='?id=" & rsl(0) & "'>" & RSl(1) & "</a></li>"
			 RSL.MoveNext
			 Loop
			 RSL.Close
			 Set RSL=Nothing
			 Template=Replace(Template,"{$GetZXPackage}",str)
		   End If
		   str=""
		   if instr(Template,"{$GetTHPackage}")<>0 Then
			 Set RSL=Conn.Execute("Select top 10 ID,PackName From KS_ShopPackage Where PackType=1 and Status=1 Order by id desc")
			 Do While Not RSL.Eof 
			  str=str & "<li><a href='?id=" & rsl(0) & "'>" & RSl(1) & "</a></li>"
			 RSL.MoveNext
			 Loop
			 RSL.Close
			 Set RSL=Nothing
			 Template=Replace(Template,"{$GetTHPackage}",str)
		   End If
		   
		   '特惠礼包
		   If InStr(Template,"{$ShowProductList}")<>0 Then
		     Dim SqlStr,PXml,Pnode,Price,RealPrice,n,totalPrice
			 SqlStr="Select a.id,a.title,a.price,a.Price_Member,a.photourl,b.id as packproid from ks_product a inner join KS_ShopPackagePro b on a.id=b.proid where b.PackID=" &PackID & " and a.verific=1 order by a.id desc"
             Set RSL=Conn.Execute(SqlStr)
			 If Not RSL.Eof Then
			   Set PXML=KS.RsToXml(RSL,"row","")
			 End If
			 RSL.Close : Set RSL=Nothing
			 str="<h3>套装详情:</h3>"
			 str=str & "<table border='0' width='100%'>"
			 str=str & "<tr align='center' class='thead'><td>商品名称</td><td>数量</td><td>原价</td><td>现价</td></tr>"
			 If IsObject(Pxml) Then
			        n=1:totalPrice=0
			   For Each Pnode In Pxml.DocumentElement.SelectNodes("row")
				   IF KS.C("UserName")<>"" Then
					   Price=pNode.SelectSingleNode("@price_member").text
					Else
					  Price=pNode.SelectSingleNode("@price").text
					End If
			        RealPrice=Price*rs("discount")/10
					totalPrice=totalPrice+realprice
			     str=str & "<tr class='td'><td title='" & PNode.SelectSingleNode("@title").text& "'>" & n &"、" & KS.Gottopic(PNode.SelectSingleNode("@title").text,36) & "</td><td align='center'>1</td><td align='center'><strike>￥" & formatnumber(Price,2,-1) &"元</strike></td><td align='center'><font color=#ff6600>￥" & formatnumber(RealPrice,2,-1) & " 元</font></td></tr>"
				  n=N+1
			   Next
			      str=str &"<tr><td colspan=""8"" style=""text-align:right""><strong>本套装合计：<font color=""red"">￥" & formatnumber(totalprice,2,-1) & "元</font></strong></td></tr>"
			      str=str &"<tr><td colspan=""8"" style=""height:50px;text-align:center;""><a href=""javascript:checkCS()""><img src=""images/hesuan.gif"" border=""0"" /></a></td></tr>"
			 Else
			      str=str &"<tr><td align=""center"">该套装下还没有添加商品!</td></tr>"
			 End If
			 str=str &"</table>"
			 Template=Replace(Template,"{$ShowProductList}",str)
			 Template=Replace(Template,"{$TotalPrice}",totalprice)
		   End If
		   
		   
		End Sub
		
		'选购品列表
		Sub ProductList()
		 MaxPerPage=12
		 Page=KS.ChkClng(Request("page"))
		 If Page=0 Then Page=1
	
		 Dim SqlStr:SqlStr="Select a.id,a.title,a.price,a.Price_Member,a.photourl,b.id as packproid from ks_product a inner join KS_ShopPackagePro b on a.id=b.proid where b.PackID=" &PackID & " and a.verific=1 order by a.id desc"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open SqlStr,conn,1,1
		 If RS.EOF And RS.Bof Then
		    RS.Close
			Set RS=Nothing
			KS.Die "对不起,该礼包还没有添加选购品!"
		 Else
			totalPut = RS.Recordcount
			If Page < 1 Then Page = 1
			If (totalPut Mod MaxPerPage) = 0 Then
				PageNum = totalPut \ MaxPerPage
			Else
				PageNum = totalPut \ MaxPerPage + 1
			End If

			If (Page - 1) * MaxPerPage < totalPut Then
				RS.Move (Page - 1) * MaxPerPage
			Else
				Page = 1
			End If
			Set XML=KS.ArrayToxml(Rs.GetRows(MaxPerPage),Rs,"row","xml")					
			RS.Close : Set RS=Nothing
			Dim PhotoUrl
			For Each Node In XML.DocumentElement.SelectNodes("row")  
			  PhotoUrl=Node.SelectSingleNode("@photourl").text
			  If KS.IsNul(PhotoUrl) Then PhotoUrl="/Images/Nopic.gif"
			  KS.Echo "<li>"
			  KS.Echo "<a href='javascript:showproduct(" & Node.SelectSingleNode("@id").text & ")'><img src='" & PhotoUrl & "' border='0' alt='" & Node.SelectSingleNode("@title").text & "' /></a><br/><span class='t'>" & KS.Gottopic(Node.SelectSingleNode("@title").text,25) & "</span><br/><strike>￥" & Node.SelectSingleNode("@price").text & "</strike> 元    <span style=""font-size:14px;font-weight:bold;color:Red;"">￥" & Node.SelectSingleNode("@price_member").text & "元</span>"
			  KS.Echo "</li>"
			Next
			
	     End If
		 
		      KS.Echo "<div class=""qspage"" style=""clear:both"" >页次<font color=red>" & Page & "</font>/" & PageNum & "页,共<font color=red>" & TotalPut & "</font>条信息"
			  if page>1 then
			  KS.Echo " <a href=""javascript:loading(1);"">首页</a>"
			  KS.Echo " <a href=""javascript:loading(" & page-1 & ");"">上一页</a>"
			  Else
			  KS.Echo " 首页"
			  KS.Echo " 上一页"
			  end if
			  
			  If page<>PageNum Then
			  KS.Echo " <a href=""javascript:loading(" & page+1 & ");"">下一页</a>"
			  KS.Echo " <a href=""javascript:loading(" & pagenum & ");"">末页</a>"
			  Else
			  KS.Echo " 下一页"
			  KS.Echo " 末页"
			  End If
			  KS.Echo "</div>"
		End Sub
		
		'显示详情
		Sub showdetail()
		 Dim RS,PhotoUrl,ProID
		 ProID=KS.ChkClng(Request("proid"))
		 Set RS=Server.CreateObject("Adodb.recordset")
		 RS.Open "Select top 1 * From KS_Product Where verific=1 and deltf=0 and ID=" & ProID,conn,1,1
		 If RS.Eof And RS.BOf Then
		    RS.Close:Set RS=Nothing
		    KS.Die "找不到商品,可能已删除!"
		 End If
		 PhotoUrl=RS("PhotoUrl")
		 If KS.IsNul(PhotoUrl) Then PhotoUrl="/images/nopic.gif"
		%>
	 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
     <HTML xmlns="http://www.w3.org/1999/xhtml">
		<head>
		 <script src="../ks_inc/jquery.js"></script>
		 <script src="js/shop.detail.js"></script>
		 <script type="text/javascript">
		  var dir='<%=KS.Setting(3)%>';
		  function addChoseCart(proid,a){
		     if (shop.check(a)){
			      $.get("pack.asp",{action:"addChoseCart",attrid:$("#attrid").val(),id:<%=PackID%>,proid:proid,num:$("#num").val(),AttributeCart:escape($("#AttributeCart").val())},function(d){
				   if (d!=''&&d!=null){
				      if (d!='success'){
					   alert(d);
					  }else{
					   parent.loadcart();
					   top.box.close();
					  }
				   }
				});
			}
		  }
		 </script>
		  <style type="text/css">
		   body{font-size:12px}
		   td{font-size:12px}
		.carbox{margin:5px;padding:10px;border:1px solid #f9c943;background:#FFFFF6;}
		.carbox span{float:left;min-width:20px;display:block;text-align:center;}
		.carbox span.txt{display:block;cursor:default;border:1px #c9c8ce solid; padding:1px;margin-right:8px; color:#646464; font-family:Arial, Helvetica, sans-serif; background:#fff; margin-bottom:0px; white-space:nowrap;padding-top:3px;}
		.carbox span.txt i{display:none;}
		.carbox span.txt:Hover{border:2px #ff6701 solid;padding:0px;padding-top:2px;}
		.carbox span.curr{position:relative;  margin-right:8px; padding:0px;padding-top:2px; border:2px #ff6701 solid;}
		.carbox span.curr i{display:inline;background:url(/shop/images/item_sel.gif) no-repeat 0 0;height:12px;overflow:hidden;width:12px;position:absolute;bottom:-1px;right:-2px;text-indent:-9999em;}
		.carbox div{clear:both}
 </style>
		</head>
		<body>
		
		<table width="98%" align="center">
			<tr>
			<td width="320">
			<table border="0" cellpadding="0" cellspacing="3" class="borderX">
			  <tr><td width="310" height="310" align="center"><A href="../item/show.asp?d=<%=proid%>&m=5" target="_blank"><img src='<%=PhotoUrl%>' width='310' height='310' border="0" id="current_img"/></A></td>
			</tr>
			<tr><td align="center"><br/><a onClick="window.open('ShowPic.asp?id=<%=rs("id")%>&u='+jQuery('#current_img').attr('src'))" href="javascript:;"><img src="images/look.gif" border="0"/></a></td></tr>
			</table>
			</td>
			<td valign="top" style="border:#ccc 1px solid;"><table width="100%" border="0" cellspacing="5" cellpadding="0">
			  <tr>
				<td height="25"><strong><%=RS("Title")%></strong></td>
			  </tr>
			  <tr>
				<td height="25">产品编号：<%=RS("ProID")%></td>
			  </tr>

			  <tr>
				<td height="25">上市时间：<%=formatdatetime(RS("AddDate"),2)%></td>
			  </tr>
			  <tr>
				<td height="25">原&nbsp;&nbsp;价：￥<%=RS("Price")%> 元</td>
			  </tr>
			  <tr>
				<td height="25">商城价：￥<%=RS("Price_Member")%> 元</td>
			  </tr>
			  <tr>
				<td height="25">VIP价格：<span style="color:#ff6600;font-weight:bold" id="vipprice"><script src="GetGroupPrice.asp?t=p&ProID=<%=proid%>" type="text/javascript"></script></span></td>
			  </tr>
			</table>
			<table width="100%">
			<tr><td height="25">
			<form name="cartform" id="cartform" method="post" action="pack.asp">
			<%
			  KS.Echo "<div class=""carbox"">"
			  KS.echo "<input type=""hidden"" name=""action"" value=""Add"">"
				dim attr,attrarr,attrvalue,attrname,varr,k,vlen,alen,i
				attr=rs("attributecart") : alen=0 : vlen=0
				If Not KS.IsNul(attr) Then
				   
				    dim rss:set rss=conn.execute("select * from KS_ShopSpecificationPrice where proid=" & rs("id") & " order by id")
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
					jsstr=jsstr & "</script>"
					rss.close
				    ks.echo jsstr
					
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
				  	ks.echo "<input type='hidden' id='attrid' value='0' name='attrid'>"
				  alen=ubound(attrarr)+1
				  for i=1 to alen
				   if Not KS.IsNul(attrarr(i-1)) Then
					   ks.echo "<input type='hidden' id='attr"&i&"' name='attr" & i & "'>"
					   ks.echo "<div style='height:40px' id=""showattr" & i & """><span id='attrname" & i & "'>" & attrarr(i-1) & "：</span>"
				       
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
							ks.echo "<span id=""att" & i & k & """ class=""txt"" onclick=""shop.getAttr(this," & i & "," & alen & ","&vlen&",2,'" & split(new1str(k),"^")(1) & "')"">" & itemvalue & "<i></i></span> "

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
							ks.echo "<span id=""att" & i & k & """ class=""txt"" onclick=""shop.getAttr(this," & i & "," & alen & ","&vlen&")"">" & itemvalue & "<i></i></span> "
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
							ks.echo "<span id=""att" & i & k & """ class=""txt"" onclick=""shop.getAttr(this," & i & "," & alen & ","&vlen&")"">" & itemvalue & "<i></i></span> "
						 next
						end if
				      ks.echo "</div><div style='clear:both'></div>"
				   
				   
				      
					End If
				  next
				End If
				
				ks.echo "<div>我要买：<input type=""text"" onkeyup=""this.value=this.value.replace(/\D/,'');"" name=""Q_" & rs("id") & """ id=""num"" size=""4"" value=""1"" style=""text-align:center""> " & rs("unit") & "<label style=""color:#999"">(库存<label id='stock'><script src='getstock.asp?id=" & rs("id") &"'></script></label>"& rs("unit") & ")</label></div><br/><div id=""buyselect""></div><div><input type='hidden' name='AttributeCart' id='AttributeCart' value=''><input type='hidden' name='ID' id='ID' value='" & rs("id") &"'> </div></div>"
			%>
			</td></tr>
			<tr><td align=right height=30><img src="images/addcart.gif"  onclick="addChoseCart(<%=rs("id")%>,<%=alen%>)" border=0 style="CURSOR: hand"></td></tr>
			</form>
			<tr><td align="right">
			</td></tr>
			</table></td>
			</tr></table>
		</body>
		</html>
		<%
		  RS.Close
		  Set RS=Nothing
		End Sub
		
		'商品选择列表
		Sub choselist()
		 Dim i,ProductList
		 Dim Num:Num=rs("num")

		   Dim RSObj,RealPrice,totalPrice ,n,PhotoUrl,str,productnamestr,haspro
		   totalprice=0:haspro=false
		   Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select a.id,a.title,a.Price_Member,a.Price,a.PhotoUrl,b.Amount,b.AttributeCart,b.attrid from KS_Product a inner join KS_ShopPackageSelect b on a.id=b.proid where b.packid=" & packid & " and a.verific=1 order by a.id desc",conn,1,1
		   if not RSObj.eof then
			  str=str & Escape("您已选择<font color=red>" & RSObj.recordcount & "</font>样商品")
			  str=str & "<table border=0>"
			  n=0
			   Do While Not RSObj.Eof
			    haspro=true
				
				If RSObj("AttrID")<>0 Then 
				  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RSObj("AttrID"))
				  If Not RSAttr.Eof Then
					RealPrice=RSAttr("Price")
				  Else
					RealPrice=RSObj("Price_Member")
				  End If
				  RSAttr.CLose:Set RSAttr=Nothing
				 Else
					RealPrice=RSObj("Price_Member")
				 End If
				
			
				
				Dim ProNum,ProID
				ProNum=KS.ChkClng(rsobj("Amount"))
				If ProNum=0 Then ProNum=1
				
				PhotoUrl=RSObj("PhotoUrl")
				ProID=RSObj(0)
				If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
				
				For I=1 To ProNum
				 productnamestr=productnamestr & "<span style='cursor:pointer' title='" &RSOBJ(1) &"'>" & KS.Gottopic(RSOBJ(1),23) & "</span><br/>"

                dim nowprice:nowprice=round(RealPrice*(rs("discount")/10),2)
				if nowprice<1 then nowprice="0" & nowprice
				totalPrice=totalPrice+nowprice
				str=str & Escape("<tr><td style=""border-bottom:1px dashed #ccc;""><table border='0'><tr><td><img src='" & PhotoUrl & "' width='72' height='96' title='" & rsobj(1) & "'></td><td><font color=brown><span style='cursor:pointer' title='" &RSOBJ(1) &"'>" & KS.Gottopic(RSOBJ(1),23) & "</span> " & rsobj("AttributeCart") & "</font><br/>原价:<strike>￥" & RealPrice & "</strike> 元 <br/> 现价:￥" & nowprice  & "</td></tr><tr><td colspan=2 align='center'><a href='javascript:removeproduct(" & ProID & ")'><img src='images/ks_index_058.gif' border='0'/></a></td></tr></table> </td></tr>")
				 n=n+1
				Next
				
			   RSObj.MoveNext
			   Loop
			   str=str & "</table><br/>"
		   end if
		   RSObj.Close
		   Set RSObj=Nothing
		   
		   IF totalPrice<1 Then totalPrice="0" & totalPrice
	  %>
	  		<h3>礼包总件数：<span style=" font-size:14px; font-weight:bold; color:red;"><%=rs("Num")%></span></h3>
			<hr size="1" />
                <SPAN style="color:green"><%=productnamestr%></SPAN>
				礼包优惠折扣：  <span style=" font-size:14px; font-weight:bold; color:red;"><%=RS("Discount")%></span>折<BR>
                 你选商品的总价是：<span style=" font-size:20px; font-weight:bold; color:red;">￥<%=totalPrice%></span>元
		
		<hr size="1" />
		<h3>以下是您所选的商品清单:</h3>
	  <%
	      KS.Echo Str
		%>
		
		
		<table width="100%" border="0">
		 <%
		 if n=0 then n=1 else n=n+1
		 for i=n to num%>
		  <tr>
			<td>
			商品名称<br/>
			原价:￥0 元<br/>
			现价:￥0 元
			</td>
		  </tr>
		  <tr>
		    <td colspan=2 align=center></td></tr>
		 <%next%>
		</table>
		<%
		 if haspro Then
		  KS.Echo "<div style='margin-top:10px;text-align:center'><a href='javascript:checkCS()'><img src='images/hesuan.gif' border='0'/></a></div>"
		 End If

		End Sub
		
		'用户名
		Function GetUserID()
		  If KS.C("UserName")="" Then
		    If KS.C("CartID")="" Then
		    Response.Cookies(KS.SiteSn)("CartID")=KS.R(KS.MakeRandomChar(20))
			End If
			GetUserID=KS.C("CartID")
		  Else
		    GetUserID=KS.C("UserName")
		  End If
		End Function 
		
		
		'加入左边选择列表,ajax调用
		Sub addChoseCart()
		   Dim Prodid:Prodid=KS.ChkClng(request("proid"))
		   Dim AttrID:AttrID=KS.ChkClng(Request("AttrID"))
		   if Prodid=0 then KS.Die ""
		   Dim Num:Num=KS.ChkClng(Request("Num"))
		   If Num=0 Then Num=1
		   If Num > rs("Num") Then
		     KS.Die "对不起,此礼包最多只能选择" & rs("num") & "件商品!"
		   End If
		   
		   Dim Total:Total=0
		   Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
		   RSS.Open "Select * From KS_ShopPackageSelect Where PackID=" & PackID & " and UserName='" & GetUserID &"'",conn,1,1
		   If Not RSS.Eof Then
		     Do While Not RSS.Eof 
			   Total=Total+RSS("Amount")
			   RSS.MoveNext
			 Loop
		   End If
		   RSS.Close
		   If Total>rs("Num") Then
		    KS.Echo "对不起,此礼包最多只能选择" & rs("num") & "件商品!"
			Exit Sub
		   ElseIf Total+Num>rs("Num") Then
		    KS.Echo "对不起,此礼包最多只能选择" & rs("num") & "件商品!"
			Exit Sub
		   End If
		  
		   RSS.Open "Select Top 1 * From KS_ShopPackageSelect Where PackID=" & PackID & " and UserName='" & GetUserID &"' and proid=" & prodid &" and AttrID=" & AttrID ,conn,1,3
		   If RSS.Eof Then
		      RSS.AddNew
		   End If
		     RSS("AttrID")=AttrID
		     RSS("Proid")=Prodid
			 RSS("PackID")=PackId
			 RSS("AttributeCart")=KS.DelSQL(UnEscape(Request("AttributeCart")))
			 RSS("Amount")=Num
			 RSS("UserName")=GetUserID
			 RSS("AddDate")=NOW
           RSS.Update
		   RSS.Close:Set RSS=Nothing
		   
		   '删除超过5天的记录
		   Conn.Execute("Delete From KS_ShopPackageSelect Where datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>5")
		   KS.Echo "success"
		End Sub
		
		'移除商品
		Sub removeproduct()
		   Dim Prodid:Prodid=KS.ChkClng(request("proid"))
		   if Prodid=0 then KS.Die Escape("没有选择商品!")
		   Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
		   RSS.Open "Select top 1 * From KS_ShopPackageSelect Where PackID=" & PackID & " and proid=" & Prodid,conn,1,3
		   If Not RSS.Eof Then
		      If RSS("Amount")>1 Then
			    RSS("Amount")=RSS("Amount")-1
				RSS.Update
			  Else
			    RSS.Delete
			  End If
		   End If
		   RSS.Close:Set RSS=Nothing
		   KS.Echo "success"
		End Sub
		
		'检查自选礼包合法性
		Sub checksubmit()
		   Dim Total:Total=0
		   Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
		   RSS.Open "Select * From KS_ShopPackageSelect Where PackID=" & PackID & " and UserName='" & GetUserID &"'",conn,1,1
		   If Not RSS.Eof Then
		     Do While Not RSS.Eof 
			   Total=Total+RSS("Amount")
			   RSS.MoveNext
			 Loop
		   End If
		   RSS.Close
		   If Total>rs("Num") Then
		    KS.Echo "对不起,此礼包最多只能选择" & rs("num") & "件商品!"
		   ElseIf Total<>rs("Num") Then
		    KS.Echo "对不起,商品数量不够,您还需要挑选 " & Rs("Num")-Total & " 件!"
		   Else
		    KS.Echo "success"
		   End If
		End Sub
		'将特惠礼包的商加添加到KS_ShopPackageSelect
		Sub checkthsubmit()
			Dim SQLStr:SqlStr="Select a.id,b.id as packproid from ks_product a inner join KS_ShopPackagePro b on a.id=b.proid where b.PackID=" &PackID & " and a.verific=1 order by a.id desc"
			Dim RSG:Set RSG=Conn.Execute(SQLStr)
			If Not RSG.Eof Then
				Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
				Do While Not RSG.Eof
				RSS.Open "Select Top 1 * From KS_ShopPackageSelect Where PackID=" & PackID & " and UserName='" & GetUserID &"' and proid=" & rsg("id") &"",conn,1,3
				   If RSS.Eof Then
					  RSS.AddNew
				   End If
					 RSS("Proid")=rsg("id")
					 RSS("PackID")=PackId
					 RSS("AttributeCart")=""
					 RSS("Amount")=1
					 RSS("UserName")=GetUserID
					 RSS("AddDate")=NOW
				   RSS.Update
				   RSS.Close
				 RSG.MoveNext
				 Loop
				 Set RSS=Nothing
				 
				 '删除超过5天的记录
				 Conn.Execute("Delete From KS_ShopPackageSelect Where datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>5")
				 KS.Echo "success"
            Else
			     KS.Echo escape("对不起，该礼包下还没有添加商品!")
			End If
			RSG.Close:Set RSG=Nothing
		End Sub
		
End Class
%>

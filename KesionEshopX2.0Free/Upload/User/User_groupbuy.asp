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
Set KSCls = New Admin_MyShop
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MyShop
        Private KS,KSUser,ChannelID,DomainStr
		Private CurrentPage,totalPut,Status,ProducerName,FieldXML,FieldNode,FNode,FieldDictionary
		Private RS,MaxPerPage,ComeUrl,SelButton,Price_Original,Price,Price_Market,Price_Member,Point,Discount
		Private ClassID,Title,KeyWords,ProModel,ProSpecificat,ProductType,Unit,TotalNum,AlarmNum,TrademarkName,Content,Verific,PhotoUrl,RSObj,I,UserClassID,ShowONSpace,Weight,FileIds
		Private CurrentOpStr,Action,ID,ErrMsg,Hits,BigPhoto,BigClassID,SmallClassID,flag,BrandID,totalscore
		Private Sub Class_Initialize()
			MaxPerPage =12
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
       Public Sub loadMain()
		 IF KS.S("ComeUrl")="" Then
     		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		 Else
     		ComeUrl=KS.S("ComeUrl")
		 End If

		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=5
		If KS.C_S(ChannelID,6)<>5 Then Response.End()
		if conn.execute("select usertf from ks_channel where channelid=" & channelid)(0)=0 then
		  Response.Write "<script>alert('本频道关闭投稿!');window.close();</script>"
		  Exit Sub
		end if
		'设置缩略图参数
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
		
		Call KSUser.Head()
		%>
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" and  KS.S("Action")<>"groupbuyOrder" and KS.S("Action")<>"ShowOrder" then response.write " class='puton'"%>><a href="User_groupbuy.asp?ChannelID=<%=ChannelID%>">我发布的团购(<span class="red"><%=Conn.Execute("Select count(id) from KS_GroupBuy where username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='puton'"%>><a href="User_groupbuy.asp?ChannelID=<%=ChannelID%>&Status=1">已审核(<span class="red"><%=conn.execute("select count(id) from KS_GroupBuy where Verific=1 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='puton'"%>><a href="User_groupbuy.asp?ChannelID=<%=ChannelID%>&Status=0">待审核(<span class="red"><%=conn.execute("select count(id) from KS_GroupBuy where Verific=0 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='puton'"%>><a href="User_groupbuy.asp?ChannelID=<%=ChannelID%>&Status=3">被退(<span class="red"><%=conn.execute("select count(id) from KS_GroupBuy where Verific=3 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
                <% Dim ii,OrderIDstr
				   OrderIDstr=""
					Dim RSD:Set RSD=Conn.Execute("Select a.* From KS_OrderItem a inner join KS_GroupBuy b on a.ProID=b.id where b.Verific=1 and b.username='"&  KSUser.UserName &"'")
					
					Do While Not RSD.Eof
				  		ii=ii+1
						if ii=1 then 
							OrderIDstr="'"&RSD("OrderID")&"'" 
						else
							OrderIDstr= OrderIDstr & "," & "'"& RSD("OrderID") &"'" 
						end if
						RSD.MoveNext
					Loop
					RSD.Close : Set RSD=Nothing
				%>
               
                <li<%If KS.S("Action")="groupbuyOrder" or KS.S("Action")="ShowOrder" then response.write " class='puton'"%>><a href="User_groupbuy.asp?Action=groupbuyOrder">团购订单(<span class="red">
                
				<%
				if OrderIDstr="" then
					Response.Write("0") 
				else
					Response.Write conn.execute("Select count(id) From KS_Order Where ordertype=1 and OrderID in ("& OrderIDstr &") ")(0)
				end if
				%>
                </span>)</a></li>
               
                
			</ul>
		  </div>
		<%
		Action=KS.S("Action")
		Select Case Action
		 Case "Del"  GroupBuyDel()
		 Case "Add","Edit"
		  Call ShopAdd
		 Case "AddSave","EditSave"
          Call ShopSave()
		 Case "refresh" Call KSUser.RefreshInfo(KS.C_S(ChannelID,2))
		 Case "groupbuyOrder" Call groupbuyOrder()
		 Case "ShowOrder" Call ShowOrder()
		 Case Else
		  Call ShopList
		 End Select
       End Sub
	   Sub ShopList
		    CurrentPage = KS.ChkClng(KS.S("page")): If CurrentPage<=0 Then  CurrentPage = 1
           	Dim Param:Param=" Where username='"& KSUser.UserName &"'"
									Verific=KS.S("status")
									If Verific="" or not isnumeric(Verific) Then Verific=4
                                    IF Verific<>4 Then 
									   Param= Param & " and Verific=" & Verific
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
									Dim Sql:sql = "select a.*,CategoryName from KS_GroupBuy a inner join KS_GroupBuyClass b on a.ClassID=b.id "& Param &" order by AddDate DESC"

								  Select Case Verific
								   Case 0 
								    Call KSUser.InnerLocation("待审"& KS.C_S(ChannelID,3) & "列表")
								   Case 1
								    Call KSUser.InnerLocation("已审"& KS.C_S(ChannelID,3) & "列表")
								   'Case 2
								   'Call KSUser.InnerLocation("草稿"& KS.C_S(ChannelID,3) & "列表")
								   Case 3
								   Call KSUser.InnerLocation("退稿"& KS.C_S(ChannelID,3) & "列表")
                                   Case Else
								    Call KSUser.InnerLocation("所有"& KS.C_S(ChannelID,3) & "列表")
								   End Select
			   %>
			    <div class="writeblog"><img src="images/ico_05.gif" align="absmiddle"> <a href="User_groupbuy.asp?ChannelID=<%=ChannelID%>&Action=Add">发布团购</a></div>
                <script src="../ks_inc/jquery.imagePreview.1.0.js"></script>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                    <tr class="title">
                          <td width="6%" height="22" align="center">选中</td>
                          <td align="center" width="40">图片</td>
                          <td align="center"><%=KS.C_S(ChannelID,3)%>名称</td>
						  <td align="center"><%=KS.C_S(ChannelID,3)%>录入</td>
                          <td align="center">添加时间</td>
                          <td align="center">状态</td>
                          <td align="center">管理操作</td>
                   </tr>
                     <%
								Set RS=Server.CreateObject("AdodB.Recordset")
								RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' colspan='6' height=30 valign=top>没有你要的"& KS.C_S(ChannelID,3) & "!</td></tr>"
								 Else
									totalPut = RS.RecordCount
								   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									End If
										Call showContent
								End If
     %>                      <tr class='tdbg'>
                                     <form action="User_groupbuy.asp" method="post" name="searchform">
								  <td colspan="6">
										<strong><%=KS.C_S(ChannelID,3)%>搜索：</strong>
										  <select name="Flag">
										   <option value="0">名称</option>
										   <option value="1">关键字</option>
									      </select>
										  
										  关键字
										  <input type="text" name="KeyWord" onfocus="if (this.value=='关键字'){this.value=''}" class="textbox" value="关键字" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" 搜 索 ">
							      </td>
								    </form>
                                </tr>
							<tr class="title">
							 <td colspan=6>
							 <%=KS.C_S(ChannelID,3)%>销售说明：
                              </td>
                              </tr>
                              <tr>
                              <td>
							  1、用户在本站发布商品销售，购物方将货款首先支付到本网站；<br/>
							  2、购物方在本站支付成功后，本站将负责对货款及订单的有效性进行审核及通知销售方发货等；<br>
							  3、促成交易后
							  ，本站将收取货款总价的 <font color=red><%=KS.Setting(79)%>% </font>作为交易管理费,并将货款支付给销售方；<br>
							  3、请确保所发布商品真实性，一旦发现您在本站所发布信息含有虚假，期骗行为,我们将立即冻结您在本站的交易账户。
							 </td>
							</tr>
</table>
		  <%
  End Sub
  
  Sub ShowContent()
     Dim I
    Response.Write "<FORM Action=""User_groupbuy.asp?Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
	  Dim PhotoStr:PhotoStr=RS("PhotoUrl")
	 if PhotoStr="" Or IsNull(PhotoStr) Then PhotoStr=KS.GetDomain & "images/Nopic.gif"
	 %>
		 <tr class='tdbg'>
                   <td class="splittd" height="22" align="center">
					<INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID">
				   </td>
				  <td class="splittd"><a href="<%=PhotoStr%>" title="<%=rs("Subject")%>" class="preview"><img src="<%=photostr%>" width="32" height="32" /></a></td>
                  <td class="splittd" align="left">[<%=RS("CategoryName")%>]
					<a title="<%=rs("Subject")%>"  href="../shop/groupbuyshow.asp?id=<%=rs("id")%>" target="_blank" class="link3"><%=KS.GotTopic(trim(RS("Subject")),32)%></a>
				  </td>
				  <td class="splittd" align="center"><%=rs("username")%></td>
                  <td class="splittd" align="center"><%=formatdatetime(rs("AddDate"),2)%></td>
                   <td class="splittd" align="center">
											  <%Select Case KS.ChkClng(rs("Verific"))
											   Case 0
											     Response.Write "<span class=""font10"">待审</span>"
											   Case 1
											     Response.Write "<span class=""font11"">已审</span>"
                                              ' Case 2
											   '  Response.Write "<span class=""font13"">草稿</span>"
											   Case 3
											     Response.Write "<span class=""font14"">被退</span>"
                                              end select
											  %></td>
                     <td class="splittd" height="22" align="center">
					    <%If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))=1 Then%>
						 <a href="?ChannelID=<%=ChannelID%>&action=refresh&id=<%=rs("id")%>" class="box">刷新</a>
						<%end if%>
											<%if rs("Verific")<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a class='box' href="User_groupbuy.asp?channelid=<%=channelid%>&id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>">修改</a> <a class='box' href="User_groupbuy.asp?channelid=<%=channelid%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除<%=KS.C_S(ChannelID,3)%>吗?'))">删除</a>
											<%else
											 If KS.C_S(ChannelID,42)=0 Then
											  Response.write "---"
											 Else
											  Response.Write "<a  class='box' href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"'>修改</a> <a class='box' href='#' disabled>删除</a>"
											 End If
											end if%>
											</td>
			</tr>
					   <tr><td colspan=6 background='images/line.gif'></td></tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
</table>
 <table width="100%" class="border">
         			<tr class='tdbg'>
					 <td valign=top style="padding-left:22px;">
							<label><INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中</label>&nbsp;<button class="pn pnc" onClick="return(confirm('确定删除选中的<%=KS.C_S(ChannelID,3)%>吗?'));" type="submit"><strong>删除选定</strong></button>  </FORM>       
					  </td>
					  <td align="right">
					<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>			
					  </td>
						
        </tr>
								<%
  End Sub
  
  Sub groupbuyOrder()
   dim SqlStr
 	%>
	<style>
	.tdbg-mouseover{ border:1px #FF9900 solid; }
	.mouseover-box{border-bottom:1px #006699 dashed}
    </style>
    <div class="writeblog">
				 <table border="0">
				<form action="user_order.asp" method="post" name="search">
				 <tr><td><strong>订单状态:</strong></td>
				 <td>
				<select class="select" name="OrderStatus">
				 <option value="">不限制</option>
				  <option value="0">等待确认</option>
				  <option value="1">已经确认</option>
				  <option value="2">已结清</option>
				</select></td>
				<td><strong>订单编号:</strong></td>
				<td>
				 <input type="text" name="keyword" class="textbox">
				 <input type="submit" value="快速搜索" class="button">
				</td>
				</tr>
				</form>				   
				</table>
				</div>

				<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					<tr class=title align=middle>
					  <td height="25" nowrap="nowrap">商品编号</td>
					  <td>商品名称</td>
					  <td>单价</td>
					  <td>数量</td>
					  <td>金额</td>
					  <td>赠送积分</td>
					  <td>其它</td>
					</tr>
					<%
					Dim ii,OrderIDstr
					OrderIDstr="'safgdfdf0000000233333322'"
					Dim RSD:Set RSD=Conn.Execute("Select a.* From KS_OrderItem a inner join KS_GroupBuy b on a.ProID=b.id where b.Verific=1 and b.username='"&  KSUser.UserName &"'")
					Do While Not RSD.Eof
				  		ii=ii+1
						if ii=1 then 
							OrderIDstr="'"&RSD("OrderID")&"'" 
						else
							OrderIDstr= OrderIDstr & "," & "'"& RSD("OrderID") &"'" 
						end if
						RSD.MoveNext
					Loop
					RSD.Close : Set RSD=Nothing
					  Dim Param:Param=" Where ordertype=1 and OrderID in ("& OrderIDstr &")"
					 
					  If KS.S("OrderStatus")<>"" Then 
					    Param=Param & " and status=" & KS.ChkClng(KS.S("OrderStatus"))
					  End If
					  If KS.S("KeyWord")<>"" Then  
					    Param=Param & " and OrderID like '%" & KS.S("KeyWord") & "%'"
					  End If
					     
						 SqlStr="Select * From KS_Order " & Param & " order by id desc"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1
                
				If RS.EOF And RS.BOF Then
					Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>您没有下任何订单!</td></tr>"
				Else
					totalPut = RS.RecordCount
					If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
					End If
					Call groupbuyShowContent
				End If
           %>
					
          </table>
		  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  </td>
		  </tr>
</table>
      </div>
    
	<%
  End Sub
  
  
  Sub groupbuyShowContent()
   Dim i
   Do While Not RS.Eof
		%>

         <tr class='tdbg'>
                 
                  <td height='30' nowrap="nowrap" class="splittd bgtitle" colspan="5">
					   
                        &nbsp;订单编号：<a href="User_Order.asp?Action=ShowOrder&ID=<%=RS("ID")%>"><%=rs("orderid")%></a> 
						 <%
						 if rs("ordertype")="1" then  response.write "<font color=red><b>团</li></b></font>"%>
                         <font color="#FF6600">购买用户:<%=rs("username")%></font>
                         <br /> 合计：<%=formatnumber(rs("NoUseCouponMoney"),2)%>元(含运费：<%=rs("Charge_Deliver")%>元) <strong>应付
                         <span style='color:brown'>￥<%=formatnumber(rs("Moneytotal"),2)%></span>元</strong>
						<%
                         if KS.ChkClng(rs("UseScoreisshop"))>0 then
                        Response.Write("<strong> + <span style='color:brown'>"&KS.ChkClng(rs("UseScoreisshop")) &"</span> 积分 </strong>")	
                        end if
						%>
						 <br />
                         订单状态：
				<%
				 response.write GetOrderStatus(rs)
				 If RS("isservice")="1" Then Response.Write "&nbsp;<a href='?Action=ShowOrder&ID=" & rs("id") & "#service'>查看服务记录</a>"
				 %>
                 
				  </td>
                  
                  
				  <td class="splittd bgtitle" align="center">
					   <%
					    if rs("totalscore")>0 and rs("DeliverStatus")<>3 then
						   response.write "<font color=green>" & rs("totalscore") & " 分</font>"
						   if rs("scoretf")=1 then
						     response.write "<font color=#999999>,已送</font>"
						   else
						     response.write "<font color=red>,未送</font>"
						   end if
						else
						   response.write "无"
						end if
					    %>
					   </td>
					  <td class="splittd bgtitle" nowrap="nowrap" style="text-align:center">
					  <a href="?Action=ShowOrder&ID=<%=rs("id")%>">订单详情</a>
					  </td>
               </tr>
         <%
		 
			 Dim OrderDetailStr,TotalPrice,attributecart,RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			   OrderDetailStr=""	   
			   RSI.Open "Select top 100 * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id",conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
			  Else
			  Do While Not RSI.Eof
			  attributecart=rsi("attributecart")
			  if not ks.isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
				OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"    
				OrderDetailStr=OrderDetailStr & "<td align='center' class='splittd'>" & rsi("proid") & "</td><td class='splittd'>" 
				 Dim OrderType:OrderType=KS.ChkClng(rs("ordertype"))

		
		
			  Dim PhotoUrl,SqlStr,RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  If OrderType=1 Then
			    SqlStr="Select top 1 Subject as title,PhotoUrl,'件' as unit,0 as IsLimitBuy,0 as LimitBuyPrice,0 as LimitBuyPayTime From KS_GroupBuy Where ID=" & RSI("ProID")
			  Else
			    SqlStr="Select top 1 I.Title,I.PhotoUrl,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID")
			  End If
			  RSP.Open SqlStr,conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  PhotoUrl=rsp("photourl")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(换购)"
				  else
				    if RSP("LimitBuyPayTime") then
					   If LimitBuyPayTime="" Then
					   LimitBuyPayTime=RSP("LimitBuyPayTime")
					   ElseIf LimitBuyPayTime>RSP("LimitBuyPayTime") Then
						LimitBuyPayTime=RSP("LimitBuyPayTime")
					   End If
					end if
				  end  if
				   OrderDetailStr=OrderDetailStr & "<img style=""width:53px;height:53px;border:1px solid #f1f1f1;padding:1px;"" class='img' onerror=""this.src='../images/nopic.gif';"" src='" & photourl &"' align='left'/>"
				  If OrderType=1 Then
				   OrderDetailStr=OrderDetailStr & "<a href='../shop/groupbuyshow.asp?id=" & RSi("proid") & "' target='_blank'>" & title & "</a>"
                  Else
				   OrderDetailStr=OrderDetailStr & "<a href='../item/show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & title & "</a>"
				  End If
				  
				  If RSI("IsLimitBuy")="1" Then OrderDetailStr=OrderDetailStr & "<span style='color:green'>(限时抢购)</span>"
				  If RSI("IsLimitBuy")="2" Then OrderDetailStr=OrderDetailStr & "<span style='color:blue'>(限量抢购)</span>"
			  End If
			  RSP.Close:Set RSP=Nothing
		
		OrderDetailStr=OrderDetailStr &  attributecart & "</td><td class='splittd' align='center'>" & formatnumber(rsi("realprice"),2) & "</td>"
		OrderDetailStr=OrderDetailStr & " <td class='splittd' align='center'>" & rsi("amount") &" " & Unit & "</td>    "
		OrderDetailStr=OrderDetailStr & " <td class='splittd' align='center'>" & formatnumber(rsi("realprice")*rsi("amount"),2) & "</td>"
		OrderDetailStr=OrderDetailStr & " <td class='splittd' align='center'>" & ks.chkclng(rsi("score")*rsi("amount")) & " 分</td>    "
		OrderDetailStr=OrderDetailStr & "<td class='splittd' align='center'>" 
		totalscore=totalscore+ks.chkclng(rsi("score")*rsi("amount"))

		
		OrderDetailStr=OrderDetailStr & "</td>  "
		OrderDetailStr=OrderDetailStr & " </tr> " 
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  '取得捆绑销售商品
		
		
			  TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
		End If
		
		OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '超值礼包 
			   

		 
		 response.write OrderDetailStr
		 
		 
		 
	
		 
		 %>
		  <tr class='tdbg'>
          		<td class="splittd bgtitle" height="5"   colspan="7" nowrap="nowrap" style="text-align:center">
                    <div  class="mouseover-box" ></div>
         		</td>
          </tr>
		 <%
		 
		 
		 
				RS.MoveNext
				I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  Loop

  End Sub
  
  
  '添加
  Sub ShopAdd
        Dim Subject,ActiveDate,AddDate,Intro,Highlights,Protection,Notes,Locked,EndTF,PhotoUrl,BigPhoto,ClassID,AllowBMFlag,AllowArrGroupID,minnum,Comment,Changes,ChangesUrl
		Dim Price_Original,Price,Discount,limitbuynum,weight,recommend,ProvinceID,CityID,HasBuyNum,MustPayOnline,CleanCart,showdelivery
		Call KSUser.InnerLocation("发布"& KS.C_S(ChannelID,3) & "")
		Action=KS.S("Action")
		ID=KS.ChkClng(KS.S("ID"))
                 If Action="Edit" Then
				  CurrentOpStr=" OK,修改 "
				  Action="EditSave"
				   Dim ShopRS:Set ShopRS=Server.CreateObject("ADODB.RECORDSET")
				   ShopRS.Open "Select top 1  * From KS_GroupBuy Where username='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
				   IF ShopRS.Eof And ShopRS.Bof Then
				     call KS.Alert("参数传递出错!",ComeUrl)
					 Exit Sub
				   Else
						If KS.C_S(ChannelID,42) =0 And ShopRS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
						   ShopRS.Close():Set ShopRS=Nothing
						   Response.Redirect "../plus/error.asp?action=error&message=" & server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
						End If   
						   Subject=ShopRS("Subject")
						   Price_Original=ShopRS("Price_Original")
						   Price=ShopRS("Price")
						   Discount=ShopRS("Discount")
						   ActiveDate=ShopRS("ActiveDate")
						   AddDate=ShopRS("AddDate")
						   Intro=ShopRS("Intro")
						   PhotoUrl=ShopRS("PhotoUrl")
						   BigPhoto=ShopRS("BigPhoto")
						   Highlights=ShopRS("Highlights")
						   Protection=ShopRS("Protection")
						   ClassID=ShopRS("ClassID")
						   Notes=ShopRS("Notes")
						   Locked=ShopRS("Locked")
						   EndTF=ShopRS("EndTF")
						   Comment=ShopRS("Comment")
						   AllowArrGroupID=ShopRS("AllowArrGroupID")
						   AllowBMFlag=ShopRS("AllowBMFlag")
						   minnum=ShopRS("minnum")
						   limitbuynum=ShopRS("limitbuynum")
						   Weight=ShopRS("Weight")
						   recommend=ShopRS("recommend")
						   ProvinceID=ShopRS("ProvinceID")
						   CityID=ShopRS("CityID")
						   HasBuyNum=ShopRS("HasBuyNum")
						   MustPayOnline=ShopRS("MustPayOnline")
						   CleanCart=ShopRS("CleanCart")
						   showdelivery=ShopRS("showdelivery")
						   Changes=KS.ChkClng(ShopRS("Changes"))
						   ChangesUrl=ShopRS("ChangesUrl")
						'ProductType=1:Discount=9:Hits = 0:TotalNum = 1000: AlarmNum = 10:Comment = 1
                   End If
				   SelButton=KS.C_C(ClassID,1)
				Else
				 Call KSUser.CheckMoney(ChannelID)
				 CurrentOpStr=" OK,添加 "
				 Action="AddSave"
				 ProductType=1 : Weight=0
				 ShowOnSpace=1
				 MustPayOnline=1
				 limitbuynum=1
				 minnum=1
				 HasBuyNum=0
				 AddDate=Now
				 ActiveDate=Now+10
				 ClassID=KS.S("ClassID")
				 If ClassID="" Then ClassID="0"
				  SelButton="选择栏目..."
				End IF	
				Response.write EchoUeditorHead
				Response.Write "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>"&vbcrlf
		%>
			<script language = "JavaScript">
			function displaydiscount(){
			 if (document.myform.ProductType[2].checked==true)
			   $("#discountarea").show();
			 else
			   $("#discountarea").hide();
			}
			function getprice(Price_Original){
			  if(Price_Original==''|| isNaN(Price_Original)){Price_Original=0;}
			  if(document.myform.ProductType[2].checked==true){
			  document.myform.Price.value=Math.round(Price_Original*Math.abs(document.myform.Discount.value/10)*100)/100;}
		
			  else{document.myform.Price.value=Price_Original;}
			}
			function regInput(obj, reg, inputStr)
			{
				var docSel = document.selection.createRange()
				if (docSel.parentElement().tagName != "INPUT")    return false
				oSel = docSel.duplicate()
				oSel.text = ""
				var srcRange = obj.createTextRange()
				oSel.setEndPoint("StartToStart", srcRange)
				var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
				return reg.test(str)
			}
			function insertHTMLToEditor(codeStr) 
			{ 
			editor.execCommand('insertHtml', codeStr);
			} 
		


function CheckForm()
{
	if ($('#Subject').val()=='')
	{
	 alert('请输入团购主题!');
	 $("#Subject").focus();
	 return false;
	}
	if ($('#ClassID').val()=='')
	{
	 alert('请选择团购分类!');
	 $("#ClassID").focus();
	 return false;
	}

	if (editor.hasContents()==false)
	{
	 alert('请输入本单详情!');
	 editor.focus();
	 return false;
	}
	if ($("#Price_Original").val()=='')
	{
	 alert('请输入原价!');
	 $("#Price_Original").focus();
	 return false;
	}
	if ($("#Discount").val()=='')
	{
	 alert('请输入折扣！');
	 $("#Discount").focus();
	 return false;
	}
	if (parseFloat($("#Discount").val())>10){
	 alert('折扣不能大于10！');
	 $("#Discount").focus();
	 return false;
	}
	if ($("#Price").val()=='')
	{
	 alert('请输入团购价！');
	 $("#Price").focus();
	 return false;
	}
  document.myform.submit();
}
function regInput(obj, reg, inputStr)
{
		var docSel = document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")    return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange = obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
}
function getprice(discount){
     if (parseFloat(discount)>10){
	 alert('折扣不能大于10！');
	 $("#Discount").val(10);
	 return false;
	 }
     var Price_Original=$("#Price_Original").val();
	 if(Price_Original==''|| isNaN(Price_Original)){Price_Original=0;}
	 document.myform.Price.value=Math.round(Price_Original*(discount/10));
  }
$(document).ready(function(){
 if ($("#Changes").prop('checked')){ChangesNews();}
});
function ChangesNews(){ 
			 if ($("#Changes").prop('checked'))
			  {
			  $("#ChangesUrl").attr("disabled",false);
			  }
			  else
			   {
			  $("#ChangesUrl").attr("disabled",true);
			   }
}
</script>
                <iframe src="about:blank" name="hidframe" style="display:none"></iframe> 
                  <form  action="User_groupbuy.asp?Action=<%=Action%>" method="post" name="myform" id="myform">
                    <input type="hidden" value="<%=ID%>" name="id">
                    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
				  <tr class="title">
				   <td colspan=3>
							 <%IF KS.S("Action")="Edit" Then
							   response.write "修改团购"
							   Else
							    response.write "发布团购"
							   End iF
							  %>				   </td>
				  </tr> 
          <tr class="tdbg" >
            <td   height='30' align='right' class='clefttitle'><strong>团购主题：</strong></td>
            <td  height='30'>&nbsp;<input class='textbox' type='text' name='Subject' id='Subject' value='<%=Subject%>' size="40"> <font color=red>*</font></td>
           <td width="243" rowspan="7" style="text-align:center"><div  style="margin:0 auto;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:100px;width:95px;border:1px solid #777777">
						<img src="<%=PhotoUrl%>" onerror="this.src='../images/logo.png';" id="pic" style="height:100px;width:95px;">
		    </div></td>
          </tr>  
          <tr class="tdbg">
            <td height='30' align='right' class='clefttitle'><strong>团购分类：</strong></td>
            <td height='30'>&nbsp;<select name="ClassID" class="ClassID select">
			<option value='0'>---选择分类---</option>
			<%Dim RSC:Set RSC=Conn.Execute("select * From KS_GroupBuyClass Order By OrderID,ID")
			Do While Not RSC.Eof
			  If KS.ChkClng(ClassID)=RSC("ID") Then
			   Response.Write "<option value='" & RSC("ID") & "' selected>" & RSC("CategoryName") & "</option>"
			  Else
			   Response.Write "<option value='" & RSC("ID") & "'>" & RSC("CategoryName") & "</option>"
			  End If
			  RSC.MoveNext
			Loop
			RSC.Close
			Set RSC=Nothing
			%>
			</select></td>
          </tr>
		  <tr class="tdbg"  id='ContentLink'>
		     <td class='clefttitle'><div align='right'><strong>外部链接:</strong></div></td>
			 <td>
				&nbsp;<%
				If ChangesUrl = "" Then
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='35' class='textbox'>")
				Else
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl & "' size='35' class='textbox'>")
				End If
				If Changes = 1 Then
				 Response.Write ("<input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
				Else
				 Response.Write ("<input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'> 使用转向链接</font>")
				End If
			 %>
	   </td></tr>
		  
          <tr style="display:none" class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>地区：</strong></td>
            <td height='30'>&nbsp;<script src="../plus/area.asp?flag=getid"></script> <span style='color:red'>tips:地区不选择的话该团购切换所有地区都会显示</span>
			<script type="text/javascript">
			<%if KS.ChkClng(ProvinceID)<>0 then%>
				  $('#Province').val('<%=provinceid%>');
			<%end if%>
			 <%if KS.ChkClng(CityID)<>0 Then%>
				$('#City').val(<%=CityID%>);
			<%end if%>
			</script>
			</td>
          </tr> 


          <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>时间设置：</strong></td>
            <td height='30'>开始：<input type='text' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" class='textbox' name='AddDate' value='<%=AddDate%>' size="20" /> <br /><br />
			结束：<input type='text' class='textbox' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" name='ActiveDate' value='<%=ActiveDate%>' size="20" /><br/><span class='tips'>如<%=year(now)%>-<%=month(now)%>-<%=day(now)%> 10:10</span> </td>
          </tr> 

		  
          <tr class="tdbg" >
            <td height='30' align='right' class='clefttitle'><strong>购物车设置：</strong></td>
            <td height='30'>
			&nbsp;需要在线支付订单才生效：<label><input type="radio" name="MustPayOnline" value="0"<%if MustPayOnline="0" then response.write " checked"%>/>不需要</label>
			<label><input type="radio" name="MustPayOnline" value="1"<%if MustPayOnline="1" then response.write " checked"%>/>需要</label>
          <div style="display:none">
			&nbsp;当购物车里有商品时先清空：<label><input type="radio" onClick="$('#delivery').show();" name="cleancart" value="1" checked/>是</label>
			<label><input type="radio" name="cleancart" onClick="$('#delivery').hide();" value="0"/>否</label>
			<br/>&nbsp;<span class="tips">当选择购物车里有商品时先清空，则订单里只能有这件商品。</span>
            </div>
			<%if cleancart="1" then%>
			<div style="" id="delivery">
			<%else%>
			<div style="display:none" id="delivery">
			<%end if%>
            
			&nbsp;显示送货方式：<label><input type="radio" name="showdelivery" value="1"<%if showdelivery="1" then response.write " checked"%>/>显示</label><label><input type="radio" name="showdelivery" value="0"<%if showdelivery="0" then response.write " checked"%>/>不显示</label>
			<br/>
			 &nbsp;<span class="tips">如本地商家打折等团购建议选择不显示</span>
             
			</div>
			</td>
          </tr>
		   <tr class="tdbg" >
		     <td height='30' align='right' class='clefttitle'><strong>商品图片：</strong></td>
		     <td height='30'>小图：<input class="textbox"  type="text" name="PhotoUrl" id="PhotoUrl" size="30" value="<%=photourl%>" /> <input class="button" type='button' name='Submit' value='选择小图...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&groupbuy=ok&Photo_id=PhotoUrl&pagetitle=<%=Server.URLEncode("选择图片")%>&channelid=5',500,360,window,document.myform.PhotoUrl);">
			 <br/>
			 大图：<input value="<%=bigphoto%>" class="textbox" type="text" name='BigPhoto' id='BigPhoto' size="30" /> <input class="button" type='button' name='Submit' value='选择大图...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&groupbuy=ok&Photo_id=BigPhoto&pagetitle=<%=Server.URLEncode("选择图片")%>&channelid=5',500,360,window,document.myform.BigPhoto);">
  </tr>
		   <tr class="tdbg" >
		     <td height='30' align='right' class='clefttitle'><strong>上传图片：</strong></td>
			 <td colspan="2">
             
             <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=5&Type=Pic' frameborder=0 scrolling=no width='95%' height='30'></iframe>
 		 </td>
		</tr>

 
		   <tr class="tdbg" >
		     <td height='30' align='right' class='clefttitle'><strong>价格设置：</strong></td>
		     <td height='30' colspan="2">
             原价<input type="text" class="textbox" style="width:120px;" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" name="Price_Original" id="Price_Original" size="6" value="<%=Price_Original%>" style="text-align:center" />元 
             折扣<input class="textbox" style="width:120px;" onChange="getprice(this.value);" type="text" name="Discount" id="Discount" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" size="6" value="<%=Discount%>" style="text-align:center" />折<br /><br />
             团购价<input type="text" style="width:120px;" name="Price" id="Price" size="6" value="<%=Price%>" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" style="text-align:center" />元
			 
			 重量：<input class="textbox" style="width:120px;" type='text' name='Weight' style="text-align:center" id='Weight' value='<%=Weight%>' size="6">KG<br /><br />
			 <span style='color:#999999'>计算运费用的,包邮请输入-1。</span></td>
  </tr>
		   <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>最低人数：</strong></td>
           <td height='30' colspan="2"><input class="textbox" type='text' name='minnum' style="text-align:center" id='minnum' value='<%=minnum%>' size="6"> 人 </td>
           </tr>
		   <tr class="tdbg" >
               <td  width='143' height='30' align='right' class='clefttitle'><strong>每人限制购买：</strong></td>
               <td height='30' colspan="2"><input class="textbox" type='text' name='limitbuynum' style="text-align:center" id='limitbuynum' value='<%=limitbuynum%>' size="6"> 件 <font color=red>*</font> 不限制输入0
               </td>
            </tr>
		   <tr class="tdbg" >
               <td  width='143' height='30' align='right' class='clefttitle'><strong>初始已销售：</strong></td>
               <td height='30' colspan="2"><input type='text' name='hasbuynum' style="text-align:center" class="textbox" id='hasbuynum' value='<%=hasbuynum%>' size="6"> 件 <span style='color:#999999'>(作弊用的)</span> 
               </td>
            </tr>
            
           
          </tr>  
		   <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>本单详情：</strong></td>
            <td height='30' colspan="3">
			 <%
				 Response.Write "<script id=""Intro"" name=""Intro"" type=""text/plain"" style="" width:90%;height:200px;"">" &KS.ClearBadChr(Intro)&"</script>"
	             Response.Write "<script>setTimeout(""editor = " & GetEditorTag() &".getEditor('Intro',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source', '|',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:200,elementPathEnabled:false });"",10);</script>"
				%>
				</td>
          </tr>
		   <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>精彩卖点：</strong></td>
           <td height='30' colspan="2">&nbsp;<textarea name='Highlights' cols="60" rows="4" class="textbox"><%=Highlights%></textarea></td>
          </tr>  
		   <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>团购保障：</strong></td>
           <td height='30' colspan="2">&nbsp;<textarea name='Protection' cols="60" rows="4" class="textbox"><%=Protection%></textarea></td>
          </tr>  
		  <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>温馨提示：</strong></td>
            <td height='30' colspan="2">&nbsp;<textarea name='Notes' cols="60" rows="4" class="textbox"><%=Notes%></textarea></td>
          </tr>  
		  <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>允许参加团购的权限：</strong></td>
            <td height='30' colspan="2">	    
		    <label><input type="radio" onclick="$('#showgroup').hide()" name="AllowBMFlag" value="0"<%if AllowBMFlag=0 then response.write " checked"%>>允许所有人报名参加,包括游客</label><br/>
			<label><input type="radio" onclick="$('#showgroup').hide()" name="AllowBMFlag" value="1"<%if AllowBMFlag=1 then response.write " checked"%>>只允许会员报名参加</label>
			<br/><label><input type="radio" onclick="$('#showgroup').show()" name="AllowBMFlag" value="2"<%if AllowBMFlag=2 then response.write " checked"%>>只允许指定的会员组报名参加</label>			</td>
          </tr>  
		  <tr class="tdbg" <%if AllowBMFlag<>2 then response.write " style='display:none'"%> id="showgroup" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>允许参加团购的会员组：</strong><br/>
           </td>
           <td height='30' colspan="2" valign="top"><%=KS.GetUserGroup_CheckBox("AllowArrGroupID",AllowArrGroupID,5)%>			</td>
          </tr> 
 
		  <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>是否允许评论：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="comment" value="0"<%if comment=0 then response.write " checked"%>>不允许（关闭）<Br/>
			&nbsp;&nbsp;<input type="radio" name="comment" value="1"<%if comment=1 then response.write " checked"%>>允许，评论内容需要审核<Br/>	
			&nbsp;&nbsp;<input type="radio" name="comment" value="2"<%if comment=2 then response.write " checked"%>>允许，评论不需要审核		   </td>
          </tr>  
		   
		  <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>是否锁定：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="locked" value="0"<%if locked=0 then response.write " checked"%>>否
			<input type="radio" name="locked" value="1"<%if locked=1 then response.write " checked"%>>是		   </td>
          </tr>  
		  <tr class="tdbg" >
            <td  width='143' height='30' align='right' class='clefttitle'><strong>是否结束：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="endtf" value="0"<%if endtf=0 then response.write " checked"%>>否
			<input type="radio" name="endtf" value="1"<%if endtf=1 then response.write " checked"%>>是		   </td>
          </tr>  
                
          <tr class="tdbg">
			 <td></td>
             <td><button class="pn"  onClick="return CheckForm();" id="btn" type="button"><strong><%=CurrentOpStr%></strong></button></td>
          </tr>
        </table>
   </form>
		  <%
		  If IsObject(ShopRS) Then
  			If ShopRS.status<>0 Then  ShopRS.Close:Set ShopRS=Nothing
          End If
  End Sub
  
  Function GetBrandByClassID(ClassID,BrandID)
		  Dim SQL,K
		  Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		  RS.Open "Select B.ID,B.BrandName From KS_ClassBrand B inner join KS_ClassBrandR R On B.id=R.BrandID where R.classid='" & classid & "' order by B.orderid",conn,1,1
		  If Not RS.Eof  Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
		  If Not IsArray(SQL) Then
		   'GetBrandByClassID="Null" 
		  Else
		     GetBrandByClassID = "所属品牌：<select name='brandid'>"
			 GetBrandByClassID = GetBrandByClassID & "<option value='0'>-请选择品牌-</option>"
		     For K=0 To Ubound(SQL,2)
			  If BrandID=SQL(0,K) Then
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & sql(0,k) & "' selected>" & sql(1,k) & "</option>"
			  Else
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & sql(0,k) & "'>" & sql(1,k) & "</option>"
			  End If
			 Next
			 GetBrandByClassID = GetBrandByClassID &  "</select>"
			 Erase Sql
		  End If
  End Function
	   
  Sub ShopSave()
        Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Subject:Subject=KS.LoseHtml(KS.G("Subject"))
       Dim ActiveDate:ActiveDate=KS.G("ActiveDate")
			if not isdate(ActiveDate) then
			 Response.Write "<script>alert('本单载止日期格式不正确！');history.back();</script>"
			 Exit Sub
			End If	  
       Dim AddDate:AddDate=KS.G("AddDate")
			if not isdate(AddDate) then
			 Response.Write "<script>alert('发布时间格式不正确！');history.back();</script>"
			 Exit Sub
		End If	 
			


	   Dim PhotoUrl:PhotoUrl=KS.G("PhotoUrl")
	   Dim BigPhoto:BigPhoto=KS.G("BigPhoto")

			 
	   Dim Intro:Intro=Request.Form("Intro")
	   Dim Fax:Fax=KS.LoseHtml(KS.G("Fax"))
	   Dim Highlights:Highlights=KS.LoseHtml(KS.G("Highlights"))
	   Dim Protection:Protection=KS.LoseHtml(KS.G("Protection"))
	   Dim Notes:Notes=KS.LoseHtml(KS.G("Notes"))
	   Dim Locked:Locked=KS.ChkClng(KS.G("Locked"))
	   Dim EndTF:EndTF=KS.ChkClng(KS.G("EndTf"))
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   Dim ClassID:ClassID=KS.ChkClng(KS.G("ClassID"))
	   Dim AllowBMFlag:AllowBMFlag=KS.ChkClng(KS.G("AllowBMFlag"))
	   Dim minnum:minnum=KS.ChkClng(KS.G("minnum"))
	   Dim AllowArrGroupID:AllowArrGroupID=KS.G("AllowArrGroupID")
	   Dim Price_Original:Price_Original=KS.G("Price_Original")
	   Dim Discount:Discount=KS.G("Discount")
	   Dim Price:Price=KS.G("Price")
	   Dim Weight:Weight=KS.G("Weight")
	   If Not IsNumeric(Weight) Then Weight=0
	   Dim LimitBuyNum:LimitBuyNum=KS.ChkCLng(KS.G("LimitBuyNum"))
	   Dim ProvinceID:ProvinceID=KS.ChkClng(KS.G("province"))
	   Dim CityID:CityID=KS.ChkClng(KS.G("city"))
	   Dim HasBuyNum:HasBuyNum=KS.ChkClng(KS.G("hasbuynum"))
	   Dim MustPayOnline:MustPayOnline=KS.ChkClng(KS.G("MustPayOnline"))
	   Dim CleanCart:CleanCart=KS.ChkClng(KS.G("CleanCart"))
	   Dim Comment:Comment=KS.ChkClng(KS.G("Comment"))
	   Dim showdelivery:showdelivery=KS.ChkClng(KS.G("showdelivery"))
	   dim Verific:Verific=KS.ChkClng(KS.G("Verific"))
	   If KS.IsNul(Subject) Then KS.Die "<script>alert('团购主题必须输入!');history.back();</script>"
	   If not isnumeric(Price_Original) Then KS.Die "<script>alert('原价必须输入正确的数字!');history.back();</script>"
	   If not isnumeric(Discount) Then KS.Die "<script>alert('折扣必须输入正确的数字!');history.back();</script>"
	   If not isnumeric(Price) Then KS.Die "<script>alert('团购价必须输入正确的数字!');history.back();</script>"			  
	   Set RSObj=Server.CreateObject("Adodb.Recordset")
		
			RSObj.Open "Select top 1 * From KS_GroupBuy Where username='" & KSUser.UserName & "' and ID=" & ID,Conn,1,3
				If RSObj.Eof And RSObj.Bof Then
				   RSObj.AddNew
				   RSObj("IsSuccess")=0
				   RSObj("PostTable")= LFCls.GetCommentTable()
				   RSObj("CmtNum") = 0
				   RSObj("username") =KSUser.UserName
				   RSObj("Verific") =Verific
				   RSObj("recommend")=0
				End If
					 RSObj("AddDate")=AddDate
					 RSObj("Subject")=Subject
					 RSObj("ActiveDate")=ActiveDate
					 RSObj("Intro")=Intro
					 RSObj("PhotoUrl")=PhotoUrl
					 RSObj("BigPhoto")=BigPhoto
					 RSObj("Highlights")=Highlights
					 RSObj("Protection")=Protection
					 RSObj("ClassID")=ClassID
					 RSObj("Notes")=Notes
					 RSObj("Locked")=Locked
					 RSObj("EndTF")=EndTF
					 RSObj("minnum")=minnum
					 RSObj("LimitBuyNum")=LimitBuyNum
					 RSObj("Weight")=Weight
					 RSObj("AllowBMFlag")=AllowBMFlag
					 RSObj("AllowArrGroupID")=AllowArrGroupID
					 RSObj("Price_Original")=Price_Original
					 RSObj("Discount")=Discount
					 RSObj("Price")=Price
					 
					 RSObj("HasBuyNum")=HasBuyNum
					 RSObj("MustPayOnline")=MustPayOnline
					 RSObj("CleanCart")=CleanCart
					 RSObj("Comment")=Comment
					 RSObj("showdelivery")=showdelivery
					 RSObj("ProvinceID")=ProvinceID
					 RSObj("CityID")=CityID
					 RSObj("Changes")=KS.ChkClng(request("changes"))
					 RSObj("ChangesUrl")=Request.Form("ChangesUrl")
					 
					 if RSObj("Verific")=3 then
					  RSObj("Verific")=0
					 end if
				RSObj.Update
				If ID=0 Then
				   RSObj.MoveLast
                   Call KS.FileAssociation(1005,RSObj("ID"),Intro&RSObj("PhotoUrl"),0)
				 Else
                   Call KS.FileAssociation(1005,ID,Intro&RSObj("PhotoUrl"),1)
				 End If
				 RSObj.Close:Set RSObj=Nothing
              If ID=0 Then
				 KS.Echo "<script>if (confirm('团购信息添加成功，继续添加吗?')){top.location.href='User_groupbuy.asp?Action=Add&ClassID=" & ClassID &"';}else{top.location.href='User_groupbuy.asp';}</script>"
			  Else
				KS.Echo "<script>alert('团购信息修改成功!');top.location.href='" & ComeUrl & "';</script>"
			  End If
		
  End Sub
  
  
  
  Sub ShowOrder()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * from ks_order where username='" & KSUser.UserName & "' and id=" & ID ,conn,1,1
		 IF RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   response.end
		 End If

		response.write OrderDetailStr(RS)

	
	
	if rs("isservice")="1" then%>
		<a name="service"></a><strong> 服务记录明细：<br/></strong>
		 <%
		         dim times,sytimes,validity,firstservicetime
				  times=conn.execute("select count(1) from ks_orderservice where orderid=" & rs("id"))(0)
				  if times>rs("servicetimes") then sytimes=0 else sytimes=rs("servicetimes")-times
				  dim rsi:set rsi=conn.execute("select top 1 adddate from ks_orderservice where orderid=" & rs("id"))
				  if not rsi.eof then
					firstservicetime=rsi(0)
					validity=dateadd("m",rs("validity"),firstservicetime)
				  else
					validity=dateadd("m",rs("validity"),now)
				  end if
				  rsi.close
				  set rsi=nothing
				  %>
				   <div style="border:#B2D9F6 1px solid; line-height:26px;padding-left:5px;margin-bottom:10px;background:#F3F9FF;">
				  
				  服务商品名称：<%=rs("servicename")%>&nbsp;&nbsp;服务次数：<%=rs("servicetimes")%>次,剩余：<font color=red><%=sytimes%></font>次&nbsp;服务有效期：<%=rs("validity")%>个月,载止日期：<%=year(validity) & "-" & month(validity) & "-" & day(validity)%>
				   
				   </div>
				 <table  cellpadding="1" style="margin-bottom:6px;border:1px solid #999;" cellspacing="1" width="100%">

				   <tr style="background:#f1f1f1;height:23px;text-align:center">
					  <td width="50">次数</td>
					  <td width="350">内容</td>
					  <td width="70">时间</td>
					  <td width="70">签收人</td>
				   </tr>
				   <%
				   dim rss:set rss=server.CreateObject("adodb.recordset")
				   RSS.Open "select * from ks_orderservice where orderid=" & rs("id") & " order by id desc",conn,1,1
				   if RSS.Eof Then
					str="<tr><td colspan=4 class=""splittd"">没有找到服务记录！</td></tr>"
				   Else
					dim totalnum:totalnum=rss.recordcount
					dim str,num,qsr
					num=0
					if totalnum<5 then
					str="<tr><td colspan=4><div>"
					else
					str="<tr><td colspan=4><div style=""overflow-x:hidden;overflow-y:auto;height:130px"">"
					end if
					do while not rss.eof
					str=str &"<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
					str=str &"<tr id='tr1" & rss("id") & "'>"
					str=str &"<td width=""50"" height=""25"" class=""splittd"">第" & totalnum-num & "次</td>"
					str=str &"<td width=""350"" class=""splittd"" style=""width:290px;word-break:break-all;"">" & rss("content") & "</td>"
					str=str &"<td width=""70"" class=""splittd"" style='text-align:center'>" & year(rss("adddate")) & "-" & month(rss("adddate")) & "-" & day(rss("adddate")) & "</td>"
					qsr=rss("qsr")
					if ks.isnul(qsr) then qsr="---"
					str=str &"<td width=""70"" class=""splittd"" style='text-align:center'>&nbsp;" & qsr & "&nbsp;</td>"
					str=str &"</tr></table>"
					num=num+1
					rss.movenext
					loop
					str=str & "</div></td></tr>"
				  end if
				  rss.close
					response.write str
					
				  %>
				 </table>
		<%
		end if
		 rs.close:set rs=nothing
		End Sub
  
  '删除
	Sub GroupBuyDel()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Delete From KS_UploadFiles Where UserName='"& KSUser.UserName &"' and  ChannelID=1005 and InfoID In("& id & ")")
	 Conn.execute("Delete From KS_GroupBuy Where UserName='"& KSUser.UserName &"' and   id In("& id & ")")
	 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	'取得捆绑销售商品
	Dim ProIds
	Function GetBundleSalePro(ByRef TotalPrice,ProID,OrderID)
	  If KS.FoundInArr(ProIDS,ProID,",")=true Then Exit Function
	  ProIds=ProIDs & "," & ProID
	  Dim Str,RS,XML,Node
	  Set RS=Server.CreateObject("adodb.recordset")
	  RS.Open "Select I.Title,I.Unit,O.* From KS_OrderItem O inner join KS_Product I On O.ProID=I.ID Where O.SaleType=6 and BundleSaleProID=" & ProID & " and OrderID='" & OrderID & "' order by O.id",conn,1,1
	  If Not RS.Eof Then
		Set XML=KS.RsToXml(rs,"row","")
	  End If
	  RS.Close:Set RS=Nothing
	  If IsObject(XML) Then
			 str=str & "<tr height=""25"" align=""left""><td></td><td colspan=9 style=""font-weight:bold"">选购捆绑促销:</td></tr>"
		   For Each Node In Xml.DocumentElement.SelectNodes("row")
			 str=str & "<tr>"
			 str=str &" <td style='color:#999999'></td>"
			 str=str &" <td style='color:#999999'>&nbsp;" & Node.SelectSingleNode("@title").text &" (原价：" & formatnumber(Node.SelectSingleNode("@price_original").text,2,-1) &")" &"</td>"
			 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
			 str=str &" <td align='center'>" & Node.SelectSingleNode("@amount").text & " " & Node.SelectSingleNode("@unit").text & "</td>"
			 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2,-1) &"</td>"
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
							  str=str & "<tr class='tdbg'  align=""center""><td></td><td align='left'><strong><a href='" & DomainStr & "shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a></strong></td>"
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
								  Else
									Price=RSS("Price_member")
								  End If
								  RSAttr.CLose:Set RSAttr=Nothing
								 Else
									Price=RSS("Price_member")
								 End If
								
								   TotalPackPrice=TotalPackPrice+Price
								  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
								  n=n+1
							  Next
							  RSS.MoveNext
							Loop
							
							str=str &"<td>￥" & TotalPackPrice & "</td><td>1</td><td>￥" & TotalPackPrice & "</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) &"</td><td>---</td>"
						   
							str=str & "</tr><tr><td></td><td align='left' colspan=9 valign='top' style='color:#999'>您选择的套装详细如下:<br/>" & tempstr & "</td></tr>" 
							
							TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1))   '将礼包金额加入总价
							
							RSS.Close
							Set RSS=Nothing
						
				End If
				RSB.Close
			   Next
				
			End If
			GetPackage=str
			
	End Function
	 '得到订单的状态		
  function GetOrderStatus(rs)
    dim str:str=""
           if rs("status")=2 then
				   str="<span style='color:#999'>已结清</span>"
				   If RS("DeliverStatus")=3 Then
					 str=str & " <font color=#ff6600>退货</font>"
				   end if
				else
				if rs("alipaytradestatus")<>"" and RS("Status")<>2 then
				  select case rs("alipaytradestatus")
				    Case "WAIT_BUYER_PAY" str=str & "<font color=red>未付款</font>"
					Case "WAIT_SELLER_SEND_GOODS" str=str & "<font color=brown>已付款等待发货</font>"
					Case "WAIT_BUYER_CONFIRM_GOODS" str=str & "<font color=blue>等待买家确认收货</font>"
					Case "TRADE_FINISHED" str=str & "<font color=#a7a7a7>交易完成</font>"
				  end select
				else
					if rs("paystatus")="100" then
					  str=str & "<font color=""green"">凭单消费</font>"
					elseif rs("paystatus")="3" then
					  str=str & "<font color=blue>退款</font>"
					elseIf RS("MoneyReceipt")<=0 Then
					  str=str & "<font color=red>未付款</font>"
					ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
					 str=str & "<font color=blue>已收定金</font>"
					Else
					 str=str & "<font color=green>已经付清</font>"
					End If
				end if	  
			    str=str & "/"
		 	   If RS("DeliverStatus")=0 Then
		 		 str=str & "<font color=red>未发货</font>"
			   ElseIf RS("DeliverStatus")=1 Then
				 str=str & "<font color=blue>已发货</font>"
			   ElseIf RS("DeliverStatus")=2 Then
				 str=str & "<font color=green>已签收</font>"
			   ElseIf RS("DeliverStatus")=3 Then
					   str=str & "<font color=#ff6600>退货</font>"
			   ElseIf RS("DeliverStatus")=4 Then
					str=str & "<font color=brown>已申请退货退款</font>"
			  End If
  	end if
	GetOrderStatus=str
end function

'返回订单详细信息
		Function  OrderDetailStr(RS)
		 OrderDetailStr="<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr><td class='title' style='padding:3px'>订单信息：</td></tr>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr><td class='bgtitle' style='padding-left:10px'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "订单编号：" & RS("ORDERID") & " (购买日期：" & formatdatetime(rs("inputtime"),2) &")<br/>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "订单状态：" & GetOrderStatus(rs) & "<br/>"&vbcrlf
		 
		 if KS.ChkClng(rs("UseScoreisshop"))>0 then
					OrderDetailStr=OrderDetailStr & "  <font color=""#006600""  style=""font-size:14px"">本单为积分兑换订单，已支付积分<font color=""#FF0000"">"& KS.ChkClng(rs("UseScoreisshop")) & "</font> 积分</font><br/>"  
		 end if
		 OrderDetailStr=OrderDetailStr & "<span style=""font-weight:bold;font-size:14px"">订单金额￥" & formatnumber(rs("NoUseCouponMoney"),2) & " 元 "
		 if rs("Charge_Deliver")>0 then OrderDetailStr=OrderDetailStr & "<span style='color:#999;font-weight:normal'>(含运费" & rs("Charge_Deliver")&" 元)</span>"
	If KS.ChkClng(RS("CouponUserID"))<>0 and RS("UseCouponMoney")>0 Then
	OrderDetailStr=OrderDetailStr & "使用优惠券 <font color=#ff6600>￥" & formatnumber(RS("UseCouponMoney"),2) & " 元</font><br>"
    ElseIf RS("UseScoreMoney")<>"0" Then
	OrderDetailStr=OrderDetailStr & "花费<font color=green>" &RS("UseScore") & "</font>积分抵扣了<font color=#ff6600>" & formatnumber(RS("UseScoreMoney"),2) & "</font>元 "
	End If
	OrderDetailStr=OrderDetailStr & " 应付￥" & formatnumber(rs("MoneyTotal"),2) & "元 已付￥<font color=green>" & formatnumber(rs("MoneyReceipt"),2) & "</font>元"
	If RS("MoneyReceipt")<RS("MoneyTotal") Then
	OrderDetailStr=OrderDetailStr & " 尚欠￥<font color=red>" & formatnumber(RS("MoneyTotal")-RS("MoneyReceipt"),2) &"</font>元"
	End If
		 
		 
		 
		 OrderDetailStr=OrderDetailStr & " 获赠积分：" 
		if rs("totalscore")=0  or rs("DeliverStatus")=3 then
			OrderDetailStr=OrderDetailStr & "无"
		else
			if rs("scoretf")=1 then
			OrderDetailStr=OrderDetailStr & "<font color=green>" & rs("totalscore") & "分,已送出</font>"
			else
			OrderDetailStr=OrderDetailStr & "<font color=red>" & rs("totalscore") & "分,未送出</font>"
			end if
		end if
		OrderDetailStr=OrderDetailStr & "</div>"
		
		 
		
		 OrderDetailStr=OrderDetailStr & "</td><tr><td class='title' style='padding:3px'>收货信息：</td></tr>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr><td class='bgtitle' style='padding-left:10px'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "收 货 人：" & rs("contactman") & "<br/>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "联系电话：" & rs("phone") & " " & rs("mobile")& "<br/>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "收货地址：" & rs("address") & "（邮编：" & rs("zipcode") &"）<br/>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "快递公司：" 
	if rs("tocity")="" then
    OrderDetailStr=OrderDetailStr & "免运费订单，由商家指定" 
	else
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
	
	
	OrderDetailStr=OrderDetailStr & " 发往<span style='color:red'>" & rs("tocity") & "</span>"
	end if
	
	 OrderDetailStr=OrderDetailStr & "<br/><table cellspacing='0' cellpadding='0' border='0'><tr><td>发票信息：</td><td>" 
If RS("NeedInvoice")=0 Then
	  OrderDetailStr=OrderDetailStr & "不需要"
	ElseIf RS("NeedInvoice")=1 Then
	  OrderDetailStr=OrderDetailStr & "发票类型：普通发票"
	  If RS("Invoiced")=1 Then OrderDetailStr=OrderDetailStr & " <font color=green>已开</font>" Else OrderDetailStr=OrderDetailStr & " <font color=red>未开</font>"
	  OrderDetailStr=OrderDetailStr & "<br/>单位名称：" &rs("InvoiceContent")
	Else
	  OrderDetailStr=OrderDetailStr & "发票类型：增值税发票"
	  If RS("Invoiced")=1 Then OrderDetailStr=OrderDetailStr & " <font color=green>已开</font>" Else OrderDetailStr=OrderDetailStr & " <font color=red>未开</font>"
	  OrderDetailStr=OrderDetailStr & "<br/>单位名称：" &rs("InvoiceContent") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "纳税人识别码："&rs("InvoiceCode") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "注册地址："&rs("InvoiceAddress") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "注册电话："&rs("InvoiceTel") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "开户银行："&rs("Invoicebank") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "银行账号："&rs("Invoicebankcard")
	End If
	 OrderDetailStr=OrderDetailStr & "</td></tr></table>"
	  if not ks.isnul(rs("Remark")) then
	 OrderDetailStr=OrderDetailStr & "<br/>备注/留言：" & rs("Remark")
	  end if


	OrderDetailStr=OrderDetailStr & "		<tr><td>"
	OrderDetailStr=OrderDetailStr & "		<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "
	OrderDetailStr=OrderDetailStr & "		  <tr align='center' class='title' height='25'>  "  
	OrderDetailStr=OrderDetailStr & "		   <td nowrap><b>编号</b></td> "   
	OrderDetailStr=OrderDetailStr & "		   <td><b>商品名称</b></td> "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>您的价格</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>数量</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>金 额</b></td>   " 
	OrderDetailStr=OrderDetailStr & "		   <td><b>赠送积分</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td nowrap><b>备注</b></td>  "
	OrderDetailStr=OrderDetailStr & "		  </tr> "
			 Dim TotalPrice,attributecart,RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			   RSI.Open "Select * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id",conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
			  Else
			  Do While Not RSI.Eof
			  attributecart=rsi("attributecart")
			  if not ks.isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"    
		OrderDetailStr=OrderDetailStr & "	   <td>" & rsi("proid") &"</td>" 
		OrderDetailStr=OrderDetailStr & "	   <td>" 
		 Dim OrderType:OrderType=KS.ChkClng(rs("ordertype"))
		 If OrderType=1 Then
		  OrderDetailStr=OrderDetailStr & "<a href='" & DomainStr & "shop/groupbuyshow.asp?id=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 subject from ks_groupbuy where id=" & rsi("proid"))(0)
		 Else
		  OrderDetailStr=OrderDetailStr & "<a href='" & DomainStr & "item/show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 title from ks_product where id=" & rsi("proid"))(0) 
		 End If
		If RSI("IsChangedBuy")="1" Then OrderDetailStr=OrderDetailStr & "(换购)"
		
		
			  Dim SqlStr,RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  If OrderType=1 Then
			    SqlStr="Select top 1 Subject as title,'件' as unit,0 as IsLimitBuy,0 as LimitBuyPrice,0 as LimitBuyPayTime From KS_GroupBuy Where ID=" & RSI("ProID")
			  Else
			    SqlStr="Select top 1 I.Title,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID")
			  End If
			  RSP.Open SqlStr,conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(换购)"
				  else
				    if RSP("LimitBuyPayTime") then
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
		
		OrderDetailStr=OrderDetailStr & "</a>" & attributecart & "" & "(参考价：" & formatnumber(rsi("price_original"),2) &"元 商城价：" & formatnumber(rsi("price"),2) & "元)"
		OrderDetailStr=OrderDetailStr & "</td><td width='65' align='center'>" & formatnumber(rsi("realprice"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "<td width='55' align='center'>" & rsi("amount") &""& unit & "</td>    "
		OrderDetailStr=OrderDetailStr & "<td width='85' align='center'>" & formatnumber(rsi("realprice")*rsi("amount"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "<td width='65' align=center>" & ks.chkclng(rsi("score")*rsi("amount")) & " 分</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td align=center'>" 
		totalscore=totalscore+ks.chkclng(rsi("score")*rsi("amount"))
		Set RSP=Conn.Execute("Select Top 1 DownUrl From KS_Product Where ID=" & RSI("ProID"))
		If Not RSP.Eof Then
			If Not KS.IsNul(RSP("DownUrl")) Then
				If RS("MoneyReceipt")>=RS("MoneyTotal") Then
				  OrderDetailStr=OrderDetailStr & "<a href='?action=OrderDown&orderid=" & rs("id") & "&proid=" & rsi("proid") &"'><img src='../images/default/download.gif'></a>"
				Else
				 OrderDetailStr=OrderDetailStr & "<a href='#' disabled>未付清</a>"
				End If
			Else
				 OrderDetailStr=OrderDetailStr & "---"
			End If
		Else
		  OrderDetailStr=OrderDetailStr & "---"
		End If
		RSP.Close :Set RSP=Nothing
		
		OrderDetailStr=OrderDetailStr & "</td>  "
		OrderDetailStr=OrderDetailStr & "	   </tr> " 
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  '取得捆绑销售商品
		
		
			  TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
		End If
		
		OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '超值礼包
		
		
	OrderDetailStr=OrderDetailStr & "</table></td>  "
	OrderDetailStr=OrderDetailStr & "</tr>"  
	OrderDetailStr=OrderDetailStr & "     <tr><td><br><b>注：</b><br/>1、“<font color='blue'>参考价</font>”指商品的市场参考价，“<font color='green'>商城价</font>”指本商城的销售价格，“<font color='red'>您的价格</font>”指根据会员灯级折扣系统自动算出的优惠价。商品的最终销售价格以“您的价格”为准。“订单金额”指系统自动算出来的价格，本订单的最终价格以“<font color=#ff6600>应付金额</font>”为准。<br>2、积分的赠送为结清订单后送出；<br/>3、限时抢购商品、特惠礼包、使用优惠券及使用积分抵扣费用的订单将不再赠送积分。"
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
		
End Class
%> 

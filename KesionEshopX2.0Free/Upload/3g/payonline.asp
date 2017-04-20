<!--#include file="../user/payfunction.asp"-->
<%
	   
Sub PayOnline()
	    %>
		<div class="tabs">	
			<ul class="clearfix">
				<li<%if ks.s("flag")="" then response.write" class='puton'"%>><a href="?action=payonline">在线充值</a></li>
				<li<%if ks.s("flag")="card" then response.write" class='puton'"%>><a href="?action=payonline&flag=card">充值卡充值</a></li>
			</ul>
        </div>
		<%
		if ks.s("flag")="cardsave" then
		  call cardsave()
		elseif ks.s("flag")="card" then
		 '充值卡充值
		 %>
		 <script type="text/javascript">
	     function Confirm(){
		  if (document.myform.CardNum.value==""){
		   $.dialog.alert('请输入充值卡卡号!',function(){
		   document.myform.CardNum.focus();});
		   return false;
		  }
		  if (document.myform.CardPass.value==""){
		   $.dialog.alert('请输入充值卡密码!',function(){
		   document.myform.CardPass.focus();});
		   return false;
		  }
		  return true;
		  }
	   </script>
	   <div class="tableBGw2">
		<FORM name=myform action="user.asp" method="post">
		  <table class=border cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
			<tr class=tdbg>
			  <td align=right width=120>用户名：</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right>计费方式：</td>
			  <td><%if KSUser.ChargeType=1 Then 
		  Response.Write "扣点数</font>计费用户"
		  ElseIf KSUser.ChargeType=2 Then
		   Response.Write "有效期</font>计费用户,到期时间：" & cdate(KSUser.GetUserInfo("BeginDate"))+KSUser.GetUserInfo("Edays") & ","
		  ElseIf KSUser.ChargeType=3 Then
		   Response.Write "无限期</font>计费用户"
		  End If
		  %>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=120>资金余额：</td>
			  <td><input type='hidden' value='<%=KSUser.GetUserInfo("Money")%>' name='Premoney'><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1,0,-1)%> 元</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用<%=KS.Setting(45)%>：</td>
			  <td><%=formatnumber(KSUser.GetUserInfo("Point"),2,-1,0,-1)%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>剩余天数：</td>
			  <td>
			  <%if KSUser.ChargeType=3 Then%>
			  无限期
			  <%else%>
			  <%=KSUser.GetEdays%>&nbsp;天
			  <%end if%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right>充值卡卡号：</td>
			  <td>&nbsp;<input name="CardNum" type="text" class="textbox" size="25" maxlength="50"></td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=120>充值卡密码：</td>
			  <td>&nbsp;<input name="CardPass" type="text" class="textbox" size="25" maxlength="50"></td>
			</tr>
			<tr class=tdbg>
			  <td align=left height=40></td>
			  <td align=left height=40>
			  <Input  type=hidden value="payonline" name="Action"> 
			  <Input type=hidden value="cardsave" name="flag"> 
				<Input class="button" id=Submit type=submit value="确定充值" onClick="return(Confirm())" name=Submit></td>
			</tr>
		  </table>
		</FORM>
		</div>
		<%else %>
		   <script type="text/javascript">
			 function Confirm(v)
			 {
			  $("#paytype").val(v);
			  if (v==1){
				return(confirm('此操作不可逆，确定使用余额支付购买吗？'));
			  }
			  if (document.myform.Money.value=="")
			  {
			   alert('请输入你要充值的金额!')
			   document.myform.Money.focus();
			   return false;
			  }
			  return true;
			  }
		   </script>
		   <div class="tableBGw2">
			<FORM name=myform action="user.asp" method="post">
			  <table class=border cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
				<tr class="tdbg">
				  <td align=right  nowrap="nowrap" style="min-width:6rem;">用户名：</td>
				  <td><%=KSUser.UserName%></td>
				</tr>
				<tr class="tdbg">
				  <td align=right  nowrap="nowrap" style="min-width:6rem;">余额：</td>
				  <td><input type='hidden' value='<%=KSUser.GetUserInfo("Money")%>' name='Premoney'><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1)%> 元</td>
				</tr>
				<%If KSUser.ChargeType=1 then%>
				<tr class=tdbg>
				  <td align=right  nowrap="nowrap" style="min-width:6rem;">可用<%=KS.Setting(45)%>：</td>
				  <td><%=KSUser.GetUserInfo("Point")%>&nbsp;<%=KS.Setting(46)%></td>
				</tr>
				<%end if%>
				<%If KSUser.ChargeType=2 then%>
				<tr class=tdbg>
				  <td align=right nowrap="nowrap" style="min-width:6rem;">剩余天数：</td>
				  <td>
				  <%if KSUser.ChargeType=3 Then%>
				  无限期
				  <%else%>
				  <%=KSUser.GetEdays%>&nbsp;天
				  <%end if%></td>
				</tr>
			   <%end if%>
				<tr class=tdbg>
				  <td align=right nowrap="nowrap" style="min-width:6rem;">当前级别：</td>
				  <td><%=KS.U_G(KSUser.GroupID,"groupname")%></td>
				</tr>
	            </table>
				</div>
				<div class="fgTitle">选择在线充值方式</div>
				<div class="tableBGw2">
				<table class=border cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
				<tr class=tdbg>
				  <td colspan="2" style="padding-left:0.75rem;">
				  <%
				   Dim HasCard:HasCard=false
				   Dim RSC,AllowGroupID:Set RSC=Conn.Execute("Select ID,GroupName,Money,AllowGroupID From KS_UserCard Where CardType=1 and DateDiff(" & DataPart_S & ",EndDate," & SqlNowString& ")<0")
				   Do While NOt RSC.Eof 
					  AllowGroupID=RSC("AllowGroupID") : If IsNull(AllowGroupID) Then AllowGroupID=" "
					 If KS.IsNul(AllowGroupID) Or KS.FoundInArr(AllowGroupID,KSUser.GroupID,",")=true Then
					  HasCard=true
					response.write "&nbsp;&nbsp; <label><input checked name=""UserCardID"" onclick=""$('#m').hide();$('#paybutton').attr('disabled',false);"" type=""radio"" value=""" & rsc("ID") & """/>" & rsc(1) & " (需要花费 <span style='color:red'>" & formatnumber(RSC(2),2,-1) & "</span> 元)</label><br/>"
					End If
					RSC.MoveNext
				   Loop
				   RSC.Close
				   Set RSC=Nothing
				  %>
				  <%If Mid(KS.Setting(170),6,1)="1" Then%>
				  <label><input <%if HasCard=false Then response.write " checked"%> onClick="$('#m').show();$('#paybutton').attr('disabled',true);" type="radio" value="0" name="UserCardID" style="vertical-align:middle; margin-right:0.25rem;">自由充(您可以任意输入要充值的金额)</label>
				  <%end if%>
				  <span id='m'<%if HasCard=true Then response.write " style=""display:none"""%> style="display:block; margin-top:0.5rem; margin-bottom:0.5rem;">请输入你要充值的金额：<input style="text-align:center; width:auto;line-height:1.1rem;color:#555;" name="Money" type="text" class="textbox" value="100" size="10" maxlength="10"> 元</span>
				  </td>
				</tr>
				<tr class=tdbg>
				  <td align=middle colSpan=2 height=40>
					<Input id="Action" type="hidden" value="paystep2" name="Action"> 
					<Input class="button" id=Submit type=submit value=" 进入在线支付 " onClick="return(Confirm(0))" name=Submit>
					<%if HasCard then%>
					<input type='hidden' name='paytype' id='paytype' value='1'/>
					<Input class="button" id="paybutton" type=submit value=" 使用余额支付 " onClick="return(Confirm(1))"  name=Submit>
					<%end if%>
					 </td>
				</tr>
			  </table>
			</FORM>
			</div>
	   <%
	    end if
	   End Sub
	   
	   Sub PayStep2()
	    Dim UserCardID:UserCardID=KS.ChkClng(KS.G("UserCardID"))
	   	Dim Money:Money=KS.S("Money")
		Dim Title,PayType
		PayType=KS.ChkClng(KS.S("PayType"))
		
		If UserCardID<>0 Then
		   Dim RS:Set RS=Conn.Execute("Select Top 1 Money,GroupName From KS_UserCard Where ID=" & UserCardID)
		   If Not RS.Eof Then
		    Title=RS(1)
		    Money=RS(0)
			RS.Close : Set RS=Nothing
		   Else
		    RS.Close : Set RS=Nothing
		    Call KS.AlertHistory("出错啦！",-1)
			Exit Sub 
		   End If
		   '判断用户有没有选择按余额购买
		   If PayType=1 Then
		     If round(KSUser.GetUserInfo("money"),2)<round(Money,2) Then
		      Call KS.AlertHistory("对不起，您可用金额不足，本充值卡需要消费" & Money & "元，您当前的可用余额为" & Formatnumber(KSUser.GetUserInfo("money"),2,-1,-1) & "元，请选择按在线购买支付！",-1)
			  Exit Sub
			 End If
			 Call UpdateByCard(1,UserCardID,KSUser.UserName,KSUser.GetUserInfo("RealName"),KSUser.GetUserInfo("Edays"),KSUser.GetUserInfo("BeginDate"),UserCardID,"")
			 Session(KS.SiteSN&"UserInfo")=empty
			 Response.Write("<script>alert('恭喜，[" & title & "]购买成功！');location.href='user_logmoney.asp';</script>")
			 response.End()
		   End If 
		   
		   
		ElseIf Mid(KS.Setting(170),6,1)="0" Then
		  KS.AlertHintScript "对不起，本站不允许会员自由充值！"
		  Exit Sub
		Else
		   Title="为自己的账户充值"
		End If

		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("对不起，您输入的充值金额不正确！",-1)
		  exit sub
		End If
		
		If Money=0 Then
		  Call KS.AlertHistory("对不起，充值金额最低为0.01元！",-1)
		  exit sub
		End If
		Dim OrderID:OrderID=KS.Setting(72) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&hour(Now)&minute(Now)&second(Now)
		
		%>
	   <FORM name=myform action="user.asp" method="post">
	      <div class="fgTitle">确认款项</div>
		  <div class="tableBGw">
		  <table id="c1" class=Cpayment cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">用户名</td>
			  <td class="aRight"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付编号</td>
			  <td class="aRight"><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付金额</td>
			  <td class="aRight"><input type='hidden' value='<%=Money%>' name='Money'><%=FormatNumber(Money,2,-1)%> 元</td>
			</tr>
			<%If title<>"" then%>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付用途</td>
			  <td style="color:red" class="aRight">“<%=title%>”</td>
			</tr>
			<%end if%>
            <tr><td class="fgTitle" colspan="2" style="padding: 0 !important;background: #efefef;font-size: 0.8rem;">支付平台</td></tr>
			<tr class=tdbg>
			  <td class="aRight" colspan="2" style="text-align:left;">
			  <%
			   Dim SQL,K,Param
			   If UserCardID<>0 Then
			    Param=" and id in(1,10,6,7,12,13,14)"
			   End IF
			   Set RS=Server.CreateOBject("ADODB.RECORDSET")
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 " & Param & " Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>对不起，本站暂不开通在线支付功能！</font>"
			   Else
			     For K=0 To Ubound(SQL,2)
				   Response.Write "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
				   If SQL(3,K)="1" Then Response.Write " checked"
				   Response.Write ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
				 Next
			   End If
			  %>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="paystep3" name="Action"> 
		        <Input id=Action type=hidden value="<%=UserCardID%>" name="UserCardID"> 
		        <Input type=hidden value="user" name="PayFrom"> 
				<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();" style="margin-right:0.5rem;"> 
				<Input class="button" id=Submit type=submit value=" 下一步 " name=Submit>
				</td>
			</tr>
		  </table>
		</FORM>
		</div>
		<%
	   End Sub
	   
	   
	   '支付商城订单
	   Sub PayShopOrder()
	  	 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 OrderID,MoneyTotal,DeliverType,Status,OrderType From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  KS.Die "<script>alert('出错啦!');history.back();</script>"
		 End If 
		If KS.ChkCLng(KS.Setting(49))=1 Then
		  If RS("Status")=0 Then
		    RS.Close:Set RS=Nothing
		   	KS.Die "<script>alert('对不起，该订单还未确认，本站启用只有后台确认过的订单才能在线支付!');history.back();</script>"
		  End If
		End If
		
		Dim OrderID:OrderID=RS("OrderID")
	   	Dim Money:Money=RS("MoneyTotal")
		Dim DeliverType:DeliverType=RS("DeliverType")
		Dim OrderType:OrderType=RS("OrderType")
		RS.Close
		Dim DeliverName,ProductName
		RS.Open "Select Top 1 TypeName From KS_Delivery Where Typeid=" & DeliverType,conn,1,1
		If Not RS.Eof Then
		 DeliverName=RS(0)
		End IF
		RS.Close
		If OrderType=1 Then
		RS.Open "Select top 10 subject as title From KS_GroupBuy Where ID in(Select proid From KS_OrderItem Where OrderID='" & OrderID& "')",conn,1,1
		Else
		RS.Open "Select top 10 Title From KS_Product Where ID in(Select proid From KS_OrderItem Where OrderID='" & OrderID& "')",conn,1,1
		End If
		If RS.Eof And RS.Bof Then
		 ProductName=OrderID
		Else
			Do While Not RS.Eof
			 if ProductName="" Then
			   ProductName=rs(0)
			 Else
			   ProductName=ProductName&","&rs(0)
			 End If
			 RS.MoveNext
			Loop
		End If
		RS.Close
		
		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("对不起，订单金额不正确！",-1)
		  exit sub
		End If
		If Money=0 Then
		  Call KS.AlertHistory("对不起，订单金额最低为0.01元！",-1)
		  exit sub
		End If
		%>
	   <FORM name=myform action="user.asp" method="post">
	      <div class="fgTitle">确认款项</div>
		  <div class="tableBGw">
		  <table id="c1" class=Cpayment cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">用户名</td>
			  <td class="aRight"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">商品名称</td>
			  <td class="aRight"><input type='hidden' value='<%=ProductName%>' name='ProductName'><%=ProductName%>&nbsp;
			  <input type='hidden' value='<%=DeliverName%>' name='DeliverName'>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付编号</td>
			  <td class="aRight"><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付金额</td>
			  <td class="aRight">
			  <%
			   Dim LessPayMoeny:LessPayMoeny=0
			   Dim PArr:Parr=Split(KS.Setting(82)&"||||||||","|")
			  If Parr(0)="1" Then
			  %><input type='hidden' value='<%=Money%>' name='Money'><%=Money%> 元<%
			  ElseIf Parr(0)="2" Then
			   LessPayMoeny=round(Parr(1),2)/100*Money
			   if ks.chkclng(Parr(3))<>0 and Money<ks.chkclng(Parr(3)) then
				  LessPayMoeny=Money
			   end if
			  %>
			  <input type='hidden' value="1" name="zfdj" />
			  <strong>预交<input type='hidden' value='<%=Parr(1)%>' name='Money'><%=LessPayMoeny%> 元定金</strong><%
			  Else %>
			   <input type='hidden' value="1" name="zfdj" />
			   <input type='text' size='8' style="height:21px;line-height:21px" name='money' value='<%=Money%>'/> 元
			 <%
			  End If
			 %>
			  </td>
			</tr>
			<tr class=tdbg><td align=left nowrap="nowrap" class="fgTitle" colspan="2" style="padding: 0 !important;background: #efefef; font-size:0.8rem;">支付平台</td></tr>
			<tr class=tdbg>
			  <td class="aRight" colspan="2" style="text-align: left; line-height:1.2rem;">
			  <%
			   Dim SQL,K
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>对不起，本站暂不开通在线支付功能！</font>"
			   Else
			     For K=0 To Ubound(SQL,2)
				   Response.Write "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
				   If SQL(3,K)="1" And KS.ChkClng(KS.S("PaymentPlat"))=0 Then Response.Write " checked"
				   iF KS.ChkClng(SQL(0,K))=KS.ChkClng(KS.S("PaymentPlat")) Then Response.Write " checked"
				   Response.Write ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
				 Next
			   End If
			  %>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
	            <input type="hidden" name="oid" value="<%=id%>"/>
		        <Input id=Action type=hidden value="paystep3" name="Action"> 
		        <Input type=hidden value="shop" name="PayFrom"> 
				<Input class="button" id=Submit type=submit value=" 下一步 " name=Submit style="margin-right:0.5rem;">
				<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
		</div>
		<%
	   End Sub
	   
	   Sub PayStep3()
	    Dim UserCardID,Title
		UserCardID=KS.ChkClng(KS.S("UserCardID"))
	    Dim Money:Money=KS.S("Money")
		Dim MoneyTotal:MoneyTotal=0
		Dim Oid:Oid=KS.ChkClng(request("oid"))
		if oid<>0 then
		  dim rs:set rs=conn.execute("select top 1 MoneyTotal from ks_order where id=" & oid)
		  if not rs.eof then
		    MoneyTotal=rs(0)
		  end if
		  rs.close:set rs=nothing
		end if
		Dim LessPayMoney:LessPayMoney=0
		If KS.S("zfdj")="1" Then
			Dim PArr:Parr=Split(KS.Setting(82)&"||||||||","|")
			If Parr(0)="1" Then
			ElseIf Parr(0)="2" Then
			 if ks.chkclng(Parr(3))<>0 and MoneyTotal<ks.chkclng(Parr(3)) then
			  money=MoneyTotal
			 end if
			Else 
			 Money=KS.S("Money"): If Not Isnumeric(Parr(2)) Then Parr(2)=0
			 If Not IsNumerIc(Money) Then
				KS.Die "<script>alert('对不起，订单金额不正确！');history.back();</script>"
			 End If
			 
			 	 If Parr(2)<>0 then  lessPayMoney=round(Parr(2),2)/100*MoneyTotal
				 If Not IsNumerIc(Money) Then  KS.Die "<script>$.dialog.tips('对不起，订单金额不正确！',1,'error.gif',function(){window.close();});</script>"
				 
				if ks.chkclng(Parr(3))<>0 and round(money,2)<ks.chkclng(Parr(3)) and MoneyTotal>ks.chkclng(Parr(3)) then KS.Die "<script>$.dialog.tips('对不起，支付金额不能少于" & ks.chkclng(Parr(3)) & "元！',1,'error.gif',function(){window.close();});</script>"
				
				If (LessPayMoney<>0 and Round(Money,2)<round(LessPayMoney,2)) Or Money="0" Then KS.Die "<script>$.dialog.tips('对不起，支付金额必须大于订单总额的" & parr(2) & "%,即不能少于" & round(LessPayMoney,2) & "元！',1,'error.gif',function(){window.close();});</script>"

			End If
		End If
		Dim OrderID:OrderID=KS.S("OrderID")
		Dim ProductName:ProductName=KS.CheckXSS(KS.S("ProductName"))
		Dim PaymentPlat:PaymentPlat=KS.ChkClng(KS.S("PaymentPlat"))
		Dim PayUrl,PayMentField,ReturnUrl,RealPayMoney,RealPayUSDMoney,RateByUser,PayOnlineRate
        Call GetPayMentField(OrderID,PaymentPlat,Money,UserCardID,ProductName,KS.S("PayFrom"),KSUser,PayMentField,PayUrl,ReturnUrl,Title,RealPayMoney,RealPayUSDMoney,RateByUser,PayOnlineRate)
		
		 %>
	   	  <FORM name="myform"  id="myform" action="<%=PayUrl%>" <%if PaymentPlat=11 or PaymentPlat=9 then response.write "method=""get""" else response.write "method=""post"""%>  target="_blank">
		  <div class="fgTitle">确认款项</div>
		  <div class="tableBGw">
		  <table id="c1" class=Cpayment cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">用户名</td>
			  <td class="aRight"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付编号</td>
			  <td class="aRight"><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付金额</td>
			  <td class="aRight"><%=formatnumber(Money,2,-1)%> 元</td>
			</tr>
			<%if title<>"" then%>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">支付用途</td>
			  <td style="color:red" class="aRight">“<%=title%>”</td>
			</tr>
			<%end if%>
			<%
			if RateByUser=1 then
			%>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">手续费</td>
			  <td class="aRight"><%=PayOnlineRate%>%</td>
			</tr>
			<%end if%>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">实际支付</td>
			  <td class="aRight">
			  <%=formatnumber(RealPayMoney,2,-1)%></td>
			</tr>
			<%If PaymentPlat=12 Then%>
			<tr class=tdbg>
			  <td align=left nowrap="nowrap" class="clefttitle">实际支付美金</td>
			  <td style="color:#FF6600;font-weight:bold" class="aRight">
			  $<%=formatnumber(RealPayUSDMoney,2,-1)%> USD</td>
			</tr>
			<%End If%>
			<tr class=tdbg>
			  <td colspan=2 class="aRight" style="text-align:left;">点击“确认支付”按钮后，将进入在线支付界面，在此页面选择您的银行卡。</td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
			    <%=PayMentField%>
				<%if PaymentPlat=9 then%>
				<Input class="button" id=Submit type=button onClick="$('#myform').submit()" value=" 确定支付 " style="margin-right:0.5rem;">
				<%else%>
				<Input class="button" id=Submit type=submit value=" 确定支付 " style="margin-right:0.5rem;">
				<%end if%>
				<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
	   </div>	  
	   <%
	   End Sub
	   
	   '充值卡充值保存
	   Sub cardsave()
	     Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 DiM CardNum:CardNum=KS.S("CardNum")
		 Dim CardPass:CardPass=KS.S("CardPass")
		 If CardNum="" Or CardPass="" Then 
		   Call KS.AlertHistory("请输入的充值卡号及密码！",-1)
		   exit sub
		 end if
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_usercard where cardtype=0 and cardnum='" & CardNum & "'",conn,1,1
		 if rs.bof and rs.eof then
		  rs.close:set rs=nothing
		  KS.Die "<script> $.dialog.alert('对不起，您输入的充值卡号不正确！',function(){ history.back(-1);});</script>"
		  exit sub
		 end if
		 if rs("cardpass")<>KS.Encrypt(cardpass) then
		  rs.close:set rs=nothing
		   KS.Die "<script> $.dialog.alert('对不起，您输入的充值卡密码不正确！',function(){ history.back(-1);});</script>"
		  exit sub
		 end if
		 
		 if rs("isused")=1 then
		   rs.close:set rs=nothing
		   KS.Die "<script> $.dialog.alert('对不起，您输入的充值卡已被使用！',function(){ history.back(-1);});</script>"
		  exit sub
		 end if
		 
		 if datediff("d",rs("enddate"),now())>0 then
		  rs.close:set rs=nothing
		   KS.Die "<script> $.dialog.alert('对不起，您输入的充值卡已过期！',function(){ history.back(-1);});</script>"
		  exit sub
		 end if
		 
		 if not KS.IsNul(rs("allowgroupid")) then
		    If KS.FoundInArr(rs("allowGroupID"),KSUser.GroupID,",")=false Then
			  rs.close:set rs=nothing
		      KS.Die "<script> $.dialog.alert('对不起，您所在的用户组没有使用本充值卡的权限,请联系本站管理员！',function(){ history.back(-1);});</script>"
			  exit sub
			End If
		 end if
		 
		  Dim ValidNum:ValidNum=rs("ValidNum")
		  Dim ValidUnit:ValidUnit=rs("ValidUnit")
		  Dim UserCardID:UserCardID=rs("id")
		  Dim GroupID:GroupID=rs("GroupID")
		  rs.close
		  rs.open "select top 1 * from ks_user Where UserName='" & KSUser.UserName & "'",conn,1,1
		  if not rs.eof then
		    if rs("ChargeType")=3 and ValidUnit<>3 then
				  rs.close:set rs=nothing
		          KS.Die "<script> $.dialog.alert('由于你的账户永不过期，如需充值资金，请购买资金卡！',function(){ history.back(-1);});</script>"
				  exit sub
			end if
			dim ValidDays,tmpdays
		    select case ValidUnit
			  case 1 '点数
			   'rs("point")=rs("point")+ValidNum
			   Call KS.PointInOrOut(0,0,rs("UserName"),1,ValidNum,"System","通过充值卡获得的点数",0)
			  case 2 '天数
			    ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
				    conn.execute("update ks_user set chargetype=2,edays=edays+" & validnum & " where username='" & ksuser.username & "'")
				else
					conn.execute("update ks_user set chargetype=2,begindate=" & sqlnowstring & ",edays=" & validnum & " where username='" & ksuser.username & "'")
				end if
				Call KS.EdaysInOrOut(rs("UserName"),1,ValidNum,"System","通过充值卡[" & CardNum & "]获得的有效天数")
			  case 3 '金币
			    Call KS.MoneyInOrOut(rs("UserName"),RS("RealName"),ValidNum,4,1,now,0,"System","通过充值卡[" & CardNum & "]获得的资金",0,0,0)
			  case 4 '积分
			    Call KS.ScoreInOrOut(rs("UserName"),1,ValidNum,"System","通过充值卡[" & CardNum & "]获得的积分!",0,0)
			end select
			if GroupID<>0 then conn.execute("update ks_user set groupid=" & GroupID & " where userName='" & KSUser.UserName & "'")
			conn.execute("update ks_user set usercardid="&usercardid &" where userName='" & KSUser.UserName & "'")
		  end if
		  '置充值卡已使用、已售出
		  Conn.Execute("Update KS_UserCard Set Isused=1,issale=1,username='" & KSUser.UserName & "',UseDate=" & SqlNowString & " where cardnum='" & cardnum & "'")
		 Session(KS.SiteSN&"UserInfo")=""
		 if GroupID<>0 then
		 Response.Write "<script>alert('恭喜您，充值成功并升级为"""& KS.U_G(GroupID,"groupname") &"""!');location.href='use.asp';</script>"
		 else
		 Response.Write "<script>alert('恭喜您，充值成功!');location.href='user.asp';</script>"
		 end if
		 RS.Close:Set RS=Nothing
	   End Sub
	   
	   
	    
    '消费记录
   Sub Logmoney() 
   %>
   <div class="tabs exRecord">	
			<ul>
				<li<%if ks.s("flag")="" then response.write" class='puton'"%>><a href="?action=logmoney">资金明细</a></li>
				<li<%if ks.s("flag")="point" then response.write" class='puton'"%>><a href="?action=logmoney&flag=point">点券明细</a></li>
				<li<%if ks.s("flag")="score" then response.write" class='puton'"%>><a href="?action=logmoney&flag=score">积分明细</a></li>
			</ul>
   </div>
  <% 
     Select case KS.S("flag")
	   case "point"
	      Call LogPoint()
	   case "score"
	      Call LogScore()
	   Case else
	      Call LogMoneyMain()
	 End Select
					   
   End Sub
   
    '资金明细
   Sub LogMoneyMain()
          Dim SQLStr,RS
					    If KS.ChkClng(KS.S("IncomeOrPayOut"))=1 Or KS.ChkClng(KS.S("IncomeOrPayOut"))=2 Then
						  SqlStr="Select * From KS_LogMoney Where IncomeOrPayOut=" & KS.ChkClng(KS.S("IncomeOrPayOut")) & " And  UserName='" & KSUser.UserName &"' order by id desc"
 					    Else
						  SqlStr="Select * From KS_LogMoney Where UserName='" & KSUser.UserName &"' order by id desc"
						End if
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<table clas=""border""><tr class='tdbg'><td align=center height=25 colspan=9 class=""empty"" valign=top><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe723;</i></div>找不到您要的记录!</div></td></tr></table>"
								 Else
									totalPut = RS.RecordCount
						            If (CurrentPage - 1) * MaxPerPage < totalPut Then
											RS.Move (CurrentPage - 1) * MaxPerPage
									End If
									 Dim I,intotalmoney,outtotalmoney
									  intotalmoney=0
									 outtotalmoney=0
									 Do While Not rs.eof 
									%>
                                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="myCollection3" >
								<tr>
					               <td class="ContentTitle" style="line-height: 1rem;padding: 0.75rem;font-size: 0.65rem; color:#888;">说明：<%=rs("Remark")%> 	<br/>交易时间：<%=rs("LogTime")%> 
								   </td>
								</tr>
		                        <tr>
					              <td class="Contenttips">
									  交易用户：<%=rs("username")%> <br/>
									  交易金额：<%Response.Write formatnumber(rs("money"),2,-1)%>元
									  <%
									    If rs("IncomeOrPayOut")=1 Then
										 intotalmoney=intotalmoney+rs("money")
										ElseIf rs("IncomeOrPayOut")=2 Then
										 outtotalmoney=outtotalmoney+rs("money")
										End If
										%>
									  <br/>交易类型：<% If rs("IncomeOrPayOut")=1 Then
										  Response.Write "<font color=red>收入</font>"
										 Else
										  Response.Write "<font color=green>支出</font>"
										 End If
										 %>
									  <br/>
									  当前余额：
									  <%=formatnumber(RS("CurrMoney"),2,-1)%>元
										</td>
									</tr>
									</table>
									<%
												
												I = I + 1
												RS.MoveNext
												If I >= MaxPerPage Then Exit Do
								
									 loop
									%>
							      <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="pageMoney" >
									<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
										 <td class="splittd"  align=right>本页收入：￥<span><%=formatnumber(intotalmoney,2,-1)%></span>&nbsp;&nbsp;&nbsp;本页支出：￥<span><%=formatnumber(outtotalmoney,2,-1)%></span></td>
									</tr>
									<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
									  <td class="splittd" align=right>
									 <%
									    intotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where UserName='" & KSUser.UserName & "' And IncomeOrPayOut=1")(0)
										outtotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where UserName='" & KSUser.UserName & "' And IncomeOrPayOut=2")(0)
										if not isnumeric(intotalmoney) then intotalmoney=0
										if not isnumeric(outtotalmoney) then outtotalmoney=0
									  %>总计收入：￥<span><%=formatnumber(intotalmoney,2,-1)%></span>&nbsp;&nbsp;&nbsp;总计支出：￥<span><%=formatnumber(outtotalmoney,2,-1)%></span>
									 </td>
									</tr>
									<tr><td class="splittd"></td></tr>
								  </table>
										<%
				         End If
						 %>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
	End Sub
   
   '点券明细
   Sub LogPoint()
          Dim SQLStr,RS
					    If KS.ChkClng(KS.S("IncomeOrPayOut"))=1 Or KS.ChkClng(KS.S("IncomeOrPayOut"))=2 Then
						  SqlStr="Select * From KS_LogPoint Where InOrOutFlag=" & KS.ChkClng(KS.S("IncomeOrPayOut")) & " And  UserName='" & KSUser.UserName &"' order by id desc"
 					    Else
						  SqlStr="Select * From KS_LogPoint Where UserName='" & KSUser.UserName &"' order by id desc"
						End if
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<table class=""noborder"" width=""100%""><tr class='tdbg'><td align=center height=25 colspan=9 valign=top><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe723;</i></div>找不到您要的记录!</div></td></tr></table>"
								 Else
									totalPut = RS.RecordCount
						            If (CurrentPage - 1) * MaxPerPage < totalPut Then
											RS.Move (CurrentPage - 1) * MaxPerPage
									End If
									 Dim I,intotalmoney,outtotalmoney
									 intotalmoney=0
									 outtotalmoney=0
									 Do While Not rs.eof 
									%>
                           <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="PointsDetails" >
									<tr class="title" style=" background:#4dbdf5 ;">
					                  <td class="ContentTitle" style=" line-height:1rem">说明：<%=rs("descript")%><br/>交易时间：<%=rs("adddate")%> 
									  </td>
									  </tr>
		                        <tr>
					              <td class="Contenttips">
									  交易用户：<%=rs("username")%> <br/>
									  交易点券：<%Response.Write rs("point")%>
									  <%
									    If rs("InOrOutFlag")=1 Then
										 intotalmoney=intotalmoney+rs("point")
										ElseIf rs("InOrOutFlag")=2 Then
										 outtotalmoney=outtotalmoney+rs("point")
										End If
										%>
									  <br/>交易类型：<% If rs("InOrOutFlag")=1 Then
										  Response.Write "<font color=red>收入</font>"
										 Else
										  Response.Write "<font color=green>支出</font>"
										 End If
										 %>
									  <br/>
									  当前点券余额：
									  <%=rs("currpoint")%>
										</td>
									</tr>
									</table>
									<%
												I = I + 1
												RS.MoveNext
												If I >= MaxPerPage Then Exit Do
								
									 loop
									%>
							 <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="pageMoney" >
									<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
										 <td class="splittd"  align=right>
											本页收入：￥<span><%=intotalmoney%></span>&nbsp;&nbsp;&nbsp;本页支出：￥<span><%=outtotalmoney%></span></td>
									</tr>
									<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
									  <td class="splittd" align=right>
									 <%
									    intotalmoney=Conn.execute("Select Sum(point) From KS_LogPoint Where UserName='" & KSUser.UserName & "' And InOrOutFlag=1")(0)
										outtotalmoney=Conn.execute("Select Sum(point) From KS_LogPoint Where UserName='" & KSUser.UserName & "' And InOrOutFlag=2")(0)
										if not isnumeric(intotalmoney) then intotalmoney=0
										if not isnumeric(outtotalmoney) then outtotalmoney=0
									  %>总计收入：￥<span><%=intotalmoney%></span>&nbsp;&nbsp;&nbsp;总计支出：￥<span><%=outtotalmoney%></span>
									 </td>
									</tr>
								  </table>
										<%
				         End If
						 %>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
	End Sub
	
   '积分明细
   Sub LogScore()
          Dim SQLStr,RS
					    If KS.ChkClng(KS.S("IncomeOrPayOut"))=1 Or KS.ChkClng(KS.S("IncomeOrPayOut"))=2 Then
						  SqlStr="Select * From KS_LogScore Where InOrOutFlag=" & KS.ChkClng(KS.S("IncomeOrPayOut")) & " And  UserName='" & KSUser.UserName &"' order by id desc"
 					    Else
						  SqlStr="Select * From KS_LogScore Where UserName='" & KSUser.UserName &"' order by id desc"
						End if
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<table clas=""border""><tr class='tdbg'><td align=center height=25 colspan=9 valign=top><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe723;</i></div>找不到您要的记录!</div></td></tr></table>"
								 Else
									totalPut = RS.RecordCount
						            If (CurrentPage - 1) * MaxPerPage < totalPut Then
											RS.Move (CurrentPage - 1) * MaxPerPage
									End If
									 Dim I,intotalmoney,outtotalmoney
									 intotalmoney=0
									 outtotalmoney=0
									 Do While Not rs.eof 
									%>
                           <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="PointsDetails" >
									<tr class="title">
					                  <td class="ContentTitle" style="line-height: 1rem;">说明：<%=rs("descript")%> 	<br/>交易时间：<%=rs("adddate")%> 
									  </td>
									  
									  </tr>
		                        <tr>
					              <td class="Contenttips">
									  交易用户：<%=rs("username")%> <br/>
									  交易积分：<%Response.Write rs("score")%>分
									  <%
									    If rs("InOrOutFlag")=1 Then
										 intotalmoney=intotalmoney+rs("score")
										ElseIf rs("InOrOutFlag")=2 Then
										 outtotalmoney=outtotalmoney+rs("score")
										End If
										%>
									  <br/>交易类型：<% If rs("InOrOutFlag")=1 Then
										  Response.Write "<font color=red>收入</font>"
										 Else
										  Response.Write "<font color=green>支出</font>"
										 End If
										 %>
									  <br/>
									  当前余额：
									  <%=rs("currscore")%>分
										</td>
									</tr>
									</table>
									<%
												I = I + 1
												RS.MoveNext
												If I >= MaxPerPage Then Exit Do
								
									 loop
									%>
							 <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="pageMoney" >
									<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
										 <td class="splittd"  align=right>本页收入：<span><%=intotalmoney%></span>分&nbsp;&nbsp;&nbsp;本页支出：<span><%=outtotalmoney%></span>分</td>
									</tr>
									<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
									  <td class="splittd" align=right>
									 <%
									    intotalmoney=Conn.execute("Select Sum(score) From KS_Logscore Where UserName='" & KSUser.UserName & "' And InOrOutFlag=1")(0)
										outtotalmoney=Conn.execute("Select Sum(score) From KS_Logscore Where UserName='" & KSUser.UserName & "' And InOrOutFlag=2")(0)
										if not isnumeric(intotalmoney) then intotalmoney=0
										if not isnumeric(outtotalmoney) then outtotalmoney=0
									  %>总计收入：<span><%=intotalmoney%></span>分&nbsp;&nbsp;&nbsp;总计支出：<span><%=outtotalmoney%></span>分
									 </td>
									</tr>
									<tr><td class="splittd"></td></tr>
								  </table>
										<%
				         End If
						 %>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
	End Sub
	
%> 

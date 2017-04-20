<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.membercls.asp"-->
<!--#include file="../plus/md5.asp"-->
<!--#include file="../user/payfunction.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
%>
<!DOCTYPE html>
<html>
<head>
<title>正在为您接入在线支付平台...</title>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<script src="../ks_inc/jquery.js" type="text/javaScript"></script>
<script src="../ks_inc/common.js"></script>
</head>
<body>
<%
Dim KS:Set KS=New PublicCls

         Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 OrderID,MoneyTotal,DeliverType,Status,OrderType From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  KS.Die "<script>$.dialog.tips('出错啦,没有找到订单！',1,'error.gif',function(){window.close();});</script>"
		 End If 
		 
		 If KS.ChkCLng(KS.Setting(49))=1 Then
		  If RS("Status")=0 Then
		  KS.Die "<script>$.dialog.tips('对不起，该订单还未确认，本站启用只有后台确认过的订单才能付款！',1,'error.gif',function(){window.close();});</script>"
			
		  End If
		End If
		 
		Dim OrderID:OrderID=RS("OrderID")
	   	Dim Money:Money=RS("MoneyTotal")
		Dim DeliverType:DeliverType=RS("DeliverType")
		Dim OrderType:OrderType=RS("OrderType")
		Dim MoneyTotal:MoneyTotal=RS("MoneyTotal")
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
		  KS.Die "<script>$.dialog.tips('对不起，订单金额不正确！',1,'error.gif',function(){window.close();});</script>"
		End If
		If Money=0 Then
		  KS.Die "<script>$.dialog.tips('对不起，订单金额最低为0.01元！',1,'error.gif',function(){window.close();});</script>"
		End If
Dim KSUser:Set KSUser=New UserCls
KSUser.UserLoginChecked
Dim PaymentPlat:PaymentPlat=KS.ChkClng(KS.S("PaymentPlat"))  '支付平台
If PaymentPlat=0 Then PaymentPlat=7
Dim PayMentField,PayUrl,ReturnUrl,Title,RealPayMoney,RateByUser,PayOnlineRate,RealPayUSDMoney

Dim LessPayMoney:LessPayMoney=0
Dim PArr:Parr=Split(KS.Setting(82)&"||||||||","|")
If Parr(0)="1" Then
ElseIf Parr(0)="2" Then
 Money=round(Parr(1),2)/100*MoneyTotal
 if ks.chkclng(Parr(3))<>0 and MoneyTotal<ks.chkclng(Parr(3)) then
  money=MoneyTotal
 end if
Else 
	 if isnumeric(KS.S("Money")) Then
	  Money=KS.S("Money"): If Not Isnumeric(Parr(2)) Then Parr(2)=0
	 End If
	 
	 If Parr(2)<>0 then  lessPayMoney=round(Parr(2),2)/100*MoneyTotal
	 If Not IsNumerIc(Money) Then  KS.Die "<script>$.dialog.tips('对不起，订单金额不正确！',1,'error.gif',function(){history.back();});</script>"
	 If round(Money)>round(MoneyTotal) Then KS.Die "<script>$.dialog.tips('对不起，本单只需支付" & MoneyTotal& "元！',1,'error.gif',function(){history.back();});</script>"
	if ks.chkclng(Parr(3))<>0 and round(money,2)<ks.chkclng(Parr(3)) and MoneyTotal>ks.chkclng(Parr(3)) then KS.Die "<script>$.dialog.tips('对不起，支付金额不能少于" & ks.chkclng(Parr(3)) & "元！',1,'error.gif',function(){history.back();});</script>"
	
	If (LessPayMoney<>0 and Round(Money,2)<round(LessPayMoney,2)) Or Money="0" Then KS.Die "<script>$.dialog.tips('对不起，支付金额必须大于订单总额的" & parr(2) & "%,即不能少于" & round(LessPayMoney,2) & "元！',1,'error.gif',function(){history.back();});</script>"
End If
Call GetPayMentField(OrderID,PaymentPlat,Money,0,ProductName,"shop",KSUser,PayMentField,PayUrl,ReturnUrl,Title,RealPayMoney,RealPayUSDMoney,RateByUser,PayOnlineRate)
%>
<%if PaymentPlat<>16 then%>
正在为您接入在线支付平台，请稍等....
<%end if%>
<FORM name="myform"  id="myform" action="<%=PayUrl%>" <%if PaymentPlat=11 or PaymentPlat=9 then response.write "method=""get""" else response.write "method=""post"""%> >
		  <table <%if PaymentPlat<>16 then response.write " style='display:none'"%> id="c1" class=border cellSpacing=1 cellPadding=2 width="80%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 确 认 款 项 并 支 付</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>用户名：</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>支付编号：</td>
			  <td><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>支付金额：</td>
			  <td><%=formatnumber(Money,2,-1)%> 元</td>
			</tr>
			
             <%If PaymentPlat=12 Then%>
			<tr class=tdbg>
			  <td align=right width=167>实际支付美金：</td>
			  <td style="color:#FF6600;font-weight:bold">
			  $<%=formatnumber(RealPayUSDMoney,2,-1)%> USD</td>
			</tr>
			<%End If%>
						<%if title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>支付用途：</td>
			  <td style="color:red">“<%=title%>”</td>
			</tr>
			<%end if%>
			<%
			if RateByUser=1 then
			%>
			<tr class=tdbg>
			  <td align=right width=167>手续费：</td>
			  <td><%=PayOnlineRate%>%</td>
			</tr>
			<%end if%>

			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
			  <%if PaymentPlat=16 then%>
			  请打开微信APP，找到“扫一扫“对准以下二维码完成支付操作。
			  <br/>
			  <img src="../user/wxpay/images/wxpay.jpg" align="left"/>
			  <%end if%>
			    <%=PayMentField%>
				<%if PaymentPlat<>16 then%>
					<%if PaymentPlat=9 then%>
					<Input class="button" id=Submit type=button onClick="$('#myform').submit()" value=" 确定支付 " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
					<%else%>
					<Input class="button" id=Submit type=submit value=" 确定支付 " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
					<%end if%>
					<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();"> 
				<%end if%>	
				</td>
			</tr>
		  </table>
		</FORM>
		<%if PaymentPlat<>16 then%>
		<script type="text/javascript">
		 $(function(){
		  $("#myform").submit();
		 });
		</script>
		<%end if%>
<%
CloseConn
Set KSUser=Nothing
Set KS=Nothing
%>
</body>
</html>
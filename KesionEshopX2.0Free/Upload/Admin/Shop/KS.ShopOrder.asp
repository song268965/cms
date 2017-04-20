<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../user/payfunction.asp"-->
<!--#include file="../../plus/md5.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls,Md5Key
Set KSCls = New Admin_ShopOrder
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ShopOrder
        Private KS,KSCls
		Private totalPut, CurrentPage, MaxPerPage,DomainStr
		Private SqlStr,PageTotalMoney1,PageTotalMoney2,SqlTotalMoney,RS,SqlParam,SearchType
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		
		
		Sub ModifyOrder()
		 dim moneytotal:moneytotal=ks.g("moneytotal")
		 dim id:id=ks.g("id")
		%>
		<style>
		</style>
		 <dl class="dtable" style="padding-top: 37px;background: #fff;padding-bottom: 37px;">
            <dd><div>当前价格:</div>
   <iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='KS.ShopOrder.asp?action=ModifyTotalPrice' method='post'>￥<%=moneytotal%>元<br/><input type='hidden' value='<%=moneytotal%>' name='oprice'><input type='hidden' name='Id' value='<%=id%>'>
		</dd>
		<dd><div>将订单总价格改为:</div>
		<input type='text' value='<%=moneytotal%>' name='price' style='width:60px;text-align:center; margin-right:10px; margin-left:0;' class="textbox">元<br /><input style='margin-top:7px' class='button' type='submit' value='确定修改'></form>
		</dd>
	   </dl>
	   
		<%
		End Sub
		
		Public Sub Kesion()
		
		   If Not KS.ReturnPowerResult(5, "M510020") Then                  '权限检查
			Call KS.ReturnErr(1, "")   
			Response.End()
		  End iF

			SearchType=KS.ChkClng(KS.G("SearchType"))
	if KS.G("Action")<>"printexpress" then%>
<!DOCTYPE html><html>
<head><title>订单处理</title>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<link href="../include/Admin_Style.css" type=text/css rel=stylesheet>
<script src="../../ks_inc/jquery.js"></script>
<script src="../../ks_inc/common.js"></script>
<script type="text/javascript">
  var box='';
  function modifyPrice(ev,title,orderid,id,price)
  {
    box=top.$.dialog({title:"商品价格",content:"<iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='shop/KS.ShopOrder.asp?action=ModifyPrice' method='post' style='line-height:30px;'>商品名称:"+title+"<br/><input type='hidden' value='"+price+"' name='oprice'><input type='hidden' name='orderId' value='"+orderid+"'><input type='hidden' name='Id' value='"+id+"'>实收价格:<input type='text' value='"+price+"' name='price' style='width:40px;text-align:center; height: 25px;border: 1px solid #ddd;margin: 0 5px;'>元<br /><input style='margin-top:7px;background: #3c85ba;color: #fff;border: 1px solid #3c85ba;padding: 0 10px;height: 30px;' class='button' type='submit' value='确定修改'></form>",width:240});
  }
  function modifytotalprice(id,moneytotal){
  
    top.openWin('修改订单总价','shop/KS.ShopOrder.asp?Action=ModifyPrice&id='+id+'&moneytotal='+moneytotal,false,520,265)

    //box=top.$.dialog({title:"修改订单总价",content:"<iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='shop/KS.ShopOrder.asp?action=ModifyTotalPrice' method='post'>当前价格:￥"+moneytotal+"元<br/><input type='hidden' value='"+moneytotal+"' name='oprice'><input type='hidden' name='Id' value='"+id+"'>将订单总价格改为:<input type='text' value='"+moneytotal+"' name='price' style='width:60px;text-align:center'>元<br /><input style='margin-top:7px' onclick='top.box.close();' class='button' type='submit' value='确定修改'></form>",width:240});
  }
  function addservice(id){
    top.openWin('配送服务明细','shop/KS.ShopOrder.asp?action=addservice&id='+id,false,960,420);
  }
  function modifyInfo(id){
   top.openWin('修改送货资料','shop/KS.ShopOrder.asp?action=modifyinfo&id='+id,false,960,420);
  }
  function modifyproduct(id){
   top.openWin('修改/添加商品','shop/KS.ShopOrder.asp?action=modifyproduct&id='+id,true,960,420);
  }
</script>
</head>
<body>
<%end if%>
 <%
   If KS.G("Action")="PrintOrder" Then
     Call PrintOrder()
     Response.end
   elseif KS.G("Action")="ModifyPrice" then
     Call ModifyOrder()
	 Response.End()
   End IF
   If KS.G("Action")<>"modifyinfo" and KS.G("Action")<>"printexpress" and KS.G("Action")<>"modifyproduct" and KS.G("Action")<>"addservice" Then
  %>
  <div class="tableTop mt20">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
<FORM name=form1 action=KS.ShopOrder.asp method=get>
      <td><strong class="mr0">订单处理：</strong></td>
      <td valign="top"><span class="mr0">快速查询：</span> 
<Select onchange=javascript:submit() size=1 name=SearchType> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>所有订单</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>24小时之内的新订单</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>最近10天内的新订单</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>最近一月内的新订单</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>未确认的订单</Option> 
  <Option value=5<%If SearchType="5" Then Response.write " selected"%>>未付款的订单</Option> 
  <Option value=6<%If SearchType="6" Then Response.write " selected"%>>未付清的订单</Option> 
  <Option value=7<%If SearchType="7" Then Response.write " selected"%>>未送货的订单</Option> 
  <Option value=8<%If SearchType="8" Then Response.write " selected"%>>未签收的订单</Option> 
  <Option value=9<%If SearchType="9" Then Response.write " selected"%>>未开发票的订单</Option> 
  <Option value=11<%If SearchType="11" Then Response.write " selected"%>>未结清的订单</Option> 
  <Option value=12<%If SearchType="12" Then Response.write " selected"%>>已结清的订单</Option>
  <Option value=13<%If SearchType="13" Then Response.write " selected"%>>需要服务跟踪的订单</Option>
      </Select></td></FORM>
<FORM name=form2 action=KS.ShopOrder.asp method=post>
      <td><B>高级查询：</B> 
	<Select id="Field" name="Field"> 
  <Option value=1>订单编号</Option> 
  <Option value=2>收货人</Option> 
  <Option value=3>用户名</Option> 
  <Option value=4>联系地址</Option> 
  <Option value=5>联系电话</Option> 
  <Option value=6>下单时间</Option>
  <Option value=7>推荐人</Option>
</Select> 
  <Input class='textbox' id=Keyword maxLength=30 name=Keyword> 
  <Input type=submit value=" 查 询 " class='button' name=Submit2> 
        <Input id=SearchType type=hidden value=10 name=SearchType> </td></FORM>
    </tr>
  </table>
  </div>
  <%
  Response.Write ""	
   End If
  
  
		  Select Case KS.G("Action")
		   case "BankRefundOK" BankRefundOK  '退款妥协成功，恢复正常
		   case "addservice" addservice
		   Case "ModifyTotalPrice"
		    Call ModifyTotalPrice()
		   Case "modifyinfo"
		    Call modifyinfo()
		   Case "printexpress"
		    Call printexpress()
		   Case "DoModifyInfoSave"
		    Call DoModifyInfoSave()
		   Case "modifyproduct"
		    Call modifyproduct()
		   Case "doModifyProductSave"
		    Call doModifyProductSave()
		   Case "ProAddToOrder"
		    Call ProAddToOrder()
		   Case "delproduct"
		    Call delproduct()
		   Case "ShowOrder"
		    Call ShowOrder()
		   Case "DelOrder"
		    Call DelOrder()
		   Case "OrderConfirm"
		    Call OrderConfirm()
		   Case "BankPay"     '付款
		    Call BankPay() 
		   Case "DoBankPay"    '银行付款操作
		    Call DoBankPay()
		   Case "BankRefund"    '退款
		    Call BankRefund()
		   Case "DoRefundMoney" '退款操作
		    Call DoRefundMoney()
		   Case "DeliverGoods"  '发货
		    Call DeliverGoods()
		   Case "DoDeliverGoods" '发货操作 
		    Call DoDeliverGoods()
		   Case "BackGoods"     '退货
		    Call BackGoods()
		   Case "SaveBack"     '退货操作
		     Call SaveBack()
		   Case "PayMoney"      '支付货款给卖方
		    Call PayMoney()
		   Case "DoPayMoney"    '支付货款
		    Call DoPayMoney()
		   Case "Invoice"   '开发票
		    Call Invoice()
		   Case "DoSaveInvoice"
		    Call DoSaveInvoice()
		   Case "ClientSignUp"   '已签收商品
		    Call ClientSignUp()
		   Case "FinishOrder"     '结算清单
		    Call FinishOrder()
		   Case "ModifyPrice"    '修改指定价
		    Call ModifyPrice()
		   Case Else
		    Call OrderList
		  End Select
		End Sub
		
		sub BankRefundOK()
		  dim id:id=ks.g("id")
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 1 * from KS_Order Where OrderID='" & id & "'",conn,1,1
		  if rs.eof and rs.bof then
		   rs.close
		   set rs=nothing
		  end if
		  dim myid:myid=rs("id")
		  rs.close
		  set rs=nothing
		  dim isdelivery:isdelivery=0
		  if not conn.execute("select top 1 * from KS_LogDeliver where orderid='" & id & "' and DeliverType=1").eof then
		   isdelivery=1
		  end if
		  conn.execute("update ks_order set DeliverStatus=" & isdelivery & " where orderid='" & id & "'")
		  conn.execute("update KS_LogDeliver set DeliverType=4,status=1 where orderid='" & id & "' and DeliverType=3")
		  ks.die "<script>alert('恭喜，订单状态恢复正常成功！');location.href='KS.ShopOrder.asp?Action=ShowOrder&ID=" & myid &"';</script>" 
		end sub
		
		
		'服务明细
		Sub addservice()
		 dim id:id=ks.chkclng(request("id"))
		 if id=0 then ks.die "<script>alert('出错!');top.box.close();</script>"
		 Dim RS:SET RS=Server.CreateObject("ADODB.RECORDSET")
		 if request("Flag")="dosave" then  '保存
		   if Request("ServiceName")="" then ks.die "<script>alert('请输入服务名称!');history.back();</script>"
		   if KS.ChkCLng(Request("ServiceTimes"))=0 then ks.die "<script>alert('请输入服务次数!');history.back();</script>"
		   if KS.ChkCLng(Request("Validity"))=0 then ks.die "<script>alert('请输入服务有效期!');history.back();</script>"
		   rs.open "select top 1 * From KS_Order Where ID=" & ID,conn,1,3
		   If Not RS.Eof Then
		     RS("IsService")=KS.ChkClng(Request("IsService"))
			 RS("ServiceName")=Request("ServiceName")
			 RS("ServiceTimes")=KS.ChkCLng(Request("ServiceTimes"))
			 RS("Validity")=KS.ChkClng(Request("Validity"))
			 RS.Update
		   End If
		   RS.Close:Set RS=Nothing
		   KS.Die "<script>alert('恭喜，修改成功!');location.href='" & Request.ServerVariables("HTTP_REFERER") &"';</script>"
		 elseif request("flag")="additem" then
           if Request("content")="" then ks.die "<script>alert('请输入服务内容!');history.back();</script>"
           if not isdate(Request("adddate")) then ks.die "<script>alert('时间格式不正确，请重输!');history.back();</script>"
		   rs.open "select top 1 * from ks_orderservice where 1=0",conn,1,3
		   rs.addnew
		     rs("orderid")=id
			 rs("content")=request("content")
			 rs("adddate")=request("adddate")
			 rs("qsr")=request("qsr")
		   rs.update
		   rs.close
		   set rs=nothing
		   KS.Die "<script>alert('恭喜，服务记录添加成功!');location.href='" & Request.ServerVariables("HTTP_REFERER") &"';</script>"
		 elseif request("flag")="modifyitem" then
           if Request("content")="" then ks.die "<script>alert('请输入服务内容!');history.back();</script>"
           if not isdate(Request("adddate")) then ks.die "<script>alert('时间格式不正确，请重输!');history.back();</script>"
		   rs.open "select top 1 * from ks_orderservice where id="&KS.ChkClng(request("itemid")),conn,1,3
			 rs("content")=request("content")
			 rs("adddate")=request("adddate")
			 rs("qsr")=request("qsr")
		   rs.update
		   rs.close
		   set rs=nothing
		   KS.Die "<script>alert('恭喜，服务记录修改成功!');location.href='" & Request.ServerVariables("HTTP_REFERER") &"';</script>"
		 elseif request("flag")="delitem" then
		  conn.execute("delete from ks_orderservice where id=" & ks.chkclng(request("itemid")))
		   KS.Die "<script>alert('恭喜，服务记录删除成功!');location.href='" & Request.ServerVariables("HTTP_REFERER") &"';</script>"
		 end if
		 
		 RS.OPEN "SELECT TOP 1 * From KS_Order Where id=" & id,conn,1,1
		 if RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   ks.die "<script>alert('出错!');top.box.close()</script>"
		 End If
		 Dim ServiceName,IsService,ServiceTimes,Validity,firstservicetime
		 ServiceName=rs("ServiceName")
		 IsService=rs("IsService")
		 ServiceTimes=rs("ServiceTimes")
		 Validity=rs("Validity")
		 RS.Close
		 RS.Open "select top 1 * from ks_orderservice where orderid=" & id & " order by id",conn,1,1
		 if not rs.eof then
		  firstservicetime=year(rs("adddate")) &"-" & month(rs("adddate")) & "-" & day(rs("adddate"))
		 else
		  firstservicetime="无"
		 end if
		 rs.close
		%>
		<div style="background:#fff;height: 423px;">
		<table width='99%' border='0' align='center' cellpadding='0' cellspacing='0' class='border' style="background:#fff;"> 		 
		 <tr align='center' class='title' height='25'>
		 <td><b>配送服务明细清单</b></td></tr>
		 <tr>
		  <td style="padding:10px;">
			<form name="myform" action="KS.ShopOrder.asp" method="post"/>
		    <div style="margin-bottom:10px;border:1px solid #999;padding:5px">
			<input type="hidden" name="id" value="<%=id%>"/>
			<input type="hidden" name="action" value="addservice"/>
			<input type="hidden" name="flag" value="dosave"/>
			 <div class="pt10">配送服务名称：<input type="text" name="servicename" id="servicename" value="<%=ServiceName%>" class="textbox" size="30"/></div>
			 <div class="pt10">是否开启配送记录：<input type="checkbox" onclick="if(this.checked){$('#service').show();}else{$('#service').hide();}" value="1" name="IsService"<%if IsService="1" then response.write " checked"%>/></div>
			 <div class="pt10">配送次数：<input type="text" name="ServiceTimes" id="ServiceTimes" value="<%=ServiceTimes%>" class="textbox" size="5" style="text-align:center"/>次&nbsp;&nbsp;&nbsp; 第一次配送：<%=firstservicetime%> &nbsp;&nbsp;&nbsp; 有效期：<input type="text" name="Validity" id="Validity" value="<%=Validity%>" class="textbox" size="5" style="text-align:center"/>月</div>
			</div>  
			 <div style="text-align:center;margin-bottom:10px;">
			   <input type="submit" value="确定保存修改" onclick="return(checkform())" class="button"/> <Input type="button" class="button" value="关闭取消" onclick="top.box.close();"/>
			 </div>
			</form>
			 
			 <script type="text/javascript">
			  function checkform(){
			    if ($("#servicename").val()==''){
				 alert('请输入服务名称!');
				 $("#servicename").focus();
				 return false;
				}
				if ($("#ServiceTimes").val()==''){
				 alert('请输入配送次数!');
				 $("#ServiceTimes").focus();
				 return false;
				}
				if ($("#Validity").val()==''){
				  alert('请输入有效期!');
				  $("#Validity").focus();
				  return false;
				}
				return true;
			  }
			  function modifyitem(id){
			    $("#tr1"+id).hide();
			    $("#tr2"+id).show();
			  }
			  function checkitem(id){
			    if ($("#content"+id).val()==''){
				  alert('请输入服务内容!');
				  $("#content"+id).focus();
				  return false;
				}
			    if ($("#adddate"+id).val()==''){
				  alert('请输入服务时间!');
				  $("#adddate"+id).focus();
				  return false;
				}
				return true;
			  }
			 </script>
			 <table cellpadding="0" id="service" style="<%if IsService<>"1" then response.write "display:none;"%>border:1px solid #999;" cellspacing="0" width="100%">
			   <tr style="background:#f1f1f1;height:23px;text-align:center">
			      <td style="padding:10px;">次数</td>
				  <td style="padding:10px;">内容</td>
				  <td style="padding:10px;">时间</td>
				  <td style="padding:10px;">签收人</td>
				  <td style="padding:10px;">操作</td>
			   </tr>
			   <%
			   RS.Open "select * from ks_orderservice where orderid=" & id & " order by id desc",conn,1,1
			   if Not RS.Eof Then
			    dim totalnum:totalnum=rs.recordcount
				dim str,num
				num=0
				do while not rs.eof
			    str=str &"<tr id='tr1" & rs("id") & "'>"
				str=str &"<td class=""splittd"">第" & totalnum-num & "次</td>"
				str=str &"<td class=""splittd"" style=""width:290px;word-break:break-all;"">" & rs("content") & "</td>"
				str=str &"<td class=""splittd"" style='text-align:center'>" & year(rs("adddate")) & "-" & month(rs("adddate")) & "-" & day(rs("adddate")) & "</td>"
				str=str &"<td class=""splittd"" style='text-align:center'>&nbsp;" & rs("qsr") & "&nbsp;</td>"
				str=str &"<td class=""splittd"" style='text-align:center'><a href=""javascript:modifyitem(" & rs("id") &");"">修改</a> <a href=""?action=addservice&flag=delitem&itemid=" & rs("id") & "&id=" & id &""" onclick=""return(confirm('删除后不可恢复，确定删除吗？'));"">删除</a></td>"
				str=str &"</tr>"
			    str=str &"<form name='form" & rs("id") & "' method='post' action='KS.ShopOrder.asp'/><input type='hidden' name='id' value='" & id & "'/><input type='hidden' name='itemid' value='" & rs("id") & "'/><input type='hidden' name='action' value='addservice'/><input type='hidden' name='flag' value='modifyitem'/>"
				str=str & "<tr style='display:none' id='tr2" & rs("id") &"'>"
				str=str &"<td class=""splittd"">第" & totalnum-num & "次</td>"
				str=str &"<td class=""splittd""><textarea id='content" & rs("id") & "' name='content' class='textbox'  style='line-height:30px;'/>" & rs("content") &"</textarea></td>"
				str=str &"<td class=""splittd""><input type='text' id='adddate" & rs("id") & "' name='adddate' class='textbox' value='" &rs("adddate") & "'/></td>"
				str=str &"<td class=""splittd""><input type='text' name='qsr' value='" & rs("qsr") & "' class='textbox'/></td>"
				str=str &"<td class=""splittd"" style='text-align:center'><input type='submit' onclick=""return(checkitem(" & rs("id") & "))"" value='保存' class='button'/></td>"
				str=str &"</tr></form>"
			    num=num+1
				rs.movenext
				loop
			  end if
			  rs.close
			    str="<form name='form1' method='post' action='KS.ShopOrder.asp'/><input type='hidden' name='id' value='" & id & "'/><input type='hidden' name='action' value='addservice'/><input type='hidden' name='flag' value='additem'/><tr><td class=""splittd"">第" & totalnum+1 & "次</td><td class=""splittd""><textarea name='content' class='textbox' style='line-height:30px;'/></textarea></td><td class=""splittd""><input type='text' name='adddate' class='textbox' value='" & year(now) & "-" & month(now) & "-" & day(now) & "'/></td><td class=""splittd""><input type='text' name='qsr' class='textbox'/></td><td class=""splittd"" style='text-align:center'><input type='submit' value='确定' class='button'/></td></tr></form>" & str
				
				response.write str
				
			  %>
			 </table>
			 
		  </td>
		 </tr>
		</table>
		</div>
 		<%
		End Sub
		
		Sub FinishOrder()
		 dim totalscore,AllianceUser,orderid,username,scoretf,DeliverStatus,paystatus,usescore
		 dim id:id=KS.ChkClng(Request("id"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "select top 1 * from ks_order where id=" & id ,conn,1,1
		 If RS.Eof And RS.Bof Then
		   rs.close:set rs=nothing
		   KS.AlertHintScript "出错啦，找不到订单！"
		 End If
		 totalscore=rs("totalscore")
		 orderid=rs("orderid")
		 username=rs("username")
		 scoretf=rs("scoretf")
		 DeliverStatus=rs("DeliverStatus")
		 paystatus=KS.ChkClng(rs("paystatus"))
		 usescore=KS.ChkClng(rs("usescore"))
		 rs.close
		 if totalscore>0 and scoretf="0" and DeliverStatus<>3  then
		    Call KS.ScoreInOrOut(username,1,totalscore,"系统","商城购物赠送的积分，订单号：" & orderid & "。",0,0)
			rs.open "select top 1 AllianceUser from ks_user where username='" & username & "'",conn,1,1
			if not rs.eof then
		    AllianceUser=rs("AllianceUser")
			end if
			rs.close
			if not ks.isnul(AllianceUser) then
			  rs.open "select top 1 groupid from ks_user where username='" & AllianceUser &"'",conn,1,1
			  if not rs.eof then
			    if KS.U_S(rs("GroupID"),19)="1"  then   '享受推广获积分
				   dim per:per=KS.U_S(rs("GroupID"),20)
				   if not isnumeric(per) then per=0
				   if per>0 then
				      dim myscore:myscore=KS.ChkClng(totalscore*per/100)
					  if myscore>0 then
					   	Call KS.ScoreInOrOut(AllianceUser,1,myscore,"系统","您推荐的用户[" & UserName & "]在商城购物成功,订单号：" & orderid & "，您享受该订单总赠送积分(" & totalscore & "分)的 " & per& "% 奖励。",0,0)

					  end if
				   end if
				end if
			  end if
			  rs.close
			end if
		 elseif paystatus=3 or DeliverStatus=3 and usescore>0 then  '退货或是退款时返还积分
			Session("ScoreHasUse")="-" '设置只累计消费积分
			Call KS.ScoreInOrOut(UserName,1,usescore,"系统","购物失败，返还积分。订单号<font color=red>" & orderid & "</font>!",0,0)
		 end if
		
		 
		 set rs=nothing
		 Conn.Execute("update ks_order set status=2,scoretf=1 where id=" & id)
		

		 
		 KS.Die "<script>alert('恭喜，订单已结清!');location.href='KS.ShopOrder.asp?Action=ShowOrder&ID=" & KS.G("ID") & "';</script>"
		End Sub
		
		'修改商品
		Sub modifyproduct()
		  If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');top.box.close();</script>"
			response.end
		  End If
		 Dim RSI,OrderID
		 OrderID=KS.G("ID")
		 %>
         <div style="background:#fff;">
		 <table width='100%' border='0' style='text-align:center' cellpadding='0' cellspacing='0' class='border'> 		  
		 <tr style='text-align:center' class='title' height='25'>  		   
		   <td><b>商 品 名 称</b></td> 		   
		   <td width='45'><b>单位</b></td>  		   
		   <td width='55'><b>数量</b></td>  		   
		   <td width='65'><b>原价</b></td>  		   
		   <td width='65'><b>实价</b></td>  		   
		   <td width='85'><b>小计</b></td>   		   
		   <td width='45'><b>操作</b></td>  		  
		  </tr> 
		  <form name="myform" action="KS.ShopOrder.asp" method="post">
		  <input type="hidden" name="action" value="doModifyProductSave"/>
		  <input type="hidden" name="orderid" value="<%=orderid%>"/>
		 <%
		 Dim SQLStr
		 SQLStr="Select i.*,P.Title,P.Unit From KS_OrderItem I Left Join KS_Product P  On I.ProID=P.ID Where I.SaleType<>5 and I.SaleType<>6 and I.OrderID='" & OrderID & "' order by i.ischangedbuy,i.id"
		 Set RSI=Server.CreateObject("ADODB.RECORDSET")
		 RSI.Open sqlstr,conn,1,1
		 If RSI.Eof And RSI.Bof Then
		    Response.Write "<tr class='tdbg'><td colspan=10 style='text-align:center'>该订单里没有商品!</td></tr>"
		 Else
		   Do While Not RSI.Eof
		   %>
		   <tr valign='middle' class='tdbg' height='20'>	  
		    <td width='*'><%=RSI("Title")%>
			 <br/>
			 <font color=#999999>属性：<input style="color:#999;width:200px;border:1px solid #ccc" class="textbox" type="text" name="attr<%=rsi("id")%>" value="<%=RSI("AttributeCart")%>"/>
			 </font>
			 </td> 
			<td width='45' style='text-align:center'><%=RSI("Unit")%></td>
			<td width='55' style='text-align:center'>
			 <input type="hidden" value="<%=rsi("id")%>" name="id" />
			<input class="textbox" type="text" name="amount<%=rsi("id")%>" value="<%=RSI("Amount")%>" size="4" style="text-align:center"></td>
			<td width='65' style='text-align:center'><input class="textbox" type="text" name="price_original<%=rsi("id")%>" value="<%=RSI("Price_Original")%>" size="5"/></td>    	   
			<td width='65' style='text-align:center'><input class="textbox" type="text" name="realprice<%=rsi("id")%>" value="<%=rsi("RealPrice")%>" size="5"/></td>    	   
			<td width='85' style='text-align:right'><%=formatnumber(rsi("realprice")*rsi("amount"),2,-1,-1)%> 元</td>
			 <td  style='text-align:center' width='45'>
			  <a href="?action=delproduct&orderid=<%=rsi("orderid")%>&id=<%=rsi("id")%>" onclick="return(confirm('确定将该商品从本订单中移除吗?'))">删除</a>
			 </td>  	   
			 </tr>
		   <%
		   RSI.MoveNext
		   Loop
		 End If
		 %>
		 <tr class="tdbg">
		   <td colspan=8>
		     <input type="submit" value="批量修改" class="button" /> <font color="blue">说明：批量修改将会重新计算订单的运费，订单总额等。</font>
		   </td>
		  </tr>
		  </form>
		 </table>

		<script type="text/javascript">
		  function getProduct()
		  {			 
		     $(parent.document).find("#ajaxmsg").toggle("fast");
			 var key=escape($('input[name=key]').val());
			 var tid=$('#tid>option:selected').val();
			 var priceType=$('#PriceType>option:selected').val();
			 var minPrice=$("#minPrice").val();
			 var maxPrice=$("#maxPrice").val();
			 var str='';
			 if (key!=''){
			   str='商品名称:'+key;
			 } 
			 if (tid!=''){
			   str+=' 栏目:'+$('#tid>option:selected').get(0).text
			 }
			 if (priceType!=0){
			   str+= minPrice +' 元';
			   switch (parseInt(priceType)){
			     case 1 :
				  str+='<=当前零售价<=';
				  break;
			     case 2 :
				   str+='<=会员价<=';
				   break;
			     case 3 :
				  str+='<=原始零售价<=';
				  break;
			   }
			   str+= maxPrice +' 元';
			   
			 }
			 if (str!='') str='<strong>条件:</strong><font color=red>'+str+'</font>';
			 $("#keyarea").html(str);
			 $.get("../../plus/ajaxs.asp", { action: "GetPackagePro", proid:$("#proids").val(),pricetype:priceType,key: key,tid:tid,minPrice:minPrice,maxPrice:maxPrice},
			 function(data){
					$(parent.document).find("#ajaxmsg").toggle("fast");
					$("#prolist").empty().append(data);
			  });
		  }
		</script>
		<div style="border:1px dashed #cccccc;margin-top:20px;padding:4px">
		<table width="100%" border="0">
		  <tr>
			<td style="text-align:left">
			  <div class="pt10">&nbsp;<strong>快速搜索=></strong></div>
			  <div class="pt10">&nbsp;商品编号: <input type="text" class="textbox" name="proids" id="proids" size='15'> 可留空</div>
			  <div class="pt10">&nbsp;商品名称: <input type="text" class='textbox' name="key"></div>
			  <div class="pt10">&nbsp;所属栏目: <select size='1' name='tid' id='tid' class="textbox"><option value=''>--栏目不限--</option><%=KS.LoadClassOption(5,false)%></select></div>
			  <div class="pt10">&nbsp;价格范围:
			<input type='text' name='minPrice' size='5' style='text-align:center' id='minPrice' value='10' class="textbox"> 元
			<= <select name="PriceType" id="PriceType" class="textbox">
			  <option value=0>--不限制--</option>
			  <option value=1>当前零售价</option>
			  <option value=2>会员价</option>
			  <option value=3>原始零售价</option>
			 </select>
			 <= <input type='text' name='maxPrice' size='5' style='text-align:center' id='maxPrice' value='100' class="textbox"> 元
			  
			  </div>
			  <div class="pt10">&nbsp;<input type="button" onclick="getProduct()" value="开始搜索" class="button" name="s1"></div>
			
			</td>
			<form name="myform" id="myform" action="KS.ShopOrder.asp?action=ProAddToOrder" method="post">
		    <input type="hidden" name="orderid" value="<%=orderid%>" class="textbox"/>
			<td>
				<div id='keyarea'></div>
				<div class="pt10"><strong>查询到的商品:</strong></div>		
				<div class="pt10">
				 <select name="prolist" size="5" style="width:260px;height:140px" multiple="multiple" id="prolist"></select>
				</div>
				<div class="pt10"><input type="submit" value="将选中的商品加入到本订单" class="button"></div>
			</td>
			</form>
		  </tr>
		</table>
		 </div>
		 </div>
		 <%RSI.Close
		 Set RSI=Nothing
		End Sub
		
		'保存修改
		Sub doModifyProductSave()
		 dim orderid:orderid=ks.s("orderid")
		 dim id:id=ks.filterids(ks.s("id"))
		 if id="" then ks.alerthintscript "没有商品!"
		 dim idarr,i
		 idarr=split(id,",")
		 for i=0 to ubound(idarr)
		    conn.execute("update ks_orderitem set amount=" & KS.G("amount" & trim(IDArr(i))) & ",price_original=" & KS.G("price_original"&Trim(IDArr(i))) &",realprice=" & KS.G("realprice"&Trim(IDArr(i))) & ",AttributeCart='" & KS.G("Attr" & trim(IDArr(i))) & "' Where ID=" & IDArr(i))
		 next
		 call updateorderprice(orderid)
		 KS.Die "<script>alert('恭喜，订单商品修改成功');parent.frames['MainFrame'].location.reload();top.box.close();</script>"
		 
		End Sub
		
		'商品加入订单
		Sub ProAddToOrder()
		 dim orderid:orderid=ks.g("orderid")
		 dim prolist:prolist=ks.filterids(ks.g("prolist"))
		 if orderid="" then ks.die "error!"
		 if ks.isnul(prolist) then ks.alerthintscript "对不起，您没有选择商品!"
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select * from ks_product where id in("&prolist&")",conn,1,1
		 if not rs.eof then
			 do while not rs.eof 
			  dim rsi:set rsi=server.CreateObject("adodb.recordset")
			  rsi.open "select top 1 * from ks_orderitem where proid=" & rs("id"),conn,1,3
			  if rsi.eof then
				  rsi.addnew
				  rsi("orderid")=orderid
				  rsi("proid")=rs("id")
				  rsi("Price_Original")=RS("Price")
				  rsi("Price")=RS("Price_member")
				  rsi("IsChangedBuy")=0
				  rsi("LimitBuyTaskID")=0
				  rsi("IsLimitBuy")=0
				  rsi("RealPrice")=RS("Price_Member")
				  rsi("Amount")=1
				  rsi("AttributeCart")=""
				  rsi("TotalPrice")=RS("Price_Member")
				  rsi("BeginDate")=Now
				  rsi("ServiceTerm")=RS("ServiceTerm")
				  rsi("PackID")=0
				  rsi("BundleSaleProID")=0
				  rsi.update
			 end if
			 rsi.close:set rsi=nothing
			 rs.movenext
			 loop 
			 call updateorderprice(orderid)
		 end if
			 rs.close
			 set rs=nothing
		 ks.alertHintscript "恭喜，已成功将选中的商品加入订单中!"
		End Sub
		
		Sub delproduct()
		 If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');</script>"
			response.end
		  End If
		  dim id:id=KS.ChkClng(KS.S("ID"))
		  dim orderid:orderid=ks.s("orderid")
		  Conn.Execute("Delete From KS_OrderItem Where ID=" & ID)
			 call updateorderprice(orderid)
		 ks.alertHintscript "恭喜，已成功将选中的商品从订单中移除!"
		End Sub
		
		'打印快递单
		Sub PrintExpress()
		 dim id:id=KS.ChkClng(Request("id"))
		 if id=0 then ks.die "error!"
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_order where id=" & id,conn,1,1
		 if rs.eof and rs.bof then
		   rs.close : set rs=nothing
		   ks.die "error!"
		 end if
		
		 dim rss:set rss=server.CreateObject("adodb.recordset")
		 dim tid:tid=KS.ChkClng(request("tid"))
		 if tid=0 then 
		   rss.open "select top 1 * from ks_shopexpress where expressid=" & rs("DeliverType"),conn,1,1
		   if not rss.eof then
		    tid=rss("id")
		   end if
		   rss.close
		 End If
		 
		 rss.open "select shopsetting from KS_Config",conn,1,1
		 dim setting:setting=split(rss(0)&"^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^","^#^")
		 rss.close
		 rss.open "select top 1 * from KS_ShopExpress Where ID=" & tid,conn,1,1
		 if rss.eof and rss.bof then
		  rss.close 
		  rss.open "select top 1 * from KS_ShopExpress where status=1",conn,1,1
		 end if
		 if rss.eof and rss.bof then
		   rss.close
		   ks.die  "<script>alert('对不起，您没有添加任何快递单模板,按确定进入添加!');location.href='KS.ShopExpress.asp';</script>"
		 end if
		 dim template:template=rss("template")
		 dim photourl:photourl=rss("photourl")
		 rss.close
		 '替换寄件人标签
		 template=replace(template,"{$寄件人_单位}",setting(0))
		 template=replace(template,"{$寄件人_姓名}",setting(1))
		 template=replace(template,"{$寄件人_地址}",setting(2))
		 template=replace(template,"{$寄件人_邮编}",setting(3))
		 template=replace(template,"{$寄件人_手机}",setting(4))
		 template=replace(template,"{$寄件人_电话}",setting(5))
		 template=replace(template,"{$寄件人_始发地}",setting(6))
		 '替换收件人标签
		 template=replace(template,"{$收件人_姓名}",rs("contactman"))
		 template=replace(template,"{$收件人_地址}",rs("address"))
		 template=replace(template,"{$收件人_电话}",rs("phone"))
		 template=replace(template,"{$收件人_手机}",rs("mobile"))
		 template=replace(template,"{$收件人_邮编}",rs("zipcode"))
		 template=replace(template,"{$收件人_目的地}",rs("tocity"))
		 template=replace(template,"{$年}",year(now))
		 template=replace(template,"{$月}",right("0"&month(now),2))
		 template=replace(template,"{$日}",right("0"&day(now),2))
		 
		 template=replace(template,"{$订单_备注留言}",rs("remark"))
		 template=replace(template,"{$订单_总金额}",rs("NoUseCouponMoney"))
		 
		%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
	<meta http-equiv="content-type" content="text/html;charset=utf-8">
	<head>
	<link href="../include/Admin_Style.css" type=text/css rel=stylesheet>
	<style type="text/css">
	body{margin:0px;padding:0px;}
	.mo{display:block;
	border:0px solid #ff6600;
    padding:0px;
    height:22px;
	font-size:20px;
	line-height:22px;
    position:absolute;
	}
	 .box{border:1px solid #cccccc;width:874px;height:483px;background:url(../../shop/express/<%=photourl%>) no-repeat;}
	 .bar{font-size:12px;border:1px solid #ccc;margin-bottom:6px;margin-top:5px;padding:2px;background:#f1f1f1;height:25px;width:870px;}
	 .noprint{border-radius: 0 0 5px 5px;margin: 0 20px 20px;background: #fff;padding: 0 20px 20px;border: 0; width:874px;}
	 #mybody{padding: 20px 20px 0px 20px;background-color: #fff;margin: 20px 20px 0;border-radius: 5px 5px 0px 0px;border: none;background-position: 20px;}
	 @media print {     
            .noprint{display: none; }     
    }     
    </style>
	<script type="text/javascript" src="../../ks_inc/jquery.js"></script>
	<script type="text/javascript">
	 function changesize(v){
	   if (v==0) return;
$("#mybody").find("label").each(function(){
       $(this)[0].style.fontSize=v+"px"; 
});	   
	   
	 }
	</script>
	</head>
	<body>
	
	<div id="mybody" class="box"><%=template%></div>
	<div class="noprint bar">
		<strong>选择快递单模板：</strong><select onChange="if(this.value!='0'){location.href='?id=<%=id%>&action=printexpress&tid='+this.value}" name="tid">
		 <option value='0'>请选择...</option>
		 <%rss.open "select * from KS_ShopExpress Where status=1 order by id",conn,1,1
		 do while not rss.eof
		   if tid=rss("id") then
		   response.write "<option value='" & rss("id") & "' selected>" & rss("title") & "</option>"
		   else
		   response.write "<option value='" & rss("id") & "'>" & rss("title") & "</option>"
		   end if
		 rss.movenext
		 loop
		 rss.close
		 set rss=nothing
		 %>
		</select>
		打印字号：<select name="psize" onChange="changesize(this.value);">
		 <option value='0'> 请选择字号...</option>
		 <%
		 dim n
		  for n=8 to 50  
		   if n=20 then
		   response.write "<option value='" & n & "' selected>" & n & " px</option>"
		   else
		   response.write "<option value='" & n & "'>" & n & " px</option>"
		   end if
		   next
		 %>
		</select>

		<input type="button" value=" 开始打印 "  onclick="window.print()" class="button"/>
    </div>
 </body>
</html>
		<%
		End Sub
		
		
		'修改送货信息
		Sub modifyinfo()
		If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');top.box.close();</script>"
			response.end
		  End If
		 dim id:id=KS.ChkClng(Request("id"))
		 if id=0 then ks.die "error!"
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_order where id=" & id,conn,1,1
		 if rs.eof and rs.bof then
		   rs.close : set rs=nothing
		   ks.die "error!"
		 end if
		%>
		
			<table border="0" cellpadding="2" cellspacing="1" class="border" width="100%" style="background:#fff;">
			<form name="myform" action="KS.ShopOrder.asp" method="post">
				<tr align="middle" class="title">
					<td colspan="2" height="25">
						<b>修 改 送 货 资 料</b></td>
				</tr>
				<tr class="tdbg">
					<td width="15%" style="text-align:right;">
						收货人：</td>
					<td><input type="text" name="contactman" class="textbox" maxlength="20" value="<%=rs("contactman")%>"/></td>
				</tr>
				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						收货地址：</td>
					<td><input type="text" name="address" class="textbox" maxlength="120" value="<%=rs("address")%>"/>
					
					邮政编码：<input type="text" name="zipcode" class="textbox" maxlength="20"  size="10" value="<%=rs("zipcode")%>"/>
					</td>
				</tr>

				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						联系电话：</td>
					<td>
						<input type="text" name="phone" class="textbox" maxlength="20" value="<%=rs("phone")%>"/>
						
						联系手机：<input type="text" class="textbox" name="mobile" maxlength="20" value="<%=rs("mobile")%>"/></td>
				</tr>

				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						电子邮件：</td>
					<td>
						<input type="text" name="email" class="textbox" maxlength="60" value="<%=rs("email")%>"/>
						联系QQ：<input type="text" name="qq" class="textbox" maxlength="20" value="<%=rs("qq")%>"/>
						</td>
				</tr>
          <%if rs("tocity")<>"" then%>
				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						发货方式：</td>
					<td>
					   <style>
					   	  .provincename{color:#ff6600}
						  .tocity{border:1px solid #006699;text-align:center;background:#C6E7FA;height:23px;width:130px;}
						  .showcity{position:absolute;background:#C6E7FA;border:#278BC6 1px solid;width:340px;display:none;height:230px;overflow-y:scroll;overflow-x:hidden;} 
						  .delivery{width:530px;padding:5px;margin-left: 5px;border:1px solid #cccccc;background:#f1f1f1}
						  .jgxx{color:#ff3300}
						  .jgxx span{color:blue}
						 </style>
							 <script type="text/javascript">
								  function ajshowdata(city)
									{ 
											  $.get("../../shop/ajax.getdate.asp",{city:escape(city),expressid:$("#DeliverType option:selected").val()},function(d){
											  var r=unescape(d).split('|');
											  if (r[0]=='error'){
											   alert(r[1]);
											   $("#jgxx").html('选择发往路线确定运费!');
											   $("#tocity").val('');
											  }else{ 
											   $("#jgxx").html(r[1]);
											   $("#tocity").val(city);
											   }
											  });
									} 
                                   $(document).ready(function(){
								   ajshowdata('<%=rs("tocity")%>');
								   })
							  </script>
						<div class="delivery">			  
						<%=GetDeliveryTypeStr(rs("DeliverType"),rs("tocity"))%>
						</div>
						</td>
				</tr>
		<%end iF%>	
				
				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						付款方式：</td>
					<td>
						<%=GetPaymentTypeStr(rs("PaymentType"))%></td>
				</tr>
				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						发票信息：</td>
					<td>
						<input type="radio" onClick="$('#zzs').hide();$('#fp').hide();" name="NeedInvoice" <%if rs("NeedInvoice")=0 then response.write " checked"%> value=0>不需要发票
						<input type="radio" onClick="$('#zzs').hide();$('#fp').show();" name="NeedInvoice" <%if rs("NeedInvoice")=1 then response.write " checked"%> value="1">普通发票
						<input type="radio" onClick="$('#zzs').show();$('#fp').show();" name="NeedInvoice" <%if rs("NeedInvoice")=2 then response.write " checked"%> value="2">增值税发票
					
						<br/>
						<div id='fp' style="<%if rs("NeedInvoice")=0 then response.write "display:none"%>">
						单位名称<input name="InvoiceContent" class="textbox" value="<%=rs("InvoiceContent")%>">
						</div>
						<div id="zzs" style="<%if rs("NeedInvoice")<>2 then response.write "display:none"%>">
						纳税人识别码<input name="InvoiceCode" class="textbox" value="<%=rs("InvoiceCode")%>"><br/>
						注册地址<input name="InvoiceAddress" class="textbox" value="<%=rs("InvoiceAddress")%>"><br/>
						注册电话<input name="InvoiceTel" class="textbox" value="<%=rs("InvoiceTel")%>"><br/>
						开户银行<input name="InvoiceBank" class="textbox" value="<%=rs("InvoiceBank")%>"><br/>
						银行账号<input name="InvoiceBankCard" class="textbox" value="<%=rs("InvoiceBankCard")%>">
						</div>
						
						</td>
				</tr>
				<tr class="tdbg">
					<td style="text-align:right;" width="15%">
						备注留言：</td>
					<td>
						<textarea name="Remark" cols="40" rows="3"><%=rs("Remark")%></textarea></td>
				</tr>

				<tr align="middle" class="tdbg">
					<td colspan="2" height="30" style="text-align:center">
						<input id="Action" name="Action" type="hidden" value="DoModifyInfoSave" /> <input id="ID" name="ID" type="hidden" value="<%=id%>" /> <input class="button" name="Submit" type="submit" value="确定保存修改" />&nbsp;<input class="button" name="Submit" type="button" onClick="window.open('KS.ShopOrder.asp?action=printexpress&id=<%=id%>');" value="打印快递单" />&nbsp;<input class="button" name="Submit" onClick="javascript:top.box.close();" type="button" value="关闭取消" /></td>
				</tr>
			</form>
		</table>
		
		<%
		End Sub
		
		'保存修改
      Sub DoModifyInfoSave()
		Dim ID:ID=KS.ChkClng(KS.G("id"))
		If id=0 Then KS.Die "error!"
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		If Not RS.Eof Then
		  RS("ContactMan")=KS.G("ContactMan")
		  RS("Address")=KS.G("Address")
		  RS("ZipCode")=KS.G("ZipCode")
		  RS("Phone")=KS.G("Phone")
		  RS("Mobile")=KS.G("Mobile")
		  RS("Email")=KS.G("Email")
		  RS("qq")=KS.G("qq")
		  RS("PaymentType")=KS.ChkClng(KS.G("PaymentType"))
		  RS("ToCity")=KS.G("ToCity")
		  RS("NeedInvoice")=KS.ChKClng(KS.G("NeedInvoice"))
		  If KS.ChKClng(KS.G("DeliverType"))<>0 Then
		  RS("ToCity")=KS.G("ToCity")
		  RS("DeliverType")=KS.ChKClng(KS.G("DeliverType"))
		  End If
		  RS("InvoiceContent")=KS.G("InvoiceContent")
				RS("InvoiceCode")=KS.S("InvoiceCode")
				RS("InvoiceAddress")=KS.S("InvoiceAddress")
				RS("InvoiceTel")=KS.S("InvoiceTel")
				RS("InvoiceBank")=KS.S("InvoiceBank")
				RS("InvoiceBankCard")=KS.S("InvoiceBankCard")
		  RS("Remark")=KS.G("Remark")
		  RS.Update
		End If
		RS.Close :Set RS=Nothing
		KS.Die "<script>alert('恭喜，修改成功!');top.frames[""MainFrame""].location.reload();top.box.close();</script>"
  End Sub
		
  '付款方式
  Function GetPaymentTypeStr(PaymentType)
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
     If trim(SQL(0,I))=trim(PaymentType) Then
    GetPaymentTypeStr=GetPaymentTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & " " & DiscountStr & "</option>"
	 Else
    GetPaymentTypeStr=GetPaymentTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & " " & DiscountStr & "</option>"
	End If
   Next
   GetPaymentTypeStr=GetPaymentTypeStr & "</select>"
  End Function
	
	 '发货方式
  Function GetDeliveryTypeStr(typeid,tocity)
   Dim j,rss,rsss
   Dim DiscountStr,SQL,I,RS


   Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select TypeID,TypeName,IsDefault from KS_DeliveryType order by orderid,TypeID",conn,1,1
   If Not RS.Eof Then
     SQL=RS.GetRows(-1)
   End IF
   RS.Close:Set RS=Nothing
   GetDeliveryTypeStr="<strong>快递公司：</strong><select name='DeliverType' id='DeliverType'>"
   For I=0 To UBound(SQL,2)
     If trim(typeid)=trim(sql(0,i)) Then
    GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & "</option>"
	 Else
    GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & "</option>"
	End If
   Next
   GetDeliveryTypeStr=GetDeliveryTypeStr & "</select>"
   if tocity="" then tocity="选择送货地点"

   GetDeliveryTypeStr=GetDeliveryTypeStr & "<br/> <input type=""hidden"" name=""tocity"" id=""tocity""/> <span style='position:relative'><input class=""tocity"" style='text-align;left' name='' id='choosecity' type='button' value='" & tocity & "'  onclick=""showprovn.style.display='block';if(this.getBoundingClientRect().top>300){showprovn.style.top=(this.offsetHeight-showprovn.offsetHeight)}else{showprovn.style.top='0'}""><span id='showprovn' onclick=""this.style.display='none'"" class='showcity'>"&_
			 "<table width='92%' style='text-align:center' border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
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
						IF (j MOD 4) = 1 THEN GetDeliveryTypeStr=GetDeliveryTypeStr&"<tr>"&vbcrlf
						GetDeliveryTypeStr=GetDeliveryTypeStr&"<td id='ccity' onclick=""choosecity.value=this.innerHTML;ajshowdata(this.innerHTML)"" style='cursor:hand' onmouseover=""this.style.color='red'"" onmouseout=""this.style.color=''"">"&pnode.selectsinglenode("@city").text&"</td>"&vbcrlf
						if (j mod 4)=0 then GetDeliveryTypeStr=GetDeliveryTypeStr&"</tr>"&vbcrlf
						j=j+1
						Next
						
					 Next
					End If
 
			        
					 
			 GetDeliveryTypeStr=GetDeliveryTypeStr&"</table>"&vbcrlf&_
			"</span></span>"&_
		" <span id='jgxx' class='jgxx'>选择送货路线确定运费！</span>"&vbcrlf


  End Function
	
  '修改总价
  Sub ModifyTotalPrice()
          If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');top.box.close();</script>"
			response.end
		  End If
		  dim id:id=ks.chkclng(request("id"))
		  dim price:price=request("price")
		  dim oprice:oprice=request("oprice")
		  if id=0 then
		    response.write "<script>alert('参数出错!');</script>"
			response.end
		  end if
		  if not isnumeric(price) then
		    response.write "<script>alert('输入的价格不对,请输入正确的数字!');</script>"
			response.end
		  end if
		  if oprice=price then
		    response.write "<script>alert('价格与修改前一样,没有更新!');</script>"
			response.end
		  end if
		  conn.execute("update ks_order set moneytotal=" & price  & " where id=" & id)
		  response.write "<script>alert('恭喜,订单总价修改成功!');top.frames[""MainFrame""].location.reload();top.box.close();</script>"
  End Sub	

  '更新订单价格
  sub updateorderprice(orderid)
          dim totalrealprice:totalrealprice=0
		  Dim totalweight:totalweight=0
		  dim ordertype:ordertype=0
		  dim totalscore:totalscore=0
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select ordertype from ks_order where orderid='" & orderid & "'",conn,1,1
		  if not rs.eof then
		    ordertype=rs(0)
		  end if
		  rs.close
		  if ordertype=1 then
		  rs.open "select i.*,p.weight from ks_orderitem i left join ks_groupbuy p on i.proid=p.id where i.orderid='" & orderid & "'",conn,1,1
		  else
		  rs.open "select i.*,p.weight from ks_orderitem i left join ks_product p on i.proid=p.id where i.orderid='" & orderid & "'",conn,1,1
		  end if
		  do while not rs.eof
		    totalscore=rs("score")+totalscore
		    totalrealprice=totalrealprice+Round(rs("totalprice"),2)
			if isnumeric(rs("weight")) then
		    totalweight=totalweight+Round(rs("weight")*rs("amount"),2)
			end if
		  rs.movenext
		  loop
		  rs.close
		  if totalrealprice<>0 then
		    conn.execute("update ks_order set weight=" & totalweight & " where orderid='" & orderid & "'")
		    rs.open "select top 1 * from ks_order where orderid='" & orderid & "'",conn,1,3
			if not rs.eof then
			   rs("moneygoods")=totalrealprice
			   Dim TaxRate:TaxRate=KS.Setting(65)
			   Dim IncludeTax:IncludeTax=KS.Setting(64)
			   Dim TaxMoney,RealMoneyTotal,Freight
			   Freight=KS.GetFreight(RS("DeliverType"),RS("ToCity"),RS("weight"),"")
			   If IncludeTax=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+Taxrate/100
				'总金额 = (总价*付费方式折扣+运费)*(1+税率)
                if rs("PaymentType")=0 then
				RealMoneyTotal=Round((totalrealprice+Freight)*TaxMoney,2)
				else
				RealMoneyTotal=Round((totalrealprice*KS.ReturnPayment(rs("PaymentType"),1)/100+Freight)*TaxMoney,2)
				end if
				
				RS("Charge_Deliver")=Freight
			  rs("NoUseCouponMoney")=RealMoneyTotal
			  if rs("CouponUserID")<>0 then
			     'dim facevalue:facevalue=conn.execute("select facevalue from KS_ShopCoupon where id=" &rs("CouponUserID"))(0) 
			   	' If FaceValue>0 Then
				   RealMoneyTotal=Round(RealMoneyTotal-rs("usecouponmoney"),2)
				' End If
			  elseif rs("UseScoreMoney")>0 then
				   RealMoneyTotal=Round(RealMoneyTotal-rs("UseScoreMoney"),2)
			  end if
			  rs("MoneyTotal")=RealMoneyTotal
			  
			  if rs("UseCouponMoney")>0 or rs("UseScoreMoney")>0 then  '使用优惠券不送积分
			   rs("totalscore")=0
			  else
              rs("totalscore")=KS.ChkClng(totalscore)
			  end if
  
			   rs.update
			end if
			rs.close
		  end if
		  set rs=nothing
  end sub		
		
  '修改指定价
  sub ModifyPrice()
           If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单价格!');top.box.close();</script>"
			response.end
		   End If
		  dim id:id=ks.chkclng(request("id"))
		  dim price:price=request("price")
		  dim orderid:orderid=ks.g("orderid")
		  dim oprice:oprice=request("oprice")
		  if id=0 then
		    response.write "<script>alert('参数出错!');</script>"
			response.end
		  end if
		  if not isnumeric(price) then
		    response.write "<script>alert('输入的价格不对,请输入正确的数字!');</script>"
			response.end
		  end if
		  if oprice=price then
		    response.write "<script>alert('价格与修改前一样,没有更新!');</script>"
			response.end
		  end if
		  dim rs:set rs=server.createobject("adodb.recordset")
		  
		  dim username,groupid,JFDiscount
		  rs.open "select top 1 * from ks_order where orderid='" & orderid & "'",conn,1,1
		  if rs.eof and rs.bof then
		    response.write "<script>alert('订单已不存在了!');</script>"
			response.end
		  end if
		  username=rs("username")
		  rs.close
		  if not ks.isnul(username) then
		   rs.open "select top 1 groupid from ks_user where username='" & username & "'",conn,1,1
		   if not rs.eof then
		    groupid=rs(0)
		   end if
		   rs.close
		   JFDiscount=KS.U_S(GroupID,18)
		  end if
		   If Not IsNumeric(JFDiscount) Then JFDiscount=0
		  
		  rs.open "select top 1 * from ks_orderitem where id=" &id,conn,1,3
		  if not rs.eof then
		     rs("realprice")=price
			 rs("score")=KS.ChkClng(price*JFDiscount)* rs("amount")
			 rs("totalprice")=price * rs("amount")
			 rs.update
		  end if
		  rs.close
		  set rs=nothing
		  call updateorderprice(orderid)
		  response.write "<script>alert('恭喜,指定价修改成功!');top.frames[""MainFrame""].location.reload();top.box.close();</script>"
		end sub
		
	
		
		Sub OrderList()
%>
  <table width="100%" cellSpacing=0 cellPadding=0  border="0" class="hide">
    <tr>
      <td align=left height="28"><i class='icon mainer'></i> 您现在的位置：<a href="KS.ShopOrder.asp">订单处理</a>&nbsp;&gt;&gt;&nbsp;
	  <%
	     Dim SearchTypeStr,Keyword
		 Keyword=KS.G("Keyword")
	    Select Case SearchType
	    Case 0	SearchTypeStr= "所有订单"
		Case 1	SearchTypeStr= "24小时之内的新订单"
		Case 2	SearchTypeStr= "最近10天内的新订单"
		Case 3	SearchTypeStr= "最近一月内的新订单"
		Case 4	SearchTypeStr="未确认的订单"
		Case 5	SearchTypeStr="未付款的订单"
		Case 6	SearchTypeStr="未付清的订单"
		Case 7	SearchTypeStr="未送货的订单"
		Case 8	SearchTypeStr="未签收的订单"
		Case 9	SearchTypeStr="未开发票的订单"
		Case 10
		   Select Case  KS.ChkClng(KS.G("Field"))
		    Case 1:SearchTypeStr="订单编号含有<font color=red>""" & KeyWord & """</font>"
		    Case 2:SearchTypeStr="收货人含有<font color=red>""" & KeyWord & """</font>"
		    Case 3:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
		    Case 4:SearchTypeStr="联系地址含有<font color=red>""" & KeyWord & """</font>"
		    Case 5:SearchTypeStr="联系电话含有<font color=red>""" & KeyWord & """</font>"
		    Case 6:SearchTypeStr="下单时间含有<font color=red>""" & KeyWord & """</font>"
		    Case 7:SearchTypeStr="推荐人为<font color=red>""" & KeyWord & """</font>"
		   End Select
		Case 11	SearchTypeStr="未结清的订单"
		Case 12	SearchTypeStr="已结清的订单"
		Case 13 SearchTypeStr="需要服务跟踪"
		Case 14 SearchTypeStr="团购订单"
		Case 15 SearchTypeStr="积分兑换订单"
		End Select
		Response.Write SearchTypeStr
	  %>
	  </td>
    </tr>
  </table>
  
 
<div class="tabs_header">
    <ul class="tabs">
    <li<%if SearchType=0 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?<%=KS.QueryParam("SearchType")%>"><span>所有订单</span></a></li>
    <li<%if SearchType=1 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=1&<%=KS.QueryParam("SearchType")%>"><span style='color:red'>24小时内新订单</span></a></li>
    <li<%if SearchType=14 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=14&<%=KS.QueryParam("SearchType")%>"><span>团购订单</span></a></li>
    <li<%if SearchType=15 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=15&<%=KS.QueryParam("SearchType")%>"><span>积分兑换订单</span></a></li>
    <li<%if SearchType=7 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=7&<%=KS.QueryParam("SearchType")%>"><span>未送货订单</span></a></li>
    <li<%if SearchType=8 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=8&<%=KS.QueryParam("SearchType")%>"><span>未签收的订单</span></a></li>
    <li<%if SearchType=11 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=11&<%=KS.QueryParam("SearchType")%>"><span>未结清订单</span></a></li>
    <li<%if SearchType=12 then response.write " class='active'"%>><a href="KS.ShopOrder.asp?SearchType=12&<%=KS.QueryParam("SearchType")%>"><span>已结清订单</span></a></li>
    </ul>
</div>
<div class="pageCont"> 
  <table cellSpacing=0 cellPadding=0  style="width:100%" border=0>
    <tr>
<FORM name=myform onSubmit="return confirm('确定要删除选定的订单吗？');" action=KS.ShopOrder.asp method=post>
      <td>
        <table cellSpacing="0" cellPadding="0" width="100%" border=0>
          <tr class=sort align=middle>
            <td>选中</td>
            <td>订单编号</td>
            <td nowrap="nowrap">客户</td>
            <td>用户名</td>
            <td>下单时间</td>
            <td>赠送积分</td>
           <!-- <td width=60>总金额</td>-->
            <td>应付金额</td>
            <td>已收金额</td>
            <td>发票</td>
            <td>已开</td>
            <td>状态</td>
            <td>付款</td>
            <td>物流状态</td>
            <td>服务跟踪</td>
          </tr>
		  <%
		  	MaxPerPage=20
			If KS.G("page") <> "" Then
				  CurrentPage = KS.ChkClng(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			
			SqlParam="1=1"

			
			If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1 SqlParam=SqlParam &" And datediff(" & DataPart_H & ",inputtime," & SqlNowString & ")<25"
			   Case 2 SqlParam=SqlParam &" And datediff(" & DataPart_D & ",inputtime," & SqlNowString & ")<=10"
			   Case 3 SqlParam=SqlParam &" And datediff(" & DataPart_D & ",inputtime," & SqlNowString & ")<=30"
			   Case 4:SqlParam=SqlParam &" And Status=0"
			   Case 5:SqlParam=SqlParam &" And MoneyReceipt=0"
			   Case 6:SqlParam=SqlParam &" And MoneyReceipt<=MoneyTotal"
			   Case 7:SqlParam=SqlParam &" And DeliverStatus=0"
			   Case 8:SqlParam=SqlParam &" And DeliverStatus=1"
			   Case 9:SqlParam=SqlParam &" And NeedInvoice=1 And Invoiced=0"
			   Case 10
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1 SqlParam=SqlParam &" And OrderID Like '%" & Keyword & "%'"
				   Case 2 SqlParam=SqlParam &" And ContactMan Like '%" & Keyword & "%'"
				   Case 3 SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 4 SqlParam=SqlParam &" And Address Like '%" & Keyword & "%'"
				   Case 5 SqlParam=SqlParam &" And Phone Like '%" & Keyword & "%'"
				   Case 6 SqlParam=SqlParam &" And InputTime Like '%" & Keyword & "%'"
				   Case 7 SqlParam=SqlParam & " and UserName in(select username from ks_user where AllianceUser='" & KeyWord & "')"
				  End Select
			   Case 11:SqlParam=SqlParam &" And status=1"
			   Case 12:SqlParam=SqlParam &" And status=2"
			   Case 13:SqlParam=SqlParam &" And isservice=2"
			   Case 14:SqlParam=SqlParam &" And ordertype=1"
			   Case 15:SqlParam=SqlParam &" And UseScoreisshop>0"
			  End Select
			End If

		   Set RS=Server.CreateObject("ADODB.RECORDSET")
		   SqlStr="Select * From KS_Order where " & SqlParam & " order by inputtime desc"
		   RS.Open SqlStr ,Conn,1,1
		   If RS.Eof And RS.Bof Then
		    Response.Write "<tr class='list' onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'""><td height='30' colspan=15 style='text-align:center' class='splittd'>找不到" & SearchTypeStr & "!</td></tr>"
		  Else
		  	               totalPut = RS.RecordCount
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent()
		  End If
			 RS.Close:Set RS=Nothing
		%>
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td height=30 class="operatingBox">
              &nbsp;<label><Input id=chkAll onclick=CheckAll(this.form) type=checkbox value=checkbox name=chkAll> 选中本页显示的所有订单</label>
  <Input id=Action type=hidden value=DelOrder name=Action> 
              <Input type=submit value="删除选定的订单" class="button" name=Submit>
		   </td>
		   <td>
		   <%
		   	  '显示分页信息
			   Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		   %>
		   </td>
          </tr>
        </table>
		</FORM>
      </td>
    </tr>
  </table>
</div>
<div class="footerTable pt10">
	<div class="attention">
		<font color=red><strong>说明：</strong><br/>1、为便于销售统计已结清或已收到汇款(包括仅收到预付款)的订单不能删除;<br/>2、订单号后面有“团”字表示该订单是团购订单;<br/>
		3、订单号后面有“兑换”字表示该订单是积分兑换订单;</font>
	</div>
</div> 
</body>
<html>
		<%
		End Sub
		
		Sub ShowContent()
		      Dim I
			  Do While Not RS.Eof 
		   %>
			  <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
				<td class='splittd' height='25'><Input id=ID type=checkbox value="<%=rs("id")%>" name="ID"></td>
				<td class='splittd' style="text-align:left"><a href="KS.ShopOrder.asp?Action=ShowOrder&ID=<%=RS("ID")%>"><%=RS("OrderID")%></a>
				<%
				
				  if rs("ordertype")="1" then
				  response.write "<font color=red><b><i>团</li></b></font>"
				  end if
				  if KS.ChkClng(rs("UseScoreisshop"))>0 then
				 	 response.write "<font color=""006600""><b>兑换</b></font>"
				  end if
					 
					  %>
				</td>
				<td class='splittd'><%=RS("ContactMan")%></td>
				<td class='splittd'><%=RS("UserName")%></td>
				<td class='splittd' title="<%=RS("InputTime")%>"><%=formatdatetime(RS("InputTime"),2)%></td>
				<td class='splittd'><%
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
					    %></td>
<!--				<td class='splittd' style='text-align:right'>￥<%=RS("NoUseCouponMoney")%>元</td>-->
				<td  class='splittd'style='text-align:right'>￥<%If RS("MoneyTotal")<1 and RS("MoneyTotal")>0 then 
				  response.write "0" & RS("MoneyTotal")
				else
				 response.write RS("MoneyTotal")
				end if
				%>元
                <%if KS.ChkClng(rs("UseScoreisshop"))>0 then
					 	Response.Write "<B>+</b> "& KS.ChkClng(rs("UseScoreisshop")) &" 积分"
					 end if%>
                </td>
				<td  class='splittd'style='text-align:right'><font color=red><%If RS("MoneyReceipt")<1 and RS("MoneyReceipt")>0then 
				  response.write "0" & RS("MoneyReceipt")
				else
				 response.write RS("MoneyReceipt")
				end if
				%></font></td>
				<td class='splittd'>
				<%If RS("NeedInvoice")=1 Then
				  Response.Write "<Font color=red>√</font>"
				  Else
				   Response.Write "&nbsp;"
				  End If
				  %>
				</td>
				<td class='splittd'>
				<%
				if RS("NeedInvoice")=1 Then
				  If RS("Invoiced")=1 Then
				   Response.Write "<font color=green>√</font>"
				  Else
				   Response.Write "<font color=red>×</font>"
				  End If
				Else
				  Response.Write "&nbsp;"
				End If
				 %>
				</td>
				<td class='splittd'>
				<%If RS("Status")=0 Then
				  Response.Write "<font color=red>等待确认</font>"
				  ElseIf RS("Status")=1 Then
				  Response.WRITE "<font color=green>已经确认</font>"
				  ElseIf RS("Status")=2 Then
				  Response.Write "<font color=#a7a7a7>已结清</font>"
				  ElseIf RS("Status")=3 Then
				  Response.Write "<font color=#a7a7a7>无效订单</font>"
				  End If
				%>
				  </td>
				<td class='splittd'>
				<%
				if rs("alipaytradestatus")<>"" and RS("Status")<>2 then
				  select case rs("alipaytradestatus")
				    Case "WAIT_BUYER_PAY" Response.Write "<font color=red>等待汇款</font>"
					Case "WAIT_SELLER_SEND_GOODS" Response.Write "<font color=brown>已付款等待发货</font>"
					Case "WAIT_BUYER_CONFIRM_GOODS" Response.Write "<font color=blue>等待买家确认收货</font>"
					Case "TRADE_FINISHED" Response.Write "<font color=#a7a7a7>交易完成</font>"
				  end select
				else
					if rs("paystatus")="100" then
					  Response.WRITE "<font color=""green"">凭单消费</font>"
					elseif rs("paystatus")="3" then
					  Response.WRITE "<font color=blue>退款</font>"
					elseif rs("paystatus")="1" then
					  Response.Write "<font color=green>已经付清</font>"
					elseIf RS("MoneyReceipt")<=0 Then
					   Response.Write "<font color=red>等待汇款</font>"
					ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
					   Response.WRITE "<font color=blue>已收定金</font>"
					Else
					   Response.Write "<font color=green>已经付清</font>"
					End If
				end if
				  %></td>
				<td class='splittd'>
				<% If RS("DeliverStatus")=0 Then
				 Response.Write "<font color=red>未发货</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  Response.Write "<font color=blue>已发货</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  Response.Write "<font color=green>已签收</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  Response.Write "<font color=#ff6600>退货</font>"
				ElseIf RS("DeliverStatus")=4 Then
				 Response.Write "<font color=brown>客户申请退货退款</font>"
				 End If
				 %></td>
				<td class='splittd'>
				<% If RS("isservice")="1" Then
				  Response.Write "<font color=blue>需要</font>"
				 Else
				  Response.Write "<font color=""#999999"">不需要</font>"
				 End If
				 %></td>
			  </tr>
			  <%
			    PageTotalMoney1=PageTotalMoney1+RS("MoneyTotal")
				PageTotalMoney2=PageTotalMoney2+RS("MoneyReceipt")
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do
			  Loop
		  %>
          <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
            <td  class='splittd' style='text-align:right' colSpan=6><B>本页合计：</B></td>
            <td  class='splittd' style='text-align:right'>￥<%=PageTotalMoney1%>元</td>
            <td  class='splittd' style='text-align:right'><%=PageTotalMoney2%></td>
            <td  class='splittd' colSpan=6>&nbsp;</td>
          </tr>
          <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
            <td class='splittd' style='text-align:right' colSpan=6><B>本次查询合计：</B></td>
            <td class='splittd' style='text-align:right'>￥<%=Conn.execute("Select Sum(MoneyTotal) From KS_Order where " & SqlParam)(0)%>元</td>
            <td class='splittd' style='text-align:right'><%=Conn.execute("Select Sum(MoneyReceipt) From KS_Order where " & SqlParam)(0)%></td>
            <td class='splittd' colSpan=6>&nbsp;</td>
          </tr>
          <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
            <td class='splittd' style='text-align:right' colSpan=6><B>总计金额：</B></td>
            <td class='splittd' style='text-align:right'>￥<%=Conn.execute("Select Sum(MoneyTotal) From KS_Order")(0)%>元</td>
            <td class='splittd' style='text-align:right'><%=Conn.execute("Select Sum(MoneyReceipt) From KS_Order")(0)%></td>
            <td class='splittd' colSpan=6>&nbsp;</td>
          </tr>
        </table>
		<%End Sub
		
		Sub ShowOrder()
		%>
		<div class="pageCont"> 
		<%
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * from ks_order where id=" & ID ,conn,1,1
		 IF RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   response.end
		 End If
		 
		response.write "<br>"
		response.write KS.OrderDetailStr(RS,1)
		%>

        <br>
   <div style='text-align:center'> 
           <% 
			 If RS("Status")=3 Then
			   response.write "本订单在指定时间内没有付款作废!"
			 ElseIf rs("DeliverStatus")=4 Then  '处理客户申请的退款
			 %>
			 <input type='button' class='button' name='Submit' value=' 同意退款，并结清订单 ' onClick="window.location.href='KS.ShopOrder.asp?Action=BankRefund&ID=<%=RS("id")%>'">&nbsp;
			 <input type='button' class='button' name='Submit'  value=' 已和客户妥协订单恢复正常 ' onClick="if (confirm('确定将本订单恢复正常状态吗？')){window.location.href='KS.ShopOrder.asp?Action=BankRefundOK&ID=<%=RS("orderid")%>'}">&nbsp;
			 <%
			 Else

		   If RS("Status")<>2 Then%>   
			 <% IF RS("Status")=0 Then%>
			 <input type='button' class='button' name='Submit' value='确认订单' onClick="javascript:if(confirm('请仔细检查此订单的所有信息，确认后将发送一封站内短信、手机短信及邮件通知客户!')){window.location.href='KS.ShopOrder.asp?Action=OrderConfirm&ID=<%=RS("ID")%>';}">&nbsp;&nbsp;
			 <%ElseIf RS("Status")=1 And RS("MoneyReceipt")=0 and rs("paystatus")<>"3" Then%>
			 <input type='button' class='button' name='Submit' value='删除订单' onClick="javascript:if(confirm('确定删除此订单吗!')){window.location.href='KS.ShopOrder.asp?Action=DelOrder&ID=<%=RS("ID")%>';}">&nbsp;&nbsp;
			 <%End iF%>
			 <%
			if rs("paystatus")<>"3"  then
			 If RS("MoneyReceipt")<RS("MoneyTotal") Then%>
			 <input type='button'class='button'  name='Submit' value='银行汇款支付' onClick="window.location.href='KS.ShopOrder.asp?Action=BankPay&ID=<%=RS("id")%>'">&nbsp;
			 <%Else%>
			 <input type='button' class='button' name='Submit' value=' 退款 ' onClick="window.location.href='KS.ShopOrder.asp?Action=BankRefund&ID=<%=RS("id")%>'">&nbsp;
			 <%End IF
			end if 
			%>
			 
			 <%If RS("NeedInvoice")=1 And RS("Invoiced")=0 Then%>
			 <input type='button' class='button' name='Submit' value=' 开发票 ' onClick="window.location.href='KS.ShopOrder.asp?Action=Invoice&ID=<%=RS("ID")%>'">&nbsp;
			 <%End IF%>
			 <%If RS("Status")=1 and RS("DeliverStatus")<>3 and RS("DeliverStatus")<>2 Then%>
			 <input type='button' class='button' name='Submit' value='客户已签收' onClick="if(confirm('确定客户已收到货了吗?')){window.location.href='KS.ShopOrder.asp?Action=ClientSignUp&ID=<%=RS("ID")%>';}">&nbsp;
			 <%End If
			 dim tipsstr,days:days=datediff("d",RS("Deliverydate"),now)
			 IF RS("Deliverydate")="2000-1-1" then tipsstr="" else tipsstr="已发货 " & days &" 天,"
			 
			 If RS("Status")<>2 Then
				 If RS("MoneyReceipt")>=RS("MoneyTotal") And RS("DeliverStatus")<>0 Then
				 %>
				  <%if rs("totalscore")>0 and rs("DeliverStatus")<>3 and rs("usescoremoney")<=0 then%>
				 <input type='button' class='button' name='Submit' value='<%=tipsstr%>结清订单,并赠送<%=rs("totalscore")%>分积分' onClick="if(confirm('订单一旦结算，该订单就不可进行任何操作，确定结清订单吗?')){window.location.href='KS.ShopOrder.asp?Action=FinishOrder&ID=<%=RS("ID")%>';}">&nbsp;<br/><br/>
				 <%else%>
				 <input type='button' class='button' name='Submit' value='<%=tipsstr%>结清订单' onClick="if(confirm('订单一旦结算，该订单就不可进行任何操作，确定结清订单吗?')){window.location.href='KS.ShopOrder.asp?Action=FinishOrder&ID=<%=RS("ID")%>';}">&nbsp;
				 <%end if%>
			  <%ElseIf RS("PayStatus")="3" or rs("DeliverStatus")=3 Then
			    if rs("usescore")>0 then
				%>
			   <input type="button" value="结清订单,返还用户 <%=rs("usescore")%>分积分" onClick="if (confirm('结清订单将立即返还用户的积分，确定结清吗？')){window.location.href='KS.ShopOrder.asp?Action=FinishOrder&ID=<%=RS("ID")%>';}" class="button" />
				<%
				else
				%>
			   <input type="button" value="结清订单" onClick="if (confirm('确定结清吗？')){window.location.href='KS.ShopOrder.asp?Action=FinishOrder&ID=<%=RS("ID")%>';}" class="button" />
				<%
				end if
				 
				 End if
			 End If
			 
			 IF RS("DeliverStatus")=0 Then%>
			 <input type='button' class='button' name='Submit' value=' 发货 ' onClick="window.location.href='KS.ShopOrder.asp?Action=DeliverGoods&ID=<%=rs("id")%>'">&nbsp;
			 <%ElseIf RS("DeliverStatus")<>3 Then%>
			 <input type='button' class='button' name='Submit' value=' 客户退货 ' onClick="window.location.href='KS.ShopOrder.asp?Action=BackGoods&ID=<%=rs("id")%>'">&nbsp;
			 <%End If%>
			 <%End If%>
			 <%If RS("DeliverStatus")<>3 Then%>
			 <input type='button' class='button' name='Submit' value=' 支付货款给卖方 ' onClick="window.location.href='KS.ShopOrder.asp?Action=PayMoney&ID=<%=rs("id")%>'">&nbsp;
			 <%end if%>
			 <%
			End If
			 %>
			 <input type='button' class='button' name='Submit' value='打印订单' onClick="window.location.href='KS.ShopOrder.asp?Action=PrintOrder&ID=<%=RS("ID")%>'">
			 <input type='button' class='button' name='Submit' value='打印快递单' onClick="window.location.href='KS.ShopOrder.asp?Action=printexpress&ID=<%=RS("ID")%>'">
			 &nbsp;<input type='button' class='button' name='Submit' value='取消返回' onClick="javascript:history.back();">
			</div>
			<br/><br/>
			</div>
</body></html>
		<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		


	
 '删除订单
 Sub DelOrder()
         dim UserName_Order,UseScoreisshop_Order,OrderID
		 If Not KS.ReturnPowerResult(5, "M510021") Then                  '权限检查
			Call KS.ReturnErr(1, "")   
			Response.End()
		  End iF
 
		 Dim ID:ID=KS.G("ID")
		 If ID="" Then KS.echo "<script>history.back();</script>" : Exit Sub
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select OrderID,CouponUserID,UserName,UseScoreisshop From KS_Order Where paystatus<>3 and Status<>2 And MoneyReceipt=0 And ID In(" & ID  &")",Conn,1,1
		 If Not RS.Eof Then
		
		  Do While Not RS.Eof
		   UserName_Order=rs("UserName")
		   UseScoreisshop_Order=rs("UseScoreisshop")
		   OrderID=rs("OrderID")
		   Conn.execute("Update KS_ShopCouponUser Set UseFlag=0,OrderID='' Where ID=" & rs(1))
		   Conn.Execute("Delete From KS_OrderItem Where OrderID='" & RS(0) & "'")
		   if KS.ChkClng(UseScoreisshop_Order)>0 then 
		 	    Session("ScoreHasUse")="-" 
				Call KS.ScoreInOrOut(UserName_Order,1,KS.ChkClng(UseScoreisshop_Order),"系统","购物失败，返还积分，订单号：<font color=red>" & OrderID & "</font>",0,0)	
		  end if
		   RS.MoveNext
		  Loop
		 End If
		 RS.Close:Set RS=Nothing
		 
		 Conn.Execute("Delete From KS_Order Where paystatus<>3 and Status<>2 And MoneyReceipt=0 And ID In(" & ID  &")") 
		  %>
		  <script language=JavaScript>
			$.dialog.alert('恭喜,订单删除成功!',function(){location.replace('<%=Request.ServerVariables("HTTP_REFERER")%>');
			});</script>

		  <%
End Sub
		
		'确认订单
		Sub  OrderConfirm()
		  Dim MailContent:MailContent=KS.Setting(73)
		  Dim ID:ID=KS.G("ID")
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 * From KS_Order Where ID=" & ID ,Conn,1,3
		  If Not RS.Eof Then
		    RS("Status")=1
			RS.Update
			Dim RSA:Set RSA=Server.CreateObject("ADODB.RECORDSET")
			RSA.Open "Select ProID,Amount,AttrID From KS_OrderItem Where OrderID='" & RS("OrderID") & "'",conn,1,1
			do while not rsa.eof
			 If RSA("AttrID")<>0 Then
			  Conn.Execute("update KS_ShopSpecificationPrice set amount=amount-" & RSA(1) & " Where amount>=" & rsa(1) & " and ID=" & RSA(2))
			 Else
			  Conn.Execute("update ks_product set TotalNum=TotalNum-" & RSA(1) & " Where TotalNum>=" & rsa(1) & " and ID=" & RSA(0))
			 End If
			 RSA.MoveNext
			loop
			rsa.close:set rsa=nothing
		    If Trim(RS("UserName"))<>"游客" Then   '游客下的订单不允许发送站内信件
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"订单确认通知",KS.ReplaceOrderLabel(MailContent,RS))
			End If
			If RS("Email")<>"" Then
				Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "收到汇款通知", RS("Email"),RS("ContactMan"), KS.ReplaceOrderLabel(MailContent,rs),KS.Setting(11))
			 End If
			 Dim Mobile:Mobile=RS("Mobile")
			 '发短信
		 Dim Rstr
		    Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(9)
			If Not KS.IsNul(SmsContent) Then
			   If KS.IsNul(Mobile) and Trim(RS("UserName"))<>"游客" Then
			      Dim RSU:Set RSU=Conn.Execute("select top 1 Mobile From KS_User Where UserName='" & RS("UserName") &"'")
				  If Not RSU.Eof Then
				    Mobile=RSU(0)
				  End If
				  RSU.Close
				  Set RSU=Nothing
			   End If 
			   If Not KS.IsNul(Mobile) Then
			      SmsContent=Replace(SmsContent,"{$contactman}",RS("ContactMan"))
			      SmsContent=Replace(SmsContent,"{$orderid}",rs("orderid"))
			      SmsContent=Replace(SmsContent,"{$time}",now)
			      SmsContent=Replace(SmsContent,"{$money}",rs("NoUseCouponMoney"))
				  Rstr=KS.SendMobileMsg(Mobile,SmsContent)
			   End If
			End If
			 
		 %>
		     <div class="pageCont2"><table align='center' width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>订单确认成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他订单已经确认！
			  <%end if%>
			   <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=rs("Email")%>发送了一封邮件通知，通知他订单已确认！
			  <%end if%>
			   <%IF rstr="1" Then%>
			  <br><br>已经向手机<%=rs("mobile")%>发送了一条短信通知，通知他订单已确认！
			  <%end if%>
			  
			  </td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
			</div>
		 <%
		  Else
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		  End If
		  RS.Close:Set RS=Nothing
		End Sub
		
		'银行付款
		Sub BankPay()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		  %>
		  <DIV class="pageCont2">
		<form name='form4' method='post' action='KS.ShopOrder.asp' onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')">  
		<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    <tr style='text-align:center' class='title'>      <td height='25' colspan='2'><b>添 加 订 单 汇 款 信 息</b></td>    </tr>    <tr class='tdbg'>      <td width='15%' style='text-align:right'>客户姓名：</td>      <td><%=rs("contactman")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' style='text-align:right'>用户名：</td>      <td><%=rs("username")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' style='text-align:right'>支付内容：</td>      <td><table  border='0' cellspacing='2' cellpadding='0'>        <tr class='tdbg'>          <td width='15%' style='text-align:right'>订单编号：</td>          <td><%=rs("orderid")%></td>          <td>&nbsp;</td>        </tr>        <tr class='tdbg'>          <td width='15%' style='text-align:right'>订单金额：</td>          <td><%=rs("MoneyTotal")%>元</td>          <td></td>        </tr>        <tr class='tdbg'>          <td width='15%' style='text-align:right'>已 付 款：</td>          <td><%=rs("MoneyReceipt")%>元</td>          <td>&nbsp;</td>        </tr>      </table>      </td>    </tr>    <tr class='tdbg'>      <td width='15%' style='text-align:right'>汇款日期：</td>      <td><input name='PayDate' class="textbox" type='text' id='PayDate' value='<%=formatdatetime(now,2)%>' size='15' maxlength='30'></td>    </tr>    <tr class='tdbg'>      <td width='15%' style='text-align:right'>汇款金额：</td>      <td><input name='Money' class="textbox" type='text' id='Money' value='<%=rs("MoneyTotal")-rs("MoneyReceipt")%>' size='10' maxlength='10'> 元</td>    </tr>       <tr class='tdbg'>      <td width='15%' style='text-align:right'>备注：</td>      <td><input name='Remark' type='text' id='Remark' size='50' maxlength='200' class="textbox" value="支付订单费用，订单号：<%=rs("orderid")%>"></td>    </tr>    
		<tr class='tdbg'>      
		<td width='15%' style='text-align:right'>通知会员：</td>      
		<td>
		 <input type='checkbox' name='SendMessageToUser' value='1' checked>同时使用站内短信通知会员已经收到汇款
		<br><input type='checkbox' name='SendMailToUser' value='1' checked>同时发送邮件通知会员已经收到汇款
		<br><input type='checkbox' name='SendSmsToUser' value='1' checked>同时发送手机短信通知会员
		</td>    
		</tr>    
		<tr class='tdbg'>      <td height='30' colspan='2'><b><font color='#FF0000'>注意：汇款信息一旦录入，就不能再修改或删除！所以在保存之前确认输入无误！</font></b></td>    </tr>   
		 <tr class='tdbg'>      <td height='30' colspan='2' style='text-align:center'>
		 <input name='Action' type='hidden' id='Action' value='DoBankPay'>   
		    <input name='ID' type='hidden' id='ID' value='<%=rs("id")%>'>      
			<input  class='button' type='submit' name='Submit' value='保存汇款信息'>&nbsp;
			<input type='button' class='button' onclick='javascript:history.back();' name='Submit' value='取消返回'></td>    </tr>  </table></form></DIV>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'开始银行支付操作
		Sub DoBankPay()
		 Dim ID:ID=KS.G("ID")
		 Dim PayDate:PayDate=KS.G("PayDate")
		 Dim Money:Money=KS.G("Money")
		 Dim Remark:Remark=KS.G("Remark")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 Dim SendSmsToUser:SendSmsToUser=KS.ChkClng(KS.G("SendSmsToUser"))
		 If Not IsDate(PayDate) Then Response.Write "<script>alert('付款日期格式有误');history.back();</script>":response.end
		 If Not IsNumeric(Money) Then 
		  Response.Write "<script>alert('输入的汇款金额不合法!');history.back();</script>":response.end
		 else
		  If Money<=0 Then
		  Response.Write "<script>alert('汇款金额必须大于0!');history.back();</script>":response.end
		  End If
		 End If
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		    rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		  If Remark="" Then Remark="支付订单费用，订单号：" & rs("orderid")
          Dim HasMoney:HasMoney=RS("MoneyReceipt")
		  If Round(RS("MoneyReceipt")+Money,2)>Round(RS("MoneyTotal"),2) Then
		   RS("MoneyReceipt")=RS("MoneyTotal")
		  Else
           RS("MoneyReceipt")=RS("MoneyReceipt")+Money
		  End If
		  Dim OrderStatus:OrderStatus=RS("Status")
		  RS("Status")=1
		  RS("PayTime")=now   '记录付款时间
		  RS.Update
		  Dim Email:Email=RS("Email")
		  Dim Mobile:Mobile=RS("Mobile")
		  Dim Money1
		  If RS("MoneyReceipt")>=RS("MoneyTotal") And HasMoney=0 Then
		  Money1=RS("MoneyTotal")
		  Else
		  Money1=Money
		  End If
		  
		  If RS("MoneyReceipt")>=RS("MoneyTotal") Then
		  	 RS("PayStatus")=1  '已付清
		  ElseIf RS("MoneyReceipt")<>0 Then
		     RS("PayStatus")=2  '已收定金
		  Else
		     RS("PayStatus")=0  '未付款
		  End If
		  RS.Update

		  Dim ContactMan:ContactMan=RS("ContactMan")
		  Call KS.MoneyInOrOut(rs("UserName"),ContactMan,Money,2,1,now,rs("orderid"),KS.C("AdminName"),"银行汇款",0,0,0)
		  Call KS.MoneyInOrOut(rs("UserName"),ContactMan,Money1,4,2,now,rs("orderid"),KS.C("AdminName"),Remark,0,0,0)
		 If SendMessageToUser=1 and Trim(RS("UserName"))<>"游客" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"收到汇款通知",KS.ReplaceOrderLabel(KS.Setting(74),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "收到汇款通知", Email,ContactMan, KS.ReplaceOrderLabel(KS.Setting(74),rs),KS.Setting(11))
		 End If
		 '发短信
		 Dim Rstr
		 If SendSmsToUser=1 Then
		    Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(5)
			If Not KS.IsNul(SmsContent) Then
			   If KS.IsNul(Mobile) and Trim(RS("UserName"))<>"游客" Then
			      Dim RSU:Set RSU=Conn.Execute("select top 1 Mobile From KS_User Where UserName='" & RS("UserName") &"'")
				  If Not RSU.Eof Then
				    Mobile=RSU(0)
				  End If
				  RSU.Close
				  Set RSU=Nothing
			   End If 
			   If Not KS.IsNul(Mobile) Then
			      SmsContent=Replace(SmsContent,"{$contactman}",ContactMan)
			      SmsContent=Replace(SmsContent,"{$orderid}",rs("orderid"))
			      SmsContent=Replace(SmsContent,"{$time}",now)
			      SmsContent=Replace(SmsContent,"{$money}",Money)
				  Rstr=KS.SendMobileMsg(Mobile,SmsContent)
			   End If
			End If
		 End If
		 %>
		 <DIV class="pageCont2"><table align='center' width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存汇款信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信通知，通知他已经收到汇款！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经收到汇款！
			  <%end if%>
			  <%If Rstr="1" Then%>
			  <br><br>已经向手机号<%=Mobile%>发送了一条短信通知，通知他已经收到汇款！
			  <%End If%>
			  </td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table></DIV>
		 <%
		 
					'====================更新库存量========================
					Dim rsp:set rsp=conn.execute("select id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
					do while not rsp.eof
					  dim rsi:set rsi=conn.execute("select amount,attrid from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(0))
					  if not rsi.eof then
						  if OrderStatus<>1 Then  '扣库存量
						   If RSI("AttrID")<>0 Then
			                  Conn.Execute("update KS_ShopSpecificationPrice set amount=amount-" & RSI(0) & " Where amount>=" & RSI(0) & " and ID=" & RSI(1))
			              Else
						   conn.execute("update ks_product set totalnum=totalnum-" & rsi(0) &" where totalnum>=" & rsi(0) &" and id=" & rsp(0))        
						  End If
						  End If
					  end if
					  rsi.close
					  set rsi=nothing
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
		  RS.Close:Set RS=Nothing
		End Sub
		
		
		'退款
		Sub BankRefund()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * From KS_Order Where OrderID='" & ID & "'",Conn,1,1
		 If RS.Eof Then
		    RS.Close
		    RS.Open "select top 1 * From KS_Order Where ID=" & ks.chkclng(ID),conn,1,1
			If RS.Eof Then
			 RS.Close
			 Set RS=Nothing
		     Response.Write "<script>alert('参数错误！');history.back();</script>"
			End If
		 End IF
		 id=rs("id")
		  %>
<DIV class="pageCont2"><form name='form4' method='post' action='KS.ShopOrder.asp' onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')">  
<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    <tr style='text-align:center' class='title'>      <td height='25' colspan='2'><b>处 理 订 单 退 款</b></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>客户姓名：</td>      <td><%=rs("contactman")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>用户名：</td>      <td><%=rs("username")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>支付内容：</td>      <td><table  border='0' cellspacing='2' cellpadding='0'>        <tr class='tdbg'>          <td width='15%' class='tdbg' style='text-align:right'>订单编号：</td>          <td><%=rs("orderid")%></td>          <td>&nbsp;</td>        </tr>        <tr class='tdbg'>          <td width='15%' class='tdbg' style='text-align:right'>订单金额：</td>          <td><%=rs("moneytotal")%>元</td>        </tr>        <tr class='tdbg'>          <td width='15%' class='tdbg' style='text-align:right'>已 付 款：</td>          <td><%=rs("MoneyReceipt")%>元</td>          <td>&nbsp;</td>        </tr>      </table>      </td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>退款日期：</td>      <td><input name='PayDate' class="textbox" type='text' id='PayDate' value='<%=FormatDateTime(Now,2)%>' size='15' maxlength='30'></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>退款金额：</td>      <td><input name='Money' class="textbox" type='text' id='Money'  size='10' value='<%=rs("MoneyReceipt")%>' maxlength='10'> 元&nbsp;&nbsp;<font color='#0000FF'>退款金额将从已付款中扣除。</font></td>    </tr>    
<tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>退款方式：</td>      <td><input type='radio' name='RefundType' value='1' onClick="Remark.value='订单退款金额，订单号：<%=RS("orderid")%>'" <%if rs("username")<>"游客" then Response.Write " checked"%>>扣除的金额添加到会员资金余额中<br><input type='radio' name='RefundType' value='2' onClick="Remark.value='退款'+$('#Money').val()+'元，退款方式采用银行转账，订单号：<%=rs("orderid")%>'"<%if rs("username")="游客" then Response.Write " checked"%>>采用其它方式：如银行转帐，现金交付等等</td>    </tr>    
<tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>备注：</td>      <td><input name='Remark' class="textbox" type='text' id='Remark' value=<%if rs("username")<>"游客" then Response.Write "'订单退款金额，订单号："&rs("orderid") &"'"  Else Response.Write "'订单退款金额，退款方式采用其它方式，订单号：" & rs("orderid") & "'"%> size='50' maxlength='200'></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' style='text-align:right'>通知会员：</td>      <td><input type='checkbox' name='SendMessageToUser' value='1' checked>同时使用站内短信通知会员已经退款
<br><input type='checkbox' name='SendMailToUser' value='1' checked>同时发送Email通知会员已经退款
<br><input type='checkbox' name='SendMailToSms' value='1' checked>同时发送手机短信通知会员已经退款

</td>    </tr>   
 <tr class='tdbg'>     
  <td height='30' colspan='2'><b><font color='#FF0000'>注意：退款信息一旦录入，就不能再修改或删除,并且订单将自动结清作废！所以在保存之前确认输入无误！</b></td>    
  </tr>    
  <tr style='text-align:center' class='tdbg'>      
  <td height='30'></td><td><input name='Action' type='hidden' id='Action' value='DoRefundMoney'>      <input name='ID' type='hidden' id='ID' value='<%=rs("id")%>'>      <input class='button' type='submit' name='Submit' value=' 确认提交退款 '></td>    </tr>  </table></form></DIV>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'开始退款相关操作
		Sub DoRefundMoney()
		 Dim ID:ID=KS.G("ID")
		 Dim PayDate:PayDate=KS.G("PayDate")
		 Dim Money:Money=KS.G("Money")
		 Dim Remark:Remark=KS.G("Remark")
		 Dim RefundType:RefundType=KS.ChkClng(KS.G("RefundType"))
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 Dim SendSmsToUser:SendSmsToUser=KS.ChkClng(KS.G("SendSmsToUser"))
		 If Not IsDate(PayDate) Then Response.Write "<script>alert('退款日期格式有误');history.back();</script>":response.end
		 If KS.ChkClng(Money)=0 Then Response.Write "<script>alert('退款金额必须大于0!');history.back();</script>":response.end
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		    rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		   
		  If round(Money,2)>round(RS("MoneyReceipt"),2) Then Response.Write "<script>alert('退款金额必须小于已付款金额!');history.back();</script>":response.end
		  If Remark="" Then Remark="订单已退款" &Money & "元，订单号：" & rs("orderid")
          RS("MoneyReceipt")=RS("MoneyReceipt")-Money
		  RS("PayStatus")=3
		  RS("DeliverStatus")=3 '退货
		  RS("Status")=2   '标志结清
		  RS("Remark")=RS("Remark") & "<div>" & Remark & "</div>"
		  RS.Update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=RS("ContactMan")
		  Dim Mobile:Mobile=RS("Mobile")
		  
		  If RefundType=1 Then  '退回账户余额中
		  Call KS.MoneyInOrOut(rs("UserName"),ContactMan,Money,4,1,now,rs("orderid"),KS.C("AdminName"),Remark,0,0,0)
		  End If
		  
		   '返还积分
		  if rs("usescore")>0 and rs("UseScoreMoney")>0 then
			Session("ScoreHasUse")="-" '设置只累计消费积分
			Call KS.ScoreInOrOut(rs("UserName"),1,rs("usescore"),"系统","购物失败，返还积分。订单号<font color=red>" & rs("orderid") & "</font>!",0,0)				
		  end if
		  
		  if KS.ChkClng(rs("UseScoreisshop"))>0 then
			  Session("ScoreHasUse")="-" '设置只累计消费积分
			  Call KS.ScoreInOrOut(rs("UserName"),1,KS.ChkClng(rs("UseScoreisshop")),"系统","购物失败，返还积分。订单号<font color=red>" & rs("orderid") & "</font>!",0,0)	
		  end if
		  '返还抵用券
		  if rs("CouponUserID")>0 and rs("UseCouponMoney")>0 then
		       Dim RSC:Set RSC=Server.CreateObject("adodb.recordset")
			   RSC.Open "Select top 1 * From KS_ShopCouponUser WHERE ID=" & KS.ChkClng(rs("CouponUserID")),conn,1,3
			   If Not RSC.Eof Then

				   If KS.IsNul(RSC("Note")) Then
				     RSC("Note")="购买订单[" & rs("OrderID") & "]失败，返回抵用券金额" & formatnumber(Round(rs("UseCouponMoney"),2),2,-1) & "元;"
				   Else
				     RSC("Note")=rsC("Note") & "<br/>购买订单[" & rs("OrderID") & "]失败，返回抵用券金额" & formatnumber(Round(rs("UseCouponMoney"),2),2,-1) & "元;"
				   End If
				   RSC("AvailableMoney")=RSC("AvailableMoney")-rs("UseCouponMoney")
				   RSC.Update 
			   End If
			   RSC.Close
			   Set RSC=nothing
		  end if
		  
		  
           conn.execute("update KS_LogDeliver set DeliverType=2,status=1 Where OrderID='" & rs("Orderid") & "' and DeliverType=3")
		  

		  
		 If SendMessageToUser=1 and Trim(RS("UserName"))<>"游客" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"退款通知",KS.ReplaceOrderLabel(KS.Setting(75),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "退款退货处理成功通知", Email,ContactMan, KS.ReplaceOrderLabel(KS.Setting(75),rs),KS.Setting(11))
		 End If
		 
		 '发短信
		 Dim Rstr
		 If SendSmsToUser=1 Then
		    Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(7)
			If Not KS.IsNul(SmsContent) Then
			   If KS.IsNul(Mobile) and Trim(RS("UserName"))<>"游客" Then
			      Dim RSU:Set RSU=Conn.Execute("select top 1 Mobile From KS_User Where UserName='" & RS("UserName") &"'")
				  If Not RSU.Eof Then
				    Mobile=RSU(0)
				  End If
				  RSU.Close
				  Set RSU=Nothing
			   End If 
			   If Not KS.IsNul(Mobile) Then
			      SmsContent=Replace(SmsContent,"{$contactman}",ContactMan)
			      SmsContent=Replace(SmsContent,"{$orderid}",rs("orderid"))
			      SmsContent=Replace(SmsContent,"{$time}",now)
			      SmsContent=Replace(SmsContent,"{$money}",Money)
				  Rstr=KS.SendMobileMsg(Mobile,SmsContent)
			   End If
			End If
		 End If
		 
		 %>
		 <div class="pageCont2"><table align='center' width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存退款信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他已经退款！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经退款！
			  <%end if%>
			  <%IF Rstr="1" Then%>
			  <br><br>已经向手机号<%=Mobile%>发送了一条短信通知，通知他已经退款！
			  <%end if%>
			  </td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table></div>
		 <%
					
		  RS.Close:Set RS=Nothing
		End Sub
		
		'发货操作
		Sub DeliverGoods()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
        %><div class="pageCont2">
<FORM name=form4 onSubmit="return confirm('确定录入的发货信息都正确无误了吗？');" action="KS.ShopOrder.asp" method=post>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=25><B>录 入 发 货 信 息</B></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">客户名称：</td>
      <td><%=rs("contactman")%></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">用户名：</td>
      <td><%=rs("username")%></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">收货人姓名：</td>
      <td><%=rs("contactman")%></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">订单编号：</td>
      <td><%=rs("orderid")%></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">订单金额：</td>
      <td><%=formatnumber(rs("MoneyTotal"),2,-1,-1)%>元</td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">已 付 款：</td>
      <td><%=formatnumber(rs("MoneyReceipt"),2,-1,-1)%>元</td>
    </tr>
	<%if not ks.isnul(rs("alipaytradeno")) then%>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">支付宝交易号：</td>
      <td><%=rs("alipaytradeno")%> <span style='font-weight:Bold;color:green'>本单采用支付宝担保交易,发货操作同时改变支付宝订单状态。</span>
	  </td>
    </tr>
	<%end if%>
	
	<%if rs("tocity")<>"" then%>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">客户指定：</td>
      <td>快递公司:<%
	  dim rst,foundexpress,companyname
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
		companyname=rst("typename")
	response.write "<span style='color:green'>" & companyname & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
	  If DataBaseType=1 Then
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
	 Else
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
	 End If
	  if rst.eof then
	    rst.close : set rst=nothing
	  else
	response.write "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	   rst.close
	  end if
	 End If
	 set rst=nothing
	
	
	response.write " 发往<span style='color:red'>" & rs("tocity") & "</span>"
	  
	  %></td>
    </tr>
<%end if%>
	
	
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">发货日期：</td>
      <td>
        <Input id="DeliverDate" class="textbox" maxLength=30 size=15 value="<%=formatdatetime(now,2)%>" name="DeliverDate"></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">快递公司：</td>
      <td>
   <Input id="ExpressCompany2"  maxLength=30 size=15 name="ExpressCompany2"  type="hidden" value=""  >  
  <Input id="ExpressCompany" maxLength=30 size=15 class="textbox" name="ExpressCompany" value="<%=companyname%>"> <=
  <select id="Code" name="Code" onChange="$('#ExpressCompany').val($('#Code option:selected').text());$('#ExpressCompany2').val(this.value)"> 
          <option value=''>---快速选择快递公司---</option>
           <%
		    dim rss:set rss=conn.execute("select * from KS_Deliverytype")
			do while not rss.eof
			  response.write "<option value='" & rss("typename_e") &"'>" & rss("typename") & "</option>"
			  rss.movenext
			loop
			rss.close
			set rss=nothing
		   %>

			 </select>
			 
     </td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">快递单号：</td>
      <td>
        <Input id="ExpressNumber" class="textbox" maxLength=30 size=15 name="ExpressNumber"></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">经 手 人：</td>
      <td>
        <Input id="HandlerName" class="textbox" maxLength=50 size=30 value="<%=KS.C("AdminName")%>" name="HandlerName"></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td>
      <td>
        <Input id=Remark class="textbox" maxLength=200 size=50 name="Remark" value="订单号：<%=rs("orderid")%>的货物已送出"></td>
    </tr>
    <tr class=tdbg>
      <td  style='text-align:right' width="15%">通知会员：</td>
      <td>
  <Input type=checkbox CHECKED value="1" name="SendMessageToUser">同时使用站内短信通知会员已经发货<br>
  <input type="checkbox" checked value="1" name="SendMailToUser">同时发送Email通知会员已经发货<br/>
  <input type="checkbox" checked value="1" name="SendSmsToUser">同时发送手机短信通知会员已经发货</td>
    </tr>
    <tr class=tdbg>
	 <td></td>
      <td height=30>
	  <Input id=Action type=hidden value="DoDeliverGoods" name="Action"> 
	  <Input id=OrderFormID type=hidden value="<%=rs("id")%>" name="ID"> 
      <Input class='button' type=submit value=" 保 存 发 货" name=Submit></td>
    </tr>
  </table>
</FORM>
</div>
		<% rs.close:set rs=nothing
		End Sub
		
		'发货操作
		Sub DoDeliverGoods()
		 Dim ID:ID=KS.G("ID")
		 Dim DeliverDate:DeliverDate=KS.G("DeliverDate")
		 Dim ExpressCompany:ExpressCompany=KS.G("ExpressCompany")
		 Dim ExpressCompany2:ExpressCompany2=KS.G("ExpressCompany2")
		 Dim ExpressNumber:ExpressNumber=KS.G("ExpressNumber")
		 Dim HandlerName:HandlerName=KS.G("HandlerName")
		 Dim Remark:Remark=KS.G("Remark")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 Dim SendSmsToUser:SendSmsToUser=KS.ChkClng(KS.G("SendSmsToUser"))
		 
		 If Not IsDate(DeliverDate) Then Response.Write "<script>alert('发货日期格式有误');history.back();</script>":response.end
		 If (HandlerName="") Then Response.Write "<script>alert('经手人必须填写');history.back();</script>":response.end
		 If (ExpressCompany="") Then Response.Write "<script>alert('快递公司必须填写');history.back();</script>":response.end
		 If (ExpressNumber="") Then Response.Write "<script>alert('快递单号必须填写');history.back();</script>":response.end
		 
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		 If rs("DeliverStatus")=1 Then  Response.Write "<script>alert('此订单已经发过货!');history.back();</script>":Response.end 
		 
		 
		 '================================================同步支付宝发货接口====================================
		 Dim alipaytradeno:alipaytradeno=RS("alipaytradeno")
		 Dim PaymentPlatId:PaymentPlatId=KS.ChkClng(rs("PaymentPlatId"))
		 Dim SendToAlipayTF:SendToAlipayTF=false
		 Dim alipayErr
		 if Not KS.IsNul(alipaytradeno) and (PaymentPlatId=9 or PaymentPlatId=15) Then
		    Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			RSP.Open "Select top 1 * From KS_PaymentPlat where id=" & PaymentPlatId,conn,1,1
			If Not RSP.Eof Then
			Dim AccountID:AccountID=RSP("AccountID")
			MD5Key=RSP("MD5Key")
			End If
			RSP.Close:Set RSP=Nothing
		    Dim Partner
			Dim ArrMD5Key
			If InStr(MD5Key, "|") > 0 Then
				ArrMD5Key = Split(MD5Key, "|")
				If UBound(ArrMD5Key) = 1 Then
					MD5Key = ArrMD5Key(0)
					Partner = ArrMD5Key(1)
				End If
			End If
			
			if Partner<>"" Then
				 dim surl:surl="https://mapi.alipay.com/gateway.do"  '发货接口地址
				input_charset="utf-8"
				dim invoice_no:invoice_no=ExpressNumber           '发货单号
				dim logistics_name:logistics_name=ExpressCompany  '物流公司,中文会出错
				dim trade_no:trade_no=alipaytradeno               '支付宝交易号
				dim transport_type:transport_type="EXPRESS"
				
				dim mystr:mystr = Array("service=send_goods_confirm_by_platform","partner="&partner,"trade_no="&trade_no,"logistics_name="&logistics_name,"invoice_no="&invoice_no,"transport_type="&transport_type,"_input_charset="&input_charset)
				
				Dim mysign:mysign = GetMysign(mystr)
				surl=surl & "?_input_charset=" & input_charset &"&invoice_no=" & invoice_no & "&logistics_name=" & server.URLEncode(logistics_name) & "&partner=" & Partner &"&service=send_goods_confirm_by_platform&trade_no="& trade_no& "&transport_type=EXPRESS&sign="& mysign &"&sign_type=MD5"
				
				Dim objHttp:Set objHttp=KS.InitialObject("Msxml2.ServerXMLHTTP.3.0")
				objHttp.open "GET", sUrl, False, "", ""
				objHttp.send()
				Dim XMLDoc:Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		        if XMLDoc.loadxml(objHttp.Responsetext) then
				  Dim Node:Set Node=XMLDoc.getElementsByTagName("alipay").item(0)
				  if node.selectsinglenode("is_success").text="T" Then 
				   SendToAlipayTF=true
				  else
				   alipayErr=node.selectsinglenode("error").text
				  end if
				end if
		  end if
		   		  rs("alipaytradestatus")="WAIT_BUYER_CONFIRM_GOODS"   '更新状态
       end if
	   '============================================================================================
		 
		 
		  rs("DeliverStatus")=1
		  rs("DeliveryDate")=now
		  rs.update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=rs("Contactman")
		  Dim Mobile:Mobile=RS("Mobile")
		  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		  RSLog.Open "Select top 1 * From KS_LogDeliver",Conn,1,3
		   RSLog.AddNew
		    RSLog("OrderID")=RS("OrderID")
			RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("ContactMan")
			RSLog("Inputer")=KS.C("AdminName")
			RSLog("HandlerName")=HandlerName  
			RSLog("DeliverDate")=DeliverDate
			RSLog("DeliverType")=1  '发货
			RSLog("Remark")=Remark
			RSLog("ExpressCompany")=ExpressCompany
			RSLog("ExpressCompany2")=ExpressCompany2
			RSLog("ExpressNumber")=ExpressNumber
			RSLog("Status")=0
		 RSLog.Update
		 RSLog.Close:Set RSLog=Nothing
		  If SendMessageToUser=1 and trim(rs("UserName"))<>"游客" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"发货通知",KS.ReplaceOrderLabel(KS.Setting(77),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "发货通知", Email,ContactMan, KS.ReplaceOrderLabel(KS.Setting(77),rs),KS.Setting(11))
		 End If
		 
		 '发短信
		 Dim Rstr
		 If SendSmsToUser=1 Then
		    Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(6)
			If Not KS.IsNul(SmsContent) Then
			   If KS.IsNul(Mobile) and Trim(RS("UserName"))<>"游客" Then
			      Dim RSU:Set RSU=Conn.Execute("select top 1 Mobile From KS_User Where UserName='" & RS("UserName") &"'")
				  If Not RSU.Eof Then
				    Mobile=RSU(0)
				  End If
				  RSU.Close
				  Set RSU=Nothing
			   End If 
			   If Not KS.IsNul(Mobile) Then
			      SmsContent=Replace(SmsContent,"{$contactman}",ContactMan)
			      SmsContent=Replace(SmsContent,"{$orderid}",rs("orderid"))
			      SmsContent=Replace(SmsContent,"{$time}",now)
			      SmsContent=Replace(SmsContent,"{$express}",ExpressCompany)
			      SmsContent=Replace(SmsContent,"{$expressno}",ExpressNumber)
				  Rstr=KS.SendMobileMsg(Mobile,SmsContent)
			   End If
			End If
		 End If
		 
%>
		 <div class="pageCont2"><table align='center' width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存发货信息成功！
			  <%
			  if alipaytradeno<>"" then
			   if SendToAlipayTF=true then%>
			   <br/><br/><font color=green>已经同步更新了支付宝的交易单的发货状态！</font>
			  <%else%>
			   <br/><br/><font color=red>无法同步更新支付宝交易单的发货状态，支付宝出错返回码：<%=alipayErr%></font>
			  <%end if
			  end if%>
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他已经发货！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经发货！
			  <%end if%>
			  <%IF Rstr="1" Then%>
			  <br><br>已经向手机号<%=Mobile%>发送了一条短信通知，通知他已经发货！
			  <%end if%>
			  </td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table></div>
			<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		'退货操作
		Sub BackGoods()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		%>
		<div class="pageCont2">
		<FORM name=form4 onSubmit="return confirm('确定录入的退货信息都正确无误了吗？');" action=KS.ShopOrder.asp method=post>
		<table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
		  <tr class=title align=middle>
			<td colSpan=2 height=22><B>录 入 退 货 信 息</B></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">客户名称：</td>
			<td><%=rs("contactman")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">用户名：</td>
			<td><%=rs("username")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">收货人姓名：</td>
			<td><%=rs("contactman")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">订单编号：</td>
			<td><%=rs("orderid")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">订单金额：</td>
			<td><%=formatnumber(rs("MoneyTotal"),2,-1,-1)%>元</td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">已 付 款：</td>
			<td><%=formatnumber(rs("MoneyReceipt"),2,-1,-1)%>元</td>
		  </tr>
		  
		<%If rs("tocity")<>"" then%>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">送货方式：</td>
			<td>快递公司:<%
	  dim rst,foundexpress,companyname
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
		companyname=rst("typename")
	response.write "<span style='color:green'>" & companyname & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
		 If DataBaseType=1 Then
			  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
		Else
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
		 End If
	  if rst.eof then
	    rst.close : set rst=nothing
	  else
	response.write "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	rst.close
	  end if
	 End If
	 set rst=nothing
	
	
	response.write " 发往<span style='color:red'>" & rs("tocity") & "</span>"
	  
	  %>&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>客户指定的送货方式</font></td>
		  </tr>
		<%end if%>  
		  
		  
		  
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">退货日期：</td>
			<td>
			  <Input id=DeliverDate maxLength=30 size=15 value="<%=now%>" name=DeliverDate></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">经 手 人：</td>
			<td>
			  <Input id=HandlerName maxLength=50 size=30 value="<%=KS.C("AdminName")%>" name=HandlerName></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 style='text-align:right' width="15%">退货原因：</td>
			<td>
			  <Input id=Remark maxLength=200 size=50 name=Remark></td>
		  </tr>
		  <tr class=tdbg align=middle>
			<td colSpan=2 height=30>
		  <Input id=Action type=hidden value="SaveBack" name=Action> 
		  <Input id=ID type=hidden value=<%=rs("id")%> name=ID> 
			  <Input type=submit value=" 保 存 " class="button" name=Submit></td>
		  </tr>
		</table>
		</FORM>
		</div>
		<%
		rs.close:set rs=nothing
		End Sub
		
		'退货操作
		Sub SaveBack()
		 Dim ID:ID=KS.G("ID")
		 Dim DeliverDate:DeliverDate=KS.G("DeliverDate")
		 Dim HandlerName:HandlerName=KS.G("HandlerName")
		 Dim Remark:Remark=KS.G("Remark")
		 
		 If Not IsDate(DeliverDate) Then Response.Write "<script>alert('退货日期格式有误');history.back();</script>":response.end
		 If (HandlerName="") Then Response.Write "<script>alert('经手人必须填写');history.back();</script>":response.end
		 If Remark="" Then Response.Write "<script>alert('请输入退货原因!');history.back();</script>":response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		  dim DeliverStatus:DeliverStatus=rs("DeliverStatus")
		  rs("DeliverStatus")=3
		  rs.update
		  
		  
		  if DeliverStatus<>3 then
		   '====================为用户减少购物应得积分========================
					Dim rsp:set rsp=conn.execute("select id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & ID & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select top 1 amount from ks_orderitem where orderid='" &ID & "' and proid=" & rsp(0))(0)
					  conn.execute("update ks_product set totalnum=totalnum+" & amount &" where id=" & rsp(0))         '扣库存量
					 ' response.write rs("orderid") & "=55<br>"
					 ' response.write amount & "<br>"
					 ' response.write username & "<br>"
					  
					 '  Call KS.ScoreInOrOut(UserName,2,KS.ChkClng(rsp(0))*amount,"系统","商品退货<font color=red>" & rsp("title") & "</font>扣除!",0,0)

					  
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
		  end if
		  
		  
		  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		  RSLog.Open "Select top 1 * From KS_LogDeliver where DeliverType=2 and orderid='" & RS("OrderID") & "'",Conn,1,3
		  If RSLog.Eof Then
		   RSLog.AddNew
		  End If
		    RSLog("OrderID")=RS("OrderID")
			RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("ContactMan")
			RSLog("Inputer")=KS.C("AdminName")
			RSLog("HandlerName")=HandlerName  
			RSLog("DeliverDate")=DeliverDate
			RSLog("DeliverType")=2  '退货
			RSLog("Remark")=Remark
			RSLog("ExpressCompany")=""
			RSLog("ExpressNumber")=""
			RSLog("Status")=1
		 RSLog.Update
		 RSLog.Close:Set RSLog=Nothing
%>
		 <div class="pageCont2"><table align='center' width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存退货信息成功！
			 <br><br></td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table></div>
			<%
		 RS.Close:Set RS=Nothing		
		 End Sub
		
		'开发票
		Sub Invoice()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		%>
		<div class="pageCont2"><FORM name=form4 onSubmit="return confirm('确定录入的发票信息都正确无误了吗？');" action="KS.ShopOrder.asp" method=post>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=22><B>录 入 开 发 票 信 息</B></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">客户名称：</td>
      <td><%=RS("ContactMan")%></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">用户名：</td>
      <td><%=RS("UserName")%></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">订单编号：</td>
      <td><%=RS("OrderID")%></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">订单金额：</td>
      <td><%=RS("MoneyTotal")%>元</td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">已 付 款：</td>
      <td><%=RS("MoneyReceipt")%>元</td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">发票信息：</td>
      <td><%=RS("InvoiceContent")%></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">开票日期：</td>
      <td>
        <Input id="InvoiceDate" maxLength=30 size=15 class="textbox" value="<%=FormatDateTime(Now,2)%>" name="InvoiceDate"></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">发票类型：</td>
      <td>
<Select name="InvoiceType">
  <Option value="地税普通发票" selected>地税普通发票</Option>
  <Option value="国税普通发票">国税普通发票</Option>
  <Option value="增值税发票">增值税发票</Option>
      </Select></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">发票号码：</td>
      <td>
        <Input id=InvoiceNum maxLength=30 size=15 class="textbox" name="InvoiceNum"></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">发票抬头：</td>
      <td>
        <Input id=InvoiceTitle maxLength=50 size=50 class="textbox" value="" name="InvoiceTitle"></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">发票内容：</td>
      <td><TEXTAREA name=InvoiceContent rows=4 class="textbox" cols=50><%=RS("InvoiceContent")%></TEXTAREA></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">发票金额：</td>
      <td>
        <Input id="MoneyTotal" maxLength=15 size=15 class="textbox" value="<%=RS("MoneyTotal")%>" name="MoneyTotal"></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">开 票 人：</td>
      <td>
        <Input id="HandlerName" maxLength=30 size=15 class="textbox" value="<%=KS.C("AdminName")%>" name="HandlerName"></td>
    </tr>
    <tr class=tdbg>
      <td style='text-align:right' width="15%">通知会员：</td>
      <td>
  <Input type=checkbox CHECKED value="1" name="SendMessageToUser">同时使用站内短信通知会员已经开具发票<br>
  <Input type=checkbox CHECKED value="1" name="SendMailToUser">同时发送Email通知会员已经开具发票<br>
  <Input type=checkbox CHECKED value="1" name="SendSmsToUser">同时发送手机短信通知会员已经开具发票<br>
  </td>
    </tr>
    <tr class=tdbg align=middle>
      <td height=30></td>
	  <td>
  <Input id=Action type=hidden value="DoSaveInvoice" name="Action"> 
  <Input id="ID" type=hidden value="<%=RS("ID")%>" name="ID"> 
        <Input type=submit class='button' value=" 保存开票记录 " name=Submit></td>
    </tr>
  </table>
</FORM>
</div>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'保存发票
		Sub DoSaveInvoice()
		 Dim ID:ID=KS.G("ID")
		 Dim InvoiceDate:InvoiceDate=KS.G("InvoiceDate")
		 Dim InvoiceType:InvoiceType=KS.G("InvoiceType")
		 Dim InvoiceNum:InvoiceNum=KS.G("InvoiceNum")
		 Dim InvoiceTitle:InvoiceTitle=KS.G("InvoiceTitle")
		 Dim InvoiceContent:InvoiceContent=KS.G("InvoiceContent")
		 Dim MoneyTotal:MoneyTotal=KS.G("MoneyTotal")
		 Dim HandlerName:HandlerName=KS.G("HandlerName")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 Dim SendSmsToUser:SendSmsToUser=KS.ChkClng(KS.G("SendSmsToUser"))
		 If Not IsDate(InvoiceDate) Then Response.Write "<script>alert('开票日期格式有误');history.back();</script>":response.end
		 If (HandlerName="") Then Response.Write "<script>alert('开票人必须填写');history.back();</script>":response.end
		 If (InvoiceTitle="") Then Response.Write "<script>alert('发票抬头必须填写');history.back();</script>":response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
           rs("Invoiced")=1
		  rs.update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=rs("ContactMan")
		  Dim Mobile:Mobile=rs("Mobile")
		  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		  RSLog.Open "Select top 1 * From KS_LogInvoice",Conn,1,3
		   RSLog.AddNew
			RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("ContactMan")		    
			RSLog("OrderID")=RS("OrderID")
            RSLog("InvoiceType")=InvoiceType
			RSLog("InvoiceNum")=InvoiceNum
			RSLog("InvoiceTitle")=InvoiceTitle
			RSLog("InvoiceContent")=InvoiceContent
			RSLog("InvoiceDate")=InvoiceDate
			RSLog("InputTime")=Now
			RSLog("MoneyTotal")=MoneyTotal
			RSLog("Inputer")=KS.C("AdminName")
			RSLog("HandlerName")=HandlerName  
		 RSLog.Update
		 RSLog.Close:Set RSLog=Nothing
		  If SendMessageToUser=1 and Trim(RS("UserName"))<>"游览" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"开发票通知",KS.ReplaceOrderLabel(KS.Setting(76),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "开发票通知", Email,ContactMan, KS.ReplaceOrderLabel(KS.Setting(76),rs),KS.Setting(11))
		 End If
		 
		 
		 '发短信
		 Dim Rstr
		 If SendSmsToUser=1 Then
		    Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(8)
			If Not KS.IsNul(SmsContent) Then
			   If KS.IsNul(Mobile) and Trim(RS("UserName"))<>"游客" Then
			      Dim RSU:Set RSU=Conn.Execute("select top 1 Mobile From KS_User Where UserName='" & RS("UserName") &"'")
				  If Not RSU.Eof Then
				    Mobile=RSU(0)
				  End If
				  RSU.Close
				  Set RSU=Nothing
			   End If 
			   If Not KS.IsNul(Mobile) Then
			      SmsContent=Replace(SmsContent,"{$contactman}",ContactMan)
			      SmsContent=Replace(SmsContent,"{$orderid}",rs("orderid"))
			      SmsContent=Replace(SmsContent,"{$time}",now)
			      SmsContent=Replace(SmsContent,"{$company}",InvoiceTitle)
			      SmsContent=Replace(SmsContent,"{$content}",InvoiceContent)
			      SmsContent=Replace(SmsContent,"{$money}",MoneyTotal)
				  Rstr=KS.SendMobileMsg(Mobile,SmsContent)
			   End If
			End If
		 End If
%>
		 <div class="pageCont2"><table align='center' width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存开发票信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他已经开发票！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经开发票！
			  <%end if%>
			  <%IF Rstr="1" Then%>
			  <br><br>已经向手机号<%=Mobile%>发送了一条短信通知，通知他已经开发票！
			  <%end if%>
			  
			  </td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table></div>
			<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		'已签收商品
		Sub ClientSignUp()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		   rs("DeliverStatus")=2
		   rs.update
		   Conn.execute("update KS_LogDeliver Set Status=1 Where OrderID='" & RS("OrderID") & "'")
		 RS.Close:Set RS=Nothing
		 Response.Redirect "KS.ShopOrder.asp?Action=ShowOrder&ID=" & ID
		End Sub
		
		'打印清单
		Sub PrintOrder() 
		 Dim ID:ID=KS.G("ID")
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Order Where ID=" & ID,Conn,1,1
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		 response.write "<div class=""pageCont2 mt20"">" & KS.OrderDetailStr(RS,0)
		 response.write "</div>"
		 RS.Close:Set RS=Nothing
		 %> 
		 <div id='Varea' style='text-align:center'>
		 	 <input type='button' class='button' name='Submit' value='开始打印' onClick="document.all.Varea.style.display='none';window.print();">&nbsp;<input type='button' class='button' name='Submit' value='取消打印' onClick="javascript:history.back();">
             </div>
		 <%
		End Sub
		
	 '支付货款给卖方
		Sub PayMoney()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		  %>
		<div class="pageCont2"><form name='form4' method='post' action='KS.ShopOrder.asp' onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')">  
		<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    <tr style='text-align:center' class='title'>      <td height='25' colspan='2'><b>支 付 货 款 给 卖 方</b></td>    </tr>    
		 <tr class='tdbg'>      <td width='15%' style='text-align:right'>订单信息：</td>      <td><table  border='0' cellspacing='2' cellpadding='0'>        <tr class='tdbg'>          <td width='15%' style='text-align:right'>订单编号：</td>          <td><%=rs("orderid")%></td>          <td>&nbsp;</td>        </tr>        <tr class='tdbg'>          <td width='15%' style='text-align:right'>订单金额：</td>          <td><%=rs("MoneyTotal")%>元</td>          <td></td>        </tr>        <tr class='tdbg'>          <td width='15%' style='text-align:right'>已 付 款：</td>          <td><%=rs("MoneyReceipt")%>元</td>          <td>&nbsp;</td>        </tr>      </table>      </td>    
		 </tr> 
		 <tr class='tdbg'>      
	  	  <td width='15%' style='text-align:right'>支付明细：</td>      
		  <td>
		  <%dim rso:set rso=server.createobject("adodb.recordset")
		  rso.open "select sum(a.TotalPrice),Inputer from ks_orderitem a inner join ks_product b on a.proid=b.id where a.orderid='" & rs("orderid") & "' group by inputer",conn,1,1
		  do while not rso.eof
		    response.write "卖方<font color=red>" & rso("inputer") & "</font>总价款:" & rso(0) & "元 本次应支付<font color=green>" & rso(0)-(rso(0) * ks.setting(79))/100 & "</font>元<br>"
		  rso.movenext
		  loop
		  rso.close
		  set rso=nothing
		  %>
		  </td>    
		 </tr>  
		
		  <tr class='tdbg'>      <td width='15%' style='text-align:right'>支付时间：</td>      <td><input name='PayDate' class="textbox" type='text' id='PayDate' value='<%=now%>' size='25' maxlength='30'></td>    </tr>  
		  
		   <tr class='tdbg'>      <td width='15%' style='text-align:right'>备注：</td>      <td><input name='Remark' class="textbox" type='text' id='Remark' size='50' maxlength='200' value="收到货款费用，订单号：<%=rs("orderid")%>"></td>    </tr>    <tr class='tdbg'>      <td width='15%' style='text-align:right'>通知会员：</td>      <td><input type='checkbox' name='SendMessageToUser' value='1' checked>同时使用站内短信通知卖主已经支付
		   <br><input type='checkbox' name='SendMailToUser' value='1' checked>同时发送邮件通知卖主已经支付
		   <br><input type='checkbox' name='SendSmsToUser' value='1' checked>同时发送手机短信通知卖主已经支付
		   </td>    
		   </tr>    
		   <tr class='tdbg'>     
		   <td height='30' colspan='2'><b><font color='#FF0000'>注意：一旦按确定支付，就不能再修改或删除！所以在保存之前确认输入无误！</font></b>
		   </td>    
		   </tr>   
		    <tr style='text-align:center' class='tdbg'>     
			 <td></td> <td height='30'><input name='Action' type='hidden' id='Action' value='DoPayMoney'>      <input name='OrderID' type='hidden' id='orderID' value='<%=rs("orderid")%>'>
		   <input name='ID' type='hidden' id='ID' value='<%=rs("id")%>'>
		   <input  class='button' type='submit' name='Submit' value='确定支付'>&nbsp;<input type='button' class='button' onclick='javascript:history.back();' name='Submit' value='取消返回'></td>    </tr>  
		   </table></form></div>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'开始支付货款给卖家操作
		Sub DoPayMoney()
		 Dim OrderID:OrderID=KS.G("OrderID")
		 Dim PayDate:PayDate=KS.G("PayDate")
		 Dim Remark:Remark=KS.G("Remark")
		 If Remark="" Then Remark="收到货款费用，订单号：" & rs("orderid")

		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 Dim SendSmsToUser:SendSmsToUser=KS.ChkClng(KS.G("SendSmsToUser"))
		 If Not IsDate(PayDate) Then Response.Write "<script>alert('支付日期格式有误');history.back();</script>":response.end
		 If not Conn.Execute("Select top 1 PayToUser From ks_Order Where Paytouser=1 and OrderID='" & OrderID & "'").eof Then
		   response.write "<script>alert('对不起，该订单已支付过。不能重复支付!');history.back();</script>"
		   response.end
		 End If
		 
		 
		 dim rso,rsu
		 set rso=server.createobject("adodb.recordset")
		  rso.open "select sum(a.TotalPrice),Inputer from ks_orderitem a inner join ks_product b on a.proid=b.id where a.orderid='" & OrderID & "' group by inputer",conn,1,1
		  do while not rso.eof
		     set rsu=server.createobject("adodb.recordset")
			 rsu.open "select top 1 * from ks_user where username='" & rso(1) & "'",conn,1,1
			 if not rsu.eof then
			    Dim TotalMoney:TotalMoney=rso(0)
				Dim ServiceMoney:ServiceMoney=(TotalMoney * ks.setting(79))/100
				Dim MustPayMoney:MustPayMoney=(TotalMoney-ServiceMoney)
				
				Call KS.MoneyInOrOut(rsu("UserName"),rsu("RealName"),TotalMoney,4,1,PayDate,OrderID,KS.C("AdminName"),Remark,0,0,0)
				Call KS.MoneyInOrOut(rsu("UserName"),rsu("RealName"),ServiceMoney,4,2,PayDate,OrderID,KS.C("AdminName"),"支付订单:"& OrderID & "的服务费",0,0,0)

				 
				 Dim Email:Email=RSU("Email")
				 Dim ContactMan:ContactMan=RSU("RealName")
				 Dim Mobile:Mobile=RSU("Mobile")
				 Dim SiteMessage,Mail,MailContent
				 If ContactMan="" or isnull(ContactMan) Then ContactMan=RSU("UserName")
				 
				 MailContent=KS.Setting(80)
				 MailContent=Replace(MailContent,"{$ContactMan}",ContactMan)
				 MailContent=Replace(MailContent,"{$OrderID}",orderid)
				 MailContent=Replace(MailContent,"{$TotalMoney}",TotalMoney)
				 MailContent=Replace(MailContent,"{$ServiceCharges}",ServiceMoney)
				 MailContent=Replace(MailContent,"{$RealMoney}",TotalMoney-ServiceMoney)
				 
		
				 If SendMessageToUser=1 Then
					'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
					Call KS.SendInfo(rsu("username"),KS.C("AdminName"),"支付货款通知",MailContent)
					SiteMessage="已经向卖方" & rsu("username") & "发送了一条站内短信通知，通知他已经支付货款<br>"
				 End If
				 If SendMailToUser=1 and Email<>"" Then
					Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0)&"向您支付货款通知", Email,ContactMan, MailContent,KS.Setting(11))
					If ReturnInfo="OK" Then
					 Mail="已经向" & Email  &"发送了一封邮件通知，通知他已经支付货款！<br>"
					End If
				 End If
				 
				 
				  '发短信
				 Dim Rstr
				 If SendSmsToUser=1 and Not KS.IsNul(Mobile)Then
					Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(10)
					If Not KS.IsNul(SmsContent) Then
						  SmsContent=Replace(SmsContent,"{$contactman}",ContactMan)
						  SmsContent=Replace(SmsContent,"{$orderid}",orderid)
						  SmsContent=Replace(SmsContent,"{$time}",now)
						  SmsContent=Replace(SmsContent,"{$totalmoney}",TotalMoney)
						  SmsContent=Replace(SmsContent,"{$servicecharges}",ServiceMoney)
						  SmsContent=Replace(SmsContent,"{$realmoney}",TotalMoney-ServiceMoney)
						  Rstr=KS.SendMobileMsg(Mobile,SmsContent)
					End If
				 End If
		 
			 end if
			 rsu.close
		  rso.movenext
		  loop
		  rso.close
		  set rso=nothing
		  set rsu=nothing

          '标志已支付
		  Conn.Execute("Update KS_Order Set PayToUser=1 where orderid='" & OrderID & "'")
		
		 %>
		 <div class="pageCont2"><table align='center' width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr style='text-align:center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>已将货款支付给卖方！
			  <br><br>
			  <%=SiteMessage%>
			  <br>
			  <%=Mail%>
			  <br/>
			  <%if  Rstr="1" then%>
			   已经向手机<%=mobile%>发送了一条短信通知，通知他已经支付货款！
			  <%end if%>
			  </td></tr>
			<tr class='tdbg'><td height=25 style='text-align:center'><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=KS.G("ID")%>'><<点此返回</a></td></tr>
			</table></div>
		 <%
		End Sub		
		
		
End Class
%> 

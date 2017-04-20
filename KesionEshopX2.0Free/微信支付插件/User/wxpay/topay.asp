<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Config.asp"-->
<!--#include file="lib/md5.asp"-->
<!--#include file="lib/Class.asp"-->
<%  
  
	dim out_trade_no,body,total_fee,openid,prepay_id,paySign,PayFrom,UserCardID,Title,money,OrderID,attach
	
	OrderID=KS.S("O")
	PayFrom=KS.S("f")
	UserCardID=KS.ChkClng(KS.S("CID"))
	money=KS.S("m")
	attach=UserCardID&"___" & KS.S("n")&"___"& payfrom
	
	 If UserCardID<>0 Then
		   Dim RS:Set RS=Conn.Execute("Select Top 1 Money,GroupName From KS_UserCard Where ID=" & UserCardID)
		   If Not RS.Eof Then
		    Title=RS(1)
		    Money=RS(0)
			RS.Close : Set RS=Nothing
		   Else
		    RS.Close : Set RS=Nothing
		    Call KS.AlertHistory("出错啦！",-1)
		   End If
	 ElseIf PayFrom="shop" Then
		   Title="购买商品"
	 Else
		   Title="""" & KS.Setting(0) & """账户在线充值,订单号:" & OrderID
	 End If
		
	
	'下面三个参数，需要在商城转过来，包括唯一的订单号，商品名称，总金额
	'out_trade_no	= request("out_trade_no")
	'body			= request("body")
	'total_fee		= request("total_fee")
	out_trade_no	= getStrNow & getStrRandNumber(9999,1000) '唯一订单号，可以自行生成
	out_trade_no	= OrderID  '唯一订单号，可以自行生成
	body			= Title  	    '商品名称
	total_fee		= Money*100  	'以分为单位	
		
	openid			= GetOpenId
	prepay_id		= get_prepay_id
	paySign			= get_paySign()
%>
<!DOCTYPE html>
<html>
<head>
<title>微信支付</title>
<script Language="javascript">
var prepay_id="<%=prepay_id%>";
var paySign="<%=paySign%>";
function Pay_ok()
{
	alert ("支付成功");
	self.location='<%=redirect_url%>?body=<%=body%>&total_fee=<%=total_fee%>&out_trade_no=<%=out_trade_no%>'; 
}

function callpay()
{	
	WeixinJSBridge.invoke('getBrandWCPayRequest',{"appId":"<%=getAppId%>","timeStamp":"<%=timeStamp%>","nonceStr":"<%=nonce_str%>","package":"prepay_id=<%=prepay_id%>","signType":"MD5","paySign":"<%=paySign%>"},function(res){if(res.err_msg=="get_brand_wcpay_request:ok"){Pay_ok();}else{alert(res.err_code+res.err_desc+res.err_msg);}});
}
</script>
<meta http-equiv="content-type" content="text/html;charset=utf-8"/>
<meta id="viewport" name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1; user-scalable=no;" />
<style>
body {
	margin:0;
	padding:0;
	background:#eae9e6;
}
body, p, table, td, th {
	font-size:14px;
	font-family:helvetica, Arial, Tahoma;
}
h1 {
	font-family:Baskerville, HelveticaNeue-Bold, helvetica, Arial, Tahoma;
}
a {
	text-decoration:none;
	color:#385487;
}
.title h1 {
	font-size:22px;
	font-weight:bold;
	padding:0;
	margin:0;
	line-height:1.2;
	color:#1f1f1f;
}
</style>
</head>
<body>
<div style="margin:0 auto;TEXT-ALIGN: center;">
<p><br></p>
  <p id="test">商品名称：<%=body%></p>
  <p>商品金额：<%=FormatNumber(total_fee*0.01,2,-1)%>元</p>
  <p></p>
  <a href="javascript:callpay();">
  <h1 class="title">点击支付商品</h1>
  </a><br><%=prepay_id%>
</div>
</body>
</html>
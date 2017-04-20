<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html>
<head>
<title>微信支付显示结果</title>
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
	color: #F00;
}
</style>
</head>
<body>
<div>
  <p>商品订单：<%=request("out_trade_no")%></p>
  <p>商品名称：<%=request("body")%></p>
  <p>商品金额：<%=FormatNumber(request("total_fee")*0.01,2,-1)%>元</p>
  <p></p>
  <h1 class="title">支付成功！！！！</h1>
  </div>
</body>
</html>
<!--#include file="../../../conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->

<%
	dim ks:set ks=new publiccls
	dim getPartnerKey,getAppId,getSecret,getMCHID,notify_url,redirect_url
	
	dim rsconfig:set rsconfig=conn.execute("select top 1 * From KS_PaymentPlat where id=16")
	if not rsconfig.eof then
	 dim arr:arr=split(rsconfig("Md5Key")&"|||||","|")
	'以下参数自行修改
	getMCHID		= rsconfig("accountID")								'微信支付分配的商户号mch_id
	getPartnerKey	= arr(2)		'财付通密钥
	getAppId		= arr(0)						'微信分配的公众账号 appid
	getSecret		= arr(1) 		'微信分配的公众账号 srcret	
	notify_url		= KS.Setting(2) & KS.Setting(3) & "user/wxpay/notify.asp" 	'支付完成后微信将在后台发送回调处理信息,由本页面接受是否成功支付的信息
	redirect_url	= KS.Setting(2) & KS.Setting(3) & "user/wxpay/pay_ok.asp"		'支付完成后，跳转到本页面，用于展示订单支付提示，本页面可以自己修改
   else
    ks.die "配置有误，不支持微信支付！"
   end if
   
   rsconfig.close
   set rsconfig=nothing
%>
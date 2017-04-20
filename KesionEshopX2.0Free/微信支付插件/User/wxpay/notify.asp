<!--#include file="../../conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../payfunction.asp"-->

<%
Dim KS:Set KS=New PublicCls


'本接口为用户支付成功后，微信后台通知结果的接口，包括Url上的参数及XML内的参数
'可通过产品唯一订单号和支付状态，确定用户支付成功后执行的一系列操作

dim xml_dom
set xml_dom=Server.CreateObject("MSXML2.DOMDocument")
xml_dom.load Request
	
dim return_code,return_msg,result_code,err_code_des
return_code=xml_dom.getelementsbytagname("return_code").item(0).text
if return_code="FAIL" then
	'协议级错误
	return_msg=xml_dom.getelementsbytagname("return_msg").item(0).text
else
	result_code=xml_dom.getelementsbytagname("result_code").item(0).text
	if result_code="FAIL" then
		'业务级错误
		err_code_des=xml_dom.getelementsbytagname("err_code_des").item(0).text
	else
		if return_code="SUCCESS" and result_code="SUCCESS" then
			'数据正常
			dim openid,is_subscribe,trade_type,bank_type,total_fee,transaction_id,out_trade_no,time_end,attach
			openid=xml_dom.getelementsbytagname("openid").item(0).text
			is_subscribe=xml_dom.getelementsbytagname("is_subscribe").item(0).text
			trade_type=xml_dom.getelementsbytagname("trade_type").item(0).text
			bank_type=xml_dom.getelementsbytagname("bank_type").item(0).text
			total_fee=xml_dom.getelementsbytagname("total_fee").item(0).text
			transaction_id=xml_dom.getelementsbytagname("transaction_id").item(0).text
			out_trade_no=xml_dom.getelementsbytagname("out_trade_no").item(0).text
			time_end=xml_dom.getelementsbytagname("time_end").item(0).Text
			attach=xml_dom.getelementsbytagname("attach").item(0).Text
			call AddData()
		end if			
	end if
end if

dim returnXml
returnXml="<xml>"&_
		"<return_code><![CDATA[SUCCESS]]></return_code>"&_
		"</xml>"
		
sub AddData()
        dim v_amount:v_amount= total_fee / 100
		SUserName=attach
        Call UpdateOrder(v_amount,"在线充值，订单号为:" &out_trade_no,out_trade_no,"微信") 
	    Response.Write returnXml	'返回SUCCESS给微信
end Sub
%>
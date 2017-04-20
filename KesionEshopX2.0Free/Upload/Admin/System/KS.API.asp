<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<!--#include file="../../api/cls_api.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
If Not KS.ReturnPowerResult(0, "KMST10002") Then          '检查是否有基本信息设置的权限
	Call KS.ReturnErr(1, "")
	Response.End
End If

Dim Action
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveConformify
	Case Else
		Call showmain
End Select
Sub showmain()
Response.Write "<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><title>多系统整合接口设置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<link href='../include/Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<script src='../../ks_inc/jquery.js'></script>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
%>
<div class="pageCont2 mt20">
<div class="tabTitle">API整合接口设置</div>
<form name="myform" id="myform" method="post" action="?action=save">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="ctable">
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>通用设置：</strong></td>
	<td>&nbsp;<b>首次登录自动创建账号并登录：</b>
	<label><input type="radio" name="API_QuickLogin"  value="false"<%
	If Not API_QuickLogin Then Response.Write " checked"
	%>> 不启用</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_QuickLogin"  value="true"<%
	If API_QuickLogin Then Response.Write " checked"
	%>> 启用</label>
	<br/>
	&nbsp;<b>默认注册的会员用户组：</b>
	<%
	If KS.ChkClng(Api_GroupID)=0 Then Api_GroupID=2 '默认用户组
	Dim Node
	Call KS.LoadUserGroup()
	For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
	    if KS.ChkClng(Api_GroupID)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
		response.write "<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"" checked>" & Node.SelectSingleNode("@groupname").text  & "</label>"
		Else
		response.write "<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"">" & Node.SelectSingleNode("@groupname").text  & "</label>"
		End If
	Next
	%>
	</td>
</tr>

<!-- 微信登录 begin-->
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启微信账号登录：</strong></td>
	<td>
	<label><input type="radio" name="API_WeiXinEnable" onclick="$('#weixin').hide()" value="false"<%
	If Not API_WeiXinEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_WeiXinEnable" onclick="$('#weixin').show()" value="true"<%
	If API_WeiXinEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="weixin"<%if cbool(API_WeiXinEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>微信登录AppID：</strong></td>
	<td><input type="text" class="textbox" name="API_WeiXinAppId" size="35" value="<%=API_WeiXinAppId%>"> 
		<font color="red">open.weixin.qq.com 申请到的appid,<a href="https://open.weixin.qq.com/" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>微信登录AppSecret：</strong></td>
	<td><input type="text" class="textbox" name="API_WeiXinAppKey" size="35" value="<%=API_WeiXinAppKey%>"> 
		<font color="red">open.weixin.qq.com 申请到的AppSecret。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>微信登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="API_WeiXinCallBack" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/weixin/callback.asp"%>"> 
		<font class="tips">微信账号登录成功后跳转的地址,不可改。</font>
	</td>
</tr>
</tbody>
<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>
<!-- 微信登录 end-->

<!-- QQ登录 begin-->
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启QQ登录：</strong></td>
	<td>
	<label><input type="radio" name="API_QQEnable" onclick="$('#qq').hide()" value="false"<%
	If Not API_QQEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_QQEnable" onclick="$('#qq').show()" value="true"<%
	If API_QQEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="qq"<%if cbool(API_QQEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>QQ登录AppID：</strong></td>
	<td><input type="text" class="textbox" name="API_QQAppId" size="35" value="<%=API_QQAppId%>"> 
		<font color="red">opensns.qq.com 申请到的appid,<a href="http://connect.qq.com/" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>QQ登录AppKey：</strong></td>
	<td><input type="text" class="textbox" name="API_QQAppKey" size="35" value="<%=API_QQAppKey%>"> 
		<font color="red">opensns.qq.com 申请到的appkey。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>QQ登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="API_QQCallBack" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/qq/callback.asp"%>"> 
		<font class="tips">QQ登录成功后跳转的地址,不可改。</font>
	</td>
</tr>
</tbody>
<!--QQ登录 end-->



<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>

<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启新浪微博登录：</strong></td>
	<td>
	<label><input type="radio" name="API_SinaEnable" onclick="$('#sina').hide()" value="false"<%
	If Not API_SinaEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_SinaEnable" onclick="$('#sina').show()" value="true"<%
	If API_SinaEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="sina"<%if cbool(API_SinaEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>新浪微博登录App Key：</strong></td>
	<td><input type="text" class="textbox" name="API_SinaId" size="35" value="<%=API_SinaId%>"> 
		<font color="red">新浪微博登录API申请网址：http://open.weibo.com/<a href="http://open.weibo.com/" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>新浪微博登录App Secret：</strong></td>
	<td><input type="text" class="textbox" name="API_SinaKey" size="35" value="<%=API_SinaKey%>"> 
		<font color="red">新浪微博登录申请到的App Secret</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>新浪微博登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="api_sinacallback" id="api_sinacallback" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/sina/callback.asp"%>"> <font class="tips">新浪微博登录成功后跳转的地址,不可改。</font>
	</td>
</tr>

</tbody>

<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启支付宝快捷登录：</strong></td>
	<td>
	<label><input type="radio" name="API_AlipayEnable" onclick="$('#alipay').hide()" value="false"<%
	If Not API_AlipayEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_AlipayEnable" onclick="$('#alipay').show()" value="true"<%
	If API_AlipayEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="alipay"<%if cbool(API_AlipayEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>支付宝合作者身份ID：</strong></td>
	<td><input type="text" class="textbox" name="API_AlipayPartner" size="35" value="<%=API_AlipayPartner%>"> 
	<font color=red>如果还没有与支付宝签约，请<a href="https://b.alipay.com/order/slaverIndex.htm?rewardIds=vtq05uWfOIk-Ht9P1HzAYTlNX7GOvULv" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>安全检验码Key：</strong></td>
	<td><input type="text" class="textbox" name="API_AlipayKey" size="35" value="<%=API_AlipayKey%>"> 
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>支付宝快捷登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="api_alipayreturnurl" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/alipay/return_url.asp"%>"> <font class="tips">支付宝快捷登录成功后跳转的地址,不可改。</font>
	</td>
</tr>
</tbody>





<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启物流跟踪类接口：</strong></td>
	<td>
	<label><input type="radio" name="API_Deliveryapi" onclick="$('#Deliveryapi').hide()" value="false"<%
	If Not API_Deliveryapi Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_Deliveryapi" onclick="$('#Deliveryapi').show()" value="true"<%
	If API_Deliveryapi Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="Deliveryapi" <%if Not API_Deliveryapi then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>身份授权key：</strong></td>
	<td><input type="text" class="textbox" name="API_Deliveryapi_Key" size="35" value="<%=API_Deliveryapi_Key%>"> 
	<font color=red>如果还没有快递查询身份授权key，请<a href="http://www.kuaidi100.com/openapi/applyapi.shtml" target="_blank" >点此申请</a> 。</font>
	</td>
</tr>
</tbody>


<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>





<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启UCenter整合程序：</strong></td>
	<td>
	<input type="radio" name="API_Enable" onclick="$('#api').hide()" value="false"<%
	If Not API_Enable Then Response.Write " checked"
	%>> 关闭&nbsp;&nbsp;
	<input type="radio" name="API_Enable" onclick="$('#api').show()" value="true"<%
	If API_Enable Then Response.Write " checked"
	%>> 开启
	</td>
</tr>
<tbody id="api"<%if Cbool(Api_Enable)=false Then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>Ucenter应用密钥：</strong></td>
	<td><input type="text" name="API_ConformKey" class="textbox" size="35" value="<%=API_ConformKey%>"> 
		<font color="red">Ucenter 中此应用的密钥。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>Ucenter应用ID(APP ID)：</strong></td>
	<td><input type="text" name="API_Debug" class="textbox"  size="35" value="<%=API_Debug%>"> 
	&nbsp;&nbsp;<font color="red">Ucenter中此应用的ID号</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>Ucenter安装地址(URL)：</strong></td>
    <td><input type="text" name="API_Urls"  class="textbox" size="45" value="<%=API_Urls%>"> </td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>Ucenter安装地址(IP)：</strong></td>
	<td><input type="text" name="API_LoginUrl" class="textbox"  size="45" value="<%=API_LoginUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>DZ数据库连接字符串：</strong></td>
	<td><textarea name="API_ReguserUrl"  class="textbox" style="width:600px;height:80px"><%=API_ReguserUrl%></textarea><br/> 
		<font color="red">不设置请留空,但新注册会员,首页注册无法直接在DZ论坛同步登录,需要在DZ论坛登录后,第二次才能同步登录。</font>
	</td>
</tr>
<tr class="tdbg"  style="display:none">
	<td height="30" class="clefttitle" align="right"><strong>整合用户注销后转向URL：</strong></td>
	<td><input type="text" name="API_LogoutUrl"  class="textbox" size="45" value="<%=API_LogoutUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
</tbody>
</table>
</form>
</div>
<script>
 function CheckForm()
 {
  $("#myform").submit();
 }
</script>
<%
End Sub

Sub SaveConformify()
	Dim XslDoc,XslNode,Xsl_Files
	Xsl_Files = API_Path & "api.config"
	Xsl_Files = Server.MapPath(Xsl_Files)
	Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If Not XslDoc.Load(Xsl_Files) Then
		Response.Write "初始数据不存在！"
		Response.End
	Else
		Set XslNode = XslDoc.documentElement.selectSingleNode("rs:data/z:row")
		XslNode.attributes.getNamedItem("api_enable").text = Trim(Request.Form("API_Enable"))
		XslNode.attributes.getNamedItem("api_conformkey").text = ChkRequestForm("API_ConformKey")
		XslNode.attributes.getNamedItem("api_urls").text = ChkRequestForm("API_Urls")
		XslNode.attributes.getNamedItem("api_debug").text = ChkRequestForm("API_Debug")
		XslNode.attributes.getNamedItem("api_loginurl").text = ChkRequestForm("API_LoginUrl")
		XslNode.attributes.getNamedItem("api_reguserurl").text = ChkRequestForm("API_ReguserUrl")
		XslNode.attributes.getNamedItem("api_logouturl").text = ChkRequestForm("API_LogoutUrl")
		'XslNode.attributes.setNamedItem(XslDoc.createNode(2,"date","")).text = Now()
		'XslNode.appendChild(XslDoc.createNode(1,"pubDate","")).text = Now()
		XslNode.attributes.getNamedItem("api_quicklogin").text =trim(Request.Form("API_QuickLogin"))
		XslNode.attributes.getNamedItem("api_groupid").text =trim(Request.Form("GroupID"))
		XslNode.attributes.getNamedItem("api_qqenable").text =trim(Request.Form("API_QQEnable"))
		XslNode.attributes.getNamedItem("api_qqappid").text =ChkRequestForm("API_QQAppId")
		XslNode.attributes.getNamedItem("api_qqappkey").text =ChkRequestForm("API_QQAppKey")
		XslNode.attributes.getNamedItem("api_qqcallback").text =ChkRequestForm("API_QQCallBack")
		
		XslNode.attributes.getNamedItem("api_weixinenable").text =trim(Request.Form("API_WeiXinEnable"))
		XslNode.attributes.getNamedItem("api_weixinappid").text =ChkRequestForm("API_WeiXinAppId")
		XslNode.attributes.getNamedItem("api_weixinappkey").text =ChkRequestForm("API_WeiXinAppKey")
		XslNode.attributes.getNamedItem("api_weixincallback").text =ChkRequestForm("API_WeiXinCallBack")

		
		XslNode.attributes.getNamedItem("api_alipayenable").text =trim(Request.Form("API_AlipayEnable"))
		XslNode.attributes.getNamedItem("api_alipaypartner").text =ChkRequestForm("API_AlipayPartner")
		XslNode.attributes.getNamedItem("api_alipaykey").text =ChkRequestForm("API_AlipayKey")
		XslNode.attributes.getNamedItem("api_alipayreturnurl").text =ChkRequestForm("API_AlipayReturnUrl")
		
		XslNode.attributes.getNamedItem("api_deliveryapi").text =trim(Request.Form("API_Deliveryapi"))
		XslNode.attributes.getNamedItem("api_deliveryapi_key").text =ChkRequestForm("API_Deliveryapi_Key")
		
		XslNode.attributes.getNamedItem("api_sinaenable").text =trim(Request.Form("API_SinaEnable"))
		XslNode.attributes.getNamedItem("api_sinaid").text =ChkRequestForm("API_SinaId")
		XslNode.attributes.getNamedItem("api_sinakey").text =ChkRequestForm("API_SinaKey")
		XslNode.attributes.getNamedItem("api_sinacallback").text =ChkRequestForm("API_SinaCallBack")

		XslDoc.save Xsl_Files
		Set XslNode = Nothing
	End If
	Set XslDoc = Nothing
	Response.Write ("<script>alert('恭喜您！保存设置成功。');location.href='KS.Api.asp';</script>")
End Sub
Function ChkRequestForm(reform)
	Dim strForm
	strForm = Trim(Request.Form(reform))
	If IsNull(strForm) Then
		strForm = "0"
	Else
		strForm = Replace(strForm, Chr(0), vbNullString)
		strForm = Replace(strForm, Chr(34), vbNullString)
		strForm = Replace(strForm, "'", vbNullString)
		strForm = Replace(strForm, """", vbNullString)
	End If
	If strForm = "" Then strForm = ""
	ChkRequestForm = strForm
End Function

%>
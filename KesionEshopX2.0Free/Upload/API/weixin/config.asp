<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.membercls.asp"-->
<!--#include file="../cls_api.asp"-->
<%
'请将下面信息更改成自己申请的信息
Dim appid  : appid   = API_WeiXinAppId  'open.weixin.qq.com 申请到的appid
Dim appkey : appkey  = API_WeiXinAppKey 'open.weixin.qq.com 申请到的appkey
Dim callback:callback = API_WeiXinCallBack '微信登录成功后跳转的地址


'生成时间戳 
Function ToUnixTime(strTime, intTimeZone)
If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now
If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0
ToUnixTime = DateAdd("h",-intTimeZone,strTime)
ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)
End Function

'生成随机数
Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
End Function

%>

<%
	dim phpapiurl,create_ip,nonce_str,timeStamp,xmlhttp,preCookies
	phpapiurl=KS.Setting(2) & KS.Setting(3) & "user/wxpay/turn.aspx"
	preCookies		= "WOS23294I3"	'Cookie前缀，同一个网站下，多个支付需要修改
	create_ip		= Request.ServerVariables("REMOTE_ADDR")
	nonce_str		= GetRnd(10)
	timeStamp		= ToUnixTime(now())
	xmlhttp			= "Msxml2.ServerXMLHTTP.6.0"
	
	'微信支付V3获取Prepay_Id
	function get_prepay_id()
		dim postData,signValue,post_url,sign,returnXml,xml_dom,return_code,result_code
		sign="appid="&getAppId&"&attach=" &attach& "&body="&body&"&mch_id="&getMCHID&"&nonce_str="&nonce_str&"&notify_url="&notify_url&"&openid="&openid&"&out_trade_no="&out_trade_no&"&spbill_create_ip="&create_ip&"&total_fee="&total_fee&"&trade_type=JSAPI&key="&getPartnerKey
		signValue=UCase(MD5(sign,"UTF-8"))
		postData="<xml>"&_
			"<appid><![CDATA["&getAppId&"]]></appid>"&_
			"<attach><![CDATA["&attach&"]]></attach>"&_
			"<body><![CDATA["&body&"]]></body>"&_
			"<mch_id><![CDATA["&getMCHID&"]]></mch_id>"&_
			"<nonce_str><![CDATA["&nonce_str&"]]></nonce_str>"&_
			"<notify_url><![CDATA["&notify_url&"]]></notify_url>"&_
			"<openid><![CDATA["&openid&"]]></openid>"&_
			"<out_trade_no><![CDATA["&out_trade_no&"]]></out_trade_no>"&_
			"<spbill_create_ip><![CDATA["&create_ip&"]]></spbill_create_ip>"&_
			"<total_fee><![CDATA["&total_fee&"]]></total_fee>"&_
			"<trade_type><![CDATA[JSAPI]]></trade_type>"&_
			"<sign><![CDATA["&signValue&"]]></sign>"&_
			"</xml>"
		returnXml=PostURL(phpapiurl&"?rnd="&now,postData)
		get_prepay_id=server.HTMLEncode(returnXml)
		
	end Function
	
	'微信支付V3，返回最后提交的paySign
	function get_paySign()
		dim sign
		sign="appId="&getAppId&"&nonceStr="&nonce_str&"&package=prepay_id="&prepay_id&"&signType=MD5&timeStamp="&timeStamp&"&key="&getPartnerKey
		get_paySign=UCase(MD5(sign,"UTF-8"))
	end function
	
	
	function GetOpenId()
		'获取用户OpenID
		if request.Cookies(preCookies&"openid")="" then		
			dim code,myurl,url,strJson,access_token,openids
			code=request("code")
			myurl="http://"&Request.ServerVariables("Server_Name")&Request.ServerVariables("URL")
			if request.ServerVariables("QUERY_STRING")<>"" then 
				myurl = myurl &"?"& Request.ServerVariables("QUERY_STRING")
			end if
			myurl=Server.URLEncode(myurl)
			if code="" then
				response.Redirect("https://open.weixin.qq.com/connect/oauth2/authorize?appid="&getAppId&"&redirect_uri="&myurl&"&response_type=code&scope=snsapi_base&state=STATE#wechat_redirect")
				response.End()
			else
				url="https://api.weixin.qq.com/sns/oauth2/access_token?appid="&getAppId&"&secret="&getSecret&"&code="&code&"&grant_type=authorization_code"
				strJson=GetURL(url)
				dim objTest
				Call InitScriptControl:Set objTest = getJSONObject(strJson)
				if InStr(strJson,"errcode")>0 then
					response.Write "获取Openid出错："&strJson
					response.End()
				else
					openids=objTest.openid	'获取openid
					Response.Cookies(preCookies&"openid")=openids
					Response.Cookies(preCookies&"openid").Expires=DateAdd("m",60,now())
					GetOpenId=openids
				end if
			end if
		else
			GetOpenId=request.Cookies(preCookies&"openid")
		end if
	End function

	'返回当前日期20140105024523
	Function getStrNow()
		dim strNow:strNow = Now()
		strNow = Year(strNow) & Right(("00" & Month(strNow)),2) & Right(("00" & Day(strNow)),2) & Right(("00" & Hour(strNow)),2) & Right(("00" &  Minute(strNow)),2) & Right(("00" & Second(strNow)),2)
		getStrNow = strNow
	End Function
	
	'获取随机数,返回 [min,max]范围的数
	Function getRandNumber(max, min)
		Randomize 
		getRandNumber = CInt((max-min+1)*Rnd()+min) 
	End Function
		
	'获取随机数字的字符串,返回[min,max]范围的数字字符串
	Function getStrRandNumber(max, min)
		dim randNumber:randNumber = getRandNumber(max, min)
		getStrRandNumber = CStr(randNumber)
	End Function	
	
	Function GetRnd(t0)
		randomize
		dim n1,n2,n3
		do while len(getrnd)<t0 '随机字符位数 
			n1=cstr(chrw((57-48)*rnd+48)) '0~9 
			n2=cstr(chrw((90-65)*rnd+65)) 'a~z 
			n3=cstr(chrw((122-97)*rnd+97)) 'a~z 
			getrnd=getrnd&n1&n2&n3 
		loop
	End Function	
	
	'时间戳转换成普通日期
	Function FromUnixTime(intTime) 
		If IsEmpty(intTime) Or Not IsNumeric(intTime) Then 
			FromUnixTime = Now() 
			Exit Function 
		End If 	
		FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0") 
		FromUnixTime = DateAdd("h", 8, FromUnixTime) 
	End Function
	
	'普通日期转换成时间戳
	Function ToUnixTime(strTime)        
		If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now   
		 ToUnixTime = DateAdd("h",-8,strTime)        
		 ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)        
	End Function 	
		
	Dim sc4Json   
	Sub InitScriptControl
		Set sc4Json = Server.CreateObject("MSScriptControl.ScriptControl")    
		sc4Json.Language = "JavaScript"    
		sc4Json.AddCode "var itemTemp=null;function getJSArray(arr, index){itemTemp=arr[index];}"    
	End Sub 
	Function getJSONObject(strJSON)    
		sc4Json.AddCode "var jsonObject = " & strJSON    
		Set getJSONObject = sc4Json.CodeObject.jsonObject    
	End Function 
	Sub getJSArrayItem(objDest,objJSArray,index)    
		On Error Resume Next    
		sc4Json.Run "getJSArray",objJSArray, index    
		Set objDest = sc4Json.CodeObject.itemTemp    
		If Err.number=0 Then Exit Sub    
		objDest = sc4Json.CodeObject.itemTemp    
	End Sub
	
	Function PostURL(url,PostStr)
		dim http
		Set http = Server.CreateObject(xmlhttp)
		With http
			.Open "POST", url, false ,"" ,""
			.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			.Send(PostStr)
			PostURL = .responsetext
		End With
		Set http = Nothing
	End Function
	
	Function GetURL(url)	
		dim http
		set http=server.createobject(xmlhttp)
		http.open "GET",url,false
		http.setRequestHeader "If-Modified-Since","0"
		http.send()
		GetURL=http.responsetext
		set http=nothing
	End Function	

	Function IsInstall(byval t0)
		err.clear
		on error resume next
		IsInstall=false
		dim obj
		set obj=server.createobject(t0)
			if err.number=0 then IsInstall=true
		set obj=nothing
		err.clear()
	End Function

		''转换HTML代码，过滤代码
	Function enhtml(byval t0)
		if isnull(t0) then enhtml="":exit function
		if t0="<p>&nbsp;</p>" then enhtml="":exit function
		t0=replace(t0,"&","&amp;")
		t0=replace(t0,"'","&#39;")
		t0=replace(t0,"""","&#34;")
		t0=replace(t0,"<","&lt;")
		t0=replace(t0,">","&gt;")
		enhtml=t0
	End Function
	
	sub OutPutTxt(str)
		dim FilePath,Fso,fopen
		filepath=server.mappath("wx.txt")
		Set fso = Server.CreateObject("scripting.FileSystemObject")
		set fopen=fso.OpenTextFile(filepath, 8 ,true)
		fopen.writeline(str)
		set fso=nothing
		set fopen=Nothing
	end sub
%>
<!--#include file="config.asp"-->
<%

Function uc_api_post(module,action,arg)

	dim s,postdata
	s = arg
	postdata = uc_api_requestdata(module,action,s,"")
	'response.write postdata
	uc_api_post =  uc_fopen2(UC_API & "/index.php", 500000, postdata, "", true, UC_IP, 20,true)
End Function

Function uc_user_synlogin(uid)
    uc_user_synlogin = uc_api_post("user", "synlogin", "uid="&uid)
End Function

Function uc_user_login(user,pwd)
	uc_user_login = uc_api_post("user","login","username="&user&"&password="&pwd&"&isuid=0&checkques=0&questionid=&answer=")
	'返回一个XML数组
End Function 

Function uc_user_synlogout()
	uc_user_synlogout = uc_api_post("user","synlogout","uid="&uid)

End Function 

Function uc_user_edit(username , oldpw , newpw , email ,ignoreoldpw, questionid,answer)
	uc_user_edit = uc_api_post("user","edit","username="&username&"&newpw="&newpw&"&oldpw="&oldpw&"&email="&email&"&questionid="&questionid&"&answer="&answer&"&ignoreoldpw="&ignoreoldpw)
End Function 

Function uc_user_register(user,pwd,email)
	uc_user_register = uc_api_post("user","register","username="&user&"&password="&pwd&"&email="&email&"&questionid=&answer=&regip=")
End Function 

Function uc_user_delete(user)
    Dim XML:XML = uc_api_post("user","get_user","username="&user)
	Dim arr_login:arr_login =  xml2array(XML)
	Dim uid:uid=KS.ChkClng(arr_login(0))
	if uid>0 Then
	uc_user_delete = uc_api_post("user","delete","uid="&uid)
	end if
End Function 

Function uc_avatar(uid,stype)
	uid = CInt(uid)	
	input = uc_api_input("uid="&uid)
	uc_avatarflash = UC_API&"/images/camera.swf?inajax=1&appid="&UC_APPID&"&input="&input&"&agent="&MD5(Request.ServerVariables ("HTTP_USER_AGENT"),32)&"&ucapi="&Server.URLEncode(replace(UC_API,"http://", ""))&"&avatartype="&stype&"&uploadSize=2048"
	uc_avatar="<object classid='clsid:d27cdb6e-ae6d-11cf-96b8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0' width='450' height='253' id='mycamera' align='middle'><param name='allowScriptAccess' value='always' /><param name='scale' value='exactfit' /><param name='wmode' value='transparent' /><param name='quality' value='high' /><param name='bgcolor' value='#ffffff' /><param name='movie' value='"&uc_avatarflash&"' /><param name='menu' value='false' /><embed src='"&uc_avatarflash&"' quality='high' bgcolor='#ffffff' width='450' height='253' name='mycamera' align='middle' allowScriptAccess='always' allowFullScreen='false' scale='exactfit'  wmode='transparent' type='application/x-shockwave-flash' pluginspage='http://www.macromedia.com/go/getflashplayer' /></object>"
End Function 

Function uc_api_requestdata(module,action,arg,extra)
	dim input
	input = uc_api_input(arg)
	uc_api_requestdata = "m="&module&"&a="&action&"&inajax=2&input="&input&"&appid="&UC_APPID&extra
	
End Function

Function uc_api_input(data)
	'Response.write "<br>"&data&"<br>"
	uc_api_input = Server.URLEncode(uc_authcode(data&"&agent="&MD5(Request.ServerVariables("HTTP_USER_AGENT"),32)&"&time="&php_time(), "ENCODE", UC_KEY))
End Function

Function uc_fopen(url, limit, post, cookie, bysocket, ip, stimeout, block)
	dim objXmlHttp
	set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
	objXmlHttp.open "post",url,False
	objXmlHttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    objXmlHttp.setRequestHeader "Accept", "*/*"   
    objXmlHttp.setRequestHeader "Accept-Language","zh-cn"   
    objXmlHttp.setRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")    
    objXmlHttp.setRequestHeader "Connection","Close"   
    objXmlHttp.setRequestHeader "Cache-Control","no-cache"   
	objXmlHttp.send(post)
	Dim binFileData,s
	binFileData = objXmlHttp.responseBody
	Dim ObjStream
	Set ObjStream = CreateObject("Adodb.Stream")
	With ObjStream
		.Type = 1
		.Mode = 3
		.Open
		.write binFileData
		.Position = 0
		.Type = 2
		.Charset = "utf-8"
		s = .ReadText
		.Close
	End With
	Set ObjStream = Nothing
	uc_fopen = s
End function


function uc_fopen2(url, limit, post, cookie,bysocket, ip , stimeout, block)
	Dim times
	if request("__times__")<>"" Then
		times = Cint(request("__times__")) + 1
	Else
		times =  1
	End If
	if times > 2 Then
		uc_fopen2 = ""
	End If
	If instr(url,"?")>0 Then
		url = url & "&__times__=" & times
	Else
		url = url & "?__times__=" & times
	End If
	uc_fopen2 = uc_fopen(url, limit, post, cookie, bysocket, ip, stimeout, block)
End Function


Function xml2array(xmldoc)
	Dim objdoc,i
	set objdoc=Server.CreateObject("msxml2.FreeThreadedDOMDocument.3.0")
	objdoc.async = false
	objdoc.LoadXml(xmldoc)
	Dim objNodeList
	Dim Counters(14)
	set objNodeList = objdoc.selectSingleNode("//root").childNodes
	For i = 0 To (objNodeList.length - 1)
		Counters(i)= objNodeList.Item(i).text
	Next
	xml2array = Counters
End Function 
%>
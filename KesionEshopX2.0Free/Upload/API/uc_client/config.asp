<!--#include file="php_base64.asp"-->
<%
'**************************************************
' ASP-Client to UCenter-Server 通信应用范例
' ************************************************
' 网上的一些代码号称UCenter 整合ASp，但是大部分都没有一个实例，只有一个加密解密函数，很多新手无法利用其进行实际应用。
' 公司网站项目--课堂无忧，需整合asp和ucenter，研究了一把UCenter，写了这个范例程序，虽然只简单写了登陆和注册部分功能的实现，其余模块您可以依葫芦画瓢。
' 程序发布人QQ：59466966 
'**************************************************


'----------------------------------------------
Dim UC_APPID:UC_APPID = API_Debug	'ucenter中此应用的ID号
Dim UC_KEY:UC_KEY = API_ConformKey	'PHP ucenter中此应用密钥
Dim UC_IP:UC_IP = API_LoginUrl
Dim UC_API:UC_API = API_Urls	 'PHP ucenter地址 ,末尾不要/
'-------------------------------------------

function uc_authcode(ByVal str, ByVal operation, ByVal key)
  uc_authcode = ""
  if Len(str)<4 then Exit Function
  
  Const ckey_length = 4
  Dim keya, keyb, keyc, md5MT, cryptkey, key_length, string_length, box, rndkey, i, j, tmp, iTmp, tmpResult, a, x1, x2
  
  key = MD5(key,32)
  keya = md5(Left(key, 16),32)
  keyb = MD5(Right(key, 16),32)
  md5MT = MD5(microtime(),32)
  if ckey_length > 0 then
  	if operation="DECODE" then
  		keyc = Left(str, ckey_length)
  	else
  		keyc = Right(md5MT, ckey_length)
  	end if
	else
		keyc = ""
	end if

  cryptkey = keya & MD5(keya & keyc,32)
  key_length = Len(cryptkey)
  
	Redim box(255)
  for i = 0 to 255
  	box(i) = i
  next

  Redim rndkey(255)
  for i = 0 to 255
    rndkey(i) = Asc(Mid(cryptkey, i Mod key_length + 1, 1))
  next

	j = 0
  for i = 0 to 255
      j = (j + box(i) + rndkey(i)) Mod 256
      tmp = box(i)
      box(i) = box(j)
      box(j) = tmp
  next  

	if operation="DECODE" then
		tmp = Right(str, Len(str)-ckey_length)
  	tmpResult = php_Base64Decode(tmp)
  	string_length = UBound(tmpResult)+1

	  a = 0
	  j = 0
	  for i = 0 to string_length-1
	      a = (a + 1) Mod 256
	      j = (j + box(a)) Mod 256
	      tmp = box(a)
	      box(a) = box(j)
	      box(j) = tmp
	
	      x1 = tmpResult(i)
	      iTmp = (box(a) + box(j)) Mod 256
	      x2 = box(iTmp)
	      tmpResult(i) = x1 Xor x2
	  next

		tmp = ""
		for i=0 to 9
			tmp = tmp & ChrB(tmpResult(i))
		next
		If IsNumeric(tmp) then iTmp = CLng(tmp) else iTmp = 0

    x1 = ""
    for i=10 to 25
    	x1 = x1 & Chr(tmpResult(i))
  	next

  	tmp = ""
    for i=26 to UBound(tmpResult)
    	tmp = tmp & ChrB(tmpResult(i))
  	next
  	tmp = strAnsi2Unicode(tmp)
    x2 = Left(MD5(tmp & keyb,32), 16)

    if (iTmp = 0 Or iTmp > php_time()) And (x1=x2) then
        uc_authcode = tmp
    else
        uc_authcode = ""
    end if
      
  else
  	str = "0000000000" & Left(MD5(str + keyb,32), 16) & str
  	str = strUnicode2Ansi(str)
  	string_length = LenB(str)
  	Redim tmpResult(string_length-1)

	  a = 0
	  j = 0
	  for i = 0 to string_length-1
	      a = (a + 1) Mod 256
	      j = (j + box(a)) Mod 256
	      tmp = box(a)
	      box(a) = box(j)
	      box(j) = tmp
	
	      x1 = AscB(MidB(str, i+1, 1))
	      iTmp = (box(a) + box(j)) Mod 256
	      x2 = box(iTmp)
	      tmpResult(i) = x1 Xor x2
	  next

	  uc_authcode = keyc & Replace(php_Base64Encode(tmpResult),"=","")
  end if
end function


Const TimeZone=8	'服务器所在时区

function php_time()
	php_time = dateadd("h", TimeZone*-1, now())
	php_time = datediff("s", "1970-01-01 00:00:00", php_time)
end function

function microtime()
  Dim sec, msec, i, s
  sec = php_time()
  msec = timer()*1000 Mod 1000
  i = Max(0, 8-Len(Cstr(msec)))
  s = String(i,"0")
  microtime = "0." & msec & s & " " & sec
end function

Private Function Min(x,y)
  if x<y then Min=x else Min=y
End Function

Private Function Max(x,y)
  if x>y then Max=x else Max=y
End Function


%>
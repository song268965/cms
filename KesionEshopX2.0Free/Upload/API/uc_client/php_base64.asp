<%

Const BASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
dim sBASE_64_CHARACTERS
dim len1,k
dim asc1,asContents1
dim varchar,varasc,varHex,varlow,varhigh
sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"  
sBASE_64_CHARACTERS = strUnicode2Ansi(sBASE_64_CHARACTERS)

Public Function php_Base64Encode(ByVal ByteArr)
    Dim lnPosition
    Dim lsResult
    Dim Char1
    Dim Char2
    Dim Char3
    Dim Char4
    Dim Byte1
    Dim Byte2
    Dim Byte3
    Dim SaveBits1
    Dim SaveBits2
    Dim lsGroupBinary
    Dim lsGroup64
    Dim M4, len1, len2
    Dim i
    
    len1 = UBound(ByteArr)+1
    If len1 < 1 Then
        php_Base64Encode = ""
        Exit Function
    End If

    M4 = len1 Mod 3
    If M4 > 0 Then
    	Redim Preserve ByteArr(len1+(3 - M4)-1)
    	for i=len1 to UBound(ByteArr)
    		ByteArr(i) = 0
    	next
   	End If
    '补足位数是为了便于计算

    If M4 > 0 Then
        len1 = len1 + (3 - M4)
        len2 = len1 - 3
    Else
        len2 = len1
    End If

    lsResult = ""
    
    For lnPosition = 0 To len2-1 Step 3
        Byte1 = ByteArr(lnPosition): SaveBits1 = Byte1 And 3
        Byte2 = ByteArr(lnPosition+1): SaveBits2 = Byte2 And 15
        Byte3 = ByteArr(lnPosition+2)

        Char1 = Mid(BASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
        Char2 = Mid(BASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Char3 = Mid(BASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
        Char4 = Mid(BASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
        
        lsResult = lsResult & (Char1 & Char2 & Char3 & Char4)
    Next

    '处理最后剩余的几个字符
    If M4 > 0 Then
        Byte1 = ByteArr(len2): SaveBits1 = Byte1 And 3
        Byte2 = ByteArr(len2+1): SaveBits2 = Byte2 And 15
        Byte3 = ByteArr(len2+2)

        Char1 = Mid(BASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
        Char2 = Mid(BASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Char3 = Mid(BASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)

        If M4 = 1 Then
            lsResult = lsResult & (Char1 & Char2 & Chr(61) & Chr(61)) '用=号补足位数
        Else
            lsResult = lsResult & (Char1 & Char2 & Char3 & Chr(61)) '用=号补足位数
        End If
    End If
    
		php_Base64Encode = lsResult
End Function


Public Function php_Base64Decode(ByVal asContent)
	  while Len(asContent) Mod 4 <> 0
	      asContent = asContent & "="
	  wend

    Dim lsResult()
    Dim lnPosition
    Dim lsGroup64, lsGroupBinary
    Dim Char1, Char2, Char3, Char4
    Dim Byte1, Byte2, Byte3
    Dim M4, len1, len2
    Dim iPos

    len1 = Len(asContent)
    M4 = len1 Mod 4

    If len1 < 1 Or M4 > 0 Then
        '字符串长度应当是4的倍数
        php_Base64Decode = ""
        Exit Function
    End If

    '判断最后一位是不是 = 号
    '判断倒数第二位是不是 = 号
    '这里m4表示最后剩余的需要单独处理的字符个数
    If Mid(asContent, len1, 1) = Chr(61) Then M4 = 3
    If Mid(asContent, len1 - 1, 1) = Chr(61) Then M4 = 2

    If M4 = 0 Then
        len2 = len1
    Else
        len2 = len1 - 4
    End If
    
    iPos = 0
    Redim lsResult(((Len(asContent) \ 4)-1)*3+10)

    For lnPosition = 1 To len2 Step 4
        lsGroup64 = Mid(asContent, lnPosition, 4)
        
        Char1 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 1, 1)) - 1
        Char2 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 2, 1)) - 1
        Char3 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 3, 1)) - 1
        Char4 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 4, 1)) - 1
        
        Byte1 = ((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF
        Byte2 = ((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF
        Byte3 = (((Char3 And 3) * 64) And &HFF) Or (Char4 And 63)

        lsResult(iPos) = Byte1 : iPos = iPos+1
        lsResult(iPos) = Byte2 : iPos = iPos+1
        lsResult(iPos) = Byte3 : iPos = iPos+1
    Next

    '处理最后剩余的几个字符
    If M4 > 0 Then
        lsGroup64 = Mid(asContent, len2 + 1, M4) & Chr(65) 'chr(65)=A，转换成值为0
        If M4 = 2 Then '补足4位，是为了便于计算
            lsGroup64 = lsGroup64 & Chr(65)
        End If
        
        Char1 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 1, 1)) - 1
        Char2 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 2, 1)) - 1
        Char3 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 3, 1)) - 1
        Char4 = InStr(BASE_64_CHARACTERS, Mid(lsGroup64, 4, 1)) - 1
        
        Byte1 = ((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF
        Byte2 = ((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF
        Byte3 = (((Char3 And 3) * 64) And &HFF) Or (Char4 And 63)

        If M4 = 2 Then
            lsResult(iPos) = Byte1 : iPos = iPos+1
        ElseIf M4 = 3 Then
            lsResult(iPos) = Byte1 : iPos = iPos+1
            lsResult(iPos) = Byte2 : iPos = iPos+1
        End If
    End If

		Redim Preserve lsResult(iPos-1)
    php_Base64Decode = lsResult
End Function

Private Function strUnicodeLen(ByVal asContents)
    '计算unicode字符串的Ansi编码的长度
    Dim asContents1
    Dim len1
    Dim k
    Dim i
    Dim asc1
    
    asContents1 = "a" & asContents
    len1 = Len(asContents1)
    k = 0
    For i = 1 To len1
        asc1 = Asc(Mid(asContents1, i, 1))
        If asc1 < 0 Then asc1 = 65536 + asc1
        If asc1 > 255 Then
            k = k + 2
            Else
            k = k + 1
        End If
    Next
    strUnicodeLen = k - 1
End Function

Private Function strUnicode2Ansi(ByVal asContents)
    '将Unicode编码的字符串，转换成Ansi编码的字符串
    Dim len1
    Dim i
    Dim VarChar
    Dim varAsc
    Dim varHex, varlow, varhigh
    
    strUnicode2Ansi = ""
    len1 = Len(asContents)
    For i = 1 To len1
        VarChar = Mid(asContents, i, 1)
        varAsc = Asc(VarChar)
        If varAsc < 0 Then varAsc = varAsc + 65536
        If varAsc > 255 Then
            varHex = Hex(varAsc)
            varlow = Left(varHex, 2)
            varhigh = Right(varHex, 2)
            strUnicode2Ansi = strUnicode2Ansi & ChrB("&H" & varlow) & ChrB("&H" & varhigh)
        Else
            strUnicode2Ansi = strUnicode2Ansi & ChrB(varAsc)
        End If
    Next
End Function

Private Function strAnsi2Unicode(asContents)
    '将Ansi编码的字符串，转换成Unicode编码的字符串
    Dim len1
    Dim i
    Dim VarChar
    Dim varAsc
    
    strAnsi2Unicode = ""
    len1 = LenB(asContents)
    If len1 = 0 Then Exit Function
    For i = 1 To len1
        VarChar = MidB(asContents, i, 1)
        varAsc = AscB(VarChar)
        If varAsc > 127 Then
            strAnsi2Unicode = strAnsi2Unicode & Chr(AscW(MidB(asContents, i + 1, 1) & VarChar))
            i = i + 1
        Else
            strAnsi2Unicode = strAnsi2Unicode & Chr(varAsc)
        End If
    Next
End Function



Function Base64encode(asContents)  
'将Ansi编码的字符串进行Base64编码
'asContents应当是ANSI编码的字符串（二进制的字符串也可以）
Dim lnPosition  
Dim lsResult  
Dim Char1  
Dim Char2  
Dim Char3  
Dim Char4  
Dim Byte1  
Dim Byte2  
Dim Byte3  
Dim SaveBits1  
Dim SaveBits2  
Dim lsGroupBinary  
Dim lsGroup64  
Dim m3,m4,len1,len2

len1=Lenb(asContents)
if len1<1 then 
   Base64encode=""
   exit Function
end if

m3=Len1 Mod 3 
If M3 > 0 Then asContents = asContents & String(3-M3, chrb(0))  
'补足位数是为了便于计算

IF m3 > 0 THEN 
   len1=len1+(3-m3)
   len2=len1-3
else
   len2=len1
end if

lsResult = ""  

For lnPosition = 1 To len2 Step 3  
    lsGroup64 = ""  
    lsGroupBinary = Midb(asContents, lnPosition, 3)  

    Byte1 = Ascb(Midb(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3  
    Byte2 = Ascb(Midb(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15  
    Byte3 = Ascb(Midb(lsGroupBinary, 3, 1))  

    Char1 = Midb(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)  
    Char2 = Midb(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)  
    Char3 = Midb(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)  
    Char4 = Midb(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)  
    lsGroup64 = Char1 & Char2 & Char3 & Char4  
    
    lsResult = lsResult & lsGroup64  
Next  

'处理最后剩余的几个字符
if M3 > 0  then
    lsGroup64 = ""  
    lsGroupBinary = Midb(asContents, len2+1, 3)
    Byte1 = Ascb(Midb(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3  
    Byte2 = Ascb(Midb(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15  
    Byte3 = Ascb(Midb(lsGroupBinary, 3, 1))  

    Char1 = Midb(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)  
    Char2 = Midb(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)  
    Char3 = Midb(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)  

    if M3=1 then
       lsGroup64 = Char1 & Char2 & ChrB(61) & ChrB(61)   '用=号补足位数
    else
       lsGroup64 = Char1 & Char2 & Char3 & ChrB(61)      '用=号补足位数
    end if
    
    lsResult = lsResult & lsGroup64  
end if

Base64encode = lsResult  

End Function  


Function Base64decode(asContents)  
'将Base64编码字符串转换成Ansi编码的字符串
'asContents应当也是ANSI编码的字符串（二进制的字符串也可以）
Dim lsResult  
Dim lnPosition  
Dim lsGroup64, lsGroupBinary  
Dim Char1, Char2, Char3, Char4  
Dim Byte1, Byte2, Byte3  
Dim M4,len1,len2

len1= Lenb(asContents) 
M4 = len1 Mod 4

if len1 < 1 or M4 > 0 then
   '字符串长度应当是4的倍数
   Base64decode = ""  
   exit Function  
end if
       
'判断最后一位是不是 = 号
'判断倒数第二位是不是 = 号
'这里m4表示最后剩余的需要单独处理的字符个数
if midb(asContents, len1, 1) = chrb(61)   then   m4=3 
if midb(asContents, len1-1, 1) = chrb(61) then   m4=2

if m4 = 0 then
   len2=len1
else
   len2=len1-4
end if

For lnPosition = 1 To Len2 Step 4  
    lsGroupBinary = ""  
    lsGroup64 = Midb(asContents, lnPosition, 4)  
    Char1 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 1, 1)) - 1  
    Char2 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 2, 1)) - 1  
    Char3 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 3, 1)) - 1  
    Char4 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 4, 1)) - 1  
    Byte1 = Chrb(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)  
    Byte2 = lsGroupBinary & Chrb(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)  
    Byte3 = Chrb((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))  
    lsGroupBinary = Byte1 & Byte2 & Byte3  
    
    lsResult = lsResult & lsGroupBinary  
Next 

'处理最后剩余的几个字符
if M4 > 0 then 
    lsGroupBinary = ""  
    lsGroup64 = Midb(asContents, len2+1, m4) & chrB(65)   'chr(65)=A，转换成值为0
    if M4=2 then                                          '补足4位，是为了便于计算 
        lsGroup64 = lsGroup64 & chrB(65)                  
    end if
    Char1 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 1, 1)) - 1  
    Char2 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 2, 1)) - 1  
    Char3 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 3, 1)) - 1  
    Char4 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 4, 1)) - 1  
    Byte1 = Chrb(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)  
    Byte2 = lsGroupBinary & Chrb(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)  
    Byte3 = Chrb((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))  
  
    if M4=2 then
       lsGroupBinary = Byte1
    elseif M4=3 then
       lsGroupBinary = Byte1 & Byte2
    end if
    
    lsResult = lsResult & lsGroupBinary  
end if

Base64decode = lsResult  

End Function  

%>

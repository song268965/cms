<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.5
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
Dim KSUser:Set KSUser = New UserCls
Dim ID:ID=KS.ChkClng(Request("id"))
If ID=0 Then KS.Die "error!"
Select Case Request("action")
 case "checklogin" checklogin
 case "savemsg" savemsg
 case "loadmsg" loadmsg
 case "loadtextrecord" loadtextrecord
End Select

Set KS=Nothing
Set KSUser=Nothing

Sub checklogin()
   Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select top 1 * From KS_InterView Where ID=" & ID,conn,1,1
   If RS.Eof And RS.Bof Then
      KS.Echo Escape("找不到访谈主题!")
   Else
	   If RS("locked")="1" Then
	     KS.Echo Escape("该访谈已锁定!")
	   ElseIf RS("MessageTF")="0" Then
	     KS.Echo "该访谈不允许网友留言！"
	   ElseIf KS.C("UserName")="" And RS("MessageLoginTF")="1" Then
		  KS.Echo "login"
	   Else
	      KS.Echo "success"
	   End If
  End If 
  RS.Close
  Set RS=Nothing
End Sub

Sub SaveMsg()
   Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select top 1 * From KS_InterView Where ID=" & ID,conn,1,1
   If RS.Eof And RS.Bof Then
      KS.Echo Escape("找不到访谈主题!")
   Else
	   If RS("locked")="1" Then
	     KS.Echo Escape("该访谈已锁定!")
	   ElseIf RS("MessageTF")="0" Then
	     KS.Echo Escape("该访谈不允许网友留言！")
	   ElseIf Now<RS("Begindate") Then
	     KS.Echo escape("该访谈未开始，暂不能留言互动!")
	   ElseIf Cbool(KSUser.UserLoginChecked)=false And RS("MessageLoginTF")="1" Then
		  KS.Echo "login"
	   Else
	      Dim Content:Content=UnEscape(KS.S("Content"))
		  If KS.IsNul(Content) Then
		    KS.Echo Escape("请输入提问内容！")
		  ElseIf KS.IsNul(KS.S("NickName")) Then
		   KS.Echo Escape("请输入您的昵称！")
		  Else
	      Dim RSA:Set RSA=Server.CreateObject("Adodb.recordset")
		  RSA.Open "select top 1 * From KS_InterViewMsg",conn,1,3
		   RSA.AddNew
		   RSA("NickName")=UnEscape(KS.S("NickName"))
		   RSA("UserName")=KSUser.UserName
		   RSA("UserIP")=KS.GetIP
		   RSA("Content")=Content
		   RSA("AddDate")=Now
		   RSA("InterviewID")=ID
		   If RS("MessageVerifyTF")="1" Then
		    RSA("Verify")=0
		   Else
		   RSA("Verify")=1
		   End If
		  RSA.Update
		  RSA.Close
		  Set RSA=Nothing
		  If RS("MessageVerifyTF")="1" Then
		  KS.Echo escape("恭喜，您的提提问已成功提交,审核后才能显示！")
		  Else
		  KS.Echo escape("恭喜，您的提提问已成功提交!")
		  End If
		 End If
	   End If
  End If 
  RS.Close
  Set RS=Nothing
End Sub

'显示网友互动
Sub LoadMsg()
  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "select * from KS_InterViewMsg Where InterviewID=" & id & " and verify=1 Order by id desc",conn,1,1
  Do While Not RS.EOf
    Response.Write Escape("<div><strong>" & rs("nickname") & ":</strong>" & replace(rs("content")&"",chr(10),"<br/>") & " <span style='color:#999' class='date'>" & rs("adddate") & "</span></div>")
  RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
End Sub

'显示文字实录
Sub LoadTextRecord()
 Dim Str,OrderStr
  Str="<table border=""0"" class=""mytable"" style=""margin-top:5px"" width=""99%"" align=""center"" cellspacing=""1"" cellpadding=""1"">" &vbcrlf
  
  OrderStr=" id desc"
  If KS.S("myorder")="0" then OrderStr=" id asc"
  
  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "select * from KS_InterViewRecord Where verify=1 and InterviewID=" & id & " Order by "&OrderStr,conn,1,1
  Do While Not RS.EOf
     if rs("username")="主持人" then
     str=str &"<tr class='zrr'>"
     elseif rs("username")="网友" then
     str=str &"<tr class='wy'>"
	 else
     str=str &"<tr class='jb'>"
	 end if
	 if instr(rs("content")&"","【网友：")<>0 then
		 str=str &"<td height='25' style='word-break;break-all'>" & replace(rs("content")&"",chr(10),"<br/>")
	 else
	     str=str &"<td height='25' style='word-break;break-all'>【" & rs("username") &"】" & replace(rs("content")&"",chr(10),"<br/>")
	 End If
	 str=str & "<span class=""date"">" & rs("adddate") & "</span></td></tr>" &vbcrlf
  RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
  response.write escape(str)
End Sub

%>

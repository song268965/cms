<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="session.asp"-->
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
 case "loadmsg" loadmsg
 case "delmsg" delmsg
 case "verifymsg" verifymsg
 case "savetextrecord" savetextrecord
 case "loadtextrecord" loadtextrecord
 case "deltextrecord" deltextrecord
End Select

Set KS=Nothing
Set KSUser=Nothing


Sub LoadMsg()
  Dim Str
  Str="<table border=""0"" class=""mytable"" width=""99%"" align=""center"" cellspacing=""1"" cellpadding=""1"">" &vbcrlf
  Str=str &"  <tr class='title'><td>网友</td><td>内容</td><td>时间</td><td width=""110"" align=""center"">操作</td></tr>" &vbcrlf
  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "select * from KS_InterViewMsg Where InterviewID=" & id & " Order by id desc",conn,1,1
  If RS.Eof And RS.Bof Then
    str=str &"<tr class='splittd'><td colspan=10 align='center'>没有留言互动记录！</td></tr>"
  Else
	  Do While Not RS.EOf
		 str=str &"<tr class='splittd'><td>【<span id='r"&rs("id")&"'>" & rs("nickname") &"</span>】</td><td id='m" &rs("id") &"'>" & replace(rs("content")&"",chr(10),"<br/>")
		 if rs("verify")="0"then str=str &"&nbsp;<span style='color:red'>未审</span>"
		 str=str & "</td><td style='color:#999'>"
		 if ks.s("editTime2")="1" then
		 str=str &"<input type=""hidden"" name=""ids"" value=""" & rs("id") &"""/>"
		 str=str &"<input type=""text"" onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" class=""textbox myadddate"" name=""adddate" & rs("id") &""" value=""" &  rs("adddate") &"""/>"
		 else
		 str=str & rs("adddate") 
		 end if
		 str=str & "</td><td align='center'><a href='javascript:mselect(" & rs("id") &")'>选择</a> <a href='javascript:delMsg(" & rs("id") &")'>删除</a>"
		 if rs("verify")="1" then
		  str=str &" <a href='javascript:verifyMsg(" & rs("id") &",0)'>取消审核</a>"
		 else
		  str=str &" <a href='javascript:verifyMsg(" & rs("id") &",1)'>审核通过</a>"
		 end if
		 str=str &"</td></tr>" &vbcrlf
	  RS.MoveNext
	  Loop
  End If
  str=str &"</table>"
  RS.Close
  Set RS=Nothing
  response.write escape(str)
End Sub
'文字实录
Sub LoadTextRecord()
 Dim Str
  Str="<table border=""0"" class=""mytable"" style=""margin-top:5px"" width=""99%"" align=""center"" cellspacing=""1"" cellpadding=""1"">" &vbcrlf
  
  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "select * from KS_InterViewRecord Where verify=1 and InterviewID=" & id & " Order by id desc",conn,1,1
  If RS.Eof And RS.Bof Then
    str=str &"<tr class='jb'><td colspan=10 height='30' align='center'>暂没有文字实录！</td></tr>"
  Else
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
		 end if
		 str=str & "</td><td align='center' nowrap style='color:#999'>" 
		 if request("editTime")="1" then
		 
		 str=str &"<input type=""hidden"" name=""ids"" value=""" & rs("id") &"""/>"
		 str=str &"<input type=""text"" onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" class=""textbox myadddate"" name=""adddate" & rs("id") &""" value=""" &  rs("adddate") &"""/>"
		 else
		 str=str & rs("adddate") 
		 end if
		 str=str &  "</td><td align='center' nowrap><a href='javascript:delTextRecord(" & rs("id") &")'>删除</a>"
		 
		 str=str &"</td></tr>" &vbcrlf
	  RS.MoveNext
	  Loop
  End If
  str=str &"</table>"
  RS.Close
  Set RS=Nothing
  response.write escape(str)
End Sub

Sub DelMsg()
 Dim MsgID:MsgID=KS.FilterIds(KS.S("Msgid"))
 If Not KS.IsNul(MsgID) Then
  Conn.Execute("delete from KS_InterViewMsg Where InterViewID=" & ID &" and ID in(" & MsgID &")")
 End If
 KS.Die "success"
End Sub

Sub VerifyMsg()
 Dim MsgID:MsgID=KS.FilterIds(KS.S("Msgid"))
 If Not KS.IsNul(MsgID) Then
  Conn.Execute("update KS_InterViewMsg set verify=" & ks.chkclng(request("v")) & " Where InterViewID=" & ID &" and ID in(" & MsgID &")")
 End If
 KS.Die "success"
End Sub

Sub SaveTextRecord()
 Dim Role:Role=UnEscape(Request("Role"))
 Dim Content:Content=UnEscape(Request("Content"))
 If KS.IsNul(role) Then
  KS.Die escpae("请选择角色!")
 End If
 If KS.IsNul(Content) Then
  KS.Die escpae("请输入实录文字内容!")
 End If
  Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select top 1 * From KS_InterView Where ID=" & ID,conn,1,1
   If RS.Eof And RS.Bof Then
      KS.Echo Escape("找不到访谈主题!")
   Else
	   If RS("locked")="1" Then
	     KS.Echo Escape("该访谈已锁定!")
	   Else
	      
	      Dim RSA:Set RSA=Server.CreateObject("Adodb.recordset")
		  RSA.Open "select top 1 * From KS_InterViewRecord",conn,1,3
		   RSA.AddNew
		   RSA("UserName")=Role
		   IF KS.C("AdminName")<>"" Then
		   RSA("AdminName")=KS.C("AdminName")
		   Else
		   RSA("AdminName")=KS.C("InterViewUserName")
		   End If
		   RSA("UserIP")=KS.GetIP
		   RSA("Content")=Content
		   if isdate(request("adddate")) then
		     RSA("AddDate")=request("adddate")
		   else
		    RSA("AddDate")=Now
		   end if
		   RSA("InterviewID")=ID
		   RSA("Verify")=1
		  RSA.Update
		  RSA.Close
		  Set RSA=Nothing
		  KS.Echo escape("success")
	   End If
  End If 
  RS.Close
  Set RS=Nothing 
End Sub
'删除文字实录
Sub DelTextRecord()
 Dim MsgID:MsgID=KS.FilterIds(KS.S("Msgid"))
 If Not KS.IsNul(MsgID) Then
  Conn.Execute("delete from KS_InterViewRecord Where InterViewID=" & ID &" and ID in(" & MsgID &")")
 End If
 KS.Die "success"
End Sub

%>

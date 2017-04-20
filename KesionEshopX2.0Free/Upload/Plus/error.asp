<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>操作提示消息！</title>
<script src="../ks_inc/jquery.js" type="text/javascript"></script>
<script src="../ks_inc/common.js" language="javascript"></script>
<style type="text/css">
ul,li{list-style-type:none;}
</style>
</head>
<body >
<%
Dim KS:Set KS=New PublicCls
Dim Message:Message = KS.CheckXSS(KS.S("message"))
Message=replace(Message&"","'","\'")
Select Case KS.S("action")
        Case "error"
                Call Error_Msg()
        Case "succeed"
                Call Succeed_Msg()
        Case Else
                Call Error_Msg()
End Select
Set KS=Nothing
Sub Error_Msg()
 If KS.IsNul(Request.ServerVariables("HTTP_REFERER")) Then
%>
 <script>$.dialog.tips('<%= Message%>',2,'error.gif',function(){window.close();});</script>
<%ElseIf instr(lcase(Request.ServerVariables("HTTP_REFERER")),"user/")>0 then%>
 <script>$.dialog.tips('<%= Message%><br/><font color=red>3</font> 秒后自动返回！！！',3,'error.gif',function(){location.href='../';});</script>
<%Else%>
 <script>$.dialog.tips('<%= Message%><br/><font color=red>3</font> 秒后自动返回！！！',3,'error.gif',function(){location.href='<%= Request.ServerVariables("HTTP_REFERER")%>';});</script>
<%End If
End Sub

'********成功提示信息****************
Sub Succeed_Msg()
%>
 <script>$.dialog.tips('<%= Message%><br/><font color=red>3</font> 秒后自动返回！！！',3,'success.gif',function(){location.href='<%= Request.ServerVariables("HTTP_REFERER")%>';});</script>

<%
End Sub

%>
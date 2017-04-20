<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KS
Set KS=New PublicCls
Dim ID
ID = KS.ChkClng(KS.S("ID"))
ChannelID=KS.ChkClng(KS.S("m"))
If ChannelID=0 Then Response.End()
Response.Write "document.write('" & Conn.Execute("Select Score From " & KS.C_S(ChannelID,2) &" Where ID=" & ID)(0) & "');"
Call CloseConn()
Set KS=Nothing
%> 

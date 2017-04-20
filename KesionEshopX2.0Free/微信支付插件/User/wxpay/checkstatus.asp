<!--#include file="../../conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../payfunction.asp"-->

<%
Dim KS:Set KS=New PublicCls
Dim OrderID:OrderID=KS.S("OrderID")
Dim PayFrom:PayFrom=KS.S("PayFrom")
dim ispay:ispay=false
if orderid<>"" then
	dim rs:set rs=server.CreateObject("adodb.recordset")
	rs.open "select top 1 id from KS_LogMoney Where orderid='" & OrderID & "'",conn,1,1
	if not rs.eof  then
	  ispay=true
	end if
	rs.close
	set rs=nothing
end if

if cbool(ispay)=true then
	if payfrom="shop" then
	  ks.die "success|" & KS.Setting(3) & "user/user_order.asp"
	else
	  ks.die "success|" & KS.Setting(3) & "user/user_logmoney.asp"
	end if
else
   ks.die "1"
end if


set ks=nothing
conn.close
set conn=ntohing
%>
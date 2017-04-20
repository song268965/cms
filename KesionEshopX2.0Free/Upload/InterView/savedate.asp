<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="session.asp"-->
<%
Dim KS:Set KS=New Publiccls
dim i,ids,idsarr
if KS.S("Action")="savedate" then
  ids=KS.FilterIDs(request("ids"))
  if ids<>"" then
     idsarr=split(ids,",")
	 for i=0 to ubound(idsarr)
	   if isdate(request("adddate"&idsarr(i))) then
	   conn.execute("update KS_InterViewRecord set adddate='" & request("adddate"&idsarr(i)) & "' where id=" & idsarr(i))
	   end if
	 next
	 ks.die "<script>alert('时间批量修改成功!');top.location.href='main.asp?id="  & request("id") &"';</script>"
  end if
ElseIf KS.S("Action")="savedate2" then
  ids=KS.FilterIDs(request("ids"))
  if ids<>"" then
     idsarr=split(ids,",")
	 for i=0 to ubound(idsarr)
	   if isdate(request("adddate"&idsarr(i))) then
	   conn.execute("update KS_InterViewMsg set adddate='" & request("adddate"&idsarr(i)) & "' where id=" & idsarr(i))
	   end if
	 next
	 ks.die "<script>alert('时间批量修改成功!');top.location.href='main.asp?id="  & request("id") &"';</script>"
  end if
end if

%>

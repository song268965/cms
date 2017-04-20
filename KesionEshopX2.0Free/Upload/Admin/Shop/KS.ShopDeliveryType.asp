<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New PaymentType
KSCls.Kesion()
Set KSCls = Nothing

Class PaymentType
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	    If Not KS.ReturnPowerResult(5, "M520005") Then  Call KS.ReturnErr(1, ""):Exit Sub   
	     Dim RS
		 Dim TypeID:TypeID=2 
         With Response
		   .Write "<!DOCTYPE html><html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
			.Write"</head>"
			.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			.Write "<div class='tabs_header'><ul id='menu_top' class='tabs'>"
			.Write "<li id='p7'><a href='KS.ShopDelivery.asp'><span>送货方式</span></a></li>"
			.Write "<li id='p8'><a href='KS.ShopPaymentType.asp'><span>付款方式</span></a></li>"
			.Write "<li id='p9' class='active'><a href='KS.ShopDeliveryType.asp'><span>快递公司</span></a></li>"
			.Write	" </ul></div>"
		End With
%>		
		<div class="pageCont">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="otable">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>编号</strong></td>
			<td width="217"><strong>快递公司名称</strong></td>
			<td width="197"><strong>排序</strong></td>
			<td width="250"><strong>订单查询api名称</strong></td>
			<td width="197"><strong>是否默认</strong></td>
			<td width="196"><strong>管理操作</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_DeliveryType order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""tdbg"">还没有添加任何的收货方式!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?x=a">
				<tr  class='' onMouseOver="this.className=''" onMouseOut="this.className=''"> 
				  <td width="87" align="center" class="splittd"><%=rs("typeid")%> <input name="typeid" type="hidden" id="typeid" value="<%=rs("typeid")%>"></td>
				  <td width="217" align="center" class="splittd"><input  name="TypeName" type="text" class="textbox" id="TypeName" value="<%=rs("TypeName")%>" size="25"></td>
				  			  
				  <td width="197" align="center" class="splittd"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">
				  </td>
				  <td width="250" align="center" class="splittd"><input style="text-align:center" name="typename_e" type="text" class="textbox" id="typename_e" value="<%=rs("typename_e")%>" size="10">
				  <a href="http://code.google.com/p/kuaidi-api/wiki/Open_API_API_URL" class="aco" target="_blank"><font color="#006699">快递公司参数</font></a>
				  </td>
				  <td width="197" align="center" class="splittd">
				  <a href="?x=d&typeid=<%=rs("typeid")%>">
				  <%If RS("IsDefault") Then
				     Response.Write "<font color=red>是</font>"
					Else
					 Response.Write "否"
					End If
				  %>
				  </a>
				  </td>
				  <td align="center" class="splittd"><input name="Submit" class="button" type="submit"value=" 修改 ">&nbsp;<input  onclick='if (confirm("确定删除吗？")==true){window.location="?x=c&typeid=<%=rs("typeid")%>";}' name="Submit2" type="button" class="button" value=" 删除 "></td>
				</tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?x=b" method="post" name="myform" id="form">
		    <tr class="sort"><td colspan="6" style="text-align: left;">&nbsp;&nbsp;<strong>新增付款方式</strong></td>
		    </tr>
			<tr valign="middle" class="list"> 
			  <td class="splittd"></td>
			  <td class="splittd" align="center"><input name="TypeName" type="text" class="textbox" id="TypeName" size="25"></td>
			  
			  <td class="splittd" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8"></td>
			   <td class="splittd" align="center"><input style="text-align:center" name="typename_e" type="text" value="" class="textbox" id="typename_e" size="10">
			  <a href="http://code.google.com/p/kuaidi-api/wiki/Open_API_API_URL" class="aco" target="_blank"><font color="#006699">快递公司参数</font></a>
			  </td>
			  <td class="splittd" align="center"><input name="isdefault" type="checkbox" value="1" size="8">设为默认
</td>
			  <td class="splittd" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
           
		</form>
</table>
</div>
<div class="footerTable pt10">
		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_DeliveryType set TypeName='" & KS.G("TypeName") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"',typename_e='"& KS.G("typename_e") &"' where Typeid="&KS.G("typeid")&"")
				Response.Redirect "?"
		   case "b"
		       If KS.G("TypeName")="" Then Response.Write "<script>alert('请输入付款方式名称!');history.back();</script>":response.end
				conn.execute("Insert into KS_DeliveryType(TypeName,orderid,typename_e)values('" & KS.G("TypeName") & "','" & KS.ChkClng(KS.G("OrderID")) &"','" & KS.G("typename_e") & "')")
				If KS.G("isdefault")="1" Then
				 Conn.execute("update KS_DeliveryType Set IsDefault=0")
				 Conn.execute("update KS_DeliveryType Set IsDefault=1 Where TypeID=" & Conn.execute("select max(typeid) from KS_DeliveryType")(0))
				End If
				Response.Redirect "?"
		   case "c"
				conn.execute("Delete from KS_DeliveryType where Typeid="&KS.G("typeid")&"")
				Response.Redirect "?"
		   case "d"
				 Conn.execute("update KS_DeliveryType Set IsDefault=0")
				 Conn.execute("update KS_DeliveryType Set IsDefault=1 Where TypeID=" & KS.ChkClng(KS.G("TypeID")))
				Response.Redirect "?"
		End Select
		%></div></body>
		</html>
<%End Sub
End Class
%> 

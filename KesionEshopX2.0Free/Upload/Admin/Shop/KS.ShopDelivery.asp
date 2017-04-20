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
Set KSCls = New Delivery
KSCls.Kesion()
Set KSCls = Nothing

Class Delivery
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		function checkbox(str,char)
		  dim myarray,i
		  myarray=Split(str,",")
			For i = Lbound(myarray) to Ubound(myarray)
			  if myarray(i)=char then
				 checkbox="checked"
			  end if
			next
	   end function
       Sub Kesion()
	     If Not KS.ReturnPowerResult(5, "M520004") Then  Call KS.ReturnErr(1, ""):Exit Sub
	     Dim RS
         With Response
		   .Write "<!DOCTYPE html><html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
			.Write "<style>input{border:solid 1px #A7A7A7}</style>" &vbcrlf
			.Write"</head>"
			.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			if KS.G("Action")="Deliveryapi" or KS.G("Action")="SaveKey"  then
				Call Deliveryapi
				Response.end()
			end if
			.Write "<div class='tabs_header'><ul id='menu_top' class='tabs'>"
			.Write "<li id='p7' class='active'><a href='KS.ShopDelivery.asp'><span>送货方式</span></a></li>"
			.Write "<li id='p8'><a href='KS.ShopPaymentType.asp'><span>付款方式</span></a></li>"
			.Write "<li id='p9'><a href='KS.ShopDeliveryType.asp'><span>快递公司</span></a></li>"
			.Write	" </ul></div>"
		End With
%>		
		<script>
		   var dialogbox=""
			var k_box=""
			//function Deliveryapi(){ 
			//	k_box({s_title:"物流查询API设置",s_width:"600px",s_height:"120px",s_url:"url:KS.Delivery.asp?Action=Deliveryapi"})
			//}
		    function   delright(str,n)   
			{   
			  return   str.substr(0,str.length-n)   
			} 
			function getcheck(id,id2,id3){
			   var value="";
			   for (var i=0;i<id.length;i++ ){
				 if(id[i].checked){ 
				value=value+id[i].value + ",";
			  }
			 } 
			   document.getElementById(id2).value=(value.slice(0,8)+"...");
			   document.getElementById(id3).value=delright(value,1);
			   //alert(document.getElementById(id2).value);
			   if (value==""){
				   alert("没有选择地址，则执行全国统一运费！");
			       document.getElementById(id2).value="全国统一运费";
				   document.getElementById(id2).style.color="red";
			   }else{
				   document.getElementById(id2).style.color="";
			   }
			}
			
        </script> 
		<div class="pageCont">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="otable">
		  <tr align="center"  class="sort">
		    <td><strong>编号</strong></td>
		    <td><strong>快递公司</strong></td>
		    <td><strong>送货地点</strong></td>
		    
		    <td ><strong>首重价格&nbsp;&nbsp;&nbsp;</strong></td>
		    <td><strong>续重价格&nbsp;&nbsp;&nbsp;</strong></td>
		    <td><strong>排序</strong></td>
		    <td><strong>默认</strong></td>
		    <td><strong>默认地区</strong></td>
		    <td width="190"><strong>管理操作</strong></td>
	      </tr>
		  <%dim orderid
		  dim rsp,pxml,pnode,panode,exml,enode
				set rsp=conn.execute("select id,city,parentid from ks_province order by orderid,id")
				if not rsp.eof then
				 Set pxml=KS.RSToXml(rsp,"row","")
				end if
				rsp.close
				'快递公司
				set rsp=conn.execute("select typeid,typename,isdefault from KS_Deliverytype order by orderid")
				if not rsp.eof then
				 set exml=KS.RSToXml(rsp,"row","")
				end if
				rsp.close : set rsp=nothing
		  
		  
		  set rs = conn.execute("select * from KS_Delivery order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""10"" height=""25"" align=""center"" class=""tdbg"">还没有添加任何的送货方式!</td></tr>"
			else
			   
				
			
			   do while not rs.eof
			   dim i,j:j=1
			   %>
		  <form name="form1" method="post" action="?x=a">
		    <tr onmouseover="this.className=''" onmouseout="this.className=''">
		      <td align="center" class="splittd"><%=rs("typeid")%>
		        <input name="ids" type="hidden" id="ids" value="<%=rs("typeid")%>" /></td>
                
		      <td  align="center" class="splittd"><select name="expressid">
		        <%
				if isobject(exml) Then
				   for each enode in exml.documentelement.selectnodes("row")
				  %>
		        <option value="<%=enode.selectsinglenode("@typeid").text%>" <%if trim(rs("expressid"))=trim(enode.selectsinglenode("@typeid").text) then response.Write "selected"%>><%=enode.selectsinglenode("@typename").text%></option>
		        <% next
			   End If
				  %>
		        </select></td>
		      <td align="center" class="splittd"><span style="position:relative">
		        <%dim tocity:tocity=rs("tocity")
				    if KS.IsNul(tocity) then tocity="全国统一运费"
				   %>
		        <input name="prvnc<%=i%>" class="button" style="text-align:center;width:116px;" type="button"  id="btnprvnc<%=i%>" value="<%=left(tocity,8)%>..." onclick="$('span[class=sss]').hide();$('#spprovince<%=i%>').show();if(this.getBoundingClientRect().top>300){spprovince<%=i%>.style.top=(this.offsetHeight-spprovince<%=i%>.offsetHeight)}else{spprovince<%=i%>.style.top='0'}"/>
		        <input name="tocity" id="tocity<%=i%>" value="<%=tocity%>" type="hidden" />
		        <span class="sss" id="spprovince<%=i%>" style="position:absolute;background:#C6E7FA;border:#278BC6 1px solid;width:400px;display:none;height:380px;overflow-y:scroll;overflow-x:hidden;">
		          <table width="100%" height="100%" border="0" style="margin:10px;">
		            <%
					
					dim ischecked
					if isobject(pxml) then
					    for each panode in pxml.documentelement.selectnodes("row[@parentid=0]")
							response.write "<tr><td colspan='10' style='text-align:left'><strong>" & panode.selectsinglenode("@city").text & "</strong></td></tr>"
							 j=0
							 for each pnode in pxml.documentelement.selectnodes("row[@parentid=" & panode.selectsinglenode("@id").text &"]")
								IF (j MOD 4) = 1 THEN response.Write "<tr>"&vbcrlf
								response.Write "<td id=""prvnclist""><input style='border:none' name='prvnc"&i&"'  id='prvnc"&i&"' type='checkbox' value='"&pnode.selectsinglenode("@city").text&"' "&checkbox(tocity,pnode.selectsinglenode("@city").text)&"/> <span id=prvnchtml"&i&">"&pnode.selectsinglenode("@city").text&"</span></td>"&vbcrlf
								if (j mod 4)=0 then response.Write "</tr>"&vbcrlf
								j=j+1
							 next
						next
				  end if
				  %>
	            </table>
				
		          <span style="position:absolute;text-align:right;top:1px;right:6px">
				  <script type="text/javascript">
					function  all_checked(){
						$('td[id=prvnclist]').each(function(){
							$(this).find('input').prop("checked",'true')
						})
					}
					function  del_checked(){
						$('td[id=prvnclist]').each(function(){
							$(this).find('input').prop("checked","")
						})
					}
				</script>
				  <input type="button" class="button" name="allchecked" value="全选"  onclick="all_checked();" />
				  <input type="button" class="button" name="allchecked" value="不选"  onclick="del_checked();" />
				  <input type="button" class="button" value="确定" onclick="spprovince<%=i%>.style.display='none';getcheck(prvnc<%=i%>,'btnprvnc<%=i%>','tocity<%=i%>');"/>
		            <!--<img src="images/close.jpg" onclick="spprovince<%'=i%>.style.display='none';"/>-->
	            </span> </span></span></td>
		      <td align="center" class="splittd"><input class='textbox' name="carriage" type="text"  value="<%=rs("carriage")%>" size="3" />
		        元/
		        <input name="fweight" class='textbox' type="text"  value="<%=rs("fweight")%>" size="3" />
		        kg</td>
		      <td align="center" class="splittd"><input class='textbox' name="cfee"  type="text"  value="<%=rs("C_fee")%>" size="3" />
		        元/
		        <input name="wfee" class='textbox' type="text"  value="<%=rs("W_fee")%>" size="3" />
		        kg</td>
		      <td align="center" class="splittd"><input class='textbox' name="OrderID" type="text"  id="OrderID" value="<%=rs("OrderID")%>" size="3" /></td>
		      <td  align="center" class="splittd"><a href="?x=d&id=<%=rs("typeid")%>">
		        <%If RS("IsDefault")="1" Then
				     Response.Write "<font color=red>是</font>"
					Else
					 Response.Write "否"
					End If
				  %>
		        </a></td>
		      <td  align="center" class="splittd"><input size="8"class='textbox'  type="text" name="defaultcity" id="defaultcity" value="<%=rs("defaultcity")%>"/></td>
		      <td align="center" class="splittd"><input name="Submit" class="button" type="submit"value="修改" />
		        &nbsp;
		        <input  onclick='if (confirm("确定删除吗？")==true){window.location="?x=c&id=<%=rs("typeid")%>";}' name="Submit2" class="button" type="button"  value="删除"/></td>
	        </tr>
	      </form>
		  <%orderid=rs("orderid")
		   i=i+1
		   rs.movenext
		   loop
		 End IF
		rs.close%>
		  <form action="?x=b" method="post" name="myform" id="form">
		    <tr class="sort">
		      <td colspan="9" style="text-align: left;">&nbsp;&nbsp;<strong>新增送货方式</strong></td>
	        </tr>
		    <tr valign="middle">
		      <td class="splittd"></td>
		      <td class="splittd" align="center"><select name="expressid">
				<%
				if isobject(exml) Then
				   for each enode in exml.documentelement.selectnodes("row")
				    dim dtpisdefault:if enode.selectsinglenode("@isdefault").text="1" then dtpisdefault="selected" else dtpisdefault=""
				  %>
		        <option value="<%=enode.selectsinglenode("@typeid").text%>" <%=dtpisdefault%>><%=enode.selectsinglenode("@typename").text%></option>
		        <% next
			   End If
				  %>
				
		        </select></td>
		      <script>
              function getTop(e){   
              var offset=e.offsetTop;   
              if(e.offsetParent!=null) offset+=getTop(e.offsetParent);   
              return offset;   
              }  
              </script>
		      <td align="center" class="splittd"><span style="position:relative">
		        <input name="prvncadd" style="text-align:center;width:112px" class="button" type="button"  id="prvncadd" value="全国统一运费"  onclick="spprovinceadd.style.display='block';if(this.getBoundingClientRect().top>320){spprovinceadd.style.top=(this.offsetHeight-spprovinceadd.offsetHeight)}else{showprovn.style.top='0'}"/>
		        <input id="tocityadd" name="tocityadd" value="" type="hidden" />
		        <span id="spprovinceadd" style="position:absolute;background:#C6E7FA;border:#278BC6 1px solid;width:400px;display:none;height:380px;overflow-y:scroll">
		          <table width="100%" height="100%" border="0" style="margin:10px;">
				  
				  <%
				  if isobject(pxml) then
					    for each panode in pxml.documentelement.selectnodes("row[@parentid=0]")
							response.write "<tr><td colspan='10' style='text-align:left'><strong>" & panode.selectsinglenode("@city").text & "</strong></td></tr>"
							 j=0
							 for each pnode in pxml.documentelement.selectnodes("row[@parentid=" & panode.selectsinglenode("@id").text &"]")
								IF (j MOD 4) = 1 THEN response.Write "<tr>"&vbcrlf
								response.Write "<td><input style='border:none' name='checkprvncadd' id='prvnc"&i&"' type='checkbox' value='"&pnode.selectsinglenode("@city").text&"'/> <span id=prvnchtml"&i&">"&pnode.selectsinglenode("@city").text&"</span></td>"&vbcrlf
								if (j mod 4)=0 then response.Write "</tr>"&vbcrlf
								j=j+1
							 next
						next
				  end if
				  %>
				  
				  
				  
	            </table>
		          <span style="position:absolute;text-align:right;top:2px;right:6px"><input class="button" type="button" value="确定"  onclick="spprovinceadd.style.display='none';getcheck(checkprvncadd,'prvncadd','tocityadd');"/> </span> </span></span></td>
		      <td align="center" class="splittd"><input class='textbox' name="carriage" type="text"  value="10" size="3" />
		        元/
		        <input name="fweight" class='textbox' type="text"  value="2" size="3" />
		        kg</td>
		      <td align="center" class="splittd"><input class='textbox' name="cfee" type="text" value="2" size="3" />
		        元/
		        <input name="wfee" type="text" class='textbox' value="1" size="3" />kg</td>
		      <td align="center" class="splittd"><input  name="orderid" class='textbox' type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="3" /></td>
		      <td align="center" class="splittd"><input name="isdefault" style="border:none" type="checkbox" value="1" size="8" />
		        设为默认 </td>
		      <td align="center" class="splittd"><input class='textbox' name="defaultcity" type="text" value="" class="textbox" id="defaultcity" size="8" /></td>
		      <td align="center" class="splittd"><input name="Submit3" class="button" type="submit" value="OK,提交" /></td>
	        </tr>
	      </form>
	    </table>
		</div>
		<div class="footerTable pt10">
		<div class="attention">
		<font color=red>说明：<br/>
		1、先添加快递公司再添加送货方式;<br/>
		2、同一家快递公司同个地区，请不要重复选择;<br/>
		3、送货地区不选择时，表示该快运公司执行全国统一运费，这样该快递公司只需添加一条数据。</font>
		</div>
<% 
		dim expressid:expressid=KS.ChkClng(KS.G("expressid"))

Select case request("x")
		   case "a"
		   
		   		If Not Isnumeric(KS.G("carriage")) or not isnumeric(KS.G("fweight")) or not isnumeric(KS.G("cfee")) or not isnumeric(KS.G("wfee")) Then Response.Write "<script>alert('运费/首重等必须用数字!');history.back();<//script>":response.end
				
				conn.execute("Update KS_Delivery set defaultcity='" & KS.G("defaultcity") & "',expressid=" & Expressid & ",TypeName='" & KS.G("deliverytype") & "',tocity='" & KS.G("tocity") & "',fweight='" & KS.G("fweight") & "',carriage='" & KS.G("carriage") & "',C_fee='" & KS.G("cfee") & "',W_fee='" & KS.G("wfee") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where typeid="&KS.ChkClng(KS.G("ids")))
				KS.AlertHintScript "恭喜，修改成功！"
		   case "b"

			   If Not Isnumeric(KS.G("carriage")) Then Response.Write "<script>alert('运费必须用数字!');history.back();<//script>":response.end
				conn.execute("Insert into KS_Delivery(expressid,TypeName,tocity,fweight,carriage,C_fee,W_fee,orderid,defaultcity) values('" & expressid & "','" & KS.G("deliverytype") & "','" & KS.G("tocityadd") & "','" & KS.G("fweight") & "','" & KS.G("carriage") & "','" & KS.G("cfee") & "','" & KS.G("wfee") & "','" & KS.ChkClng(KS.G("OrderID")) &"','" & KS.G("defaultcity") & "')")
				If KS.G("isdefault")="1" Then
				 Conn.execute("update KS_Delivery Set IsDefault=0")
				 Conn.execute("update KS_Delivery Set IsDefault=1 Where TypeID=" & Conn.execute("select max(typeid) from KS_Delivery")(0))
				End If
				Response.Redirect "?"
		   case "c"
				conn.execute("Delete from KS_Delivery where typeid="&KS.ChkClng(KS.G("id"))&"")
				Response.Redirect "?"
		   case "d"
				 Conn.execute("update KS_Delivery Set IsDefault=0")
				 Conn.execute("update KS_Delivery Set IsDefault=1 Where TypeID=" & KS.ChkClng(KS.G("ID")))
				Response.Redirect "?"
		End Select
		%></div></body>
		</html>
<%End Sub

sub Deliveryapi()
'	dim API_Key,RS
'	if KS.G("action")="SaveKey" then
'		Conn.Execute("UPDATE KS_Deliverytype SET API_Key = '"& KS.G("API_Key") &"'")
'		KS.echo "<script>alert('修改成功!');top.frames['MainFrame'].dialogbox.close();</"script>"
'		Response.end()
'	else
'		Set RS = Server.CreateObject("AdoDb.RecordSet")
'		RS.Open "select top 1 API_Key from KS_Deliverytype", conn, 1, 1
'		If not RS.Eof And not RS.Bof Then
'			API_Key=rs(0)
'		end if
'		Rs.close
'		Set Rs = Nothing
		if 2=1 then
		%>
		<div style="height:10px;overflow:hidden;"></div>
		<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" >
		 <form action="?action=SaveKey" id="form1" method="post" name="myform">
		<tr class="tdbg" >
			<td height="30" class="clefttitle" align="right"><strong>身份授权key：</strong></td>
			<td>
			 <input type="text" class="textbox" name="API_Key" size="35" value="<%=API_Key%>"> 设为<font color=red>0</font>不启用
			<br /><font color=red>如果还没有快递查询身份授权key，请<a href="http://www.kuaidi100.com/openapi/api_2_02.shtml" target="_blank" >点此申请</a> 。</font>
			</td>
		</tr>
		</form>
		</table>
		<div>
		<ul id='save'>
		<li class='parent' onClick="$('#form1').submit();"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='../images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>
		<li class='parent' onClick="top.frames['MainFrame'].dialogbox.close();"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='../images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>
		</ul>
		</div>
		<%
		end if
'	end if
end sub

End Class

%> 
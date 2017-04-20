<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->

<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Main
KSCls.Kesion()
Set KSCls = Nothing

Class Main
        Private KS,Action,PKID,XMLStr,FieldXML,Node,NodeXML,ItemID,tya,FieldXMLb,NodeXMLb,XML2,TN,NewNode,FieldXMLall,NodeXMLall
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  FieldXML.async = false
		  FieldXML.setProperty "ServerHTTPRequest", true 
		  FieldXML.load(Server.MapPath(KS.Setting(3)&"config/filmk/"& KS.C("AdminName") &"_mk_a.xml"))
		  Set NodeXML=FieldXML.DocumentElement.SelectNodes("item")
		  
		  set FieldXMLb = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  FieldXMLb.async = false
		  FieldXMLb.setProperty "ServerHTTPRequest", true 
		  FieldXMLb.load(Server.MapPath(KS.Setting(3)&"config/filmk/"& KS.C("AdminName") &"_mk_b.xml"))
		  Set NodeXMLb=FieldXMLb.DocumentElement.SelectNodes("item")
		  
		  set FieldXMLall = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  FieldXMLall.async = false
		  FieldXMLall.setProperty "ServerHTTPRequest", true 
		  FieldXMLall.load(Server.MapPath(KS.Setting(3)&"config/filmk/mkall_user.xml"))
		  Set NodeXMLall=FieldXMLall.DocumentElement.SelectNodes("item")
		  
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			'If Not KS.ReturnPowerResult(0, "KSMS20014") Then
			'  Call KS.ReturnErr(1, "")
			'  exit sub
			'End If
			
			PKID=KS.ChkClng(Request("PKID"))
			Action=KS.G("Action")
			Select Case Action
			  Case "add","edit"
			      Call MainList()
			  Case "FastSave"
			      Call FastSave()
			  case "FastSave_all"
			  	call FastSave_all	  
			 Case "del_mk"
			      Call del_mk()
			 case "del_mkb"
			 	  Call del_mkb()	  
			 Case Else
			   Call MainList()
			End Select
	    End Sub
		
		
		sub FastSave()
			 Response.Write("<script src=""../../ks_inc/jquery.js"" type=""text/javascript""></script>")
			 ItemID=KS.ChkClng(Request("id"))

			 if  Request("tya")="add" then
				 Dim ItemID,mm
				 mm=1
				 '取得唯一任务ID号  
				  For Each Node In FieldXMLb.DocumentElement.SelectNodes("item")
					if FieldXMLb.DocumentElement.SelectNodes("item").length=mm  then
						ItemID=KS.ChkClng(Node.SelectSingleNode("@id").text)+1
					end if 
					mm=mm+1
				  Next
				  
				   if not FieldXML.DocumentElement.SelectSingleNode("/field/item[Fast='" & Request("Fast_name") &"']") is nothing then
					   KS.AlertHintScript "您输入的的快捷菜单名称存在，请输入其它的名称！"
					 end if
				   if not FieldXMLall.DocumentElement.SelectSingleNode("/field/item[Fast='" & Request("Fast_name") &"']") is nothing then
					   KS.AlertHintScript "您输入的的快捷菜单名称存在，请输入其它的名称！"
					 end if
				  
				  if FieldXMLb.DocumentElement.SelectNodes("item").length=0 then ItemID=1
				 XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
				 XMLStr=XMLStr&"<item isenable="""& Request("order") &""" id="""& ItemID &""">" &vbcrlf
				 XMLStr=XMLStr&" <Fast>" & Request("Fast_name") &"</Fast>" &vbcrlf
				 XMLStr=XMLStr&" <Fasturl>" & Request("Fast_url") &"</Fasturl>" &vbcrlf
				 XMLStr=XMLStr&" <Attribute>" & Request("Attribute") &"</Attribute>" &vbcrlf
				 XMLStr=XMLStr&" <Fastico>" & Request("ico_url") &"</Fastico>" &vbcrlf
				 XMLStr=XMLStr&" <order>" &Request("order") &"</order>" &vbcrlf
				 XMLStr=XMLStr&"</item>" &vbcrlf
				 XMLStr=Replace(XMLStr, "&", "|Fast|")
				 set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 XML2.LoadXml(XMLStr)
				 set NewNode=XML2.documentElement
				 Set TN=FieldXMLb.DocumentElement
				 TN.appendChild(NewNode)
				 FieldXMLb.Save(Server.MapPath(KS.Setting(3)&"config/filmk/"& KS.C("AdminName")&"_mk_b.xml"))
				 %>
				  <script type="text/javascript">
						location.href='ks.index_mk.asp'
				   </script>
				 <%
			end if
			if  Request("tya")="addmka" then
				 mm=1
				 '取得唯一任务ID号
				
				  
				  For Each Node In FieldXML.DocumentElement.SelectNodes("item")
				  	if FieldXML.DocumentElement.SelectNodes("item").length=mm then
						ItemID=KS.ChkClng(Node.SelectSingleNode("@id").text)+1
					end if 
					mm=mm+1
				  Next
				
				 
				 Set Node=FieldXMLb.DocumentElement.SelectSingleNode("item[@id=" & KS.ChkClng(Request("id")) & "]")
				 If Not Node Is Nothing Then
					XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
					XMLStr=XMLStr&"<item isenable="""& Node.SelectSingleNode("order").text  &""" id="""& ItemID &""">" &vbcrlf
					XMLStr=XMLStr&" <Fast>" & Node.SelectSingleNode("Fast").text  &"</Fast>" &vbcrlf
					XMLStr=XMLStr&" <Fasturl>" & Replace(Node.SelectSingleNode("Fasturl").text,"|Fast|","&") &"</Fasturl>" &vbcrlf
					XMLStr=XMLStr&" <Attribute>" & Node.SelectSingleNode("Attribute").text &"</Attribute>" &vbcrlf
					XMLStr=XMLStr&" <Fastico>" & Node.SelectSingleNode("Fastico").text &"</Fastico>" &vbcrlf
					XMLStr=XMLStr&" <order>" &Node.SelectSingleNode("order").text &"</order>" &vbcrlf
					XMLStr=XMLStr&"</item>" &vbcrlf
				 	XMLStr=Replace(XMLStr, "&", "|Fast|")
					set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					XML2.LoadXml(XMLStr)
					set NewNode=XML2.documentElement
					Set TN=FieldXML.DocumentElement
					TN.appendChild(NewNode)
					FieldXML.Save(Server.MapPath(KS.Setting(3)&"config/filmk/"&KS.C("AdminName")&"_mk_a.xml"))
				 End If	
				 %>
				  <script type="text/javascript">
					$(document).ready(function(){
						var str=""
						str=str+("<div id=\"cpbox\"  class=\"Cp_box\" onMouseOver=\"cpboxbut(this,\'ok\');\"  onmouseout=\"cpboxbut(this,\'no\');\">");
						str=str+("<img id=\"img_<%=ItemID%>\" src=\"<% =Node.SelectSingleNode("Fastico").text %>\"/>");
						str=str+("<div class=\"name\"><a id=\"url_<%=ItemID%>\" href=\"javascript:void(0)\" onclick=\"SelectObjItem1(this,\'<%=Node.SelectSingleNode("Fast").text %> >><font color=red><%=Node.SelectSingleNode("Fast").text%></font>\',\'<%=Node.SelectSingleNode("Attribute").text%>\',\'<%=Replace(Node.SelectSingleNode("Fasturl").text,"|Fast|","&")%>\')\" ><%=Node.SelectSingleNode("Fast").text %></a></div>");
						str=str+("				<span class=\"delico\"><img title=\"删除快捷\" src=\"images/mk_del.png\" onClick=\"delbox(this,\'<%=ItemID%>\')\"/></span>");
						str=str+("				</div>");
						var mm=$(parent.frames["MainFrame"].document).find('.Cp_box').length;
						var newmk=$(parent.frames["MainFrame"].document).find('#cpboxadd').before(str);	
						//alert($(parent.frames["frame2"]).html())
						
						top.frames['MainFrame'].boxi.close();
					});
				   </script>
				 <%
			end if
			
			if  Request("tya")="edit" then
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
				 Node.Attributes.getNamedItem("isenable").text=Request("order")
				 Node.childnodes(0).text=Request("Fast_name")
				 Node.childNodes(1).text=Request("Fast_url")
				 Node.childNodes(2).text=Request("Attribute")
				 Node.childNodes(3).text=Request("ico_url") 
				 Node.childNodes(4).text=Request("order")
				FieldXML.Save(Server.MapPath(KS.Setting(3)&"config/filmk/"&KS.C("AdminName")&"_mk_a.xml"))
				%>
				<script type="text/javascript">
					$(document).ready(function(){
						//top.frames['MainFrame'].location.href='index.asp?action=Main'
					    //alert("---------")
						//var str=""
						
						//var mm=$(parent.frames["MainFrame"].document).find("#cpbox").length
						//alert(mm)
						$(".hiddiv").append(str);
						//$(parent.frames["MainFrame"].document).find("#zzname").val(mytext);
						//$(parent.frames["MainFrame"].document).find("#zzid").val(myid);
						top.frames['MainFrame'].boxi.close();
					});
				</script>
				<%
			end if
		end sub
		
		sub FastSave_all()
			 dim mm
			 Response.Write("<script src=""../../ks_inc/jquery.js"" type=""text/javascript""></script>")
			 ItemID=KS.ChkClng(Request("id"))
			if  Request("tya")="addmka" then
				 mm=1
				 
				 
				 '取得唯一任务ID号
				  For Each Node In FieldXML.DocumentElement.SelectNodes("item")
				  	if FieldXML.DocumentElement.SelectNodes("item").length=mm then
						ItemID=KS.ChkClng(Node.SelectSingleNode("@id").text)+1
					end if 
					mm=mm+1
				  Next
				  
				  
				  

				 Set Node=FieldXMLall.DocumentElement.SelectSingleNode("item[@id=" & KS.ChkClng(Request("id")) & "]")
				 If Not Node Is Nothing Then
				     if not FieldXML.DocumentElement.SelectSingleNode("/field/item[Fast='" & Node.SelectSingleNode("Fast").text &"']") is nothing then
					   KS.AlertHistory "您选择的项目已添加过了！",-1
					 end if
				 
					XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
					XMLStr=XMLStr&"<item isenable="""& Node.SelectSingleNode("order").text  &""" id="""& ItemID &""">" &vbcrlf
					XMLStr=XMLStr&" <Fast>" & Node.SelectSingleNode("Fast").text  &"</Fast>" &vbcrlf
					XMLStr=XMLStr&" <Fasturl>" & Replace(Node.SelectSingleNode("Fasturl").text,"|Fast|","&") &"</Fasturl>" &vbcrlf
					XMLStr=XMLStr&" <Attribute>" & Node.SelectSingleNode("Attribute").text &"</Attribute>" &vbcrlf
					XMLStr=XMLStr&" <Fastico>" & Node.SelectSingleNode("Fastico").text &"</Fastico>" &vbcrlf
					XMLStr=XMLStr&" <order>" &Node.SelectSingleNode("order").text &"</order>" &vbcrlf
					XMLStr=XMLStr&"</item>" &vbcrlf
				 	XMLStr=Replace(XMLStr, "&", "|Fast|")
					set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					XML2.LoadXml(XMLStr)
					set NewNode=XML2.documentElement
					Set TN=FieldXML.DocumentElement
					TN.appendChild(NewNode)
					FieldXML.Save(Server.MapPath(KS.Setting(3)&"config/filmk/"&KS.C("AdminName")&"_mk_a.xml"))
				 End If	
				 %>
				  <script type="text/javascript">
					$(document).ready(function(){
						var str=""
						str=str+("<div id=\"cpbox\"  class=\"Cp_box\" onMouseOver=\"cpboxbut(this,\'ok\');\"  onmouseout=\"cpboxbut(this,\'no\');\">");
						str=str+("				<img id=\"img_<%=ItemID%>\" src=\"<% =Node.SelectSingleNode("Fastico").text %>\"/>");
						str=str+("				<a class=\"name\" id=\"url_<%=ItemID%>\" href=\"javascript:void(0)\"   onclick=\"SelectObjItem1(this,\'<%=Node.SelectSingleNode("Fast").text %> >> <font color=red><%=Node.SelectSingleNode("Fast").text%></font>\',\'<%=Node.SelectSingleNode("Attribute").text%>\',\'<%=Replace(Node.SelectSingleNode("Fasturl").text,"|Fast|","&")%>\')\" ><%=Node.SelectSingleNode("Fast").text %></a>");
						str=str+("				<span class=\"delico\"><img title=\"删除快捷\" src=\"images/mk_del.png\" onClick=\"delbox(this,\'<%=ItemID%>\')\"/></span>");
						str=str+("				</div>");
						$(parent.frames["MainFrame"].document).find('#cpboxadd').before(str);					
						top.frames['MainFrame'].boxi.close();
					});
				   </script>
				 <%
			end if	
		end sub
		
		sub del_mk()
			  ItemID=KS.ChkClng(Request("id"))
			 ' If ItemID=0 Then KS.AlertHintScript "对不起,参数出错!"
			  Dim DelNode,ID
			  Set DelNode=FieldXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
			  If DelNode Is Nothing  Then
			   KS.AlertHintScript "对不起,参数出错!"
			  End If
			  FieldXML.DocumentElement.RemoveChild(DelNode)
			  
			  '更新比当前任务ID大的ID号,依次减一
			 ' For Each Node In FieldXML.DocumentElement.SelectNodes("item")
				' ID=KS.ChkClng(Node.SelectSingleNode("@id").text)
				' If ID>ItemID Then
			'		Node.SelectSingleNode("@id").text=ID-1
			'	 End If
			'  Next
			  '保存
			  FieldXML.Save(Server.MapPath(KS.Setting(3)&"config/filmk/"&KS.C("AdminName")&"_mk_a.xml"))
			  'KS.AlertHintScript "恭喜删除!"
		end sub 
		
		sub del_mkb()
			  ItemID=KS.ChkClng(Request("id"))
			  If ItemID=0 Then KS.AlertHintScript "对不起,参数出错!"
			  Dim DelNode,ID
			  Set DelNode=FieldXMLb.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
			  If DelNode Is Nothing  Then
			   KS.AlertHintScript "对不起,参数出错!"
			  End If
			  FieldXMLb.DocumentElement.RemoveChild(DelNode)
			  FieldXMLb.Save(Server.MapPath(KS.Setting(3)&"config/filmk/"&KS.C("AdminName")&"_mk_b.xml"))
			  %><script type="text/javascript">location.href='ks.index_mk.asp'</script><%
		end sub
		
		Sub MainList()
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			tya="add"
			dim Fast,Fasturl,Attribute,Fastico,Nodek
			if Request("Action")="edit" then
				ItemID=KS.ChkClng(Request("id"))
				If ItemID=0 Then KS.AlertHintScript "对不起,参数出错!"
				Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
				If Not Node Is Nothing Then
					Fast=Node.SelectSingleNode("Fast").text 
					Fasturl=Replace(Node.SelectSingleNode("Fasturl").text,"|Fast|","&")
					Attribute=Node.SelectSingleNode("Attribute").text
					Fastico=Node.SelectSingleNode("Fastico").text
					tya="edit"
				End If	
		   end if
		   if Fastico="" then Fastico="images/mk_1.png"
			With Response
			.Write "<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"" src=""../../ks_inc/jquery.js""></script>"
			.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"" style=""background:#F2F2F2"" >"

			%>
			<ul id='menu_top' style='font-weight:bold;text-align:center;padding-top:14px; background:none; margin-bottom:10px;' >
			<span class="addstyle"><img src="../images/titleka.png"  onMouseOver="addbut(this,'ok');"  onmouseout="addbut(this,'no')" onClick="$('#mkselsec').toggle(200);" style="border:none;cursor:pointer;"/></span>		</ul>
			<script type="text/javascript">
			 function addbut(obj,csk){
			 if (csk=="ok")
			 {
			 $(obj).attr("src","../images/titlekb.png")
			 }
			 else
			 {
			 $(obj).attr("src","../images/titleka.png")
			 }
			 }
			 function add_mk(f_id){
			 	location.href='ks.index_mk.asp?Action=FastSave&tya=addmka&id='+f_id
			 }
			 function add_mk_all(f_id){
			 	location.href='ks.index_mk.asp?Action=FastSave_all&tya=addmka&id='+f_id
			 }
			function CheckForm()
			 {
			   if ($("input[name=Fast_name]").val()=='')
			   {
				 alert('请输入--快捷名称!');
				 $("input[name=Fast_name]").focus();
				 return false
			   }
			   if ($("input[name=Fast_url]").val()=='')
			   {
				 alert('请输入--快捷功能地址!');
				 $("input[name=Fast_url]").focus();
				 return false;
			   }
			   if ($("input[name=ico_url]").val()=='')
			   {
				 alert('请输入--图标地址!');
				 $("input[name=ico_url]").focus();
				 return false;
			   }
			   if ($("input[name=order]").val()=='')$("input[name=order]").val("0");
			   return true;
			 }
			 function del_mkb(id){ 
			 	location.href='ks.index_mk.asp?Action=del_mkb&id='+id
			}
			 function boxclose(){
			 	top.frames["MainFrame"].boxi.close();
			 }
			
			 function delbox(id){
				 top.$.dialog({
					title: '警告',
					width: '110px',
					height: '50px',
					content: '确定要删除快捷!',
					ok: function(){
						del_mkb(id);
						return false
					},
					cancelVal: '取消',
					cancel: true 
				});
			}
			</script>
			<style>
			.mk_list{ padding:0 20px 20px;}
			.mk_list .cp_box{ background:#FFFFFF; padding:5px; margin-bottom:15px;border:1px solid #fff; float:left; margin-right:8px;margin-left:8px; width:98px; height:98px; font-size:13px; font-weight:500; text-align:center; position:relative;border-radius: 5px;}
			
			.mk_list .cp_box .but{background:url(../Images/but_delmkb.png); width:18px; height:18px;cursor:pointer;border:none;}
			.mk_list .cp_box .but1{width:70px; height:25px;cursor:pointer ;border:none;border:1px solid #ddd;margin-top: 5px;}
			</style>
			
			<div id="mkselsec" style="display:none">
			<form name='myform' method='Post' action='KS.index_mk.asp'>
			<table style="margin:0 auto 10px; background:#fff;border-radius: 5px;" width="93%"  border="0" align="center" cellpadding="0" cellspacing="0" >
			<input type="hidden" value="<%=ItemID%>" name="id" id="id">
		    <input type="hidden" value="FastSave" name="action" id="action">
			<input type="hidden" value="<%=tya%>"  name="tya" id="tya">
		    <input type="hidden" value="1" name="v">
			<tr><td colspan="2" class="pt10"></td></tr>
			<tr>
			<td width="20%" height="50" align="center" >快捷名称:</td>
			<td width="80%" style="padding-left:5px"><input class="textbox" name="Fast_name" type="text" value="<%=Fast%>" size="10"  /> </td>
			</tr>
			<tr>
			<td width="20%" height="50" align="center" >快捷功能地址:</td>
			<td width="80%" style="padding-left:5px">
			<input name="Fast_url" class="textbox" type="text" value="<%=Fasturl%>"  size="30"/>
			属性:
			<select name="Attribute">
			<%if Attribute<>"" then 
				 dim sel_1,sel_2,sel_3
				 select case Attribute
					case "disabled"
					sel_1="selected='selected'"
					case "GO"
					sel_2="selected='selected'"
					case "SetParam"
					sel_3="selected='selected'"
				 end select	
			else
				sel_1="selected='selected'"
			end if
			%>
			
			  <option value="disabled" <%=sel_1%>>管理查看</option>
			  <option value="GO" <%=sel_2%>>新建增加</option>
			  <option value="SetParam" <%=sel_3%>>修改保存</option>
			  
			</select>
			</td>
			</tr>
			<tr>
			<td width="20%" height="50" align="center"> 图标选择:<br/><font color="#FF0000">图标40X40</font></td>
			<td width="80%" style="padding-left:5px">
			<input name="ico_url" type="text"  class="textbox" value="<%=Fastico%>" size="30" />
			<img id="img_ico" src="../<%=Fastico%>" style='vertical-align: middle;height: 20px;'/>
			<input name="icobut" value="选择图标.." onclick="$('#Fastico').toggle(200);" type="button" class='button'/>
				<div id="Fastico" style="display:none; padding:5px; overflow:hidden;" >
				
				<%for i=1 to 11 %>
				
				<a href="#" style="display:block;float:left;text-align:center;margin:2px;border:1px solid #ddd; width:45px; height:45px;" onclick="$('input[name=ico_url]').val('images/mk_<%=i%>.png');$('#img_ico').attr('src','../images/mk_<%=i%>.png');$('#Fastico').toggle(200);" ><img src="../images/mk_<%=i%>.png" style="border:none" width="40" height="40" /></a>
				<%next%>
				</div>
			
			</td>
			</tr>
			<tr style="display:none">
			<td width="30%" height="30" bgcolor="#E6F3FB" align="center" style="border-bottom:1px solid #ddd; display:none">序号排列:</td>
			<td width="70%" bgcolor="#FFFFFF" style="border-bottom:1px solid #CCCCCC;padding-left:5px;display:none">
			<input name="order" size="5" class="textbox" type="text" value="1" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /> 
			</td>
			</tr>
            <tr><td colspan="2" class="pd10"></td></tr>
			</table>
			<div  align="center" class='md10'>
			<%if Request("Action")="add" then%>
			<input name="" value="添加" type="submit" onClick="return(CheckForm())" class='button'/>
			<%else%>
			<input name="" value="保存" type="submit" onClick="return(CheckForm())" class='button'/>
			<%end if%>
			<%if Request("Action")="add" then%>
			<input name="" value="返回"  type="button" class='button' onclick="$('#mkselsec').toggle(200);"  />
			<%else%>
			<input name="" value="返回"  type="button" class='button'  onclick="$('#mkselsec').toggle(200);"  />
			<%end if%>
			</div>
			</form>
			</div>
			<script src="../images/jquery.nicescroll.js"></script>
			<script>
			
				$(function(){
					
					$("body").niceScroll({  
					cursorcolor:"#a5a8a9",  
					cursoropacitymax:1,  
					touchbehavior:false,  
					cursorwidth:"5px",  
					cursorborder:"0",  
					cursorborderradius:"15px"  
					}); 
				});
				
				
			</script>

			<%
			Response.Write "<div id=""cpbox""  class=""mk_list clearfix"" >"
			dim Nodekall
			For Each Nodekall In NodeXMLall
			Response.Write "<div id=""cpbox""  class=""cp_box"" >"
			%>
			<img id="img_<%=Nodekall.SelectSingleNode("@id").text%>" src="../<% =Nodekall.SelectSingleNode("Fastico").text %>"/> <br/>
			 <span><%=Nodekall.SelectSingleNode("Fast").text%></span><br />
			 <input name="" class="but1 button2"  value="+添加" type="button" onclick="add_mk_all('<%=Nodekall.SelectSingleNode("@id").text%>');" /> 
			<%	
			Response.Write "</div>"
			Next
			
			
			For Each Nodek In NodeXMLb
			Response.Write "<div id=""cpbox""  class=""cp_box"" >"
			%>
			<img id="img_<%=Nodek.SelectSingleNode("@id").text%>" src="../<% =Nodek.SelectSingleNode("Fastico").text %>"/> <br/>
			 <span><%=Nodek.SelectSingleNode("Fast").text%></span><br />
			 <input name="" class="but1"  value="" type="button" onclick="add_mk('<%=Nodek.SelectSingleNode("@id").text%>');" /> 
			 <span style="position:absolute; top:-6px; right:-6px;" >
			 <input name="dfgdfg" class="but" type="button" onClick="if(confirm('是否删除!')){del_mkb('<%=Nodek.SelectSingleNode("@id").text%>')};"/>
			 </span>
			<%	
			Response.Write "</div>"
			Next
			Response.Write "</div>"
			%>
			<div style="height:20px;clear:both; overflow:hidden;clear:both;"></div>
			</body>
			</html>
			<%
			End With
			End Sub

End Class
%>
 

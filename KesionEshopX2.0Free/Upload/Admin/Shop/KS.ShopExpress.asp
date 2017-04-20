<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_EnterPrise
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrise
        Private KS,Param,KSCls
		Private Action,i,strClass,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  maxperpage = 30 '###每页显示数
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()

		
		 With Response
			  If Not KS.ReturnPowerResult(5, "M520014") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
			  End If
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>添加快递单模板</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加快递单模板</span></li>"
			 

			  .Write "</ul>"
		End With

		Select Case KS.G("action")
		 Case "Add","Edit" Call ExpressManage()
		 Case "dosave" Call DoSave()
		 Case "dosaveinfo" Call dosaveinfo()
		 Case "ExpressType"  Call ExpressType()
		 Case "CancelExpressType" Call CancelExpressType()
		 Case "Del"  Call ExpressDel()
		 Case Else
		  Call showmain
		End Select
		
		
End Sub

Private Sub showmain()
  CurrentPage=KS.ChkClng(Request("page")):If CurrentPage<0 Then CurrentPage=1
   Param=" where 1=1"
%>
<div class="pageCont2">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>快递单模板名称</td>
	<td nowrap>是否启用</td>
	<td nowrap>管理操作</td>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_ShopExpress " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>对不起,找不到符合条件的快递单模板！</td></tr>"
	Else
		totalPut = RS.RecordCount
		If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
				RS.Move (CurrentPage - 1) * MaxPerPage
		End If
		i = 0
%>
<form name=selform method=post action=?action=Del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td  class="splittd"><%=Rs("Title")%>
	</td>
	<td align="center" class="splittd"><%
	    if rs("status")=1 then
		  response.write "<font color=blue>正常</font>"
		else
		  response.write " <font color=red>禁用</font>"
		end if
	%></td>
	<td align="center" class="splittd">
	 
	<a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>修改快递单模板</font>')+'&ButtonSymbol=GOSave';" class='setA'>修改</a>|<a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('删除快递单模板不可恢复，确定删除吗？'));" class='setA'>删除</a>|
		
		<%IF rs("Status")="1" then %><a href="?Action=CancelExpressType&id=<%=rs("id")%>" class='setA'>关闭</a><%else%><a href="?Action=ExpressType&id=<%=rs("id")%>" class='setA'>启用</a><%end if%>

	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='pt10' onMouseOver="this.className='pt10'" onMouseOut="this.className='pt10'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选&nbsp;&nbsp;
	<input class="button" type="submit" name="Submit2" value="删除选中的记录" onclick="{if(confirm('删除快递单模板不可恢复，确定删除吗？')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
<br/>

<%dim Setting,rss
set rss=server.CreateObject("adodb.recordset")
rss.open "select shopsetting from ks_config",conn,1,1
Setting=split(rss(0) & "^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^^#^","^#^")
rss.close
set rss=nothing
%>
   <div class="attention" style=" background:#fff;font-weight:bold;font-size:14px">配置发货信息：
   <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="otable">
  <form name="myform" id="myform" action="KS.ShopExpress.asp" method="post">
    <input type="hidden" name="action" value="dosaveinfo"/>
    <tr>
      <td height='26' class='clefttitle'><strong>公司名称：</strong></td><td><input class="textbox" type='text' name='Setting(0)' id='CompanyName' value='<%=Setting(0)%>' size="30" /></td>
      <td height='26' class='clefttitle'><strong>发货人姓名：</strong></td><td><input class="textbox" type='text' name='Setting(1)' id='RealName' value='<%=Setting(1)%>' size="30" /></td>
	</tr> 
    <tr>
      <td height='26' class='clefttitle'><strong>发货详细地址：</strong></td><td><input class="textbox" type='text' name='Setting(2)'  value='<%=Setting(2)%>' size="30" /></td>
      <td height='26' class='clefttitle'><strong>发货地邮编：</strong></td><td><input class="textbox" type='text' name='Setting(3)' value='<%=Setting(3)%>' size="30" /></td>
	</tr> 
    <tr>
      <td height='26' class='clefttitle'><strong>发货人手机：</strong></td><td><input class="textbox" type='text' name='Setting(4)'  value='<%=Setting(4)%>' size="30" /></td>
      <td height='26' class='clefttitle'><strong>发货人电话：</strong></td><td><input class="textbox" type='text' name='Setting(5)' value='<%=Setting(5)%>' size="30" /></td>
	</tr> 
    <tr>
      <td height='26' class='clefttitle'><strong>始发地：</strong></td><td><input class="textbox" type='text' name='Setting(6)'  value='<%=Setting(6)%>' size="30" /></td>
      <td height='26' class='clefttitle'></td><td></td>
	</tr> 
    <tr>
      <td height='26' class='cefttitle' colspan="4" style="text-align:center"><input type="submit" value=" 保存 " class="button"/></td>
	</tr> 
  </form>
  </table>
</div>

<div class="attention">
  <b>注意事项：</b>
  <li>设置打印机的尺寸，开始-&gt;打印机和传真-&gt;右击 服务器属性-&gt;创建新格式-&gt;填写快递单尺寸（一般大小为：23cm*12.7cm）</li>
  <li>打印机后进纸的时候，纸张一定靠左，以左对齐，然后再对齐右边。这样不会打歪</li>
  <li>扫描好的快递单图片大小应该改成874*483</li>
  <li>把浏览器的页面设置量的 上和下改成0,页眉页脚都要设置为空（非常重要）</</li>
  <li>要开始打印时需给打印机设置下纸张大小：选择打印机-&gt;打印首选项-&gt;高级-&gt;选择纸张规格，选择刚第一步添加的纸规格即可</li>
</div>
</div>
<%
End Sub

Sub dosaveinfo()
 dim n,setstr
 for n=0 to 10
  if n=0 then
   setstr=request("setting(" & n &")")
  else
   setstr=setstr & "^#^" &request("setting(" & n &")")
  end if
 next
 dim rs:set rs=server.CreateObject("adodb.recordset")
 rs.open "select * from ks_config",conn,1,3
 rs("shopsetting")=setstr
 rs.update
 rs.close
 set rs=nothing
 ks.alerthintscript "恭喜，发货信息保存成功!"
End Sub

Sub ExpressManage()
Dim Title,Status,Template,PhotoUrl,ExpressID
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select top 1 * From KS_ShopExpress Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	  Response.End
	 Else
	   Title=RS("Title")
	   Status=RS("Status")
	   Template=RS("Template")
	   PhotoUrl=RS("PhotoUrl")
	   ExpressID=RS("ExpressID")
	 End If
	  rs.close

Else
  PhotoUrl="zt.jpg" : status=1 : ExpressID=0
 End If
%>
<script>
function CheckForm()
{
   $("#template").val($("#mybody").html());
	if ($('#Title').val()=='')
	{
	 alert('请输入快递模板名称!');
	 $("#Express").focus();
	 return false;
	}
   if ($("#template").val()==''){
	 alert('模板内容必须放入标签!');
	 $("#template").focus();
	 return false;
   }

document.myform.submit();
}

var rDrag = {
    o:null,
    init:function(o){
    o.onmousedown = this.start;
    },
    start:function(e){
    var o;
    e = rDrag.fixEvent(e);
    e.preventDefault && e.preventDefault();
    rDrag.o = o = this;
    o.x = e.clientX - rDrag.o.offsetLeft;
    o.y = e.clientY - rDrag.o.offsetTop;
    document.onmousemove = rDrag.move;
    document.onmouseup = rDrag.end;
    },
    move:function(e){
    e = rDrag.fixEvent(e);
    var oLeft,oTop;
    oLeft = e.clientX - rDrag.o.x;
    oTop = e.clientY - rDrag.o.y;
    rDrag.o.style.left = oLeft + 'px';
    rDrag.o.style.top = oTop + 'px';
    },
    end:function(e){
    e = rDrag.fixEvent(e);
    rDrag.o = document.onmousemove = document.onmouseup = null;
    },
    fixEvent: function(e){
    if (!e) {
    e = window.event;
    e.target = e.srcElement;
    e.layerX = e.offsetX;
    e.layerY = e.offsetY;
    }
    return e;
    }
    }

var domid=1;	
function add_div()
{   var label=$("#mylabel").val();
   if (label==''){alert('请选择选标签!');return;}
    var o=document.createElement("label");
    o.className="mo";
    o.id="m"+domid;
    $("#mybody").append(o);
    o.innerHTML=label;
    domid++;
	rDrag.init(o);
	
	$("#mybody").find(".mo").each(function(){ //双击删除标签
	   $(this).dblclick(function(){
	    if (confirm('确定删除该标签吗？')){
		 $(this).remove();
		 }
	   });
    });
	
}
<%If id<>0 then%>
$(function(){
 $("#mybody").find("label").each(function(){
   rDrag.init($(this)[0]);
   $(this).dblclick(function(){
	    if (confirm('确定删除该标签吗？')){
		 $(this).remove();
		 }
   });
 });
});
<%end if%>

function changebg(v){
 if (v=='') return;
 $("#mybody").attr("style","background:url(../../shop/express/"+v+")");
}
    </script>

    <style type="text/css">
	.mo{
	 display:block;
	border:1px solid #ff6600;
    padding:0px;
    height:22px;
	font-size:14px;
	line-height:22px;
    position:absolute;
	}
	
	 .box{position:relative;border:1px solid #ccc;border-top:2px solid #000;width:874px;height:483px;background:url(../../shop/express/<%=photourl%>) no-repeat;}
	 .noprint{display:naone;}
	 @media print {     
            .noprint{     
              display: none;    
        }     
    }     

    </style>
<div class="pageCont2">
	<div class="noprint">

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="otable">
  <form name="myform" id="myform" action="KS.ShopExpress.asp" method="post">
    <input type="hidden" name="action" value="dosave"/>
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
    <tr>
      <td height='26' class='pt10' style="text-align:left"><strong>模板名称：</strong><input class="textbox" type='text' name='Title' id='Title' value='<%=Title%>' size="10" />&nbsp;&nbsp;<strong>状态</strong> <input type="checkbox" name="Status" value="1"<%if Status=1 then response.write " checked"%> />正常
       &nbsp;&nbsp;<strong>背景图片：</strong><select name="photourl" onchange="changebg(this.value);">
		<option value=''>请选择...</option>
		<%dim fs:Set fs = KS.InitialObject(KS.Setting(99)) 
		  dim f:set f = fs.GetFolder(server.MapPath("../../shop/express/")) 
		  dim fc:Set fc = f.Files 
		  dim i,f1
					For Each f1 in fc 
					 if instr(lcase(f1.name),".jpg")<>0 or instr(lcase(f1.name),".gif")<>0 or instr(lcase(f1.name),".png")<>0  then
					  if photourl=f1.name then
					   response.write "<option value='" & f1.name & "' selected>" & f1.name & "</option>"
					  else
					   response.write "<option value='" & f1.name & "'>" & f1.name & "</option>"
					  end if
					 end if
		Next 
	  %>
		</select>
		&nbsp;&nbsp;
		绑定快递公司：<select name="expressid">
		 <option value='0'>请选择...</option>
		 <%
		  rs.open "select * from KS_Deliverytype order by orderid,typeid",conn,1,1
		  do while not rs.eof
		    IF ExpressID=RS("Typeid") then
		     response.write "<option value='" & rs("typeid") & "' selected>" & rs("typename") & "</option>"
			Else
		     response.write "<option value='" & rs("typeid") & "'>" & rs("typename") & "</option>"
		   End If
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		 %>
		</select>
	  </td>
      
    </tr>
<tr>
      <td height='26' class='pt10' style="text-align:left">
		  <strong>插入标签：</strong><select name="mylabel" id="mylabel">
		   <option value=''>请选择标签...</option>
		   <option value='{$寄件人_姓名}'>{$寄件人_姓名}</option>
		   <option value='{$寄件人_地址}'>{$寄件人_地址}</option>
		   <option value='{$寄件人_电话}'>{$寄件人_电话}</option>
		   <option value='{$寄件人_手机}'>{$寄件人_手机}</option>
		   <option value='{$寄件人_邮编}'>{$寄件人_邮编}</option>
		   <option value='{$寄件人_始发地}'>{$寄件人_始发地}</option>
		   <option value='{$寄件人_单位}'>{$寄件人_单位}</option>
		   
		   <option value='{$收件人_姓名}'>{$收件人_姓名}</option>
		   <option value='{$收件人_地址}'>{$收件人_地址}</option>
		   <option value='{$收件人_电话}'>{$收件人_电话}</option>
		   <option value='{$收件人_手机}'>{$收件人_手机}</option>
		   <option value='{$收件人_邮编}'>{$收件人_邮编}</option>
		   <option value='{$收件人_目的地}'>{$收件人_目地的}</option>
		   
		   <option value='{$年}'>{$当前日期_年}</option>
		   <option value='{$月}'>{$当前日期_月}</option>
		   <option value='{$日}'>{$当前日期_日}</option>
		   <option value='{$订单_备注留言}'>{$订单_备注留言}</option>
		   <option value='{$订单_总金额}'>{$订单_总金额}</option>
		   
		   <option value='√'>{$打勾_√}</option>
		  </select>
		  <textarea name="template" style="display:none" id="template"></textarea>
		  <input type='button' class="button" value=" 插入标签 " onclick="add_div();"/>
		  
		  <input type="button" class="button" value=" 保存模板 " onclick="CheckForm();"/>
		  
		   <span class="tips">请将扫描好的快递单图片放到/shop/express/目录下，规格：874*483</span>
   </td>
  </tr>
  </form>
</table>
	</div>

   <div id="mybody" class="box mt10"><%=template%></div>
   
   <br/>
   <div class="attention">
     提示，鼠标点中标签可以拖动，双击标签可以删除。
   </div>
</div>
<%
End Sub


Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Title:Title=KS.LoseHtml(KS.G("Title"))
       Dim photourl:photourl=KS.G("photourl") 
	   Dim Status:Status=KS.ChkClng(KS.G("Status"))
	   Dim Template:Template=Request.Form("Template")
	   Dim ExpressID:ExpressID=KS.ChkClng(KS.G("ExpressID"))
	  
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   If Title="" Then Response.Write "<script>alert('快递模板名称必须输入');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_ShopExpress Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
			  End If
				 RS("Title")=Title
				 RS("PhotoUrl")=PhotoUrl
			     RS("Status")=Status
				 RS("Template")=Template
				 RS("ExpressID")=ExpressID
		 		 RS.Update
			     RS.Close
				 Set RS=Nothing
				 If ID=0 Then
				  Response.Write "<script>if (confirm('快递单模板添加成功!')){location.href='?action=Add';}else{$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>快递单模板管理</font>") & "';location.href='KS.ShopExpress.asp';}</script>"
				 Else
				  Response.Write "<script>alert('快递单模板修改成功！');$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>快递单模板管理</font>") & "';location.href='"& ComeUrl & "';</script>"
				 End If

EnD Sub

'删除
Sub ExpressDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Delete From KS_ShopExpress Where id In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub ExpressType()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_ShopExpress Set Status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub CancelExpressType()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_ShopExpress Set Status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
End Class
%> 

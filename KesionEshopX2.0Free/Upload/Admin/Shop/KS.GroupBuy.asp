<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_GroupBuy
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_GroupBuy
        Private KS,Param,KSCls
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
		      .Write "<!DOCTYPE html><html>"&vbcrlf
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>"&vbcrlf
			  .Write EchoUeditorHead()
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div id='menu_top'>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('团购系统 >> <font color=red>添加团购商品</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加团购</span></li>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('团购系统 >> <font color=red>团购分类管理</font>')+'&ButtonSymbol=Disabled';location.href='?action=ClassManage';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon move'></i>分类管理</span></li>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('团购系统 >> <font color=red>添加团购分类</font>')+'&ButtonSymbol=GOSave';location.href='?action=AddClass';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon audit'></i>添加分类</span></li>"
			 
			  .Write "<li class='changeWay' style='margin-top: 5px;'>&nbsp;&nbsp;<strong>查看方式:</strong><a href=""KS.GroupBuy.asp"">所有团购</a> <a href=""KS.GroupBuy.asp?flag=1"">进行中的团购</a>  <a href=""KS.GroupBuy.asp?flag=2"">已结束的团购</a> <a href=""KS.GroupBuy.asp?flag=3"">已锁定的团购</a></li>"

			  .Write "</div>"
		End With
		
		   	 If Not KS.ReturnPowerResult(5, "M530001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
			 End If

		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		
		Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and Subject like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and Intro like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		if KS.G("Verific")<>"" then
		  If KS.G("Verific")="1" Then Param=Param & " and Verific=1"
		  If KS.G("Verific")="0" Then Param=Param & " and Verific=0"
		  If KS.G("Verific")="3" Then Param=Param & " and Verific=3"
		end if
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="1" Then Param=Param & " and locked=0 and endtf=0"
		  If KS.G("Flag")="2" Then Param=Param & " and endtf=1"
		  If KS.G("Flag")="3" Then Param=Param & " and locked=1"
		  
		End If

		totalPut = Conn.Execute("Select Count(id) From KS_GroupBuy " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call SubjectManage()
		 Case "EditSave" Call DoSave()
		 Case "Verific" Call GroupBuyVerific()
		 Case "Verific_t" Call GroupBuyVerific_t()
		 
		 Case "Del"  Call GroupBuyDel()
		 Case "Recommend" Call Recommend()
		 Case "UnRecommend" Call UnRecommend()
		 Case "lock"  Call GroupBuyLock()
		 Case "unlock" Call GroupBuyUnLock()
		 Case "endtf"  Call GroupBuyendtf()
		 Case "Cancelendtf" Call GroupBuyCancelendtf()
		 Case "AddClass" Call AddClass()
		 Case "AddClassSave" Call AddClassSave()
		 Case "ClassManage" Call ClassManage()
		 Case "DelClass" Call DelClass()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<script src="../../ks_inc/jquery.imagePreview.1.0.js"></script>
<script type="text/javascript">
function ShowSale(id,title)
 { top.openWin("查看商品销售详情","shop/KS.ShopProSale.asp?proid="+id+"&title="+escape(title),false)}
</script>
<div class="tableTop">
<table><tr><td>
<form action="KS.GroupBuy.asp" name="myform" method="get">
   <div>
      <strong class="mr0">快速搜索=></strong>
	 <span class="tiaoJian">关键字:</span><input type="text" class='textbox' name="keyword"><span class="tiaoJian">条件:</span>
	 <select name="condition" class="h30">
	  <option value=1>按商品名称</option>
	  <option value=2>按商品介绍</option>
	 </select>
	  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
    </div>
</form></td></tr>
</table>
</div>
<div class="tabs_header mt20">
    <ul class="tabs">
    <li<%if KS.S("Verific")="" then response.write " class='active'"%>><a href="KS.GroupBuy.asp"><span>所有团购信息</span></a></li>
    <li<%if KS.S("Verific")="0" then response.write " class='active'"%>><a href="KS.GroupBuy.asp?Verific=0"><span>待审核的团购信息(<label style="color:red"><%=Conn.Execute("select count(1) From KS_GroupBuy Where Verific=0")(0)%></label>)</span></a></li>
    <li<%if KS.S("Verific")="1" then response.write " class='active'"%>><a href="KS.GroupBuy.asp?Verific=1"><span>审核过的团购信息</span></a></li>
    <li<%if KS.S("Verific")="3" then response.write " class='active'"%>><a href="KS.GroupBuy.asp?Verific=3"><span>被退回的团购信息</span></a></li>
    </ul>
</div>
<div class="pageCont">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="30" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>团购主题</td>
	<td nowrap>开始/结束时间</td>
	<td nowrap>原价</td>
	<td nowrap>现价</td>
	<td nowrap>团购状态</td>
	<td nowrap>管理操作</td>
</tr>
<%
	sFileName = "KS.GroupBuy.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_GroupBuy " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""30"" align=center class='splittd' colspan='7'>还没有添加任何团购信息！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="KS.GroupBuy.asp">
<input type="hidden" name="action" id="action" value=""/>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="30"  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd">
	<a href='<%=rs("bigphoto")%>' onclick='window.open("../../shop/groupbuyshow.asp?id=<%=rs("id")%>");return false;'><img onerror="this.src='../../images/nopic.gif';" style='margin:2px;padding:1px;border:1px solid #ccc' src='<%=rs("photourl")%>' title="点击预览" border='0' width='40' height='40' align='left'/></a>
	<span style="cursor:default;margin-top:3px;display:block;">
	<a href='../../shop/groupbuyshow.asp?id=<%=rs("id")%>' target='_blank' style='font-size:13px;'><%=KS.Gottopic(Rs("Subject"),35)%></a>
	<%If rs("recommend")="1" then response.write " <font color=green>荐</font>"%>
	<div class="tips" style="margin-top:4px">[总销量：<font color=blue><%=KS.ChkClng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.proid=" & rs("id"))(0))%></font> 件，已付：<font color=green><%=ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and o.MoneyReceipt>0 and i.proid=" & rs("id"))(0))%> </font>件，未付：<font color=red><%=ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and o.MoneyReceipt<=0 and i.proid=" & rs("id"))(0))%></font> 件]</div>
	</span>
		</td>
	<td align="center" width="120" class="splittd">
	<%=Rs("adddate")%><br/> 至<br/> <%=Rs("ActiveDate")%>
	</td>
	
	<td align="center" class="splittd">
	<span style='color:#999999;text-decoration:line-through;'><%=rs("price_original")%> 元</span>
	</td>
	<td align="center" class="splittd">
	<span style='color:#ff6600'><%=rs("price")%> 元</span>
	</td>
	<td align="center" class="splittd">
	<%
	if DateDiff("s",now,RS("AddDate"))>0 Then
		response.write " <font color=green>未开始</font>"
	ElseIf DateDiff("s",now,RS("ActiveDate"))<0 Then
		response.write " <font color=#cccccc>已结束</font>"
	elseif rs("locked")=0 and rs("endtf")=0 then
	 response.write "<font color=red>进行中</font>"
	else
		if rs("locked")=1 then
		  response.write "<font color=blue>锁定</font>"
		end if
		if rs("endtf")=1 then
		  response.write " <font color=#cccccc>已结束</font>"
		end if
	end if
	%></td>
	<td align="center" class="splittd">
    <%if KS.S("Verific")="0" or KS.S("Verific")="3"  then%>
    	<a href="?Action=Verific&id=<%=rs("id")%>" class="setA">审核</a>|
   		<%if  KS.S("Verific")<>"3"  then%>
    	<a href="?Action=Verific_t&id=<%=rs("id")%>" class="setA" class="setA">退团购</a>|
    	<%end if%>
   		<a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('团购系统 >> <font color=red>修改团购信息</font>')+'&ButtonSymbol=GOSave';" class="setA">修改</a>|
    <%else%>
    <a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('团购系统 >> <font color=red>修改团购信息</font>')+'&ButtonSymbol=GOSave';" class="setA">修改</a>| 
    <a href="?Action=Del&ID=<%=rs("id")%>" onClick="return(confirm('确定删除该团购吗？'));"  class="setA">删除</a>| 
	
	&nbsp;<%if rs("locked")=0 then%><a href="?Action=lock&id=<%=rs("id")%>" class="setA">锁定</a>|<%else%><a href="?Action=unlock&id=<%=rs("id")%>">解锁</a>|<%end if%>
		
		&nbsp;<%IF rs("endtf")="1" then %><a href="?Action=Cancelendtf&id=<%=rs("id")%>" class="setA"><font color=red>打开</font></a>|<%else%><a href="?Action=endtf&id=<%=rs("id")%>"  class="setA">结束</a>|<%end if%>
		
		<a href="javascript:ShowSale(<%=rs("id")%>,'<%=rs("subject")%>');" class="setA">销售</a>
    <%end if%>
    
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
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的团购 " onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){$('#action').val('Del');this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" name="Submit2" value=" 批量推荐 " onClick="$('#action').val('Recommend');this.document.selform.submit();return true;">
	<input class="button" type="submit" name="Submit2" value=" 批量取消推荐 " onClick="$('#action').val('UnRecommend');this.document.selform.submit();return true;">
	
	
	
	</td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
</div>
<div class="footerTable pt10">

</div>

<%
End Sub

Sub SubjectManage()
Dim Subject,ActiveDate,AddDate,Intro,Highlights,Protection,Notes,Locked,EndTF,PhotoUrl,BigPhoto,ClassID,AllowBMFlag,AllowArrGroupID,minnum,Comment,Changes,ChangesUrl
Dim Price_Original,Price,Discount,limitbuynum,weight,recommend,ProvinceID,CityID,HasBuyNum,MustPayOnline,CleanCart,showdelivery
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select top 1 * From KS_GroupBuy Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	  Response.End
	 Else
	   Subject=RS("Subject")
	   Price_Original=RS("Price_Original")
	   Price=RS("Price")
	   Discount=RS("Discount")
	   ActiveDate=RS("ActiveDate")
	   AddDate=RS("AddDate")
	   Intro=RS("Intro")
	   PhotoUrl=RS("PhotoUrl")
	   BigPhoto=RS("BigPhoto")
	   Highlights=RS("Highlights")
	   Protection=RS("Protection")
	   ClassID=RS("ClassID")
	   Notes=RS("Notes")
	   Locked=RS("Locked")
	   EndTF=RS("EndTF")
	   Comment=RS("Comment")
	   AllowArrGroupID=RS("AllowArrGroupID")
	   AllowBMFlag=RS("AllowBMFlag")
	   minnum=RS("minnum")
	   limitbuynum=RS("limitbuynum")
	   Weight=RS("Weight")
	   recommend=RS("recommend")
	   ProvinceID=RS("ProvinceID")
	   CityID=RS("CityID")
	   HasBuyNum=RS("HasBuyNum")
	   MustPayOnline=RS("MustPayOnline")
	   CleanCart=RS("CleanCart")
	   showdelivery=RS("showdelivery")
	   Changes=KS.ChkClng(rs("Changes"))
	   ChangesUrl=rs("ChangesUrl")
	 End If
Else
  AllowBMFlag=0:Comment=1:Changes=0
  AddDate=Now: MustPayOnline=1 : CleanCart=1 : showdelivery=0
  ActiveDate=Now+10
  Locked=0:EndTF=0 :minnum=0:recommend=0:HasBuyNum=0
  Intro=" ":ProvinceID=0:CityID=0
 End If
%>
<script>
function CheckForm()
{
	if ($('#Subject').val()=='')
	{
	 top.$.dialog.alert('请输入团购主题!',function(){
	 $("#Subject").focus();});
	 return false;
	}
	if ($('#ClassID').val()=='0' || $('#ClassID').val()==undefined)
	{
	 top.$.dialog.alert('请选择团购分类!',function(){
	 $("#ClassID").focus();});
	 return false;
	}
	if (<%=GetEditorContent("Intro")%>==false)
	{
	 top.$.dialog.alert('请输入本单详情!',function(){
	<%=GetEditorFocus("Intro")%>});
	 return false;
	}

	if ($("#Price_Original").val()=='')
	{
	 top.$.dialog.alert('请输入原价!',function(){
	 $("#Price_Original").focus();});
	 return false;
	}
	if ($("#Discount").val()=='')
	{
	 top.$.dialog.alert('请输入折扣！',function(){
	 $("#Discount").focus();});
	 return false;
	}
	if (parseFloat($("#Discount").val())>10){
	 top.$.dialog.alert('折扣不能大于10！',function(){
	 $("#Discount").focus();});
	 return false;
	}
	if ($("#Price").val()=='')
	{
	 top.$.dialog.alert('请输入团购价！',function(){
	 $("#Price").focus();});
	 return false;
	}
   $("#myform").submit();
}
function regInput(obj, reg, inputStr)
{
		var docSel = document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")    return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange = obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
}
function getprice(discount){
     if (parseFloat(discount)>10){
	 alert('折扣不能大于10！');
	 $("#Discount").val(10);
	 return false;
	 }
     var Price_Original=$("#Price_Original").val();
	 if(Price_Original==''|| isNaN(Price_Original)){Price_Original=0;}
	 document.myform.Price.value=Math.round(Price_Original*(discount/10));
  }
$(document).ready(function(){
if ($("#Changes").prop('checked')){ChangesNews();}
});
function ChangesNews(){ 
		 if ($("#Changes").prop('checked'))
			  $("#ChangesUrl").attr("disabled",false);
		else
			  $("#ChangesUrl").attr("disabled",true);
}
</script>
<div class="pageCont2">
 <form name="myform" id="myform" action="?action=EditSave" method="post">
   <input type="hidden" value="<%=ID%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
<dl class="dtable">
          <dd>
            <div class="firstd">团购主题：</div>
            <input class='textbox' type='text' name='Subject' id='Subject' value='<%=Subject%>' size="60"> <font color=red>*</font>
          </dd>
          <dd>
            <div class="firstd">团购分类：</div>
           <select name="ClassID" id="ClassID">
			<option value='0'>---选择分类---</option>
			<%Dim RSC:Set RSC=Conn.Execute("select * From KS_GroupBuyClass Order By OrderID,ID")
			Do While Not RSC.Eof
			  If KS.ChkClng(ClassID)=RSC("ID") Then
			   Response.Write "<option value='" & RSC("ID") & "' selected>" & RSC("CategoryName") & "</option>"
			  Else
			   Response.Write "<option value='" & RSC("ID") & "'>" & RSC("CategoryName") & "</option>"
			  End If
			  RSC.MoveNext
			Loop
			RSC.Close
			Set RSC=Nothing
			%>
			</select>
          </dd>
		  <dd id='ContentLink'>
		     <div class="firstd">外部链接:</div><%
				If ChangesUrl = "" Then
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='60' class='textbox'>")
				Else
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl & "' size='60' class='textbox'>")
				End If
				If Changes = 1 Then
				 Response.Write (" <input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
				Else
				 Response.Write (" <input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'> 使用转向链接</font>")
				End If
			 %>
	   </dd>
		  
          <dd style="display:none">
            <div class="firstd">地区：</div>
            <script src="../../plus/area.asp?flag=getid"></script> <span style='color:red'>tips:地区不选择的话该团购切换所有地区都会显示</span>
			<script type="text/javascript">
			<%if KS.ChkClng(ProvinceID)<>0 then%>
				  $('#Province').val('<%=provinceid%>');
			<%end if%>
			 <%if KS.ChkClng(CityID)<>0 Then%>
				$('#City').val(<%=CityID%>);
			<%end if%>
			</script>
          </dd> 


          <dd>
            <div class="firstd">时间设置：</div>
            开始<input type='text' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  class='textbox' name='AddDate' value='<%=AddDate%>' size="40" /> 
			结束：<input type='text' class='textbox'  onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" name='ActiveDate' value='<%=ActiveDate%>' size="40" /> 
            &nbsp;<span class='tips'>如：<%=now%></span>
          </dd> 

		  
          <dd>
            <div class="firstd">购物车设置：</div>
            需要在线支付订单才生效：<label><input type="radio" name="MustPayOnline" value="0"<%if MustPayOnline="0" then response.write " checked"%>/>不需要</label>
			<label><input type="radio" name="MustPayOnline" value="1"<%if MustPayOnline="1" then response.write " checked"%>/>需要</label>&nbsp;<span class="tips">如凭订单号享受打折的团购，建议选择不需要在线支付。</span>

			<br/>当购物车里有商品时先清空：<label><input type="radio" onClick="$('#delivery').show();" name="cleancart" value="1"<%if cleancart="1" then response.write " checked"%>/>是</label>
			<label><input type="radio" name="cleancart" onClick="$('#delivery').hide();" value="0"<%if cleancart="0" then response.write " checked"%>/>否</label>
			<span class="tips">当选择购物车里有商品时先清空，则订单里只能有这件商品。</span>
			<%if cleancart="1" then%>
			<div id="delivery" style="font-weight:normal;font-size:12px;">
			<%else%>
			<div style="display:none;font-weight:normal;font-size:12px;" id="delivery">
			<%end if%>
			显示送货方式：<label><input type="radio" name="showdelivery" value="1"<%if showdelivery="1" then response.write " checked"%>/>显示</label><label><input type="radio" name="showdelivery" value="0"<%if showdelivery="0" then response.write " checked"%>/>不显示</label>
			<span class="tips">如本地商家打折等团购建议选择不显示</span>
			</div>
			</dd>
		   <dd>
		     <div>商品图片：</div>
			 <table width="100%" border="0">
			  <tr>
			   <td>
			 <div class="mt10">
		     <label>小图：</label><input class="textbox"  type="text" name="PhotoUrl" id="PhotoUrl" size="30" value="<%=photourl%>" /> <input class="button" type='button' name='Submit' value='选择小图...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.PhotoUrl,'pic');">
			 </div>
			 <div class="mt10">
			 <label>大图：</label><input value="<%=bigphoto%>" class="textbox" type="text" name='BigPhoto' id='BigPhoto' size="30" /> <input class="button" type='button' name='Submit' value='选择大图...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.BigPhoto,'pic');">
			 </div>
			 <div class="mt10">
              <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?showpic=pic&ChannelID=5&UpType=Pic' frameborder=0 scrolling=no width='100%' height='30'></iframe>
			 </div>
             </td>
			 <td>
             <div  style="float:right;margin:0 auto;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:100px;width:95px;border:1px solid #777777">
				<img src="<%=PhotoUrl%>" onerror="this.src='../../images/logo.png';" id="pic" style="height:100px;width:95px;">
		    </div>
            </td>
            </tr>
            </table>
    </dd>
		   
 
		   <dd>
		     <div class="firstd">价格设置：</div>
		     原价<input type="text" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" name="Price_Original" id="Price_Original" size="6" value="<%=Price_Original%>" style="text-align:center" />元 折扣<input class="textbox" onChange="getprice(this.value);" type="text" name="Discount" id="Discount" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" size="6" value="<%=Discount%>" style="text-align:center" />折  团购价<input type="text" name="Price" id="Price" size="6" value="<%=Price%>" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" style="text-align:center" />元
			 
			 重量：<input class="textbox" type='text' name='Weight' style="text-align:center" id='Weight' value='<%=Weight%>' size="6">KG
			 <span style='color:#999999'>计算运费用的,包邮请输入-1。</span>
  </dd>
		   <dd>
            <div class="firstd">最低人数：</div>
           <input class="textbox" type='text' name='minnum' style="text-align:center" id='minnum' value='<%=minnum%>' size="6"> 人 &nbsp;每人限制购买<input class="textbox" type='text' name='limitbuynum' style="text-align:center" id='limitbuynum' value='<%=limitbuynum%>' size="6"> 件 <font color=red>*</font> <span>不限制输入0</span>  初始已销售<input type='text' name='hasbuynum' style="text-align:center" class="textbox" id='hasbuynum' value='<%=hasbuynum%>' size="6"> 件 <span>(作弊用的)</span>
          </dd>  
		  <dd>
            <div class="firstd">本单详情：</div>
		  </dd>
		   <dd>
            <div class="firstd"></div>
			   <%
			   Response.Write EchoEditor("Intro",Intro,"Basic","96%","220px")
				%>
          </dd>
		   <dd>
            <div class="firstd">精彩卖点：</div>
           <textarea name='Highlights' cols="60" rows="4"><%=Highlights%></textarea>
          </dd>  
		   <dd>
            <div class="firstd">团购保障：</div>
           <textarea name='Protection' cols="60" rows="4"><%=Protection%></textarea>
          </dd>  
		  <dd>
            <div class="firstd">温馨提示：</div>
            <textarea name='Notes' cols="60" rows="4"><%=Notes%></textarea>
          </dd>  
		  <dd>
            <div class="firstd">允许参加团购的权限：</div>
		    <label><input type="radio" name="AllowBMFlag" value="0"<%if AllowBMFlag=0 then response.write " checked"%>>允许所有人报名参加,包括游客</label>
			<label><input type="radio" name="AllowBMFlag" value="1"<%if AllowBMFlag=1 then response.write " checked"%>>只允许会员报名参加</label>
			<label><input type="radio" name="AllowBMFlag" value="2"<%if AllowBMFlag=2 then response.write " checked"%>>只允许指定的会员组报名参加</label>		
          </dd>  
		  <dd>
            <div class="firstd">允许参加团购的会员组：<font>(当上面选择只允许指定的会员组参加时，请在此指定会员组)</font></div>
            <%=KS.GetUserGroup_CheckBox("AllowArrGroupID",AllowArrGroupID,5)%>	
          </dd> 
		  <dd>
            <div class="firstd">是否推荐：</div>
		    <input type="radio" name="recommend" value="0"<%if recommend=0 then response.write " checked"%>>否
			<input type="radio" name="recommend" value="1"<%if recommend=1 then response.write " checked"%>>是	
          </dd>  
		  <dd>
            <div class="firstd">是否允许评论：</div>
		    <input type="radio" name="comment" value="0"<%if comment=0 then response.write " checked"%>>不允许（关闭）
			<input type="radio" name="comment" value="1"<%if comment=1 then response.write " checked"%>>允许，评论内容需要审核
			<input type="radio" name="comment" value="2"<%if comment=2 then response.write " checked"%>>允许，评论不需要审核		
          </dd>  
		   
		  <dd>
            <div class="firstd">是否锁定：</div>
		    <input type="radio" name="locked" value="0"<%if locked=0 then response.write " checked"%>>否
			<input type="radio" name="locked" value="1"<%if locked=1 then response.write " checked"%>>是	
          </dd>  
		  <dd>
           <div class="firstd">是否结束：</div
		    ><input type="radio" name="endtf" value="0"<%if endtf=0 then response.write " checked"%>>否
			<input type="radio" name="endtf" value="1"<%if endtf=1 then response.write " checked"%>>是
          </dd>  
		 </dl>
	</form>    
</div> 
<%
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Subject:Subject=KS.LoseHtml(KS.G("Subject"))
       Dim ActiveDate:ActiveDate=KS.G("ActiveDate")
			if not isdate(ActiveDate) then
			 Call KS.AlertDoFun("本单载止日期格式不正确！","history.back();")
			 Exit Sub
			End If	  
       Dim AddDate:AddDate=KS.G("AddDate")
			if not isdate(AddDate) then
			 Call KS.AlertDoFun("发布时间格式不正确！","history.back();")
			 Exit Sub
		End If	 
			


	   Dim PhotoUrl:PhotoUrl=KS.G("PhotoUrl")
	   Dim BigPhoto:BigPhoto=KS.G("BigPhoto")


			 
	   Dim Intro:Intro=Request.Form("Intro")
	   Dim Fax:Fax=KS.LoseHtml(KS.G("Fax"))
	   Dim Highlights:Highlights=KS.LoseHtml(KS.G("Highlights"))
	   Dim Protection:Protection=KS.LoseHtml(KS.G("Protection"))
	   Dim Notes:Notes=KS.LoseHtml(KS.G("Notes"))
	   Dim Locked:Locked=KS.ChkClng(KS.G("Locked"))
	   Dim EndTF:EndTF=KS.ChkClng(KS.G("EndTf"))
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   Dim ClassID:ClassID=KS.ChkClng(KS.G("ClassID"))
	   Dim AllowBMFlag:AllowBMFlag=KS.ChkClng(KS.G("AllowBMFlag"))
	   Dim minnum:minnum=KS.ChkClng(KS.G("minnum"))
	   Dim AllowArrGroupID:AllowArrGroupID=KS.G("AllowArrGroupID")
	   Dim Price_Original:Price_Original=KS.G("Price_Original")
	   Dim Discount:Discount=KS.G("Discount")
	   Dim Price:Price=KS.G("Price")
	   Dim Weight:Weight=KS.G("Weight")
	   If Not IsNumeric(Weight) Then Weight=0
	   Dim recommend:recommend=KS.ChkClng(KS.G("recommend"))
	   Dim LimitBuyNum:LimitBuyNum=KS.ChkCLng(KS.G("LimitBuyNum"))
	   Dim ProvinceID:ProvinceID=KS.ChkClng(KS.G("province"))
	   Dim CityID:CityID=KS.ChkClng(KS.G("city"))
	   Dim HasBuyNum:HasBuyNum=KS.ChkClng(KS.G("hasbuynum"))
	   Dim MustPayOnline:MustPayOnline=KS.ChkClng(KS.G("MustPayOnline"))
	   Dim CleanCart:CleanCart=KS.ChkClng(KS.G("CleanCart"))
	   Dim Comment:Comment=KS.ChkClng(KS.G("Comment"))
	   Dim showdelivery:showdelivery=KS.ChkClng(KS.G("showdelivery"))
	   
		
	   If KS.IsNul(Subject) Then Call KS.AlertDoFun("团购主题必须输入！","history.back();") 
	   If not isnumeric(Price_Original) Then Call KS.AlertDoFun("原价必须输入正确的数字！","history.back();") 
	   If not isnumeric(Discount) Then Call KS.AlertDoFun("折扣必须输入正确的数字！","history.back();") 
	   If not isnumeric(Price) Then Call KS.AlertDoFun("团购价必须输入正确的数字！","history.back();") 

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_GroupBuy Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
				 RS("IsSuccess")=0
				 RS("PostTable")= LFCls.GetCommentTable()
				 RS("CmtNum") = 0
			  End If
				 RS("AddDate")=AddDate
			     RS("Subject")=Subject
				 RS("ActiveDate")=ActiveDate
				 RS("Intro")=Intro
				 RS("PhotoUrl")=PhotoUrl
				 RS("BigPhoto")=BigPhoto
				 RS("Highlights")=Highlights
				 RS("Protection")=Protection
				 RS("ClassID")=ClassID
				 RS("Notes")=Notes
				 RS("Locked")=Locked
				 RS("EndTF")=EndTF
				 RS("minnum")=minnum
				 RS("LimitBuyNum")=LimitBuyNum
				 RS("Weight")=Weight
				 RS("AllowBMFlag")=AllowBMFlag
				 RS("AllowArrGroupID")=AllowArrGroupID
				 RS("Price_Original")=Price_Original
				 RS("Discount")=Discount
				 RS("Price")=Price
				 RS("recommend")=recommend
				 RS("HasBuyNum")=HasBuyNum
				 RS("MustPayOnline")=MustPayOnline
				 RS("CleanCart")=CleanCart
				 RS("Comment")=Comment
				 RS("showdelivery")=showdelivery
				 RS("ProvinceID")=ProvinceID
				 RS("CityID")=CityID
				 RS("Changes")=KS.ChkClng(request("changes"))
				 RS("ChangesUrl")=Request.Form("ChangesUrl")
		 		 RS.Update
				 If ID=0 Then
				   RS.MoveLast
                   Call KS.FileAssociation(1005,RS("ID"),Intro&RS("PhotoUrl"),0)
				 Else
                   Call KS.FileAssociation(1005,ID,Intro&RS("PhotoUrl"),1)
				 End If
				 
			     RS.Close
				 Set RS=Nothing
				 If ID=0 Then
				  Call KS.ConfirmDoFun("团购信息发布成功!","location.href='Shop/KS.GroupBuy.asp?action=Add';","parent.frames['BottomFrame'].location.href='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("团购系统 >> <font color=red>管理首页</font>") & "';location.href='Shop/KS.GroupBuy.asp';")
				 Else
				  Call KS.AlertDoFun("团购信息修改成功！","parent.frames['BottomFrame'].location.href='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("团购系统 >> <font color=red>管理首页</font>") & "';location.href='"& ComeUrl & "';")
				 End If

EnD Sub

	'删除
	Sub GroupBuyDel()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 Conn.execute("Delete From KS_UploadFiles Where ChannelID=1005 and InfoID In("& id & ")")
	 Conn.execute("Delete From KS_GroupBuy Where id In("& id & ")")
	 KS.AlertHintScript "删除成功!"
	End Sub
	
	Sub Recommend()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 Conn.execute("Update KS_GroupBuy  set recommend=1 Where id In("& id & ")")
	 KS.AlertHintScript "恭喜，批量设置推荐成功!"
	End Sub
	Sub UnRecommend()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 Conn.execute("Update KS_GroupBuy  set recommend=0 Where id In("& id & ")")
	  KS.AlertHintScript "恭喜，批量取消推荐成功!"
	End Sub
	
	Sub GroupBuyendtf()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 Conn.execute("Update KS_GroupBuy Set endtf=1 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	Sub GroupBuyCancelendtf()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy Set endtf=0 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	Sub GroupBuyVerific()
		Dim ID:ID=KS.G("ID")
		If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 	Conn.execute("Update KS_GroupBuy Set Verific=1 Where id ="& id )
		Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	Sub GroupBuyVerific_t()
		Dim ID:ID=KS.G("ID")
		If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 	Conn.execute("Update KS_GroupBuy Set Verific=3 Where id ="& id )
		Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	
	'锁定
	Sub GroupBuyLock()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 Conn.execute("Update KS_GroupBuy Set locked=1 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	'解锁
	Sub GroupBuyUnLock()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
	 Conn.execute("Update KS_GroupBuy Set locked=0 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
    
	'添加团购分类
    Sub AddClass()
	 Dim CategoryName,OrderID,ID,RS
	 ID=KS.ChkClng(KS.G("ID"))
	 If ID=0 Then
	   OrderID=Conn.Execute("select Max(OrderID) From KS_GroupBuyClass")(0)
	   OrderID=KS.ChkClng(OrderID)+1
	 Else
	   Set RS=Conn.Execute("Select top 1 * From KS_GroupBuyClass Where ID=" & ID)
	   If Not RS.Eof Then
	    CategoryName=RS("CategoryName")
		OrderID=RS("OrderID")
	   End If
	   RS.Close
	   Set RS=Nothing
	 End If
	 
	%>
    <script>
	 function CheckForm()
   {
	    if ($("#CategoryName").val()==''){
			 top.$.dialog.alert('请输入分类名称!',function(){
				  $("#CategoryName").focus();
			 });
			 return false;
		}
		 $("#myform").submit(); 
	 }
	</script>
	<div class="pageCont2">
	<div style="margin:0 20px;font-weight:bold; font-size:14px;">
	<%If ID=0 Then%>
	添加团购分类
	<%else%>
	修改团购分类
	<%end if%>
	</div>
    <div class="pageCont2"><form name="myform" id="myform" action="?action=AddClassSave" onsubmit="return(CheckForm());" method="post">
    <input type="hidden" name="id" value="<%=id%>"/>
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
	<dl class="dtable">
          <dd>
            <div class='firstd'>分类名称：</div>
            <input type='text'  class="textbox" name='CategoryName' id='CategoryName' value='<%=CategoryName%>' size="40"> <font color=red>*</font>
          </dd> 
          <dd>
            <div class='firstd'>排列序号：</div>
           <input type='text'  class="textbox" name='OrderID' id='OrderID' value='<%=OrderID%>' size="5" style="text-align:center"> <font color=red>*</font></dd> 
          <dd>
           <input type="submit" value="确定保存" class="button"/>
          </dd> 
	</dl> 
	</form>
	</div>
	</div>
	<%
	End Sub
	
	Sub AddClassSave()
	 Dim CategoryName,OrderID,ID
	 CategoryName=KS.G("CategoryName")
	 OrderID=KS.ChkClng(KS.G("OrderID"))
	 ID=KS.ChkClng(KS.G("ID"))
	 If KS.IsNul(CategoryName) Then Call KS.AlertDoFun("请输入团购分类名称!","history.back();")
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 If ID=0 Then
	   RS.open "select top 1 * from KS_GroupBuyClass Where CategoryName='"& CategoryName & "'",CONN,1,1
	 Else
	   RS.open "select top 1 * from KS_GroupBuyClass Where ID<>" & ID & " and CategoryName='"& CategoryName & "'",CONN,1,1
	 End If
	 If Not RS.Eof Then
	  RS.Close:Set RS=Nothing
	  Call KS.AlertDoFun("对不起，您输入的团购分类名称已存在!","history.back();")
	 End If
	 RS.Close
	 
	 RS.Open "select top 1 * From KS_GroupBuyClass Where ID=" & ID,conn,1,3
	 If RS.Eof And RS.Bof Then
	 RS.AddNew
	 End If
	  RS("CategoryName")=CategoryName
	  RS("OrderID")=OrderID
	 RS.Update
	 RS.Close
	 Set RS=Nothing
	 
	 If ID=0 Then
	   Call KS.ConfirmDoFun("恭喜，团购分类添加成功,继续添加吗？","location.href='?action=AddClass'","location.href='?action=ClassManage';")
	 Else
	   Call KS.AlertDoFun("恭喜，团购分类修改成功!","location.href='?action=ClassManage';")
	 End If
	End Sub

Private Sub ClassManage()
%>
<div class="pageCont2">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>分类名称</th>
	<td nowrap>序号</th>
	<td nowrap>管理操作</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_GroupBuyClass order by orderid,id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>还没有添加任何团购分类！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=DelClass>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="30"  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td align="center" width="120" class="splittd">
	<%=Rs("CategoryName")%>
	</td>
	<td align="center" class="splittd">
	<%=Rs("OrderID")%>
	</td>

	<td align="center" class="splittd">
		
		<a href="?Action=AddClass&id=<%=RS("ID")%>" class="setA">修改</a>|
		<a href="?Action=DelClass&id=<%=RS("ID")%>" onClick="return(confirm('确定删除该分类吗?'))" class="setA">删除</a>
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
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的团购分类 " onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
</div>
<%
End Sub

 Sub DelClass()
   Dim ID:ID=KS.FilterIds(KS.G("ID"))
   If ID="" Then KS.Die "<script>alert('没有选择分类ID!');history.back();</script>"
   Conn.Execute("Delete From KS_GroupBuyClass Where  ID In (" & ID & ")")
   KS.Die "<script>alert('恭喜，删除成功!');location.href='KS.GroupBuy.asp?action=ClassManage';</script>"
 End Sub

End Class
%> 

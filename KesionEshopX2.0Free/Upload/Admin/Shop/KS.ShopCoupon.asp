<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../../plus/md5.asp"-->
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
			
		 If KS.G("action")="SearchUser" Then
		  SearchUser
		  Exit Sub
		 End If

		
		 With Response
			  If Not KS.ReturnPowerResult(5, "M520007") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
			  End If
			  If KS.S("from")<>"excel" Then
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>添加优惠券类型</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加优惠券类型</span></li>"
			 
			  .Write "<li class='changeWay'><strong>查看方式:</strong><a href=""KS.ShopCoupon.asp"">所有优惠券</a> <a href=""KS.ShopCoupon.asp?flag=1"">正常</a>  <a href=""KS.ShopCoupon.asp?flag=2"">未启用</a> <a href=""KS.ShopCoupon.asp?flag=3"">已到期</a></li>"

			  .Write "</ul>"
			 Else
			   Response.AddHeader "Content-Disposition", "attachment;filename=" & formatdatetime(now,2)&".xls" 
			   Response.ContentType = "application/vnd.ms-excel" 
			   Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			 End If
		End With
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
        If CInt(CurrentPage) = 0 Then CurrentPage = 1
		Select Case KS.G("action")
		 Case "Add","Edit" Call CouponManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call CouponDel()
		 Case "CouponType"  Call CouponType()
		 Case "CancelCouponType" Call CancelCouponType()
		 Case "Show" Call ShowDetail()
		 Case "Pub" Call PubCoupon()
		 Case "PubByUserSave" Call PubByUserSave()
		 Case "PubByUserGroupSave" Call PubByUserGroupSave()
		 Case "PubByxx" Call PubByxx()
		 Case "DelCoupon" Call DelCoupon()
		 Case Else
		  Call showmain
		End Select
		
		
End Sub

Private Sub showmain()
        Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		   Param= Param & " and title like '%" & KS.G("KeyWord") & "%'"
		End If
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="1" Then Param=Param & " and Status=1"
		  If KS.G("Flag")="2" Then Param=Param & " and Status=0"
		  If KS.G("Flag")="3" Then Param=Param & " and datediff("& DataPart_S & ",enddate," & SqlNowString & ")>0"
		End If

%>
<div class="tableTop">
<table><tr><td>
<form action="KS.ShopCoupon.asp" name="myform" method="get">
   <div>
     <strong class="mr0">快速搜索=></strong>
	 <span class="tiaoJian">关键字:</span><input type="text" class='textbox' name="keyword">&nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
    </div>
</form>
</td></tr></table>
</div>
<div class="pageCont2 mt20">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>优惠券名</td>
	<td nowrap>面值</td>
	<td nowrap>订单下限</td>
	<td nowrap>发放方式</td>
	<td nowrap>发放数量</td>
	<td nowrap>有效期</td>
	<td nowrap>状态</td>
	<td nowrap>管理操作</td>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_ShopCoupon " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan='10' class='splittd'>对不起,找不到符合条件的优惠券！</td></tr>"
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
	<td align="center" class="splittd"><font color=red><%=Rs("FaceValue")%></font> 元</td>
	<td align="center" class="splittd"><font color=red><%=Rs("MinAmount")%></font> 元</td>
	<td align="center" class="splittd">
	<%
	 Select Case RS("CouponType")
	  case 0 Response.Write "按用户发放"
	  case 1 Response.Write "线下发放"
	  case 2 Response.Write "按订单金额"
	  case 3 Response.Write "<font color=red>新会员自动发放</font>"
	 End Select
	%>
	</td>
	<td align="center" class="splittd"><span style="color:red;font-weight:bold"><%=LFCls.GetSingleFieldValue("Select Count(id) From KS_ShopCouponUser Where CouponId=" & RS("ID"))%></span>
	(<a href="#" onclick="javascript:window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>查看优惠券</font>')+'&ButtonSymbol=Disabled';location.href='KS.ShopCoupon.asp?Action=Show&id=<%=RS("ID")%>'">查看</a>)
	</td>
	<td align="center" class="splittd">
	<font color=#999999> <%=rs("begindate")%>
	 <br />至<br />
	 <%=rs("enddate")%>
	</font>
	</td>
	<td align="center" class="splittd"><%
	    if datediff("s",rs("enddate"),now)>0 then
		 response.write "<font color=green>已过期</font>"
		elseif rs("status")=1 then
		  response.write "<font color=#cccccc>正常</font>"
		else
		  response.write " <font color=red>禁用</font>"
		end if
	%></td>
	<td align="center" class="splittd">
	 <%if rs("coupontype")<>2 and RS("CouponType")<>3 Then%>
	<a href="?action=Pub&ID=<%=RS("ID")%>"  onclick="window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>发放优惠券</font>')+'&ButtonSymbol=Disabled';" class='setA'>发放</a>|
	<%end if%>
	
	<a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>修改优惠券信息</font>')+'&ButtonSymbol=GOSave';" class='setA'>修改</a>|<a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('删除优惠券,将同时删除原已分配数据且不可恢复,确定删除该优惠券吗？'));" class='setA'>删除</a>|
		
		<%IF rs("Status")="1" then %><a href="?Action=CancelCouponType&id=<%=rs("id")%>"  class='setA'>关闭</a><%else%><a href="?Action=CouponType&id=<%=rs("id")%>"  class='setA'>启用</a><%end if%>

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
	<td class='pt10' onMouseOver="this.className='pt10'" onMouseOut="this.className='pt10'" height='25' colspan=9>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的优惠券 " onclick="{if(confirm('删除优惠券,将同时删除原已分配数据且不可恢复,确定删除该优惠券吗？')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=9 style="text-align:right">
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

</div>
<%
End Sub

Sub CouponManage()
Dim Coupon,ActiveDate,BeginDate,EndDate,FaceValue,Score,Telphone,Intro,MinAmount,Protection,BuyFlow,Notes,CouponType,Status,PhotoUrl,MaxDiscount
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select * From KS_ShopCoupon Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	  Response.End
	 Else
	   Coupon=RS("Title")
	   BeginDate=RS("BeginDate")
	   EndDate=RS("EndDate")
	   FaceValue=RS("FaceValue")
	   MinAmount=RS("MinAmount")
	   CouponType=RS("CouponType")
	   Status=RS("Status")
	   MaxDiscount=RS("MaxDiscount")
	 End If
Else
  BeginDate=Now
  EndDate=dateadd("m",1,now)
  MinAmount=0:Score=10
  CouponType=0:Status=1
  FaceValue=10:MaxDiscount=0
  Intro=" "
  PhotoUrl="../../images/nopic.gif"
 End If
%>
<script>
function CheckForm()
{
	if ($('#Coupon').val()=='')
	{
	 top.$.dialog.alert('请输入优惠券名称!',function(){
	 $("#Coupon").focus();
	 });
	 return false;
	}
	if ($('input[name=FaceValue]').val()=='')
	{
	 top.$.dialog.alert('请输入优惠券面值!',function(){
	 $("input[name=FaceValue]").focus();});
	 return false;
	}
	if ($('input[name=MinAmount]').val()=='')
	{
	 top.$.dialog.alert('请输入最小订单金额!',function(){
	 $("input[name=MinAmount]").val();});
	 return false;
	}
	
document.myform.submit();
}
</script>
<script language="javascript" src="../../KS_Inc/DatePicker/WdatePicker.js"></script>
<div class="pageCont2">
<dl class="dtable">
  <form name="myform" id="myform" action="?action=EditSave" method="post">
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
    <dd>
      <div class="firstd">优惠券名称：</div>
          <input type='text' class="textbox"  name='Coupon' id='Coupon' value='<%=Coupon%>' size="50" />
          <font color=red>*</font></td>
    </dd>
    <dd>
      <div class="firstd">使用起始日期：</div>
	  <input name='BeginDate' type='text' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" id='BeginDate' value='<%=BeginDate%>' size='50'  class='textbox'>           
    </dd>
    <dd>
      <div class="firstd">使用结束日期：</div>
           <input name='EndDate' type='text' id='EndDate' value='<%=EndDate%>' size='50' onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" class='textbox'>      
        <span>过了这个时间将不能兑换 </span>
    </dd>
    <dd>
      <div class="firstd">优惠券面值：</div>
          <input type='text' name='FaceValue' class="textbox" style="text-align:center" value='<%=FaceValue%>' size="10" />
          元<span>可以抵销的最大金额,一旦设定建议不要再修改。</span>
    </dd>
    <dd>
      <div class="firstd">优惠券最大抵扣订单总金额的百分数：</div>
          <input type='text' name='MaxDiscount' class="textbox"  style='text-align:center' value='<%=formatnumber(MaxDiscount,2,-1)%>' size="10" />
          %
          <span>不限制请输入0,如一张200元的优惠券,最大抵用金额为订单总金额的50%,假设购物中订单总金额为240元， 这时系统只能使用优惠券抵扣 240 * 50% =120元，这样这张 200元的优惠券就剩 80元。只要在有效期内，剩余的80元优惠券 还可以用在其它订单当中。</span>
    </dd>
	
    <dd>
      <div class="firstd">优惠券发放方式：</div>
        <input type="radio" name="CouponType" value="0"<%if CouponType=0 then response.write " checked"%> />
按用户手工发放<br />

<input type="radio" name="CouponType" value="1"<%if CouponType=1 then response.write " checked"%> />
线下发放优惠券号 
<br />

<input type="radio" name="CouponType" value="3"<%if CouponType=3 then response.write " checked"%> />
<font color=green>给每位新注册用户自动发放(新增)</font>
<br />

<span style="display:none">
&nbsp;
<input type="radio" name="CouponType" value="2"<%if CouponType=2 then response.write " checked"%> />
按订单金额发放
</span>
</dd>
    <dd>
      <div class="firstd">最小订单金额：</div>
          <input name='MinAmount' class="textbox"  value="<%=MinAmount%>" stype="text-align:center" size="10" />
          元<span>只有商品总金额达到这个数的订单才能使用这种优惠券</span>
    </dd>
    <dd>
      <div class="firstd">状态：</div>
          <input type="radio" name="Status" value="1"<%if Status=1 then response.write " checked"%> />
          正常
          <input type="radio" name="Status" value="0"<%if Status=0 then response.write " checked"%> />
        禁用 
    </dd>
</dl>
  </form>
<div class="attention">
			 <font color=red><strong>操作说明:</strong></font><br />
			 1.优惠券必须发放后,才可以使用<br />
			 2.第一步先添加优惠券类型,选择优惠券发放方式<br />
			 3.第二步根据对应的发放方式,对优惠券进行发放,按“新注册用户自动发放”时，则不需要人工发放，系统会自动给新注册的会员发放<br />
</div>
</div>
<%
End Sub

Sub ShowDetail()
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim K,SQL,Subject,CouponType
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 1 Title,CouponType FROM KS_ShopCoupon Where ID=" & ID,Conn,1,1
 If RS.Eof And RS.Bof Then
  RS.Close
  Set RS=Nothing
  KS.AlertHintScript "对不起,参数出错!"
  Exit Sub
 Else
  Subject=RS(0)
  CouponType=RS(1)
 End If
 RS.Close
 
 RS.Open "Select a.ID,a.CouponNum,a.OrderID,a.UserName,a.UseFlag,a.UseTime,a.AvailableMoney,b.FaceValue,b.maxdiscount,a.note From KS_ShopCouponUser a inner join KS_ShopCoupon b on a.couponid=b.id Where a.CouponID=" & ID & " Order By a.ID Desc",conn,1,1
 If Not RS.Eof Then
 	    totalPut = RS.Recordcount
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
        If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		if KS.S("isall")="1" THEN
           SQL=RS.GetRows(-1)
        ELSE
           SQL=RS.GetRows(MaxPerPage)
		END IF
 End If
 RS.Close:Set RS=Nothing
%>
<div class="pageCont2">
<div style="height:45px;font-size:14px;line-height:45px;text-align:center;font-weight:bold">查看 <font color=red>[<%=Subject%>]</font> 的优惠券列表</div>
<form name="selform" method="post" action="?action=DelCoupon">
<%If KS.S("from")<>"excel" Then%>
<table width="100%" border="0" align="center" style="border-top:1px solid #cccccc" cellspacing="0" cellpadding="0">
<%else%>
<table width="100%" border="1" align="center" style="border-top:1px solid #cccccc" cellspacing="0" cellpadding="0">
<%end if%>
<tr height="25" align="center" class='sort'>
  <%If KS.S("from")<>"excel" Then%>
	<td width='5%' nowrap>选择</td>
  <%end if%>
	<td nowrap>优惠券号</td>
	<td nowrap>优惠券面值</td>
	<td nowrap>最大抵扣额</td>
	<td nowrap>可用金额</td>
	<td nowrap>使用人</td>
	<td nowrap>使用情况</td>
 <%If KS.S("from")<>"excel" Then%>
	<td nowrap>管理操作</td>
 <%end if%>
</tr>
<%
If Not IsArray(SQL) Then 
  Response.Write "<tr><td class='splittd'  colspan='10' height='25' align='center'>对不起,找不到优惠券!</td></tr>"
Else
  For K=0 To Ubound(SQL,2)
   Response.Write "<tr>"
   If KS.S("from")<>"excel" Then
   Response.Write "	<td class='splittd' align='center'><input type=checkbox name=ID value='" & SQL(0,k) & "'></td>"
   end if
   Response.Write "	<td class='splittd' nowrap>" & SQL(1,k) & "</td>"
   Response.Write "	<td class='splittd' align='center'>" & formatnumber(SQL(7,k),2,-1) & " 元</td>"
   Response.Write " <td class='splittd' align='center'>" 
   If SQL(8,K)=0 Then
    Response.Write "实际优惠券面值"
   Else
    Response.Write "按订单总额的" & formatnumber(SQL(8,K),2,-1) & "%,但不超过实际优惠券面值"
   End If
   Response.Write "</td>"
   Response.Write "	<td class='splittd' align='center'>" & formatnumber(SQL(6,k),2,-1) & " 元</td>"
   Response.Write " <td class='splittd' nowrap align='center'>&nbsp;"
   If SQL(3,K)<>"" Then
    Response.Write SQL(3,K)
   Else
    Response.Write "<font color=#999999>未分配</font>"
   End If
   Response.Write "&nbsp;</td>"
   Response.Write "	<td class='splittd' align=""center"">" 
    If SQL(4,K)=1 Then
	 if SQL(6,k)>0 then
	  response.write "已使用,未用完"
	 else
	  response.write "已用完"
	 end if
	 response.write "<span style='cursor:pointer' onclick=""top.$.dialog({title:'说明',content:'" & SQL(9,K) & "',width:350})""><font color=red>(详情)</font></span>"
	Else
	 Response.Write "<font color=#999999>未使用</font>"
	End If
   Response.Write "</td>"
   If KS.S("from")<>"excel" Then
   Response.Write " <td class='splittd' nowrap align='center'><a href='?action=DelCoupon&id=" & sql(0,k) & "' onclick=""return(confirm('确定删除吗?'))"">删除</a></td>"
   end if
   Response.Write "</tr>"

  Next
End If
If KS.S("from")<>"excel" Then%>
<tr>
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=10>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的优惠券 " onclick="{if(confirm('此操作不可恢复,确定删除该优惠券吗？')){this.document.selform.submit();return true;}return false;}">
	<%If CouponType<>2 and CouponType<>3 Then%>
	<input class="button" type="button" value=" 发 放 " onclick="location.href='?action=Pub&id=<%=ID%>'">
	<%End If%>
	
	<input class="button" type="button" value=" 打 印 " onclick="window.print()">
	<input class="button" type="button" value=" 导出本页至Excel " onclick="location.href='KS.ShopCoupon.asp?page=<%=currentpage%>&Action=Show&id=<%=ID%>&from=excel';">
	<input class="button" type="button" value=" 导出所有至Excel " onclick="location.href='KS.ShopCoupon.asp?isall=1&Action=Show&id=<%=ID%>&from=excel';">
	</td>
</tr>
</form>
<tr>
	<td  colspan=8 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
</div>
<%end if%>

<%
End Sub


'发放优惠券
Sub PubCoupon()
 Dim CouponType,EndDate,Title
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 1 * FROM KS_ShopCoupon Where ID=" & ID,conn,1,1
 IF RS.Eof And RS.Bof Then
  RS.Close:Set RS=Nothing
  KS.AlertHintScript "出错啦!"
  Exit Sub
 End If
 CouponType=RS("CouponType")
 Title=RS("Title")
 EndDate=RS("EndDate")
 RS.Close:Set RS=Nothing
 If DateDiff("s",EndDate,Now)>0 Then
  KS.AlertHintScript "该优惠券已过期,不能再分配!"
  Exit Sub
 End If
 Select Case CouponType
   Case 0    '按用户发放
            %>
				 <div class="pageCont2">
				 <div class="tabTitle">1.按用户组发放</div>
				 <table width="99%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
					  <form name="myform" action="?action=PubByUserGroupSave" method="post">
						<input type="hidden" value="<%=ID%>" name="id" />
						<input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
					  <tr class="tdbg">
						<td height='25' width="100" class='clefttitle' align="right"><strong>按用户组发放：</strong></td>
						<td width="400">
						<%=KS.GetUserGroup_CheckBox("GroupID","",5)%>
						</td>
						<td style="text-align:left"><input type="checkbox" name="sendtips" value="1" checked="checked" />发送站内消息通知会员<br/>
						<input type="checkbox" name="sendemail" value="1" checked="checked" />发送电子邮件通知会员<br/>
						<input type="checkbox" name="sendsms" value="1" checked="checked" />发送手机短信通知会员<br/>
						</td>
						<td>
						 <input type="hidden" name="EndDate" value="<%=EndDate%>"/>
						<input type="submit" name="send_user" value="确定发放优惠券" class="button" /></td>
					  </tr>
					  
					  </form>
					</table>
					<div class="blank20"></div>
					<div class="tabTitle">2.搜索指定会员发放</div>
					<table width="99%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
					 <tr>
					  <td colspan="3">
					   <div class="tdbg">
					<form name="theForm" action="KS.ShopCoupon.asp" method="post" onsubmit="return validate();">
					<div class="form-div">
					  关键字：<input type="text" name="keyword" class="textbox" id="keyword" size="30" />
					  <input type="button" class="button" name="search" value=" 搜索用户 " onclick="searchUser();" />
					</div>
					</form>
					  </td>
					  </tr>
					  <form name="myform" action="?action=PubByUserSave" method="post">
					    <input type="hidden" name="EndDate" value="<%=EndDate%>"/>
						<input type="hidden" value="<%=ID%>" name="id" />
						<input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
					  <tr align='center'>
						<td height='25' class='clefttitle' style="text-align:center">会员列表</td>
						<td  class='clefttitle' style="text-align:center">操作</td>
						<td  class='clefttitle' style="text-align:left;padding-left:10px;" colspan="3">给下列用户发放优惠券</td>
					  </tr>
					  <tr>
						<td class='tdbg' align="center">
						  <select name="user_search" id="user_search" size="15" style="width:260px;height: 100px;" ondblclick="addUser()" multiple="true">
						  </select>
						</td>
						<td class='tdbg' align="center">
						  <p><input type="button" value=" &gt; " onclick="addUser()" class="button" /></p>
						  <p><input type="button" value=" &lt; " onclick="delUser()" class="button" /></p>
						</td>
						<td class='tdbg' style="padding-left:10px">
						  <select name="user" id="user" multiple="true" size="15" style="width:260px;height: 100px;" ondblclick="delUser()">
						  </select>
						</td>
						<td style="text-align:left"><input type="checkbox" name="sendtips" value="1" checked="checked" />发送站内消息通知会员<br/>
						<input type="checkbox" name="sendemail" value="1" checked="checked" />发送电子邮件通知会员<br/>
						<input type="checkbox" name="sendsms" value="1" checked="checked" />发送手机短信通知会员<br/>
						</td>
						<td><input type="submit" name="send_user" onclick="return(check())" value="确定发放优惠券" class="button" /></td>
					  </tr>
					  </form>
					</table>
					</div>
				<script type="text/javascript">
					 function searchUser()
					 {
					  var url = 'KS.ShopCoupon.asp'; 
					    
					  $.get(url,{action:"SearchUser",keyword:escape($("#keyword").val())},function(s){
					   showResponse(unescape(s))
					  })    
					 }
					 function showResponse(s)
					  {
						  var result=s;
						  if (result!='')
						  {
							  var obj = $('#user_search')[0];
							  obj.length = 0;
							  var rarr=result.split('|');
								for (var i = 0; i < rarr.length; i++)
								{
									 if (rarr[i]!=''){
									  var opt = document.createElement('OPTION');
									  opt.value = rarr[i];
									  opt.text  = rarr[i];
									  obj.options.add(opt);
									   }
								}
						  }
					 }
				function addUser()
				  {
					  var src = document.getElementById('user_search');
					  var dest = document.getElementById('user');
				
					  for (var i = 0; i < src.options.length; i++)
					  {
						  if (src.options[i].selected)
						  {
							  var exist = false;
							  for (var j = 0; j < dest.options.length; j++)
							  {
								  if (dest.options[j].value == src.options[i].value)
								  {
									  exist = true;
									  break;
								  }
							  }
							  if (!exist)
							  {
								  var opt = document.createElement('OPTION');
								  opt.value = src.options[i].value;
								  opt.text = src.options[i].text;
								  dest.options.add(opt);
							  }
						  }
					  }
				  }
				
				  function delUser()
				  {
					  var dest = document.getElementById('user');
				
					  for (var i = dest.options.length - 1; i >= 0 ; i--)
					  {
						  if (dest.options[i].selected)
						  {
							  dest.options[i] = null;
						  }
					  }
				  }
				function check()
				{
					var idArr = new Array();
					var dest = document.getElementById('user');
					for (var i = 0; i < dest.options.length; i++)
					{
						dest.options[i].selected = "true";
						idArr.push(dest.options[i].value);
					}
					if (idArr.length <= 0)
					{
						alert("你没有选择用户!");
						return false;
					}
					else
					{
						return true;
					}
				  
				}
				</script>
<%
   Case 1   '线下发放
 %>
            <br />
				 <strong>&nbsp;<font color=red size=2>批量生成优惠券号</font></strong>
				 <table width="99%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
					  <form name="myform" action="?action=PubByxx" method="post">
						<input type="hidden" value="<%=ID%>" name="id" />
						<input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
					  <tr class="tdbg">
						<td height='25' width="200" class='clefttitle' align="right"><strong>归属优惠券类型：</strong></td>
						<td>
						<%=Title%>
						</td>
					  </tr>
					  <tr class="tdbg">
						<td height='25' width="200" class='clefttitle' align="right"><strong>生成数量：</strong></td>
						<td>
						<input type="text" name="Num" value="100" class="textbox" style="text-align:center;width:60px">
						 张 <font color=red>*</font></td>
					  </tr>
					  <tr class="tdbg">
						<td colspan='2' height="30" align="center">
						<input type="submit" name="send_user" value="确定批量生成优惠券号" class="button" />
						</td>
					  </tr>
					  </form>
					</table>
					
					<div style="margin-top:20px;border:1px solid #f1f1f1;line-height:21px;padding-left:10px">
			 <font color=red><strong>操作说明:</strong></font><br />
			 1.当优惠券类型为线下发放时,必须先生成优惠券号<br />
			 2.您可以将生成的优惠券号发放给用户使用<br />
			 3.用户得到优惠券号后,可以在会员中心启用,每张优惠券号只能用一次<br />
</div>

 <%End Select
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Coupon:Coupon=KS.LoseHtml(KS.G("Coupon"))
       Dim BeginDate:BeginDate=KS.G("BeginDate") 
	   if not isdate(BeginDate) then
	       Call KS.AlertDoFun("起始时间格式不正确！","history.back();")
			 Exit Sub
	   End If	 
       Dim EndDate:EndDate=KS.G("EndDate")
	   if not isdate(EndDate) then
	    Call KS.AlertDoFun("截止时间格式不正确！","history.back();")
			 Exit Sub
	   End If	 
			
			 
	   Dim FaceValue:FaceValue=KS.G("FaceValue")
	   If Not IsNumeric(FaceValue) Then
	     Call KS.AlertHistory("优惠券面值不正确!")
		 Exit Sub
	   End If
	   Dim MinAmount:MinAmount=KS.G("MinAmount")
	   If Not IsNumeric(MinAmount) Then
	     Call KS.AlertHistory("最小订单金额不正确,请重输!")
		 Exit Sub
	   End If
	   Dim CouponType:CouponType=KS.ChkClng(KS.G("CouponType"))
	   Dim Status:Status=KS.ChkClng(KS.G("Status"))
	   Dim MaxDiscount:MaxDiscount=KS.G("MaxDiscount")
	   If Not IsNumeric(MaxDiscount) Then
	     Call KS.AlertHistory("最大抵扣百分比不正确,请重输!")
		 Exit Sub
	   End If
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
		
	        If Coupon="" Then Call KS.AlertDoFun("优惠券名称必须输入！","history.back();") :response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_ShopCoupon Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
				 RS("Inputer")=KS.C("AdminName")
				 RS("AddDate")=Now
			  End If
				 RS("BeginDate")=BeginDate
				 RS("EndDate")=EndDate
			     RS("Title")=Coupon
				 RS("FaceValue")=FaceValue
				 RS("MinAmount")=MinAmount
				 RS("CouponType")=CouponType
				 RS("MaxDiscount")=MaxDiscount
				 RS("Status")=Status
		 		 RS.Update
			     RS.Close
				 Set RS=Nothing
				 If ID=0 Then
				  Call KS.ConfirmDoFun("优惠券类型添加成功,继续添加吗？","location.href='shop/KS.ShopCoupon.asp?action=Add';","$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>优惠券管理</font>") & "';location.href='shop/KS.ShopCoupon.asp';")
				 Else
				  Call KS.AlertDoFun("优惠券类型修改成功！","$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>优惠券管理</font>") & "';location.href='"& ComeUrl & "';")
				 End If

EnD Sub

'删除
Sub CouponDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
 Conn.execute("Delete From KS_ShopCoupon Where id In("& id & ")")
 Conn.execute("Delete From KS_ShopCouponUser Where couponid In("& id & ")")
 Call KS.Alert("删除成功！",Request.servervariables("http_referer"))
End Sub

'删除单张
Sub DelCoupon()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
 Conn.execute("Delete From KS_ShopCouponUser Where id In("& id & ")")
 Call KS.Alert("删除成功！",Request.servervariables("http_referer"))
End Sub

Sub CouponType()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
 Conn.execute("Update KS_ShopCoupon Set Status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub CancelCouponType()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Call KS.AlertDoFun("对不起，您没有选择!","history.back();")
 Conn.execute("Update KS_ShopCoupon Set Status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

Sub PubByUserSave()
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim User:User=Replace(KS.G("User")," ","")
 Dim K,UserArr,RS
 If User="" Then
   KS.AlertHintScript "对不起,您没有选择用户"
   Exit Sub
 End If
 UserArr=Split(User,",")
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 1 * From KS_ShopCoupon Where ID=" & ID,conn,1,1
 If RS.Eof  And RS.Bof Then
    RS.Close:Set RS=Nothing
	Exit Sub
 End If
 Dim Money,EndDate
 Money=RS("FaceValue")
  EndDate=RS("EndDate")
 RS.Close
 
 For K=0 To Ubound(UserArr)
  Dim RSUser:Set RSUser=Conn.Execute("Select top 1 * From KS_User Where UserName='" & UserArr(k) & "'")
  If NOT RSUser.Eof Then
      Dim CouponNum:CouponNum=GetCouponNum()
	  RS.Open "Select * From KS_ShopCouponUser Where 1=0",conn,1,3
	  RS.AddNew
	  RS("CouponID")=ID
	  RS("CouponNum")=CouponNum
	  RS("UserName")=UserArr(k)
	  RS("OrderID")=""
	  RS("UseFlag")=0
	  RS("AddDate")=Now
	  RS("AvailableMoney")=Money
	  RS.Update
	  RS.Close
	  Call SendTips(RSUser("username"),RSUser("RealName"),RSUser("Email"),RSUser("Mobile"),CouponNum,Money,EndDate)
  End If
  RSUser.Close
  Set RSUser=Nothing
 Next
 KS.Alert "恭喜,成功发放" & K & "个用户!","Shop/KS.ShopCoupon.asp?Action=Show&ID=" & ID
End Sub

Sub PubByUserGroupSave()
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim GroupID:GroupID=KS.FilterIds(KS.G("GroupID"))
 If GroupID="" Then
   KS.AlertHintScript "对不起,您没有选择用户组"
   Exit Sub
 End If
 
 Dim RS,K:K=0
 Dim RSUser:Set RsUser=Conn.Execute("Select UserName,Email,Mobile,RealName From KS_User Where GroupID in(" & GroupID & ")")
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 1 * From KS_ShopCoupon Where ID=" & ID,conn,1,1
 If RS.Eof  And RS.Bof Then
    RS.Close:Set RS=Nothing
	Exit Sub
 End If
 Dim Money,EndDate
 Money=RS("FaceValue")
 EndDate=RS("EndDate")
 RS.Close
 
 Do While Not RSUser.Eof
      K=K+1
	  Dim CouponNum:CouponNum=GetCouponNum()
	  RS.Open "Select * From KS_ShopCouponUser Where 1=0",conn,1,3
	  RS.AddNew
	  RS("CouponID")=ID
	  RS("CouponNum")=CouponNum
	  RS("UserName")=RSUser(0)
	  RS("OrderID")=""
	  RS("UseFlag")=0
	  RS("AddDate")=Now
	  RS("AvailableMoney")=Money
	  RS.Update
	  RS.Close
	  Call SendTips(RSUser(0),RSUser("RealName"),RSUser("Email"),RSUser("Mobile"),CouponNum,Money,EndDate)
  RSUser.MoveNext
 Loop
 RSUser.Close
 Set RSUser=Nothing
 KS.Alert "恭喜,成功发放" & K & "个用户!","Shop/KS.ShopCoupon.asp?Action=Show&ID=" & ID
End Sub

Sub SendTips(UserName,RealName,Email,Mobile,CouponNum,Money,EndDate)
 Dim sendtips:sendtips=KS.ChkClng(KS.G("sendtips"))
 Dim sendEmail:sendEmail=KS.ChkClng(KS.G("sendEmail"))
 Dim sendSms:sendSms=KS.ChkClng(KS.G("sendSms"))
 If KS.IsNul(RealName) Then RealName=UserName
 
 Dim MailContent:MailContent=KS.Setting(186)
 MailContent=Replace(MailContent,"{$UserName}",UserName)
 MailContent=Replace(MailContent,"{$CouponNum}",CouponNum)
 MailContent=Replace(MailContent,"{$Money}",Money)
 MailContent=Replace(MailContent,"{$EndDate}",EndDate)
 Dim SmsContent:SmsContent=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(11)
 SmsContent=Replace(SmsContent,"{$username}",UserName)
 SmsContent=Replace(SmsContent,"{$couponnum}",CouponNum)
 SmsContent=Replace(SmsContent,"{$money}",Money)
 SmsContent=Replace(SmsContent,"{$enddate}",EndDate)
 
 
 
 If SendTips=1 and Not KS.IsNul(MailContent) Then   
	'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
	Call KS.SendInfo(UserName,KS.C("AdminName"),"获得购物优惠券通知",MailContent)
 End If
 If SendEmail=1 and Not KS.IsNul(MailContent) And Not KS.IsNul(Email) Then 
    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "在[" & KS.Setting(0) & "]网站获得购物优惠券通知", Email,RealName, MailContent,KS.Setting(11))  
 End If
 
 If sendSms=1 And Not KS.IsNul(SmsContent) And Not KS.IsNul(Mobile) Then
   Call KS.SendMobileMsg(Mobile,SmsContent)
 End If

End Sub

'批量生成线下优惠券号
Sub PubByxx()
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim Num:Num=KS.ChkClng(KS.G("Num"))
 Dim K,RS
 If Num=0 Then
   KS.AlertHintScript "对不起,输入的生成张数必须大于0"
   Exit Sub
 End If
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 1 * From KS_ShopCoupon Where ID=" & ID,conn,1,1
 If RS.Eof  And RS.Bof Then
    RS.Close:Set RS=Nothing
	Exit Sub
 End If
 Dim Money
 Money=RS("FaceValue")
 RS.Close
 For K=1 To Num
  RS.Open "Select top 1 * From KS_ShopCouponUser Where 1=0",conn,1,3
  RS.AddNew
  RS("CouponID")=ID
  RS("CouponNum")=GetCouponNum()
  RS("UserName")=""
  RS("OrderID")=""
  RS("UseFlag")=0
  RS("AddDate")=Now
  RS("AvailableMoney")=Money
  RS.Update
  RS.Close
 Next
 KS.Alert "恭喜,批量生成了" & K-1 & "张优惠券号!","Shop/KS.ShopCoupon.asp?Action=Show&ID=" & ID
End Sub


Function GetCouponNum()
   Do While True
	 GetCouponNum = "C" & KS.MakeRandom(10)
	 If Conn.Execute("Select CouponNum from KS_ShopCouponUser Where CouponNum='" & GetCouponNum & "'").Eof Then Exit Do
   Loop
End Function

Sub SearchUser()
 Dim KeyWord:KeyWord=KS.DelSQL(Unescape(Request("KeyWord")))
 Dim RS,UserArr,I
 If KeyWord="" Then Exit Sub
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 500 UserName FROM KS_User Where UserName like '%" & KeyWord & "%'",conn,1,1
 If Not RS.Eof Then
  UserArr=RS.GetRows(-1)
 End If
 RS.Close:Set RS=Nothing
 If IsArray(UserArr) Then
  For I=0 To Ubound(UserArr,2)
   Response.Write Escape(UserArr(0,i)) & "|"
  Next
 End If
End Sub

End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
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
			  If Not KS.ReturnPowerResult(5, "M520006") Then          '检查是权限
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
			  If KS.S("Action")="Add" or KS.S("Action")="Edit" Then
			  	.Write "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon save'></i>确定保存</span></li>"
			   .Write "<li class='parent' onclick=""location.href='?';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消返回</span></li>"

			  Else
			  .Write "<li class='parent' onclick=""window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>添加商品规格</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加商品规格</span></li>"
			  End If
			  .Write "</ul>"
		End With
		
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
		Select Case KS.G("action")
		 Case "Add","Edit" Call ShopSpecificationManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call DoDelete()
		 Case Else
		  Call showmain
		End Select
		
		
End Sub

Private Sub showmain()
        Param=" where 1=1"


		totalPut = Conn.Execute("Select Count(id) From KS_ShopSpecification " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
%>
<div class='pageCont2'>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>规格名称</td>
	<td >绑定分类</td>
	<td nowrap>显示类型</td>
	<td nowrap>序号</td>
	<td nowrap>管理操作</td>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_ShopSpecification " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>到不起,还没有添加任何商品规格！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=Del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td  class="splittd"><%=Rs("Title")%></td>
	<td class="splittd" style="width:300px">
	<%
	Dim RSC,Str:Set RSC=Conn.Execute("select FolderName from ks_class c inner join KS_ShopSpecificationR r on c.id=r.classid where c.channelid=5 and sid= " & RS("ID"))
	If Not RSC.Eof Then
	      str=""
		  Do While Not RSC.Eof
		   str=str & RSC(0) & " "
		   RSC.MoveNext
		  Loop
		  RSC.Close:Set RSC=Nothing
		  response.write "<span style=""cursor:hand"" onclick=""top.$.dialog({title:'绑定分类',content:'"& str &"',min:false,max:false})"">" & KS.Gottopic(str,100) & "</span>"
	End If
	%>
	</td>
	<td align="center" class="splittd">
	<%if rs("showtype")="1" then
	 response.write "文字"
	 else
	 response.write "图片"
	 end if%>
	</td>
	<td align="center" class="splittd"><%=ks.chkclng(rs("orderid"))%></td>
	<td align="center" class="splittd">
	
	
	<a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('商城系统 >> <font color=red>修改商品规格</font>')+'&ButtonSymbol=GOSave';" class="setA">修改</a>|<a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('删除商品规格不可恢复,确定删除该商品规格吗？'));" class="setA">删除</a> 

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
	&nbsp;&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的规格 " onclick="{if(confirm('删除商品规格不可恢复,确定删除吗？')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
<div>
   <div class="attention">
      &nbsp;<strong>说明：</strong>
	 &nbsp;每个商品分类系统最多只能绑定三个规格，即对应顾客购买时可以选择三级规格属性，如一件衣服可以选择颜色为红色，尺码为大码等。
    </div>
</div>
</div>
<%
End Sub

Sub ShopSpecificationManage()
Dim Title,ShowType,SValue,orderid
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select top 1 * From KS_ShopSpecification Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	  Response.End
	 Else
	   Title    = RS("Title")
	   ShowType = RS("ShowType")
	   SValue   = RS("Svalue")
	   orderid  = RS("orderid")
	 End If
Else 
    ShowType = 1 : orderid=1
End If
%>
<script>
function CheckForm()
{
	if ($('#Title').val()==''){
	 top.$.dialog.alert('请输入规格名称!',function(){
	 $("#Title").focus();
	 });
	 return false;
	}
	if ($("#classid option:selected").val()==undefined){
	 top.$.dialog.alert('请选择要绑定的商品分类!');
	 return false;
	}

 var num=parseInt($("#num").val());
 if (num==0){
  top.$.dialog.alert('规格值必须输入!');
  return false;
 }

 for (var i=0;i<num;i++){
   if ($("#item"+i).val()==''){
     top.$.dialog.alert('第'+(i+1)+'个规格值必须输入!',function(){
	 $("#item"+i).focus();
	 });
	 return false;
   }
   if (parseInt($("input[name=ShowType]:checked").val())==2){
	   if ($("#itempic"+i).val()==''){
		 top.$.dialog.alert('第'+(i+1)+'个规格值的图片必须输入!',function(){
		 $("#itempic"+i).focus();});
		 return false;
	   }
   }
 }
 
document.myform.submit();
}
function showpic(v){
 if (v==1){
  $("#table1").find("[name='tp']").hide();
 }else{
  $("#table1").find("[name='tp']").show();
 }
}
function doadd(){
    var str="";
	var ss='display:none;';
	if (parseInt($("input[name='ShowType']:checked").val())==2){
	ss=''
	}

	var num=parseInt($("#num").val());
    str=str+"<tr id='tr"+num+"'><td class='splittd'><input type=text class=textbox name=item id=item"+num+" size=30></td><td name='tp' class='splittd' style='"+ss+";text-align:center'><input type=text class=textbox name=itempic id=itempic"+num+" size=26>&nbsp;<input class='button'  type='button' name='Submit' value='选择...' onClick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir%>',550,290,window,document.myform.itempic" +num+ ");\"></td><td class='splittd' style='text-align:center'><a href='javascript:deltd("+num+")' name='delbtn'>删除</a></td></tr>";
     jQuery("#additem").append(str);
	 jQuery("#num").val(parseInt(jQuery("#num").val())+1);
}
function deltd(i){
 $("#tr"+i).remove();
}


</script>
<div class="pageCont2">
  <form name="myform" id="myform" action="?action=EditSave" method="post">
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
<dl class="dtable">
    <dd><div class='firstd'>规格名称：</div>
          <input type='text' class='textbox' name='Title' id='Title' value='<%=Title%>' size="40" />
          <font color=red>*</font> <span>如颜色，尺寸等。</span>
    </dd>
    
    <dd>
      <div class='firstd'>显示类型：</div>
     <label><input type='radio' onclick="showpic(1)" name='ShowType' value='1' <%if ShowType="1" then response.write " checked"%>/>文字</label> <label><input type="radio" onclick="showpic(2)" value="2" name="ShowType"<%if ShowType="2" then response.write " checked"%>/>图片</label> 
    </dd>
    <dd>
      <div class='firstd'>序号：</div>
      <input type='text' class='textbox' style='text-align:center' name='orderid' value='<%=orderid%>' size='5'/> <span class='tips'>值越小排在越前面</span>
    </dd>
	<dd>
	 <div>绑定分类：<font>(可以绑定到多个分类下，按ctrl键进行多选)</div>
	<select size='10' style='width:280px;height:80px' multiple name='classid' id='classid'>
	 <%
			Dim C_L_Str:C_L_Str=KS.LoadClassOption(5,true)
			If ID<>0 Then
				Dim RSB:Set RSB=Conn.Execute("Select ClassID From KS_ShopSpecificationR Where SID=" & ID)
				iF Not RSB.Eof Then
				  Do While Not RSB.Eof
				  C_L_Str=Replace(C_L_Str,"value='" & RSB(0) & "'","value='" & RSB(0) &"' selected")
				  RSB.MoveNext
				  Loop
				End If
				RSB.Close:Set RSB=Nothing
			End If
	%>
	<%=C_L_Str%>
	</select>
   </dd>
   </dl>
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="ctable">
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td colspan="2" style="padding:20px">
	   <input type='button' class="button" onclick="doadd()" value="添加规格值">
	    <table width="98%" id='table1' class="ctable mt20" cellpadding="0" cellspacing="0" border="0">
		  <tr class="sort">
		    <td style="text-align:center">规则格值名称</td>
			<td name='tp' style='<%if ShowType="1" then response.write "display:none;"%>width:400px;text-align:center'>图片</td>
			<td style="width:200px;text-align:center">操作</td>
		  </tr>
		  
			<%if SValue<>"" then
			 Dim v1,v2,ss,num,i,SArr:SArr=Split(SValue,",")
			 num=ubound(sarr)+1
			 for i=0 to ubound(sarr)
			  if showtype="1" then
			    ss="display:none"
				v1=trim(sarr(i))
				v2=""
			  else
			    ss=""
				v1=trim(split(sarr(i),"|")(0))
				v2=trim(split(sarr(i),"|")(1))
			  end if
			  response.write "<tr id='tr" & i & "'><td class='splittd'><input class='textbox' type=text value='" & v1 &"' name=item id=item" & i & " size=30></td><td name='tp' class='splittd' style='" & ss & ";text-align:center'><input class='textbox' value='" & v2 &"' type=text name=itempic id=itempic" & i & " size=26>&nbsp;<input class='button'  type='button' name='Submit' value='选择...' onClick=""OpenThenSetValue('../Include/SelectPic.asp?ChannelID=5&CurrPath=" & KS.GetUpFilesDir&"',550,290,window,document.myform.itempic" & i & ");""> </td><td class='splittd' style='text-align:center'><a href='javascript:deltd(" & i & ")' name='delbtn'>删除</a></td></tr>"
			 next
			else
			  num=0
			end if%>
		  <tbody id="additem"> </tbody>
		</table>
	  	   <input type='hidden' name="num" id="num" value="<%=num%>"/>

	   </td>
    </tr>
</table>
  </form>
</div>
<%
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Title:Title=KS.G("Title")
       Dim ShowType:ShowType=KS.ChkClng(KS.G("ShowType")) 
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   Dim Items:Items=KS.G("Item")	
	   Dim itempic:itempic=KS.G("itempic")
	   Dim OrderID:OrderID=KS.ChkClng(KS.G("OrderID"))
	   Dim ClassID:ClassID=Replace(KS.G("ClassID")," ","")
	        If Title="" Then KS.Die "<script>alert('商品规格名称必须输入');history.back();</script>"
            If Items="" Then KS.Die "<script>alert('商品规格值必须输入');history.back();</script>"
            If ShowType=2 and ItemPic="" Then KS.Die "<script>alert('商品规格图片必须输入');history.back();</script>"
            Dim SValue,iarr,parr,ii,k,ClassID_Arr
			If ClassID="" Then Call KS.AlertHistory("请选择商品规格归属分类！",-1)
			ClassID_Arr=Split(ClassID,",")
			For K=0 To Ubound(ClassID_Arr)
			  dim bindnum
			  if ID<>0 Then
			  bindnum=KS.ChkClng(Conn.Execute("Select count(*) From KS_ShopSpecificationR Where Sid<>" & id  & " AND ClassID='" & ClassID_Arr(K) &"'")(0))
			  Else
			  bindnum=KS.ChkClng(Conn.Execute("Select count(*) From KS_ShopSpecificationR Where ClassID='" & ClassID_Arr(K) &"'")(0))
			  End If
			  If bindnum>=3 Then
			  Call KS.AlertHistory("对不起，每个商品分类最多只能绑定三个规格，分类[" & KS.C_C(ClassID_Arr(K),1) & "]已绑定" & bindnum &"个规格了！",-1)
			  End If
			Next 
			iarr=split(Items,",")
			parr=split(ItemPic,",")
			for ii=0 to ubound(iarr)
			  if trim(iarr(ii))<>"" then
			    if showtype=1 then
				  if svalue="" then
				   svalue=trim(iarr(ii))
				  else
				   svalue=svalue&","&trim(iarr(ii))
				  end if
				else
				  if svalue="" then
				   svalue=trim(iarr(ii))&"|"&trim(parr(ii))
				  else
				   svalue=svalue&","&trim(iarr(ii))&"|"&trim(parr(ii))
				  end if
				end if
			  end if
			next
            
            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_ShopSpecification Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
			  End If
				 RS("Title")   =Title
				 RS("ShowType")=ShowType
				 RS("Svalue")  =Svalue
				 RS("OrderID") =OrderID
		 		 RS.Update
				 If ID=0 Then
				    RS.MoveLast
				    ID=RS("ID")
					RS.Close:Set RS=Nothing
				  For K=0 To Ubound(ClassID_Arr)
					 Conn.Execute("Insert Into KS_ShopSpecificationR(ClassID,SID) values('" & ClassID_Arr(K) & "'," & ID & ")")
				  Next
				  Call KS.ConfirmDoFun("商品规格添加成功!","location.href='shop/KS.ShopSpecification.asp?action=Add';","$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>商品规格列表管理</font>") & "';location.href='shop/KS.ShopSpecification.asp';")
				 Else
				   RS.Close:Set RS=Nothing
				    Conn.Execute("Delete From KS_ShopSpecificationR Where SID=" & ID)
					For K=0 To Ubound(ClassID_Arr)
					 Conn.Execute("Insert Into KS_ShopSpecificationR(ClassID,SID) values('" & ClassID_Arr(K) & "'," & ID & ")")
					Next
				  Call KS.AlertDoFun("商品规格修改成功！","$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>商品规格列表管理</font>") & "';location.href='"& ComeUrl & "';")
				 End If

EnD Sub

'删除
Sub DoDelete()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>top.$.dialog.alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Delete From KS_ShopSpecification Where id In("& id & ")")
 Conn.execute("Delete From KS_ShopSpecificationR Where sid In("& id & ")")
 KS.AlertHintScript "恭喜，删除成功！"
End Sub


End Class
%> 

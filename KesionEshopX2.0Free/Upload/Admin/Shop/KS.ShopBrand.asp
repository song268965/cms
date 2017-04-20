<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Brand
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Brand
        Private KS,Action,ComeUrl,Page,ItemName,Table,ClassID
		Private I,totalPut,CurrentPage,KeySql,RS,MaxPerPage,KSCls
		Private Sub Class_Initialize()
		  MaxPerPage =18
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		Call KS.CreateBrandCache()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		
		  With Response
			.Write "<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
			.Write "<title>商品品牌管理</title>"
			.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
	        .Write "<script language=""JavaScript"" src=""../../KS_Inc/Jquery.js""></script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>" & vbCrLf
			.Write EchoUeditorHead
             Action=KS.G("Action")
			   If Not KS.ReturnPowerResult(5, "M510018") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End if
			
			CurrentPage = KS.ChkClng(KS.G("page"))
			If CurrentPage<1 Then CurrentPage = 1

			 ItemName=KS.C_S(5,3)
			 Page=KS.G("Page")
			 ClassID=KS.G("ClassID")
			 
			 Select Case Action
			  Case "Add" Call BrandAddOrEdit("Add")
			  Case "Edit" Call BrandAddOrEdit("Edit")
			  Case "Del"  Call BrandDel()
			  Case "DoSave" Call DoSave()
			  Case "Create" Call CreateJS()
			  Case "SaveCreateJS" Call SaveCreateJS()
			  Case "ShowDetail" Call ShowBrandClass()
			  Case "ProList" Call ProList()
			  Case "BrandProList" Call BrandProList()
			  Case "remove" Call Remove()
			  Case Else
			   Call BrandList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		  End With
		End Sub
			  
		Sub ShowBrandClass()
		  Dim RS:Set RS=Conn.Execute("select FolderName from ks_class c inner join ks_classbrandr r on c.id=r.classid where brandid= " & KS.ChkCLng(KS.G("ID")))
		  If Not RS.Eof Then
		   Response.Write "<div style='margin:5px'><br><strong>品牌<font color=red>" & KS.G("BrandName") & "</font>属于分类：</strong><br>"
		  Do While Not RS.Eof
		   Response.Write RS(0) & " "
		   RS.MoveNext
		  Loop
		  RS.Close:Set RS=Nothing
		  Response.Write "</div>"
		  End If
		End Sub
		
		Sub ProList()
		 dim BrandID:BrandID=KS.ChkClng(request("BrandID"))
		 If BrandID=0 Then KS.Die "error"
          
  		 Dim RS:Set rs=server.CreateObject("adodb.recordset")
		 RS.Open "select top 1 BrandName from ks_classbrand where id=" & brandid,conn,1,1
		 if rs.eof and rs. bof then
		  rs.close
		  set rs=nothing
		  ks.alerthintscript "出错啦!"
		 end if
		 dim brandname:brandname=rs(0)
		 rs.close

		  
		  With Response
			.Write "<div class=""topdashed sort"">查看管理品牌[<span style='color:green'>" & brandname &"</span>]下的商品</div>"
		  
		  %>
		  <script type="text/javascript">
		  function getProduct()
		  {			 
		     $(parent.document).find("#ajaxmsg").toggle("fast");
			 var key=escape($('input[name=key]').val());
			 var tid=$('#tid>option:selected').val();
			 var priceType=$('#PriceType>option:selected').val();
			 var minPrice=$("#minPrice").val();
			 var maxPrice=$("#maxPrice").val();
			 var str='';
			 if (key!=''){
			   str='商品名称:'+key;
			 } 
			 if (tid!=''){
			   str+=' 栏目:'+$('#tid>option:selected').get(0).text
			 }
			 if (priceType!=0){
			   str+= minPrice +' 元';
			   switch (parseInt(priceType)){
			     case 1 :
				  str+='<=当前零售价<=';
				  break;
			     case 2 :
				   str+='<=商城价<=';
				   break;
			     case 3 :
				  str+='<=原始零售价<=';
				  break;
			   }
			   str+= maxPrice +' 元';
			   
			 }
			 if (str!='') str='<strong>条件:</strong><font color=red>'+str+'</font>';
			 $("#keyarea").html(str);
			 
			 $.get("../../plus/ajaxs.asp", { action: "GetPackagePro", proid:$("#proids").val(),pricetype:priceType,key: key,tid:tid,minPrice:minPrice,maxPrice:maxPrice},
			 function(data){
					$(parent.document).find("#ajaxmsg").toggle("fast");
					$("#prolist").empty().append(data);
			  });
		  }
		</script>
    <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
		<table width="100%" border="0">
		  <tr>
			<td style="text-align:left">
			  &nbsp;<strong>快速搜索=></strong>
			  <br/>
			   &nbsp;商品编号: <input type="text" class="textbox" name="proids" id="proids" size='15'> 可留空<br/>
			 &nbsp;商品名称: <input type="text" class='textbox' name="key">
			 <br/>&nbsp;所属栏目: <select size='1' name='tid' id='tid'><option value=''>--栏目不限--</option><%=KS.LoadClassOption(5,false)%></select>
			 <br/>&nbsp;价格范围:
			<input type='text' name='minPrice' size='5' class='textbox' style='text-align:center' id='minPrice' value='10'> 元
			<= <select name="PriceType" id="PriceType">
			  <option value=0>--不限制--</option>
			  <option value=1>当前零售价</option>
			  <option value=2>商城价</option>
			  <option value=3>原始零售价</option>
			 </select>
			 <= <input type='text' name='maxPrice' size='5' class='textbox' style='text-align:center' id='maxPrice' value='100'> 元
			  
			  <br/> <br/>
			  &nbsp;<input type="button" onclick="getProduct()" value="开始搜索" class="button" name="s1">
			
			</td>
			<form name="myform" id="myform" action="KS.ShopBrand.asp?action=BrandProList&flag=add" method="post" target="packframe">
			<input type="hidden" name="brandid" value="<%=BrandID%>"/>
			<td>
			<div id='keyarea'></div>
			<strong>查询到的商品:</strong>			
			<br/>
			 <select name="prolist" size="5" style="width:260px;height:140px" multiple="multiple" id="prolist"></select>
			 <br/>
			 <input type="submit" value="将选中的商品加入选购品" class="button">
			</td>
			</form>
		  </tr>
		</table>
		 </div>	
		<iframe name="packframe" src="?action=BrandProList&brandid=<%=BrandID%>" width="100%" height="100%" frameborder="0" scrolling="auto"></iframe>	  
		  <%
	  End With
End Sub

Sub BrandProList()
		 dim BrandID:BrandID=KS.ChkClng(request("BrandID"))
		 If BrandID=0 Then KS.Die "error"
		 
		 if request("flag")="add"  then
		   dim proids:proids=KS.FilterIds(request("prolist"))
		   If Proids<>"" then
		     conn.execute("update ks_product set brandid=" & brandid & " where id in(" & Proids & ")")
		   end if
		     Call KS.AlertDoFun("恭喜,在该品牌下加入商品成功!","location.href='?action=BrandProList&brandid=" & brandid & "';")
		 end if
          
  		 Dim RS:Set rs=server.CreateObject("adodb.recordset")
	  With Response
		 	.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.Write ("<form name='myform' method='Post' action='?'>")
	        .Write ("<input type='hidden' name='action' id='action' value='remove'>")
			.Write "  <tr>"
			.Write "          <td class=""sort"" width='35' align='center'>选择</td>"
			.Write "          <td class='sort' align='center'>小图</td>"
			.Write "          <td class='sort' align='center'>商品名称</td>"
			.Write "          <td class='sort' align='center'>会员价</td>"
			.Write "          <td class='sort' align='center'>管理操作</td>"
			.Write "  </tr>"

		 
		 RS.Open "select id,title,photourl,price_member,tid,fname,adddate from KS_Product Where BrandID=" & BrandID,conn,1,1
		  If RS.EOF And RS.Bof Then
		       .Write "<tr><td class='splittd' colspan=6>找不到属于该品牌的商品!</td></tr>"
		  Else
			   totalPut = RS.RecordCount
			   if CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrentPage - 1) * MaxPerPage
			   End If
					Do While Not RS.EOF
					 dim photourl:PhotoUrl=rs("photourl")
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../../images/nopic.gif"
			          .Write "<tr class=""list"" onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
			          .Write "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
					  .write "<td class='splittd' align='center'><img style='margin:2px;border:1px solid #f1f1f1;padding:1px' src='" & photourl & "' width='40'/></td>"
					  .Write "<td class='splittd'><a href='" & KS.GetItemUrl(5,rs("tid"),rs("id"),rs("fname"),rs("adddate")) & "' target='_blank'>" & RS("title") & "</a></td>"
					  .Write "  <td class='splittd' align='center'><font color=brown>" & rs("price_member") & "</font> 元</td>"
					  .Write "  <td class='splittd' align='center'> <a href='?action=remove&id=" & rs("id") & "'>移除</a> </td>"
					  .Write "</tr>"
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RS.MoveNext
					Loop
					  RS.Close
						
		      End If
			
			.Write "<tr><Td colspan=5><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> <input type='submit' value='批量移除' class='button'/></div></td></tr></form><tr><td colspan=6 align='right'>" 
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .Write ("</td></tr></table><br/><br/>")
			  
			End With			   
		End Sub
		
		Sub Remove()
		  Dim id:id=KS.FilterIds(KS.G("ID"))
		  If ID="" Then KS.AlertHintScript "对不起，没有选择商品!"
		  Conn.Execute("Update KS_Product Set BrandID=0 Where ID in(" & Id &")")
		  KS.AlertHintScript "恭喜，操作成功!"
		End Sub
			 
		Sub BrandList()			
			With Response
            .Write "<script language='JavaScript'>"
			.Write "var Page='" & CurrentPage & "';"
			.Write "var ItemName='" & ItemName & "';"
			.Write "</script>"
			%>
			<script language="javascript">
			    function set(v)
				{
				 if (v==1)
				 BrandControl(1);
				 else if (v==2)
				 BrandControl(2);
				}
				function BrandAdd()
				{
					location.href='KS.ShopBrand.asp?Action=Add';
					window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape(ItemName+"管理中心 >> 品牌管理 >> <font color=red>新增品牌</font>")+'&ButtonSymbol=GO';
				}
				function EditBrand(id)
				{
					location="KS.ShopBrand.asp?Page="+Page+"&Action=Edit&ID="+id;
					window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape(ItemName+"管理中心 >> 品牌管理 >> <font color=red>编辑品牌</font>")+'&ButtonSymbol=GOSave';
				}
				function DelBrand(id)
				{
				if (confirm('真的要删除该品牌吗?'))
				 location="KS.ShopBrand.asp?Action=Del&Page="+Page+"&id="+id;
				  SelectedFile='';
				}
				function BrandControl(op)
				{  var alertmsg='';
	               var ids=get_Ids(document.myform);
					if (ids!='')
					 {  if (op==1)
						{
						if (ids.indexOf(',')==-1) 
							EditBrand(ids)
						  else alert('一次只能编辑一个品牌tags!')	 
						}	
					  else if (op==2)    
						 DelBrand(ids);
					 }
					else 
					 {
					 if (op==1)
					  alertmsg="编辑";
					 else if(op==2)
					  alertmsg="删除"; 
					 else
					  {
					  alertmsg="操作" 
					  }
					 alert('请选择要'+alertmsg+'的品牌');
					  }
				}
				function GetKeyDown()
				{ 
				if (event.ctrlKey)
				  switch  (event.keyCode)
				  {  case 90 : location.reload(); break;
					 case 65 : Select(0);break;
					 case 78 : event.keyCode=0;event.returnValue=false; BrandAdd();break;
					 case 69 : event.keyCode=0;event.returnValue=false;BrandControl(1);break;
					 case 68 : BrandControl(2);break;
				   }	
				else	
				 if (event.keyCode==46)BrandControl(2);
				}
			</script>
			<%
			.Write "<body topmargin='0' leftmargin='0'  onkeydown='GetKeyDown();' onselectstart='return false;'>"
			%>
			<script type="text/javascript">
			 function ShowDetail(param){  
				top.openWin("查看品牌所属分类","shop/KS.ShopBrand.asp?Action=ShowDetail&"+param+"&rnd="+Math.random(),false,400,150);
			 }
			</script>
			<%
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onClick=""BrandAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加品牌</span></li>"
			.Write "<li class='parent' onClick=""BrandControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon write'></i>编辑品牌</span></li>"
			.Write "<li class='parent' onClick=""BrandControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon delete'></i>删除品牌</span></li>"
			.Write "</ul>"
			.Write "<div class='tableTop'><table><tr><td><form name=""sform"" action=""KS.ShopBrand.asp"" method=""post""><div id='go'>关键词：<input type=""text"" size=""15"" value=""" & KS.G("Key") & """ class=""textbox"" name=""key"" /> <input type=""submit"" value=""搜索品牌"" class=""button""/>"
			.Write "&nbsp;&nbsp;&nbsp;<select size='1' name='classid' onchange=""location.href='?classid='+this.value;""><option value='0'>--按分类快速查看--</option>"
			.Write Replace(KS.LoadClassOption(5,false),"value='" & ClassID & "'","value='" & ClassID &"' selected") & " </select>"
			.Write "</form></div></td></tr></table></div>"
			
			If Not KS.IsNul(Request("Key")) Then
			 .write "<div class='pageCont2 mt20'>搜索关键词“<font color=red>" & KS.G("Key") & "</font>”搜索结果:</div>"
			End If
			.Write ("<div class='pageCont2 mt20'><form name='myform' method='Post' action='?'>")
			.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	        .Write ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
			.Write "        <tr class='sort'>"
			.Write "          <td class=""sort"" width='35' align='center'>选择</td>"
			.Write "          <td class='sort' style='width:400px' align='center'>品牌名称</td>"
			.Write "          <td class='sort' align='center'>所属分类</td>"
			.Write "          <td class='sort' align='center'>商品数</td>"
			.Write "          <td align='center' class='sort'>是否顶级显示</td>"
			.Write "          <td align='center' class='sort'>是否推荐</td>"
			.Write "          <td class='sort' align='center'>排序号</td>"
			.Write "  </tr>"
			  
			  Set RS = Server.CreateObject("ADODB.RecordSet")
			  Dim Param:Param=" Where 1=1"
			  If Request("key")<>"" Then
			    Param=Param & " and (brandname like '%" & KS.G("Key") & "%' or brandename like '%" & KS.G("Key") & "%')"
			  End If
			  If ClassID<>"" and classid<>"0" Then 
			    KeySql="select a.* from KS_ClassBrand a inner join KS_ClassBrandR  r On a.id=r.brandid" & Param & " and r.classid='" & ClassID & "'"
			  Else
			  	KeySql = "SELECT * FROM [KS_ClassBrand] " & Param & " order by iD desc"
              End If
			   RS.Open KeySql, conn, 1, 1
					 If Not RS.EOF Then
						totalPut = RS.RecordCount
						If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
						End If
						Call showContent
				End If
			.Write "</table>"
			.Write ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .Write ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	        .Write ("</td>")
	        .Write ("<td><select style='height:30px' onchange='set(this.value)' name='setattribute'><option value=0>快速选项...</option><option value='1'>编辑品牌</option><option value='2'>执行删除</option></select></td>")
	        .Write ("</form><td align='right'>")
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .Write ("</td></tr></table></div>")
            End With
			End Sub
			
			Sub showContent()
			   With Response
					Do While Not RS.EOF
			          .Write "<tr class=""list"" onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
			          .Write "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
					  .Write "<td class='splittd' height='20'><span BrandID='" & RS("ID") & "' ondblclick=""EditBrand(" & RS("ID") & ")""><i class='icon photo'></i>"
					  .Write "  <span style='cursor:default;'>" & RS("BrandName") & "</span></span>"
					  .Write "<span class='noshow'><a href=""../../shop/brand.asp?brandid=" & RS("ID") &""" target=""_blank"">预览</a></span>"
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'><a  href='#' onclick='javascript:ShowDetail(""id=" & rs("id") & "&brandname=" & RS("BrandName") & """);'>已属于<font color=red>" & conn.execute("select count(*) from ks_class c inner join ks_classbrandr r on c.id=r.classid where brandid= " & rs("id"))(0) & "</font>个分类</a>" 
					   
					  .Write " </td>"
					  
					  .Write "  <td class='splittd' align='center'><font color=green>" & conn.execute("select count(1) from ks_product where brandid=" & rs("id"))(0) & "</font> 件 <a href='?action=ProList&brandid=" & rs("id") & "'>管理</a></td>"
					  .Write "  <td class='splittd' align='center'>"
					  if rs("showintop")=1 then
					   .write "<FONT Color=red>是</font>"
					  else
					   .write "否"
					  end if
					  .Write " </td>"
					  .Write "  <td class='splittd' align='center'>"
					  if rs("recommend")=1 then
					   .write "<FONT Color=red>是</font>"
					  else
					   .write "否"
					  end if
					  .Write " </td>"
					  .Write "  <td class='splittd' align='center'>" & RS("orderid") & " </td>"
					  .Write "</tr>"
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RS.MoveNext
					Loop
					  RS.Close
				End With
			End Sub
			
			
			
			Sub BrandAddOrEdit(OpType)
			With Response
			 Dim RS,Action, BrandName,BrandEName ,ID, PageRS, KeySql,Page,ClassID,OrderID,ShowInTop,Recommend,PhotoUrl,Intro,firstAlphabet
			  ID = KS.ChkClng(KS.G("ID"))
			  Page = KS.G("Page")
			 If OpType = "Edit" Then
				 Set RS = Server.CreateObject("ADODB.RECORDSET")
				 KeySql = "Select top 1 * From [KS_ClassBrand] Where ID=" & ID
				 RS.Open KeySql, conn, 1, 1
				 If Not RS.EOF Then 
				  BrandName = RS("BrandName")
				  BrandEname=RS("BrandEname")
				  firstAlphabet=RS("firstAlphabet")
				  Intro=RS("Intro")
				  OrderID = RS("OrderID")
				  ShowInTop=RS("ShowInTop")
				  Recommend=RS("Recommend")
				  PhotoUrl=RS("PhotoUrl")
				 End If
				 RS.Close:Set RS=Nothing
			 Else
			   ShowInTop=1:Recommend=0:orderid=1:classid=ks.g("classid"):ID=0
			   orderid=KS.ChkClng(conn.execute("select max(orderid) from ks_classbrand ")(0))+1
			 End If
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon save'></i>确定保存</span></li>"
			.Write "<li class='parent' onclick=""location.href='?';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消返回</span></li>"
		    .Write "</ul>"
			.Write "" & vbCrLf
			.Write "<div class='pageCont2'><form  action='?Action=DoSave' method='post' name='BrandForm' onsubmit='return(CheckForm())'>" & vbCrLf
			.Write "<dl class=""dtable"">" & vbCrLf
			.Write "    <dd>" & vbCrLf
			.Write "      <div>品牌名称：</div>" & vbCrLf
			.Write "        <input name='BrandName' type='text' size='50' id='BrandName' value='" & BrandName & "' class='textbox'>" & vbCrLf
			.Write "        *<span>如三星，飞利浦等</span>" & vbCrLf
			.Write "    </dd>" & vbCrLf
			.Write "    <dd>" & vbCrLf
			.Write "      <div>品牌英文名称：</div>" & vbCrLf
			.Write "      <input name='BrandEName' type='text' size='50' id='BrandEName' value='" & BrandEName & "' class='textbox'>" & vbCrLf
			.Write "     <span>如SamSung等</span>" & vbCrLf
			.Write "    </dd>" & vbCrLf
			.Write "    <dd>" & vbCrLf
			.Write "      <div>品牌首字母：</div>" & vbCrLf
			.Write "        <Select name='firstAlphabet' id='firstAlphabet'>"
			Dim I
			For I=65 To 90
				If firstAlphabet=chr(I) Then
				.Write "<option value='" & I& "' selected>" & chr(I) &"</option>"
				Else
				.Write "<option value='" & I& "'>" & chr(I) &"</option>"
				End If
			Next
			
			.Write "        </select>" & vbCrLf
			.Write "        * <span>方便前台按字母查找品牌</span>" & vbCrLf
			.Write "    </dd>" & vbCrLf
			
			
			.Write "    <dd>" & vbCrLf
			.Write "      <div>绑定分类：<font>(当一个品牌有多个分类时，可以按ctrl键进行多选)</font></div>" & vbCrLf
			.Write " <select size='12' style='width:350px' multiple name='classid'>"
			Dim C_L_Str:C_L_Str=KS.LoadClassOption(5,false)
			If ID<>0 Then
				Dim RSB:Set RSB=Conn.Execute("Select ClassID From KS_ClassBrandR Where BrandID=" & ID)
				iF Not RSB.Eof Then
				  Do While Not RSB.Eof
				  C_L_Str=Replace(C_L_Str,"value='" & RSB(0) & "'","value='" & RSB(0) &"' selected")
				  RSB.MoveNext
				  Loop
				End If
				RSB.Close:Set RSB=Nothing
			End If
			.Write C_L_Str & " </select>"
			.Write "    </dd>" & vbCrLf
			.Write "    <dd>" & vbCrLf
			.Write "      <div>品牌介绍：</div>" & vbCrLf
			
			Response.Write EchoEditor("Intro",Intro,"Basic","96%","220px")
			
			
			.Write "     </dd>" & vbCrLf			
			
			.Write "    <dd>" & vbCrLf
			.Write "      <div>品牌图片：</div>" & vbCrLf
			.Write "        <input name='PhotoUrl' style='width:300px' id='PhotoUrl' type='text' value='" & PhotoUrl & "' class='textbox'>&nbsp;<input class=""button"" type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=" & KS.GetUpFilesDir() & "',550,290,window,$('#PhotoUrl')[0]);""> " & vbCrLf
			.Write "    </dd>" & vbCrLf			
			.Write "    <dd>" & vbCrLf
			.Write "      <div>排 序 号：</div>" & vbCrLf
			.Write "        <input name='OrderID' type='text' value='" & OrderID & "' class='textbox'>" & vbCrLf
			.Write "        *<span>数据越小，前台排在越前面 </span>" & vbCrLf
			.Write "    </dd>" & vbCrLf
			.Write "    <dd style='display:none'>" & vbCrLf
			.Write "      <div>是否顶级树型菜单显示：</div>" & vbCrLf
			.Write "        <input name='ShowInTop' type='radio' value='1'"
			If ShowInTop="1" Then .Write " checked"
			.Write ">是" & vbCrLf
			.Write "        <input name='ShowInTop' type='radio' value='0'"
			If ShowInTop="0" Then .Write " checked"
			.Write ">否" & vbcrlf 
			.Write "    </dd>" & vbCrLf
			.Write "    <dd>" & vbCrLf
			.Write "      <div>是否推荐：</div>" & vbCrLf
			.Write "        <input name='Recommend' type='radio' value='1'"
			If Recommend="1" Then .Write " checked"
			.Write ">是" & vbCrLf
			.Write "        <input name='Recommend' type='radio' value='0'"
			If Recommend="0" Then .Write " checked"
			.Write ">否" & vbcrlf 
			.Write "    </dd>" & vbCrLf
			
			
			.Write "    <input type='hidden' value='" & ID & "' name='ID'>" & vbCrLf
			.Write "    <input type='hidden' value='" & Page & "' name='Page'>" & vbCrLf
			.Write "</dl>" & vbCrLf
			.Write "  </form></div>" & vbCrLf

			.Write "<Script Language='javascript'>" & vbCrLf
			.Write "<!--" & vbCrLf
			.Write "function CheckForm()" & vbCrLf
			.Write "{ var form=document.BrandForm;" & vbCrLf
			.Write "   if (form.BrandName.value=='')" & vbCrLf
			.Write "    {" & vbCrLf
			.Write "     top.$.dialog.alert('请输入品牌!',function(){" & vbCrLf
			.Write "     form.BrandName.focus();});" & vbCrLf
			.Write "     return false;" & vbCrLf
			.Write "    }" & vbCrLf
			.Write "    form.submit();" & vbCrLf
			.Write "}" & vbCrLf
			.Write "//-->" & vbCrLf
			.Write "</Script>" & vbCrLf
			.Write "</body>" & vbCrLf
			.Write "</html>" & vbCrLf
			End With
			End Sub

			
			Sub DoSave()
			    Dim RS,ClassID,ClassID_Arr,ID,K
				Dim  BrandName:BrandName = KS.G("BrandName")
				ID=KS.ChkClng(KS.G("ID"))
				 ClassID=Replace(KS.G("ClassID")," ","")
				 If BrandName = "" Then  KS.Die "<script> $.dialog.alert('请输入品牌名称！',function(){ history.back(-1);});</script>" 
				 If ClassID="" Then  KS.Die "<script> $.dialog.alert('请选择品牌归属分类！',function(){ history.back(-1);});</script>" 
				 ClassID_Arr=Split(ClassID,",")
				  If not Conn.Execute("Select ID From KS_ClassBrand Where ID<>" & ID & " and BrandName='" & BrandName & "'").eof Then KS.Die "<script> $.dialog.alert('您输入的品牌名称已存在！',function(){ history.back(-1);});</script>" 
				 Set RS = Server.CreateObject("ADODB.RECORDSET")
				  KeySql = "Select * From [KS_ClassBrand] Where ID=" & ID
				 RS.Open KeySql, conn, 1, 3
				 If RS.EOF And RS.BOF Then
				  RS.AddNew
				 End If
				  RS("BrandName") = BrandName
				  RS("OrderID")=KS.ChkClng(KS.G("OrderID"))
				  RS("ShowInTop")=KS.ChkClng(KS.G("ShowIntop"))
				  RS("PhotoUrl")=KS.G("PhotoUrl")
				  RS("Recommend")=KS.ChkClng(KS.G("Recommend"))
				  RS("Intro")=Request.Form("Intro")
				  RS("firstAlphabet")=Chr(request("firstAlphabet"))
				  RS("BrandEName")=Request("BrandEname")
				  RS.Update
				  If ID=0 Then
				    RS.MoveLast
				    ID=RS("ID")
					RS.Close
					For K=0 To Ubound(ClassID_Arr)
					 Conn.Execute("Insert Into KS_ClassBrandR(ClassID,BrandID) values('" & ClassID_Arr(K) & "'," & ID & ")")
					Next
                    Call KS.FileAssociation(1003,ID,Request.Form("Intro")&KS.G("PhotoUrl"),0)
					KS.Die "<script> $.dialog.confirm('品牌增加成功,继续添加吗?',function(){ location.href='" & KS.Setting(3) & KS.Setting(89) &"shop/KS.ShopBrand.asp?classid=" & KS.G("ClassID") &"&Action=Add';},function(){location.href='" & KS.Setting(3) & KS.Setting(89) &"shop/KS.ShopBrand.asp';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr="& server.URLEncode(ItemName & "管理中心 >> 品牌管理")+"&ButtonSymbol=Disabled';});</script>"
				  Else
				   
					Conn.Execute("Delete From KS_ClassBrandR Where BrandID=" & ID)
					For K=0 To Ubound(ClassID_Arr)
					 Conn.Execute("Insert Into KS_ClassBrandR(ClassID,BrandID) values('" & ClassID_Arr(K) & "'," & RS("ID") & ")")
					Next
					 Call KS.FileAssociation(1003,ID,Request.Form("Intro")&KS.G("PhotoUrl"),1)
					 KS.die "<script> $.dialog.alert('恭喜，品牌修改成功！', function (){ location.href='" & KS.Setting(3) & KS.Setting(89) &"shop/KS.ShopBrand.asp?Page=" & Page &"';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr="& server.URLEncode(ItemName & "管理中心 >> 品牌管理") &"&ButtonSymbol=Disabled';; });</script>"
					 
				  End If
			End Sub
			
			Sub BrandDel()
			  Dim ID,Page
			  Page=KS.G("Page")
			  ID = KS.G("ID")
			  ID = Replace(ID, " ", "")
			  Conn.Execute("Delete From [KS_UploadFiles] Where ChannelID=1003 And InfoID in(" & ID & ")")
			  Conn.Execute("Delete from [KS_ClassBrand] Where ID in(" & ID & ")")
			  Conn.Execute("Delete From [KS_ClassBrandR] Where BrandID in(" & ID & ")")
			  Response.Redirect "?Page=" & Page
			End Sub
			
			Sub CreateJS()
				Response.Write "<div class='topdashed' style='text-align:center'>"
				Response.Write "<b>生成品牌树型菜单</b>"		
				Response.Write "</div>"
				Response.Write "<form method='POST' action='?Action=SaveCreateJS' id='myform' name='myform'>"
				Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
				Response.Write "  <tr  class='title'>"
				Response.Write "    <td height='22' colspan='6'><strong>品牌树型菜单参数设置</strong> </td>"
				Response.Write "  </tr>"
				Response.Write "  <tr class='tdbg'> "
				Response.Write "    <td width='130' height='25' class='clefttitle'><strong>选择频道：</strong></td>"
				Response.Write "    <td>"
				Response.Write ReturnAllChannel()
				Response.Write "    </td>"
				Response.Write " </tr>"
				Response.Write " <tr class='tdbg' style='display:none'>"
				Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成样式：</strong></td>"
				Response.Write "    <td>"
				Response.Write "      <select name='fsostyle' onchange=""if (this.value==2) document.all.cols.style.display='';else document.all.cols.style.display='none';"">"
				Response.Write "        <option value=1>样式一</option>"
				Response.Write "        <option value=2>样式二</option>"
				Response.Write "      </select>"
				Response.Write "    </td>"
				Response.Write "</tr>"
				Response.Write "<tbody id='cols' style='display:none'>"
				Response.Write " <tr class='tdbg'>"
				Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成列数：</strong></td>"
				Response.Write "    <td>"
				Response.Write "      <input type='text' name='col' value='1' size=""6"">列"
				Response.Write "    </td>"
				Response.Write "</tr>"
				Response.Write "</tbody>"
				Response.Write "<tr class='tdbg'>"
				Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成文件名：</strong></td>"
				Response.Write "    <td>"
				Response.Write "      <input name='JsFileName' type='text' id='JsFileName' value='Brand.js' size='10' maxlength='10'>"
				Response.Write "    </td>"
				Response.Write "  </tr>"
				Response.Write "</table>"
				Response.Write "<br><div style='text-align:center'><input type='submit' name='Submit' value=' 生成树型导航 ' class='button'></div>"
				Response.Write "</form>"
			End Sub
			
			'取得网站的所有频道及其子栏目
			Function ReturnAllChannel()
				  Dim SQL,K,ChannelStr
				   ChannelStr = "<select class='textbox' name=""ChannelID"" style=""width:200;border-style: solid; border-width: 1""><option value='0'>不指定栏目（默认）</option>"
						ChannelStr=ChannelStr & KS.LoadClassOption(5,false)
				   ReturnAllChannel = ChannelStr &"</select>"
			End Function
			
			Sub SaveCreateJS()
				Call KS.WriteTOFile(KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName"), HtreeList)
				Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='ctable'>"
				Response.Write "  <tr class='sort'>"
				Response.Write "    <td height='22' align='center'><strong> 生 成 树 形 导 航 菜 单 </strong></td>"
				Response.Write "  </tr>"
				Response.Write "  <tr class='tdbg'>"
				Response.Write "    <td height='150'>"
				Response.Write "<br><p align='center'><font color=red><b>恭喜您！品牌树形导航菜单成功生成,请按以下提示完成最好操作。</b></font></p>"
				Response.Write "<p><b>将以下代码复制到在模板里要显示的地方。</b></p>"
				Response.Write "<input name='s2' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName") & "&quot;&gt;&lt;/script&gt;' size='80'>&nbsp;<input class=""button"" onClick=""jm_cc('s2')"" type=""button"" value=""复制到剪贴板"" name=""button1"">"
				Response.Write "    </td>"
				Response.Write "  </tr>"
				Response.Write "</table>"
			 %>
			 <script>
			function jm_cc(ob)
			{
				var obj=MM_findObj(ob); 
				if (obj) 
				{
					obj.select();js=obj.createTextRange();js.execCommand("Copy");}
					top.$.dialog.alert('复制成功，粘贴到你要调用的模板里即可!');
				}
				function MM_findObj(n, d) { //v4.0
			  var p,i,x;
			  if(!d) d=document;
			  if((p=n.indexOf("?"))>0&&parent.frames.length)
			   {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
			   }
			  if(!(x=d[n])&&d.all) x=d.all[n];
			  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
			  </script>
			 <%
			End Sub
			
			Function EncodeJS(str)
			EncodeJS = Replace(Replace(Replace(Replace(Replace(str, Chr(10), ""), "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "")
			End Function
			
			Function HtreeList()
			   Dim RS,TreeStr,ID,i,Param,ChannelID
			   ChannelID=KS.S("ChannelID")
			   Param=" Where ChannelID=5"
				If KS.S("ChannelID")="0" Then  
				 Param=Param & " and tj=1"
				Else
				 Param=Param & " and TN='" & ChannelID & "'"
				End If
				Set  RS=Server.CreateObject("ADODB.Recordset")
				RS.Open ("select ID,FolderName,ClassID from KS_Class "  & Param & " Order BY FolderOrder ASC"), Conn, 1, 1
						TreeStr=TreeStr & "document.writeln('<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"">');" & vbcrlf
						Do While Not RS.EOF
						  TreeStr=TreeStr & "document.writeln('<tr>');" & vbcrlf
						  For I=1 To KS.ChkClng(KS.G("Col"))
						   TreeStr = TreeStr & "document.writeln('<td valign=""top"" width=""" & 100 / KS.ChkCLng(KS.G("Col")) & "%"">');" & vbcrlf
						   TreeStr = TreeStr & "document.writeln('<div class=""classtitle"" style=""font-weight:bold""><img src=""" & KS.Setting(3) & "../images/gw12.gif"" align=""absmiddle"">&nbsp;<a href=""" & KS.GetDomain & "shop/showbrand.asp?id=" & rs(2) & """>" & rs(1)& "</a></div>');" & vbnewline 
						   TreeStr = TreeStr & SubList(RS(0),RS(2))
						   TreeStr = TreeStr & "document.writeln('</td>');" & vbcrlf
						   RS.MoveNext
						   If RS.EOF Then Exit For
						  Next
						   TreeStr = TreeStr & "document.writeln('</tr>');" & vbcrlf
						  if rs.eof then exit do
						Loop
						TreeStr =TreeStr & "document.writeln('</table>');" & vbcrlf
						RS.Close:Set RS=Nothing
				HtreeList=TreeStr
			End Function	
			
			Function SubList(TID,ClassID)
			  Dim RS:Set RS=Conn.Execute("select top 20 b.id,brandname from ks_classbrand b inner join ks_classbrandr r on b.id=r.brandid where r.classid ='" & TID & "' order by b.orderid")
			  Dim SQL,I
			  If Not RS.Eof Then
				 SQL=RS.GetRows(-1)
				 SubList="document.writeln('<div class=""list"">"
				 For I=0 To Ubound(SQL,2)
				   SubList=SubList & "<a href=""" & KS.Setting(3) & "shop/brand.asp?id=" & ClassID & "&brandid=" & SQL(0,I) & """ target=""_blank"">" & SQL(1,I) & "</a>&nbsp;"
				   If I <> Ubound(SQL,2) Then SubList=SubList & "<img src=""" & KS.Setting(3) & "../images/nl.gif"" align=""absmiddle"">&nbsp;"
				   
				 Next
				 SubList=SubList & "</div>');"& vbcrlf
			  End IF
			  RS.Close:Set RS=Nothing
			End Function 

End Class
%>
 

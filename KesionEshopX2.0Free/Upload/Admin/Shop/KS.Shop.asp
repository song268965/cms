<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Shop
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Shop
        Private KS,KSCls
		'=====================================定义本页面全局变量=====================================
		Private ID, I, totalPut, Page, RS,ComeFrom
		Private FolderName, CreateDate, TempStr
		Private IsChangedBuy,ChangeBuyNeedPrice,ChangeBuyPresentPrice
		Private IsLimitbuy,LimitBuyPrice,LimitBuyAmount,LimitBuyTaskID
		Private KeyWord, SearchType, StartDate, EndDate, VerificTF, SearchParam
		Private MaxPerPage,UserDefineFieldArr,UserDefineFieldValueStr
		Private T, TitleStr, VerificStr, ShortName, SpecialID,FileName
		Private TypeStr,AttributeStr,HitsByDay, HitsByWeek, HitsByMonth
		Private FolderID,TN, TI, TJ,Action,TemplateID,GroupPrice,ProductID,oTid,OID,RelatedID
		Private ProID, Title, PhotoUrl,BigPhoto, ProIntro,Recommend,IsSpecial,IsTop,IsScore,Score
		Private Popular, Verific,Strip, Comment, Slide,Rolls, KeyWords,ProSpecificat, ProModel,ServiceTerm, ProducerName, TrademarkName, AddDate, Rank, Hits,Unit, TotalNum, AlarmNum,IsDiscount,Price_Member,Price,VIPPrice,BrandID,membernum,visitornum,arrGroupID
		Private CurrPath,UpPowerFlag,AddType
		Private Inputer,ComeUrl,AttributeCart,Weight,DownUrl
		Private SqlStr,Errmsg,Makehtml,FnameType,Tid,Fname,KSRObj
		Private ChannelID,FieldXML,FieldNode,FNode,FieldDictionary
		Private SEOTitle,SEOKeyWord,SEODescript,FreeShipping,WholesalePrice,WholesaleNum,Changes,ChangesUrl
		'=============================================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		If ChannelID=0 Then ChannelID=5
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		Call KSCls.LoadModelField(ChannelID,FieldXML,FieldNode)
		
		If Request("action")="CheckTitle" Then
				Call KSCls.CheckTitle()  
				Exit Sub  
		ElseIf Request("action")="SelectClass" Then
			   Call KSCls.SelectMutiClass()
			   Exit Sub
		End If
		
		If KS.G("page") <> "" Then
			  Page = KS.ChkClng(KS.G("page"))
		Else
			  Page = 1
		End If
		If Instr(Request("action"),"LimitBuy")<>0 then
            If Not KS.ReturnPowerResult(5, "M520008") Then                  '权限检查
			 Call KS.ReturnErr(1, "")   
			 Response.End()
		    End If	
		ElseIf Instr(Request("action"),"Package")<>0 then
            If Not KS.ReturnPowerResult(5, "M520011") Then                  '权限检查
			 Call KS.ReturnErr(1, "")   
			 Response.End()
		    End If	
		end if
		
		Select Case Request("action")
		  case "StockAlarm" StockAlarm:Exit Sub
		  cASE "BatchSaveStock" BatchSaveStock
		  case "Package"    Package:Exit Sub
		  case "AddPackage"  AddPackage:Exit Sub
		  case "PackageSave" PackageSave:Exit Sub
		  case "DelPackage"  DelPackage:Exit Sub
		  case "AddPackPro" AddPackPro:Exit Sub
		  case "Packprolist" Packprolist:Exit Sub
		  case "DelPackagepro" DelPackagepro: Exit Sub
		  case "ChangedBuy" ChangedBuy:Exit Sub
		  case "DelChangeBuy" DelChangeBuy:Exit Sub
		  case "BatchSave" BatchSave
		  case "LimitBuy" LimitBuy:Exit Sub
		  case "DelLimitBuy" DelLimitBuy:Exit Sub
		  case "DelLimitBuyTask" DelLimitBuyTask:Exit Sub
		  case "BatchSaveLimitBuy" BatchSaveLimitBuy: Exit Sub
		  case "AddLimitBuyTask" AddLimitBuyTask:Exit Sub
		  case "LimitBuySave" LimitBuySave:Exit Sub
		  case "LimitBuyProduct" LimitBuyProduct:Exit Sub
		  case "selectProductAddLimitBuy" selectProductAddLimitBuy:Exit Sub
		  case "SetKBXSPrice" SetKBXSPrice:Exit Sub
		  case "BundleSale" BundleSale:Exit Sub
		  case "DelBundleSale" DelBundleSale:Exit Sub
		End Select
		
		'收集搜索参数
		KeyWord   = KS.G("KeyWord")
		SearchType= KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate   = KS.G("EndDate")
		Action     = KS.G("Action")
		ComeFrom   = KS.G("ComeFrom")
		SearchParam = "ChannelID=" & ChannelID
		If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
		If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
		If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
		If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		IF KS.G("Method")="Save" Then
				 Call DoSave()
		Else 
				 Call ShopAdd()
		End If
	   End Sub
	   
	   '超值礼包
	   Sub PackAge()
		 With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'><li class='parent'><span class='child'><i class='icon set'></i>超值礼包管理</span></li> <li class='parent' onclick=""window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('超值礼包管理 >> <font color=red>添加添加礼包</font>')+'&ButtonSymbol=GOSave';location.href='?action=AddPackage';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon gift'></i>添加礼包</span></li></ul>"
			  .Write  "<div class='pageCont2'><table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>序号</td>"
	          .Write "   <td>小图</td>"
	          .Write "   <td>礼包名称</td>"
			  .Write "	 <td>礼包类型</td>"
			  .Write "	 <td>商品数</td>"
			  .Write "   <td>折扣率</td>"
			  .Write "   <td>可选商品</td>"
			  .Write "   <td>添加时间</td>"
			  .Write "   <td>状态</td>"
			  .Write "   <td>操作</td>"
			  .Write " </tr>"
			  .Write "<form name='myform' action='?' method='get'>"
			  .Write "<input type='hidden' name='action' value='BatchSave'>"
			  
			  Dim XML,Node,Param,TotalPages,PhotoUrl
			  MaxPerPage=100
			  Param=" 1=1"
			  SQLStr=KS.GetPageSQL("KS_ShopPackage","id",MaxPerPage,Page,1,Param,"*")
              Set RS = Server.CreateObject("AdoDb.RecordSet")
		      RS.Open SQLStr, conn, 1, 1
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>还没有添加任何礼包!</td>"
					 .Write "</tr>"
			  Else
					totalPut = Conn.Execute("Select count(id) from [KS_ShopPackage] where " & Param)(0)
					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 PhotoUrl=Node.SelectSingleNode("@photourl").text
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@id").text & "<input type='hidden' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
					 .Write "<td align='center' class='splittd'><Img style='margin:2px;border:1px solid #efefef;padding:1px;' src='" &PhotoUrl & "' width='40' height='40' /></td>"
					 .Write "<td class='splittd'>" & Node.SelectSingleNode("@packname").text & "</td>"
					 .Write "<td align='center' class='splittd'>"
					 if Node.SelectSingleNode("@packtype").text="0" then
					  .write "自选礼包"
					 else
					  .Write "<span style='color:red'>特惠礼包</span>"
					 end if
					 .Write "</td>"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@num").text & "</td>"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@discount").text & " 折</td>"
					 .Write "<td align='center' class='splittd'><a href='?Action=AddPackPro&id=" & Node.SelectSingleNode("@id").text & "' title='添加选购品' class='setA'>" & Conn.Execute("Select Count(id) From KS_ShopPackagePro Where packid=" & Node.SelectSingleNode("@id").text)(0) &" 件</a></td>"
					 .Write "<td align='center' class='splittd'>" & FormatDateTime(Node.SelectSingleNode("@adddate").text,2) & "</td>"
					 .Write "<td align='center' class='splittd'>"
					 if Node.SelectSingleNode("@status").text="0" then
					  .write "<span style='color:red'>关闭</span>"
					 else
					  .Write "<span style='color:green'>正常</span>"
					 end if
					 .Write "</td>"
					 .Write "<td align='center' class='splittd'><a href='?Action=AddPackPro&id=" & Node.SelectSingleNode("@id").text & "' class='setA'>添加选购品</a>|<a href='KS.Shop.asp?channelid=5&id="& Node.SelectSingleNode("@id").text&"&Action=DelPackage' onclick=""return(confirm('删除后不可恢复,确定删除吗?'))"" class='setA'>删除</a>|<a href='?Page=" & Page & "&Action=AddPackage&ID=" &Node.SelectSingleNode("@id").text & "' onclick='parent.frames[""BottomFrame""].location.href=""../Post.Asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("超值礼包管理 >> <font color=red>编辑礼包</font>") & "&ButtonSymbol=GOSave"";' class='setA' class='setA'>修改</a>|<a href='../../shop/pack.asp?id=" & Node.SelectSingleNode("@id").text& "' target='_blank' class='setA'>查看</a></td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "</form>"
			 .Write "</table>"
			 Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			 .Write "</div>"
			 RS.Close
			 Set RS=Nothing

		End With	   
	   
	   End Sub
	   
	   '添加超值礼品
	   Sub AddPackage()
	        Dim PackName,PhotoURL,BigPhoto,Num,PackType,Discount,Status,Content,ID,TemplateID
	        Num=2 : Discount=8 : PackType=0 : Status=1 : Content=""
			ID=KS.ChkClng(Request("id"))
			If ID<>0 Then
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From KS_ShopPackage Where ID=" & ID,conn,1,1
			   If Not RS.Eof Then
			     PackName=RS("PackName") : PhotoURL=RS("PhotoUrl") : BigPhoto=RS("BigPhoto") :Num=RS("Num") : PackType=RS("PackType") :Discount=RS("Discount") : Status=RS("Status") :Content=RS("Content"):TemplateID=RS("TemplateID")
			   End If
			   RS.Close:Set RS=Nothing
			End If
	   
	   		 With Response
              .Write"<!DOCTYPE html>"
		      .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js""></script>"
			  .Write EchoUeditorHead
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort"">添加超值礼包</div>"
			  %>
			  <script type="text/javascript">
			  function CheckForm()
			  {
			  if ($('#PackName').val()=='')
				{
				 top.$.dialog.alert('请输入礼包名称!',function(){
				 $('#PackName').focus();});
				 return false;
				}
			  if ($('#Num').val()=='')
				{
				 top.$.dialog.alert('请输入限定商品数!',function(){
				 $('#Num').focus();});
				 return false;
				}
			  if ($('#Discount').val()=='')
				{
				 top.$.dialog.alert('请输入礼包折扣率!',function(){
				 $('#Discount').focus();});
				 return false;
				}
				<%
			  Call LFCls.ShowDiyFieldCheck(FieldXML,1)
			     %>
			  $("#myform").submit();
				
			  }
			  </script>

			 <div class="pageCont2">
			 <form name="myform" id="myform" action="?action=PackageSave" method="post">
					<input type="hidden" value="<%=ID%>" name="id" />
					<input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
                    <dl class="dtable">
					<dd>
					      <div class="firstd">礼包名称：</div>
						  <input type='text' name='PackName' class='textbox' id='PackName' value='<%=PackName%>' size="50" />
						  <font color=red>*</font>
					</dd>
					<dd>
					  <div class="firstd">礼包图片：</div>
					  <dd class="mt10">
					  <div class="firstd">小图</div><input name='PhotoUrl' type='text' id='PhotoUrl' size='50' value='<%=PhotoUrl%>' class='textbox'>
			          <input class="button" type='button' name='Submit' value='选择礼包小图片...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.PhotoUrl);"> <input class="button" type='button' name='Submit' value='远程抓取小图片...' onClick="OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('抓取远程图片')+'&ItemName=小图片&CurrPath=<%=KS.GetUpFilesDir()%>',300,100,window,document.myform.PhotoUrl);">
					  <input class="button"  type='button' name='Submit' value='裁剪...' onClick="if($('#PhotoUrl').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindow(1,'<%=KS.Setting(3)%>',$('#PhotoUrl').val())}"> 

					   </dd>
					   <dd class="mt10">
					   <div class="firstd">大图</div><input name='BigPhoto' type='text' id='BigPhoto' size='50' value='<%=BigPhoto%>' class='textbox'>
			           <input class="button" type='button' name='Submit' value='选择礼包大图片...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.BigPhoto);"> <input class="button" type='button' name='Submit' value='远程抓取大图片...' onClick="OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('抓取远程图片')+'&ItemName=大图片&CurrPath=<%=KS.GetUpFilesDir()%>',300,100,window,document.myform.BigPhoto);">
					   <input class="button"  type='button' name='Submit' value='裁剪...' onClick="if($('#BigPhoto').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindows(1,'<%=KS.Setting(3)%>',$('#BigPhoto').val(),$('#BigPhoto')[0])}"> 
                       </dd>
					   <dd class="mt10">
					   <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?ChannelID=5&UpType=Pic' frameborder=0 scrolling=no width='100%' height='30'></iframe>
					   </dd>
					   
					</dd>
					<dd>
					   <div class="firstd">礼包类型：</div>
					   <input type="radio" name="PackType" value="0"<%if PackType="0" then response.write " checked"%>>自选礼包
					   <input type="radio" name="PackType" value="1"<%if PackType="1" then response.write " checked"%>>特惠礼包
					</dd>
					<dd>
					   <div class="firstd">商 品 数：</div>
					   <input type="text" name="Num" id="Num" class="textbox" style="text-align:center" size="4" value="<%=Num%>" />                 <span>限定礼包内的商品数</span>
					 </dd>
					<dd>
					   <div class="firstd">礼包折扣：</div>
					   <input type="text" name="Discount" class="textbox" id="Discount" style="text-align:center" size="4" value="<%=Discount%>" />折 <span>礼包内的商品按此折扣率计算</span>
					</dd>
					<dd>
					   <div class="firstd">礼包介绍：</div>
					  <%
					  Response.Write EchoEditor("Content",Content,"Basic","96%","220px")
					%>
					</dd>
					
					<dd>
					   <div class="firstd">礼包状态：</div>
					   <input type="radio" name="Status" value="1"<%if Status="1" then response.write " checked"%>>正常
					   <input type="radio" name="Status" value="0"<%if Status="0" then response.write " checked"%>>关闭
					</dd>
					
					<dd>
					 <div class="firstd">绑定模板:</div>
			         <input id='TemplateID' name='TemplateID' readonly maxlength='255' size=50 class='textbox' value='<%=TemplateID%>'>&nbsp;<%=KSCls.Get_KS_T_C("$('#TemplateID')[0]")%>
				      </dd>
				</dl>
			  </form>
			  </div>
			  <%
			 End With
	   End Sub
	   
	   '保存礼包
	   Sub PackageSave()
	     Dim PackName,PhotoURL,BigPhoto,Num,PackType,Discount,Status,Content,ID,TemplateID
		 PackName=KS.G("PackName")
		 PhotoUrl=KS.G("PhotoURL")
		 BigPhoto=KS.G("BigPhoto")
		 Num=KS.ChkClng(KS.G("Num"))
		 Discount=KS.G("Discount")
		 Status=KS.ChkClng(KS.G("Status"))
		 Content=Request.Form("Content")
		 ID=KS.ChkClng(KS.G("ID"))
		 PackType=KS.ChkClng(KS.G("PackType"))
		 TemplateID=KS.G("TemplateID")
		 If Not IsNumeric(Discount) Then
		   Call KS.AlertDoFun("折扣率填写不正确!","history.back(-1);")
		 End If
		 If KS.IsNul(PackName) Then
		  Call KS.AlertDoFun("请输入礼包名称!","history.back(-1);")
		 End If
		 
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_ShopPackage Where ID=" & ID,conn,1,3
		 If RS.Eof Then
		   RS.AddNEW
		   RS("AddDate")=Now
		 End If
		   RS("PackName")=PackName
		   RS("PhotoURL")=PhotoUrl
		   RS("BigPhoto")=BigPhoto
		   RS("Num")     =Num
		   RS("PackType")=PackType
		   RS("Discount")=Discount
		   RS("Status")  =Status
		   RS("Content") =Content
		   RS("TemplateID")=TemplateID
		 RS.Update
		 RS.MoveLast
		 Dim NewID:NewID=RS("ID")
		 RS.Close
		 Set RS=Nothing
		
		 If ID=0 Then
		  '关联上传文件
		  Call KS.FileAssociation(1034,NewID,Content & PhotoUrl & BigPhoto,0)
		  Call KS.ConfirmDoFun("超值礼包添加成功,继续添加吗?","location.href='KS.Shop.asp?action=AddPackage';","$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>超值礼包管理</font>") & "';location.href='KS.Shop.asp?action=Package';")
		 Else
		  '关联上传文件
		  Call KS.FileAssociation(1034,NewID,Content & PhotoUrl & BigPhoto,1)
		  Call KS.AlertDoFun("超值礼包修改成功！","$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>超值礼包管理</font>") & "';location.href='"& KS.G("ComeUrl") & "';")
		 End If
	   End Sub
	   Sub DelPackage()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要删除的礼包"
		ID=KS.FilterIDs(ID)
		Conn.Execute("Delete From KS_ShopPackagePro Where PackID in(" & ID & ")")
		Conn.Execute("Delete From KS_ShopPackage Where ID in(" & ID & ")")
		KS.AlertHintScript "恭喜,已将选中的礼包删除!"
	   End Sub
	   
	   Sub DelLimitBuyTask()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要删除的礼包"
		ID=KS.FilterIDs(ID)
		Conn.Execute("Update KS_Product Set LimitBuyTaskID=0,IsLimitbuy=0 Where LimitBuyTaskID in(" & ID & ")")
		Conn.Execute("Delete From KS_ShopLimitBuy Where ID in(" & ID & ")")
		KS.AlertHintScript "恭喜,限时/限量抢购任务删除成功!"
	   End Sub
	   
	   '将商品加入礼包
	   Sub AddPackPro()
	      Dim ID:id=KS.ChkCLng(Request("id"))
		  Dim RS,PackName,PackType,num
		  If ID=0 Then KS.AlertHintScript "对不起,参数出错啦!"
		  Set RS=Conn.Execute("Select top 1 PackName,PackType,num From KS_ShopPackage Where ID=" & ID)
		  If RS.Eof And RS.Bof Then
		    RS.Close :Set RS=Nothing
			KS.AlertHintScript "对不起,参数出错啦!"
		  End If
		  PackName=RS(0)
		  PackType=RS(1)
		  num=rs(2)
		  RS.Close:Set RS=Nothing
	  	  With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort""><strong>礼包名称:</strong><font color=red>" & PackName & "</font> <strong>礼包类型:</strong><font color=blue>"
			  if PackType="0" then
			   .Write "自选礼包"
			  else
			   .Write "特惠礼包"
			  end if
			  .Write "</font></div>"
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
		<div class="pageCont2">
		<table width="100%" border="0">
		  <tr>
			<td style="text-align:left">
			   <div class="pt10">&nbsp;<strong>快速搜索=></strong></div>
			   <div class="pt10">&nbsp;商品编号: <input type="text" class="textbox" name="proids" id="proids" size='15'> 可留空</div>
			   <div class="pt10">&nbsp;商品名称: <input type="text" class='textbox' name="key"></div>
			   <div class="pt10">&nbsp;所属栏目: <select size='1' name='tid' id='tid' class="textbox"><option value=''>--栏目不限--</option><%=KS.LoadClassOption(ChannelID,false)%></select></div>
			  <div class="pt10">&nbsp;价格范围:
			   <input type='text' name='minPrice' class="textbox" size='5' style='text-align:center' id='minPrice' value='10'> 元
			<= <select name="PriceType" id="PriceType">
			  <option value=0>--不限制--</option>
			  <option value=1>当前零售价</option>
			  <option value=2>商城价</option>
			  <option value=3>原始零售价</option>
			 </select>
			 <= <input type='text' name='maxPrice' class="textbox" size='5' style='text-align:center' id='maxPrice' value='100'> 元
			  
			  </div>
			  <div class="pt10">&nbsp;<input type="button" onClick="getProduct()" value="开始搜索" class="button" name="s1"></div>
			
			</td>
			<form name="myform" id="myform" action="KS.Shop.asp?action=Packprolist&flag=add" method="post" target="packframe">
			<input type="hidden" name="packid" value="<%=ID%>"/>
			<input type="hidden" name="PackType" value="<%=PackType%>"/>
			<input type="hidden" name="num" value="<%=num%>"/>
			<td>
			<div id='keyarea'></div>
			<div class="pt10"><strong>查询到的商品:</strong></div>		
            <div class="pt10">
			 <select name="prolist" size="5" style="width:260px;height:140px" multiple="multiple" id="prolist"></select>
			</div>
			<div class="pt10">
			 <input type="submit" value="将选中的商品加入选购品" class="button">
			</div>
			</td>
			</form>
		  </tr>
		</table>
		 </div>
		 
		 <iframe name="packframe" id="optionBox" src="KS.Shop.asp?action=Packprolist&packid=<%=ID%>" width="100%" height="100%" frameborder="0" scrolling="no"></iframe>
		 <script>
		 $("#optionBox").load(function(){
			var mainheight = $(this).contents().find("body").height();
			$(this).css({height:""+mainheight+"px"});
			$(this).contents().find("body").click(function(){
				var mainheight = $(this).height();
				$("#optionBox").css({height:""+mainheight+"px"});
			});		
		});
		 </script>
	   <%
	      End With
	   End Sub
	   
	   Sub Packprolist()
	      dim packid:packid=KS.ChkClng(KS.S("packid"))
		  dim rs,i
		  if request("flag")="add" then
		   dim proids:proids=ks.s("prolist")
		   dim PackType:PackType=ks.chkclng(ks.s("PackType"))
		   dim num:num=KS.Chkclng(ks.s("num"))
		   if proids<>"" then
		      proids=split(KS.FilterIds(proids),",")
			  if packtype=1 then
			   if ubound(proids)+1>num then
			    Call KS.AlertDoFun("对不起,这个特惠礼包只能添加" & num & "件商品!","history.back();")
			   end if
			   dim nownum:nownum=ks.chkclng(conn.execute("select count(id) from KS_ShopPackagePro where packid=" &packid)(0))
			   if ubound(proids)+1+nownum>num then
			     Call KS.AlertDoFun("对不起,这个特惠礼包只能添加" & num & "件商品,请先删除再添加!","history.back();")
			   end if
			  end if
			  for i=0 to ubound(proids)
			    set rs=server.createobject("adodb.recordset")
				rs.open "select top 1 * from KS_ShopPackagePro where packid=" & packid & " and proid=" & proids(i),conn,1,3
				if rs.eof then
				 rs.addnew
				 rs("packid")=packid
				 rs("proid")=proids(i)
				 rs.update
				end if 
				rs.close
				set rs=nothing
			  next
			   Call KS.Alert("恭喜,加入成功!","Shop/KS.Shop.asp?action=Packprolist&packid=" & packid)
		   else
		      response.write "<script>top.$.dialog.alert('您没有选择要加入该礼包的商品!');</script>"
		   end if
		  end if
		  
	  	  With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select a.id,a.title,a.price,a.Price_Member,a.photourl,b.id as packproid from ks_product a inner join KS_ShopPackagePro b on a.id=b.proid where b.PackID=" &packid & " and a.verific=1 order by a.id desc",conn,1,1
			  
			  MaxPerPage=20
              dim TotalPages,xml,node,photourl
			  
			  .Write "<div class='pageCont2'><table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>选择</td>"
	          .Write "   <td>小图</td>"
	          .Write "   <td>商品名称</td>"
			  .Write "	 <td>当前价格</td>"
			  .Write "	 <td>商城价</td>"
			  .Write "   <td>操作</td>"
			  .Write " </tr>"
			  .Write "<form name=""myform"" id=""myform"" action=""ks.shop.asp"" method=""post"">"
			  .Write "<input type='hidden' name='action' id='action' value='DelPackagepro'/>"
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>该礼包下还没有添加选购的商品!</td>"
					 .Write "</tr>"
			  Else
					totalPut = RS.Recordcount

					if Page > TotalPages then Page=TotalPages
					If TotalPages > 1 then Rs.Move (Page - 1) * maxperpage
					
					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 PhotoUrl=Node.SelectSingleNode("@photourl").text
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'><input type='checkbox' name='id' value='" & Node.SelectSingleNode("@packproid").text & "'></td>"
					 .Write "<td align='center' class='splittd'><Img style='border:1px solid #efefef;padding:1px;margin:2px' src='" &PhotoUrl & "' width='40' height='40' /></td>"
					 .Write "<td class='splittd'>" & Node.SelectSingleNode("@title").text & "</td>"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@price").text & " 元</td>"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@price_member").text & " 元</td>"
					 .Write "<td align='center' class='splittd'><a href='?Action=DelPackagepro&id=" & Node.SelectSingleNode("@packproid").text & "' onclick=""return(confirm('确定移除吗?'));""  class='setA'>移除</a>|<a href='javascript:void(0)' onclick='parent.parent.frames[""BottomFrame""].location.href=""../Post.Asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("编辑" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & ID & """;parent.location.href=""KS.Shop.asp?Page=" & Page & "&Action=Edit&ID=" &Node.SelectSingleNode("@id").text & """;' class='setA'>修改</a></td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			 .Write "<td colspan=10 class='operatingBox'>&nbsp;<label><input id=""chkAll"" onClick=""CheckAll(this.form)"" type=""checkbox"" value=""checkbox""  name=""chkAll"">全选</label> <input type='submit' value='批量移除选中的选购品' class='button'> "
			 .Write "</tr>"
			 .Write "</form>"
			 .Write "</table>"
			 
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			  .Write "</div>"
			 RS.Close
			 Set RS=Nothing
			   
			   
			   
			   
	       End With
	   End Sub
	   
	   '将产品从移除
	   Sub DelPackagepro()
	    dim id:id=KS.G("ID")
		IF KS.IsNul(id) Then
		  KS.AlertHintScript "请选择要移除的商品!"
		End If
		id=KS.FilterIds(id)
		Conn.Execute("Delete From KS_ShopPackagePro Where ID in(" & id & ")")
		KS.AlertHintScript "恭喜,选购品移除成功!"
	   End Sub
	   
	   '捆绑销售商品管理
	   Sub BundleSale()
            If Not KS.ReturnPowerResult(5, "M520009") Then                  '权限检查
			 Call KS.ReturnErr(1, "")   
			 Response.End()
		    End If		   
	       With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort"">捆绑销售的商品管理</div>"
			  .Write  "<div class='pageCont2'><table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>序号</td>"
	          .Write "   <td>小图</td>"
	          .Write "   <td>商品名称</td>"
			  .Write "	 <td>当前零售价</td>"
			  .Write "	 <td>商城价</td>"
			  .Write "   <td>添加时间</td>"
			  .Write "   <td>捆绑商品数</td>"
			  .Write "   <td>操作</td>"
			  .Write " </tr>"
			  .Write "<form name='myform' action='?' method='get'>"
			  .Write "<input type='hidden' name='action' value='BatchSave'>"
			  
			  Dim XML,Node,Param,TotalPages,PhotoUrl
			  MaxPerPage=100
			  Param=" ID in(select distinct(proid) from KS_ShopBundleSale) "
			  SQLStr=KS.GetPageSQL("KS_Product","id",MaxPerPage,Page,1,Param,"*")
		      Set RS = Server.CreateObject("AdoDb.RecordSet")
		      RS.Open SQLStr, conn, 1, 1
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>没有捆绑销售的商品!</td>"
					 .Write "</tr>"
			  Else
					totalPut = Conn.Execute("select count(1) from KS_Product where " & Param)(0)
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if Page > TotalPages then Page=TotalPages
					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 PhotoUrl=Node.SelectSingleNode("@photourl").text
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@id").text & "<input type='hidden' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
					 .Write "<td align='center' class='splittd'><Img src='" &PhotoUrl & "' width='40' height='40' /></td>"
					 .Write "<td class='splittd'>" & Node.SelectSingleNode("@title").text & "</td>"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@price").text & " 元</td>"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@price_member").text & " 元</td>"
					 .Write "<td align='center' class='splittd'>" & FormatDateTime(Node.SelectSingleNode("@adddate").text,2) & "</td>"
					 .Write "<td align='center' class='splittd'><font color=red>" &conn.execute("select count(1) from KS_ShopBundleSale where proid=" &Node.SelectSingleNode("@id").text &"")(0) & "</font> 件</td>"
					 .Write "<td align='center' class='splittd'><a href='?Action=DelBundleSale&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('确定移除所有捆绑定吗?'));""  class='setA'>移除</a>|<a href='../System/KS.ItemInfo.asp?channelid=5&id="& Node.SelectSingleNode("@id").text&"&Action=Delete' onclick=""return(confirm('删除后不可恢复,确定删除吗?'))"" class='setA'>永久删除</a>|<a href='?Page=" & Page & "&Action=Edit&ID=" &Node.SelectSingleNode("@id").text & "' onclick='parent.frames[""BottomFrame""].location.href=""../Post.Asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("编辑" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & ID & """;' class='setA'>修改</a></td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			' .Write "<td colspan=10>&nbsp;<input type='submit' value='批量设置换购价' class='button'> "
			 .Write "</tr>"
			 .Write "</form>"
			 .Write "</table>"
			 
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			 .Write "</div>"
			 RS.Close
			 Set RS=Nothing

		End With
	   End Sub
	   
	   '移除捆绑销售的商品
	   Sub DelBundleSale()
	     Dim ID:ID=KS.ChkClng(Request("id"))
		 if id<>0 then
		   conn.execute("delete from KS_ShopBundleSale where proid=" & id)
		 end if
		 call KS.AlertHintScript("恭喜,移除成功!")
	   End Sub
	   
	   '限时抢购商品管理
	   Sub LimitBuy()
	     With Response
			   .Write"<!DOCTYPE html>"
		      .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  %>
			  <script type="text/javascript">
			  function addLimitBuyTask(){
			  window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('限时/限量抢购任管理 >> <font color=red>添加抢购任务</font>')+'&ButtonSymbol=GOSave';
			  location.href='?action=AddLimitBuyTask';
			  }
			  function addLimitBuy(){	
				location.href='KS.Shop.asp?addtype=limitbuy&ChannelID=<%=channelid%>&Action=Add';
                $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ChannelID=<%=channelid%>&OpStr='+escape("添加商品")+'&ButtonSymbol=AddInfo';
			}</script>
			  <%
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class='topdashed sort'>限时/限量抢购任管理</div>"
			  .Write  "<div class='pageCont2'><table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>序号</td>"
	          .Write "   <td>任务名称</td>"
			  .Write "	 <td>任务类型</td>"
			  .Write "   <td>活动时间</td>"
			  .Write "   <td>最后付款时间</td>"
			  .Write "   <td>状态</td>"
			  .Write "   <td>操作</td>"
			  .Write " </tr>"
			  .Write "<form name='myform' action='?' method='get'>"
			  .Write "<input type='hidden' name='action' value='BatchSave'>"
			  
			  Dim XML,Node,Param,TotalPages,PhotoUrl
			  MaxPerPage=100
			  Param=" 1=1"
			  SQLStr=KS.GetPageSQL("KS_ShopLimitBuy","id",MaxPerPage,Page,1,Param,"*")
              Set RS = Server.CreateObject("AdoDb.RecordSet")
		      RS.Open SQLStr, conn, 1, 1
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>还没有添加任何限时/限量任务!</td>"
					 .Write "</tr>"
			  Else
					totalPut = Conn.Execute("Select count(id) from [KS_ShopLimitBuy] where " & Param)(0)
					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@id").text & "<input type='hidden' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
					 .Write "<td class='splittd'>" & Node.SelectSingleNode("@taskname").text & "</td>"
					 .Write "<td align='center' class='splittd'>"
					 if Node.SelectSingleNode("@tasktype").text="1" then
					  .write "<span style='color:green'>限时抢购</span>"
					 else
					  .Write "<span style='color:blue'>限量抢购</span>"
					 end if
					 .Write "</td>"
					 .Write "<td align='center' class='splittd'>从" & Node.SelectSingleNode("@limitbuybegintime").text & "至<br/>" &Node.SelectSingleNode("@limitbuyendtime").text & "止</td>"
					 .Write "<td align='center' class='splittd'>下单后<font color=red>" &Node.SelectSingleNode("@limitbuypaytime").text & "</font>小时内</td>"
					 .Write "<td align='center' class='splittd'>"
					 if Node.SelectSingleNode("@status").text="0" then
					  .write "<span style='color:red'>关闭</span>"
					 else
					  .Write "<span style='color:green'>正常</span>"
					 end if
					 .Write "</td>"
					 .Write "<td align='center' class='splittd'><a href='?taskType=" & Node.SelectSingleNode("@tasktype").text & "&Action=LimitBuyProduct&id=" & Node.SelectSingleNode("@id").text & "' class='setA'>活动商品(<font color=red>" & conn.execute("select count(id) from ks_product where limitbuytaskid=" & node.selectsinglenode("@id").text)(0) & "</font>)</a> | <a href='KS.Shop.asp?channelid=5&id="& Node.SelectSingleNode("@id").text&"&Action=DelLimitBuyTask' onclick=""return(confirm('删除后不可恢复,确定删除吗?'))"" class='setA'>删除</a> | <a href='?Page=" & Page & "&Action=AddLimitBuyTask&ID=" &Node.SelectSingleNode("@id").text & "' onclick='parent.frames[""BottomFrame""].location.href=""../Post.Asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("超值礼包管理 >> <font color=red>编辑礼包</font>") & "&ButtonSymbol=GOSave"";' class='setA'>修改</a> </td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "</form>"
			 .Write "<tr><td height='40' colspan=10 class='pt10'>&nbsp;<input type='button' class='button' value='添加抢购任务' onclick=""addLimitBuyTask()""/>&nbsp;<input type='button' class='button' value='添加抢购商品' onclick=""addLimitBuy()""/></td></tr>"
			 .Write "</table>"
			 Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			 .Write "</div>"
			 RS.Close
			 Set RS=Nothing

		End With
	 End Sub
	 
	 '添加限时抢购
	 Sub AddLimitBuyTask()
	  Dim TaskName,LimitBuyBeginTime,LimitBuyEndTime,LimitBuyPayTime,TaskType,Intro,AddDate,ID,TemplateID,Status
	        Status=1 :Intro="":TaskType=1:LimitBuyBeginTime=Now:LimitBuyEndTime=Now+10:LimitBuyPayTime=48
			ID=KS.ChkClng(Request("id")):AddDate=Now
			If ID<>0 Then
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From KS_ShopLimitBuy Where ID=" & ID,conn,1,1
			   If Not RS.Eof Then
			     TaskName=RS("TaskName") : Status=RS("Status") :Intro=RS("Intro"):TemplateID=RS("TemplateID") : TaskType=RS("TaskType")
				 LimitBuyBeginTime=RS("LimitBuyBeginTime") : LimitBuyEndTime=RS("LimitBuyEndTime") : LimitBuyPayTime=rs("LimitBuyPayTime") 
			   End If
			   RS.Close:Set RS=Nothing
			End If
	   
	   		 With Response
              .Write"<!DOCTYPE html>"
		      .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>" &vbcrlf
			  .Write EchoUeditorHead
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort"">"
			  If ID<>0 Then  .Write "修改"  Else .Write "添加"
			  .Write "限时限量抢购任务</div>"
			  %>
			  <script type="text/javascript">
			  function CheckForm()
			  {
			  if ($('#TaskName').val()=='')
				{
				 top.$.dialog.alert('请输入任务名称!',function(){
				 $('#TaskName').focus();});
				 return false;
				}
			  
			  $("#myform").submit();
				
			  }
			  </script>

			  <div class="pageCont2">
				  <form name="myform" id="myform" action="?action=LimitBuySave" method="post">
				<dl class="dtable">
					<input type="hidden" value="<%=ID%>" name="id" />
					<input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
					<dd>
					  <div>限时/限量任务名称：</div>
						  <input type='text' name='TaskName' id='TaskName' class='textbox' value='<%=TaskName%>' size="50" />
						  <font color=red>*</font> <span>如:7天限时抢购活动</span>
					</dd>
					<dd>
					  <div>任务类型：</div>
					  <label><input type='radio' name='TaskType' onClick="$('#limitbuytime').show();" value='1'<%if tasktype="1" then response.write " checked"%>>限时抢购</label>
					  <label>
					  <input type='radio' name='TaskType' onClick="$('#limitbuytime').hide();" value='2'<%if tasktype="2" then response.write " checked"%>>限量抢购</label>
				    </dd>
					<dd <%if tasktype="2" then response.write " style='display:none'" %> id='limitbuytime'>
					  <div>抢购时间限制：</div>
					  <input type='text' class='textbox Wdate' onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  name='LimitBuyBeginTime' id='LimitBuyBeginTime' value='<%=LimitBuyBeginTime%>' size='20'>至<input onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  class='textbox Wdate' type='text' name='LimitBuyEndTime' id='LimitBuyEndTime' value='<%=LimitBuyEndTime%>' size='20'>  <font color=red>在此时间内才能享受抢购价</font>
					</dd>
					<dd>
					  <div>最迟付款时间：</div>
					   下单后<input type='text' class='textbox' name='LimitBuyPayTime' id='LimitBuyPayTime' style='text-align:center' value='<%=LimitBuyPayTime%>' size='6'>小时内没有付款,视为抢购无效。<span>如果不限制请录入"0"。</span>
					</dd>
				
					
					<dd>
					  <div>任务介绍：</div>
					  <%
					  Response.Write EchoEditor("Intro",Intro,"Basic","96%","220px")
			     	%>
					</dd>
					<dd>
					  <div>任务添加时间：</div>
					   <input type='text' class='textbox Wdate' onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" name='AddDate' id='AddDate' value='<%=AddDate%>' size='50'><span>格式：<%=now%></span>
					</dd>
					<dd>
					  <div>任务状态：</div>
					   <input type="radio" name="Status" value="1"<%if Status="1" then response.write " checked"%>>正常
					   <input type="radio" name="Status" value="0"<%if Status="0" then response.write " checked"%>>关闭
					</dd>
					<dd style="display:none">
					 <div>绑定模板:</div>
			         <input id='TemplateID' name='TemplateID' readonly maxlength='255' size=30 class='textbox' value='<%=TemplateID%>'>&nbsp;<%=KSCls.Get_KS_T_C("$('#TemplateID')[0]")%>
				      </dd>
				</dl>
				</form>
			  </div>
				</div>

			  <%
			 End With
	 End Sub
	 '保存限时抢购任务
	 Sub LimitBuySave()
	     Dim TaskName,LimitBuyBeginTime,LimitBuyEndTime,LimitBuyPayTime,TaskType,Intro,AddDate,ID,TemplateID,Status
		 TaskName=KS.G("TaskName")
		 LimitBuyBeginTime= KS.S("LimitBuyBeginTime")
		 LimitBuyEndTime = KS.S("LimitBuyEndTime")
		 LimitBuyPayTime=KS.ChkClng(KS.S("LimitBuyPayTime"))
		 If Not IsDate(LimitBuyBeginTime) Then LimitBuyBeginTime=now
		 If Not IsDate(LimitBuyEndTime) Then LimitBuyEndTime=Now+10


		 Status=KS.ChkClng(KS.G("Status"))
		 Intro=Request.Form("Intro")
		 ID=KS.ChkClng(KS.G("ID"))
		 TaskType=KS.ChkClng(KS.G("TaskType"))
		 TemplateID=KS.G("TemplateID")
		 AddDate=KS.G("AddDate")
		 If Not IsDate(AddDate) Then AddDate=Now
		 
		 If KS.IsNul(TaskName) Then
		  KS.Echo ("<script>alert('请输入任务名称!');history.back(-1);</script>")
		 End If
	     
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_ShopLimitBuy Where ID=" & ID,conn,1,3
		 If RS.Eof Then
		   RS.AddNEW
		 End If
		   RS("TaskName")=TaskName
		   RS("LimitBuyBeginTime")=LimitBuyBeginTime
		   RS("LimitBuyEndTime")=LimitBuyEndTime
		   RS("LimitBuyPayTime") =LimitBuyPayTime
		   RS("TaskType")=TaskType
		   RS("Intro")=Intro
		   RS("Status")  =Status
		   RS("AddDate") =AddDate
		   RS("TemplateID")=TemplateID
		 RS.Update
		 RS.MoveLast
		 Dim NewID:NewID=RS("ID")
		 RS.Close
		 Set RS=Nothing
		
		 Response.Write "<script src='../../ks_inc/jquery.js'></script>" 
		 If ID=0 Then
		  Response.Write "<script>if (confirm('限时/限量抢购任务添加成功,继续添加吗?')){location.href='KS.Shop.asp?action=AddLimitBuyTask';}else{$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>抢购任务管理</font>") & "';location.href='KS.Shop.asp?action=LimitBuy';}</script>"
		 Else
		  Response.Write "<script>alert('限时/限量抢购任务修改成功！');$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("商城系统 >> <font color=red>抢购任务管理</font>") & "';location.href='"& KS.G("ComeUrl") & "';</script>"
		 End If
	End Sub
	 
	 Sub selectProductAddLimitBuy()
	  Dim LimitBuyTaskID:LimitBuyTaskID=KS.ChkClng(Request("LimitBuyTaskID"))
	  Dim taskType:taskType=KS.ChkClng(Request("taskType"))
	  If Request("flag")="save" then
	    Dim ID:ID=KS.FilterIds(KS.S("ID"))
		If KS.IsNul(ID) Then
		 KS.AlertHintScript "对不起，您没有选择要加入抢购的商品!"
		End If
		Dim IDArr:IDArr=Split(ID,",")
		Dim i,LimitBuyPrice,LimitBuyNum
		For i=0 To Ubound(IDArr)
		   LimitBuyPrice=KS.S("LimitBuyPrice"&IDArr(i))
		   LimitBuyNum=KS.S("LimitBuyNum"&IDArr(i))
		   If Not Isnumeric(LimitBuyPrice) Then LimitBuyPrice=0
		   If Not Isnumeric(LimitBuyNum) Then LimitBuyNum=0
		   Conn.Execute("Update KS_Product Set IsLimitbuy=" & taskType & ",LimitBuyTaskID=" & LimitBuyTaskID&",LimitBuyPrice=" & LimitBuyPrice&",LimitBuyAmount=" & LimitBuyNum&" where id=" & KS.ChkClng(IDArr(i)))
		Next
		KS.Die "<script>alert('恭喜，成功将商品加入抢购任务！');top.box.close();</script>"
	  End If
	  With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "</head>"
	  End With
	  %>
	        <body>
			<script>
			 function doSearch(){
			  var proids=$("#proids").val();
			  var title=$("#title").val();
			  if (proids=='' && title==''){
			   alert('商品编号或商品名称必须填一个！');
			   $("#title").focus();
			   return false
			  }
			  $("#prolist").empty();
			  $(parent.document).find("#ajaxmsg").toggle(true);
			   var url='../../shop/ajax.getdate.asp';
			   $.get(url,{action:"Shop_SearchProduct",proids:proids,title:escape(title)},function(d){
				    $(parent.document).find("#ajaxmsg").toggle(false);
					 if (d==''){
					   alert('对不起，找不到商品！');
					 }else{
						 var darr=d.split('§');
						 for(var i=0;i<darr.length;i++){
						   var rrr=darr[i].split('◇');
						   addRecord(rrr[0],rrr[1],rrr[2]);
						 }
					 }
				   });
			  
			 }
			 function addRecord(id,title,price){
			    var str="<tr id='tr"+id+"'><td class='splittd'>"+id+
				         "<input type='hidden' name='id' value='"+id+"'/>"+
				         "</td><td class='splittd'>"+title+
				         "</td><td class='splittd' style='text-align:center'><input type='text' class='textbox' style='width:50px' name='limitbuyprice"+id+"' value='"+price+"'/>元"+
				         "</td><td class='splittd' style='text-align:center'><input type='text' class='textbox' style='width:30px' name='limitbuynum"+id+"' value='100'/>"+
				         "</td><td class='splittd' style='text-align:center'><a href='javascript:;' onclick='del("+id+");'>删除</a>"+
						 "</td></tr>";
				$("#prolist").append(str);		 
			 }
			 function del(id){
			  $("#tr"+id).remove();
			 }
			</script>
			<br/>
			 <table width="98%" align="center">
			  <tr>
			   <td>
			   <strong>第一步：搜索商品</strong><br/>
			   商品编号: <input type="text" class="textbox" name="proids" id="proids" size='15'> 或 商品名称: <input type="text" class='textbox' name="title" id="title"/>
			 <br/>
			 <input type="button" value="搜索商品" onClick="doSearch()" class="button"/>
			   </td>
			  </tr>
			  <tr>
			    <td style="padding-top:10px">
				<strong>第二步：确定加入抢购</strong><br/>
				   <form name="myform" action="KS.Shop.asp" method="post">
				   <input type="hidden" name="action" value="selectProductAddLimitBuy"/>
				   <input type="hidden" name="LimitBuyTaskID" value="<%=LimitBuyTaskID%>"/>
				   <input type="hidden" name="taskType" value="<%=taskType%>"/>
				   <input type="hidden" name="flag" value="save"/>
					  <table border='0' width='100%'>
					  <tr class='sort'>
						<td>ID</td>
						<td>商品名称</td>
						<td>抢购价</td>
						<td>抢购数量</td>
						<td>操作</td>
					  </tr>
					   <tbody id="prolist">
					     <tr>
							<td colspan='5' class='splittd' style='text-align:center'>请先搜索商品!</td>
						  </tr>
					   </tbody>
					  </table>
				     <input type="submit" value="确定加入" class="button"/>
				   </form>
				</td>
			  </tr>
			  </table>
			</body>
			</html>
	  <%
	 End Sub
	 
	 Sub LimitBuyProduct()
	          dim id:id=ks.chkclng(request("id"))
			  dim rs:set rs=server.createobject("adodb.recordset")
			  rs.open "select top 1 * from ks_shoplimitbuy where id="& id,conn,1,1
			  if rs.eof and rs.bof then 
			    rs.close : set rs=nothing
				ks.alerthintscript "出错了!"
			  end if
			  dim taskname:taskname=rs("taskname")
			  dim tasktype:tasktype=rs("tasktype")
			  rs.close: set rs=nothing
		 With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  %>
			  <script type="text/javascript">
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
			function addLimitBuyProduct(){	
				location.href='KS.Shop.asp?taskType=<%=KS.S("TaskType")%>&addtype=limitbuy&LimitBuyTaskID=<%=ID%>&ChannelID=<%=ChannelID%>&Action=Add';
                $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ChannelID=<%=channelid%>&OpStr='+escape("添加商品")+'&ButtonSymbol=AddInfo';
			}
			function selectProduct(){	
				top.openWin('选择已添加的商品加入抢购','shop/KS.Shop.asp?action=selectProductAddLimitBuy&taskType=<%=KS.S("TaskType")%>&LimitBuyTaskID=<%=ID%>&rnd='+Math.random(),true,800,500); 
			}
			  </script>
			  <%
			  
			  
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort"" style='text-align:left'>&nbsp;&nbsp;<a href='KS.Shop.asp?action=LimitBuy&channelid=5'>返 回</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;管理抢购任务[<font color=red>" & taskname & "</font>]下的商品</div>"

			  .Write "<div class='pageCont2'><form name='myform' action='KS.Shop.asp' method='post'>"
			  .Write "<input type='hidden' name='action' id='action' value='BatchSaveLimitBuy'>"
			  .Write "<input type='hidden' name='f' value='" & KS.G("F") & "'>"
			  .Write  "<table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>序号</td>"
	          .Write "   <td>小图</td>"
	          .Write "   <td>商品名称</td>"
			  .Write "	 <td>抢购价</td>"
			  .Write "	 <td>抢购数量</td>"
			  .Write "   <td>已被抢购</td>"
			  .Write "   <td>操作</td>"
			  .Write " </tr>"
			  
			  Dim XML,Node,Param,TotalPages
			  MaxPerPage=100
			  Param=" LimitBuyTaskID=" & ID
			  SQLStr=KS.GetPageSQL("KS_Product","id",MaxPerPage,Page,1,Param,"*")
		      Set RS = Server.CreateObject("AdoDb.RecordSet")
		      RS.Open SQLStr, conn, 1, 1
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>没有允许换购的商品!</td>"
					 .Write "</tr>"
			  Else
					totalPut = Conn.Execute("Select count(id) from [KS_Product] where " & Param)(0)

					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 PhotoUrl=Node.SelectSingleNode("@photourl").text
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../../images/nopic.gif"
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'><input type='checkbox' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
					 .Write "<td align='center' class='splittd'><Img onerror=""this.src='../../images/nopic.gif';"" src='" &PhotoUrl & "' width='40' height='40' /></td>"
					 .Write "<td class='splittd'><a href='../../item/show.asp?m=5&d=" & Node.SelectSingleNode("@id").text& "' target='_blank'>" & Node.SelectSingleNode("@title").text & "</a><br/>商城价:" & Node.SelectSingleNode("@price_member").text & "元</td>"
					 .Write "<td align='center' class='splittd'><input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' name='limitbuyprice" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@limitbuyprice").text & "' size='5' style='text-align:center'> 元</td>"
					 .Write "<td align='center' class='splittd'><input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' name='limitbuyamount" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@limitbuyamount").text & "' size='5' style='text-align:center'></td>"
					 
					
					 
					 .Write "<td align='center' class='splittd'><font color=red>" &conn.execute("select count(1) from ks_orderitem where proid=" &Node.SelectSingleNode("@id").text &" and islimitbuy<>0 and limitbuytaskid=" & id & "")(0) & "</font> 件</td>"
					 .Write "<td align='center' class='splittd'><a href='?Action=DelLimitBuy&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('确定移除吗?'));"">移除</a> | <a href='../System/KS.ItemInfo.asp?channelid=5&id="& Node.SelectSingleNode("@id").text&"&Action=Delete' onclick=""return(confirm('删除后不可恢复,确定删除吗?'))"">永久删除</a> | <a href='?Page=" & Page & "&Action=Edit&ID=" &Node.SelectSingleNode("@id").text & "' onclick='parent.frames[""BottomFrame""].location.href=""../Post.Asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("编辑" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & ID & """;'>修改</a></td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "<tr onmouseout=""this.className=''"" onmouseover=""this.className=''"">"
			 .Write "<td colspan=10 class='operatingBox'><input id=""chkAll"" onClick=""CheckAll(this.form)"" type=""checkbox"" value=""checkbox""  name=""chkAll"">全选&nbsp;&nbsp;<input type='submit' value='批量移除' onclick=""$('#action').val('DelLimitBuy');"" class='button'>&nbsp;<input type='submit' value='批量设置抢购商品信息' onclick=""$('#action').val('BatchSaveLimitBuy');"" class='button'> <input type='button' value='添加抢购商品' class='button' onclick=""addLimitBuyProduct();""> <input type='button' value='选择已添加商品加入抢购' class='button' onclick='selectProduct()'/>"
			 .Write "</tr>"
			 .Write "</table>"
			 .Write "</form>"
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			 .Write "</div>"
			 RS.Close
			 Set RS=Nothing

		End With
	   End Sub
	   '移出抢购品
	   Sub DelLimitBuy()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要移除的商品"
		ID=KS.FilterIDs(ID)
		Conn.Execute("Update KS_Product Set IsLimitBuy=0,LimitBuyTaskID=0 Where ID in(" & ID & ")")
		KS.AlertHintScript "恭喜,已将选中的商品移出!"
	   End Sub
	   
	   '批量设置换购价格
	   Sub BatchSaveLimitBuy()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要设置的商品"
		Dim i,IDArr,LimitBuyBeginTime,LimitBuyEndTime,LimitBuyPayTime
		IDArr=Split(KS.FilterIDs(ID),",")
		For i=0 to ubound(idArr)
		  Conn.Execute("Update KS_Product Set LimitBuyPrice=" & KS.G("LimitBuyPrice"&idArr(i)) & ",LimitBuyAmount=" & KS.ChkClng(KS.G("LimitBuyAmount"&idArr(i))) & " Where ID=" & Idarr(i))
		Next
		KS.AlertHintScript "恭喜,批量设置抢购商品成功!"
	   End Sub
	   
	   '批量设置库存
	   Sub BatchSaveStock()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要补货的商品"
		Dim i,IDArr
		IDArr=Split(KS.FilterIDs(ID),",")
		For i=0 to ubound(idArr)
		  Conn.Execute("Update KS_Product Set alarmnum=" & KS.ChkClng(KS.G("alarmnum"&idArr(i))) & ",totalNum=" & KS.ChkClng(KS.G("totalnum"&idArr(i))) &" Where ID=" & Idarr(i))
		Next
		KS.AlertHintScript "恭喜,批量补货成功!"
	   End Sub

	   
	   '库存报警管理
	   Sub StockAlarm()
	        If Not KS.ReturnPowerResult(5, "M520012") Then                  '权限检查
			 Call KS.ReturnErr(1, "")   
			 Response.End()
		    End If		   

		 With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  %>
			  <script type="text/javascript">
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
			  </script>
			  <%
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort"">以下商品库存量已不足，请及时补货</div>"
			  .Write  "<div class='pageCont2'><table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>序号</td>"
	          .Write "   <td>小图</td>"
	          .Write "   <td>商品名称</td>"
			  .Write "	 <td>当前库存量</td>"
			  .Write "   <td>报警库存量</td>"
			  .Write " </tr>"
			  .Write "<form name='myform' action='?' method='get'>"
			  .Write "<input type='hidden' name='action' value='BatchSaveStock'>"
			  
			  Dim XML,Node,Param,TotalPages,PhotoUrl
			  MaxPerPage=100
			  Param=" totalnum<=AlarmNum"
			  SQLStr=KS.GetPageSQL("KS_Product","id",MaxPerPage,Page,1,Param,"*")
		      Set RS = Server.CreateObject("AdoDb.RecordSet")
		      RS.Open SQLStr, conn, 1, 1
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>还没有发现库存量不足的商品!</td>"
					 .Write "</tr>"
			  Else
					totalPut = Conn.Execute("Select count(id) from [KS_Product] where " & Param)(0)
					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 PhotoUrl=Node.SelectSingleNode("@photourl").text
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../../images/nopic.gif"
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@id").text & "<input type='hidden' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
					 .Write "<td align='center' class='splittd'><Img src='" &PhotoUrl & "' width='40' height='40' /></td>"
					 .Write "<td class='splittd'>" & KS.Gottopic(Node.SelectSingleNode("@title").text,40) & "</td>"
					 .Write "<td align='center' class='splittd'><input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' name='totalnum" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@totalnum").text & "' size='5' style='text-align:center'> " & Node.SelectSingleNode("@unit").text & "</td>"
					 .Write "<td align='center' class='splittd'><input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' name='alarmnum" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@alarmnum").text & "' size='5' style='text-align:center'> " & Node.SelectSingleNode("@unit").text &"</td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "<tr onmouseout=""this.className=''"" onmouseover=""this.className=''"">"
			 .Write "<td colspan=10 class='pt10'>&nbsp;<input type='submit' value='保存批量补货' class='button'> "
			 .Write "</tr>"
			 .Write "</form>"
			 .Write "</table>"
			 
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			 RS.Close
			 Set RS=Nothing

		End With
		%>
		<div class="attention" style="clear:both">
		<font color=red>说明：这里只显示库存量少于等于库存量报警的商品，为保证客户顺利购买商品，请对库存量不多的商品及时补货。</font>
		</div></div>
		<%
	   End Sub
	   
	   
	   '换购品管理
	   Sub ChangedBuy()
	        If Not KS.ReturnPowerResult(5, "M520010") Then                  '权限检查
			 Call KS.ReturnErr(1, "")   
			 Response.End()
		    End If		   

		 With Response
			  .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  %>
			  <script type="text/javascript">
			  function regInput(obj, reg, inputStr){
					var docSel = document.selection.createRange()
					if (docSel.parentElement().tagName != "INPUT")    return false
					oSel = docSel.duplicate()
					oSel.text = ""
					var srcRange = obj.createTextRange()
					oSel.setEndPoint("StartToStart", srcRange)
					var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
					return reg.test(str)
				}
			function addChangeBuy(){	
				location.href='KS.Shop.asp?addtype=changebuy&ChannelID=<%=channelid%>&Action=Add';
                $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ChannelID=5&OpStr='+escape("添加商品")+'&ButtonSymbol=AddInfo';
			}
			  </script>
			  <%
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div class=""topdashed sort"">允许换购的商品管理</div>"
			  .Write  "<div class='pageCont2'><table width='100%' border='0' align='center' cellspacing='0' cellpadding='0'>"
              .Write "   <tr height='25' align='center' class='sort'>"
	          .Write "   <td width='5%' nowrap>序号</td>"
	          .Write "   <td>小图</td>"
	          .Write "   <td>商品名称</td>"
			  .Write "	 <td>订单满足金额</td>"
			  .Write "	 <td>换购价格</td>"
			  .Write "   <td>添加时间</td>"
			  .Write "   <td>已被兑换</td>"
			  .Write "   <td>操作</td>"
			  .Write " </tr>"
			  .Write "<form name='myform' action='?' method='get'>"
			  .Write "<input type='hidden' name='action' value='BatchSave'>"
			  
			  Dim XML,Node,Param,TotalPages,PhotoUrl
			  MaxPerPage=100
			  Param=" IsChangedBuy=1"
			  SQLStr=KS.GetPageSQL("KS_Product","id",MaxPerPage,Page,1,Param,"*")
		      Set RS = Server.CreateObject("AdoDb.RecordSet")
		      RS.Open SQLStr, conn, 1, 1
			  If RS.EOF Then
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' height='25' class='splittd' colspan=10>没有允许换购的商品!</td>"
					 .Write "</tr>"
			  Else
					totalPut = Conn.Execute("Select count(id) from [KS_Product] where " & Param)(0)
					Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
					For Each Node In XML.DocumentElement.SelectNodes("row")
					 PhotoUrl=Node.SelectSingleNode("@photourl").text
					 If KS.IsNul(PhotoUrl) Then PhotoUrl="../../images/nopic.gif"
					 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
					 .Write "<td align='center' class='splittd'>" & Node.SelectSingleNode("@id").text & "<input type='hidden' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
					 .Write "<td align='center' class='splittd'><Img src='" &PhotoUrl & "' width='40' height='40' /></td>"
					 .Write "<td class='splittd'>" & KS.Gottopic(Node.SelectSingleNode("@title").text,40) & "</td>"
					 .Write "<td align='center' class='splittd'><input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' name='changebuyneedprice" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@changebuyneedprice").text & "' size='5' style='text-align:center'> 元</td>"
					 .Write "<td align='center' class='splittd'><input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' name='changebuypresentprice" & Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@changebuypresentprice").text & "' size='5' style='text-align:center'> 元</td>"
					 .Write "<td align='center' class='splittd'>" & FormatDateTime(Node.SelectSingleNode("@adddate").text,2) & "</td>"
					 .Write "<td align='center' class='splittd'><font color=red>" &conn.execute("select count(1) from ks_orderitem where proid=" &Node.SelectSingleNode("@id").text &" and IsChangedBuy=1")(0) & "</font> 次</td>"
					 .Write "<td align='center' class='splittd'><a href='?Action=DelChangeBuy&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('确定移除吗?'));"">移除</a> | <a href='../System/KS.ItemInfo.asp?channelid=5&id="& Node.SelectSingleNode("@id").text&"&Action=Delete' onclick=""return(confirm('删除后不可恢复,确定删除吗?'))"">永久删除</a> | <a href='?Page=" & Page & "&Action=Edit&ID=" &Node.SelectSingleNode("@id").text & "' onclick='parent.frames[""BottomFrame""].location.href=""../Post.Asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("编辑" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & ID & """;'>修改</a></td>"
					 .Write "</tr>"
					Next
	         End If
			 .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			 .Write "<td colspan=10 class='pt10'>&nbsp;<input type='submit' value='批量设置换购价' class='button'>&nbsp;<input type='button' value='添加换购商品' onclick=""addChangeBuy()"" class='button'> </td>"
			 .Write "</tr>"
			 .Write "</form>"
			 .Write "</table>"
			 
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			 .Write "</div>"
			 RS.Close
			 Set RS=Nothing

		End With
	   End Sub
	   
	   '移出兑换品
	   Sub DelChangeBuy()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要移除的商品"
		ID=KS.FilterIDs(ID)
		Conn.Execute("Update KS_Product Set IsChangedBuy=0 Where ID in(" & ID & ")")
		KS.AlertHintScript "恭喜,已将选中的商品移出!"
	   End Sub
	   
	   '批量设置换购价格
	   Sub BatchSave()
	    Dim ID:ID=KS.G("ID")
		If ID="" Then KS.AlertHintScript "对不起,请先选择要移除的商品"
		Dim i,IDArr
		IDArr=Split(KS.FilterIDs(ID),",")
		For i=0 to ubound(idArr)
		  Conn.Execute("Update KS_Product Set changebuyneedprice=" & KS.G("changebuyneedprice"&idArr(i)) & ",changebuypresentprice=" & KS.G("changebuypresentprice"&idArr(i)) &" Where ID=" & Idarr(i))
		Next
		KS.AlertHintScript "恭喜,批量设置换购价格成功!"
	   End Sub
	   
	   
	   

       Sub ShopAdd() 
			With Response
			.Write"<!DOCTYPE html><html>"
			.Write "<head>" & vbCrlf
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrlf
			.Write "<title>添加商品</title>" & vbCrlf
			.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>" & vbCrlf
			.Write "<script src='../../KS_Inc/JQuery.js'></script>" & vbCrlf
			.Write "<script src='../../KS_Inc/common.js'></script>" & vbCrlf
			.Write "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>" & vbCrlf
			.Write "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>"
			.Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write EchoUeditorHead()
		
			CurrPath = KS.GetUpFilesDir()
			Dim WapTemplateID			
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Action = "Add" Then
			  FolderID = Trim(KS.G("FolderID"))
			  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '检查是否有添加商品的权限
			   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='../Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
			   Call KS.ReturnErr(2, "../System/KS.ItemInfo.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
			   Exit Sub
			  End If
			  ProductID=0: IsDiscount=1:Hits = 0:HitsByDay = 0: HitsByWeek = 0:HitsByMonth = 0:TotalNum = 1000: AlarmNum = 10:Comment = 1:BrandID=0:Strip=0:IsChangedBuy=0:Weight=0 : ChangeBuyPresentPrice=1 : ChangeBuyNeedPrice=100 : IsScore=0 :VIPPrice=0 : FreeShipping=0 : Price_Member=0 : Price=0 : WholesaleNum=0 : WholesalePrice=0
				IsLimitbuy     = KS.ChkClng(Request("TaskType")) : LimitBuyPrice  = 100 : LimitBuyAmount = 100 : membernum      = 0 : visitornum     = 0
				LimitBuyTaskID = KS.ChkClng(Request("LimitBuyTaskID"))
			    ProID          = KS.GetInfoID(ChannelID)
			    KeyWords       = Session("keywords")
			    ProducerName   = Session("ProducerName")
			    TrademarkName  = Session("TrademarkName")
				AddType        = KS.G("AddType")
			ElseIf Action = "Edit" Or Action="Verify" Then
			   Set RS = Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From KS_Product Where ID=" & KS.ChkClng(KS.G("ID")), conn, 1, 1
			   If RS.EOF And RS.BOF Then Call KS.Alert("参数传递出错!", ComeUrl):Exit Sub
				ID = Trim(RS("ID"))
				FolderID = Trim(RS("Tid"))
				If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '检查是否有编辑商品的权限
				RS.Close:Set RS = Nothing
				 If KeyWord = "" Then
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='../Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
				  Call KS.ReturnErr(1, "KS.Shop.asp?Page=" & Page & "&ID=" & FolderID)
				 Else
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='../Post.Asp?OpStr=" & server.URLEncode("商品管理 >> <font color=red>搜索商品结果</font>") & "&ButtonSymbol=ShopSearch';</script>")
				  Call KS.ReturnErr(1, "KS.Shop.asp?Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate)
				 End If
				 Exit Sub
			   End If
			    ProductID      = RS("ID")
				ProID          = RS("ProID")
				Title          = Trim(RS("title"))
				PhotoUrl       = Trim(RS("PhotoUrl"))
				BigPhoto       = Trim(RS("BigPhoto"))
				BigPhoto       = Trim(RS("BigPhoto"))
				ProIntro       = Trim(RS("ProIntro"))
				Rolls          = CInt(RS("Rolls"))
				Recommend      = CInt(RS("Recommend"))
				Popular        = CInt(RS("Popular"))
				Verific        = CInt(RS("Verific"))
				Comment        = CInt(RS("Comment"))
				Slide          = CInt(RS("Slide"))
				IsSpecial      = RS("IsSpecial")
				IsTop          = RS("IsTop")
				Strip          = RS("Strip")
				AddDate        = RS("AddDate")
				Rank           = Trim(RS("Rank"))
                FileName       = RS("Fname")
				TemplateID     = RS("TemplateID")
				WapTemplateID  = RS("WapTemplateID")
				Hits           = Trim(RS("Hits"))
				HitsByDay      = Trim(RS("HitsByDay"))
				HitsByWeek     = Trim(RS("HitsByWeek"))
				HitsByMonth    = Trim(RS("HitsByMonth"))
				IsChangedBuy   = RS("IsChangedBuy")
				ChangeBuyNeedPrice=RS("ChangeBuyNeedPrice")
				ChangeBuyPresentPrice=RS("ChangeBuyPresentPrice")
				Weight         = RS("Weight")
				IsLimitbuy     = KS.ChkClng(RS("IsLimitBuy"))
				LimitBuyPrice  = RS("LimitBuyPrice")
				LimitBuyTaskID = RS("LimitBuyTaskID")
				LimitBuyAmount = KS.ChkCLng(RS("LimitBuyAmount"))
				DownUrl        = RS("DownUrl")
				membernum      = RS("membernum")
				visitornum     = RS("visitornum")
				arrGroupID     = RS("arrGroupID")
				Unit           = Trim(RS("Unit"))
				TotalNum       = Trim(RS("TotalNum"))
				AlarmNum       = Trim(RS("AlarmNum"))
				Price          = RS("Price")
				IsDiscount     = RS("IsDiscount")
				Price_Member   = RS("Price_Member")
				VIPPrice       = RS("VIPPrice")
				KeyWords       = Trim(RS("KeyWords"))
				ProducerName   = Trim(RS("ProducerName"))
				TrademarkName  = Trim(RS("TrademarkName"))
				ProModel       = RS("ProModel")
				ProSpecificat  = RS("ProSpecificat")
				ServiceTerm    = RS("ServiceTerm")
				FolderID       = RS("Tid")
				oTid           = RS("oTid")
				oID            = RS("Oid")
				BrandID        = RS("BrandID")
				AttributeCart  = RS("AttributeCart")
				SEOTitle       = RS("SEOTitle")
				SEOKeyWord     = RS("SEOKeyWord")
				SEODescript    = RS("SEODescript")
				RelatedID      = RS("RelatedID")
				FreeShipping   = RS("FreeShipping")
				WholesalePrice = RS("WholesalePrice")
				WholesaleNum   = RS("WholesaleNum")
				Changes        = RS("Changes")
				ChangesUrl     = RS("ChangesUrl")
				IsScore        = RS("Istype")
				Score          = RS("Score")
				
				'自定义字段
				if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
					Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
					If diynode.length>0 Then
						Set FieldDictionary=KS.InitialObject("Scripting.Dictionary")
						For Each FNode In DiyNode
						   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),RS(FNode.SelectSingleNode("@fieldname").text)
						   If FNode.SelectSingleNode("showunit").text="1" Then
						   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text) &"_unit",RS(FNode.SelectSingleNode("@fieldname").text&"_Unit")
						   End If
						Next
					End If
				End If
			End If
			If KS.IsNul(AttributeCart) Then AttributeCart=""
			'取得上传权限
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
			.Write "<script>var BigPhoto='" & BigPhoto & "';</script>" & vbCrlf
			%>
			<script type="text/javascript">
		    $(document).ready(function(){
				$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
				$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
				$('#KeyLinkByTitle').click(function(){
			     GetKeyTags();
			    });
				$('#checkProIdButton').click(function(){
				  CheckProIdRepeat();
				});
				if ($("#Changes").attr('checked')){ChangesNews();}
				//只有一个栏目时，为选其选中
			  if ($("#tid option").length<=2){
			     $("#tid option").each(function(i){
				    if (i==$("#tid option").length-1) $(this).attr("selected",true);
				});
			  }
				
		    })
            function GetKeyTags()
			{
			  var text=escape($('input[name=title]').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{
			   top.$.dialog.alert('对不起,请先输入商品名称!');
			  }
			}
			function CheckProIdRepeat(){
			 var text=$('input[name=proid]').val();
			 if(text!=''){
			    if (text.length<8){
				 top.$.dialog.alert('编号长度不能小于8位!');
				}else{
				 $(parent.document).find("#ajaxmsg").toggle(true);
				 $.get("../../shop/ajax.getdate.asp",{action:"Shop_CheckProID",proid:escape(text),id:<%=ProductID%>},function(d){
				 $(parent.document).find("#ajaxmsg").toggle(false);
				 top.$.dialog.alert(unescape(d));
				 });
				}
			 }else{
			  top.$.dialog.alert('请输入商品编号!');
			  $("input[name=proid]").focus();
			 }
			}
			function UnSelectAll(){
			  $("#SpecialID>option").each(function(){
				$(this).attr("selected",false);
				});
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
			function GetFileNameArea(f){$('#filearea').toggle(f);}
			function GetTemplateArea(f){$('#templatearea').toggle(f);}
			function insertHTMLToEditor(codeStr){<%=InsertEditor("Content","codeStr")%>} 
            
			function SubmitFun()
			{ 
			  var picSrcs='';
			  var src='';
			  var b=false;
			  $("#thumbnails").find(".pics").each(function(){
			     if ($(this).next().val()==''){alert('第 '+($("#thumbnails").find(".pics").index($(this))+1)+' 张图片没有输入组图名称，请输入!');$(this).next().focus();b=true;return false;}
			     src=$(this).next().val().replace('$$$','').replace('|','')+"|"+$(this).val();
			     if(picSrcs==''){picSrcs=src;}else{picSrcs+='$$$'+src;}
			  });
			  
			  if (b) return false;
			    $("#PicUrls").val(picSrcs);
			  
			    if ($('input[name=title]').val()==""){
					top.$.dialog.alert("请输入商品名称！",function(){
					$('input[name=title]').focus();
					});
					return;
				}
				if ($('input[name=proid]').val()==""){
				    top.$.dialog.alert("请输入商品编号!",function(){
					$('input[name=proid]').focus();
					});
					return;
				}
				 if ($('#tid').val()=='0' || $('#tid').val()==''){
				    top.$.dialog.alert('请选择商品<%=KS.GetClassName(ChannelID)%>!');
					return ;
				 }
				  if ($('#Unit').val()==""){
				    top.$.dialog.alert("商品单位不能为空!",function(){
					$('#Unit').focus();
					});
					return;
				  }
				  
             if (parseInt($("#totalrow").val())>0){
			    var sstr='';
			    for (var i=1;i<=parseInt($("#totalrow").val());i++){
				  if ($("#aitemno"+i)[0]!=undefined){
				      if ($("#aitemno"+i).val()==''){
					   top.$.dialog.alert('第' +i+'个货号必须输入!',function(){
					   $("#aitemno"+i).focus();
					   });
					   return false;
					  }
				      if ($('input[name="attr'+i+'0"]').val()==''){
					   top.$.dialog.alert('第' +i+'个['+$('#attrtitle0').val()+']规格值必须输入!',function(){
					   $('input[name="attr'+i+'0"]').focus();
					   });
					   return false;
					  }
					  if ($('input[name="attr'+i+'1"]')[0]!=undefined){
						  if ($('input[name="attr'+i+'1"]').val()==''){
						   top.$.dialog.alert('第' +i+'个['+$('#attrtitle1').val()+']规格值必须输入!',function(){
						   $('input[name="attr'+i+'1"]').focus();
						   });
						   return false;
						  }
					  }
					  if ($('input[name="attr'+i+'2"]')[0]!=undefined){
						  if ($('input[name="attr'+i+'2"]').val()==''){
						   top.$.dialog.alert('第' +i+'个['+$('#attrtitle2').val()+']规格值必须输入!',function(){
						   $('input[name="attr'+i+'2"]').focus();
						   });
						   return false;
						  }
					  }
					  
					  if (sstr.indexOf($("#aitemno"+i).val().toLowerCase()+",")!=-1){
					   top.$.dialog.alert('货号不能相同!',function(){
					   $("#aitemno"+i).focus();
					   });
					   return false;
					  }
					  sstr=sstr+$("#aitemno"+i).val().toLowerCase()+',';
				 }
				}
			  }
				if ($("#BeyondSavePic").prop("checked")==true) {
				   $("#LayerPrompt").show();
				  window.setInterval('ShowPromptMessage()',150)
				 }
				  $('#myform').submit();
				  $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
			}

			function getBrandList(v)
			{ if (hasSpecification){getSpecification();}
			  if (v==0)
			  $("#brandarea").html("");
			  else
			  {$(parent.document).find("#ajaxmsg").toggle(true);
			   var url='../../shop/ajax.getdate.asp';
			   $.get(url,{action:"Shop_BrandOption",classid:$("#tid").val()},function(d){
			   $(parent.document).find("#ajaxmsg").toggle(false);
			   $("#brandarea").html(unescape(d));
			   });
 
			  }
			}
			var ForwardShow=true;
			function ShowPromptMessage()
			{
				var TempStr=ShowArticleArea.innerText;
				if (ForwardShow==true)
				{
					if (TempStr.length>4) ForwardShow=false;
					ShowArticleArea.innerText=TempStr+'.';
				}
				else
				{
					if (TempStr.length==1) ForwardShow=true;
					ShowArticleArea.innerText=TempStr.substr(0,TempStr.length-1);
				}
			}
			function ChangesNews()
			{ 
			 if ($("#Changes").prop('checked'))
			  {
			  $("#ChangesUrl").attr("disabled",false);
			  }
			  else
			   {
			  $("#ChangesUrl").attr("disabled",true);
			   }
			}
			
			
			var SaveBeyondInfo=''
					   +'<div id="LayerPrompt" style="position:absolute; z-index:1; left: 200px; top: 150px; background-color: #f1efd9; layer-background-color: #f1efd9; border: 1px none #000000; width: 360px; height: 63px; display: none;"> '
					   +'<table width="100%" height="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FF0000">'
					   +'<tr> '
					   +'<td align="center">'
					   +'<table width="80%" border="0" cellspacing="0" cellpadding="0">'
					   +'<tr>'
					   +' <td width="75%" nowrap>'
					   +'<div align="right">请稍候，系统正在保存远程图片到本地</div></td>'
					   +'   <td width="25%"><font id="ShowArticleArea">&nbsp;</font></td>'
					   +' </tr>'
					   +'</table>'
					   +'</td>'
					   +'</tr>'
					   +'</table>'
					   +'</div>'
			document.write (SaveBeyondInfo)
			</script>
			<%
			
			Call KSCls.EchoFormStyle(ChannelId)   '控制添加文档布局
			
			.Write "</head>"
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>"
			.Write "<div>"
			.Write "<ul id='menu_top' class='menu_top_fixed'>"
			.Write "<li onclick=""return(SubmitFun())"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon save'></i>确定保存</span></li>"
			.Write "<li onclick=""history.back();"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消返回</span></li>"
		    .Write "</ul></div>"		
			.Write "<div class=""menu_top_fixed_height""></div>"	
			
			.Write "<form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' id='myform' name='myform' >"
           .Write "<div class=tab-page id=ShopPane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""ShopPane"" ), 1 )"
			.Write " </SCRIPT>"
					 

			.Write "      <input type='hidden' value='" & ProductID & "' name='ProductID'>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
			
			Dim AckPlusTF:AckPlusTF=KS.GetAppStatus("tags")
			Call KS.LoadFieldGroupXML()
			
Dim TypeNode,TTN
TTN=0
IF IsObject(Application(KS.SiteSN & "_FieldGroupXml")) Then
  For Each TypeNode In Application(KS.SiteSN & "_FieldGroupXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &"]")
     .Write " <div class=tab-page id=""p" &TypeNode.SelectSingleNode("@id").text & """>"
	 .Write "  <H2 class=tab>" & TypeNode.SelectSingleNode("@groupname").text & "</H2>"
	 .Write "	<SCRIPT type=text/javascript>"
	 .Write "				 tabPane1.addTabPage( document.getElementById( ""p" &TypeNode.SelectSingleNode("@id").text & """ ) );"
	 .Write "	</SCRIPT>"
	 TTN=TTN+1
			
	
	  .Write " <dl class='dtable'>"
			
			
		IF TTN=1 THEN
		 if addtype="changebuy" or IsChangedBuy="1" then
			.Write " <dd>"
			.Write "  <div>换购条件:</div>"
			.Write "   <input type='hidden' name='IsChangedBuy' value='1'>"
			.Write "<font id='ChangedBuy' style='font-weight:normal;font-size:12px;border:1px solid #f9c943;background:#FFFFF6;margin-top:5px;padding:10px;'>"
			.Write "订单满足金额：<input type='text' style='text-align:center' name='ChangeBuyNeedPrice' value='" & ChangeBuyNeedPrice &"' size='6'size='4' maxlength='4' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox'>元 换购价格：<input type='text' style='text-align:center' name='ChangeBuyPresentPrice' value='" & ChangeBuyPresentPrice &"'  size='4' maxlength='4' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" size='6' class='textbox'>元 "
			.Write "<font class='tips'>当购买商品总金额>=订单满足金额时,可以根据换购价格得到此商品。</font>"
			.Write "</font>"
			.Write " </dd>"
		end if	
		
		'if addtype="limitbuy" or KS.ChkClng(LimitBuyTaskID)<>0 then	
			.Write " <dd>"
			.Write "  <div>限时限量:</div>"
			%>
			<script>
			 function loadTask(){
			  var v=$("input[name=IsLimitbuy]:checked").val();
			  if (v==0){
			   $("#showlimitbuy").hide();
			  }else{
			       $("#showlimitbuy").show();
				   $(parent.document).find("#ajaxmsg").toggle(true);
				   var url='../../shop/ajax.getdate.asp';
				   $.get(url,{action:"Shop_LimitBuyTask",TaskType:v,LimitBuyTaskID:'<%=LimitBuyTaskID%>'},function(d){
				    $(parent.document).find("#ajaxmsg").toggle(false);
				    $("#LimitBuyTaskID").empty().append(d);
				   });
				}
			 }
			 $(document).ready(function(){
			   loadTask();
			 });
			</script>
			<label><input name='IsLimitbuy' onClick="loadTask()" type='radio'  value='0' <%if KS.ChkClng(IsLimitbuy)="0" then response.write " checked"%>>不启用促销</label>
			<label><input name='IsLimitbuy' onClick="loadTask()" type='radio'  value='1' <%if IsLimitbuy="1" or addType="limitbuy" then response.write " checked"%>>限时抢购</label>
			<label><input name='IsLimitbuy' onClick="loadTask()" type='radio'  value='2' <%if IsLimitbuy="2" then response.write " checked"%>>限量抢购</label>
		
			<table id="showlimitbuy" style='border:1px solid #f9c943;background:#FFFFF6;display:none;margin-top:5px;padding:10px;'>
			<tr><td>
			抢购任务:<select name='LimitBuyTaskID' id='LimitBuyTaskID'>
			</select>
			<%
			.Write " <a href='KS.Shop.asp?action=LimitBuy&channelid=5'>抢购任务管理</a><br/>"
			.Write "抢 购 价:<input type='text' style='text-align:center' name='LimitBuyPrice' id='LimitBuyPrice' size='6'  value='" & LimitBuyPrice &"' size='4' maxlength='8' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox'>元<br/>"
			.Write "抢购数量:<input type='text' class='textbox' name='LimitBuyAmount' id='LimitBuyAmount' value='" & LimitBuyAmount & "' size='10'/>件   设置允许让抢购的商品数<br/>"
			.Write "</td></tr></table>"
			.Write "</dd>"	
	End If	
			
	For Each FNode In FieldXML.DocumentElement.SelectNodes("fielditem[showonform=1&&fieldtype!=13&&@groupid=" & TypeNode.SelectSingleNode("@id").text & "]")

	    If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
			.Write   KSCls.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary,0) '自定义字段
		Else
		 Dim XTitle:XTitle=FNode.SelectSingleNode("title").text
	     Select Case lcase(FNode.SelectSingleNode("@fieldname").text)
	       case "title"
				.Write " <dd><div>" & XTitle &":</div><input name='title' id='title' type='text'  class='rule' value='" & Title & "' size=50><font color='#FF0000'>*</font> "
				.Write "<input class='button' type='button' value='重名检测' onclick=""if($('#title').val()==''){ top.$.dialog.alert('请输入" & KS.C_S(ChannelID,3) & "标题!');}else top.openWin('" & KS.C_S(ChannelID,3) & "重名检测','Shop/KS.Shop.asp?ChannelID=" & ChannelID & "&Action=CheckTitle&title='+escape($('#title').val()),false,360,370);"">"
				.Write "<strong>商品编号:</strong><input name='proid' type='text' id='proid' class='textbox' value='" & proid & "' size=18> <input type='button' value='检测重复' id='checkProIdButton' class='button'><font color=green>商品号必须唯一</font><span id='cmsg'></span>"
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/showonform").text="1" Then
					.Write "<label><input type='checkbox' name='MakeHtml' value='1' checked>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/title").text & "</label>"
				End IF
				if RelatedID=-11 or KS.ChkClng(RelatedID)<>0 then
							.Write "<span style=""padding:5px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6""><label><input type='checkbox' name='EditNewtb' value='1' checked/> 此"  & KS.C_S(ChannelID,3) & "发布到多个栏目，选中将同步更新 <input type='hidden' name='RelatedID' value='"& RelatedID &"'/></label></span>"
				end if
				.Write "</dd>" &vbcrlf
		  case "tid"
					.Write " <dd>"
					.Write "   <div>" & Replace(XTitle,"栏目",KS.GetClassName(ChannelID)) & ":</div>"
					.Write "   <input type='hidden' name='OldClassID' value='" & FolderID & "'>"
					If Action<>"Edit" Then
						.Write "&nbsp;<input name='Istidtb' type='button' class='button' id='istidtb' value='发布多" & Replace(XTitle,"栏目",KS.GetClassName(ChannelID)) & "'  onclick=""sel();"" >"
					end if	
					.Write "<select size='1' name='tid' id='tid' style='width:335px'>"
					.Write " <option value='0'>--请选择" & KS.GetClassName(ChannelID) &"--</option>"
					.Write Replace(KS.LoadClassOption(ChannelID,true),"value='" & FolderID & "'","value='" & FolderID &"' selected") & " </select>"
					' Call KSCls.EchoSelectTid(FolderID,ChannelID)
					%>
					<input type="hidden" id="tidtb" name="tidtb" value=""/>
					<script>
					var box=''
					function sel(){
					top.openWin(false,'shop/KS.Shop.asp?channelID=<%=ChannelID%>&FolderID='+$("#tidtb").val()+'&action=SelectClass',false,400,420);
					}
					</script>
					<%
					
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/showonform").text="1" Then
					.Write FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/title").text &":"
					.Write "<label><input name='Recommend' type='checkbox' id='Recommend' value='1'"
					If Recommend = "1" Then .Write (" Checked")
					.Write ">推荐</label><label><input name='IsSpecial' type='checkbox' id='IsSpecial' value='1'"
					If IsSpecial = "1" Then .Write (" Checked")
					.Write ">特价</label><label><input name='Popular' type='checkbox' id='Popular' value='1'"
					If Popular = "1" Then .Write (" Checked")
					.Write ">热卖</label><label><input name='Comment' type='checkbox' id='Comment' value='1'"
					If Comment = "1" Then .Write (" Checked")
					.Write ">允许评论</label><label><input name='IsTop' type='checkbox' id='IsTop' value='1'"
					If IsTop = "1" Then .Write (" Checked")
					.Write ">置顶</label><label><input name='Rolls' type='checkbox' id='Rolls' value='1'"
					If Rolls = "1" Then .Write (" Checked")
					.Write ">滚动</label><label><input name='Slide' type='checkbox' id='Slide' value='1'"
					If Slide = "1" Then .Write (" Checked")
					.Write ">幻灯</label><label><input name='Strip' type='checkbox' id='Strip' value='1'"
					If Strip = "1" Then .Write (" Checked")
					.Write ">头条</label>"
					Call KSCls.GetDiyAttribute(FieldXML,FieldDictionary)
					.Write " </dd>" & vbcrlf	
				  End If
				  
				  
				.Write "<dd id='ContentLink'>"
				.Write "   <div>外部链接:</strong></div>" &vbcrlf 
				If ChangesUrl = "" Then
				 .Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='50' class='textbox'>")
				Else
				 .Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl & "' size='50' class='textbox'>")
				End If
				If Changes = 1 Then
				 .Write (" <input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
				Else
				 .Write (" <input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
				End If
				.Write " </dd>" & vbcrlf
			case "otid"
		        Call KSCls.EchoOTidInfo(FNode,OTid,Oid)		
			case "brandid"	
					.Write "<dd>"
					.Write "  <div>" & XTitle &":</div>"
					.Write "  <font id='brandarea'>"
					If Action="Edit" or (FolderID<>"" and FolderID<>"0") Then
					.Write GetBrandByClassID(FolderID,BrandID)
					Else
					 .Write "<select name='BrandID'><option value='0'>--选择品牌--</option></select>"
					End If
					.Write "</font><span>说明：品牌由分类决定自动关联</span>"
					.Write "   </dd>" &vbcrlf
		    case "photourl"
					.Write " <dd>"
					.Write "   <div>列表图片:</div>"
					.Write "   <table border='0' cellpadding='0' cellspacing='0'><tr><td nowrap>小图：<input name='PhotoUrl' type='text' id='PhotoUrl' size='50' value='" & PhotoUrl & "' class='textbox'>"
					.Write " <input class=""button"" type='button' name='Submit' value='选择小图...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=" & CurrPath & "',550,290,window,document.myform.PhotoUrl,'pic');""> <input class=""button"" type='button' name='Submit' value='远程抓图...' onClick=""top.openWin('抓取远程图片','include/SaveBeyondfile.asp?pic=pic&fieldid=PhotoUrl&CurrPath=" & CurrPath & "',false,500,100);"">"
				    .Write "        <input class=""button""  type='button' name='Submit' value='裁剪...' onClick=""if($('#PhotoUrl').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindow(1,'" & KS.Setting(3) & "',$('#PhotoUrl').val())}"">  "
					.Write "   <br/> 大图：<input name='BigPhoto' type='text' id='BigPhoto' size='50' value='" & BigPhoto & "' class='textbox'>"
					.Write " <input class=""button"" type='button' name='Submit' value='选择大图...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=" & CurrPath & "',550,290,window,document.myform.BigPhoto,'pic');""> <input class=""button"" type='button' name='Submit' value='远程抓图...' onClick=""top.openWin('抓取远程图片','include/SaveBeyondfile.asp?pic=pic&fieldid=BigPhoto&CurrPath=" & CurrPath & "',false,500,100);"">"
				   .Write " <input class=""button""  type='button' name='Submit' value='裁剪...' onClick=""if($('#BigPhoto').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindows(1,'" & KS.Setting(3) & "',$('#BigPhoto').val(),$('#BigPhoto')[0])}""></td><td width='200'>" &vbcrlf
				   .Write "<div  style=""margin:0 auto;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:50px;width:50px;border:1px solid #777777""><img src=""" & PhotoUrl & """ onerror=""this.src='../../images/logo.png';"" id=""pic"" style=""height:50px;width:50px;"">"
				    .write "</td></tr></table>"
				   
					.Write "  </dd>"
		    case "uploadphoto"
					If CBool(UpPowerFlag) = True Then
					.Write "<dd><div></div>"
				   .Write "<table width=""90%""><tr><td> <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?showpic=pic&ChannelID=" & ChannelID &"&UpType=Pic' frameborder=0 scrolling=no width='100%' height='30'></iframe></td></tr></table>"
					.Write " </dd>" & vbcrlf
					End If
			case "unit"
					.Write " <dd>"
					.Write "  <div>" & XTitle &":</div>"
					.Write "<input name='Unit' type='text' id='Unit' style='text-align:center' value='" & Unit & "' size='10' class='textbox'> "
					.Write " << <select name='sUnit' onchange=""$('#Unit').val(this.value);"">"
					.Write " <option value=''>请选择</option>"
					.Write  KSCls.Get_O_F_D(KS.C_S(ChannelID,2),"Distinct Top 10 Unit","1=1 Group by unit")
					.Write "</select>"
					.Write " 单件重量：<input name='Weight' type='text' style='text-align:center' id='Weight' value='" & Weight & "' size='6' class='textbox'>KG <span style='color:green'>用于计算运费，如果不设置请输入0，表示按首重计算运费。</span><br/>"
					.Write "<strong>免邮设置：</strong>购买<input name='FreeShipping' type='text' style='text-align:center' id='FreeShipping' value='" & FreeShipping & "' size='6' class='textbox'>件免邮,输入“0”不免邮。"
					
					.Write " </dd>" &vbcrlf
		   case "price"
					.Write " <dd><div>" & XTitle &":</div>"
					.Write "  <input id=""IsScore"" name=""IsScore"" value=""0"" onclick=""$('#jf_box').hide();"" "
					If KS.ChkClng(IsScore)=0 Then .Write "checked=""checked"" "
					.Write " type=""radio"">正常销售 "
					.Write " <input id=""IsScore"" name=""IsScore"" value=""1"" onclick=""$('#jf_box').show();"" "
					If KS.ChkClng(IsScore)=1 Then .Write "checked=""checked"" "
					.Write "type=""radio"">积分兑换 <br/>"
					.Write "<font id=""jf_box"" "
					 If KS.ChkClng(IsScore)=0 Then .Write " style='display:none'"
					.Write ">兑换积分<input type='text' size='6' style='text-align:center' class='textbox' name='Score' value='"& KS.ChkClng(Score) &"' >积分 + "
					.Write "</font>"
					.Write "<font "
					.Write ">商城价<input name='Price_Member' type='text' style='text-align:center' id='Price_Member' value='" &KS.GetPrice(Price_Member) & "' size='6' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"">元 (<font class='tips'>所有会员组都以商城价为准计算折扣</font>)"
					.Write "</font>"
					.Write "<div style=""clear:both""></div>"
					.Write "<table  id=""Price_box"" style=""font-weight:normal;font-size:12px""><tr><td>"
					.Write "参考价<input name='Price' type='text' style='text-align:center' id='Price' value='" & KS.GetPrice(Price) & "' size='6' class='textbox' onKeyPress=""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"">元&nbsp;&nbsp;"
					.Write "<br/>VIP 价<input name='VIPPrice' type='text' style='text-align:center' id='VIPPrice' value='" & KS.GetPrice(VIPPrice) & "' size='6' class='textbox' onKeyPress=""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"">元 (<font class='tips'>VIP用户以VIP价格为准。</font>)"
					.write "<br/>批发价：购买大于等于<input name='WholesaleNum' style='text-align:center' type='text' id='WholesaleNum' value='" & WholesaleNum & "' size='6' class='textbox'>件，每件按批发价<input name='WholesalePrice' type='text' style='text-align:center' id='WholesalePrice' value='" & KS.GetPrice(WholesalePrice) & "' size='6' class='textbox'>元计算。"
					.Write "</td></tr></table>"
					.Write " </dd>" & vbcrlf
		   case "isdiscount"
					.Write " <dd><div>" & XTitle &":</div>"
					.Write "<input type='radio' name='IsDiscount' value='1'"
					if isdiscount=1 then .write " checked"
					.Write ">允许"
					.Write "<input type='radio' name='IsDiscount' value='0'"
					if isdiscount=0 then .write " checked"
					.Write ">不允许"
					.Write "&nbsp;<span>(如果这里设置为不允许，那么所有VIP会员组将不再享受优惠)</span></dd>"
		   case "totalnum"
					.Write "<dd><div>" & XTitle &":</div>库存数量<input name='TotalNum' style='text-align:center' type='text' id='TotalNum' value='" & TotalNum & "' size='10' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"">&nbsp;&nbsp;&nbsp;库存报警下限数&nbsp;<input name='AlarmNum' type='text' id='AlarmNum' value='" & AlarmNum & "' style='text-align:center' size='10' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"">"
				   .Write "</dd>" &vbcrlf
		   case "cartpropert"
            %>
					 <style type="text/css">
					  .atcs li{padding:2px;float:left;border:1px #efefef solid;background:#FFFFF6;margin-right:6px}
					  .ttt{background:#FFFFF6;border:1px #F9C943 solid;width:50px;height:20px;line-height:20px;text-align:center}
					  .dtable dd input.ttt[type="text"]{ margin-bottom:0 !important;}
					 </style>
					 <script type="text/javascript">
					  var hasSpecification=false;
					  function getSpecification(){
					   var tid=$("#tid option:selected").val();
					   if (tid=='0'){
					    alert('请先选择商品分类!');
						return false;
					   }
					   $(parent.document).find("#ajaxmsg").toggle();
                       $.get("../../shop/ajax.getdate.asp",{action:"getSpecification",id:<%=KS.ChkClng(Request("ID"))%>,classid:tid},function(d){ 
					        
					        var r=unescape(d).split('|');
							if (r[0]=='error'){alert(r[1]);return;}else{hasSpecification=true;$("#cartattr").html(d);}
							$(parent.document).find("#ajaxmsg").toggle();
						});
					  }
					  function delrow(r){$("#row"+r).remove();}
					  function delrowajax(r,id){
					   if (confirm('此操作不可逆，确定删除该货号吗？')){
					   $.get("../../shop/ajax.getdate.asp",{action:"deleteproitem",id:id},function(d){ 
					        var r=unescape(d).split('|');
							if (r[0]=='error'){lert(r[1]);return;}
						});
						delrow(r);
						}
					  }
					  function additemno(c){
					    var row=parseInt($("#totalrow").val())+1;
						var imgstr1='';
						if ($("#attrimg"+(row-1)+"0")[0]!=undefined){imgstr1="<img id='i"+row+"0' src='../../images/nopic.gif' width='25' title='请上传图片' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"0' id='attrimg"+row+"0' value=''/><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?ChannelID=5&get=min&UpType=Pic&imgname=i"+row+"0&FieldName=attrimg"+row+"0' frameborder=0 scrolling=no width='50' height='26'></iframe>";}else{imgstr1='';}
						var str='<table width="98%" class="attbox" cellspacing="0" cellpadding="0"><tr id="row'+row+'"><td class="splittd"><input value="NO'+row+'" type="text" name="aitemno'+row+'" id="aitemno'+row+'" size="10" class="textbox"/></td><td width="150"><input type="text" name="attr'+row+'0" class="ttt" size="8" value=""/> '+imgstr1+'</td>';
						if (c>=1){
						  if ($("#attrimg"+(row-1)+"1")[0]!=undefined){imgstr1="<img id='i"+row+"1' src='../../images/nopic.gif' width='25' title='请上传图片' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"1' id='attrimg"+row+"1' value=''/><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?ChannelID=5&get=min&UpType=Pic&imgname=i"+row+"1&FieldName=attrimg"+row+"1' frameborder=0 scrolling=no width='50' height='26'></iframe>";}else{imgstr1='';}
						 str+='<td width="150" class="splittd"><input type="text" class="ttt" size="8" name="attr'+row+'1" value=""/> '+imgstr1+'</td>';
						 if (c>=2){
						  if ($("#attrimg"+(row-1)+"2")[0]!=undefined){imgstr1="<img id='i"+row+"2' src='../../images/nopic.gif' width='25' title='请上传图片' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"2' id='attrimg"+row+"2' value=''/><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?ChannelID=5&get=min&UpType=Pic&imgname=i"+row+"2&FieldName=attrimg"+row+"2' frameborder=0 scrolling=no width='50' height='26'></iframe>";}else{imgstr1='';}
						 str+='<td width="150" class="splittd"><input type="text" class="ttt" size="8" name="attr'+row+'2" value=""/> '+imgstr1+'</td>';
						 }
						}
						str+=rowstr(row)+"</table>";
						$("#alist").append(str);
						$("#totalrow").val(row);
					  }
					  function rowstr(row){
					   var str='<td width="100" class="splittd"><input type="text" name="aprice'+row+'" value="'+$("#Price_Member").val()+'" size="4" class="textbox"/>元</td><td width="100" class="splittd"><input type="text" name="aamount'+row+'" value="'+$("#TotalNum").val()+'" size="4" class="textbox"/>件</td><td width="100" class="splittd"><input type="text" name="aweight'+row+'" value="'+$("#Weight").val()+'" size="4" class="textbox"/>KG</td><td width="100" class="splittd"><a href="javascript:delrow('+row+');">删除</a></td></tr>';
					   return str;
					  }

					   function getlist(t){
					    var titlestr='';
						var firstct=false;
						
					    for (var i=0;i<3;i++){
							eval("str"+i+"=new Array();");
							eval("strimg"+i+"=new Array();");
							var n=0;
							var ct=false;
							$("input[name='cc"+i+"']:checked").each(function(){
							  eval("str"+i+"[n++]=$(this).val();");
							  if (parseInt($("#ashowtype"+i).val())==2){
							   eval("strimg"+i+"[n]=$(this).next().val();");
							  }
							  ct=true;
							  if (i==0 && ct){firstct=true;}
							});
							if (ct&&firstct){
							  if (titlestr==''){titlestr=$("#attrtitle"+i).val();}else{titlestr=titlestr+','+$("#attrtitle"+i).val();}
							}
						}
						
						var str=imgstr1=imgstr2=imgstr3='';
						var row=0;
						for (var i=0;i<str0.length;i++){
						  if (str1==''){
						    $("#tt0").show();
							$("#tt1").hide();
							row++;
							if (parseInt($("#ashowtype0").val())==2){imgstr1="<img src='"+strimg0[i+1]+"' width='25' title='"+str0[i]+"' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"0' value='"+strimg0[i+1]+"'/>";}
							str=str+'<tr id="row'+row+'" style="text-align:center;"><td class="splittd"><input type="text" name="aitemno'+row+'" id="aitemno'+row+'" value="NO'+row+'" size="10" class="textbox"/></td><td class="splittd" width="150">'+imgstr1+' <input type="text" class="ttt"" name="attr'+row+'0" value="'+str0[i]+'"/></td>'+rowstr(row);
							
						  }
						   else{
						     $("#tt1").show();
							 for (var j=0;j<str1.length;j++){
								  if (str2==''){
								   $("#tt0").show();
								   $("#tt2").hide();
								   row++;
								   if (parseInt($("#ashowtype0").val())==2){imgstr1="<img src='"+strimg0[i+1]+"' width='25' title='"+str0[i]+"' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"0' value='"+strimg0[i+1]+"'/>";}
								   if (parseInt($("#ashowtype1").val())==2){imgstr2="<img src='"+strimg1[j+1]+"' width='25' title='"+str1[j]+"' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"1' value='"+strimg1[j+1]+"'/>";}
							       str=str+'<tr id="row'+row+'"><td class="splittd"><input type="text" value="NO'+row+'" name="aitemno'+row+'" id="aitemno'+row+'" size="10" class="textbox"/></td><td width="150">'+imgstr1+' <input type="text" class="ttt"" name="attr'+row+'0" value="'+str0[i]+'"/></td><td width="150">'+imgstr2+' <input type="text" class="ttt"" name="attr'+row+'1" value="'+str1[j]+'"/></td>'+rowstr(row);
								  }else{
								   $("#tt0").show();
								   $("#tt2").show();
									  for (var k=0;k<str2.length;k++){
										row++;
								        if (parseInt($("#ashowtype0").val())==2){imgstr1="<img src='"+strimg0[i+1]+"' width='25' title='"+str0[i]+"' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"0' value='"+strimg0[i+1]+"'/>";}
								        if (parseInt($("#ashowtype1").val())==2){imgstr2="<img src='"+strimg1[j+1]+"' width='25' title='"+str1[j]+"' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"1' value='"+strimg1[j+1]+"'/>";}
								        if (parseInt($("#ashowtype2").val())==2){imgstr3="<img src='"+strimg2[k+1]+"' width='25' title='"+str2[k]+"' height='25' align='absmiddle'/><input type='hidden' name='attrimg"+row+"2' value='"+strimg2[k+1]+"'/>";}
							            str=str+'<tr id="row'+row+'"><td class="splittd"><input value="NO'+row+'" type="text" name="aitemno'+row+'" id="aitemno'+row+'" size="10" class="textbox"/></td><td width="150">'+imgstr1+' <input class="ttt"" type="text" name="attr'+row+'0" value="'+str0[i]+'"/></td><td width="150">'+imgstr2+' <input type="text" class="ttt"" name="attr'+row+'1" value="'+str1[j]+'"/></td><td width="150">'+imgstr3+' <input type="text" class="ttt"" name="attr'+row+'2" value="'+str2[k]+'"/></td>'+rowstr(row);
									  }
								  }
							 }
						  }
						}
						$("#totalrow").val(row);
						$("#AttributeCart").val(titlestr);
						$("#alist").empty().append(str);
					  }
					 </script>
					 <%					  
				   	 tempstr="<dd><div>规格属性: </div>"
					Dim ACartArr,tname,hasattr
					If Not KS.IsNul(AttributeCart) Then
					    if not conn.execute("select top 1 * from KS_ShopSpecificationPrice where proid=" & KS.CHkClng(KS.G("ID"))).eof  then
						 ACartArr=split(AttributeCart,",")
						else
						 AttributeCart=""
						end if
							
					End If
					
					If (Action = "Edit" Or Action="Verify") and isarray(ACartArr) Then
					  for i=0 to ubound(ACartArr)
						 tempstr=tempstr &"<input type='hidden' name='attrtitle" &i & "' id='attrtitle" & i &"' value='" & ACartArr(i) &"'/>"
					  next
					tempstr=tempstr & "<td><table class='ctable' border='0' cellspacing='0' cellpadding='0'>"
					tempstr=tempstr & "<tr><td colspan='10'><div style='margin:4px'><input onclick=""additemno(" & Ubound(ACartArr) & ")"" type='button' value='添加一个货号' class='button'/></div></td></tr>"

				     tempstr=tempstr & "       <tr class='sort' style='text-align:center;'>"
				      tempstr=tempstr & "       <td  width='100'>货号</td>"
					   If IsArray(ACartArr) Then
						 for i=0 to ubound(ACartArr)
						  tempstr=tempstr & "       <td width='150' id='tt" & i &"'>" & ACartArr(i) &"</td>"
						 next
					   End If
				      tempstr=tempstr & "       <td  width='100'>销售价</td>"
				      tempstr=tempstr & "       <td  width='100'>库存</td>"
				      tempstr=tempstr & "       <td  width='100'>重量</td>"
				      tempstr=tempstr & "       <td  width='100'>操作</td>"
					  tempstr=tempstr & "       </tr>"
					  tempstr=tempstr & "       <tr><td colspan='30' id='alist'>"
					  dim row:row=0
					   If IsArray(ACartArr) Then
					    Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
						RSC.Open "Select * From KS_ShopSpecificationPrice Where ProID=" & KS.CHkClng(KS.G("ID")) &" Order By Id",conn,1,1
						tempstr=tempstr & "<table width='98%' class='attbox' cellspacing='0' cellpadding='0'>"
						Do While Not RSC.Eof
						  row=row+1
						  Dim imgstr,attr1value:attr1value=RSC("attr1")
						  tname=split(attr1value,"|")(0)
						  if split(attr1value,"|")(1)<>"" then
						   imgstr="<img src='" & split(attr1value,"|")(1) & "' width='25' height='25' title='" & tname & "' align='absmiddle'/><input type='hidden' name='attrimg" & row & "0' id='attrimg" & row & "0' value='"& split(attr1value,"|")(1) & "'/>"
						  else
						   imgstr=""
						  end if
						  tempstr=tempstr & "<tr id=""row" & row &"""><td class=""splittd""><input type=""hidden"" name=""id" & row & """ value=""" & rsc("id") & """/><input type=""text"" value=""" & RSC("itemNo") &""" name=""aitemno" & row & """ id=""aitemno" & row & """ size=""10"" class=""textbox""/></td><td width=""150"" class=""splittd"">" & imgstr &" <input type=""text"" class=""ttt"" name=""attr" &row &"0"" value=""" & tname &"""/></td>"
						  if ubound(ACartArr)>=1 then
						   dim attr2value:attr2value=RSC("attr2")
						   tname=split(attr2value,"|")(0)
						  if split(attr2value,"|")(1)<>"" then
						   imgstr="<img src='" & split(attr2value,"|")(1) & "' width='25' height='25' title='" & tname & "' align='absmiddle'/><input type='hidden' name='attrimg" & row & "1' id='attrimg" & row & "1' value='"& split(attr2value,"|")(1) & "'/>"
						  else 
						   imgstr=""
						  end if
						  tempstr=tempstr & "<td width=""150"" class=""splittd"">" & imgstr & " <input type=""text"" class=""ttt"" name=""attr" & row & "1"" value=""" & tname &"""/></td>"
						   if ubound(ACartArr)>=2 then
							   dim attr3value:attr3value=RSC("attr3")
							   tname=split(attr3value,"|")(0)
							  if split(attr3value,"|")(1)<>"" then
							   imgstr="<img src='" & split(attr3value,"|")(1) & "' width='25' height='25' title='" & tname & "' align='absmiddle'/><input type='hidden' name='attrimg" & row & "2' id='attrimg" & row & "2' value='"& split(attr3value,"|")(1) & "'/>"
							  else
							   imgstr=""
							  end if
							   tempstr=tempstr & "<td width=""150"">" & imgstr &" <input type=""text"" class=""ttt"" name=""attr" & row & "2"" value=""" & tname &"""/></td>"
						   end if
						  end if
						  tempstr=tempstr & "<td width=""100"" class=""splittd""><input type=""text"" name=""aprice" & row & """ value=""" & RSC("Price") &""" size=""4"" class=""textbox""/>元</td><td width=""100"" class=""splittd""><input type=""text"" name=""aamount" & row & """ value=""" & RSC("Amount") & """ size=""4"" class=""textbox""/>件</td><td width=""100"" class=""splittd""><input type=""text"" name=""aweight" & row & """ value=""" & RSC("Weight") & """ size=""4"" class=""textbox""/>KG</td><td width=""100"" class=""splittd""><a href=""javascript:delrowajax(" & row &"," & rsc("id") &");"">删除</a></td></tr>"
						RSC.MoveNext
						Loop
						RSC.Close:Set RSC=Nothing
						tempstr=tempstr &"</table>"
					   End If
					 tempstr=tempstr &"</td></tr>"
					 tempstr=tempstr & "      </table>"
					 tempstr=tempstr & " <input type='hidden' name='AttributeCart' id='AttributeCart' value='" & AttributeCart &"'/></dd>"
				else
					 tempstr=tempstr & "<input type='button' value=' 开启规格属性 ' onclick=""getSpecification()"" class='button' /> tips:当一件商品有不同规格供选择或是对应不同价格时，可以启用此项。</dd><dd><font id='cartattr'></font></dd><input type='hidden' name='AttributeCart' id='AttributeCart' value=''/>"
	 
				end if
					tempstr=tempstr & "<input type='hidden' name='totalrow' id='totalrow' value='" & row &"'/>"
                    response.write tempstr
					
		   case "downurl"
		       		.Write "<dd>"
					.Write " <div>" & XTitle &":</div>"
					.Write " <input type='text' class='textbox' name='DownUrl' id='DownUrl' value='" & DownUrl & "' size='50'/>"
					.Write "       &nbsp;<input class=""button"" type='button' name='Submit' value='选择文件...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=" & CurrPath & "',550,290,window,document.myform.DownUrl);""><span>说明：如果这里的下载地址不为空，则用户购买成功后，会出现下载图标让用户下载！！</span>"
					.Write "</dd>" &vbcrlf

         End Select
		End If
   Next

        .Write "</dl>"
		.Write "</div>"
  Next
END IF
		
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='prointro']/showonform").text="1" Then
		.Write " <div class=tab-page id=intro-page>"
		.Write "  <H2 class=tab>商品介绍</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""intro-page"" ) );"
		.Write "	</SCRIPT>"

        .Write "<dl class='dtable'>"
		.Write "      <dd><div>" & KS.C_S(ChannelID,3) & "简介:<font>(<label><input name='BeyondSavePic' type='checkbox' value='1' id='BeyondSavePic'>自动下载简介里的图片</label>)</font></div>"
		.Write "<table border='0' width='90%' cellspacing='0' cellpadding='0'>"
		.Write "<tr><td height='30' width=70>&nbsp;<strong>附件上传:</div><td><iframe id='upiframe' name='upiframe' src='../../user/BatchUploadForm.asp?UPFrom=Admin&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='620' height='24'></iframe></td></tr>"
		.Write "</table>"
		.Write EchoEditor("Content",ProIntro,"Basic","90%","300px")
		
		.Write "      </dd>"
        .Write "</dl>"
		.Write "</div>"
	  END IF

	  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadphotos']/showonform").text="1" Then
		.Write " <div class=tab-page id=photo-page>"
		.Write "  <H2 class=tab>组图上传</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""photo-page"" ) );"
		.Write "	</SCRIPT>"

        .Write "<div style='clear:both'></div><dl class='dtables'>"
		.Write "      <dd><div>组图上传:<font>(<input type='checkbox' value='1' name='BeyondSavePic1' checked>采集存图)"
		%>
		    		<label<%if KS.TBSetting(5)="0" then response.write " style='display:none'"%>><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)"/>添加水印</label>
</font></div>
			<style type="text/css">
			#thumbnails{background:url(../../plus/swfupload/images/albviewbg.gif) no-repeat;min-height:200px;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
			#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:162px;height:190px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
			#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px ;margin:5px;}
			#thumbnails div.thumbshow div span{ padding-left:20px; line-height:20px; height:20px;}
			#thumbnails div.thumbshow div span a{ color:#006699}
			</style>
			<link href="../../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
			<script type="text/javascript" src="../../plus/swfupload/swfupload/swfupload.js"></script>
			<script type="text/javascript" src="../../plus/swfupload/js/handlers.js"></script>
			<script type="text/javascript" src="../../KS_inc/boxtcshow.js"></script>
<script type="text/javascript">
		var swfu;
		var pid=0;
		function SetAddWater(obj){if (obj.checked){swfu.addPostParam("AddWaterFlag","1");}else{swfu.addPostParam("AddWaterFlag","0");}}
		//删除已经上传的图片
		function DelUpFiles(pid,picid){
		 if (confirm('删除后不可恢复，确认删除吗？')){
			var p=$('#pic'+pid).val();
			   if (p!==''){
				$.ajax({
				  url: "../../plus/ajaxs.asp",
				  cache: false,
				  data: "action=DelPhoto&pic="+p+"&picid="+picid,
				  success: function(r){
				  }});
			   }
			   $("#thumbshow"+pid).remove();
			}
		}	
		
		function addImage(bigsrc,smallsrc,text,picid) {
			if (picid==undefined) picid=0;
			if (smallsrc=='') smallsrc=bigsrc;
			var newImgDiv = document.createElement("div");
			var delstr = '';
			delstr = '<a href="javascript:DelUpFiles('+pid+','+picid+')" style="color:#ff6600">[删除]</a>';
			newImgDiv.className = 'thumbshow';
			newImgDiv.id = 'thumbshow'+pid;
			document.getElementById("thumbnails").appendChild(newImgDiv);
			newImgDiv.innerHTML = '<a href="'+bigsrc+'" target="_blank"><span id="show'+pid+'"><strong><img src="'+smallsrc+'" /></strong></span></a>';
			newImgDiv.innerHTML += '<div style="margin-top:10px;text-align:left">'+delstr+' <b>组名：</b><input type="hidden" class="pics textbox" id="pic'+pid+'" value="'+bigsrc+'|'+smallsrc+'|'+picid+'" style="width:155px;"/><input class="textbox" type="text" name="picinfo'+pid+'" value="'+text+'" style="width:155px;" /> <span><a  title="左移动排序" href="javascript:;" onclick="pic_move(this,1);">←左移动</a>&nbsp;&nbsp;&nbsp;<a title="右移动排序" href="javascript:;" onclick="pic_move(this,2);">右移动→</a></span></div>';
			pid++;
			
		}
	
		window.onload = function () {
			swfu = new SWFUpload({
				// Backend Settings
				upload_url: "../include/swfupload.asp",
				post_params: {"AdminID":"<%=KS.C("AdminID") %>","AdminPass":"<%=KS.C("PassWord")%>",AddWaterFlag:"0",UPType:"ProImage","BasicType":<%=KS.C_S(ChannelID,6)%>,"ChannelID":<%=ChannelID%>,"AutoRename":4},

				// File Upload Settings
				file_size_limit : "2 MB",	// 2MB
				file_types : "*.jpg; *.gif; *.png",
				file_types_description : "图片格式,可以多选",
				file_upload_limit : 0,

				// Event Handler Settings - these functions as defined in Handlers.js
				//  The handlers are not part of SWFUpload but are part of my website and control how
				//  my website reacts to the SWFUpload events.
				swfupload_preload_handler : preLoad,
				swfupload_load_failed_handler : loadFailed,
				file_queue_error_handler : fileQueueError,
				file_dialog_complete_handler : fileDialogComplete,
				upload_start_handler : uploadStart,
				upload_progress_handler : uploadProgress,
				upload_error_handler : uploadError,
				upload_success_handler : uploadSuccess,
				upload_complete_handler : uploadComplete,

				// Button Settings
				//button_image_url : "../../plus/swfupload/images/SmallSpyGlassWithTransperancy_17x18d.png",
				button_placeholder_id : "spanButtonPlaceholder1",
				button_width: 195,
				button_height: 20,
				button_text : '<span class="button">本地批量上传(单图限制2 MB)</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#ffffff;font-size: 12px; } ',
				button_text_top_padding: 3,
				button_text_left_padding: 0,
				button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
				button_cursor: SWFUpload.CURSOR.HAND,
				
				// Flash Settings
				flash_url : "../../plus/swfupload/swfupload/swfupload.swf",
				flash9_url : "../../plus/swfupload/swfupload/swfupload_FP9.swf",

				custom_settings : {
					upload_target : "divFileProgressContainer1"
				},
				
				// Debug Settings
				debug: false
			});
		};
	</script>
	<script type="text/javascript">
	var input;
	var box='';
	function OnlineCollect(){
		box=$.dialog.open('../../editor/ksplus/remotefile.asp',{title:"网上采集图片",width:550,height:200});
	}
	function AddTJ(){
	 box=$.dialog({title:"从上传文件中选择",content:"<div style='padding:3px'><strong>小图地址:</strong><input class='textbox' type='text' name='x1' id='x1'> <input type='button' onclick=\"OpenModalDialog('../Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,$('#x1')[0]);\" value='选择小图' class='button'/><br/><strong>大图地址:</strong><input class='textbox' type='text' name='x2' id='x2'> <input type='button' onclick=\"OpenModalDialog('../Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,$('#x2')[0]);\" value='选择大图' class='button'/><br/><strong>简要介绍:</strong><input class='textbox' type='text' name='x3' id='x3'></div>",init: function(){
						},ok: function(){ 
						 var x1=this.DOM.content[0].getElementsByTagName('input')[0].value;
						 var x2=this.DOM.content[0].getElementsByTagName('input')[2].value
						 var x3=this.DOM.content[0].getElementsByTagName('input')[4].value
						   ProcessAddTj(x1,x2,x3);
						   return false; 
						}, 
						cancelVal: '关闭', 
						cancel: true });
	}
	function ProcessAddTj(x1,x2,x3){
					  if (x1==''){
					   alert('请选择一张小图地址!');
					   return false;
					  }
					  if (x2==''){
					   alert('请选择一张大图地址!');
					   return false;
					  }
					  addImage(x2,x1,x3,"")
					   box.close();
	}
	function ProcessCollect(collecthttp){
	 if (collecthttp==''){
	   alert('请输入远程图片地址,一行一张地址!');
	   return false;
	 }
	 var carr=collecthttp.split('\n');
	 for(var i=0;i<carr.length;i++){
	   if (carr[i]!=''){
	   var bigsrc=carr[i];
	   var smallsrc=carr[i];
	   addImage(bigsrc,smallsrc,'',0);
	   }
	 }
	 box.close();
	}
	</script>
	    <table>
		 <tr>
		  <td>

	    <div class="button"><span id="spanButtonPlaceholder1">上传图片</span></div>
		 </td>
		 <td>
		 <button type="button"  class="button" onClick="OnlineCollect()">网上采集</button>&nbsp;
		 <button type="button"  class="button" onClick="AddTJ();">图片库...</button>
		 </td>
		 </tr>
		</table>
		<table width="100%">
			 <tr>
			   <td id="divFileProgressContainer1"></td>
			 <tr>
			 <tr>
			   <td id="thumbnails"></td>
			 <tr>
			</table>

			<input type='hidden' name='PicUrls' id='PicUrls'>

		<%
		.Write "       </dd>"
        .Write "</dl>"
		.Write "</div>"
		
        If Action = "Edit" Or Action="Verify" Then
		   .Write "<script type=""text/javascript"">" & vbcrlf
		   Dim RSS:Set RSS=Conn.Execute("Select * From KS_ProImages Where ProID=" & ID &" order by orderid,id")
		   Do While Not RSS.Eof
		    .Write "addImage('" & RSS("BigPicUrl") & "','" & RSS("SmallPicUrl") & "','" & RSS("GroupName") & "'," & rss("id") &");" &vbcrlf
		   RSS.MoveNext
		   Loop
		   .Write "</script>"
		   RSS.Close :Set RSS=Nothing
		End If
    End If
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then	
		.Write " <div class=tab-page id=option-page>"
		.Write "  <H2 class=tab>属性设置</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""option-page"" ) );"
		.Write "	</SCRIPT>"
        .Write "<dl class=""dtable"">" & vbcrlf
       If KS.GetAppStatus("special") Then	
				Call KSCls.Get_KS_Admin_Special(ChannelID,KS.ChkClng(KS.G("ID")))
		End If
			
		.Write " <dd>"
		.Write "  <div>购买限制:</div>"
		.Write "  同个IP一天内只允许游客购买<input style='text-align:center' name='visitornum' type='text' id='visitornum' value='" & visitornum & "' size='6' class='textbox'>件 允许注册会员购买<input style='text-align:center' name='membernum' type='text' id='membernum' value='" & membernum & "' size='6' class='textbox'>件 <span>不限制请输入0</span><br/>允许购买此商品的用户组(不限制请不要勾选)<br/>"
		.Write KS.GetUserGroup_CheckBox("GroupID",arrGroupID,5)
		
		.Write "   </dd>"
			
		.Write "      <dd><div>关 键 字:</div>"
		.Write "      <input name='KeyWords' type='text' id='KeyWords' class='textbox' value='" & KeyWords & "' size=50> <<"
		.Write "                  <select name='SelKeyWords' style='width:100px' onChange='InsertKeyWords($(""#KeyWords"").get(0),this.options[this.selectedIndex].value)'>"
		.Write "<option value="""" selected> </option><option value=""Clean"" style=""color:red"">清空</option>"
		.Write KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")
		.Write "                  </select>"
		.Write " 【<a href=""javascript:;"" id=""KeyLinkByTitle"" style=""color:green"">根据商品名称自动获取Tags</a>】<input type='checkbox' name='tagstf' value='1' checked>记录"
		.Write "   </dd>"
		.Write "   <dd>"
		.Write "     <div>" & KS.C_S(ChannelID,3) & "型号:</div>"
		.Write "    <input name='ProModel' type='text' id='ProModel' value='" & ProModel & "' size=50 class='textbox'> << "
		.Write ("<select name='sProModel' style='width:100px' onChange=""$('#ProModel').val(this.options[this.selectedIndex].value)"">")
		.Write  KSCls.Get_O_F_D(KS.C_S(ChannelID,2),"Distinct Top 10 ProModel","1=1 Group by ProModel")
        .Write "</select>"
		.Write "</dd>"
		.Write "     <dd>"
		.Write "                <div>" & KS.C_S(ChannelID,3) & "规格:</div>"
		.Write "                <input name='ProSpecificat' type='text' id='ProSpecificat' value='" &ProSpecificat & "' size=50 class='textbox'> << "
		.Write ("<select name='selProSpecificat' style='width:100px' onChange=""$('#ProSpecificat').val(this.options[this.selectedIndex].value)"">")
		.Write  KSCls.Get_O_F_D(KS.C_S(ChannelID,2),"Distinct Top 10 ProSpecificat","1=1 Group by ProSpecificat")
        .Write "</select>"
		.Write "</dd>"
		.Write "  <dd><div>生 产 商:</div>"
		.Write "              <input name='ProducerName' type='text' id='ProducerName' value='" & ProducerName & "' size=50 class='textbox'>               << "
			.Write ("<select name='SelProducerName' style='width:100px' onChange=""$('#ProducerName').val(this.options[this.selectedIndex].value)"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=" & ChannelID &" And OriginType=1 Order BY AddDate Desc")
			.Write "       </select>"
			.Write "              </dd>"
			.Write "              <dd><div>商品商标:</div>"
			.Write "                <td><input name='TrademarkName' type='text' id='TrademarkName' value='" & TrademarkName & "' size=50 class='textbox'>                << "
			.Write ("<select name='selTrademarkName' style='width:100px' onChange=""$('#TrademarkName').val(this.options[this.selectedIndex].value)"">")
		    .Write  KSCls.Get_O_F_D(KS.C_S(ChannelID,2),"Distinct Top 10 TrademarkName","1=1 Group by TrademarkName")
			.Write "                </select>"
			.Write "              </dd>"
			.Write "              <dd>"
			.Write "                <div>上市时间:</div>"
		If Action <> "Edit" Then
		.Write ("<input name='AddDate' type='text' id='AddDate' value='" & Now() & "' size='60'  onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" class='textbox Wdate'>")
		Else
		.Write ("<input name='AddDate' type='text' id='AddDate' value='" & AddDate & "' size='60'  onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" class='textbox Wdate'>")
		End If
		.Write "                <b>日期格式：年-月-日 时：分：秒</b>"
			.Write "             </dd>"
			.Write "            <dd>"
			.Write "                <div>服务期限:</div>"
			.Write "             <input name='ServiceTerm' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" type='text' id='ServiceTerm' value='" & ServiceTerm & "' size=8 class='textbox'>年"
			.Write "              </dd>"
			.Write "              <dd>"
			.Write "                <div>商品等级:</div>"
			.Write "               <select name='rank'>"
			If Rank = "★" Then
			.Write "                    <option  selected>★</option>"
			Else
			.Write "                    <option>★</option>"
			End If
			If Rank = "★★" Then
			.Write "                    <option  selected>★★</option>"
			Else
			.Write "                    <option>★★</option>"
			End If
			If Rank = "★★★" Or Action = "Add" Then
			.Write "                    <option  selected>★★★</option>"
			Else
			.Write "                    <option>★★★</option>"
			End If
			If Rank = "★★★★" Then
			.Write "                    <option  selected>★★★★</option>"
			Else
			.Write "                    <option>★★★★</option>"
			End If
			If Rank = "★★★★★" Then
			.Write "                    <option  selected>★★★★★</option>"
			Else
			.Write "                    <option>★★★★★</option>"
			End If
			.Write "                  </select>"
			.Write "                  <span>请为商品评定推荐等级</span>"
			.Write "               </dd>"
			.Write "               <dd>"
			.Write "               <div>点 击 数:</div>"
			.Write "               本日：<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='10' class='textbox'> 本周：<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='10' class='textbox'> 本月：<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='10' class='textbox'> 总计：<input name='Hits' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" type='text' id='Hits' value='" & Hits & "' size='10' class='textbox'>次"
			.Write "</dd>"
			 .Write "             <dd><div>" & KS.C_S(ChannelID,3) & "模板:</div>"
			 .Write "<table>"
			IF Action <> "Edit" and  Action<>"Verify" Then
			.Write " <tr><td><input type='radio' name='templateflag' onclick='GetTemplateArea(false);' value='2' checked>继承栏目设定<input type='radio' onclick='GetTemplateArea(true);' name='templateflag' value='1'>自定义</td></tr>"
			.Write "<tr id='templatearea' style='display:none'><td>"
		    Else
			.Write "<tr id='templatearea'><td>"
			End If
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly size=50 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>3G版模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=50 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</td></tr></table>"
			
			.Write "                </dd>"
			.Write "             <dd><div>文 件 名:</div>"
			IF Action = "Edit" or Action="Verify" Then
			.Write "<input name='FileName' type='text' id='FileName' readonly  value='" & FileName & "' size='50' class='textbox'> <span>不能改</span>"
			Else
			.Write "<table>"
			.Write "<tr><td><input type='radio' value='0' name='filetype' onclick='GetFileNameArea(false);' checked>自动生成 <input type='radio' value='1' name='filetype' onclick='GetFileNameArea(true);' >自定义</td></tr>"
			.Write "<tr id='filearea' style='display:none;font-weight:normal'><td><input name='FileName' type='text' id='FileName'   value='" & FileName  & "' size='45' class='textbox'> <font class=""tips"">可带路径,如 help.html,news/news_1.shtml等</font></td></tr>"
			.Write "</table>"
			End IF
			 .Write "             </dd>"
			 .Write "              <dd><div>审核状态:</div>"
			If KS.C("Role")="1" Then   '发稿员
				.Write "<input name='verific' type='radio' value='0'"
				if verific=0 or Action="Add"  then .write " checked"
				.write ">待审核"
				If Action="Edit" Then
				.Write "<input name='verific' type='radio' value='100' checked>保持原状态"
				End If
			Else
			
				if KS.C("Role")="2" Then
					.Write "<input name='verific' type='radio' value='0'"
					if verific=0   then .write " checked"
					.write ">待审核"
					
					If KS.ChkClng(Split(KS.C_S(ChannelID,46)&"||||","|")(25)) = 1 Then
						.write "<input type='radio' name='verific' value='5'"
						if verific=5 or Action="Add"  or action="Verify" then .write "checked"
						.write ">初审通过"
					Else
						.write "<input type='radio' name='verific' value='1'"
						if verific=1 or action="Add"  or action="Verify" then .write "checked"
						.write ">审核通过"
					End If
					
					if action="Verify" Then
					.Write "<input name='verific' type='radio' value='3'"
					if verific=3   then .write " checked"
					.write ">退稿"
					End If
					
					If Action="Edit" Then
					 .Write "<input name='verific' type='radio' value='100' checked>保持原状态"
					End If
				Elseif KS.C("Role")="3" Then 
				    .Write "<input name='verific' type='radio' value='0'"
					if verific=0   then .write " checked"
					.write ">待审核"
				
					.write "<input type='radio' name='verific' value='1'"
					if verific=1 or action="Add"  or action="Verify" then .write "checked"
					.write ">终审通过"
					
					if action="Verify" Then
					.Write "<input name='verific' type='radio' value='3'"
					if verific=3   then .write " checked"
					.write ">退稿"
					End If
					
					If Action="Edit" Then
					 .Write "<input name='verific' type='radio' value='100' checked>保持原状态"
					End If
				end if
            End If
			
			.Write "        </dd>"
			 .Write "    </dl>"
			 .Write "    </div>"
	  End If		
	  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/showonform").text="1" Then	
		KSCls.LoadSeoOption ChannelID,"SEO优化选项",SEOTitle,SEOKeyWord,SEODescript
	  End If
	  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='kboption']/showonform").text="1" Then	
		.Write " <div class=tab-page id=kbxs-page>"
		.Write "  <H2 class=tab>捆绑销售</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""kbxs-page"" ) );"
		.Write "	</SCRIPT>"

        .Write "<table style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
		%>
		  <tr class="tdbg">
			<td  style="text-align:left">
			 <script type="text/javascript">
		  function getProduct()
		  {			 
		     $(parent.document).find("#ajaxmsg").toggle("fast");
			 var key=escape($('input[name=key]').val());
			 var kbtid=$('#kbtid>option:selected').val();
			 var priceType=$('#PriceType>option:selected').val();
			 var minPrice=$("#minPrice").val();
			 var maxPrice=$("#maxPrice").val();
			 var str='';
			 if (key!=''){
			   str='商品名称:'+key;
			 } 
			 if (kbtid!='' && kbtid !='0'){
			   str+=' 栏目:'+$('#kbtid>option:selected').get(0).text
			 }
			 if (priceType!=0){
			   str+= minPrice +' 元';
			   switch (parseInt(priceType)){
			     case 1 :
				  str+='<=参考价<=';
				  break;
			     case 2 :
				   str+='<=VIP价<=';
				   break;
			     case 3 :
				  str+='<=商城价<=';
				  break;
			   }
			   str+= maxPrice +' 元';
			   
			 }
			 if (str!='') str='<strong>条件:</strong><font color=red>'+str+'</font>';
			 $("#keyarea").html(str);
			 
			 $.get("../../plus/ajaxs.asp", { action: "GetPackagePro", proid:$("#proids").val(),id:<%=ks.chkclng(request("id"))%>,pricetype:priceType,key: key,tid:kbtid,minPrice:minPrice,maxPrice:maxPrice},
			 function(data){
					$(parent.document).find("#ajaxmsg").toggle("fast");
					$("#prolist").empty().append(data);
			  });
		  }
		  function addProductIn(){
		   var proid=$('#prolist option:selected').val();
		   if (proid!=undefined){
		      top.openWin('设置捆绑销售价格','shop/KS.Shop.asp?action=SetKBXSPrice&proid='+proid+'&rnd='+Math.random(),false,350,200)
		    }else{
			 alert('请选择要加入捆绑销售的商品!');
			}
		  }
		  function updateKBXS(arrstr)
		  {
			  if (arrstr!=''){
				 var finder=false;
				  var arr=arrstr.split('@@@');
				 $('#kbprolist>option').each(function(){
				 if (arr[0]==this.value.split('|')[0]){
				   this.text=arr[2]+"(捆绑销售价:￥"+arr[1]+"元)";
				   this.value=arr[0]+"|"+arr[1];
				   finder=true;}
			     });
				  if (finder==false){
					$('#kbprolist').append("<option value='"+arr[0]+"|"+arr[1]+"' selected>"+arr[2]+"(捆绑销售价:￥"+arr[1]+"元)</option>");
				 }
				 
			}
		  
		  }
		  function modifyKBXS(){
		   var l=$('#kbprolist>option:selected').length;
		   if (l==1){
		    var kb=$('#kbprolist>option:selected').val().split('|');
			 top.openWin('设置捆绑销售价格','shop/KS.Shop.asp?action=SetKBXSPrice&proid='+kb[0]+'&kbprice='+kb[1]+'&rnd='+Math.random(),false,350,200);
		   }else if(l==0){
		    alert('请选择一个商品!');
		   }else{
		    alert('一次只能选择一个商品!');
		   }
		  }
		  function delAllKBXS(){
		    $("#kbprolist").empty();
		  }
		  function delSelectKBXS(){
		      var dest = document.getElementById('kbprolist');
			  for (var i = dest.options.length - 1; i >= 0 ; i--)
					  {
						  if (dest.options[i].selected)
						  {
							  dest.options[i] = null;
						  }
					  }
		  }
		  function selectAllKBXS(){
		   $("#kbprolist option").each(function(){
		      $(this).attr("selected",true);
		   });
		  }
		</script>
			  &nbsp;<strong>快速搜索=></strong>
			  <br/>
			  &nbsp;商品编号: <input type="text" class="textbox" name="proids" id="proids" size='15'> 可留空<br/>
			 &nbsp;商品名称: <input type="text" class='textbox' name="key">
			 <br/>&nbsp;所属栏目: <select size='1' name='kbtid' id='kbtid'><option value=''>--栏目不限--</option><%=KS.LoadClassOption(ChannelID,false)%></select>
			 <br/>&nbsp;价格范围:
			<input type='text' name='minPrice' class="textbox" size='5' style='text-align:center' id='minPrice' value='10'> 元
			<= <select name="PriceType" id="PriceType">
			  <option value=0>--不限制--</option>
			  <option value=1>参考价</option>
			  <option value=2>VIP价</option>
			  <option value=3>商城价</option>
			 </select>
			 <= <input type='text' name='maxPrice' class="textbox" size='5' style='text-align:center' id='maxPrice' value='100'> 元
			  
			  <br/> <br/>
			  &nbsp;<input type="button" onClick="getProduct()" value="开始搜索" class="button" name="s1">
			
			</td>
			<td  style="text-align:left">
			<div id='keyarea'></div>
			<strong>查询到的商品:</strong>			
			<br/>
			 <select name="prolist" size="5" style="width:260px;height:140px" id="prolist"></select>
			 <br/>
			 <input type="button" onClick="addProductIn()" value="将选中的商品加入捆绑销售" class="button">
			</td>
		  </tr>
		  <tr class="tdbg">
		    <td>
			  <strong>捆绑销售商品:</strong><br/>
			  <select name="kbprolist" id="kbprolist" multiple size="6" style="width:360px;height:160px">
			  <%
			   if id<>0 then
			     Dim RSK:Set RSK=Conn.Execute("Select I.Title,i.id,K.KBPrice From KS_ShopBundleSale k Inner Join KS_Product I on i.id=k.kbproid where k.proid=" & id)
				 do while not rsk.eof
				   response.write "<option value='" & rsk(1) & "|" & rsk(2) & "' selected>" & rsk(0) & "(捆绑销售价:￥" & rsk(2) & "元)</option>"
				 rsk.movenext
				 loop
				 rsk.close : set rsk=nothing
			   end if
			  %>
			  </select>
			</td>
			<td>
			<input type="button" class="button" value="修改选中商品价格" onClick="modifyKBXS()"/><br/><br/>
			<input type="button" class="button" value="移除选中商品" onClick="delSelectKBXS()"/><br/><br/>
			<input type="button" class="button" value="全部移除" onClick="delAllKBXS()" /><br/><br/>
			<input type="button" class="button" value="全部全中" onClick="selectAllKBXS()" />
			</td>
		  </tr>
		<%
		.Write "</table>"
		.Write "</div>"
     End If
	 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='relativeoption']/showonform").text="1" Then	 
			 KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))
	 End If
	 			 .Write " </div>"

			 .Write "   </form>"
			 .Write "</body>"
			 .Write "</html>"
			 End With
			 if isobject(rs) then
			 if rs.state=1 then rs.close:Set rs=nothing
			 end if
		End Sub
		
		'保存
		Sub DoSave()
		  Dim SelectInfoList,HasInRelativeID,FileIds,WapTemplateID,Relateda,Related_s,ii
		  With Response
		    ProductID = KS.ChkClng(Request("ProductID"))
		    ProID       = KS.G("ProID")
			If ProID="" Then ProID = KS.GetInfoID(ChannelID)
			Title       = KS.G("Title")
			PhotoUrl    = KS.G("PhotoUrl")
			BigPhoto    = KS.G("BigPhoto")
			ProIntro    = Request.Form("Content")
			Hits        = KS.ChkClng(KS.G("Hits"))
			HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
			HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
			HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))
			TotalNum    = KS.ChkClng(KS.G("TotalNum"))
			AlarmNum    = KS.ChkClng(KS.G("AlarmNum"))
			Unit        = KS.G("Unit")
			Weight      = KS.G("Weight")
			If Not IsNumeric(Weight) Then KS.AlertHintScript "商品重量只能输入数字!"
			AttributeCart  = Replace(Request.Form("AttributeCart"),vbcrlf,"§")
			Price = KS.G("Price"):If Not IsNumeric(Price) Then Price=0
			Price_Member = KS.G("Price_Member"):If Price_Member="" Then Price_Member=0
			VipPrice     = KS.G("VipPrice"):If VipPrice="" Then VipPrice=0
			Recommend   = KS.ChkClng(KS.G("Recommend"))
			Rolls       = KS.ChkClng(KS.G("Rolls"))
			Popular     = KS.ChkClng(KS.G("Popular"))
			If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then
						Verific=KS.ChkClng(KS.G("Verific"))
			Else
						Verific = 1
			End if
			Comment     = KS.ChkClng(KS.G("Comment"))
			Strip       = KS.ChkClng(KS.G("Strip"))
			Slide       = KS.ChkClng(KS.G("Slide"))
			BrandID     = KS.ChkClng(KS.G("Brandid"))
			IsSpecial   = KS.ChkCLng(KS.G("IsSpecial"))
			IsTop       = KS.ChkClng(KS.G("IsTop"))
			IsDiscount  = KS.ChkClng(KS.G("IsDiscount"))
			IsScore     = KS.ChkClng(KS.G("IsScore"))
			Score     = KS.ChkClng(KS.G("Score"))
			SpecialID = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
			SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
			Makehtml    = KS.ChkClng(KS.G("Makehtml"))
			FnameType = Trim(KS.G("Fnametype"))
			Tid = KS.G("Tid")
			Relateda  =KS.G("tidtb")
			if  not ks.isnul(Relateda) then 
				 	Relateda=Replace( Replace(Relateda&"",","&Tid ,""),Tid&",","")
					if (tid<>"" and tid<>"0") then
					Relateda=Tid &","& Relateda
					end if
					Relateda=Split(Relateda,",")
					if (tid="" or tid="0") then tid=Relateda(0)

			else	
					Relateda=Array(Tid)
			end if
			If KS.ChkClng(KS.C_C(Tid,20))=0 Then
				 Response.Write "<script>alert('对不起,系统设定不能在此栏目发表,请选择其它栏目!');history.back();</script>":Exit Sub
			End IF
			
				TemplateID = KS.G("TemplateID")
				WapTemplateID=KS.G("WapTemplateID")
				Dim FnameType:FnameType=KS.C_C(TID,23)
				 If KS.ChkClng(KS.G("filetype"))=0 Then
					If Action = "Add" OR Action="Verify" Then
						Fname=KS.GetFileName(KS.C_C(TID,24), Now, FnameType)
					 End If
				 Else
				     Fname=KS.G("FileName")
				 End If
				 If KS.ChkClng(KS.G("TemplateFlag"))=2 Or TemplateID="" Then TemplateID=KS.C_C(TID,5):WapTemplateID=KS.C_C(TID,22)
           Call KSCls.CheckDiyField(FieldXML,ErrMsg)  '检查自定义字段
					
			KeyWords      = KS.G("KeyWords")
			ProducerName  = KS.G("ProducerName")
			TrademarkName = KS.G("TrademarkName")
			ProModel      = KS.G("ProModel")
			ProSpecificat = KS.G("ProSpecificat")
			ServiceTerm   = KS.ChkClng(KS.G("ServiceTerm"))
			AddDate       = KS.G("AddDate")
			Rank          = Trim(KS.G("Rank"))	
			IsChangedBuy  = KS.ChkClng(KS.G("IsChangedBuy"))
			ChangeBuyNeedPrice=KS.G("ChangeBuyNeedPrice"): If Not IsNumeric(ChangeBuyNeedPrice) Then ChangeBuyNeedPrice=0
			ChangeBuyPresentPrice=KS.G("ChangeBuyPresentPrice"): If Not IsNumeric(ChangeBuyPresentPrice) Then ChangeBuyPresentPrice=0
			visitornum    = KS.ChkClng(KS.G("visitornum"))
			membernum     = KS.ChkClng(KS.G("membernum"))
			IsLimitbuy    = KS.ChkClng(KS.S("IsLimitBuy"))
			LimitBuyPrice = KS.S("LimitBuyPrice")
			LimitBuyAmount= KS.ChkCLng(KS.S("LimitBuyAmount"))
			LimitBuyTaskID=KS.ChkCLng(KS.S("LimitBuyTaskID"))
			If IsLimitBuy<>0 And LimitBuyTaskID=0 Then ErrMsg = ErrMsg & "请选请择抢购任务! \n"
			iF IsLimitbuy<>0 and LimitBuyAmount>TotalNum Then ErrMsg=ErrMsg & "抢购数量必须小于等于商品数量!\n"
			If Not IsNumeric(LimitBuyPrice) Then LimitBuyPrice=0
			SEOTitle      = KS.G("SEOTitle")
			SEOKeyWord    = KS.G("SEOKeyWord")
			SEODescript   = KS.G("SEODescript")
			FreeShipping  = KS.ChkClng(KS.G("FreeShipping"))
			WholesaleNum  = KS.ChkClng(KS.G("WholesaleNum"))
			WholesalePrice  = KS.G("WholesalePrice")
			if Not Isnumeric(WholesalePrice) Then WholesalePrice=0
					 
			If Title = "" Then .Write ("<script>alert('商品名称不能为空!');history.back(-1);</script>")
			
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Tid = "" Then ErrMsg = ErrMsg & "[商品类别]必选! \n"
			If Title = "" Then ErrMsg = ErrMsg & "[商品标题]不能为空! \n"
			If Title <> "" And Tid <> "" And Action = "Add" Then
			  SqlStr = "select top 1 * from KS_Product where Title='" & Title & "' And Tid='" & Tid & "'"
			   RS.Open SqlStr, conn, 1, 1
				If Not RS.EOF Then
				 ErrMsg = ErrMsg & "该类别已存在此商品! \n"
			   End If
			   RS.Close
			End If
			If ProductID=0 Then
			  If Not Conn.Execute("Select top 1 ProID From KS_Product Where ProID='" & ProID & "'").eof Then
			    ErrMsg = ErrMsg & "该商品编号已存在! \n"
			  End If
			Else
			  If Not Conn.Execute("Select top 1 ProID From KS_Product Where ID<>" & ProductID & " and ProID='" & ProID & "'").eof Then
			    ErrMsg = ErrMsg & "该商品编号已存在! \n"
			  End If
			End If
			
			If ErrMsg <> "" Then
			   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
			   .End
			Else
			FileIds=LFCls.GetFileIDFromContent(ProIntro)
			
			    If KS.ChkClng(KS.G("BeyondSavePic")) = 1 Then
				    Dim SaveFilePath
					SaveFilePath = KS.GetUpFilesDir & "/"
					KS.CreateListFolder (SaveFilePath)
				    ProIntro= KS.ReplaceBeyondUrl(ProIntro, SaveFilePath)
				End If
				
				 
				
			
						
				If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)	
				  If Action = "Add" Then
				    for ii=0 to Ubound(Relateda)
					Set RS = Server.CreateObject("ADODB.RecordSet")
					SqlStr = "select top 1 * from KS_Product where 1=0"
					RS.Open SqlStr, conn, 1, 3
					RS.AddNew
					 If II=0 Then
				    	RS("ProID")       = ProID
					 Else
				    	RS("ProID")       = KS.GetInfoID(ChannelID)
					 End If
					RS("Title")       = Title
					RS("PhotoUrl")    = PhotoUrl
					RS("BigPhoto")    = BigPhoto
					RS("ProIntro")    = ProIntro&""
					RS("Recommend")   = Recommend
					RS("Rolls")       = Rolls
					RS("Popular")     = Popular
					RS("Verific")     = Verific
					RS("Strip")       = Strip
					RS("Comment")     = Comment
					   if ii=0 then
							RS("Tid")    = Tid
						else
							RS("Tid")    =RTrim(Trim(Relateda(ii)))
							Tid = RTrim(Trim(Relateda(ii))) 
						end if
					RS("oTid")        = KS.G("oTid")
					RS("oID")         = KS.ChkClng(KS.S("oid"))
					RS("TotalNum")    = TotalNum
					RS("AlarmNum")    = AlarmNum
					RS("IsDiscount")  = IsDiscount
					RS("Unit")        = Unit
					RS("Weight")      = Weight
					RS("Price")       = Price
					RS("Price_Member")= Price_Member
					RS("VipPrice")    = VipPrice
					RS("KeyWords")    = KeyWords
					RS("ProSpecificat")=ProSpecificat
					RS("ProModel")    = ProModel
					RS("ServiceTerm") = ServiceTerm
					RS("ProducerName")= ProducerName
					RS("TrademarkName") = TrademarkName
					RS("AddDate")     = AddDate
					RS("ModifyDate")  = AddDate 
					RS("Rank")        = Rank
					RS("Slide")       = Slide
					RS("IsTop")       = IsTop
					RS("IsSpecial")   = IsSpecial
					RS("Istype")      = IsScore
					If IsScore=0 Then 
					 RS("Score")       = 0
					Else
					 RS("Score")       = Score
					End If
					if ii=0 then
						 RS("TemplateID")     = TemplateID
					Else
						 RS("TemplateID")     = KS.C_C(TID,5)
					End If
					RS("WapTemplateID")  = WapTemplateID
					RS("Hits")        = Hits
					RS("HitsByDay")   = HitsByDay
					RS("HitsByWeek")  = HitsByWeek
					RS("HitsByMonth") = HitsByMonth
					RS("Fname")       = Fname
					RS("BrandID")     = BrandID
					RS("AttributeCart")=AttributeCart
					RS("IsChangedBuy")=IsChangedBuy
					RS("ChangeBuyNeedPrice")=ChangeBuyNeedPrice
					RS("ChangeBuyPresentPrice")=ChangeBuyPresentPrice
					RS("DownUrl")    = Request.Form("DownUrl")
					RS("ArrGroupID") = Request.Form("GroupID")
					RS("visitornum") = visitornum
					RS("MemberNum")  = MemberNum
					RS("IsLimitbuy") = IsLimitBuy
					RS("LimitBuyPrice") = LimitBuyPrice
					RS("LimitBuyTaskID")=LimitBuyTaskID
					RS("LimitBuyAmount")= LimitBuyAmount
                    RS("SEOTitle")   = SEOTitle
					RS("SEOKeyWord") = SEOKeyWord
					RS("SEODescript")= SEODescript
					RS("FreeShipping")=FreeShipping
					RS("WholesalePrice")=WholesalePrice
					RS("WholesaleNum")=WholesaleNum
					RS("Changes")=KS.ChkClng(request("changes"))
					RS("ChangesUrl")=Request.Form("ChangesUrl")
					RS("OrderID")  = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & Tid &"'")(0))+1
					RS("salenum")=0
					if Action="Verify" Then
						   RS("Inputer") = Trim(KS.G("Inputer"))
					Else
					       RS("Inputer") = KS.C("AdminName")
					End IF
					if KS.IsNul(KS.Setting(189))  then
						     RS("RefreshTF") = Makehtml
					else
							 RS("RefreshTF")  = 0
					end if
					RS("DelTF")     = 0
					RS("PostTable") = LFCls.GetCommentTable()
					Call KSCls.AddDiyFieldValue(RS,FieldXml)
					RS.Update
					RS.MoveLast
						dim RelatedID
						if ii=0 then
							if Ubound(Relateda)>0 then
								RS("RelatedID")=-11	
							end if
							RelatedID=RS("id")
						else
							RS("RelatedID")=RelatedID	
						end if
						RS.Update
				   Session("KeyWords") = KeyWords
				   Session("ProducerName") = ProducerName
				   Session("TrademarkName") = TrademarkName
				    RS.MoveLast
					  If Left(Ucase(Fname),2)="ID" Then
					   RS("Fname") = RS("ID") & FnameType
					   RS.Update
					  End If
					   Call SaveProAttr(rs("id"),0)
					   Call SaveProImages(rs("id"),0)
					   Call addBundleSale(rs("id"))
					   If ii=0 then
						For I=0 To Ubound(SpecialID)
						  Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
					    Next
					   End If
					    Call KSCls.UpdateRelative(ChannelID,RS("ID"),SelectInfoList,0)
						Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,Tid,ProIntro,KeyWords,PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,IsSpecial,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
					   '关联上传文件
				       Call KS.FileAssociation(ChannelID,RS("ID"),PhotoUrl & BigPhoto & ProIntro ,0)
                       If Not KS.IsNul(FileIds) Then 
					    Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & RS("ID") &",classID=" & KS.C_C(Tid,9) & " Where ID In (" & FileIds & ")")
					   End If
			        Call RefreshHtml(1)
					RS.Close:Set RS = Nothing
				 Next
					
					
				ElseIf Action = "Edit"  Or Action="Verify" Then
				   If Action="Verify" Then 
					 Call KS.ReplaceUserFile(ProIntro,ChannelID)
					 Call KS.ReplaceUserFile(PhotoUrl,ChannelID)
					 Call KS.ReplaceUserFile(BigPhoto,ChannelID)
					End If
					
					 dim RelatedArray,n:n=0
					if KS.ChkClng(KS.G("EditNewtb"))=1 then
						if KS.ChkClng(KS.G("RelatedID"))=0 or KS.ChkClng(KS.G("RelatedID"))=-11 then
							RelatedArray=KSCls.GetRelatedArray( KS.C_S(ChannelID,2), ProductID ,11)'同步文章
						else
							RelatedArray=KSCls.GetRelatedArray( KS.C_S(ChannelID,2), KS.ChkClng(KS.G("RelatedID")) ,22)'同步文章
						end if 
					else
						RelatedArray=Array(ProductID)	
					end if
					for ii=0 to UBound(RelatedArray)
					      Set RS=SERVER.CreateObject("ADODB.RECORDSET")
						   SqlStr = "SELECT Top 1 * FROM KS_Product Where ID=" & RTrim(Trim(RelatedArray(ii))) & ""
							RS.Open SqlStr, conn, 1, 3
							If RS.EOF And RS.BOF Then
							 .Write ("<script>alert('参数传递出错!');history.back(-1);</script>")
							 .End
							End If
							If KS.ChkClng(ProductID)= KS.ChkClng(Trim(RelatedArray(ii))) Then
							 RS("ProID")      = ProID
							End If
							RS("Title")      = Title
							RS("PhotoUrl")   = PhotoUrl
							RS("BigPhoto")   = BigPhoto
							RS("ProIntro")   = ProIntro&""
							RS("Recommend")  = Recommend
							RS("Rolls")      = Rolls
							RS("Strip")      = Strip
							RS("Popular")    = Popular
							If Verific<>100 Then
							 RS("Verific") = Verific
							End If
							RS("Comment")    = Comment
							if ProductID=KS.ChkClng(RTrim(Trim(RelatedArray(ii)))) then
								RS("Tid")       = KS.G("Tid")
								RS("oTid")      = KS.G("oTid")
								RS("oID")       = KS.ChkClng(KS.S("oid"))
						    end if
							RS("KeyWords")   = KeyWords
							RS("TotalNum")   = TotalNum
							RS("AlarmNum")   = AlarmNum
							RS("Unit")       = Unit
							RS("Weight")     = Weight
							RS("IsDiscount") = IsDiscount
							RS("Price")      = Price
							RS("Price_Member")=Price_Member
							RS("VipPrice")    = VipPrice
							RS("ProSpecificat")=ProSpecificat
							RS("ProModel")   = ProModel
							RS("ServiceTerm") = ServiceTerm
							RS("ProducerName") = ProducerName
							RS("TrademarkName") = TrademarkName
							RS("AddDate")    = AddDate
							RS("ModifyDate") = Now
							RS("Rank")       = Rank
							RS("Slide")      = Slide
							RS("IsTop")      = IsTop
							RS("IsSpecial")  = IsSpecial
							RS("Istype")     = IsScore
							If IsScore=0 Then 
							RS("Score")       = 0
							Else
							RS("Score")       = Score
							End If
							RS("TemplateID") = TemplateID
							RS("WapTemplateID")  = WapTemplateID
							RS("Fname") = Replace(RS("Fname"), Trim(Mid(Trim(RS("Fname")), InStrRev(Trim(RS("Fname")), "."))), FnameType)
							If Makehtml = 1 Then RS("RefreshTF") = 1
							RS("Hits")       = Hits
							RS("HitsByDay")  = HitsByDay
							RS("HitsByWeek") = HitsByWeek
							RS("HitsByMonth")= HitsByMonth
							RS("BrandID")    = BrandID
							RS("AttributeCart")=AttributeCart
							RS("IsChangedBuy")=IsChangedBuy
							RS("ChangeBuyNeedPrice")=ChangeBuyNeedPrice
							RS("ChangeBuyPresentPrice")=ChangeBuyPresentPrice
							RS("IsLimitbuy") = IsLimitBuy
							RS("LimitBuyPrice") = LimitBuyPrice
							RS("LimitBuyTaskID")=LimitBuyTaskID
							RS("LimitBuyAmount")= LimitBuyAmount
							RS("visitornum")= visitornum
							RS("MemberNum") = MemberNum
							RS("DownUrl")   = Request.Form("DownUrl")
							RS("ArrGroupID")=Request.Form("GroupID")
							RS("SEOTitle")   = SEOTitle
							RS("SEOKeyWord") = SEOKeyWord
							RS("SEODescript")= SEODescript
							RS("FreeShipping")=FreeShipping
							RS("WholesalePrice")=WholesalePrice
							RS("WholesaleNum")=WholesaleNum
							RS("Changes")=KS.ChkClng(request("changes"))
							RS("ChangesUrl")=Request.Form("ChangesUrl")
							
							Call KSCls.AddDiyFieldValue(RS,FieldXml)
							on error resume next
							RS.Update
							if err then ks.die err.description &"==" & proid & "+==" & ii
						   RS.MoveLast
						   Call SaveProAttr(rs("id"),1)
						   Call SaveProImages(rs("id"),1)
						   Call addBundleSale(rs("id"))
                          if ii=0 then
							Conn.Execute("Delete From KS_SpecialR Where InfoID=" &ProductID & " and channelid=" & ChannelID)
							For I=0 To Ubound(SpecialID)
							Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & ProductID & "," & ChannelID & ")")
							Next
						 End If
							Call KSCls.UpdateRelative(ChannelID,ProductID,SelectInfoList,1)
							Call LFCls.UpdateItemInfo(ChannelID,ProductID,Title,Tid,ProIntro,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,IsSpecial,Popular,Slide,IsTop,Comment,Verific)
							'关联上传文件
							 if ii=0 then
							Call KS.FileAssociation(ChannelID,ProductID,PhotoUrl & BigPhoto & ProIntro ,1)
							If Not KS.IsNul(FileIds) Then 
								Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & ProductID &",classID=" & KS.C_C(Tid,9) & " Where ID In (" & FileIds & ")")
							End If
							End If
							Call RefreshHtml(2)
					 RS.Close:Set RS = Nothing
				  Next	
		          
					
					If KeyWord <>"" Then
						 .Write ("<script> parent.frames['MainFrame'].focus();setTimeout(function(){alert('商品修改成功!');location.href='../System/KS.ItemInfo.asp?Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='../Post.Asp?ButtonSymbol=ShopSearch&OpStr=" & server.URLEncode("商品管理 >> <font color=red>搜索结果</font>") & "';},2500);</script>")
					End If
				End If
			End If
		 End With		
		End Sub
		
		Sub SaveProAttr(ProID,DelTF)
		   ' if deltf=1 then Conn.Execute("Delete From KS_ShopSpecificationPrice Where ProID=" & ProID)
		    Dim totalrow:totalrow=KS.ChkClng(KS.G("totalrow"))
			Dim RS
			If totalrow>0 Then
			 
			 for i=1 to totalrow
			   if request("aitemno"&i)<>"" then
			     Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "select top 1 * From KS_ShopSpecificationPrice Where ID=" & KS.ChkClng(request("id"&i)),conn,1,3
				 If RS.Eof And RS.Bof Then
				   RS.AddNew
				 End If
				   RS("ItemNo") = trim(request("aitemno"&i))
				   RS("proid")  = Proid
				   RS("Attr1")  = request("attr" & i & "0") & "|" & request("attrimg" &I &"0")
				   RS("Attr2")  = request("attr" & i & "1") & "|" & request("attrimg" &I &"1")
				   RS("Attr3")  = request("attr" & i & "2") & "|" & request("attrimg" &I &"2")
				   If IsNumeric(request("aprice" & i)) Then
				     RS("Price")  = request("aprice" & i)
				   ElseIf IsNumeric(request("price_member")) Then
				     RS("Price")  = request("price_member")
				   Else
				     RS("Price")  = 0
				   End If
				   RS("Amount")  = KS.ChkCLng(Request("aamount"&i))
				   If IsNumeric(Request("aweight"&i)) Then
				    RS("Weight") = Request("aweight"&i)
				   Else
				    RS("Weight") = 0
				   End If
				   RS.Update
				  RS.Close
				 Set RS=Nothing
			   end if
			 next
			End If
			
		End Sub
		
		Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="添加":EditStr="修改商品":AddStr="继续添加商品"
				Else
				  TempStr="修改":EditStr="继续修改商品":AddStr="添加商品"
				End If
			    With Response
 .Write "<!DOCTYPE html><html>"
			         .Write"<head>"
				     .Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<meta http-equiv=Content-Type content=""text/html; charset=utf-8"">"
					 .Write "<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>"
					 .Write"</head>"
					 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
					 .Write " <div class='pageCont2 mt20'><div class='tabTitle'>系统操作提示信息</div><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""0"" cellspacing=""0"">"
                      .Write "    <tr class='tdbg' colspan=2>"
					  .Write "          <td align='center'><table width='100%' border='0'><tr><td style='width:200px;text-align:center'><img src='../images/succeed.gif'>"
					  .Write "</td><td><div style='padding-left:30px;font-weight:bold'>恭喜，" & TempStr &"" & KS.C_S(ChannelID,3) & "成功！</div>"					  
					  
					   If Makehtml = 1 Then
					      .Write "<div style=""margin-top:15px;border: #E7E7E7;height:220; overflow: auto; width:100%"">" 
					    If KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then
						  	 .Write "<div><iframe src=""../Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						  Else
						  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "没有启用生成HTML的功能，所以ID号为 <font color=red>" & RS("ID") & "</font>  的" & KS.C_S(ChannelID,3) & "没有生成!</li></div> "
						  End If
						  
						   If KS.WSetting(0)="1" Then  '手机版
						   If KS.ChkClng(KS.M_C(ChannelID,28))=1  Or KS.ChkClng(KS.M_C(ChannelID,28))=2 Then
						  	 .Write "<div><iframe src=""../Include/RefreshHtmlSave.Asp?from=3g&ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						   End If
						  End If
						  
						  
							If KS.C_S(ChannelID,7)<>1 Then
							  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "的栏目页没有启用生成HTML的功能，所以ID号为 <font color=red>" & TID & "</font>  的栏目没有生成!</li></div> "
							Else
							 If KS.C_S(ChannelID,9)<>1 Then
								  Dim FolderIDArr:FolderIDArr=Split(left(KS.C_C(Tid,8),Len(KS.C_C(Tid,8))-1),",")
								  For I=0 To Ubound(FolderIDArr)
								  .Write "<div align=center><iframe src=""../Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe></div>"
								   Next
							 End If
						   End If
					   If Split(KS.Setting(5),".")(1)="asp" or KS.C_S(ChannelID,9)<>3 Then
					   Else
					     .Write "<div align=center><iframe src=""../Include/RefreshIndex.asp?RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div></div>"
					   End If
					    .Write   "</td></tr></table></td></tr>"
					  .Write "	  <tr class='tdbg'>"
					  .Write "		<td></td><td height=""25"" style='text-align:right'  nowrap='nowrap'>【<a href=""#"" onclick=""location.href='KS.Shop.asp?Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" & RS("ID") & "';""><strong>" & EditStr &"</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.Shop.asp?Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=" & server.URLEncode("添加商品")& "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='../System/KS.ItemInfo.asp?ChannelID=" & ChannelID &"&ID=" & Tid & "&Page=" & Page &"&keyword=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>商品管理</strong></a>】&nbsp;【<a href=""" & KS.GetDomain & "Item/Show.asp?m=" & ChannelID & "&d=" & RS("ID") & """ target=""_blank""><strong>预览商品内容</strong></a>】</td>"
					  .Write "	  </tr>"
					  .Write "	</table></div></body></html>"				
			End With
		End Sub
		
		Function GetBrandByClassID(ClassID,BrandID)
		  Dim SQL,K
		  Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		  RS.Open "Select B.ID,B.BrandName From KS_ClassBrand B inner join KS_ClassBrandR R On B.id=R.BrandID where R.classid='" & classid & "' order by B.orderid",conn,1,1
		  If Not RS.Eof  Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
		  If Not IsArray(SQL) Then
		   GetBrandByClassID="Null" 
		  Else
		     GetBrandByClassID = "<select name='brandid'>"
			 GetBrandByClassID = GetBrandByClassID & "<option value='0'>-请选择品牌-</option>"
		     For K=0 To Ubound(SQL,2)
			  If BrandID=SQL(0,K) Then
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & sql(0,k) & "' selected>" & sql(1,k) & "</option>"
			  Else
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & sql(0,k) & "'>" & sql(1,k) & "</option>"
			  End If
			 Next
			 GetBrandByClassID = GetBrandByClassID &  "</select>"
			 Erase Sql
		  End If
       End Function

	   Sub SetKBXSPrice()
	   	  with response
		      .Write "<!DOCTYPE html><html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>" & vbCrLf
			  .Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"

	    dim proid:proid=ks.chkclng(request("proid"))
		if proid=0 then ks.die "error!"
		dim kbprice:kbprice=request("kbprice")
	    dim rs:set rs=server.createobject("adodb.recordset")
		rs.open "select top 1 * from ks_product where id=" & proid,conn,1,1
		if not rs.eof then
		%>
		 <script type="text/javascript">
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
			function addIn(){
		   var title=$("#title").val();
		   var proid=$("#proid").val();
		   var kbprice=$("#kbprice").val();
		   top.frames["MainFrame"].updateKBXS(proid+"@@@"+kbprice+"@@@"+title);
		   top.box.close();
		  }
		 </script>
		<%if kbprice="" then kbprice=rs("price_member")
	     .write "<table border='0' width='100%' class='ctable' cellspacing='1' cellpadding='1'>"
		 .write "<tr class='tdbg'><td height='30' class='clefttitle' align='right' width='130'><strong>商品名称:</strong></td><td>" & rs("title") & "<input type='hidden' name='title' id='title' value='" & rs("title") & "'><input type='hidden' name='proid' id='proid' value='" & proid & "'></td></tr>"
		 .write "<tr class='tdbg'><td height='30' class='clefttitle' align='right' width='130'><strong>当前零售价:</strong></td><td>￥" & formatnumber(rs("price"),2,-1) & "元</td></tr>"
		 .write "<tr class='tdbg'><td height='30' class='clefttitle' align='right' width='130'><strong>会 员 价:</strong></td><td><font color=red>￥" & formatnumber(rs("price_member"),2,-1) & "元</font></td></tr>"
		 .write "<tr class='tdbg'><td height='30' class='clefttitle' align='right' width='130'><strong>捆绑销售价:</strong></td><td><input type='text' name='kbprice' id='kbprice' value='" & kbprice & "' size='5'  class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" style='text-align:center'> 元</td></tr>"
		 .write "<tr class='tdbg'><td height='30' style='text-align:center' colspan='2'>"
		 if request("kbprice")="" Then
		 .Write "<input onclick='addIn()' type='button' value='确定加入' class='button'>"
		 else
		 .Write "<input onclick='addIn()' type='button' value='确定修改' class='button'>"
		 end if
		 .Write "</td></tr>"
         .write "</table>"		
		end if
		rs.close
		set rs=nothing
		end with
	   End Sub
	   
	   Sub addBundleSale(ID)
	   		Dim KBProList:KBProList=KS.S("KBProList")
			Dim KBArr,I
			 Conn.Execute("Delete From KS_ShopBundleSale Where ProID=" & ID)
			 
			If Not KS.IsNul(KBProList) Then
			  KBProList=Replace(KBProList," ","")
			  KBArr=Split(KBProList,",")
			   For i=0 to ubound(KBArr)
			        Dim KBSArr:KBSArr=Split(KBArr(i),"|")
					Dim PRS:Set PRS=Server.CreateOBject("ADODB.RECORDSET")
					PRS.Open "Select top 1 * From KS_ShopBundleSale Where KBProID=" & KBSArr(0) & " And ProID=" & ID,conn,1,3
				  If PRS.Eof Then
				   PRS.AddNew
				  End If
				   PRS("ProID")=ID
				   PRS("KBProID")=KBSArr(0)
				   PRS("KBPrice")=KBSArr(1)
				  PRS.Update
				  PRS.Close:Set PRS=Nothing
				Next
		  End If
		   
	   End Sub
	   
	   Sub SaveProImages(ProID,flag)
	        'If flag<>0 Then  DelProImages ProID '删除原图
	          Dim sTemp,Url1,thumburl,ThumbFileName,SaveFilePath,PicUrls
			  PicUrls=Request.Form("PicUrls")
			  
			
				  SaveFilePath = KS.GetUpFilesDir & "/"
				  KS.CreateListFolder (SaveFilePath)
				  Dim sPicUrlArr:sPicUrlArr=Split(PicUrls,"$$$")
				   For I=0 To Ubound(sPicUrlArr)
				     If KS.ChkClng(KS.G("BeyondSavePic1"))=1 and Left(Lcase(Split(sPicUrlArr(i),"|")(1)),4)="http" and instr(Lcase(Split(sPicUrlArr(i),"|")(1)),lcase(ks.setting(2)))=0  and instr(Lcase(Split(sPicUrlArr(i),"|")(1)),"kesion.com")=0 Then
					    Url1=KS.ReplaceBeyondUrl(Split(sPicUrlArr(i),"|")(1), SaveFilePath)
					    thumburl=replace(url1,ks.setting(2),"")
					    ThumbFileName=split(thumburl,".")(0)&"_S."&split(thumburl,".")(1)
						if instr(Lcase(thumburl),"http://")=0 Then
							Dim T:Set T=New Thumb
							Dim CreateTF:CreateTF=T.CreateThumbs(thumburl,ThumbFileName)
							if CreateTF=false Then ThumbFileName=url1
							Set T=Nothing
						end if
						Call AddProImages(ProID, Url1,ThumbFileName,Split(sPicUrlArr(i),"|")(0),Split(sPicUrlArr(i),"|")(3),I+1)
					 Else
					    Call AddProImages(ProID, Split(sPicUrlArr(i)&"|||","|")(1),Split(sPicUrlArr(i)&"|||","|")(2),Split(sPicUrlArr(i)&"|||","|")(0),Split(sPicUrlArr(i)&"|||","|")(3),I+1)
					 End If
				   Next
	   End Sub
	   
	   Sub DelProImages(ProID)
		 Conn.Execute("Delete From KS_ProImages where proid=" & ProID)
	   End Sub
	  
	  Sub AddProImages(ProID,BigPicUrl,SmallPicUrl,GroupName,PicId,orderid)
	    Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open "select top 1 * From KS_ProImages Where ID=" & KS.ChkClng(PicId),conn,1,3
		If RS.Eof AND RS.Bof Then
		   RS.AddNew
		End If
		   RS("ProID")=ProID
		   RS("SmallPicUrl")=SmallPicUrl
		   RS("BigPicUrl")=BigPicUrl
		   RS("GroupName")=GroupName
		   RS("orderid")=orderid
		RS.Update
		RS.Close
		Set RS=Nothing
	      '关联上传文件
		 Call KS.FileAssociation(5,ProID,SmallPicUrl& " " &BigPicUrl  ,0)
	  End Sub
	  
End Class
%> 
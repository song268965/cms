<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_Article
KSCls.Kesion()
Set KSCls = Nothing
Class Admin_Article
        Private KS,ComeUrl,KSCls
		'=====================================声明本页面全局变量==============================================================
        Private ID, I, totalPut, Page, RS,ComeFrom
		Private KeyWord, SearchType, StartDate, EndDate,SearchParam, MaxPerPage,T, TitleStr, VerificStr
		Private TypeStr, AttributeStr, FolderID, TemplateID,WapTemplateID,FolderName, Action
		Private NewsID, TitleType, Title,Fulltitle,ShowComment, TitleFontColor, TitleFontType, PicNews, ArticleContent, PhotoUrl, Changes, Recommend,IsTop,PageTitle,IsSign,SignUser,SignDateLimit,SignDateEnd,Province,City,County,RelatedID,Otid,OID
		Private Strip, Popular, Verific, Comment, Slide,ChangesUrl, Rolls, KeyWords, Author, Origin, AddDate, Rank,  Hits, HitsByDay, HitsByWeek, HitsByMonth, SpecialID,CurrPath,UpPowerFlag,Intro,IsVideo
		Private Inputer,FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,SaveFilePath,MapMarker
		Private ReadPoint,ChargeType,PitchTime,ReadTimes,InfoPurview,arrGroupID,DividePercent
		Private SEOTitle,SEOKeyWord,SEODescript
		Private ChannelID,PostId,FieldXML,FieldNode,FNode,FieldDictionary
		'======================================================================================================================
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
			Action     = KS.G("Action") 'Add添加新文章 Edit编辑文章 Verify 审核前台投搞
			If Action="SelectUser" Then
			   Call SelectUser()
			   Exit Sub
			ElseIf Action="SelectClass" Then
			   Call KSCls.SelectMutiClass()
			   Exit Sub
			End If
			Call KSCls.LoadModelField(ChannelID,FieldXML,FieldNode)
			Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
			KeyWord    = KS.G("KeyWord")
			SearchType = KS.G("SearchType")
			StartDate  = KS.G("StartDate")
			EndDate    = KS.G("EndDate")
			ComeFrom   = KS.G("ComeFrom")
			SearchParam = "ChannelID=" & ChannelID
			If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
			If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
			If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
			If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
			If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
			If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
	
			ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
			Page = KS.ChkClng(KS.G("page"))
			If Action="CheckTitle" Then
				Call KSCls.CheckTitle()    
			Else
				Page = KS.G("page")
				Action = KS.G("Action") 
				IF KS.G("Method")="Save" Then Call DoSave()	Else Call ArticleManage()
			End If
		
	 End Sub
				
	 Sub ArticleManage()
			With Response
            .Write"<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
			.Write "<title>文章添加/修改</title>"
			.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<script src=""../images/pannel/tabpane.js""></script>" & vbCrlf
			.Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
		    .Write "<script src=""../../KS_Inc/Jquery.js""></script>" & vbCrLf
		    .Write "<script src=""../../KS_Inc/common.js""></script>" & vbCrLf
			.Write "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>" & vbCrlf
			.Write EchoUeditorHead
			CurrPath = KS.GetUpFilesDir
			 
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Action = "Add" Then			 

			  FolderID = Trim(KS.G("FolderID"))
			  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '检查是否有添加文章的权限
			   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
			   Call KS.ReturnErr(2, "../System/KS.ItemInfo.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
			   Exit Sub
			  End If
			  Hits = 0:HitsByDay = 0: HitsByWeek = 0:HitsByMonth = 0:Comment = 1 :IsTop=0:Verific=1
			  ReadPoint=0:PitchTime=24:ReadTimes=10: IsSign=0 : SignDateLimit=0 : SignDateEnd=Now : IsVideo=0
			  KeyWords = Session("keywords")
			  Author = Session("Author")
			  Origin = Session("Origin")
			ElseIf Action = "Edit"  Or Action="Verify" Then
			   Set RS = Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where ID=" & KS.G("ID") , conn, 1, 1
			   If RS.EOF And RS.BOF Then	Call KS.Alert("参数传递出错!", ComeUrl):Exit Sub
				FolderID = Trim(RS("Tid"))
				
				If Action = "Edit" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '检查是否有编辑文章的权限
					RS.Close:Set RS = Nothing
					 If KeyWord = "" Then
					  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
					  Call KS.ReturnErr(2, "../System/KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&ID=" & FolderID)
					 Else
					  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>搜索" & KS.C_S(ChannelID,3) & "结果</font>")&"&ButtonSymbol=ArticleSearch';</script>")
					  Call KS.ReturnErr(1, "../System/KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate)
					 End If
					 Exit Sub
			   End If
			   If Action="Verify" And KS.C("Role")="1" Then     '检查是否有审核前台会员投稿文章的权限
					  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
					  Call KS.ReturnErr(2, "../System/KS.ItemInfo.asp?ShowType=1&ChannelID=" & ChannelID &"&Page=" & Page & "&ID=" & FolderID)
			   End If
			   
				TitleType      = Trim(RS("TitleType"))
				Title          = Trim(RS("title"))
				Fulltitle      = Trim(RS("Fulltitle"))
				TitleFontColor = Trim(RS("TitleFontColor"))
				TitleFontType  = Trim(RS("TitleFontType"))
				PhotoUrl         = Trim(RS("PhotoUrl"))
				PicNews        = CInt(RS("PicNews"))
				Rolls          = CInt(RS("Rolls"))
				Changes        = CInt(RS("Changes"))
				Recommend      = CInt(RS("Recommend"))
				Strip          = CInt(RS("Strip"))
				Popular        = CInt(RS("Popular"))
				Verific        = CInt(RS("Verific"))
				IsTop          = Cint(RS("IsTop"))
				Comment        = CInt(RS("Comment"))
				Slide          = CInt(RS("Slide"))
				IsVideo        = RS("IsVideo")
				AddDate        = CDate(RS("AddDate"))
				Rank           = Trim(RS("Rank"))
				FileName       = RS("Fname")
                If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonform").text="1" Then MapMarker      = RS("MapMarker")
				TemplateID     = RS("TemplateID")
				WapTemplateID  = RS("WapTemplateID")
				Hits           = RS("Hits")
				HitsByDay      = RS("HitsByDay")
				HitsByWeek     = RS("HitsByWeek")
				HitsByMonth    = Trim(RS("HitsByMonth"))
				KeyWords       = Trim(RS("KeyWords"))
				Author         = Trim(RS("Author"))
				Origin         = Trim(RS("Origin"))
				Inputer        = RS("Inputer")
				Intro          = RS("Intro")
				IsSign         = RS("IsSign")
				SignUser       = RS("SignUser")
				SignDateLimit  = RS("SignDateLimit")
				SignDateEnd    = RS("SignDateEnd")
				Province       = RS("Province")
				City           = RS("City")
				County		   = RS("County")
                PostId         = RS("PostId")
				ReadPoint      = RS("ReadPoint")
				ChargeType     = RS("ChargeType")
				PitchTime      = RS("PitchTime")
				ReadTimes      = RS("ReadTimes")
				InfoPurview    = RS("InfoPurview")
				arrGroupID     = RS("arrGroupID")
				DividePercent  = RS("DividePercent")
				SEOTitle       = RS("SEOTitle")
				SEOKeyWord     = RS("SEOKeyWord")
				SEODescript    = RS("SEODescript")
				RelatedID      = RS("RelatedID")
			   If CInt(Changes) = 1 Then
				ChangesUrl     = Trim(RS("ArticleContent"))
			   Else
				ArticleContent = Trim(RS("ArticleContent"))
			   End If
			   If KS.IsNul(ArticleContent) Then ArticleContent="&nbsp;"
			    PageTitle      = RS("PageTitle")
				FolderID       = RS("Tid")
				oTid           = RS("oTid")
				OID            = RS("OID")
				'自定义字段
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
			If IsNULL(PageTitle) Then PageTitle=""
			'取得上传权限
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
			 
			%>
			<script language='JavaScript'>
			$(document).ready(function(){
				$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
			 	$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);

			 <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='turnto']/showonform").text="1" Then%>
				   if ($("#Changes").prop('checked')){ChangesNews();}
			 <%End If%>
			 <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
			  $('#KeyLinkByTitle').click(function(){GetKeyTags();});
			 <%End If%>
			  //只有一个栏目时，为选其选中
			  if ($("#tid option").length<=2){
			     $("#tid option").each(function(i){
				    if (i==$("#tid option").length-1) $(this).attr("selected",true);
				});
			  }
			  
			});
			function GetKeyTags()
			{
			  var text=escape($('#title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){$('#KeyWords').val(unescape(data)).attr("disabled",false);});
			  }else{ top.$.dialog.alert('对不起,请先输入内容!'); }
			}
			
			function ChangesNews()
			{ 
			 if ($("#Changes:checked").val()=="1")
			  {
			  <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then%>
			  $("#ContentArea").hide();
			  <%end if%>
			  $("#ChangesUrl").attr("disabled",false);
			  }
			  else
			   {
			   <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then%>
			   $("#ContentArea").show();
			   <%end if%>
			  $("#ChangesUrl").attr("disabled",true);
			   }
			}
			function UnSelectAll(){$("#SpecialID option").each(function(){$(this).attr("selected",false);});}
			function GetFileNameArea(f){$('#filearea').toggle(f);}
			function GetTemplateArea(f){$('#templatearea').toggle(f);}
			function insertHTMLToEditor(codeStr) { <%=InsertEditor("Content","codeStr")%>} 
			function insertPage(){insertHTMLToEditor("[NextPage]");}
			
			function SubmitFun()
			{ 
			   if ($("#title").val()==""){
					top.$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>标题！",function(){ $("#title").focus();});
					return false;
			   }
			   if ($("#tid option:selected").val()=='0' || $("#tid").val()=='0')
			   {
			       top.$.dialog.alert('请选择所属<%=KS.GetClassName(ChannelID)%>!');
				   return false;
			   }
			<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
			  if ($("#KeyWords").val().length>255){
			    top.$.dialog.alert('关键字不能超过255个字符!',function(){ $("#KeyWords").focus();});
				return false;}
			<%
			 End If
			 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='turnto']/showonform").text="1" and FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then
			 %>
                   if (($("#Changes:checked").val()!="1")&&(<%=GetEditorContent("Content")%>==false))
					 { top.$.dialog.alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！",function(){ <%=GetEditorFocus("Content")%>});
					  return false;
					 }
				<%end if%>
				<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='turnto']/showonform").text="1" Then%>
				if (($("#Changes:checked").val()=="1")&&($("input[name=ChangesUrl]").val()==""))
				  { $("#ChangesUrl").focus();
					top.$.dialog.alert("请输入外部链接的Url！");
					return false;
				  }
				<%end if%>
				 <%
			  Call LFCls.ShowDiyFieldCheck(FieldXML,1)
			     %>
				<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then%>
				if ($("input[name=BeyondSavePic]").attr('checked')==true)
				 {
				  $('#LayerPrompt').show();
				  window.setInterval('ShowPromptMessage()',150)
				 }
				 <%end if%>
				  $('#myform').submit();
				  $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
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
			
			
			var SaveBeyondInfo=''
					   +'<div id="LayerPrompt" style="position:absolute; z-index:1; left: 200px; top: 200px; background-color: #f1efd9; layer-background-color: #f1efd9; border: 1px none #000000; width: 360px; height: 63px; display: none;"> '
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
			function SelectUser(){
			   top.openWin('选择签收用户','article/KS.Article.asp?action=SelectUser&DefaultValue='+document.myform.SignUser.value,false,860,430);
				//var arr=showModalDialog('KS.Article.asp?action=SelectUser&DefaultValue='+document.myform.SignUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');
				//if (arr != null){
				//	document.myform.SignUser.value=arr;
				//}
			}
		  function addMap(){ top.openWin('电子地图标注','../../plus/baidumap.asp?obj=parent.frames["MainFrame"].document.getElementById("MapMark")&MapMark='+escape($("#MapMark").val()),false,860,430); }
		  function getBoardCategory(boardid){
		   if (boardid!=0){
		  $.get("../../plus/ajaxs.asp",{action:"getclubboardcategory",boardid:boardid,anticache:Math.floor(Math.random()*1000)},function(d){
		     $("#showcategory").html(unescape(d));
	       });
		    }
		  }
			</script>
			<%
			
			Call KSCls.EchoFormStyle(ChannelId)   '控制添加文档布局
			
			
			.Write "</head>"
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>"
			.Write "<div>"
			.Write "<ul id='menu_top' class='menu_top_fixed'>"
			.Write "<li class='parent' onclick=""return(SubmitFun())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon save'></i>确定保存</span></li>"
			.Write "<li class='parent' onclick=""history.back();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消返回</span></li>"
		    .Write "</ul>"
			.Write "<div class=""menu_top_fixed_height""></div>"
			
			.Write "<div class=tab-page id=ArticlePane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""ArticlePane"" ), 1 )"
			.Write " </SCRIPT>"
				 
			
			.Write "  <form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' id='myform' name='myform'>"
			.Write "      <input type='hidden' value='" & KS.G("ID") & "' name='NewsID'/>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'/>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'/>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'/>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'/>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'/>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'/>"
			.Write "      <input type='hidden' name='ArticleID' value='" & KS.G("ID") & "'/>"
			.Write "      <input type='hidden' name='Inputer' value='" &Inputer & "'/>"
			
			Dim AckPlusTF:AckPlusTF=KS.GetAppStatus("tags")
			Call KS.LoadFieldGroupXML()
			
Dim TypeNode
IF IsObject(Application(KS.SiteSN & "_FieldGroupXml")) Then
  For Each TypeNode In Application(KS.SiteSN & "_FieldGroupXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &"]")
     .Write " <div class=tab-page id=""p" &TypeNode.SelectSingleNode("@id").text & """>"
	 .Write "  <H2 class=tab>" & TypeNode.SelectSingleNode("@groupname").text & "</H2>"
	 .Write "	<SCRIPT type=text/javascript>"
	 .Write "				 tabPane1.addTabPage( document.getElementById( ""p" &TypeNode.SelectSingleNode("@id").text & """ ) );"
	 .Write "	</SCRIPT>"
	 
			
	
	  .Write " <dl class='dtable'>"
	For Each FNode In FieldXML.DocumentElement.SelectNodes("fielditem[showonform=1&&fieldtype!=13&&@groupid=" & TypeNode.SelectSingleNode("@id").text & "]")
	    If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
			.Write   KSCls.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary,0) '自定义字段
		Else
		 Dim XTitle:XTitle=FNode.SelectSingleNode("title").text
		  
	     Select Case lcase(FNode.SelectSingleNode("@fieldname").text)
	       case "title"
				
				  IF Action="Verify" and  Inputer<>KS.C("AdminName") and KS.C("Role")="3" Then
					  .Write " <dd>"
					  .Write "   <div>额外奖励:</div>" &vbcrlf
					  .Write "   点券<input style='text-align:center' name=""UserAddPoint"" class='textbox' type=""text"" id=""UserAddPoint"" value=""0"" size=""6"">点  积分<input  name=""UserAddScore"" type=""text"" class='textbox' id=""UserAddScore"" value=""0"" style='text-align:center' size=""6"">分 <span>为“0”时不增加,对优秀信息给予额外奖励</span></dd>"&vbcrlf
				  end if
					.Write "      <dd><div>" & XTitle & ":</div>"
					If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='titleattribute']/showonform").text="1" Then
					.Write "<select name='TitleType' id='TitleType' class='textbox'>"
					.Write "                    <option></option>"
					
					 Dim TitleTypeXml:Set TitleTypeXml=LFCls.GetXMLFromFile("TitleType")
					 If IsObject(TitleTypeXml) Then
						 Dim objNode,i,j,objAtr
						 Set objNode=TitleTypeXml.documentElement 
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i)
								If Trim(TitleType) = Trim(objAtr.Attributes.item(0).Text) Then 
								.Write "<option selected style='color:" &objAtr.Attributes.item(1).Text & "'>" & objAtr.Attributes.item(0).Text & "</option>"
								Else
								.Write "<option style='color:" &objAtr.Attributes.item(1).Text & "'>" & objAtr.Attributes.item(0).Text & "</option>"
								End If
						 Next
					End If		
					.Write "   </select>"
				   End If
					.Write "  <input name='title' id='title' type='text' class=""rule"" value=""" & Title & """ maxlength='160' size='50'/> <font color='#FF0000'>*</font>"
					
					If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='titleattribute']/showonform").text="1" Then
					.Write "   <select name='TitleFontType' id='TitleFontType'>"
					.Write "     <option value=''>字形</option>"
					If TitleFontType = "1" Then  .Write "<option value='1' selected>粗体</option>" Else  .Write "<option value='1'>粗体</option>"
					If TitleFontType = "2" Then  .Write "<option value='2' selected>斜体</option>" Else  .Write "<option value='2'>斜体</option>"
					If TitleFontType = "3" Then  .Write "<option value='3' selected>粗+斜</option>" Else .Write "<option value='3'>粗+斜</option>"
					If TitleFontType = "0" Then  .Write "<option value='0' selected>规则</option>"	Else .Write "<option value='0'>规则</option>"
					.Write " </select><input type='hidden' id='TitleFontColor' name='TitleFontColor' value='" & TitleFontColor &"'>"
					Dim ColorImg:If TitleFontColor="" Then ColorImg="RectNoColor.gif" Else ColorImg="rect.gif"
					.Write " <img border=0 id=""MarkFontColorShow"" src=""../images/" & ColorImg & """ style=""cursor:pointer;background-Color:" & TitleFontColor & ";"" onClick=""Getcolor('MarkFontColorShow','../../editor/ksplus/selectcolor.asp?MarkFontColorShow|TitleFontColor');this.src='../images/rect.gif';"" title=""选取颜色"">&nbsp;"
					End If
			        .Write "<input class='button' type='button' value='重名检测' onclick=""if($('#title').val()==''){ top.$.dialog.alert('请输入" & KS.C_S(ChannelID,3) & "标题!');}else top.openWin('" & KS.C_S(ChannelID,3) & "重名检测','article/KS.Article.asp?ChannelID=" & ChannelID & "&Action=CheckTitle&title='+escape($('#title').val()),false,360,370);"">"
					
					If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/showonform").text="1" Then
					.Write "<label><input type='checkbox' name='MakeHtml' value='1' checked>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/title").text & "</label>"
					End IF
					
					If Action="Edit" Then
					.Write "<label><input type='checkbox' name='AddNew' value='1'/>添加为新" & KS.C_S(ChannelID,3) & "</label>"
						if RelatedID=-11 or KS.ChkClng(RelatedID)<>0 then
							.Write "<span style=""padding:5px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6""><label><input type='checkbox' name='EditNewtb' value='1' checked/> 此"  & KS.C_S(ChannelID,3) & "发布到多个栏目，选中将同步更新 <input type='hidden' name='RelatedID' value='"& RelatedID &"'/></label></span>"
						end if
					End If
					
					If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pushtobbs']/showonform").text="1" Then
					 If KS.Setting(56)="1" Then 
					  KS.LoadClubBoard
					  if isobject(Application(KS.SiteSN&"_ClubBoard")) then
						.Write "   <input type='hidden' name='OldClassID' value='" & FolderID & "'>"
					   .Write "<select name='bid' id='bid' onchange='getBoardCategory(this.value)'><option value='0'>===推送到论坛===</option>"
						Dim RSB,Xml,Node,Node1,BoardID,CategoryID
					   If KS.ChkClng(PostId)<>0 Then
						  Set RSB=Conn.Execute("select top 1 BoardID,CategoryID From KS_GuestBook Where id=" & PostID)
						  If Not RSB.Eof Then
							BoardID=RSB(0):CategoryId=RSB(1)
						  End If
						  RSB.Close:Set RSB=Nothing
					   End If
						Set Xml=Application(KS.SiteSN&"_ClubBoard")
						for each node in xml.documentelement.selectnodes("row[@parentid=0]")
							KS.Echo "<OPTGROUP label=&nbsp;" & node.selectsinglenode("@boardname").text & " </OPTGROUP>"
							For Each Node1 In xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
							 If trim(BoardID)=trim(Node1.SelectSingleNode("@id").text) Then
								KS.Echo "<option value='" & Node1.SelectSingleNode("@id").text & "' selected>" & node1.selectsinglenode("@boardname").text &"</option>"
							 Else
								KS.Echo "<option value='" & Node1.SelectSingleNode("@id").text & "'>" & node1.selectsinglenode("@boardname").text &"</option>"
							End If
							Next
						next
					   .Write "</select>"
					   .Write " <font style=""font-weight:normal;font-size:12px;"" id=""showcategory"">"
						If KS.ChkClng(CategoryID)<>0 Then
						 If KS.ChkClng(BoardID)<>0 Then
							 Dim CategoryStr,CategoryNode
							 KS.LoadClubBoardCategory
							 For Each CategoryNode In Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" &BoardID &"]")
							   If Trim(CategoryID)=trim(CategoryNode.SelectSingleNode("@categoryid").text) Then
								CategoryStr=CategoryStr & "<option selected value='" &CategoryNode.SelectSingleNode("@categoryid").text  & "'>" & CategoryNode.SelectSingleNode("@categoryname").text &"</option>"
							   Else
								CategoryStr=CategoryStr & "<option value='" &CategoryNode.SelectSingleNode("@categoryid").text  & "'>" & CategoryNode.SelectSingleNode("@categoryname").text &"</option>"
							   End If
							Next
							If Not KS.IsNul(CategoryStr) Then
								CategoryStr="<strong>主题分类:</strong><select name=""CategoryId"" id=""CategoryId""><option value='0'>==选择分类==</option>"  & CategoryStr &"</select>"
							End If
							KS.echo (CategoryStr)
						End If
		
						End If
					   .Write "</font>" &vbcrlf
					 end if
					End If
					
					
					End If
					
					.Write "</dd>"&vbcrlf
		  case "fulltitle"
		            .Write " <dd><div>" & XTitle & ":</div>" &vbcrlf
			        .Write " <input name='Fulltitle' type='text' maxlength='200' id='Fulltitle' size='80' value='" & Fulltitle & "' class='textbox'></dd>"&vbcrlf
		  case "tid"
					.Write " <dd style=""""><div>" & Replace(XTitle,"栏目",KS.GetClassName(ChannelID)) & ":</div>"
					.Write " <input type='hidden' name='OldClassID' value='" & FolderID & "'>"
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
					top.openWin(false,'article/KS.Article.asp?channelID=<%=ChannelID%>&FolderID='+$("#tidtb").val()+'&action=SelectClass',false,400,420);
					}
					</script>
					<%
					
				 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/showonform").text="1" Then
					.Write "&nbsp;" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/title").text & "：<label><input name='Recommend' type='checkbox' id='Recommend' value='1'"
					If Recommend = 1 Then .Write (" Checked")
					.Write ">推荐</label><label><input name='IsVideo' type='checkbox' id='IsVideo' value='1'"
					If IsVideo = "1" Then .Write (" Checked")
					.Write ">视频</label><label><input name='Rolls' type='checkbox' id='Rolls' value='1'"
					If Rolls = 1 Then .Write (" Checked")
					.Write ">滚动</label><label><input name='Strip' type='checkbox' id='Strip' value='1'"
					If Strip = 1 Then .Write (" Checked")
					.Write ">头条</label><label><input name='Popular' type='checkbox' id='Popular' value='1'"
					If Popular = 1 Then .Write (" Checked")
					.Write ">热门</label><label><input name='IsTop' type='checkbox' id='IsTop' value='1'"
					If IsTop = 1 Then .Write (" Checked")
					.Write ">固顶</label><label><input name='Comment' type='checkbox' id='Comment' value='1'"
					If Comment = 1 Then .Write (" Checked")			
					.Write ">允许评论</label><label><input name='Slide' type='checkbox' id='Slide' value='1'"
					If Slide = 1 Then
					.Write (" Checked")
					End If
					.Write ">幻灯</label>"
					
					Call KSCls.GetDiyAttribute(FieldXML,FieldDictionary)
					
				End If
				   .Write " </dd>"
		case "otid"
		         Call KSCls.EchoOTidInfo(FNode,OTid,Oid)		
		case "map"
				.Write " <dd>"
				.Write "  <div>" & XTitle &":</div>经纬度：<input size='43' value=""" & MapMarker & """ type='text' name='MapMark' id='MapMark' class='textbox' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='../images/accept.gif' align='absmiddle' border='0'>添加电子地图标志</a></dd>" &vbcrlf
	    case "turnto"
		        .Write "<dd><div>" & XTitle &":</div>" &vbcrlf 
				If ChangesUrl = "" Then
				 .Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='50' class='textbox'>")
				Else
				 .Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl & "' size='50' class='textbox'>")
				End If
				If Changes = 1 Then
				 .Write (" <input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
				Else
				 .Write (" <input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'> 使用转向链接</font>")
				End If
				.Write " </dd>" & vbcrlf
		case "keywords"
		        .Write " <dd><div>" & XTitle & ":</div><input name='KeyWords' type='text' id='KeyWords' class='textbox' value='" & KeyWords & "' size=""50""/>"
				If AckPlusTF Then
				.Write " <= <select name='SelKeyWords' style='width:150px' onChange='InsertKeyWords(document.getElementById(""KeyWords""),this.options[this.selectedIndex].value)'>"
				.Write "<option value="""" selected> </option><option value=""Clean"" style=""color:red"">清空</option>"
				.Write   KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")
				.Write " </select> <input type='checkbox' name='tagstf' value='1' checked>记录"
			    End If
				.Write " 【<a href=""javascript:;"" id=""KeyLinkByTitle"" style=""color:green"">根据" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='title']/title").text & "自动获取Tags</a>】"
				.Write " </dd>"& vbcrlf
		case "author"
		        .Write " <dd><div>" & XTitle & ":</div><input name='author' type='text' id='author' value='" & Author & "' size=50 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick='$(""#author"").val(""未知"")' style='cursor:pointer;'>未知</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#author').val('佚名')"" style='cursor:pointer;'>佚名</font></font>】【<font color='blue'><font color='red' onclick=""$('#author').val('" & KS.C("AdminName") & "')"" style='cursor:pointer;'>" & KS.C("AdminName") & "</font></font>】"
								 If Author <> "" And Author <> "未知" And Author <> KS.C("AdminName") And Author <> "佚名" Then
								  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#author').va('" & Author & "')"" style='cursor:pointer;'>" & Author & "</font></font>】")
								 End If
				
				If AckPlusTF Then				 
				.Write ("<select name='SelAuthor' style='width:100px' onChange=""$('#author').val(this.options[this.selectedIndex].value)"">")
				.Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
				.Write KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=0 and OriginType=1 Order BY AddDate Desc")
				.Write " </select>"
			   End If
				.Write "</dd>" & vbcrlf
		case "origin"
                .Write "<dd><div>" & XTitle & ":</div><input name='Origin' type='text' id='Origin' value='" & Origin & "' size=50 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('不详');"" style='cursor:pointer;'>不详</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('本站原创')"" style='cursor:pointer;'>本站原创</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('互联网')"" style='cursor:pointer;'>互联网</font></font>】"
								  If Origin <> "" And Origin <> "不详" And Origin <> "本站原创" And Origin <> "互联网" Then
								  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('" & Origin & "')"" style='cursor:pointer;'>" & Origin & "</font></font>】 ")
								   End If
				If AckPlusTF Then	
				.Write ("<select name='selOrigin' style='width:100px' onChange=""$('#Origin').val(this.options[this.selectedIndex].value)"">")
				.Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
				.Write KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")
				.Write " </select>" &vbcrlf
				End If
				.Write "</dd>" &vbcrlf		
		case "area"
		        .Write "<dd>"
				.Write "<script type='text/javascript'>"
				.write "try{setCookie(""pid"",'" & province & "');setCookie(""cid"",'" & city & "');}catch(e){}" & vbcrlf
				.write "</script>"
				.Write "   <div>" & XTitle & ":</div><script src=""../../plus/area.asp"" type=""text/javascript""></script>  <font color='#999999'>指定文档的来源地或是指定具体的分站新闻</font></dd>" &vbcrlf
				.Write "<script type='text/javascript'>"
				if Province<>"" then
				  .Write "$('#Province').val('" & province & "');"
				end if
				if City<>"" Then
				  .Write "$('#City').val('" & City & "');"
				end if
				if County<>"" Then
				  .Write "$('#County').val('" & County & "');"
				end if
				.Write "</script>"&vbcrlf
		case "intro"
		        .Write "<dd>"
				.Write " <div>" & XTitle & ":<font>(<input name='AutoIntro' type='checkbox' checked value='1'>自动截取内容的200个字作为导读。)</font></div>"
				.Write "  <textarea class='textbox' name=""Intro"" style='width:90%;height:120px'>" & Intro & "</textarea>"
				.Write " </dd>" &vbcrlf
		case "content"
				.Write "<dd ID='ContentArea'>"
				.Write " <div>" & XTitle & ":<font><input name='BeyondSavePic' type='checkbox' value='1' checked>自动下载内容里的图片</font></div>"
				
				If KS.ChkClng(KS.M_C(ChannelID,30))=1 Then
			    .Write "<table width=""92%"">"
				Else
			    .Write "<table width=""100%"">"
				End If
				.Write "<tr><td>"
                 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/showonform").text="1" and CBool(UpPowerFlag) = True Then
				.Write "<table border='0' width='100%' cellspacing='0' cellpadding='0'>"
				.Write "<tr><td height='30' width=70>"
				.Write "&nbsp;<strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/title").text & ":</strong></td><td><iframe id='upiframe' name='upiframe' src='../../user/BatchUploadForm.asp?UPFrom=Admin&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='620' height='48'></iframe></td></tr>"
				.Write "</table>"
			   end if		
			  
			    .Write EchoEditor("Content",ArticleContent,"NewsTool","100%","350px")
			   		

	
				.Write "<table border='0' width='100%' cellspacing='0' cellpadding='0'>"
				.Write "<tr><td height='30' width='60'><strong>过滤设置: </strong></td>"
				.Write "<td><label><input type='checkbox' name='FilterIframe' value='1'>Iframe</label><label><input type='checkbox' name='FilterObject' value='1'>Object</label><label><input type='checkbox' name='FilterScript' value='1'>Script</label><label><input type='checkbox' name='FilterDiv' value='1'>Div</label><label><input type='checkbox' name='FilterClass' value='1'>Class</label><label><input type='checkbox' name='FilterTable' value='1'>Table</label><label><input type='checkbox' name='FilterSpan' value='1'>Span</label><label><input type='checkbox' name='FilterImg' value='1'>IMG</label><label><input type='checkbox' name='FilterFont' value='1'>Font</label><label><input type='checkbox' name='FilterA' value='1'>A链接</label><label><input type='checkbox' name='FilterHtml' value='1' onclick=""alert('所有HTML格式将被清除！');"">HTML</label><label><input type='checkbox' name='FilterTd' value='1'>TD</label>"
				.Write "</td></tr>"
				.Write "<tr><td height='30' width='60'><strong>分页方式: </strong></td><td align='left' colspan='2'><select onchange=""if (this.value==2){$('#pagearea').show()}else{$('#pagearea').hide()}"" name='PaginationType'><option value='0'>不分页</option><option value='1' selected>手工分页</option><option value=2>自动分页</option></select>&nbsp;&nbsp;<strong>注：</strong><font color=blue>手工分页符标记为<font color=red>“[NextPage]”</font>，注意大小写</font> </td></tr>"
				.Write "<tr id='pagearea' style='display:none'><td colspan=3>&nbsp;自动分页时的每页大约字符数<input type='text' name='MaxCharPerPage' value='" & KS.Setting(9) & "' size=6 class='textbox'> <font color=blue>必须大于100才能生效</font></td></tr>"          
				 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pagetitle']/showonform").text="1" Then
				.Write "<tr><td><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pagetitle']/title").text &": </strong></td><td align='left'><textarea name=""PageTitle"" style='width:480px;height:60px;line-height:26px;padding-left:20px;background:url(../images/Rule1.gif) no-repeat 0 0px;border:1px solid #999999;' ID=""PageTitle"">" & Replace(PageTitle,"§",vbcrlf) & "</textarea><font color=green>一行一个标题</font> </td></tr>"
				End If
				.Write "</table>"
				.Write "</td></tr></table>"
				
				.Write "</dd>" &vbcrlf
		case "photourl"
		        .Write "  <dd id=trpic>"
				.Write "  <div>" & XTitle & ":</div>"
				
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='photourl']/showonform").text="1" Then
				   .Write "<div style=""float:right;margin:0 auto;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:40px;width:45px;overflow:hidden;border:1px solid #ccc""><img src=""" & PhotoUrl & """ onerror=""this.src='../images/nopic.gif';"" id=""pic"" style=""height:40px;width:45px;""></div>"
				end if
				
				.Write "<input name='PhotoUrl' type='text' id='PhotoUrl' size='50' value='" & PhotoUrl & "' class='textbox'>"
				.Write "  <input class=""button""  type='button' name='Submit' value='选择...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,$('#PhotoUrl')[0],'pic');"">  <input class=""button"" type='button' name='Submit' value='远程...' onClick=""top.openWin('抓取远程图片','include/SaveBeyondfile.asp?ieditor='+$('#ieditor').prop('checked')+'&pic=pic&fieldid=PhotoUrl&CurrPath=" & CurrPath & "',false,500,100);"">"
				.Write " <input class=""button""  type='button' name='Submit' value='裁剪...' onClick=""if($('#PhotoUrl').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindow(1,'" & KS.Setting(3) & "',$('#PhotoUrl').val())}"">  "
				.Write " <input type='checkbox' name='ieditor' id='ieditor' value='1' checked>插入编辑器"
				.Write " </dd>" &vbcrlf
		case "uploadphoto"
				If CBool(UpPowerFlag) = True Then
				  .Write " <dd>"
				  .Write "<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?showpic=pic&UPType=Pic&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='30'></iframe>"
				  .Write "</dd>"
				End If
	  End Select
	End If
Next

		.Write "</dl>"
		.Write "</div>"
		
		
  Next
END IF
		
		
			
		If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then
		   .Write " <div class=tab-page id=option-page>"
		   .Write "  <H2 class=tab>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/title").text & "</H2>"
		   .Write "	<SCRIPT type=text/javascript>"
		   .Write "				 tabPane1.addTabPage( document.getElementById( ""option-page"" ) );"
		   .Write "	</SCRIPT>"

            .Write "<dl class='dtable'>"
		  If KS.GetAppStatus("special") and FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='special']/showonform").text="1" Then
	        Call KSCls.Get_KS_Admin_Special(ChannelID,KS.ChkClng(KS.G("ID")))
		  End If
	      If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='adddate']/showonform").text="1" Then
			.Write "              <dd>"
			.Write "                <div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='adddate']/title").text & ":</div>"
			
			If Action <> "Edit" Then
			.Write ("<input name='AddDate' type='text' onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" id='AddDate' value='" & Now() & "' size='50'  class='textbox Wdate'>")
			Else
			.Write ("<input name='AddDate' type='text' onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" id='AddDate' value='" & AddDate & "' size='50'  readonly class='textbox Wdate'>")
			End If
			.Write "<b>日期格式：年-月-日 时：分：秒"
			.Write "               </dd>"
	    End If
	    If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='rank']/showonform").text="1" Then
			.Write "              <dd>"
			.Write "                <div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='rank']/title").text & ":</div>"
			.Write "                <select name='rank'>"
			If Rank = "★" Then .Write "<option  selected>★</option>" Else .Write "<option>★</option>"
			If Rank = "★★" Then .Write "<option  selected>★★</option>" Else .Write "<option>★★</option>"
			If Rank = "★★★" Or Action = "Add" Then .Write "<option  selected>★★★</option>" Else .Write "<option>★★★</option>"
			If Rank = "★★★★" Then .Write "<option  selected>★★★★</option>" Else .Write "<option>★★★★</option>"
			If Rank = "★★★★★" Then .Write "<option  selected>★★★★★</option>" Else .Write "<option>★★★★★</option>"
			.Write "</select>&nbsp;请为" & KS.C_S(ChannelID,3) & "评定阅读等级"
			.Write "               </dd>"
	   End If
	   If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='hits']/showonform").text="1" Then
			.Write "              <dd>"
			.Write "                <div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='hits']/title").text & ":</div>本日：<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='6' style='text-align:center' class='textbox'> 本周：<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='6' style='text-align:center' class='textbox'> 本月：<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='6' style='text-align:center' class='textbox'> 总计：<input name='Hits' type='text' id='Hits' value='" & Hits & "' size='6' style='text-align:center' class='textbox'>&nbsp;初数点击数作弊" 
			.Write "                  </dd>"
	  End If
	  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='template']/showonform").text="1" Then
			.Write "             <dd><div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='template']/title").text & ":</div>"
			.Write "<table>"
			IF Action <> "Edit" and  Action<>"Verify" Then
			.Write " <tr><td><input type='radio' name='templateflag' onclick='GetTemplateArea(false);' value='2' checked>继承栏目设定<input type='radio' onclick='GetTemplateArea(true);' name='templateflag' value='1'>自定义</td></tr>"
			.Write "<tr id='templatearea' style='display:none'><td>"
		   Else
		    .Write "<tr style='font-weight:normal' id='templatearea'><td>"
		   End If
				If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
				.Write "<input id='TemplateID' name='TemplateID' readonly size=50 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
				If KS.WSetting(0)="1" Then 
				.Write "<br/><strong>3G版模板</strong> "
				.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=50 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
				End If
			.Write "</td></tr></table>"
			.Write "                </dd>"
	  End If
	  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='fname']/showonform").text="1" Then
			.Write "             <dd>"
			.Write "              <div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='fname']/title").text & ":</div>"
			
			IF Action = "Edit" or Action="Verify" Then
			.Write "<input name='FileName' type='text' id='FileName' readonly  value='" & FileName & "' size='50' class='textbox'> <span>不能改</span>"
			Else
			.Write "<table>"
			.Write "<tr><td><input type='radio' value='0' name='filetype' onclick='GetFileNameArea(false);' checked>自动生成 <input type='radio' value='1' name='filetype' onclick='GetFileNameArea(true);' >自定义</td></tr>"
			.Write "<tr id='filearea' style='display:none;font-weight:normal'><td><input name='FileName' type='text' id='FileName'   value='" & FileName  & "' size='45' class='textbox'> <font class=""tips"">可带路径,如 help.html,news/news_1.shtml等</font></td></tr>"
			.Write "</table>"
			End IF
			 .Write "             </dd>"
     End If
	 		.Write "              <dd>"
			.Write "               <div>审核状态:</div>"
			.Write " <input type=""hidden"" name=""oldverific"" value=""" & verific &"""/>"
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
			
			.Write "              </dd>"
			.Write "</dl>"
			.Write "</div>"
	   End If
	   
	   If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/showonform").text="1" Then
	     KSCls.LoadSeoOption ChannelID,FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/title").text,SEOTitle,SEOKeyWord,SEODescript
  	   End If
	   
	   If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='signoption']/showonform").text="1" Then
			   .Write " <div class=tab-page id=sign-page>"
			   .Write "  <H2 class=tab>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='signoption']/title").text & "</H2>"
			   .Write "	<SCRIPT type=text/javascript>"
			   .Write "				 tabPane1.addTabPage( document.getElementById( ""sign-page"" ) );"
			   .Write "	</SCRIPT>"
	
				.Write "<dl class='dtable'>"
				.Write "           <dd><div>是否签收:</div>"
				.Write "<label><input type='radio' name='issign' onclick=""$('#signs').hide();"" value='0'"
				If IsSign="0" Then .Write " checked"
				.Write">不需要</label>"
				.Write "<label><input type='radio' name='issign' onclick=""$('#signs').show();"" value='1'"
				If IsSign="1" Then .Write " checked"
				.Write ">需要</label>"
				
				.Write " </dd>"
				If IsSign="0" Then
				.Write "           <span style='display:none' id='signs'>"
				Else
				.Write "           <span  id='signs'>"
				End If
				.Write "          <dd><div>签收用户:</div>"
				.Write "             <textarea name='SignUser' id='SignUser' cols=50 rows=5>" & SignUser & "</textarea>"
				.Write "<br/>&nbsp;<input type='button' value='选择用户' onclick='SelectUser()' class='button'> <input type='button' value='清除用户' onclick=""$('#SignUser').val('')"" class='button'>"
				.Write "</dd>"
				.Write "  <dd>"
				.Write "   <div>时间限制:</div>"
				.Write "<label><input type='radio' name='SignDateLimit' onclick=""$('#signdate').hide();"" value='0'"
				If SignDateLimit="0" Then .Write " checked"
				.Write ">不启用</label>"
				.Write "<label><input type='radio' name='SignDateLimit' onclick=""$('#signdate').show();"" value='1'"
				If SignDateLimit="1" Then .Write " checked"
				.Write">启用</label>"
				.Write "</dd>"
				If SignDateLimit="1" then
				.Write "  <dd id='signdate'>"
				else
				.Write "  <dd id='signdate' style='display:none'>"
				end if
				.Write "    <div>签收结束时间:</div>"
				.Write "           <input type='text' class=""textbox Wdate"" size=""40"" onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" id='SignDateEnd'  name='SignDateEnd' value='" & SignDateEnd & "'><font class=""tips"">签收用户必须在这个时间内完成签收。格式：期格式：年-月-日 时：分：秒</font></dd>"
				.Write "        </span>"
				
				.Write "</dl>"
				.Write "</div>"
			End If   
	        If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonform").text="1" Then
	           KSCls.LoadChargeOption ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
		    End If
			If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='relativeoption']/showonform").text="1" Then
		       KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))
			End If
			 .Write "      </form></div>"
			 .Write "</body>"
			 .Write "</html>"
			End With
			if rs.state=1 then rs.close:Set rs=nothing
		End Sub


		'保存
		Sub DoSave()
			 Dim SelectInfoList,HasInRelativeID,FileIds,Relateda,Related_s,ii
			 With Response
				TitleType      = KS.G("TitleType")
				Title          = Request.Form("Title")
				Fulltitle      = Request.Form("Fulltitle")
				TitleFontColor = KS.G("TitleFontColor")
				TitleFontType  = KS.G("TitleFontType")
                ArticleContent = Request.Form("Content")
				If KS.IsNul(ArticleContent)="" Then ArticleContent="&nbsp;"
				ArticleContent = FilterScript(ArticleContent)
				PageTitle      = Replace(Request.Form("PageTitle"),vbcrlf,"§")
				Hits        = KS.ChkClng(KS.G("Hits"))
				HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
				HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
				HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))

				PhotoUrl      = KS.G("PhotoUrl")
				Changes     = KS.ChkClng(KS.G("Changes"))
				Recommend   = KS.ChkClng(KS.G("Recommend"))
				Rolls       = KS.ChkClng(KS.G("Rolls"))
				Strip       = KS.ChkClng(KS.G("Strip"))
				Popular     = KS.ChkClng(KS.G("Popular"))
				Slide       = KS.ChkClng(KS.G("Slide"))
				Comment     = KS.ChkClng(KS.G("Comment"))
				IsTop       = KS.ChkClng(KS.G("IsTop"))
				IsVideo     = KS.ChkClng(KS.G("IsVideo"))
				Makehtml    = KS.ChkClng(KS.G("Makehtml"))
				ChangesUrl  = Trim(Request("ChangesUrl"))
				SpecialID   = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
				SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
				Inputer     = KS.S("Inputer")
			
				
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
			
			
				KeyWords    = KS.G("KeyWords")
				Author      = KS.G("Author")
				Origin      = KS.G("Origin")
				AddDate     = KS.G("AddDate")
				If Not IsDate(AddDate) Then AddDate=Now
				Rank        = KS.G("Rank")
				Intro       = KS.G("Intro")
				'if Intro="" And KS.ChkClng(KS.G("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(ArticleContent),200)
				if Intro="" And KS.ChkClng(KS.G("AutoIntro"))=1 Then Intro=KS.GotTopic(replace(replace(replace(Replace(KS.LoseHtml(ArticleContent),vbcrlf,""),"　　","")," ",""),chr(32),""),200)
				
				ArticleContent = KS.FilterIllegalChar(ArticleContent)
				Intro          = KS.FilterIllegalChar(Intro)
				Title          = KS.FilterIllegalChar(Title)
				
				IsSign         = KS.ChkClng(KS.G("IsSign"))
				SignUser       = Replace(KS.G("SignUser")," ","")
				SignDateLimit  = KS.ChkClng(KS.G("SignDateLimit"))
				SignDateEnd    = KS.S("SignDateEnd")
				If Not IsDate(SignDateEnd) Then SignDateEnd=Now
				Province       = KS.S("Province")
				City           = KS.S("City")
				County         = KS.S("County")
				FileIds        = LFCls.GetFileIDFromContent(ArticleContent)
				
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then
						Verific=KS.ChkClng(KS.G("Verific"))
				Else
						Verific = 1
				End if
					 
				'收费选项
				ReadPoint   = KS.ChkClng(KS.G("ReadPoint"))
				ChargeType  = KS.ChkClng(KS.G("ChargeType"))
				PitchTime   = KS.ChkClng(KS.G("PitchTime"))
				ReadTimes   = KS.ChkClng(KS.G("ReadTimes"))
				InfoPurview = KS.ChkClng(KS.G("InfoPurview"))
				arrGroupID  = KS.G("GroupID")
				DividePercent=KS.G("DividePercent"):IF Not IsNumeric(DividePercent) Then DividePercent=0
				
				'SEO优化选项
				SEOTitle    = KS.G("SEOTitle")
				SEOKeyWord  = KS.G("SEOKeyWord")
				SEODescript = KS.G("SEODescript")
			
				TemplateID  = KS.G("TemplateID")
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
				If Changes = 1 Then	ArticleContent = ChangesUrl
				If Title = "" Then KS.die ("<script>alert('" & KS.C_S(ChannelID,3) & "标题不能为空!');history.back(-1);</script>")
				If CInt(Changes) = 1 Then
				 If ChangesUrl = "" Then KS.die ("<script>alert('请输入" & KS.C_S(ChannelID,3) & "的链接地址！');history.back(-1);</script>")
				End If
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then
				 If ArticleContent = "" and CInt(Changes)<>1 Then KS.Die("<script>alert('" & KS.C_S(ChannelID,3) & "内容不能为空!');history.back(-1);</script>")
				End If
				If KS.IsNul(ArticleContent) Then ArticleContent="&nbsp;"
				
				Set RS = Server.CreateObject("ADODB.RecordSet")
				If Tid = "" Then ErrMsg = ErrMsg & "[" & KS.C_S(ChannelID,3) & "类别]必选! \n"
				If Title = "" Then ErrMsg = ErrMsg & "[" & KS.C_S(ChannelID,3) & "标题]不能为空! \n"
				If Title <> "" And Tid <> "" And (Action = "Add") Then
				  SqlStr = "select * from " & KS.C_S(ChannelID,2) &" where Title=" & KS.WithKorean() &"'" & replace(Title,"'","''") & "' And Tid='" & Tid & "'"
				   RS.Open SqlStr, conn, 1, 1
					If Not RS.EOF Then ErrMsg = ErrMsg & "该类别已存在此篇" & KS.C_S(ChannelID,3) & "! \n"
				   RS.Close
				End If
				if KS.ChkClng(Request("NewsID"))<>0 then
				  SqlStr = "select * from " & KS.C_S(ChannelID,2) &" where id<>" & KS.ChkClng(Request("NewsID")) & " and Title=" & KS.WithKorean() &"'" & replace(Title,"'","''") & "' And Tid='" & Tid & "'"
				   RS.Open SqlStr, conn, 1, 1
					If Not RS.EOF Then ErrMsg = ErrMsg & "该类别已存在此篇" & KS.C_S(ChannelID,3) & "! \n"
				   RS.Close
				end if
				
				If ErrMsg <> "" Then
				   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
				   .End
				Else
				        If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)
						
						If KS.ChkClng(KS.G("PaginationType"))=2 Then
						 ArticleContent=KS.AutoSplitPage(Request.Form("Content"),"[NextPage]",KS.ChkClng(KS.G("MaxCharPerPage")))
						ElseIf KS.ChkClng(KS.G("PaginationType"))=0 Then
						 ArticleContent=Replace(ArticleContent,"[NextPage]","")
						End If
						If KS.ChkClng(KS.G("BeyondSavePic")) = 1 And CInt(Changes) <> 1 Then
							  SaveFilePath = KS.GetUpFilesDir & "/"
							  KS.CreateListFolder (SaveFilePath)
							  ArticleContent = KS.ReplaceBeyondUrl(ArticleContent, SaveFilePath)
						End If
						
				'自动获取一张图片
				If KS.IsNul(PhotoUrl) And Not KS.IsNul(ArticleContent) Then
				  Dim regEx:Set regEx = New RegExp
				  regEx.IgnoreCase = True
				  regEx.Global = True
				  regEx.Pattern = "src\=.+?\.(gif|jpg|png)"
				  Dim Matches:Set Matches = regEx.Execute(ArticleContent)
				  If regEx.Test(ArticleContent) Then
				   PhotoUrl=Lcase(Matches(0).value)
				   PhotoUrl=replace(PhotoUrl,"src=","")
				   PhotoUrl=replace(PhotoUrl,"""","")
				   PhotoUrl=replace(PhotoUrl,"'","")
				  End If
				End If
				
				
				If Not KS.IsNul(PhotoUrl) Then PicNews=1 Else PicNews=0

					  If Action = "Add" Or KS.ChkClng(KS.G("AddNew"))=1 Then 

						for ii=0 to Ubound(Relateda)
						Set RS = Server.CreateObject("ADODB.RecordSet")
						SqlStr = "select top 1 * from " & KS.C_S(ChannelID,2) &" where 1=0"
						RS.Open SqlStr, conn, 1, 3
						RS.AddNew
						RS("TitleType")      = TitleType
						RS("Title")          = Title
						RS("Fulltitle")      = Fulltitle
						RS("TitleFontColor") = TitleFontColor
						RS("TitleFontType")  = TitleFontType
						RS("Intro")          = Intro
						RS("ArticleContent") = ArticleContent
						RS("PageTitle")      = PageTitle
						RS("Changes")        = Changes
						RS("PicNews")        = PicNews
						RS("PhotoUrl")       = PhotoUrl
						RS("Recommend")      = Recommend
						RS("IsTop")          = IsTop
						RS("IsVideo")        = IsVideo
						RS("Rolls")          = Rolls
						RS("Strip")          = Strip
						RS("Popular")        = Popular
						RS("Verific")        = Verific
						if ii=0 then
							RS("Tid")            = Tid
						else
							RS("Tid")            =RTrim(Trim(Relateda(ii)))
							Tid = RTrim(Trim(Relateda(ii))) 
						end if
						RS("oTid")           = KS.G("oTid")
						RS("oID")            = KS.ChkClng(KS.S("oid"))
						RS("KeyWords")       = KeyWords
						RS("Author")         = Author
						RS("Origin")         = Origin
						RS("AddDate")        = AddDate
						RS("ModifyDate")     = AddDate
						RS("Rank")           = Rank
						RS("Slide")          = Slide
						RS("Comment")        = Comment
						if ii=0 then
						 RS("TemplateID")     = TemplateID
						Else
						 RS("TemplateID")     = KS.C_C(TID,5)
						End If
						RS("WapTemplateID")  = WapTemplateID
						RS("Hits")           = Hits
						RS("HitsByDay")      = HitsByDay
						RS("HitsByWeek")     = HitsByWeek
						RS("HitsByMonth")    = HitsByMonth
						RS("Fname")          = Fname
						RS("Inputer")        = KS.C("AdminName")
						if KS.IsNul(KS.Setting(189))  then
						 RS("RefreshTF")      = Makehtml
						else
						 RS("RefreshTF")      = 0
						end if
						RS("DelTF")          = 0
						RS("PostTable")      = LFCls.GetCommentTable()
						RS("CmtNum")         = 0
						RS("IsSign")         = IsSign
						RS("SignUser")       = SignUser
						RS("SignDateLimit")  = SignDateLimit
						RS("SignDateEnd")    = SignDateEnd
						RS("Province")       = Province
						RS("City")           = City
						RS("County")         = County
						RS("SEOTitle")       = SEOTitle
						RS("SEOKeyWord")     = SEOKeyWord
						RS("SEODescript")    = SEODescript
						RS("ReadPoint")      = ReadPoint
				        RS("ChargeType")     = ChargeType
				        RS("PitchTime")      = PitchTime
				        RS("ReadTimes")      = ReadTimes
						RS("InfoPurview")    = InfoPurview
						RS("arrGroupID")     = arrGroupID
						RS("DividePercent")  = DividePercent
						RS("OrderID")        = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & Tid &"'")(0))+1
						If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonform").text="1" Then	 RS("MapMarker")=KS.G("MapMark")
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
					   '写入Session,添加下一篇文章调用
					   Session("KeyWords") = KeyWords
					   Session("Author")   = Author
					   Session("Origin")   = Origin
					   RS.MoveLast
					  If Left(Ucase(Fname),2)="ID" Or KS.ChkClng(KS.G("AddNew"))=1 Then
					   RS("Fname") = RS("ID") & FnameType
					   RS.Update
					  End If
					  if ii=0 then
						  For I=0 To Ubound(SpecialID)
							Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						  Next
						  Call KSCls.UpdateRelative(ChannelID,RS("ID"),SelectInfoList,0)
					  end if
					  Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
	 				  '关联上传文件
					  Call KS.FileAssociation(ChannelID,Rs("ID"),ArticleContent & PhotoUrl,0)
                     if ii=0 then
						  If Not KS.IsNul(FileIds) Then 
							Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & Rs("ID") &",classID=" & KS.C_C(Tid,9) & " Where ID In (" & FileIds & ")")
							 '删除无用的附件记录,仅后台上传时检测
							 Conn.Execute("Delete From [KS_UploadFiles] Where Isannex=1 and infoid=0")
						  End If
					  end if
					  if ii=0 then
						  If KS.Setting(56)="1" Then '绑定论坛
							If KS.ChkClng(Request("bid"))<>0 Then
							 KS.Echo "<iframe src=""../club/KS.Push.asp?action=doPush&ChannelID=" & ChannelID &"&ItemID=" & Rs("Id") & "&Bid=" & Request("Bid") & "&CategoryId=" & Request("CategoryId") &""" width=""0"" height=""0""></iframe>"
							End If
						  End If
					  end if
					  Call RefreshHtml(1)
					  
					  next 
					  RS.Close:Set RS = Nothing	
					Response.end()
					ElseIf Action = "Edit" Or Action="Verify" Then
					
					 
						
					If Action="Verify" Then 
					 Call KS.ReplaceUserFile(ArticleContent,ChannelID)
					 Call KS.ReplaceUserFile(PhotoUrl,ChannelID)
					End If
					
					NewsID = KS.ChkClng(Request("NewsID"))
					dim RelatedArray,n:n=0
					if KS.ChkClng(KS.G("EditNewtb"))=1 then
						if KS.ChkClng(KS.G("RelatedID"))=0 or KS.ChkClng(KS.G("RelatedID"))=-11 then
							RelatedArray=KSCls.GetRelatedArray( KS.C_S(ChannelID,2), NewsID ,11)'同步文章
						else
							RelatedArray=KSCls.GetRelatedArray( KS.C_S(ChannelID,2), KS.ChkClng(KS.G("RelatedID")) ,22)'同步文章
						end if 
					else
						RelatedArray=Array(NewsID)	
					end if
					
					for ii=0 to UBound(RelatedArray)
					Set RS = Server.CreateObject("ADODB.RecordSet")
					SqlStr = "SELECT top 1 * FROM " & KS.C_S(ChannelID,2) &" Where ID=" & RTrim(Trim(RelatedArray(ii))) & ""
						RS.Open SqlStr, conn, 1, 3
						If RS.EOF And RS.BOF Then
						 .die ("<script>alert('参数传递出错!');history.back(-1);</script>")
						End If
						RS("TitleType")     = TitleType
						RS("Title")         = Title
						RS("Fulltitle")     = Fulltitle
						RS("TitleFontColor")= TitleFontColor
						RS("TitleFontType") = TitleFontType
						RS("ArticleContent")= ArticleContent
						RS("PageTitle")     = PageTitle
						RS("Changes")       = Changes
						RS("PicNews")       = PicNews
						RS("PhotoUrl")      = PhotoUrl
						RS("Recommend")     = Recommend
						RS("IsTop")         = IsTop
						RS("IsVideo")       = IsVideo
						RS("Rolls")         = Rolls
						RS("Strip")         = Strip
						RS("Popular")       = Popular
						if NewsID=KS.ChkClng(RTrim(Trim(RelatedArray(ii)))) then
							RS("Tid")       = KS.G("Tid")
							RS("oTid")      = KS.G("oTid")
							RS("oID")       = KS.ChkClng(KS.S("oid"))
						end if
						RS("KeyWords")      = KeyWords
						RS("Author")        = Author
						RS("Origin")        = Origin
						RS("AddDate")       = AddDate
						RS("ModifyDate")    = Now
						RS("Rank")          = Rank
						RS("Slide")         = Slide
						RS("Comment")       = Comment
						RS("TemplateID")    = TemplateID
						RS("WapTemplateID") = WapTemplateID
						If Action="Verify" Then
						    Inputer         = RS("Inputer")
						End If
						If Verific<>100 Then
						RS("Verific") = Verific
						End If
						If Makehtml = 1 Then
						 RS("RefreshTF") = 1
						End If
						RS("IsSign")        = IsSign
						RS("SignUser")      = SignUser
						RS("SignDateLimit") = SignDateLimit
						RS("SignDateEnd")   = SignDateEnd
						RS("SEOTitle")      = SEOTitle
						RS("SEOKeyWord")    = SEOKeyWord
						RS("SEODescript")   = SEODescript
						RS("Province")      = Province
						RS("City")          = City
						RS("County")        = County
						RS("Hits")          = Hits
						RS("HitsByDay")     = HitsByDay
						RS("HitsByWeek")    = HitsByWeek
						RS("HitsByMonth")   = HitsByMonth
						RS("ReadPoint")     = ReadPoint
				        RS("ChargeType")    = ChargeType
				        RS("PitchTime")     = PitchTime
				        RS("ReadTimes")     = ReadTimes
						RS("InfoPurview")   = InfoPurview
						RS("arrGroupID")    = arrGroupID
						RS("DividePercent") = DividePercent
						RS("Intro")         = Intro
						If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonform").text="1" Then	 RS("MapMarker")=KS.G("MapMark")
						Call KSCls.AddDiyFieldValue(RS,FieldXML)
						RS.Update
						RS.MoveLast
						If TID<>Request.Form("OldClassID") Then
					     Call KSCls.DelInfoFile(ChannelID,Request.Form("OldClassID"),Split(RS("ArticleContent"), "[NextPage]"),RS("Fname"),RS("ID"),RS("AddDate"))
					    End If
						if ii=0 then
							Conn.Execute("Delete From KS_SpecialR Where InfoID=" & NewsID & " and channelid=" & ChannelID)
							For I=0 To Ubound(SpecialID)
								Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
							Next
						end if
						Call KSCls.UpdateRelative(ChannelID,NewsID,SelectInfoList,1)
						Call LFCls.UpdateItemInfo(ChannelID,NewsID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
	 				  '关联上传文件
					  Call KS.FileAssociation(ChannelID,NewsID,ArticleContent & PhotoUrl,1)
					  if ii=0 then
						  If Not KS.IsNul(FileIds) Then 
							 Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & NewsID &",classID=" & KS.C_C(Tid,9) & " Where ID In (" & FileIds& ")")
							 '删除无用的附件记录,仅后台上传时检测
							 Conn.Execute("Delete From [KS_UploadFiles] Where Isannex=1 and infoid=0")
						  End If
					  end if
					  
					  if ii=0 then
						  If KS.Setting(56)="1" Then '绑定论坛
							If KS.ChkClng(Request("bid"))<>0 Then
							  If  KS.ChkClng(RS("PostID"))<>0 Then
							  Conn.Execute("Update KS_GuestBook Set BoardID=" & KS.ChkClng(Request("bid")) & ",categoryID=" & KS.ChkClng(Request("CategoryID")) & " Where ID=" & KS.ChkClng(RS("PostID")))
							  Else
								KS.Echo "<iframe src=""../club/KS.Push.asp?action=doPush&ChannelID=" & ChannelID &"&ItemID=" & Rs("Id") & "&Bid=" & Request("Bid") & "&CategoryId=" & Request("CategoryId") &""" width=""0"" height=""0""></iframe>"
							  End If
							End If
						  End If
					  end if
					   tid=rs("tid")
					   Call RefreshHtml(2)
					   next
					   RS.Close:Set RS = Nothing
						IF (Action="Verify" or (Not KS.IsNul(KS.S("oldverific")) AND KS.S("oldverific")="0" And Verific=1)) And Inputer<>KS.C("AdminName")  Then     '如果是审核投稿文章，对用户，进行加积分等，并返回签收文章管理
							  '对用户进行增值，及发送通知操作

							  IF Inputer<>"" And Inputer<>KS.C("AdminName") Then 
							   Call KS.SignUserInfoOK(ChannelID,Inputer,Title,NewsID)
							      session("scoremustin")="true"
								  dim Money,Point,Score
								  Point=KS.ChkClng(KS.G("UserAddPoint"))
								  Score=KS.ChkClng(KS.G("UserAddScore"))
								  if Point>0 then Call KS.PointInOrOut(ChannelID,NewsID,Inputer,1,Point,"System","投稿[" &Title & "]审核通过额外奖励!",1)
								  If Score>0 Then Call KS.ScoreInOrOut(Inputer,1,Score,"System","投稿[" &Title & "]审核通过额外奖励!",ChannelID,NewsID)	
								  							
							  End If
							     KS.Echo ("<script> parent.frames['MainFrame'].focus();alert('恭喜，" & KS.C_S(ChannelID,3) & "成功审核!');location.href='../System/KS.ItemInfo.asp?ShowType=1&ChannelID=" & ChannelID &"&Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr="& server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>签收会员" & KS.C_S(ChannelID,3)) & "</font>';</script>") 
				       End IF
					   
						If KeyWord <> "" Then
							KS.Echo  ("<script> parent.frames['MainFrame'].focus();setTimeout(function(){alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='../System/KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ArticleSearch&OpStr=" & Server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>搜索结果</font>") & "';},2500); </script>")
						End If
					End If
				End If
				End With
			End Sub
			
			
			
			Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="添加":EditStr="修改" & KS.C_S(ChannelID,3) & "":AddStr="继续添加" & KS.C_S(ChannelID,3) & ""
				Else
				  TempStr="修改":EditStr="继续修改" & KS.C_S(ChannelID,3) & "":AddStr="添加" & KS.C_S(ChannelID,3) & ""
				End If
			    With Response
				     .Write "<!DOCTYPE html><html>"
			         .Write"<head>"
				     .Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<meta http-equiv=Content-Type content=""text/html; charset=utf-8"">"
					 .Write "<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>"
					 .Write"</head>"
					 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
					 .Write " <div class='pageCont2 mt20'><div class='tabTitle'>系统操作提示信息</div><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
                      .Write "    <tr class='tdbg' colspan=2>"
					  .Write "          <td align='center'><table width='100%' border='0'><tr><td style='width:200px;text-align:center'><img src='../images/succeed.gif'>"
					  .Write "</td><td><div style='padding-left:30px;font-weight:bold'>恭喜，" & TempStr &"" & KS.C_S(ChannelID,3) & "成功！</div>"
			           '判断是否立即发布
					   If Makehtml = 1 Then
					      .Write "<div style=""float:left;margin-top:15px;height:220; overflow: auto; width:100%"">" 
						  
						  If KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then 'PC版本
						  	 .Write "<div><iframe  scrolling='no' src=""../Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						  Else
						  .Write "<div style=""height:25px;""><li>由于" & KS.C_S(ChannelID,1) & "没有启用生成HTML的功能，所以ID号为 <font color=red>" & NewsID & "</font>  的" & KS.C_S(ChannelID,3) & "没有生成!</li></div> "
						  End If
						  
						  If KS.WSetting(0)="1" Then  '手机版
						   If KS.ChkClng(KS.M_C(ChannelID,28))=1  Or KS.ChkClng(KS.M_C(ChannelID,28))=2 Then
						  	 .Write "<div><iframe  scrolling='no' src=""../Include/RefreshHtmlSave.Asp?from=3g&ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						   End If
						  End If
						  
							If KS.C_S(ChannelID,7)<>1 Then
							  .Write "<div style=""height:25px;""><li>由于" & KS.C_S(ChannelID,1) & "的栏目页没有启用生成HTML的功能，所以ID号为 <font color=red>" & TID & "</font>  的栏目没有生成!</li></div> "
							Else
							 If KS.C_S(ChannelID,9)<>1 Then
								  Dim FolderIDArr:FolderIDArr=Split(left(KS.C_C(Tid,8),Len(KS.C_C(Tid,8))-1),",")
								  For I=0 To Ubound(FolderIDArr)
								  .Write "<div align=center><iframe src=""../Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true' scrolling='no'></iframe></div>"
								   Next
							 End If
						   End If
					   If Split(KS.Setting(5),".")(1)="asp" or KS.C_S(ChannelID,9)<>3 Then
					   Else
					     .Write "<div align=center><iframe  scrolling='no' src=""../Include/RefreshIndex.asp?ChannelID=" & ChannelID &"&RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div></div>"
					End If
					.Write   "</td></tr></table></td></tr>"
					.Write "	  <tr>"
					.Write "		<td  class='tdbg' height=""25"" align=""right"" colspan=2>【<a href=""KS.Article.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" & RS("ID") & """><strong>" & EditStr &"</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.Article.asp?ChannelID=" & ChannelID &"&Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=" & Server.URLEncode("添加" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='../System/KS.ItemInfo.asp?ID=" & Tid & "&ChannelID=" & ChannelID &"&Page=" & Page&"&keyword=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>" & KS.C_S(ChannelID,3) & "管理</strong></a>】&nbsp;【<a href=""" & KS.GetDomain & "Item/Show.asp?m=" & ChannelID & "&d=" & RS("ID") & """ target=""_blank""><strong>预览" & KS.C_S(ChannelID,3) & "内容</strong></a>】</td>"
					.Write "	  </tr>"
					.Write "	</table></div>"	
					.Flush			
			End With
		End Sub
	
		

	
	Sub SelectUser()
		response.cachecontrol="no-cache"
		response.addHeader "pragma","no-cache"
		response.expires=-1
		response.expiresAbsolute=now-1
		With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
			.Write "<META HTTP-EQUIV=""pragma"" CONTENT=""no-cache"">" 
			.Write "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache, must-revalidate"">"
			.Write "<META HTTP-EQUIV=""expires"" CONTENT=""Wed, 26 Feb 1997 08:21:57 GMT"">"
            .Write "<base target='_self'>" & vbCrLf
			.Write "<title>选择用户</title>"
			.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<body>"
			%>
			
		<form method='post' name='myform' action=''>	
		<table width='98%' align="center" border='0' align='center' style="margin-top:4px" cellpadding='2' cellspacing='0' class='border'>
		  <tr class='title' height='22'>
			<td valign='top' colspan="2"><b>已经选定的用户名：</b></td>
		  </tr>
		  <tr class='tdbg'>
			<td><input type='text' name='UserList' size='40' maxlength='200' readonly='readonly'></td>
			<td align='center'><input type='button' name='del1' onclick='del(1)' class='button' value='删除最后'> <input type='button' name='del2' onclick='del(0)' value='删除全部' class='button'></td>
		  </tr>
		</table>
		<br/>
		<table width='98%' align="center" border='0' align='center' cellpadding='2' cellspacing='0' class='border'>
  <tr height='22' class='title'>
    <td><b><font color=red>会员</font>列表：</b></td><td align=right><input name='Key' type='text' size='20' value=>&nbsp;&nbsp;<input type='submit' class="button" value='查找'></td>
  </tr>
  <tr>
    <td valign='top' colspan=2>
	<table width='98%' align="center" border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>
	 <%
	 Page=KS.ChkClng(request("page"))
	 if Page=0 Then Page=1
	 MaxPerPage=40
	 dim sqlstr,AllUserList,TotalPages,param
	 if request("key")<>"" then
	   param=" where username like '%" & KS.G("Key") & "%'"
	 end if
	 
	 sqlstr="select username from ks_user " & Param & " order by userid"
	 dim rs:set rs=server.CreateObject("adodb.recordset")
	 RS.Open SQLStr, conn, 1, 1
	 If Not RS.EOF Then
			totalPut = Conn.Execute("Select count(userid) from [ks_user] " & Param)(0)
								If Page < 1 Then Page = 1
								If (Page - 1) * MaxPerPage < totalPut Then
										RS.Move (Page - 1) * MaxPerPage
								Else
										Page = 1
								End If
								
					Dim SQL:SQL=RS.GetRows(MaxPerPage)
					RS.Close : Set RS=Nothing
			  .write "<tr>"
			For I=0 To Ubound(SQL,2)
				If AllUserList = "" Then
					AllUserList = SQL(0,I)
				Else
					AllUserList = AllUserList & "," & SQL(0,I)
				End If
			  .write "<td align='center'><a href='#' onclick='add(""" &SQL(0,I) & """)'>" &SQL(0,I) & "</a></td>"
			  If ((i+1) Mod 8) = 0 And i > 0 Then Response.Write "</tr><tr>"
			Next
			  .Write "</tr>"
	End If
	%>
	  <tr class='tdbg'>
		<td align='center' colspan=8 height=30><a href='#' onclick='add("<%=AllUserList%>")'><b>增加以上所有用户名</b></a></td>
	  </tr>
	</table>
  </td>
  </tr>
 </table>
		</form>
		
	<table width='98%' align="center" border='0' cellspacing='1' cellpadding='1'>
    <tr>
	 <td>
  <%
  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
 %>	
    </td>
  </tr>
  </table>
		
  <div style="margin-top:10px;text-align:center">
  <input type="button" onClick="doBack()" class="button" value="确认返回">
  <input type="button" onClick="top.box.close()" class="button" value="取消关闭">
  </div>
		
<script language="javascript">
myform.UserList.value='<%=request("DefaultValue")%>';
var oldUser='';
function doBack(){
  top.frames['MainFrame'].document.getElementById('SignUser').value=myform.UserList.value;
  top.box.close();
}
function add(obj)
{
    if(obj==''){return false;}
    if(myform.UserList.value=='')
    {
        myform.UserList.value=obj;
       // window.returnValue=myform.UserList.value;
        return false;
    }
    var singleUser=obj.split(',');
    var ignoreUser='';
    for(i=0;i<singleUser.length;i++)
    {
        if(checkUser(myform.UserList.value,singleUser[i]))
        {
            ignoreUser=ignoreUser+singleUser[i]+" "
        }
        else
        {
            myform.UserList.value=myform.UserList.value+','+singleUser[i];
        }
    }
    if(ignoreUser!='')
    {
        alert(ignoreUser+'用户名已经存在，此操作已经忽略！');
    }
    //window.returnValue=myform.UserList.value;
}
function del(num)
{
    if (num==0 || myform.UserList.value=='' || myform.UserList.value==',')
    {
        myform.UserList.value='';
        return false;
    }

    var strDel=myform.UserList.value;
    var s=strDel.split(',');
    myform.UserList.value=strDel.substring(0,strDel.length-s[s.length-1].length-1);
   // window.returnValue=myform.UserList.value;
}
function checkUser(UserList,thisUser)
{
  if (UserList==thisUser){
        return true;
  }
  else{
    var s=UserList.split(',');
    for (j=0;j<s.length;j++){
        if(s[j]==thisUser)
            return true;
    }
    return false;
  }
}
</script>
			<%
			.Write "</body>"
			.Write "</html>"
		End With
	 End Sub
	
		'执行过滤
	Function FilterScript(ByVal Content)
		   If KS.G("FilterIframe") = "1" Then  Content = KS.ScriptHtml(Content, "Iframe", 1)
		   If KS.G("FilterObject") = "1" Then  Content = KS.ScriptHtml(Content, "Object", 2)
		   If KS.G("FilterScript") = "1" Then  Content = KS.ScriptHtml(Content, "Script", 2)
		   If KS.G("FilterDiv")    = "1" Then  Content = KS.ScriptHtml(Content, "Div", 3)
	       If KS.G("FilterTable")  = "1" Then  Content = KS.ScriptHtml(Content, "table", 3)
		   If KS.G("FilterTr")     = "1" Then  Content = KS.ScriptHtml(Content, "tr", 3)
	       If KS.G("FilterTd")     = "1" Then  Content = KS.ScriptHtml(Content, "td", 3)
		   If KS.G("FilterSpan")   = "1" Then  Content = KS.ScriptHtml(Content, "Span", 3)
		   If KS.G("FilterImg")    = "1" Then  Content = KS.ScriptHtml(Content, "Img", 3)
		   If KS.G("FilterFont")   = "1" Then  Content = KS.ScriptHtml(Content, "Font", 3)
		   If KS.G("FilterA")      = "1" Then  Content = KS.ScriptHtml(Content, "A", 3)
		   If KS.G("FilterHtml")   = "1" Then  Content = KS.LoseHtml(Content)
		   FilterScript=Content
	End Function

End Class
%> 

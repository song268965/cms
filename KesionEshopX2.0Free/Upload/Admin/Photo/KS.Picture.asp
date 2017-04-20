<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_Picture
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Picture
        Private KS,KSCls
		'=====================================定义本页面全局变量=====================================
		Private ID, I, totalPut, Page, RS,ComeFrom
		Private KeyWord, SearchType, StartDate, EndDate, ParentRs, SearchParam,MaxPerPage,SpecialID
		Private T, TitleStr, AttributeStr,RelatedID
		Private FolderID, TemplateID,WapTemplateID,TN, TI,TJ,Action,OTid,OID,Province,City,County
		Private PicID, Title, PhotoUrl, PictureContent, PicUrls, Recommend,IsTop
		Private Popular, Strip, Verific, Comment, Slide, ChangesUrl, Rolls, KeyWords, Author, Origin, AddDate, Rank, Hits, HitsByDay, HitsByWeek, HitsByMonth
		Private CurrPath, InstallDir,PreViewObj, UpPowerFlag,Inputer,SaveFilePath,MapMarker
		Private ComeUrl,ChannelID,FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,Score,ShowStyle,PageNum
		Private ReadPoint,ChargeType,PitchTime,ReadTimes,InfoPurview,arrGroupID,DividePercent
		Private FieldXML,FieldNode,FNode,FieldDictionary
		Private SEOTitle,SEOKeyWord,SEODescript
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
			Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
			Call KSCls.LoadModelField(ChannelID,FieldXML,FieldNode)
			
			'收集搜索参数
			KeyWord   = KS.G("KeyWord")
			SearchType= KS.G("SearchType")
			StartDate = KS.G("StartDate")
			EndDate   = KS.G("EndDate")
			Action     = KS.G("Action")
			ComeFrom   = KS.G("ComeFrom")
			SearchParam = "ChannelID=" & ChannelID
			If Action="CheckTitle" Then
				Call KSCls.CheckTitle()  
				Exit Sub  
			ElseIf Action="SelectClass" Then
			   Call KSCls.SelectMutiClass()
			   Exit Sub
			End If
			
			If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
			If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
			If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
			If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
			If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
			If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
			
			ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
			Action = Trim(KS.G("Action"))
			Page = KS.G("page")
			IF KS.G("Method")="Save" Then
				 Call PictureSave()
			Else 
				 Call PictureAdd()
			End If
		End Sub

        '添加
        Sub PictureAdd() 
			With Response
			CurrPath = KS.GetUpFilesDir()
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Action = "Add" Then
			  FolderID = Trim(KS.G("FolderID"))
			  
			  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '检查是否有添加图片的权限
			   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&ChannelID=" & ChannelID &"';</script>")
			   Call KS.ReturnErr(1, "")
			   Exit Sub
			  End If
			  Hits = 0:HitsByDay = 0: HitsByWeek = 0:HitsByMonth = 0:Comment = 1:IsTop=0
			  ReadPoint=0:PitchTime=24:ReadTimes=10:Score=0 : ShowStyle=4: PageNum=12
			  PreViewObj = "<br><br><br>" & KS.C_S(ChannelID,3) & "预览区"
			  KeyWords = Session("keywords")
			  Author = Session("Author")
			  Origin = Session("Origin")
			
			ElseIf Action = "Edit" Or Action="Verify" Then

			   Set RS = Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & KS.ChkClng(KS.G("ID")), conn, 1, 1
			   If RS.EOF And RS.BOF Then
				Call KS.Alert("参数传递出错!", ComeUrl)
				Set KS = Nothing:.End:Exit Sub
			   End If
				PicID = Trim(RS("ID"))
				FolderID = Trim(RS("Tid"))
				If Action ="Edit" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '检查是否有编辑图片的权限
				 RS.Close:Set RS = Nothing
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
				 Exit Sub
			   End If
			   IF Action="Verify" And KS.C("Role")="1" Then 
			     RS.Close:Set RS = Nothing
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
				 
				 Exit Sub   
			   End If
			   
				Title    = Trim(RS("title"))
				PhotoUrl = Trim(RS("PhotoUrl"))
				PreViewObj = "<img src='" & PhotoUrl & "' border='0'>"
				PicUrls  = Trim(RS("PicUrls"))
				PictureContent = Trim(RS("PictureContent")) : If KS.IsNul(PictureContent) Then PictureContent=" "
				Rolls    = CInt(RS("Rolls"))
				Strip    = CInt(RS("Strip"))
				Recommend = CInt(RS("Recommend"))
				Popular  = CInt(RS("Popular"))
				Verific  = CInt(RS("Verific"))
				Comment  = CInt(RS("Comment"))
				IsTop    = (RS("IsTop"))
				Slide    = CInt(RS("Slide"))
				AddDate  = CDate(RS("AddDate"))
				Rank     = Trim(RS("Rank"))
				FileName = RS("Fname")
				Province = RS("Province")
				City     = RS("City")
				County     = RS("County")
				
				TemplateID    = RS("TemplateID")
				WapTemplateID = RS("WapTemplateID")
				Hits          = Trim(RS("Hits"))
				HitsByDay     = Trim(RS("HitsByDay"))
				HitsByWeek    = Trim(RS("HitsByWeek"))
				HitsByMonth   = Trim(RS("HitsByMonth"))
				Score         = RS("Score")
				KeyWords      = Trim(RS("KeyWords"))
				Author        = Trim(RS("Author"))
				Origin        = Trim(RS("Origin"))
				FolderID      = RS("Tid")
				OTid          = RS("OTid")
				OId           = RS("OId")
				ShowStyle     = RS("ShowStyle")
				PageNum       = RS("PageNum")
				ReadPoint     = RS("ReadPoint")
				ChargeType    = RS("ChargeType")
				PitchTime     = RS("PitchTime")
				ReadTimes     = RS("ReadTimes")
				InfoPurview   = RS("InfoPurview")
				arrGroupID    = RS("arrGroupID")
				DividePercent = RS("DividePercent")
				SEOTitle      = RS("SEOTitle")
				SEOKeyWord    = RS("SEOKeyWord")
				SEODescript   = RS("SEODescript")
				RelatedID      = RS("RelatedID")
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonform").text="1" Then MapMarker      = RS("MapMarker")
				
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
			'取得上传权限
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
			
            .Write"<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrlf
			.Write "<title>添加</title>" & vbCrlf
			.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>" & vbCrlf
			.Write "<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>" & vbCrlf
			.Write "<script language='JavaScript' src='../../KS_Inc/common.js'></script>" & vbCrlf
			.Write "<script src=""../../KS_Inc/DatePicker/WdatePicker.js""></script>" & vbCrlf
			.Write "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
			.Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
			.Write EchoUeditorHead()
			
			Call KSCls.EchoFormStyle(ChannelId)   '控制添加文档布局
			
			.Write "</head>" & vbCrlf
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>" & vbCrlf
			.Write "<div>" & vbCrlf
			.Write "<ul id='menu_top' class='menu_top_fixed'>"
			.Write "<li onclick=""return(SubmitFun())"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon save'></i>确定保存</span></li>"
			.Write "<li onclick=""history.back();"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消返回</span></li>"
		    .Write "</ul>" & vbCrlf
			.Write "<div class=""menu_top_fixed_height""></div>"
			
			.Write "<div class=tab-page id=PhotoPane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""PhotoPane"" ), 1 )"
			.Write " </SCRIPT>"
				 
			.Write "    <form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' id='myform' name='myform' >"
			.Write "      <input type='hidden' value='" & PicID & "' name='PicID'>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
			.Write "      <input type='hidden' name='Inputer' value='" &Inputer & "'>"
			
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
				.Write " <dd><div>" & XTitle & ":</div><input name='title' type='text' id='title'  class='rule textbox' value=""" & Title & """ size='80'/> <font color='#FF0000'>*</font> "
				.Write "<input class='button' type='button' value='重名检测' onclick=""if($('#title').val()==''){ top.$.dialog.alert('请输入" & KS.C_S(ChannelID,3) & "标题!');}else top.openWin('" & KS.C_S(ChannelID,3) & "重名检测','photo/KS.Picture.asp?ChannelID=" & ChannelID & "&Action=CheckTitle&title='+escape($('#title').val()),false,360,370);"">"

				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/showonform").text="1" Then
				.Write "<input type='checkbox' name='MakeHtml' value='1' checked>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/title").text
				End IF
				if RelatedID=-11 or KS.ChkClng(RelatedID)<>0 then
							.Write "<span style=""padding:5px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6""><label><input type='checkbox' name='EditNewtb' value='1' checked/> 此"  & KS.C_S(ChannelID,3) & "发布到多个栏目，选中将同步更新 <input type='hidden' name='RelatedID' value='"& RelatedID &"'/></label></span>"
				End if
				.Write "   </dd>" &vbcrlf
		  case "tid"
					.Write " <dd style=""""><div>" & Replace(XTitle,"栏目",KS.GetClassName(ChannelID)) & ":</div>"
					.Write " <input type='hidden' name='OldClassID' value='" & FolderID & "'>"
					If Action<>"Edit" Then
						.Write "&nbsp;<input name='Istidtb' type='button' class='button' id='istidtb' value='发布多" & Replace(XTitle,"栏目",KS.GetClassName(ChannelID)) & "'  onclick=""sel();"" >"
					 	
					end if	
					.Write "<select size='1' name='tid' id='tid' style='width:335px'>"
					.Write " <option value='0'>--请选择" & KS.GetClassName(ChannelID) &"--</option>"
					.Write Replace(KS.LoadClassOption(ChannelID,true),"value='" & FolderID & "'","value='" & FolderID &"' selected") & " </select>"
					%>
					<input type="hidden" id="tidtb" name="tidtb" value=""/>
					<script>
					var box=''
					function sel(){
					top.openWin(false,'photo/KS.Picture.asp?channelID=<%=ChannelID%>&FolderID='+$("#tidtb").val()+'&action=SelectClass',false,400,420);
					}
					</script>
					<%
					
			  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/showonform").text="1" Then
				.Write "&nbsp;&nbsp;" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/title").text & " <label><input name='Recommend' type='checkbox' id='Recommend' value='1'"
				If Recommend = 1 Then .Write (" Checked")
				.Write ">推荐</label><label><input name='Rolls' type='checkbox' id='Rolls' value='1'"
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
				If Slide = 1 Then	.Write (" Checked")
				.Write ">幻灯</label>"
				Call KSCls.GetDiyAttribute(FieldXML,FieldDictionary)
				.Write " </dd>" &vbcrlf
			  End If
		  case "otid"
		        Call KSCls.EchoOTidInfo(FNode,OTid,Oid)
		  case "map"
				.Write " <dd>"
				.Write "  <div>" & XTitle &":</div>经纬度：<input size='43' value=""" & MapMarker & """ type='text' name='MapMark' id='MapMark' class='textbox' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='../images/accept.gif' align='absmiddle' border='0'>添加电子地图标志</a></dd>" &vbcrlf
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
		  case "photourl"
				.Write "<dd id='mode1'>"
				.Write "  <div>" &XTitle & ":</div><input name='PhotoUrl' type='text' id='PhotoUrl' size='50' value='" & PhotoUrl & "' class='textbox'>"
				.Write "   <font color='#FF0000'>*</font>&nbsp;<input class='button' type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID & "&CurrPath=" & CurrPath & "',550,290,window,$('#PhotoUrl')[0]);""> <input class='button' type='button' name='Submit' value='远程抓图...' onClick=""top.openWin('抓取远程图片','include/SaveBeyondfile.asp?fieldid=PhotoUrl&CurrPath=" & CurrPath & "',false,500,100);"">"
				.Write "  <input class=""button""  type='button' name='Submit' value='裁剪...' onClick=""if($('#PhotoUrl').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindow(1,'" & KS.Setting(3) & "',$('#PhotoUrl').val())}"">"
				If Action="Add" Then
				.Write "<br/><label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>使用图集的第一幅图</label>"
				End If
				.Write "     </dd>" &vbcrlf
				.Write "   <dd>"
				.Write "     <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../System/KS.UpFileForm.asp?UPType=Pic&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='30'></iframe>"
				.Write "    </dd>" &vbcrlf
		  case "content"
		  		.Write "<dd><div>显示样式:</div><table width='80%'><tr><td><input type='radio' onclick=""$('#pagenums').hide();"" name='showstyle' value='4'"
				If ShowStyle="4" Then .Write " checked"
				.Write "><img src='../../images/default/p4.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td><input type='radio' onclick=""$('#pagenums').hide();"" name='showstyle' value='1'"
				If ShowStyle="1" Then .Write " checked"
				.Write "><img src='../../images/default/p1.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>"
				.Write "<td><input type='radio' onclick=""$('#pagenums').show();"" name='showstyle' value='2'"
				If ShowStyle="2" Then .Write " checked"
				.Write "><img src='../../images/default/p2.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>"
				.Write "<td><input type='radio' onclick=""$('#pagenums').show();"" name='showstyle' value='3'"
				If ShowStyle="3" Then .Write " checked"
				.Write"><img src='../../images/default/p3.gif'></td>"
				.Write "<td><input type='radio' onclick=""$('#pagenums').hide();"" name='showstyle' value='5'"
				If ShowStyle="5" Then .Write " checked"
				.Write"><img src='../../images/default/p5.gif'></td>"
				.Write "</tr></table><font id=""pagenums"""
				If ShowStyle="1" or ShowStyle="4" Then .Write " style='display:none'"
				.Write ">每页显示<input type=""text"" name=""pagenum""  class=""textbox"" value=""" & PageNum & """ style=""text-align:center;width:30px"">张</font></td></tr>"
	
				If KS.G("Action")<>"Add" Then
				.Write " <dd style='display:none'>"
				Else
				.Write " <dd>"
				End If
				.Write " <div>添加模式:</div>"
				.Write "<label><input type='radio' name='addmode' value='0' checked onclick='$(""#addmore"").hide();$(""#addarea"").show();'>直接添加</label> <label><input type='radio' name='addmode' value='1' onclick='$(""#addmore"").show();$(""#addarea"").hide()'>批量添加</label>"
				.Write "</dd>"
				Dim CurrDate:CurrDate=Year(Now) &right("0"&Month(Now),2)
				Dim CurrDay:CurrDay=CurrDate & right("0"&day(Now),2)
				.Write "   <dd id='addmore' style='display:none'>"
				.Write "     <div>" & KS.C_S(ChannelID,3) & "地址:</div>"
				.Write "     <input  name='MorePicUrl' type='text' id='MorePicUrl' size='90' value='图片#|" &  CurrPath & "/"&CurrDate &"/" & CurrDay & "#.jpg|" &  CurrPath & "/"&CurrDate &"/" & CurrDay & "#_S.jpg' class='textbox'>&nbsp;开始ID：<input class='textbox' type='text' value='1' name='morestart' size=5> 结束ID：<input class='textbox' type='text' value='100' name='moreend' size=5><br/><font  class=""tips"">1、数字序号通配符为#，注意通配符只用一个#即可&nbsp;&nbsp;2、格式：图片介绍|大图地址|小图地址</font>"
				.Write "  </dd>"
	
	
				.Write "<span id='addarea'>"			
				.Write "<dd><div>" & XTitle & ":<font>(<input type='checkbox' value='1' name='BeyondSavePic' checked>采集存图"
				 If KS.ChkClng(KS.M_C(ChannelID,30))=1 Then Response.Write "<br/>"
				 %>
				<label<%if KS.TBSetting(5)="0" then response.write " style='display:none'"%>><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>添加水印</label>)
				</font></div>
				<style type="text/css">
				#thumbnails{background:url(../../plus/swfupload/images/albviewbg.gif) no-repeat;min-height:200px;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
				#thumbnails .zsimg {padding: 0 10px;text-align: left;}
				#thumbnails .zsimg .textbox{ margin-left:0; margin-bottom:0 !important;}

				#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:190px;height:200px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
				#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px; margin:5px;}
				#thumbnails div.thumbshow div span{ display:block; text-align:center; margin-top:5px;line-height:20px; height:20px;}
				#thumbnails div.thumbshow div span a{ color:#006699;font-weight:normal}
				.progressName{ height:auto !important;}
				</style>
				<link href="../../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
				<script type="text/javascript" src="../../plus/swfupload/swfupload/swfupload.js"></script>
				<script type="text/javascript" src="../../plus/swfupload/js/handlers.js"></script>
				<script type="text/javascript" src="../../KS_inc/boxtcshow.js"></script>
				<script type="text/javascript">
						var swfu;
						var pid=0;
						function SetAddWater(obj){
						 if (obj.checked){
						 swfu.addPostParam("AddWaterFlag","1");
						 }else{
						 swfu.addPostParam("AddWaterFlag","0");
						 }
						}
						//删除已经上传的图片
						function DelUpFiles(pid)
						{  var p=$('#pic'+pid).val();
						   if (p!==''){
							$.ajax({
							  url: "../../plus/ajaxs.asp",
							  cache: false,
							  data: "action=DelPhoto&pic="+p+"&flag=1",
							  success: function(r){
							  }
							  });
						   }
						   $("#thumbshow"+pid).remove();	
						}	
						
						function addImage(bigsrc,smallsrc,text,sorn) {
							var newImgDiv = document.createElement("div");
							var delstr = '';							
							delstr = '<a href="javascript:DelUpFiles('+pid+')" style="color:#ff6600">[删除]</a>';
							newImgDiv.className = 'thumbshow';
							newImgDiv.id = 'thumbshow'+pid;
							document.getElementById("thumbnails").appendChild(newImgDiv);
							newImgDiv.innerHTML = '<a href="'+bigsrc+'" target="_blank"><span id="show'+pid+'"><img src="'+smallsrc+'" /></span></a>';
							newImgDiv.innerHTML += '<div class="zsimg">'+delstr+' <b>注释：</b><input type="hidden" class="pics" id="pic'+pid+'" value="'+bigsrc+'|'+smallsrc+'"/><input type="text" class="textbox" name="picinfo'+pid+'" value="'+text+'" style="width:155px;" /> <span><a  title="左移动排序" href="javascript:;" onclick="pic_move(this,1);">←左移动</a>&nbsp;&nbsp;&nbsp;<a title="右移动排序" href="javascript:;" onclick="pic_move(this,2);">右移动→</a></span></div>';
							pid++;
							
						}
					
						window.onload = function () {
							swfu = new SWFUpload({
								// Backend Settings
								upload_url: "../include/swfupload.asp",
								post_params: {UPType:"pic","AdminID" : "<%=KS.C("AdminID") %>","AdminPass":"<%=KS.C("AdminPass")%>",AddWaterFlag:"1","BasicType":<%=KS.C_S(ChannelID,6)%>,"ChannelID":<%=ChannelID%>,"AutoRename":4},
				
								// File Upload Settings
								file_size_limit : 1024*2,	// 2MB
								file_types : "*.jpg; *.gif; *.png",
								file_types_description : "图片格式",
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
								//button_image_url : "../plus/swfupload/images/SmallSpyGlassWithTransperancy_17x18d.png",
								button_placeholder_id : "spanButtonPlaceholder",
								button_width: 195,
								button_height: 18,
								button_text : '<span class="button">本地批量上传(单图限制2 MB)</span>',
								button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#ffffff; } ',
								button_text_top_padding: 0,
								button_text_left_padding: 0,
								button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
								button_cursor: SWFUpload.CURSOR.HAND,
								
								// Flash Settings
								flash_url : "../../plus/swfupload/swfupload/swfupload.swf",
								flash9_url : "../../plus/swfupload/swfupload/swfupload_FP9.swf",
				
								custom_settings : {
									upload_target : "divFileProgressContainer"
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
					   box=$.dialog.open("../../editor/ksplus/remotefile.asp",{title:"网上采集图片",width:550,height:200});
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
					   top.$.dialog.alert('请选择一张小图地址!');
					   return false;
					  }
					  if (x2==''){
					   top.$.dialog.alert('请选择一张大图地址!');
					   return false;
					  }
					  addImage(x2,x1,x3,"")
					   box.close();
					}
					function ProcessCollect(collecthttp){
					 if (collecthttp==''){
					   top.$.dialog.alert('请输入远程图片地址,一行一张地址!');
					   return false;
					 }
					 var carr=collecthttp.split('\n');
					 for(var i=0;i<carr.length;i++){
					   if (carr[i]!=''){
					   var bigsrc=carr[i];
					   var smallsrc=carr[i];
					   addImage(bigsrc,smallsrc,'',"")
					   }
					 }
					 box.close();
					}
					
					</script>
					
			<table>
			 <tr>
			  <td><div class="button"><span id="spanButtonPlaceholder"></span></div></td>
			 <td><button type="button"  class="button" onclick="OnlineCollect()">网上采集</button> &nbsp;
			 <button type="button"  class="button" onClick="AddTJ();">图片库...</button></td>
			 </tr>
		    </table>
			<table width="90%">
			 <tr>
			   <td id="divFileProgressContainer"></td>
			 <tr>
			 <tr>
			   <td id="thumbnails"></td>
			 <tr>
			</table>
				<input type='hidden' name='PicUrls' id='PicUrls'>
				<%
				.Write "</td></tr>"
				.Write "</tbody>"
		  case "picturecontent"
                .Write "<dd>"
				.Write "  <div>" & XTitle & ":</div>"
				
				 .Write EchoEditor("Content",PictureContent,"Basic","90%","180px")
			 
			    .Write " </dd>"& vbcrlf	  
		   case ""
		  case ""
		  case ""
		  
	   End Select
	End If
Next
           .Write "</dl>"
		   .Write "</div>"
  Next
END IF
	
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then	 
		   .Write " <div class=tab-page id=classoption-page>"
		   .Write "  <H2 class=tab>属性设置</H2>"
		   .Write "	<SCRIPT type=text/javascript>"
		   .Write "				 tabPane1.addTabPage( document.getElementById( ""classoption-page"" ) );"
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
			.Write "                  <b>日期格式：年-月-日 时：分：秒"
			.Write "               </dd>"
	    End If
		If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='rank']/showonform").text="1" Then
			.Write "              <dd><div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='rank']/title").text & ":</div>"
			.Write "             <select name='rank'>"
			If Rank = "★" Then .Write "<option  selected>★</option>" Else .Write "<option>★</option>"
			If Rank = "★★" Then .Write "<option  selected>★★</option>" Else .Write "<option>★★</option>"
			If Rank = "★★★" Or Action = "Add" Then .Write "<option  selected>★★★</option>" Else .Write "<option>★★★</option>"
			If Rank = "★★★★" Then .Write "<option  selected>★★★★</option>" Else .Write "<option>★★★★</option>"
			If Rank = "★★★★★" Then .Write "<option  selected>★★★★★</option>" Else .Write "<option>★★★★★</option>"
			.Write "</select>&nbsp;请为" & KS.C_S(ChannelID,3) & "评定阅读等级"
			.Write "               </dd>"
	   End If
	   If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='hits']/showonform").text="1" Then
			.Write "              <dd><div>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='hits']/title").text & ":</div>本日：<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='6' style='text-align:center' class='textbox'> 本周：<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='6' style='text-align:center' class='textbox'> 本月：<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='6' style='text-align:center' class='textbox'> 总计：<input name='Hits' type='text' id='Hits' value='" & Hits & "' size='6' style='text-align:center' class='textbox'>&nbsp;" 
			.Write "得票数：<input type='text' name='score' size='6' value='" & score & "'>票"
			.Write "              </dd>"
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
		.Write "              <dd><div>审核状态:</div>"
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
			
			.Write "                  </dd>"	
			.Write "</dl>"
			.Write "</div>"
  End If
  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/showonform").text="1" Then
	     KSCls.LoadSeoOption ChannelID,FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/title").text,SEOTitle,SEOKeyWord,SEODescript
  End If
      
  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonform").text="1" Then
	       KSCls.LoadChargeOption ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
  End If
  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='relativeoption']/showonform").text="1" Then		 
	       KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))
  End If		   
			 .Write "</form>"
			 .Write " </div>"
			%>
			 <script type="text/javascript">
			 $(document).ready(function(){
				$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
				$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
			 <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
			  $('#KeyLinkByTitle').click(function(){GetKeyTags();});
			 <%End If%>
			 IniPicUrl();
			 
			 //只有一个栏目时，为选其选中
			  if ($("#tid option").length<=2){
			     $("#tid option").each(function(i){
				    if (i==$("#tid option").length-1) $(this).attr("selected",true);
				});
			  }
			 
			});
			function GetKeyTags()
			{
			  var text=escape($('input[name=title]').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){$('#KeyWords').val(unescape(data)).attr("disabled",false);});
			  }else{alert('对不起,请先输入内容!');}
			}
         function addMap(){
		 top.openWin('电子地图标注','../../plus/baidumap.asp?obj=parent.frames["MainFrame"].document.getElementById("MapMark")&MapMark='+escape($("#MapMark").val()),false,860,430);
		  }	
		  function IniPicUrl()
			{
			  var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
			  var PicUrlArr=null;
			  if (PicUrls!='')
			   { 
				PicUrlArr=PicUrls.split('|||');
			    for ( var i=1 ;i<PicUrlArr.length+1;i++){ 
			      addImage(PicUrlArr[i-1].split('|')[1],PicUrlArr[i-1].split('|')[2],PicUrlArr[i-1].split('|')[0],i);
			    }
			   }
			}
			function SelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",true);
			  });
			}
			function UnSelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",false);
			  });
			}
			function GetFileNameArea(f)
			{
			  $('#filearea').toggle(f);
			}
			function GetTemplateArea(f)
			{
			   $('#templatearea').toggle(f);
			}
			function SubmitFun()
			{ 	
			    if ($('input[name=title]').val()=="")
				  {
					top.$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>名称！",function(){
					$('input[name=title]').focus();});
					return;
				  }
			   if ($("#tid option:selected").val()=='0')
			   {
			       top.$.dialog.alert('请选择所属<%=KS.GetClassName(ChannelID)%>!');
				   return false;
			   }
			 	if ($('input[name=PhotoUrl]').val()==''<%if action="Add" Then response.write " && $('#autothumb').attr('checked')==false"%>)
				{
					top.$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>缩略图！",function(){
					$('input[name=PhotoUrl]').focus();
					});
					return;
				}
				<%
			  Call LFCls.ShowDiyFieldCheck(FieldXML,1)
			     %>
			  var picSrcs='';
			  var src='';
			  $("#thumbnails").find(".pics").each(function(){
			     src=$(this).next().val().replace('|||','').replace('|','')+'|'+$(this).val()
			     if(picSrcs==''){
			      picSrcs=src;
				 }else{
			      picSrcs+='|||'+src;
				 }
			  });
			   $('#PicUrls').val(picSrcs);
				var addmode;
				for (var i=0;i<document.myform.addmode.length;i++){
				 var KM = document.myform.addmode[i];
				if (KM.checked==true)	   
					addmode = KM.value
				}
		
				if (addmode==0 && $('input[name=PicUrls]').val()=='')
				{
				  top.$.dialog.alert('请上传图片集!',function(){
				  $('input[name=imgurl1]').focus();});
				  return false;
				}
				  $('form[name=myform]').submit();
				  $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
			}
			
			
				
		</script>
    
<%
			 .Write "</body>"
			 .Write "</html>"
			 End With
			 if rs.state=1 then rs.close:Set rs=nothing
		End Sub
		
		'保存
		Sub PictureSave()
		   Dim MoreStart,MoreEnd,MorePicUrl,MorePhotoUrl,I,SelectInfoList,HasInRelativeID,Relateda,Related_s,ii
		  With Response
			Title = Request.Form("Title")
			PictureContent= KS.FilterIllegalChar(Request.Form("Content")) : If KS.IsNul(PictureContent) Then PictureContent=" "
			Hits        = KS.ChkClng(KS.G("Hits"))
			HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
			HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
			HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))
			
			PhotoUrl     = KS.G("PhotoUrl")
			If KS.G("AddMode")="0" Then
			   PicUrls     = KS.G("PicUrls")
			Else
			   MoreStart=KS.ChkClng(KS.G("MoreStart"))
			   MoreEnd=KS.ChkClng(KS.G("MoreEnd"))
			   If MoreStart>MoreEnd Then .Write "<script>alert('批量添加的结束ID必须大小开始ID!');history.back();</script>":.end
			   MorePicUrl=KS.G("MorePicUrl")
			   For I=MoreStart to MoreEnd
			    If PicUrls="" Then
				 PicUrls=Replace(MorePicUrl,"#",I)
				Else
				 PicUrls=PicUrls & "|||" & Replace(MorePicUrl,"#",I)
				End If
			   Next
			End If
			
			Recommend   = KS.ChkClng(KS.G("Recommend"))
			Rolls       = KS.ChkClng(KS.G("Rolls"))
			Strip       = KS.ChkClng(KS.G("Strip"))
			Popular     = KS.ChkClng(KS.G("Popular"))
			Comment     = KS.ChkClng(KS.G("Comment"))
			IsTop       = KS.ChkClng(KS.G("IsTop"))
			Slide       = KS.ChkClng(KS.G("Slide"))
			Makehtml    = KS.ChkClng(KS.G("Makehtml"))
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
			
			SpecialID   = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
			SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
			
			If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then
						Verific=KS.ChkClng(KS.G("Verific"))
			Else
						Verific = 1
			End if
			
			KeyWords    = KS.G("KeyWords")
			Author      = KS.G("Author")
			Origin      = KS.G("Origin")
			AddDate     = KS.G("AddDate")
			If Not IsDate(AddDate) Then AddDate=Now
			Rank        = Trim(KS.G("Rank"))
			ShowStyle   = KS.ChkClng(KS.G("ShowStyle"))
			PageNum     = KS.ChkClng(KS.G("PageNum"))
			Province    = KS.G("Province")
			City        = KS.G("City")
			County      = KS.G("County")
			Inputer     = KS.G("Inputer")
			
			'SEO优化选项
			SEOTitle    = KS.G("SEOTitle")
			SEOKeyWord  = KS.G("SEOKeyWord")
			SEODescript = KS.G("SEODescript")
				
			'收费选项
			ReadPoint   = KS.ChkClng(KS.G("ReadPoint"))
			ChargeType  = KS.ChkClng(KS.G("ChargeType"))
			PitchTime   = KS.ChkClng(KS.G("PitchTime"))
			ReadTimes   = KS.ChkClng(KS.G("ReadTimes"))
			InfoPurview = KS.ChkClng(KS.G("InfoPurview"))
			arrGroupID  = KS.G("GroupID")
			DividePercent=KS.G("DividePercent"):IF Not IsNumeric(DividePercent) Then DividePercent=0
				
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
			 
			If Title = "" Then .Write ("<script>alert('图片名称不能为空!');history.back(-1);</script>")
			If PhotoUrl = "" And KS.ChkClng(KS.S("autothumb"))=0 Then .Write ("<script>alert('图片缩略图不能为空!');history.back(-1);</script>")
			
			
			
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Tid = "" Then ErrMsg = ErrMsg & "[图片类别]必选! \n"
			If Title = "" Then ErrMsg = ErrMsg & "[图片标题]不能为空! \n"
			If Title <> "" And Tid <> "" And Action = "Add" Then
			  SqlStr = "select top 1 * from " & KS.C_S(ChannelID,2) & " where Title=" & KS.WithKorean() &"'" & Replace(Title,"'","''") & "' And Tid='" & Tid & "'"
			   RS.Open SqlStr, conn, 1, 1
				If Not RS.EOF Then
				 ErrMsg = ErrMsg & "该类别已存在此篇图片! \n"
			   End If
			   RS.Close
			End If
			If ErrMsg <> "" Then
			   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
			   .End
			Else
			      If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)
				  
			      If KS.ChkClng(KS.G("BeyondSavePic"))=1 Then
				  	SaveFilePath = KS.GetUpFilesDir & "/"
					KS.CreateListFolder (SaveFilePath)
				   Dim sPicUrlArr:sPicUrlArr=Split(PicUrls,"|||")
				   Dim sTemp,Url1,thumburl,ThumbFileName
				   PicUrls=""
				   For I=0 To Ubound(sPicUrlArr)
				     If Left(Lcase(Split(sPicUrlArr(i),"|")(1)),4)="http" and instr(Lcase(Split(sPicUrlArr(i),"|")(1)),lcase(ks.setting(2)))=0 and instr(Lcase(Split(sPicUrlArr(i),"|")(1)),"kesion.com")=0 Then
					    Url1=SaveFilePath & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & i &".jpg"
					    Call KS.SaveBeyondFile(Url1, Split(sPicUrlArr(i),"|")(1))
					    thumburl=replace(url1,ks.setting(2),"")
					    ThumbFileName=split(thumburl,".")(0)&"_S."&split(thumburl,".")(1)
						if instr(Lcase(thumburl),"http://")=0 Then
							Dim T:Set T=New Thumb
							Dim CreateTF:CreateTF=T.CreateThumbs(thumburl,ThumbFileName)
							if CreateTF=false Then
								ThumbFileName=url1
							end if
							Set T=Nothing
						end if
					  sTemp=Split(sPicUrlArr(i),"|")(0) & "|" & Url1 &"|" &ThumbFileName
					 Else
					  sTemp=sPicUrlArr(I)
					 End If
					 If I=0 Then
					   PicUrls=sTemp
					 Else
					   PicUrls=PicUrls & "|||" & sTemp
					 End If
				   Next
				   PhotoUrl= KS.ReplaceBeyondUrl(PhotoUrl, SaveFilePath)
				  End If
				  
				  If KS.ChkClng(KS.S("autothumb"))=1 And KS.IsNul(PhotoUrl) Then  PhotoUrl=Split(Split(PicUrls,"|||")(0),"|")(2)
				  
				  
				  If Action = "Add" Then
					for ii=0 to Ubound(Relateda)
						Set RS = Server.CreateObject("ADODB.RecordSet")
						SqlStr = "select top 1 * from " & KS.C_S(ChannelID,2) &" where 1=0"
						RS.Open SqlStr, conn, 1, 3
						RS.AddNew
						RS("Title")         = Title
						RS("PhotoUrl")      = PhotoUrl
						RS("PictureContent")= PictureContent
						RS("PicUrls")       = PicUrls
						RS("PicNum")        = Ubound(split(PicUrls,"|||"))+1
						RS("Recommend")     = Recommend
						RS("Rolls")         = Rolls
						RS("Strip")         = Strip
						RS("Popular")       = Popular
						RS("Verific")       = Verific
						RS("Comment")       = Comment
						RS("IsTop")         = IsTop
						if ii=0 then
							RS("Tid")            = Tid
						else
							RS("Tid")            =RTrim(Trim(Relateda(ii)))
							Tid = RTrim(Trim(Relateda(ii))) 
						end if
						RS("OTid")          = KS.G("OTid")
						RS("oID")           = KS.ChkClng(KS.S("oid"))
						RS("Province")      = Province
						RS("City")          = City
						RS("County")        = County
						RS("KeyWords")      = KeyWords
						RS("Author")        = Author
						RS("Origin")        = Origin
						RS("AddDate")       = AddDate
						RS("ModifyDate")    = AddDate 
						RS("Rank")          = Rank
						RS("Slide")         = Slide
						if ii=0 then
						 RS("TemplateID")     = TemplateID
						Else
						 RS("TemplateID")     = KS.C_C(TID,5)
						End If
						RS("WapTemplateID") = WapTemplateID
						RS("Hits")          = Hits
						RS("HitsByDay")     = HitsByDay
						RS("HitsByWeek")    = HitsByWeek
						RS("HitsByMonth")   = HitsByMonth
						RS("Fname")         = Fname
						RS("Inputer")       = KS.C("AdminName")
						if KS.IsNul(KS.Setting(189))  then
							 RS("RefreshTF")      = Makehtml
						else
							 RS("RefreshTF")      = 0
						end if
						RS("Score")         = KS.ChkClng(KS.G("Score"))
						RS("DelTF")         = 0
						RS("PostTable")     = LFCls.GetCommentTable()
						RS("CmtNum")        = 0
						RS("ShowStyle")     = ShowStyle
						RS("PageNum")       = PageNum
						RS("ReadPoint")     = ReadPoint
						RS("ChargeType")    = ChargeType
						RS("PitchTime")     = PitchTime
						RS("ReadTimes")     = ReadTimes
						RS("InfoPurview")   = InfoPurview
						RS("arrGroupID")    = arrGroupID
						RS("DividePercent") = DividePercent
						RS("SEOTitle")      = SEOTitle
						RS("SEOKeyWord")    = SEOKeyWord
						RS("SEODescript")   = SEODescript
						RS("OrderID")        = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & Tid &"'")(0))+1
						If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonform").text="1" Then RS("MapMarker")=KS.G("MapMark")
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
						
					   '写入Session,添加下一篇图片调用
					   Session("KeyWords") = KeyWords
					   Session("Author")   = Author
					   Session("Origin")   = Origin
					   RS.MoveLast
					   If Left(Ucase(Fname),2)="ID" Then
						   RS("Fname") = RS("ID") & FnameType
						   RS.Update
						End If
					  If ii=0 then
						For I=0 To Ubound(SpecialID)
							Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						Next
					  End If
						
						Call KSCls.UpdateRelative(ChannelID,RS("ID"),SelectInfoList,0)
						Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,Tid,PictureContent,KeyWords,PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
						'关联上传文件
						 Call KS.FileAssociation(ChannelID,RS("ID"),PicUrls & PhotoUrl & PictureContent,0)
	
						Call RefreshHtml(1)
				  Next
				 
				  RS.Close:Set RS = Nothing
					
				ElseIf Action = "Edit" Or Action="Verify"  Then
				   If Action="Verify" Then 
					 Call KS.ReplaceUserFile(PictureContent,ChannelID)
					 Call KS.ReplaceUserFile(PhotoUrl,ChannelID)
					 Call KS.ReplaceUserFile(PicUrls,ChannelID)
					End If
					
				PicID = KS.ChkCLng(Request("PicID"))
				    dim RelatedArray,n:n=0
					if KS.ChkClng(KS.G("EditNewtb"))=1 then
						if KS.ChkClng(KS.G("RelatedID"))=0 or KS.ChkClng(KS.G("RelatedID"))=-11 then
							RelatedArray=KSCls.GetRelatedArray( KS.C_S(ChannelID,2), PicID ,11)'同步文章
						else
							RelatedArray=KSCls.GetRelatedArray( KS.C_S(ChannelID,2), KS.ChkClng(KS.G("RelatedID")) ,22)'同步文章
						end if 
					else
						RelatedArray=Array(PicID)	
					end if
					
					for ii=0 to UBound(RelatedArray)
					 Set RS=SERVER.CreateObject("ADODB.RECORDSET")
				     SqlStr = "SELECT top 1 * FROM " & KS.C_S(ChannelID,2) & " Where ID=" & RTrim(Trim(RelatedArray(ii))) & ""
					 RS.Open SqlStr, conn, 1, 3
						If RS.EOF And RS.BOF Then
						 .Write ("<script>alert('参数传递出错!');history.back(-1);</script>")
						 .End
						End If
						RS("Title")          = Title
						RS("PhotoUrl")       = PhotoUrl
						RS("PictureContent") = PictureContent
						RS("PicUrls")        = PicUrls
						RS("PicNum")         = Ubound(split(PicUrls,"|||"))+1
						RS("Recommend")      = Recommend
						RS("Rolls")          = Rolls
						RS("Strip")          = Strip
						RS("Popular")        = Popular
						RS("Comment")        = Comment
						RS("IsTop")          = IsTop
						if PicID=KS.ChkClng(RTrim(Trim(RelatedArray(ii)))) then
							RS("Tid")       = KS.G("Tid")
							RS("oTid")      = KS.G("oTid")
							RS("oID")       = KS.ChkClng(KS.S("oid"))
						end if
						RS("Province")       = Province
						RS("City")           = City
						RS("County")         = County
						RS("KeyWords")       = KeyWords
						RS("Author")         = Author
						RS("Origin")         = Origin
						RS("AddDate")        = AddDate
						RS("ModifyDate")     = Now
						RS("Rank")           = Rank
						RS("ShowStyle")      = ShowStyle
						RS("PageNum")        = PageNum
						RS("Slide")          = Slide
						RS("TemplateID")     = TemplateID
						RS("WapTemplateID")  = WapTemplateID
						If Makehtml = 1 Then
						 RS("RefreshTF")     = 1
						End If
						RS("Hits")           = Hits
						RS("HitsByDay")      = HitsByDay
						RS("HitsByWeek")     = HitsByWeek
						RS("HitsByMonth")    = HitsByMonth
						RS("Score")          = KS.ChkClng(KS.G("Score"))
						RS("ReadPoint")      = ReadPoint
						RS("ChargeType")     = ChargeType
						RS("PitchTime")      = PitchTime
						RS("ReadTimes")      = ReadTimes
						RS("InfoPurview")    = InfoPurview
						RS("arrGroupID")     = arrGroupID
						RS("DividePercent")  = DividePercent
						If Action="Verify" Then
						  Inputer            = RS("Inputer")
						End If
						If Verific<>100 Then
						 RS("Verific") = Verific
						End If
						RS("SEOTitle")      = SEOTitle
						RS("SEOKeyWord")    = SEOKeyWord
						RS("SEODescript")   = SEODescript
	
						If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonform").text="1" Then RS("MapMarker")=KS.G("MapMark")
						Call KSCls.AddDiyFieldValue(RS,FieldXml)
					RS.Update
                   RS.MoveLast
			       If TID<>Request.Form("OldClassID") Then
					     Call KSCls.DelInfoFile(ChannelID,Request.Form("OldClassID"), Split(RS("PicUrls"), "|||"),RS("Fname"),RS("ID"),RS("AddDate"))
				   End If
				   if ii=0 then
						Conn.Execute("Delete From KS_SpecialR Where InfoID=" & RS("ID") & " and channelid=" & ChannelID)
						For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						Next
					End If
					Call KSCls.UpdateRelative(ChannelID,PicID,SelectInfoList,1)
					Call LFCls.UpdateItemInfo(ChannelID,PicID,Title,Tid,PictureContent,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
	 				'关联上传文件
					Call KS.FileAssociation(ChannelID,PicID,PicUrls & PhotoUrl & PictureContent,1)
				    Call RefreshHtml(2)
		          Next
				  RS.Close:Set RS = Nothing
					IF (Action="Verify" or (Not KS.IsNul(KS.S("oldverific")) AND KS.S("oldverific")="0" And Verific=1)) And Inputer<>KS.C("AdminName")  Then     '如果是审核投稿图片，对用户，进行加积分等，并返回签收图片管理
							  '对用户进行增值，及发送通知操作
							  IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,Title,PicID)
							 .Write ("<script> parent.frames['MainFrame'].focus();alert('恭喜，" & KS.C_S(ChannelID,3) &"成功审核!');location.href='../System/KS.ItemInfo.asp?ShowType=1&ChannelID=" & ChannelID & "&Page=" & Page & "&ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Disabled&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) &" >> <font color=red>审核会员" & KS.C_S(ChannelID,3)) &"</font>';</script>") 
							 
				    End If
					If KeyWord <>"" Then
						 .Write ("<script> parent.frames['MainFrame'].focus();setTimeout(function(){alert('" & KS.C_S(ChannelID,3) &"修改成功!');location.href='../System/KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=PictureSearch&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) &" >> <font color=red>搜索结果</font>") & "';},2500);</script>")
					End If
				End If
			End If
		 End With		
		End Sub
		
			Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="添加":EditStr="修改" & KS.C_S(ChannelID,3):AddStr="继续添加" & KS.C_S(ChannelID,3)
				Else
				  TempStr="修改":EditStr="继续修改" & KS.C_S(ChannelID,3):AddStr="添加" & KS.C_S(ChannelID,3)
				End If
			    With Response
				     .Write "<!DOCTYPE html><html><head><link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<meta http-equiv=Content-Type content=""text/html; charset=utf-8"">"
					 .Write "<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script></head>"
					 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
					 .Write " <div class='pageCont2 mt20'><div class='tabTitle'>系统操作提示信息</div><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
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
					   ' .Write "<div style=""margin-left:140;color:blue;height:25px""><li>由于 <a href=""" & KS.GetDomain & """ target=""_blank""><font color=red>网站首页</font></a> 没有启用生成HTML的功能或发布选项没有开启，所以没有生成!</li></div>"
					   Else
					     .Write "<div align=center><iframe src=""../Include/RefreshIndex.asp?RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div></div>"
					 End If
					  .Write   "</td></tr></table></td></tr>"
					  .Write "	  <tr class='tdbg'>"
					  .Write "		<td height=""25"" colspan=""2"" style=""text-align:right"">【<a href=""#"" onclick=""location.href='KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" & RS("ID") & "';""><strong>" & EditStr &"</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.Picture.asp?ChannelID=" & ChannelID & "&Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr="&server.URLEncode("添加" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='../System/KS.ItemInfo.asp?ID=" & Tid & "&ChannelID=" & ChannelID & "&Page=" & Page&"&keyword=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>" & KS.C_S(ChannelID,3) & "管理</strong></a>】&nbsp;【<a href=""" & KS.GetDomain &"Item/Show.asp?M=" & ChannelID & "&D=" & RS("ID") & """ target=""_blank""><strong>预览" & KS.C_S(ChannelID,3) & "内容</strong></a>】</td>"
					  .Write "	  </tr>"
					  .Write "	</table></div></body></html>"				
			End With
		End Sub

End Class
%> 

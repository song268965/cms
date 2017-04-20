<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.FunctionCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"

'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New InterView_Main
KSCls.Kesion()
Set KSCls = Nothing

Class InterView_Main
        Private KS,KSCls,Action
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
			
			Action=KS.G("Action")
			If Action="Add" Then
			   If Not KS.ReturnPowerResult(0, "InterView0001") Then
				  Call KS.ReturnErr(1, "")
				  exit sub
				End If
			Else
				If Not KS.ReturnPowerResult(0, "InterView0000") Then
				  Call KS.ReturnErr(1, "")
				  exit sub
				End If
		    End If
			
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With KS
			Select Case Action
			 Case "Add","Edit" Call InterViewAddOrEdit()
			 Case "Save" Call InterViewSave()
			 Case "Del" Call InterViewDel()
			 Case "interviewpic" Call InterViewPic()
			 case "InterViewPicSave" call InterViewPicSave()
			 Case "DelInterViewPhoto" Call DelInterViewPhoto()
			 Case Else Call MainList()
			End Select
		  End With
	    End Sub
		
		'删除现场图片
		Sub DelInterViewPhoto()
			Dim UserID,i,p,picarr,pic:pic=KS.S("Pic")
			 Dim Flag:Flag=KS.ChkClng(Request("flag"))
			 Dim PicID:PicID=KS.ChkClng(Request("picid"))
			 If Not KS.IsNul(Pic) Then
				PicArr=Split(pic,"|")
				
				for i=0 to ubound(PicArr)-1
				  p=PicArr(i)
				  If Not KS.IsNul(p) Then 
					 p=replace(p,KS.Setting(2),"")
					 Call KS.DeleteFile(p)
				  End If
				next
				if picid<>0 then conn.execute("delete from KS_InterViewPic where id=" & picid)
			 End If
		End Sub
		
		Sub InterViewPic()
		Dim ID:ID=KS.ChkCLng(request("id"))
		Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open "select top 1 * From KS_InterView Where ID=" & ID,conn,1,1
		If RS.EOf And RS.Bof Then
		 RS.Close:Set RS=Nothing
		 KS.AlertHintScript "对不起，找不到访谈主题！"
		End If
		Dim Title:Title=RS("Title")
		RS.Close:Set RS=Nothing
		%>
		<!DOCTYPE html>
		<html>
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title>站点访谈</title>
		<link href="../../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<script language="javascript" src="../../../KS_Inc/jquery.js"></script>
		<script language="javascript" src="../../../KS_Inc/common.js"></script>
		<script>
		 function checkform(){
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
            if ($('#PicUrls').val()==''){
			 alert('请先上传现场图片！');
			 return false;
			}
			return true;
		 }
		</script>
		<body>
		<div class="pageCont2 mt20">
		<div class="tabTitle">管理访谈[<%=Title%>]现场图片</div>
		 <form name=InterViewForm method=post action="?Action=InterViewPicSave">  
		  <input type="hidden" name="ID" value="<%=id%>">  
		<table width="100%" border="0" cellpadding="1" cellspacing="1" class='ctable'> 
		   <tr>  
		   <td height="325" align='right' width='85' class='clefttitle'><strong>图片内容:</strong></td>  
		   <td valign="top">
		     
			 <style type="text/css">
				#thumbnails{background:url(../../../plus/swfupload/images/albviewbg.gif) no-repeat;min-height:200px;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
				#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:162px;height:155px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
				#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px}
				</style>
				<link href="../../../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
				<script type="text/javascript" src="../../../plus/swfupload/swfupload/swfupload.js"></script>
				<script type="text/javascript" src="../../../plus/swfupload/js/handlers.js"></script>
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
						function DelUpFiles(pid,picid)
						{  var p=$('#pic'+pid).val();
						   if (p!==''){
							$.ajax({
							  url: "KS.InterView.asp",
							  cache: false,
							  data: "action=DelInterViewPhoto&pic="+p+"&flag=1&picid="+picid,
							  success: function(r){
							  }
							  });
						   }
						   $("#thumbshow"+pid).remove();	
						}	
						
						function addImage(bigsrc,smallsrc,text,picid) {
							var newImgDiv = document.createElement("div");
							var delstr = '';
							delstr = '<a href="javascript:DelUpFiles('+pid+','+picid+')" style="color:#ff6600">[删除]</a>';
							newImgDiv.className = 'thumbshow';
							newImgDiv.id = 'thumbshow'+pid;
							document.getElementById("thumbnails").appendChild(newImgDiv);
							newImgDiv.innerHTML = '<a href="'+bigsrc+'" target="_blank"><span id="show'+pid+'"></span></a>';
							newImgDiv.innerHTML += '<div style="margin-top:10px;text-align:left">'+delstr+' <b>注释：</b><input type="hidden" class="pics" id="pic'+pid+'" value="'+bigsrc+'|'+smallsrc+'|'+picid+'"/><input type="text" name="picinfo'+pid+'" value="'+text+'" style="width:155px;" /></div>';
						
							var newImg = document.createElement("img");
							newImg.style.margin = "5px";
						
							document.getElementById("show"+pid).appendChild(newImg);
							if (newImg.filters) {
								try {
									newImg.filters.item("DXImageTransform.Microsoft.Alpha").opacity = 0;
								} catch (e) {
									newImg.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=' + 0 + ')';
								}
							} else {
								newImg.style.opacity = 0;
							}
						
							newImg.onload = function () {
								fadeIn(newImg, 0);
							};
							newImg.src = smallsrc;
							pid++;
							
						}
					
						window.onload = function () {
							swfu = new SWFUpload({
								// Backend Settings
								upload_url: "../../include/swfupload.asp",
								post_params: {UPType:"pic","AdminID" : "<%=KS.C("AdminID") %>","AdminPass":"<%=KS.C("AdminPass")%>",AddWaterFlag:"1","BasicType":2,"ChannelID":2,"AutoRename":4},
				
								// File Upload Settings
								file_size_limit : 1024*2,	// 2MB
								file_types : "*.jpg; *.gif; *.png",
								file_types_description : "支持.JPG.gif.png格式的图片,可以多选",
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
								flash_url : "../../../plus/swfupload/swfupload/swfupload.swf",
								flash9_url : "../../../plus/swfupload/swfupload/swfupload_FP9.swf",
				
								custom_settings : {
									upload_target : "divFileProgressContainer"
								},
								
								// Debug Settings
								debug: false
							});
						};
					</script>
					
			<table>
			 <tr>
			  <td><div class="button"><span id="spanButtonPlaceholder"></span></div>
			  
			  <label><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>图片添加水印</label>
			  </td>
			
			 </tr>
			</table>
			
			<div id="divFileProgressContainer"></div>
			<div id="thumbnails"></div>
				<input type='hidden' name='PicUrls' id='PicUrls'>
		   
		   </td>
		   </tr>
		   <tr>
		   <td colspan="2" style="text-align:center" height="40">
		    <input type="submit" value=" 保存图片 " onClick="return(checkform())" class="button"/>
		   </td>
		   </tr>
		</table>
		</form>
		</div>
		<%
		   Response.Write "<script type=""text/javascript"">" & vbcrlf
		   Dim RSS:Set RSS=Conn.Execute("Select * From KS_InterViewPic Where InterviewID=" & ID &" order by id")
		   Do While Not RSS.Eof
		    Response.Write "addImage('" & RSS("PhotoUrl") & "','" & RSS("PhotoUrl") & "','" & RSS("content") & "'," & rss("id") &");" &vbcrlf
		   RSS.MoveNext
		   Loop
		   Response.Write "</script>"
		   RSS.Close :Set RSS=Nothing
		End Sub
		
		Sub InterViewPicSave()
		  	Dim ID:ID=KS.ChkCLng(request("id"))
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "select top 1 * From KS_InterView Where ID=" & ID,conn,1,1
			If RS.EOf And RS.Bof Then
			 RS.Close:Set RS=Nothing
			 KS.AlertHintScript "对不起，找不到访谈主题！"
			End If
			Dim Title:Title=RS("Title")
			RS.Close:Set RS=Nothing
		
		
		   Dim sTemp,Url1,thumburl,ThumbFileName,SaveFilePath,PicUrls
			  PicUrls=Request.Form("PicUrls")
				  SaveFilePath = KS.GetUpFilesDir & "/"
				  KS.CreateListFolder (SaveFilePath)
				  Dim sPicUrlArr:sPicUrlArr=Split(PicUrls,"|||")
				   For I=0 To Ubound(sPicUrlArr)
				      Call AddProImages(ID, Split(sPicUrlArr(i)&"|||","|")(1),Split(sPicUrlArr(i)&"|||","|")(0),Split(sPicUrlArr(i)&"|||","|")(3))
				   Next
         KS.Die "<script>alert('恭喜，现场图片上传成功!');location.href='KS.Interview.asp?action=interviewpic&id=" & id&"';</script>"
		End Sub
		
		 Sub AddProImages(InterViewID,PhotoUrl,Content,PicId)
	    Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open "select top 1 * From KS_InterViewPic Where ID=" & KS.ChkClng(PicId),conn,1,3
		If RS.Eof AND RS.Bof Then
		   RS.AddNew
		   RS("AddDate")=now
		End If
		   RS("InterViewID")=InterViewID
		   RS("PhotoUrl")=PhotoUrl
		   RS("content")=Content
		RS.Update
		RS.Close
		Set RS=Nothing
	      '关联上传文件
		 Call KS.FileAssociation(111115,InterViewID,PhotoUrl  ,0)
	  End Sub
		
		Sub MainList()
			With KS
			.echo "<!DOCTYPE html><html>"
			.echo "<head>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<title>站点访谈</title>"
			.echo "<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script language=""JavaScript"">" & vbCrLf
			.echo "var Page='" & CurrentPage & "';" & vbCrLf
			.echo "</script>" & vbCrLf
			.echo "<script language=""JavaScript"" src=""../../../KS_Inc/common.js""></script>"
			.echo "<script language=""JavaScript"" src=""../../../KS_Inc/jquery.js""></script>"
			
			%>
			<script language="JavaScript">
			$(document).ready(function(){
				
		      $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		     })

			function InterViewAdd()
			{
				location.href='KS.InterView.asp?Action=Add';
				window.$(parent.document).find('#BottomFrame')[0].src='Post.asp?OpStr=访谈管理中心 >> <font color=red>添加新访谈</font>&ButtonSymbol=GO';
			}
			function EditInterView(id)
			{ 
			    if (id=='') id=get_Ids(document.myform);
				if (id==''){
				 alert('请选择要编辑的访谈!');
				}else if(id.indexOf(',')==-1){
				location="KS.InterView.asp?Action=Edit&Page="+Page+"&Flag=Edit&InterViewID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='Post.asp?OpStr=访谈管理中心 >> <font color=red>编辑访谈</font>&ButtonSymbol=GoSave';
				}else{
				alert('一次只能编辑一条访谈!');
				}
			}
			function DelInterView(id)
			{
			 if (id=='') id=get_Ids(document.myform);
			 if (id==''){
			   alert('请先选择要删除的访谈!')
			 }else if (confirm('真的要删除选中的访谈吗?')){
				 location="KS.InterView.asp?Action=Del&Page="+Page+"&id="+id;
				}
			 }
			</script>
			<%
			.echo "</head>"
			.echo "<body topmargin=""0"" leftmargin=""0"">"
			.echo "<ul id='menu_top'>"
			.echo "<li class='parent' onclick=""InterViewAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加访谈</span></li>"
			.echo "<li class='parent' onclick=""EditInterView('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon write'></i>编辑访谈</span></li>"
			.echo "<li class='parent' onclick=""DelInterView('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon delete'></i>删除访谈</span></li>"
			.echo "</ul>"
			.echo "<div class='pageCont2'>"
			.echo "<div class='tabTitle'>在线访谈管理</div>"
			.echo "<form name=""myform"" id=""myform"" action=""KS.InterView.asp"" method=""post"">"
			.echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.echo "<input type=""hidden"" value=""Del"" name=""Action"" ID=""Action"">"
			.echo "<input type=""hidden"" value="""& CurrentPage & """ name=""Page"" ID=""Page"">"
			.echo "  <tr  align=""center"">"			
			.echo "          <td height=""25"" class=""sort"">选择</div></td>"
			.echo "          <td  height=""25"" class=""sort"">访谈标题</div></td>"
			.echo "          <td class=""sort""><div align=""center"">访谈时间</div></td>"
			.echo "          <td class=""sort""><div align=""center"">主持人</div></td>"
			.echo "          <td align=""center"" class=""sort"">嘉宾</td>"
			.echo "          <td align=""center"" class=""sort"">状态</td>"
			.echo "          <td class=""sort"">管理操作</td>"
			.echo "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_InterView order by id desc"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					   .echo "<tr><td  class='splittd' colspan=10 style='text-align:center'>还没有添加访谈!</td></tr>"
					 Else
						       totalPut = RSObj.RecordCount
			
								If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								 Dim InterViewXMl:Set InterViewXml=KS.ArrayToXml(RSObj.GetRows(MaxPerPage),RSObj,"row","root")
							     Call showContent(InterViewXml)
								 Set InterViewXMl=Nothing

				End If
				RSObj.Close
				Set RSObj=Nothing
			.echo "    </td>"
			.echo "  </tr>"
             .echo " <tr>"
			 .echo " <td colspan='2'><div class='operatingBox'><b>选择：</b><a href='javascript:void(0)' onclick='Select(0)'>全选</a> -  <a href='javascript:void(0)' onclick='Select(1)'>反选</a> - <a href='javascript:void(0)' onclick='Select(2)'>不选</a> <input type='submit' class='button' value='删 除' onclick=""return(confirm('确定删除选中的访谈吗?'))""></td></form>"
			 .echo "   <td align=""right"" colspan=8>"
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"			
			.echo "</table>"
			.echo "</div>"
			.echo "</body>"
			.echo "</html>"
			End With
			End Sub
			Sub showContent(InterViewXML)
			  Dim ID,Node
			  With KS
			   If IsObject(InterViewXML) Then
			    For Each Node In InterViewXML.DocumentElement.SelectNodes("row")
				       ID=Node.SelectSingleNode("@id").text
					   .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &ID & "' onclick=""chk_iddiv('" & ID & "')"">")
				       .echo ("<td class='splittd' align=center><input type='hidden' value='" & ID & "' name='LinkID'><input name='id'  onclick=""chk_iddiv('" & ID & "')"" type='checkbox' id='c"& ID & "' value='" & ID & "'></td>")
				  
					   .echo "  <td class='splittd' height='20'><span InterViewID='" & ID & "' ondblclick=""EditInterView(this.InterViewID)""><img src='../../Images/qs.gif' align='absmiddle'>"
					   .echo "    <span style='cursor:default;'><a href='../../../interview/show.asp?id=" & Node.SelectSingleNode("@id").text &"' target='_blank'>" & KS.GotTopic(Node.SelectSingleNode("@title").text, 45) & "</a></span></span> "
					   .echo "  </td>"
					   .echo "  <td class='splittd' align='center'>"
					   .echo Node.SelectSingleNode("@begindate").text & "至" & Node.SelectSingleNode("@enddate").text
				 	   .echo "  <td class='splittd' align='center'>" & Node.SelectSingleNode("@host").text & "</td>"
				 	   .echo "  <td class='splittd' align='center'>" & Node.SelectSingleNode("@guests").text & "</td>"
				 	   .echo "  <td class='splittd' align='center'>"
					    if Node.SelectSingleNode("@locked").text="1" then
						 .echo "<span style='color:red'>锁定</span>"
						else
						 .echo "<span style='color:green'>正常</span>"
						end if
					   .echo "</td>"
					  
					   .echo "  <Td class='splittd' align='center'><a href=""javascript:EditInterView('');"" class='setA'>修改</a>|<a href=""javascript:DelInterView(" & ID &")"" class='setA'>删除</a>|<a href=""?action=interviewpic&id=" & id &""" class='setA'>现场图片</a>|<a href=""../../../interview/login.asp?id=" & id &""" target='_blank' class='setA'>主持人登录</a>|<a href=""../../../interview/main.asp?id=" & id &""" target='_blank' class='setA'>留言审核</a></td>"
					  .echo "</tr>"
					Next
				End If
			 End With
			End Sub
			
			'添加修改访谈
		  Sub InterViewAddOrEdit()
		  		Dim InterViewID, RSObj, SqlStr, Content, Title, PhotoUrl, NewestTF, AddDate,Flag, Page,ChannelID,GuestsIntro,Locked
				Dim BeginDate,EndDate,Host,Guests,MediaUrl,MessageVerifyTF,MessageTF,HostUserID,HostUserPass,MessageLoginTF,TemplateID
				NewestTF = 1
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					InterViewID = KS.G("InterViewID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT top 1 * FROM KS_InterView Where ID=" & InterViewID
					RSObj.Open SqlStr, Conn, 1, 1
					  Title     = RSObj("Title")
					  PhotoUrl  = RSObj("PhotoUrl")
					  AddDate   = RSObj("AddDate")
					  BeginDate = RSObj("BeginDate")
					  Content   = RSObj("Content")
					  EndDate   = RSObj("EndDate")
					  Host      = RSObj("Host")
					  Guests    = RSObj("Guests")
					  MediaUrl  = RSObj("MediaUrl")
					  MessageVerifyTF = RSObj("MessageVerifyTF")
					  MessageTF = RSObj("MessageTF")
					  HostUserID= RSObj("HostUserID")
					  HostUserPass = RSObj("HostUserPass")
					  MessageLoginTF = RSObj("MessageLoginTF")
					  GuestsIntro=RSObj("GuestsIntro")
					  Locked=RSobj("Locked")
					  TemplateID=RSObj("TemplateID")
					RSObj.Close:Set RSObj = Nothing
				Else
				  Flag = "Add":MessageTF=1:MessageVerifyTF=0:MessageLoginTF=0:Locked=0:TemplateID="{@TemplateDir}/访谈系统/访谈内容页.html"
				End If
				With KS
                .echo"<!DOCTYPE html>" & vbcrlf
			    .echo "<html>"& vbcrlf
				.echo "<head>"
				.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				.echo "<title>站点访谈</title>"
				.echo "<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.echo "<script language=""javascript"" src=""../../../KS_Inc/jquery.js""></script>"
				.echo "<script language=""javascript"" src=""../../../KS_Inc/common.js""></script>"
				.echo "<script src=""../../../KS_Inc/DatePicker/WdatePicker.js""></script>"
				.echo EchoUeditorHead()
				.echo "<body>"
				
				.echo "<div class='tabTitle mt20'>"
				If Flag = "Edit" Then
				 .echo "修改访谈"
				Else
				 .echo "添加访谈"
				End If
				.echo "</div>"
				.echo "<div class='pageCont2'>"
				.echo "  <form name='InterViewForm' id='InterViewForm' method=post action=""?Action=Save"">"
				.echo "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.echo "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.echo "   <input type=""hidden"" name=""InterViewID"" value=""" & InterViewID & """>"
				.echo "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"

				.echo "          <tr>"
				.echo "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>访谈标题:</strong></td>"
				.echo "             <td>"
				.echo "              <input name=""Title"" type=""text"" id=""Title"" value=""" & Title & """ class=""textbox"" style=""width:200px""></td>"
				 .echo "</tr>"
				 .echo "<tr>"
				.echo "  <td height=""25"" align='right' width='85' class='clefttitle'><strong>封面图片:</strong></td>"
				.echo "  <td>"
				.echo "<input name=""PhotoUrl"" type=""text"" id=""PhotoUrl""  value="""
				.echo (PhotoUrl)
				.echo """ class=""textbox"" style=""width:200px""> <input type='button' class='button' name='Submit' value='选择地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath="& KS.GetCommonUpFilesDir() & "',550,290,window,document.InterViewForm.PhotoUrl);""></td>"
				.echo "</tr>"
				 .echo "<tr>"
				.echo "  <td height=""25"" align='right' width='85' class='clefttitle'><strong>视频地址:</strong></td>"
				.echo "  <td>"
				.echo "<input name=""MediaUrl"" type=""text"" id=""MediaUrl""  value="""
				.echo (MediaUrl)
				.echo """ class=""textbox"" style=""width:200px""> <input type='button' class='button' name='Submit' value='选择地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath="& KS.GetCommonUpFilesDir() & "',550,290,window,document.InterViewForm.MediaUrl);""> <span class='tips'>如果有录制视频，可以在访谈结束后在此填上录制的视频地址。</span></td>"
				.echo "</tr>"
				 .echo "<tr>"
				.echo "  <td height=""25"" align='right' width='85' class='clefttitle'><strong>绑定模板:</strong></td>"
				.echo "  <td>"
				.echo "<input name=""TemplateID"" type=""text"" id=""TemplateID""  value="""
				.echo TemplateID
				.echo """ class=""textbox"" style=""width:200px""/> " & KSCls.Get_KS_T_C("$('#TemplateID')[0]") &"</td>"
				.echo "</tr>"
				
				.echo "          <tr>"
				.echo "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>访谈时间:</strong></td>"
				.echo "            <td>"
				.echo "              <input name=""BeginDate"" onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" type=""text"" id=""BeginDate"" value="""
				 If Flag <> "Edit" Then
				 .echo (Now)
				 Else
				 .echo (BeginDate)
				 End If
				.echo """ class=""textbox Wdate"" style=""width:200px""> 至 <input onclick=""WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"" class=""textbox Wdate"" name=""EndDate"" type=""text"" id=""EndDate"" value="""
				if Flag<>"Edit" Then
				 .echo (dateadd("h",2,now))
				Else
				  .echo (EndDate)
				End If 
				.echo """/>   <span class='tips'>请填写访谈的具体时间，如2013-9-1 10:00:00 至 2013-9-1 11:00:00           </td>"
				.echo "          </tr>"

				
				.echo "    <tr>"
				.echo "      <td align='right' width='85' class='clefttitle'><strong>访谈介绍:</strong></td>"
				.echo "      <td valign=""top"">"
				
				 .echo EchoEditor("Content",Content,"Basic","96%","150px")
				
				.echo "</td></tr>"
				.echo "    <tr>"
				.echo "      <td align='right' width='85' class='clefttitle'><strong>嘉宾介绍:</strong></td>"
				.echo "      <td valign=""top"">"
				
				 .echo EchoEditor("GuestsIntro",GuestsIntro,"Basic","96%","150px")
				
				

				
				.echo "</td></tr>"
				.echo "<tr>"
				.echo "<td height=""25"" align='right' width='85' class='clefttitle'><strong>主持人:</strong></td><td>"
				.echo " <br/><br/>主持人姓名：<input name=""Host"" type=""text"" id=""Host"" value=""" & Host & """ class=""textbox"" style=""width:200px"">"
				.echo "<br/><br/>主持人登录账号：<input name=""HostUserID"" type=""text"" id=""HostUserID"" value=""" & HostUserID & """ class=""textbox"" style=""width:200px""><br/><br/>"
				.echo "主持人登录密码：<input name=""HostUserPass"" type=""text"" id=""HostUserPass"" value=""" & HostUserPass & """ class=""textbox"" style=""width:200px"">"
				.echo "</td></tr>"
				.echo "<tr>"
				.echo "<td height=""25"" align='right' width='85' class='clefttitle'><strong>访谈嘉宾:</strong></td><td>"
				.echo "  <input name=""Guests"" type=""text"" id=""Guests"" value=""" & Guests & """ class=""textbox"" style=""width:200px"">"
				.echo "<span class='tips'>多个嘉宾，请用英文逗号隔开。</span></td></tr>"
				.echo "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>是否锁定:</strong></td><td>"
				.echo " <input type='radio' name='Locked' value='1'"
				if Locked="1" then .echo " checked"
				.echo "/>是"
				.echo " <input type='radio' name='Locked' value='0'"
				if Locked="0" then .echo " checked"
				.echo "/>否"
				.echo " <span class='tips'>锁定后将不能再留言及发表，只能查看,一般访谈结束后请设置为锁定。</span></td></tr>"
				.echo "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>开启网友留言:</strong></td><td>"
				.echo " <input type='radio' name='MessageTF' value='1'"
				if MessageTF="1" then .echo " checked"
				.echo "/>开启"
				.echo " <input type='radio' name='MessageTF' value='0'"
				if MessageTF="0" then .echo " checked"
				.echo "/>不开启"
				.echo "</td></tr>"
				.echo "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>网友留言需要登录:</strong></td><td>"
				.echo " <input type='radio' name='MessageLoginTF' value='1'"
				if MessageLoginTF="1" then .echo " checked"
				.echo "/>需要"
				.echo " <input type='radio' name='MessageLoginTF' value='0'"
				if MessageLoginTF="0" then .echo " checked"
				.echo "/>不需要"
				.echo "</td></tr>"
				.echo "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>网友留言需要审核:</strong></td><td>"
				.echo " <input type='radio' name='MessageVerifyTF' value='1'"
				if MessageVerifyTF="1" then .echo " checked"
				.echo "/>需要"
				.echo " <input type='radio' name='MessageVerifyTF' value='0'"
				if MessageVerifyTF="0" then .echo " checked"
				.echo "/>不需要"
				.echo "</td></tr>"
				
				
				
				.echo "  </form>"
				.echo "</table>"
				.echo "</div>"
				.echo "</body>"
				.echo "</html>"
				.echo "<script language=""JavaScript"">" & vbCrLf
				.echo "<!--" & vbCrLf
				.echo "function CheckForm()" & vbCrLf
				.echo "{ var form=document.InterViewForm;" & vbCrLf
				.echo "  if (form.Title.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('请输入访谈标题!');" & vbCrLf
				.echo "    form.Title.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "      if (form.BeginDate.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('请输入访谈开始日期!');" & vbCrLf
				.echo "    form.BeginDate.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "      if (form.EndDate.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('请输入访谈结束日期!');" & vbCrLf
				.echo "    form.EndDate.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				
				.echo "  if (" &GetEditorContent("Content") &"==false)" & vbCrLf
				.echo "  {" & vbCrLf
				.echo "    alert('请输入访谈介绍!');" & vbCrLf
				.echo "   " & GetEditorFocus("Content") &"" & vbcrlf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "   if (form.Host.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('请输入访谈主持人!');" & vbCrLf
				.echo "    form.Host.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "   if (form.Guests.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('请输入访谈嘉宾!');" & vbCrLf
				.echo "    form.Guests.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
								
				.echo "   form.submit();"
				.echo "}"
				.echo "//-->"
				.echo "</script>"
			 End With
		  End Sub
		  
		  '保存
		  Sub InterViewSave()
			Dim InterViewID, RSObj, SqlStr, Title, PhotoUrl, BeginDate, Content, Flag, Page, RSCheck,GuestsIntro,Locked
			Dim EndDate,Host,Guests,MediaUrl,MessageVerifyTF,MessageTF,HostUserID,HostUserPass,MessageLoginTF,TemplateID
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			InterViewID = Request("InterViewID")
			Title = Replace(Replace(Request.Form("Title"), """", ""), "'", "")
			PhotoUrl = Replace(Replace(Request.Form("PhotoUrl"), """", ""), "'", "")
			BeginDate = Replace(Replace(Request.Form("BeginDate"), """", ""), "'", "")
			Content = Replace(Request.Form("Content"), "'", "")
			TemplateID=KS.G("TemplateID")
            EndDate = KS.G("EndDate")
			Host    = KS.G("Host")
			Guests  = KS.G("Guests")
			MediaUrl= KS.G("MediaUrl")
			MessageVerifyTF=KS.ChkClng(KS.G("MessageVerifyTF"))
			MessageTF=KS.ChkClng(KS.G("MessageTF"))
			HostUserID=KS.G("HostUserID")
			HostUserPass=KS.G("HostUserPass")
			MessageLoginTF=KS.ChkClng(KS.G("MessageLoginTF"))
			GuestsIntro=KS.G("GuestsIntro")
			Locked=KS.ChkClng(request("locked"))
			If Not IsDate(BeginDate) Then Call KS.AlertHistory("访谈开始日期格式不正确!", -1)
			If Not IsDate(EndDate) Then Call KS.AlertHistory("访谈结束日期格式不正确!", -1)
		    If datediff("s",begindate,enddate)<0 Then Call KS.AlertHistory("访谈结束日期不能早于开始日期!", -1)
			If Title = "" Then Call KS.AlertHistory("访谈标题不能为空!", -1)
			If BeginDate = "" Then Call KS.AlertHistory("访谈开始日期不能为空!", -1)
			If EndDate = "" Then Call KS.AlertHistory("访谈结束日期不能为空!", -1)
			If Content = "" Then Call KS.AlertHistory("访谈内容不能为空!", -1)
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select top 1 ID From KS_InterView Where Title='" & Title & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  KS.Echo ("<script>alert('对不起,访谈标题已存在!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT top 1 * FROM KS_InterView Where (ID is Null)", Conn, 1, 3
				RSObj.AddNew
				  RSObj("Title") = Title
				  RSObj("PhotoUrl") = PhotoUrl
				  RSObj("AddDate") = Now
				  RSObj("BeginDate")=BeginDate
				  RSObj("EndDate") = EndDate
				  RSObj("Content") = Content
				  RSObj("Host")=Host
				  RSObj("Guests")=Guests
				  RSObj("MediaUrl")=MediaUrl
				  RSObj("TemplateID")=TemplateID
				  RSObj("MessageVerifyTF")=MessageVerifyTF
				  RSObj("MessageTF")=MessageTF
				  RSObj("HostUserID")=HostUserID
				  RSObj("HostUserPass")=HostUserPass
				  RSObj("MessageLoginTF")=MessageLoginTF
				  RSObj("GuestsIntro")=GuestsIntro
				  RSObj("Locked")=Locked
				RSObj.Update
				 RSObj.MoveLast
				 Call KS.FileAssociation(10119,RSObj("ID"),RSObj("Content")&RSObj("PhotoUrl")&RSObj("MediaUrl"),0)
				 RSObj.Close
			  End If
			   Set RSObj = Nothing
			   KS.Echo "<script src='../../../ks_inc/jquery.js'></script>"
			   KS.Echo ("<script> if (confirm('访谈添加成功!继续添加吗?')) {location.href='KS.InterView.asp?Action=Add';}else{location.href='KS.InterView.asp';$(parent.document).find('#BottomFrame')[0].src='Post.asp?ButtonSymbol=Disabled&OpStr=在线访谈 >> <font color=red>访谈管理中心</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_InterView Where Title='" & Title & "' And ID<>" & InterViewID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 KS.Echo ("<script>alert('对不起,标题已存在!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT  top 1 * FROM KS_InterView Where ID=" & InterViewID
			   RSObj.Open SqlStr, Conn, 1, 3
				 RSObj("Title") = Title
				  RSObj("PhotoUrl") = PhotoUrl
				  RSObj("BeginDate") = BeginDate
				  RSObj("EndDate") = EndDate
				  RSObj("Content") = Content
				  RSObj("Host")=Host
				  RSObj("Guests")=Guests
				  RSObj("MediaUrl")=MediaUrl
				  RSObj("TemplateID")=TemplateID
				  RSObj("MessageVerifyTF")=MessageVerifyTF
				  RSObj("MessageTF")=MessageTF
				  RSObj("HostUserID")=HostUserID
				  RSObj("HostUserPass")=HostUserPass
				  RSObj("MessageLoginTF")=MessageLoginTF
				  RSObj("GuestsIntro")=GuestsIntro
				  RSObj("Locked")=Locked
			   RSObj.Update
				
				Call KS.FileAssociation(101191,InterViewID,RSObj("Content"),1)
			   RSObj.Close
			   Set RSObj = Nothing
			  End If
			  KS.Echo "<script src='../../ks_inc/jquery.js'></script>"
			  KS.Echo ("<script>alert('访谈修改成功!');location.href='KS.InterView.asp?Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='Post.asp?ButtonSymbol=Disabled&OpStr=在线访谈 >> <font color=red>访谈管理中心</font>';</script>")
			End If
		  End Sub
		  
		  '删除
	Sub InterViewDel()
		  		 Dim InterViewID, Page
				 Page = KS.G("Page")
				 InterViewID = Trim(KS.G("ID"))
				 Conn.Execute("Delete From KS_UploadFiles Where ChannelID=101191 and infoid in(" & KS.FilterIds(InterViewID) & ")")
				 Conn.Execute("Delete From KS_InterView Where ID in (" & KS.FilterIds(InterViewID) & ")")
				 Conn.Execute("Delete From KS_InterViewMsg Where InterViewID in (" & KS.FilterIds(InterViewID) & ")")
				 Conn.Execute("Delete From KS_InterViewRecord Where InterViewID in (" & KS.FilterIds(InterViewID) & ")")
				 Conn.Execute("Delete From KS_InterViewPic Where InterViewID in (" & KS.FilterIds(InterViewID) & ")")
				 KS.Echo ("<script>location.href='KS.InterView.asp?Page=" & Page & "';</script>")
	End Sub
		  
	

End Class
%>
 
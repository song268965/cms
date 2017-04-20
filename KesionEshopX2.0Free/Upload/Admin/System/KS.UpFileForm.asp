<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Const UploadBySwfUpload=true  '使用swfupload插件true 是 false否

Dim KSCls
Set KSCls = New UpFileFormCls
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileFormCls
        Private KS,BasicType,UpType,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  With KS
				 .echo "<!DOCTYPE html><html>"&vbcrlf
				 .echo "<head>"&vbcrlf
				 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"&vbcrlf
				 .echo "<title>上传文件</title>"&vbcrlf
				 .echo "<link rel=""stylesheet"" href=""../Include/admin_style.css"">"&vbcrlf
				 .echo "<script type=""text/javascript"">" &vbcrlf
				 .echo "function  doSubmit(obj){LayerPrompt.style.visibility='visible';UpFileForm.submit();}"&vbcrlf
				 .echo "</script>"&vbcrlf
				 .echo "<style type=""text/css"">" & vbCrLf
				 .echo "body {margin-left: 0px; margin-top: 0px;}" & vbCrLf
		         .echo "#uploadImg{  overflow:hidden; position:absolute}" & vbcrlf
				 .echo ".file{ cursor:pointer;position:absolute; z-index:100; margin-left:-180px; font-size:55px;opacity:0;filter:alpha(opacity=0); margin-top:-5px;}" & vbcrlf
				 .echo "</style></head>"
				 .echo "<body   oncontextmenu=""return false;"" style=""background-color:#FFFFFF"">"
		   ChannelID=KS.ChkClng(KS.G("ChannelID"))
		   UpType=KS.G("UpType")
		   
		If ChannelID<5000 Then BasicType=KS.C_S(ChannelID,6) Else BasicType=ChannelID
		
		If UPType="Field" OR  UpType="UpByBar" Then
		     Call UpFileByBar()
		ElseIf UpType="Pic" Then
			 Call UpDefaultPhoto()
		End IF
		 .echo "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left:2px; top: 0px; background-color: #ffffee; layer-background-color: #00CCFF; border: 1px solid #f9c943; width: 300px; height: 28px; visibility: hidden;"">"
		 .echo "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <tr>"
		 .echo "      <td><div>&nbsp;请稍等，正在上传文件<img src='../../images/default/wait.gif' align='absmiddle'></div></td>"
		' .echo "      <td width=""35%""><div align=""left""><font id=""ShowInfoArea"" size=""+1""></font></div></td>"
		 .echo "    </tr>"
		 .echo "  </table>"
		 .echo "</div>"
		 .echo "</body>"
		 .echo "</html>"
		End With
	  End Sub
	  
	  
		'上传缩略图接口
		Sub UpDefaultPhoto()
		 If UploadBySwfUpload Then Field_UpFile:exit sub
		Dim Path:Path = KS.GetUpFilesDir() & "/" 
			With KS
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""../Include/swfupload.asp?from=Common"">"
			 .echo "<span id=""uploadImg"">"
			 .echo "          <input type=""file"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			 .echo "          <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""选择本地图片并上传..."" ><span style=""color:red"">如果选择自动生成缩略图,则不开启图片裁剪窗口</span><input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		     .echo "          <label"
			  if KS.TBSetting(0)="0" then response.write " style='display:none'"
			 .echo "><input type=""checkbox"" name=""DefaultUrl"" value=""1"">自动生成缩略图</label>"
			 .echo "          <label"
			 if KS.TBSetting(5)="0" then response.write " style='display:none'"
			 .echo "><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "加水印</label>"
			 .echo " <input type=""hidden"" name=""AutoReName"" value=""4""></td></span>"
			 .echo "    </form>"
		  End With
		End Sub


		Sub Field_UpFile()
		%>
		<script type="text/javascript" src="../../plus/swfupload/swfupload/swfupload.js"></script>
		<script type="text/javascript" src="../../plus/swfupload/js/handlers.js"></script>
		<script>
		function uploadSuccess1(file, serverData) {
			try {
				if (serverData.substring(0, 6) == "error:") {
					alert(unescape(serverData).replace("error:",""));
				} else { 
				  <%If Request("Get")="min" and BasicType=5 Then%>
					parent.document.getElementById('<%=KS.G("FieldName")%>').value=serverData;
					parent.document.getElementById('<%=KS.G("imgname")%>').src=serverData;
				  <%ElseIf UpType="Field" or KS.G("FieldName")<>"" Then%>
					parent.document.getElementById('<%=KS.G("FieldName")%>').value=serverData;
					alert('恭喜文件上传成功！');
				  <%Else%>
				    var d=serverData.split('@');
				    parent.document.myform.PhotoUrl.value=d[0];
                   <%If KS.setting(186)="1" and KS.TBSetting(0)<>"0" then%>
				    if (document.getElementById('DefaultUrl').checked!=true){
					parent.OpenImgCutWindow(0,'<%=KS.Setting(3)%>',d[0]);
					}
				   <%end if%>

					<%If BasicType=1 Or BasicType=8 Then
							 Response.Write ("try{ if (parent.document.getElementById('ieditor')==undefined || (parent.document.getElementById('ieditor')!=undefined && parent.document.getElementById('ieditor').checked)){parent.insertHTMLToEditor('<img src=""'+d[1]+'"" />');}}catch(e){}")
					  ElseIf BasicType=3 Or BasicType=5 Then
							 Response.Write ("parent.document.myform.BigPhoto.value=d[1];")
					  End If
					  If Request("showpic")<>"" then
							 Response.Write ("parent.document.getElementById('" & request("showpic") & "').src=d[0];")
					  end if
				  End If%>
				}
			} catch (ex) {
				this.debug(ex);
			}
		}
		function fileDialogComplete1(numFilesSelected, numFilesQueued){
		 if (numFilesQueued>1){
		   alert('只能选择一个文件!');
		 }else if(numFilesQueued==1){
		  this.startUpload(this.getFile(0).ID);
		 }
		}
		var swfu;
		window.onload = function () {
		
			swfu = new SWFUpload({
				// Backend Settings
				upload_url: "../include/swfupload.asp",
				post_params: {"AdminID":"<%=KS.C("AdminID") %>","AdminPass":"<%=KS.C("PassWord")%>",UpType:"<%=UPType%>",BasicType:<%=BasicType%>,"upget":"<%=request("get")%>",ChannelID:<%=ChannelID%>,"FieldID":"<%=KS.G("FieldID")%>","AutoRename":4,"AddWaterFlag":1},

				// File Upload Settings
				file_size_limit : "<%=KS.ChkClng(KS.G("MaxFileSize"))%>",	// 限制大小
				<%if KS.G("UpType")="Pic" Then%>
				file_types :"*.<%=Replace(Replace(KS.ReturnChannelAllowUpFilesType(ChannelID,1),"|",","),",",";*.")%>",
				<%Else%>
				file_types : "*.*",
				<%End If%>
				//file_types_description : "支持.JPG.gif.png格式的图片",
				file_upload_limit : 0,

				// Event Handler Settings - these functions as defined in Handlers.js
				//  The handlers are not part of SWFUpload but are part of my website and control how
				//  my website reacts to the SWFUpload events.
				swfupload_preload_handler : preLoad,
				swfupload_load_failed_handler : loadFailed,
				file_queue_error_handler : fileQueueError,
				file_dialog_complete_handler : fileDialogComplete1,
				upload_progress_handler : uploadProgress,
				upload_error_handler : uploadError,
				upload_success_handler : uploadSuccess1,
				upload_complete_handler : null,

				// Button Settings
				//button_image_url : "../plus/swfupload/images/SmallSpyGlassWithTransperancy_17x18d.png",
				button_placeholder_id : "spanButtonPlaceholder",
				<%if request("get")="min" then%>
				button_width: 30,
				button_height: 22,
				button_text : '上传',
				<%else%>
				button_width: 115,
				button_height: 18,
				button_text : '<span class="button">选择文件并上传...</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#ffffff;} ',
				<%end if%>
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
		<table cellspacing="0" cellspadding="0" border="0">
		 <tr>
		  <td width="140">
		  <%if request("get")="min" then%>
		  <div style="background:url('../images/arrow.gif') no-repeat -2px -2px;padding-left:13px;border:1px solid #efefef;margin-top:2px;"><span id="spanButtonPlaceholder">选择文件</span></div>
		  <%else%>
		  <div class="button" style="margin-top:2px;"><span id="spanButtonPlaceholder">选择文件</span></div>
		  <%end if%></td>
		  <%If UpType<>"Field" and request("get")<>"min" Then%>
		  <td><label<%if KS.TBSetting(0)="0" then response.write " style='display:none'"%>><input type="checkbox" name="DefaultUrl" id="DefaultUrl" onclick="SetDefaultUrl(this)" value="1">自动生成缩略图</label> <label<%if KS.TBSetting(5)="0" then response.write " style='display:none'"%>><input name="AddWaterFlag" type="checkbox" id="AddWaterFlag" onclick="SetAddWaterFlag(this)" value="1" checked>添加水印</label> <span <%if KS.TBSetting(0)="0" or KS.setting(186)<>"1" then response.write " style='display:none'"%>>如果选择自动生成缩略图,则不开启图片裁剪窗口</span>
		  <script type="text/javascript">
		  function SetDefaultUrl(obj){if (obj.checked){swfu.addPostParam("DefaultUrl","1");}else{swfu.addPostParam("DefaultUrl","0");}}
		  function SetAddWaterFlag(obj){if (obj.checked){swfu.addPostParam("AddWaterFlag","1");}else{swfu.addPostParam("AddWaterFlag","0");}}
		  </script>
		  </td>
		  <%End If%>
		 </tr>
		</table>
		<%
		End Sub

		
		'上传文件带进度条
		Sub UpFileByBar()
		 dim CurrPath:CurrPath=KS.G("CurrPath")
		 if currpath<>"" then
		   if left(lcase(currpath),len(KS.Setting(3) & KS.Setting(91)))<>lcase(KS.Setting(3) & KS.Setting(91)) then 
		    currpath=KS.GetUpFilesDir
		   end if
		 end if
		%>
		<script src="../../ks_Inc/jquery.js"></script>
		<script src="../../ks_Inc/common.js"></script>
		<script type="text/javascript">
			var dir='<%=KS.Setting(3)%>';  //安装目录
			var uploadUrl="../include/swfupload.asp";  //上传处理文件地址
			<%If UPType="Field" Then%>
			var limitSize=<%=KS.ChkClng(KS.G("MaxFileSize"))%>; //限制大小 KB
			var fileExt="*.<%=Replace(Replace(KS.S("AllowFileExt"),"|",","),",",";*.")%>" //限制扩展名
			<%Else%>
			var limitSize=<%=round(KS.ReturnChannelAllowUpFilesSize(ChannelID))%>; //限制大小 KB
			var fileExt="*.<%=Replace(Replace(KS.ReturnChannelAllowUpFilesType(ChannelID,0),"|",","),",",";*.")%>" //限制扩展名
			<%End If%>
			var post_params={"AdminID":"<%=KS.C("AdminID") %>","AdminPass":"<%=KS.C("PassWord")%>",BasicType:<%=BasicType%>,ChannelID:<%=ChannelID%>,"UpType":"<%=UPType%>","FieldID":"<%=KS.G("FieldID")%>",AutoRename:4,currpath:"<%=currpath%>",from:"<%=ks.s("from")%>"<%If KS.G("from")="getfile" Then%>,"AddWaterFlag":1<%end if%>};
			var buttonstyle="color:#ffffff;";
			function uploadSuccess(file, serverData) {
				try {
					if (serverData.substring(0, 6) == "error:") {
						alert(unescape(serverData).replace("error:",""));
					 }else{
					 <%If KS.G("From")="getfile" then%>
					 	  jQuery("#FileUrl",parent.document).val(serverData);
					       parent.frames["FolderList"].location.reload();
					       parent.frames["FolderList"].PreviewFile(serverData);
			               parent.frames["PreviewArea"].location="../include/Preview.asp?FilePath="+escape(serverData);
					 <%ElseIf UpType="Field" Then%>
						parent.document.getElementById('<%=KS.G("FieldName")%>').value=serverData;
					  <%Else%>
						updateDisplay.call(this, file);
						var d=unescape(serverData).split('|');
						<%Select Case basictype
						  case 3  response.write "parent.SetDownUrlByUpLoad(d[0],d[1]);"
						  case 4  response.write "parent.document.getElementById('FlashUrl').value=d[0];"
						  case 7  response.write "parent.SetMovieUrlByUpLoad(d[0]);"
						  case 9  response.write "parent.document.getElementById('DownUrl').value=d[0];"
						  End Select
						End If%>
					}
				} catch (ex) {
					this.debug(ex);
				}
		}
		function SetAutoReName(obj){if (obj.checked){swfu.addPostParam("NoReName","0");}else{swfu.addPostParam("NoReName","1");}}
		function SetAddWaterFlag(obj){if (obj.checked){swfu.addPostParam("AddWaterFlag","1");}else{swfu.addPostParam("AddWaterFlag","0");}}
		</script>
		<script type="text/javascript" src="../../Plus/swfupload/swfupload/swfupload.js"></script>
		<script type="text/javascript" src="../../Plus/swfupload/swfupload/swfupload.queue.js"></script>
		<script type="text/javascript" src="../../Plus/swfupload/swfupload/swfupload.speed.js"></script>
		<%
		if basictype=7 then
		 response.write "<script>limitnum=0;</script>"
		end if
		%>
		<table border='0' cellpadding="0" cellspacing="0">
		 <tr><td><div class="uptips" id="showspeed"><div class="button" id="shows"><span id="spanButtonPlaceholder"></span></div></div></td>
		 <%If UpType<>"Field" Then%>
		 <td><label><input type="checkbox" onclick="SetAutoReName(this)" name="AutoReName" value="4" checked>自动更名</label></td>
		 <%End If%>
		 <%If KS.G("from")="getfile" and KS.TBSetting(5)<>"0" Then%>
		 <td><label><input name="AddWaterFlag" type="checkbox" id="AddWaterFlag" onclick="SetAddWaterFlag(this)" value="1" checked>添加水印</label> </td>
		 <%end if%>
		 </tr>
		</table>
		 <div id="tipss" style="display:none">
       <div id="UploadTips" style="padding:5px"><style>#UploadTips span{color:#ff6600;}</style><div style="display:none">Files Queued:<span id="tdFilesQueued"></span>Files Uploaded:	<span id="tdFilesUploaded"></span>Errors:	<span id="tdErrors"></span></div>当前速度:<span id="tdCurrentSpeed">0</span> 平均速度：<span  id="tdAverageSpeed">0</span><br/>已上传：<span id="tdPercentUploaded">0%</span> 大小：<span id="tdSizeUploaded">0</span>&nbsp;剩余时间：<span id="tdTimeRemaining">0</span> 已用时：<span id="tdTimeElapsed">0</span><br/><strong>正在上传中... 请耐心等待!!! 直到该提示框消失。</strong><span style="display:none">Moving Average Speed:<span id="tdMovingAverageSpeed"></span>Progress Event Count:<span id="tdProgressEventCount"></span></span></div></div>
       
       </div>
		<%
	End Sub
		
End Class
%> 

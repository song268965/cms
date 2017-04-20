﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New User_Upfile
KSCls.Kesion()
Set KSCls = Nothing

Class User_Upfile
        Private KS,KSUser,ChannelID,BasicType,UploadCode,UpType
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub


		Public Sub Kesion()
		With Response
		.Write "<!DOCTYPE HTML>"
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.Write "<title>上传文件</title>"
		.Write "<link rel=""stylesheet"" href=""images/css.css"">"
		
		.Write "<script type=""text/javascript"">" &vbcrlf
		.Write "function  doSubmit(obj){LayerPrompt.style.visibility='visible';UpFileForm.submit();}"&vbcrlf
		.Write "</script>"&vbcrlf
		.Write "<style type=""text/css"">" & vbCrLf
		.Write "body {margin-left: 0px; margin-top: 0px;}" & vbCrLf
		.Write "#uploadImg{  overflow:hidden; position:absolute}" & vbcrlf
		.Write ".file{ cursor:pointer;position:absolute; z-index:100; margin-left:-180px; font-size:55px;opacity:0;filter:alpha(opacity=0); margin-top:-5px;}" & vbcrlf
		.Write "</style></head>"


		.Write "<body class=tdbg style=""background-color:transparent"">"
		If KS.ChkClng(KS.S("FormID"))<>0 Or KS.ChkClng(KS.S("ChannelID"))=101 Then UploadFile:Response.end
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=999 Then
		 BasicType=ChannelID
		ElseIf ChannelID<5000 Then
		 BasicType=KS.C_S(ChannelID,6)
		Else
		 BasicType=ChannelID
		End If
		
		If KS.ChkClng(KS.C_S(ChannelID,26))<>2 and Cbool(KSUser.UserLoginChecked)=false  Then
		   .write "<font color='#999999'>Tips:登录后可使用上传功能</font>"
		   Exit Sub
		End If
		
		Randomize
		UpType=KS.S("Type")
		UploadCode = Int(900*rnd)+1000
		Session("UploadCode") = Cstr(UploadCode)


       If KS.S("Type")="Field" or UpType="UpByBar" Then
	     Call UpFileByBar
	   ElseIf KS.S("Type")="Pic" Then
		 Call UploadPhotoForm
	   End If
		End With
	End Sub
		
	'表单或会员注册表单的上传
	Sub UploadFile
		 Dim FormID:FormID=KS.ChkClng(KS.S("FormID"))
		 Dim FieldID:FieldID=KS.ChkClng(KS.S("FieldID"))
		 ChannelID=KS.ChkClng(KS.S("ChannelID"))
		 BasicType=KS.ChkClng(KS.S("ChannelID"))
		 If FormID=0 Then FormID=ChannelID  '会员注册表单上传
		 If FormID=0 Or FieldID=0 Then KS.Die "error!"
		 Dim RS,FieldName,MaxFileSize,AllowFileExtStr
		 If ChannelID=101 Then  '会员
		  Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_Field Where ChannelID=101 and FieldID=" & FieldID)
		 Else
		  Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_FormField Where FieldID=" & FieldID)
		 End If
		 If Not RS.Eof Then
		    FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
		 Else
		    Response.End()
		 End IF
		 RS.Close:Set RS=Nothing
		%>
		<script type="text/javascript" src="../plus/swfupload/swfupload/swfupload.js"></script>
		<script type="text/javascript" src="../plus/swfupload/js/handlers.js"></script>
		<script>
		function uploadSuccess1(file, serverData) {
			try {
				if (serverData.substring(0, 6) == "error:") {
					alert(unescape(serverData).replace("error:",""));
				} else { 
					parent.document.getElementById('<%=FieldName%>').value=serverData;
					alert('恭喜,文件上传成功！');
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
				upload_url: "swfupload.asp",
				post_params: {"UserID" : "<%=KS.C("UserID") %>","UserName" : "<%=KS.C("UserName") %>","RndPassWord":"<%=KS.C("RndPassWord")%>",UpType:"Field",ChannelID:"<%=ChannelID%>",BasicType:"<%=BasicType%>",FormID:"<%=KS.S("FormID")%>","FieldID":"<%=KS.S("FieldID")%>","AutoRename":4},

				// File Upload Settings
				file_size_limit : "<%=round(MaxFileSize)%>",	// 2MB
				file_types : "*.<%=Replace(Replace(AllowFileExtStr,"|",","),",",";*.")%>",
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
				button_width: 135,
				button_height: 22,
				button_text : '<span class="button">选择文件(限<%=round(MaxFileSize/1024)%>M)</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#ffffff;font-size: 14px; } ',
				button_text_top_padding: 3,
				button_text_left_padding: 0,
				button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
				button_cursor: SWFUpload.CURSOR.HAND,
				
				// Flash Settings
				flash_url : "../plus/swfupload/swfupload/swfupload.swf",
				flash9_url : "../plus/swfupload/swfupload/swfupload_FP9.swf",

				custom_settings : {
					upload_target : "divFileProgressContainer"
				},
				
				// Debug Settings
				debug: false
			});
		};
		</script>
		<div class="pn"<%if channelid<>0 and channelid<>101 then%> onmousedown="return(parent.CheckClassID());"<%end if%> style="margin:0;width:100px;">
		 <span id="spanButtonPlaceholder"></span>
		</div>
		<%
		End Sub
		
		
		
		'普通上传接口
		Sub CommonUpload()
			With KS
			 .echo "    <form name=""UpFileForm"" id=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""swfupload.asp?from=Common"">"
			 .echo "<span id=""uploadImg"">"
			 .echo "          <input type=""file"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			 If ChannelID=9999 Then
			  .echo "          <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""拍照上传"" >"
			 Else
			  .echo "          <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""选择并上传..."" >"
			 End If
		     .echo "          <label><input type=""hidden"" name=""DefaultUrl"" value=""1""  checked></label>"
			 .echo "          <label><input name=""AddWaterFlag"" type=""hidden"" id=""AddWaterFlag"" value=""1"" checked></label>"
			 .echo " <input type=""hidden"" name=""AutoReName"" value=""4""></td></span>"
			 .echo "    </form>"
			 .echo "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left:2px; top: 0px; background-color: #ffffee; layer-background-color: #00CCFF; border: 1px solid #f9c943; width: 300px; height: 28px; visibility: hidden;"">"
		 .echo "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <tr>"
		 .echo "      <td><div>上传中<img src='../images/default/wait.gif' align='absmiddle'></div></td>"
		' .echo "      <td width=""35%""><div align=""left""><font id=""ShowInfoArea"" size=""+1""></font></div></td>"
		 .echo "    </tr>"
		 .echo "  </table>"
		 .echo "</div>"
		  End With
		End Sub
		
		'上传缩略图
		Sub UploadPhotoForm
		  If instr(request.servervariables("http_user_agent"),"Mobile")>0 Then '手机访问,自动跳到手机版
		   Call CommonUpload()
		   Exit Sub
		 end if
		%>
        <script src="/ks_inc/jquery.js" type="text/javascript"></script>
		<script type="text/javascript" src="../plus/swfupload/swfupload/swfupload.js"></script>
		<script type="text/javascript" src="../plus/swfupload/js/handlers.js"></script>
		<script>
		function uploadSuccess1(file, serverData) {
			try {
				if (serverData.substring(0, 6) == "error:") {
					alert(unescape(serverData).replace("error:",""));
				} else { 
				  <%If UpType="Field" or KS.G("FieldName")<>"" Then%>
					parent.document.getElementById('<%=KS.G("FieldName")%>').value=unescape(serverData);
					<%If ChannelID<>9994 then%>
					alert('恭喜文件上传成功！');
					<%End If%>
				  <%ElseIf ChannelID=9996 Then  '圈子封面%>
					  parent.document.myform.showimages.src=unescape(serverData);
					  parent.document.myform.PhotoUrl.value=unescape(serverData);
				  <%ElseIf ChannelID=9999 Then  '头像%> 
				   alert('恭喜，上传成功！');
				   top.location.href='User_EditInfo.asp?action=face&PhotoUrl='+serverData;
				  <%ElseIf ChannelID=55666 Then  '广告竞价%>  
				  	alert('恭喜，上传成功！');
					$(window.parent.document).find("#imgIcon").attr("src",serverData)
					parent.document.getElementById("gif_url").value=serverData;
				  <%ElseIf ChannelID=7999 Or ChannelID=7998 Then  '企业动态%> 
				    var d=unescape(serverData).split('@');
				    parent.document.myform.PhotoUrl.value=d[0];
					parent.insertHTMLToEditor('<img src="'+d[1]+'" />')
				  <%Else%>
				    var d=unescape(serverData).split('@');
				    parent.document.myform.PhotoUrl.value=d[0];
				    
					<%If BasicType=1 Or BasicType=8 Then
							 Response.Write ("try{parent.insertHTMLToEditor('<img src=""'+d[1]+'"" />');}catch(e){}")
					  ElseIf BasicType=3 Or BasicType=5 Then
							 Response.Write ("parent.document.myform.BigPhoto.value=d[1];")
					  End If
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
		var post_params={"UserID" : "<%=KS.C("UserID") %>","UserName" : "<%=KS.C("UserName") %>","RndPassWord":"<%=KS.C("RndPassWord")%>",UpType:"<%=UPType%>",BasicType:<%=BasicType%>,ChannelID:<%=ChannelID%>,"BoardID":"<%=KS.S("BoardID")%>","FieldID":"<%=KS.G("FieldID")%>","AutoRename":4<%if channelid<>"9999" and  channelid<>"55666"  and channelid<>"9994" and channelid<>"9993" and channelid<>"9990" and channelid<>"8000" then '上传头像不生成小图%>,"AddWaterFlag":1,"DefaultUrl":1<%End If%>};
		window.onload = function () {
		
			swfu = new SWFUpload({
				// Backend Settings
				upload_url: "swfupload.asp",
				post_params: post_params,
				// File Upload Settings
				file_size_limit : "<%=KS.ChkClng(KS.G("MaxFileSize"))%>",	// 限制大小
				<%If KS.G("ext")<>"" Then%>
				file_types : "<%=replace(KS.G("ext"),"＊","*")%>",
				<%Else%>
				file_types : "*.*",
				<%End If%>
				//file_types_description : "支持.JPG.gif.png格式的图片",
				file_upload_limit : 0,  //限制只能上传一个文件

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
				button_width: 75,
				button_height: 28,
				<%if channelid="9999" then%>
				button_text : '<span class="btn">上传头像</span>',
				<%Elseif channelid="55666" then %>
				button_text : '<span class="btn">上传图片</span>',
				<%Else%>
				button_text : '<span class="btn">上传图片</span>',
				<%End If%>
				button_text_style : '.btn{color:#ffffff;font-weight:bold}',
				button_text_top_padding: 3,
				button_text_left_padding: 10,
				button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
				button_cursor: SWFUpload.CURSOR.HAND,
				
				// Flash Settings
				flash_url : "../plus/swfupload/swfupload/swfupload.swf",
				flash9_url : "../plus/swfupload/swfupload/swfupload_FP9.swf",

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
		  <td width="80"><div class="uploadbutton" <%if channelid<>0 and channelid<1000 then%> onmousedown="return(parent.CheckClassID());"<%end if%>><span id="spanButtonPlaceholder">选择文件</span></div></td>
		  <%If UpType<>"Field" and channelid<>"9999" and  channelid<>"55666" and channelid<>"9990" and channelid<>"9993" and channelid<>"9994" and channelid<>"8000" and KS.TBSetting(0)<>"0" Then%>
		  <td><label class="tiy"><input type="checkbox" name="DefaultUrl" id="DefaultUrl" checked="checked" onclick="SetDefaultUrl(this)" value="1">生成缩略图</label><%If BasicType<>3 And BasicType<>2 And BasicType<>4 And BasicType<>7 and channelid<1000 Then%> <label><input name="AddWaterFlag" type="checkbox" id="AddWaterFlag" onclick="SetAddWaterFlag(this)" value="1" checked>添加水印</label><%End If%>
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
		%>
		<script src="../ks_Inc/jquery.js"></script>
		<script src="../ks_Inc/common.js"></script>
		<script type="text/javascript">
			var dir='<%=KS.Setting(3)%>';  //安装目录
			var uploadUrl="swfupload.asp";  //上传处理文件地址
		<%If UPType="Field" Then%>
			var limitSize=<%=KS.ChkClng(KS.G("MaxFileSize"))%>; //限制大小 KB
			var fileExt="*.<%=Replace(Replace(KS.S("AllowFileExt"),"|",","),",",";*.")%>" //限制扩展名
		<%ElseIf ChannelID=9995 Then%>
			var limitSize=5000; //限制大小 KB
			var fileExt="*.mp3" //限制扩展名
		<%Else%>
			var limitSize=<%=round(KS.ReturnChannelAllowUpFilesSize(ChannelID))%>; //限制大小 KB
			<%If ChannelID=4 Then%>
			var fileExt="*.<%=Replace(Replace(KS.ReturnChannelAllowUpFilesType(ChannelID,2),"|",","),",",";*.")%>" //限制扩展名
			<%ElseIf ChannelID=7 Then%>
			var fileExt="*.<%=Replace(Replace(KS.ReturnChannelAllowUpFilesType(ChannelID,2) &"|" & KS.ReturnChannelAllowUpFilesType(ChannelID,3) & "|"& KS.ReturnChannelAllowUpFilesType(ChannelID,4),"|",","),",",";*.")%>" //限制扩展名
			<%Else%>
			var fileExt="*.<%=Replace(Replace(KS.ReturnChannelAllowUpFilesType(ChannelID,0),"|",","),",",";*.")%>" //限制扩展名
			<%
			 End If
		End If%>
			var post_params={"UserID" : "<%=KS.C("UserID") %>","UserName" : "<%=KS.C("UserName") %>","RndPassWord":"<%=KS.C("RndPassWord")%>",BasicType:<%=BasicType%>,ChannelID:<%=ChannelID%>,"UpType":"<%=UPType%>","FieldID":"<%=KS.G("FieldID")%>",AutoRename:4};
			var buttonstyle="color:#ffffff;";
			function uploadSuccess(file, serverData) {
				try {
					if (serverData.substring(0, 6) == "error:") {
						alert(unescape(serverData).replace("error:",""));
					 }else{
					 <%If UpType="Field" Or KS.G("FieldName")<>"" Then%>
					   $("#<%=KS.G("FieldName")%>",parent.document).val(unescape(serverData));
					   try{
					   $("#<%=KS.G("FieldName")%>_Src",parent.document).attr("src",unescape(serverData));
					   }catch(ex){
					   }
						//parent.document.getElementById('<%=KS.G("FieldName")%>').value=unescape(serverData);
					  <%Else%>
						updateDisplay.call(this, file);
						var d=unescape(serverData).split('|');
						<%Select Case basictype
						  case 3  response.write "parent.SetDownUrlByUpLoad(d[0],d[1]);"
						  case 4  response.write "parent.document.getElementById('FlashUrl').value=d[0];"
						  case 7  response.write "parent.document.getElementById('MovieUrl').value=d[0];"
						  case 9  response.write "parent.document.getElementById('DownUrl').value=d[0];"
						  End Select
						End If%>
					}
				} catch (ex) {
					this.debug(ex);
				}
		}
		function SetAutoReName(obj){if (obj.checked){swfu.addPostParam("NoReName","0");}else{swfu.addPostParam("NoReName","1");}}
		</script>
		<script type="text/javascript" src="../Plus/swfupload/swfupload/swfupload.js"></script>
		<script type="text/javascript" src="../Plus/swfupload/swfupload/swfupload.queue.js"></script>
		<script type="text/javascript" src="../Plus/swfupload/swfupload/swfupload.speed.js"></script>
		<table border='0' cellpadding="0" cellpadding="0">
		 <tr><td><div class="uptips" id="showspeed"<%if channelid<>0 and channelid<>9 and channelid<1000 then%> onmousedown="return(parent.CheckClassID());"<%end if%>><div class="button" id="shows"><span id="spanButtonPlaceholder"></span></div></div></td>
		 <%If UpType<>"Field" Then%>
		 <td><label class="tiy"><input type="checkbox" onclick="SetAutoReName(this)" name="AutoReName" value="4" checked>自动更名</label></td>
		 <%End If%>
		 </tr>
		</table>
		 <div id="tipss" style="display:none">
       <div id="UploadTips" style="padding:5px"><style>#UploadTips span{color:#ff6600;}</style><div style="display:none">Files Queued:<span id="tdFilesQueued"></span>Files Uploaded:	<span id="tdFilesUploaded"></span>Errors:	<span id="tdErrors"></span></div>当前速度:<span id="tdCurrentSpeed">0</span> 平均速度：<span  id="tdAverageSpeed">0</span><br/>已上传：<span id="tdPercentUploaded">0%</span> 大小：<span id="tdSizeUploaded">0</span>&nbsp;剩余时间：<span id="tdTimeRemaining">0</span> 已用时：<span id="tdTimeElapsed">0</span><br/><strong>正在上传中... 请耐心等待!!! 直到该提示框消失。</strong><span style="display:none">Moving Average Speed:<span id="tdMovingAverageSpeed"></span>Progress Event Count:<span id="tdProgressEventCount"></span></span></div></div>
       
       </div>

		<%
	End Sub


End Class
%> 

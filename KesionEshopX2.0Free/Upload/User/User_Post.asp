<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New PostCls
KSCls.Kesion()
Set KSCls = Nothing

Class PostCls
        Private KS,KSUser,ChannelID,ID,ClassID,RS,Selbutton,Action
		Private LoginTF,FieldXML,FieldNode,FNode,FieldDictionary
		Private PhotoUrl,ShowStyle,PageNum
		Private DownLBList, DownYYList, DownSQList, DownPTList
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
        
		Public Sub LoadMain()
		ChannelID=KS.ChkClng(KS.S("ChannelID")) : Action=KS.S("Action")
		If ChannelID=0 Then ChannelID=1
		LoginTF=Cbool(KSUser.UserLoginChecked)
		IF LoginTF=false  Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		if KS.C_S(ChannelID,36)=0 then
		  Call KS.ShowTips("error","<li>本频道不允许投稿!</li>")
		  Exit Sub
		end if
		Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
		'设置缩略图参数
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		Call KSUser.Head()
		Select Case Action
		 Case "Add","Edit" 
		   Select Case  KS.ChkCLng(KS.C_S(ChannelID,6))
		     Case 1 Call InitialArticle()
			 Case 2 Call InitialPhoto()
			 Case 3 Call InitialDownLoad()
		   End Select
		 Case "DoSave" 
		   Select Case  KS.ChkCLng(KS.C_S(ChannelID,6))
		     Case 1 Call DoSaveArticle()
			 Case 2 Call DoSavePhoto()
			 Case 3 Call DoSaveDownLoad()
		   End Select
		End Select
	   End Sub
%>
<!--#include file="../ks_cls/UserFunction.asp"-->
<%
 '添加文章
 Sub InitialArticle()
        ID=KS.ChkClng(KS.S("id"))
        Call KSUser.InnerLocation("发布"& KS.C_S(ChannelID,3))
		If ID<>0 Then
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
	     RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
		 If Not RS.Eof Then
		  ClassID=RS("Tid") : SelButton=KS.C_C(ClassID,1)
		 End If
		Else
		 SelButton="选择栏目..."
	    End If
		Response.Write(EchoUeditorHead())
		%>
		<script type="text/javascript">
		   var box='';
		  function addMap(){
		   box=$.dialog.open('../plus/baidumap.asp?from=user&action=getcenter&MapMark='+escape($('#MapMark').val()),{title:'电子地图标注',width:'830px',height:'430px'}); }
		  </script>
		<script language = "JavaScript">
		   $(document).ready(function(){
		    <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='turnto']/showonuserform").text="1" Then%>
				   if ($("#Changes").prop('checked')){ChangesNews();}
			 <%End If%>
		   });
		   function ChangesNews()
			{ 
			 if ($("#Changes:checked").val()=="1")
			  {
			  <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonuserform").text="1" Then%>
			  $("#ContentArea").hide();
			  <%end if%>
			  $("#ChangesUrl").attr("disabled",false);
			  }
			  else
			   {
			   <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonuserform").text="1" Then%>
			   $("#ContentArea").show();
			   <%end if%>
			  $("#ChangesUrl").attr("disabled",true);
			   }
			}
		   function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){$('#KeyWords').val(unescape(data)).attr("disabled",false);});
			  }else{$.dialog.alert('对不起,请先输入标题!',function(){document.myform.Title.focus();}); }
			}
		    function CheckClassID(){
				if (document.myform.ClassID.value=="0" || document.myform.ClassID.value=='') {
					$.dialog.alert("请选择<%=KS.C_S(ChannelID,3)%><%=KS.GetClassName(channelid)%>！",function(){});
					return false;}		
				  return true;
			}
			function insertHTMLToEditor(codeStr){ <%=InsertEditor("Content","codeStr")%>} 
			function CheckForm(){
				if (document.myform.ClassID.value=="0") 
				  {
					$.dialog.alert("请选择<%=KS.C_S(ChannelID,3)%><%=KS.GetClassName(channelid)%>！",function(){});
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>标题！",function(){
					document.myform.Title.focus();
					});
					return false;
				  }
				<%Call LFCls.ShowDiyFieldCheck(FieldXML,1)%>
				
				<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='turnto']/showonuserform").text="1" Then%>
				if (($("#Changes:checked").val()=="1")&&($("input[name=ChangesUrl]").val()==""))
				  { $("#ChangesUrl").focus();
					$.dialog.alert("请输入外部链接的Url！");
					return false;
				  }
				<%end if%>
				
				<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonuserform").text="1" Then%>
				     if (($("#Changes:checked").val()!="1")&&(<%=GetEditorContent("Content")%>==false))
					{
					  $.dialog.alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！",function(){
					   <%=GetEditorFocus("Content")%>
					  });
					  return false;
					}
				<%end if%>
				 return true; 
				 }
		</script>
		<%
		Call ShowPostForm()
		if isobject(rs) then
		if rs.state=1 then rs.close:Set rs=nothing
		end if
 End Sub
 '添加图片
 Sub InitialPhoto()
        Dim PicUrls
        ID=KS.ChkClng(KS.S("id"))
        Call KSUser.InnerLocation("发布"& KS.C_S(ChannelID,3))
		If ID<>0 Then
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
	     RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
		 If Not RS.Eof Then
		  ClassID=RS("Tid") : SelButton=KS.C_C(ClassID,1) : PicUrls=RS("PicUrls")
		 End If
		Else
		 SelButton="选择栏目..."
	    End If
		Response.Write(EchoUeditorHead())
   %>
		   <style type="text/css">
			#thumbnails{background:url(../plus/swfupload/images/albviewbg.gif) no-repeat;_height:expression(document.body.clientHeight > 200? "200px": "auto" ); clear:both; margin-top:20px;}
			#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:152px;height:175px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
			#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px}
			#pagenums{clear: both;padding-top: 20px;}
			.dtable dd label input[type="checkbox"]{ vertical-align:middle;}
			.swfupload{ margin-top:5px;}
			.dtable dd label{line-height: 34px;margin-left: 10px;}
			#thumbnails .thumbshow>div{ margin-top:0 !important;}
			#thumbnails .thumbshow>div input[type="text"]{height: 22px;border: 1px solid #ddd;}
			
			</style>
			<link href="../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
			<script type="text/javascript" src="../plus/swfupload/swfupload/swfupload.js"></script>
			<script type="text/javascript" src="../plus/swfupload/js/handlers.js"></script>
			<script type="text/javascript">
		   var box='';
		  function addMap(){
		   box=$.dialog.open('../plus/baidumap.asp?from=user&action=getcenter&MapMark='+escape($('#MapMark').val()),{title:'电子地图标注',width:'830px',height:'430px'}); }
		  </script>
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
		{
		  var p=$('#pic'+pid).val();
		   if (p!==''){
		    $.ajax({
			  url: "../plus/ajaxs.asp",
			  cache: false,
			  data: "action=DelPhoto&pic="+p+"&flag=0",
			  success: function(r){
			  }
			  });
	       }
		   $("#thumbshow"+pid).remove();
		}	
		
		function addImage(bigsrc,smallsrc,text) {
			var newImgDiv = document.createElement("div");
			var delstr = '';
			delstr = '<a href="javascript:DelUpFiles('+pid+')" style="color:#ff6600">[删除]</a>';
			newImgDiv.className = 'thumbshow';
			newImgDiv.id = 'thumbshow'+pid;
			document.getElementById("thumbnails").appendChild(newImgDiv);
			newImgDiv.innerHTML = '<a href="'+bigsrc+'" target="_blank"><span id="show'+pid+'"></span></a>';
			newImgDiv.innerHTML += '<div style="margin-top:10px;text-align:left">'+delstr+' <b>注释：</b><input type="hidden" class="pics" id="pic'+pid+'" value="'+bigsrc+'|'+smallsrc+'"/><input type="text" name="picinfo'+pid+'" value="'+text+'" style="width:148px;" /></div>';
		
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
				upload_url: "swfupload.asp",
				post_params: {"UserID":"<%=KSUser.GetUserInfo("userid")%>","BasicType":<%=KS.C_S(ChannelID,6)%>,AddWaterFlag:"1","ChannelID":<%=ChannelID%>,"UserName" : "<%=KS.C("UserName") %>","RndPassWord":"<%=KS.C("RndPassWord")%>","AutoRename":4},

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
				button_image_url: "",
				button_placeholder_id : "spanButtonPlaceholder",
				button_width: 100,
				button_height: 20,
				button_text : '<span class="button">批量上传图片</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#ffffff;font-weight:bold;font-size: 14px; } ',
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
	<script type="text/javascript">
	function OnlineCollect(){
	 box=$.dialog.open("../editor/ksplus/remotefile.asp",{title:"网上采集图片",width:550,height:200});
	}
	function AddTJ(){
	box=$.dialog({max:false,title:'从上传文件中选择',content:"<div style='padding:3px'><strong>小图地址:</strong><input type='text' name='x1' id='x1'> <input type='button' onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择小图")%>&ChannelID=<%=ChannelID%>',550,290,window,$('#x1')[0]);\" value='选择小图' class='button'/><br/><strong>大图地址:</strong><input type='text' name='x2' id='x2'> <input type='button' onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择小图")%>&ChannelID=<%=ChannelID%>',550,290,window,$('#x2')[0]);\" value='选择大图' class='button'/><br/><strong>简要介绍:</strong><input type='text' name='x3' id='x3'><br/><br/><input type='button' value='加 入' onclick='ProcessAddTj()' class='button'/> <input type='button' value='取 消' class='button' onclick='box.close()'/></div>",width:420,height:200});
	}
	function ProcessAddTj(){
	  if ($("#x1").val()==''){
	   $.dialog.alert('请选择一张小图地址!',function(){
	   $("#x1").focus();});
	   return false;
	  }
	  if ($("#x2").val()==''){
	   $.dialog.alert('请选择一张大图地址!',function(){
	   $("#x2").focus();});
	   return false;
	  }
	  addImage($("#x2").val(),$("#x1").val(),$("#x3").val())
	  $("#x2").val('');
	  $("#x1").val('');
	  $("#x3").val('');
	  box.close();
	}
	function ProcessCollect(collecthttp){
	 if (collecthttp==''){
	   $.dialog.alert('请输入远程图片地址,一行一张地址!',function(){
	  });
	   return false;
	 }
	 var carr=collecthttp.split('\n');
	 for(var i=0;i<carr.length;i++){
	   if (carr[i]!=''){
	   var bigsrc=carr[i];
	   var smallsrc=carr[i];
	   addImage(bigsrc,smallsrc,'');
	   }
	 }
	 parent.box.close();
	}
	</script>
	 
	<%
	Call ShowPostForm()
	%>
  <script type="text/javascript">
	 $(document).ready(function(){IniPicUrl();})
	function IniPicUrl(){
		 var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
		 var PicUrlArr=null;
		 if (PicUrls!=''){ 
				PicUrlArr=PicUrls.split('|||');
			    for ( var i=1 ;i<PicUrlArr.length+1;i++){ 
			      addImage(PicUrlArr[i-1].split('|')[1],PicUrlArr[i-1].split('|')[2],PicUrlArr[i-1].split('|')[0]);
			    }
			   }
	}
	function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{
			   $.dialog.alert('对不起,请先输入标题!',function(){document.myform.Title.focus();});
			  }
			}
			function CheckForm()
				{
				if (document.myform.ClassID.value=="0") 
				  {
					$.dialog.alert("请选择<%=KS.C_S(ChannelID,3)%><%=KS.GetClassName(channelid)%>！",function(){});
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>名称！",function(){document.myform.Title.focus();});
					return false;
				  }		
				if (document.myform.PhotoUrl.value==''<%if KS.S("Action")="Add" Then response.write " && $('#autothumb').attr('checked')==false"%>)
				{
					$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>缩略图！",function(){document.myform.PhotoUrl.focus();});
					return false;
				}
				<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
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
				if ($('input[name=PicUrls]').val()=='')
				{
				  $.dialog.alert('请输入<%=KS.C_S(ChannelID,3)%>内容!',function(){
				  $('input[name=imgurl1]').focus();});
				  return false;
				}
				  return true;
                    
				}
				function CheckClassID()
				{
				 if (document.myform.ClassID.value=="0") 
				  {
					$.dialog.alert("请选择<%=KS.C_S(ChannelID,3)%><%=KS.GetClassName(channelid)%>！",function(){});
					return false;
				  }		
				  return true;
				}
			</script>
			 <%
		If IsOBject(RS) Then
		  If rs.status<>0 Then rs.close:set rs=nothing
		End If
 End Sub
 
 '添加下载
 Sub InitialDown(DownLb,DownYY,DownSQ)
   Dim I, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
	Set RSP = Server.CreateObject("Adodb.RecordSet")
	 RSP.Open "Select * From KS_DownParam Where ChannelID=" & ChannelID, conn, 1, 1
	 If Not RSP.Eof Then
		DownLBStr = RSP("DownLB")
		DownYYStr = RSP("DownYY")
		DownSQStr = RSP("DownSQ")
		DownPTStr = RSP("DownPT")
	End If
		RSP.Close:Set RSP = Nothing
		'下载类别
		LBArr = Split(DownLBStr, vbCrLf)
		For I = 0 To UBound(LBArr)
			If LBArr(I) = DownLb Then
			 DownLBList = DownLBList & "<option value='" & LBArr(I) & "' Selected>" & LBArr(I) & "</option>"
			Else
			 DownLBList = DownLBList & "<option value='" & LBArr(I) & "'>" & LBArr(I) & "</option>"
			End If
		Next
		'下载语言
		YYArr = Split(DownYYStr, vbCrLf)
		For I = 0 To UBound(YYArr)
		  If YYArr(I) = DownYY Then
			DownYYList = DownYYList & "<option value='" & YYArr(I) & "' Selected>" & YYArr(I) & "</option>"
		  Else
			DownYYList = DownYYList & "<option value='" & YYArr(I) & "'>" & YYArr(I) & "</option>"
		  End If
		 Next
		'下载授权
		SQArr = Split(DownSQStr, vbCrLf)
		For I = 0 To UBound(SQArr)
			If SQArr(I) = DownSQ Then
				DownSQList = DownSQList & "<option value='" & SQArr(I) & "' Selected>" & SQArr(I) & "</option>"
			Else
				DownSQList = DownSQList & "<option value='" & SQArr(I) & "'>" & SQArr(I) & "</option>"
			End If
		Next
		'下载平台
		PTArr = Split(DownPTStr, vbCrLf)
		For I = 0 To UBound(PTArr)
				DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
		Next
 End Sub
 Sub InitialDownLoad()
        Dim DownLb,DownYY,DownSQ
        ID=KS.ChkClng(KS.S("id"))
        Call KSUser.InnerLocation("发布"& KS.C_S(ChannelID,3))
		If ID<>0 Then
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
	     RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
		 If Not RS.Eof Then
		  ClassID=RS("Tid") : SelButton=KS.C_C(ClassID,1)
		  DownLb=RS("DownLb") : DownYY=RS("DownYY") :DownSQ=RS("DownSQ")
		 End If
		Else
		 SelButton="选择栏目..."
	    End If
		Response.Write(EchoUeditorHead())
		Call InitialDown(DownLb,DownYY,DownSQ)
		%>
		<script language = "JavaScript">
		   function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){$('#KeyWords').val(unescape(data)).attr("disabled",false);});
			  }else{$.dialog.alert('对不起,请先输入标题!',function(){document.myform.Title.focus();}); }
			}
		    function SetDownPT(addTitle){
					var str=document.myform.DownPT.value;
					if (document.myform.DownPT.value=="") {
						document.myform.DownPT.value=document.myform.DownPT.value+addTitle;
					}else{
						if (str.substr(str.length-1,1)=="/"){
							document.myform.DownPT.value=document.myform.DownPT.value+addTitle;
						}else{
							document.myform.DownPT.value=document.myform.DownPT.value+"/"+addTitle;
						}
					}
					document.myform.DownPT.focus();
				}

				function SetPhotoUrl()
				{
				 if (document.myform.DownUrl.value!='')
				  document.myform.PhotoUrl.value=document.myform.DownUrl.value.split('|')[1];	
				}
				function SetDownUrlByUpLoad(DownUrlStr,FileSize)
				{  $("#DownUrlS").val(DownUrlStr);
				   <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='nature']/showonuserform").text="1" Then%>
				    if (FileSize!=0)
					{ 
					  if (FileSize/1024/1024>1)
					  {
					   $("input[name=SizeUnit]")[1].checked=true;
					   document.getElementById('DownSize').value=(FileSize/1024/1024).toFixed(2); 
					  }
					  else{
					  document.getElementById('DownSize').value=(FileSize/1024).toFixed(2);
					  $("input[name=SizeUnit]")[0].checked=true;
					  }
				   }
				  <%end if%>
				var UrlStrArr;
				   UrlStrArr=DownUrlStr.split('|');
				   for (var i=0;i<UrlStrArr.length-1;i++)
				   {
				   var url=UrlStrArr[i]; 
				   if(url!=null&&url!=''){document.myform.DownUrlS.value=url;} 
				  }
				}
				function CheckClassID()
				{
				if (document.myform.ClassID.value=="0" || document.myform.ClassID.value=='') 
				  {
					$.dialog.alert("请选择<%=KS.C_S(ChannelID,3)%><%=KS.GetClassName(channelid)%>！",function(){});
					return false;
				  }		
				  return true;
				}
				function CheckForm()
				{   
					if (document.myform.ClassID.value=="0") 
						{
							$.dialog.alert("请选择<%=KS.C_S(ChannelID,3)%><%=KS.GetClassName(channelid)%>！",function(){});
							//document.myform.ClassID.focus();
							return false;
					 }		
				 if (document.myform.Title.value=="")
					  {
						$.dialog.alert("请输入<%=KS.C_S(ChannelID,3)%>名称！",function(){
						document.myform.Title.focus();});
						return false;
					  }
					if (document.myform.DownUrlS.value=='')
					{
						$.dialog.alert("请添加<%=KS.C_S(ChannelID,3)%>！",function(){
						document.myform.DownUrlS.focus();});
						return false;
					}
					<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
					return true;
				}
		</script>
		<%
		Call ShowPostForm()
		if isobject(rs) then
		if rs.state=1 then rs.close:Set rs=nothing
		end if
 End Sub 
 
Sub ShowPostForm()
  %>
<iframe src="about:blank" name="hidframe" style="display:none"></iframe>
<form  action="User_post.asp?channelid=<%=channelid%>&Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" target="hidframe">
		<%
		Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
		If KS.IsNul(XmlForm) Then 
		 GetInputForm false,ChannelID,FieldXML,FieldNode,FieldDictionary,KS.ChkClng(KS.S("id")),KSUser,rs
		Else
		   If Action="Edit" Then
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
		  Else 
		     Call KSUser.CheckMoney(ChannelID)
		  End If
		 Scan XmlForm
		End If
%>
</form>
<%
 End Sub
 
  
 Sub DoSaveArticle()
   	Dim Title,FullTitle,KeyWords,Author,Origin,Intro,Content,Verific,PhotoUrl,I,Province,City,County,FileIds,ReadPoint,ChangesUrl,Changes,AdminMobile
    ClassID=KS.S("ClassID")
	ID=KS.ChkClng(KS.S("ID"))
	If KS.ChkClng(KS.C_C(ClassID,20))=0 Then
	 Response.Write "<script>$.dialog.tips('对不起,系统设定不能在此栏目发表,请选择其它栏目!',1,'error.gif',function(){history.back();});</script>":Exit Sub
	End IF
	Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
	KeyWords=KS.LoseHtml(KS.S("KeyWords"))
	Author=KS.LoseHtml(KS.S("Author"))
	Origin=KS.LoseHtml(KS.S("Origin"))
	Content = Request.Form("Content")
	Content=KS.FilterIllegalChar(KS.ClearBadChr(content))
	FileIds=LFCls.GetFileIDFromContent(Content)
	Changes     = KS.ChkClng(KS.G("Changes"))
	ChangesUrl  = Trim(Request("ChangesUrl"))
	if Changes=1 Then Content=ChangesUrl
				 
	if KS.IsNul(Content) Then Content="&nbsp;"
	Verific=KS.ChkClng(KS.S("Status"))
	Intro  = KS.FilterIllegalChar(KS.LoseHtml(KS.S("Intro")))
	Province= KS.LoseHtml(KS.S("Province"))
	City    = KS.LoseHtml(KS.S("City"))
	County  = KS.LoseHtml(KS.S("County"))
	FullTitle = KS.LoseHtml(KS.S("FullTitle"))
	if Intro="" And KS.ChkClng(KS.S("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(Request.Form("Content")),200)
				 
	Dim Fname,FnameType,TemplateID,WapTemplateID
	If KS.ChkClng(KS.S("ID"))=0 Then
		 FnameType=KS.C_C(ClassID,23)
		 Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
		 TemplateID=KS.C_C(ClassID,5)
		 WapTemplateID=KS.C_C(ClassID,22)
		 AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
	End If
	If KS.ChkClng(KS.C_S(ChannelID,17))<>0 And Verific=0 Then Verific=1
	If ID<>0 and verific=1  Then
		If KS.ChkClng(KS.C_S(ChannelID,42))=2 Then Verific=1 Else Verific=0
	End If
	if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
	If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
	
	Content=KSUser.SaveBeyoundFile(Content)
	PhotoUrl=KSUser.SaveBeyoundFile(KS.S("PhotoUrl"))
	Call KSUser.CheckDiyField(FieldXML,false)
				
	If ClassID="" Then
		KS.Die "<script>$.dialog.tips('你没有选择" & KS.C_S(ChannelID,3) & "栏目!',1,'error.gif',function(){});</script>"
	 End IF
	If Title="" Then
		KS.Die "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "标题!',1,'error.gif',function(){});</script>"
	End IF
	If Content="" and KS.ChkClng(FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonuserform").text="1")=1 Then
		KS.Die "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "内容!',1,'error.gif',function(){});</script>"
	End IF
	Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer=" & KS.WithKorean() &"'" & KSUser.UserName & "' and ID=" & ID,Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Hits")=0
				  RSObj("TemplateID")  =TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")       = FName
				  RSObj("Rank")        = "★★★"
				  RSObj("Inputer")     = KSUser.UserName
				  RSObj("PostTable")   = LFCls.GetCommentTable()
				  RSObj("OrderID")     = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & ClassID &"'")(0))+1
				 End If
				  If isDate(KS.S("AddDate")) Then
				  RSObj("Adddate")=KS.S("AddDate")
				  ElseIf ID=0 Then
				  RSObj("Adddate")=Now
				  End If
				  RSObj("ModifyDate")=Now
				  RSObj("Title")=Title
				  RSObj("FullTitle")=FullTitle
				  RSObj("Tid")=ClassID
				  RSObj("oTid")=KS.S("oTid")
				  RSObj("oId")=KS.ChkClng(KS.S("OID"))
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Origin")=Origin
				  RSObj("Changes") = Changes
				  RSObj("ArticleContent")=Content
				  RSObj("Verific")=Verific
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("Intro")=Intro
				  RSObj("DelTF")=0
				  RSObj("Comment")=1
                  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonuserform").text="1" Then
				  RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				  RSObj("Province")=Province
				  RSObj("City")=City
				  RSObj("County")=County				  
				  if PhotoUrl<>"" Then 
				   RSObj("PicNews")=1
				  Else
				   RSObj("PicNews")=0
				  End if
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonuserform").text="1" Then	RSObj("MapMarker")=KS.S("MapMark")
				  Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
				  
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				Dim AddDate:AddDate=RSObj("AddDate")
				If Left(Ucase(Fname),2)="ID" And KS.ChkClng(KS.S("ID"))=0 Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				Fname=RSOBj("Fname")
				If ID=0 Then
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Intro,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
                End If
				If Verific=1 Then 
				    Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
					Call KSUser.RefreshHtml(RSObj,ChannelID)
				End If
				 RSObj.Close:Set RSObj=Nothing
				 
				If Not KS.IsNul(FileIds) Then 
				 Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & InfoID &",classID=" & KS.C_C(ClassID,9) & " Where ID In (" & FileIds & ")")
				End If

				 
               If ID=0 Then
                  Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl ,0)
				  Dim LogStr
				  If PhotoUrl<>"" Then
				   LogStr="[img]" & photourl & "[/img][br]" & left(KS.LoseHtml(Content),60) & "..."
				  Else
				   LogStr=left(KS.LoseHtml(Content),80) & "..."
				  End If
			    Call KSUser.AddToWeibo(KSUser.UserName,"发表了" & KS.C_S(ChannelID,3) & "：" & left(Title,40) & " [url=" & KS.GetItemURL(ChannelID,ClassID,InfoID,Fname,AddDate) & "]详情&raquo;[/url][br]"&logstr,5)
				
				  '短信通知管理员，有新稿件
				  If Not KS.IsNul(AdminMobile) and KS.Setting(157)<>"0" Then 
				     Dim SmsContent:SmsContent=KS.ReadFromFile(KS.Setting(3) & "config/smscontentTips.txt")
					 Dim MobileArr:MobileArr=Split(AdminMobile,",")
					 Dim mm
					 SmsContent=replace(SmsContent,"{$classname}",KS.C_C(ClassID,1))
					 SmsContent=replace(SmsContent,"{$title}",title)
					 For mm=0 To Ubound(MobileArr)
					   Call KS.SendMobileMsg(MobileArr(mm),SmsContent)
					 Next
				  End If
				
				 KS.Echo "<script>$.dialog.confirm('" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?',function(){top.location.href='user_post.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';},function(){top.location.href='User_ItemInfo.asp?ChannelID=" & ChannelID & "';});</script>"
			   Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Intro,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl ,1)
				 KS.Echo "<script>$.dialog.tips('" & KS.C_S(ChannelID,3) & "修改成功!',1,'success.gif',function(){top.location.href='User_ItemInfo.asp?channelid=" & channelid & "';});</script>"
			   End If
  End Sub
  
  Sub DoSavePhoto()
                Dim AdminMobile
                Dim ClassID:ClassID=KS.S("ClassID")
				If KS.ChkClng(KS.C_C(ClassID,20))=0 Then KS.Die "<script>$.dialog.tips('对不起,系统设定不能在此栏目发表,请选择其它栏目!'',1,'error.gif',function(){});</script>"
				Dim Title:Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
				Dim KeyWords:KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				Dim Author:Author=KS.LoseHtml(KS.S("Author"))
				Dim Origin:Origin=KS.LoseHtml(KS.S("Origin"))
				Dim ShowStyle:ShowStyle=KS.ChkClng(KS.S("ShowStyle"))
				Dim PageNum:PageNum=KS.ChkClng(KS.S("PageNum"))
				Dim Content
				Content = KS.FilterIllegalChar(Request.Form("Content"))
				Content=KS.ClearBadChr(content)
				If Content="" Then content=" "
				Content=KSUser.SaveBeyoundFile(Content)
				Dim Province:Province= KS.LoseHtml(KS.S("Province"))
	            Dim City:City    = KS.LoseHtml(KS.S("City"))
				Dim County:County=KS.LoseHtml(KS.S("County"))

				Dim Verific:Verific=KS.ChkClng(KS.S("Status"))
				Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
				Dim PicUrls:PicUrls=KS.S("PicUrls")
				 If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 Then
				  If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If
                 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
				 Call KSUser.CheckDiyField(FieldXML,false)
				  Dim RSObj
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>$.dialog.tips('你没有选择" & KS.C_S(ChannelID,3) & "栏目!',1,'error.gif',function(){});</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "名称!',1,'error.gif',function(){});</script>"
				    Exit Sub
				  End IF
	              If PicUrls="" Then
				    Response.Write "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "!',1,'error.gif',function(){});</script>"
				    Exit Sub
				  End IF
				 If KS.ChkClng(KS.S("autothumb"))=1 And KS.IsNul(PhotoUrl) Then  PhotoUrl=Split(Split(PicUrls,"|||")(0),"|")(2)
	              If PhotoUrl="" Then
				    Response.Write "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "缩略图!',1,'error.gif',function(){});</script>"
				    Exit Sub
				  End IF
				Dim Fname,FnameType,TemplateID,WapTemplateID
				If KS.ChkClng(KS.S("ID"))=0 Then
				 FnameType=KS.C_C(ClassID,23)
				 Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
				 TemplateID=KS.C_C(ClassID,5)
				 WapTemplateID=KS.C_C(ClassID,22)
				 AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
			    End If
				
				If KS.ChkClng(KS.Setting(92))=1 Then  '远程存图
				    Dim SaveFilePath:SaveFilePath = KS.Setting(3) & KS.Setting(91)& "user/" & KSUser.GetUserInfo("userid") & "/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/" 
					KS.CreateListFolder (SaveFilePath)
				   Dim sPicUrlArr:sPicUrlArr=Split(PicUrls,"|||")
				   Dim i,sTemp,Url1,thumburl,ThumbFileName
				   PicUrls=""
				   For I=0 To Ubound(sPicUrlArr)
				     If Left(Lcase(Split(sPicUrlArr(i),"|")(1)),4)="http" and instr(Lcase(Split(sPicUrlArr(i),"|")(1)),lcase(ks.setting(2)))=0 Then
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
				  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer=" & KS.WithKorean() &"'" & KSUser.UserName & "' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Inputer")=KSUser.UserName
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("PostTable")= LFCls.GetCommentTable()
				  RSObj("OrderID")     = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & ClassID &"'")(0))+1
				End If
				  If isDate(KS.S("AddDate")) Then
				  RSObj("Adddate")=KS.S("AddDate")
				  ElseIf KS.ChkClng(KS.S("ID"))=0 Then
				  RSObj("Adddate")=Now
				  End If
				  RSObj("ModifyDate")=Now
				  RSObj("Title")=Title
				  RSObj("Tid")=ClassID
				  RSObj("oTid")=KS.S("oTid")
				  RSObj("oId")=KS.ChkClng(KS.S("OID"))
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("PicUrls")=PicUrls
				  RSObj("PicNum") = Ubound(split(PicUrls,"|||"))+1
				  RSObj("Province")=Province
				  RSObj("City")=City
				  RSObj("County")=County
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Origin")=Origin
				  RSObj("ShowStyle")=ShowStyle
				  RSObj("PageNum")=PageNum
				  RSObj("PictureContent")=Content
				  RSObj("Verific")=Verific
				  RSObj("Comment")=1
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonuserform").text="1" Then	RSObj("MapMarker")=KS.S("MapMark")
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonuserform").text="1" Then
				   RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				   Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				Dim AddDate:AddDate=RSObj("AddDate")
				If Left(Ucase(Fname),2)="ID" And KS.ChkClng(KS.S("ID"))=0 Then
					   RSObj("Fname") = InfoID & FnameType
					   RSObj.Update
				 End If
				 Fname=RSOBj("Fname")
				 If KS.ChkClng(KS.S("ID"))=0 Then
				  Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
				 End If
				 If Verific=1 Then 
				    Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
					Call KSUser.RefreshHtml(RSObj,ChannelID)
				End If
				
				 RSObj.Close:Set RSObj=Nothing
				 If KS.ChkClng(KS.S("ID"))=0 Then
				 
				  Call KS.FileAssociation(ChannelID,InfoID,PicUrls & PhotoUrl & Content ,0)
				  Dim LogStr
				  If PhotoUrl<>"" Then
				   LogStr="[img]" & photourl & "[/img][br]" & left(KS.LoseHtml(Content),60) & "..."
				  Else
				   LogStr=left(KS.LoseHtml(Content),80) & "..."
				  End If
			    Call KSUser.AddToWeibo(KSUser.UserName,"上传了" & KS.C_S(ChannelID,3) & "：" & left(Title,40) & " [url=" & KS.GetItemURL(ChannelID,ClassID,InfoID,Fname,AddDate) & "]详情&raquo;[/url][br]"&logstr,5)
				
				
				'短信通知管理员，有新稿件
				  If Not KS.IsNul(AdminMobile) and KS.Setting(157)<>"0" Then 
				     Dim SmsContent:SmsContent=KS.ReadFromFile(KS.Setting(3) & "config/smscontentTips.txt")
					 Dim MobileArr:MobileArr=Split(AdminMobile,",")
					 Dim mm
					 SmsContent=replace(SmsContent,"{$classname}",KS.C_C(ClassID,1))
					 SmsContent=replace(SmsContent,"{$title}",title)
					 For mm=0 To Ubound(MobileArr)
					   Call KS.SendMobileMsg(MobileArr(mm),SmsContent)
					 Next
				  End If
				
				  KS.Echo "<script>$.dialog.confirm('" & KS.C_S(ChannelID,3) & "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?',function(){top.location.href='user_post.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';},function(){top.location.href='User_ItemInfo.asp?ChannelID=" & ChannelID &"';});</script>"
				Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PicUrls & PhotoUrl & Content ,1)
				 KS.Echo "<script>$.dialog.tips('" & KS.C_S(ChannelID,3) & "修改成功!',1,'success.gif',function(){top.location.href='User_ItemInfo.asp?ChannelID=" & ChannelID &"';});</script>"
				End If
  End Sub
  Sub DoSaveDownLoad()
     Dim SizeUnit,ClassID,Title,KeyWords,Author,DownLB,DownYY,DownSQ,DownSize,DownPT,YSDZ,ZCDZ,JYMM,Origin,Content,Verific,PhotoUrl,BigPhoto,DownUrls,RSObj,ID,AddDate,ComeUrl,CurrentOpStr,Action,I

    ID=KS.ChkClng(KS.S("ID"))
	ClassID=KS.S("ClassID")
	If KS.ChkClng(KS.C_C(ClassID,20))=0 Then
		KS.Die "<script>$.dialog.tips('对不起,系统设定不能在此栏目发表,请选择其它栏目!',1,'error.gif',function(){});</script>"
	End IF
				  Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
				  KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				  Author=KS.LoseHtml(KS.S("Author"))
				  DownLB=KS.LoseHtml(KS.S("DownLB"))
				  DownYY=KS.LoseHtml(KS.S("DownYY"))
				  DownSQ=KS.LoseHtml(KS.S("DownSQ"))
				  DownSize=KS.S("DownSize")
				  If DownSize = "" Or Not IsNumeric(DownSize) Then DownSize = 0
		          DownSize = DownSize & KS.S("SizeUnit")
				  DownPT=KS.LoseHtml(KS.S("DownPT"))
				  YSDZ=KS.LoseHtml(KS.S("YSDZ"))
				  ZCDZ=KS.LoseHtml(KS.S("ZCDZ"))
				  JYMM=KS.LoseHtml(KS.S("JYMM"))
				  Origin=KS.LoseHtml(KS.S("Origin"))
				  Content = KS.FilterIllegalChar(Request.Form("Content"))
				  If Content="" Then Content=" "
				  Content=KSUser.SaveBeyoundFile(Content)
				  Content=KS.ClearBadChr(content)
				  Verific=KS.ChkClng(KS.S("Status"))
				  If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 and verific=1  Then
					 If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If
				 if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
				 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
				  PhotoUrl=KS.LoseHtml(KS.S("PhotoUrl"))
				  BigPhoto=KS.LoseHtml(KS.S("BigPhoto"))
				  DownUrls=KS.S("DownUrls")
				  if (Instr(lcase(DownUrls),lcase(KS.Setting(2)))<>0 and Instr(lcase(DownUrls),"user/" & KSUser.GetUserInfo("userid") &"/")=0) or (Instr(lcase(DownUrls),"http://")=0 and Instr(lcase(DownUrls),"user/" & KSUser.GetUserInfo("userid") &"/")=0) or Instr(lcase(DownUrls),".asp")<>0 or KS.IsNul(Request.Form("DownUrls")) then
				   KS.Die "<script>$.dialog.tips('软件地址格式不正确，如果审核通过的软件要修改，请重新上传软件!',2,'error.gif',function(){});</script>"
				  end if
				  
				  
				  PhotoUrl=KSUser.SaveBeyoundFile(PhotoUrl)
				  BigPhoto=KSUser.SaveBeyoundFile(BigPhoto)
				  
				Call KSUser.CheckDiyField(FieldXML,false)		  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then KS.Die "<script>$.dialog.tips('你没有选择" & KS.C_S(ChannelID,3) & "栏目!',1,'error.gif',function(){});</script>"
				  If Title="" Then  KS.Die "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "名称!',1,'error.gif',function(){});</script>"
	              If DownUrls="" Then KS.Die "<script>$.dialog.tips('你没有输入" & KS.C_S(ChannelID,3) & "!',1,'error.gif',function(){});</script>"
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				    
				 Dim Fname,FnameType,TemplateID,WapTemplateID,AdminMobile
				 If ID=0 Then
					 FnameType=KS.C_C(ClassID,23)
					 Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
					 TemplateID=KS.C_C(ClassID,5)
					 WapTemplateID=KS.C_C(ClassID,22)
					 AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
				End If	 
					RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer=" & KS.WithKorean() &"'" & ksuser.username & "' and ID=" & ID,Conn,1,3
					If RSObj.Eof Then
						  RSObj.AddNew
						  RSObj("Inputer")=KSUser.UserName
						  RSObj("Hits")=0
						  RSObj("TemplateID")=TemplateID
						  RSObj("WapTemplateID")=WapTemplateID
						  RSObj("Fname")=FName
						  RSObj("Rank")="★★★"
						  RSObj("PostTable")= LFCls.GetCommentTable()
						  RSObj("OrderID")     = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & ClassID &"'")(0))+1
					End If
					  If isDate(KS.S("AddDate")) Then
					  RSObj("Adddate")=KS.S("AddDate")
					  ElseIf ID=0 Then
					  RSObj("Adddate")=Now
					  End If
					  RSObj("ModifyDate")=Now
					  RSObj("Title")=Title
					  RSObj("TID")=ClassID
					  RSObj("oTid")=KS.S("oTid")
					  RSObj("oId")=KS.ChkClng(KS.S("OID"))
					  RSObj("KeyWords")=KeyWords
					  RSObj("Author")=Author
					  RSObj("DownLB")=DownLB
					  RSObj("DownYY")=DownYY
					  RSObj("DownSQ")=DownSQ
					  RSObj("DownSize")=DownSize
					  RSObj("DownPT")=DownPT
					  RSObj("YSDZ")=YSDZ
					  RSObj("ZCDZ")=ZCDZ
					  RSObj("JYMM")=JYMM
					  RSObj("Origin")=Origin
					  RSObj("DownContent")=Content
					  RSObj("PhotoUrl")=PhotoUrl
					  RSObj("BigPhoto")=BigPhoto
					  RSObj("DownUrls")="0|下载地址|" & DownUrls
					  RSObj("Verific")=Verific
					   If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonuserform").text="1" Then
					   RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
					  End If
					  Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
					  RSObj.Update
					  RSObj.MoveLast
						Dim InfoID:InfoID=RSObj("ID")
						AddDate=RSObj("AddDate")
						If Left(Ucase(Fname),2)="ID" Then
							RSObj("Fname") = InfoID & FnameType
							RSObj.Update
						End If
						Fname=RSOBj("Fname")
						If ID=0 Then
			         	 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
						End If
						
						If Verific=1 Then 
							Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
							Call KSUser.RefreshHtml(RSObj,ChannelID)
						End If
						 RSObj.Close:Set RSObj=Nothing
				 
			 If ID=0 Then
				 Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & Content & DownUrls & BigPhoto ,0)
			    Dim LogStr
				  If PhotoUrl<>"" Then
				   LogStr="[img]" & photourl & "[/img][br]" & left(KS.LoseHtml(Content),60) & "..."
				  Else
				   LogStr=left(KS.LoseHtml(Content),80) & "..."
				  End If
			    Call KSUser.AddToWeibo(KSUser.UserName,"上传了" & KS.C_S(ChannelID,3) & "：" & left(Title,40) & " [url=" & KS.GetItemURL(ChannelID,ClassID,InfoID,Fname,AddDate) & "]详情&raquo;[/url][br]"&logstr,5)
				
				
				'短信通知管理员，有新稿件
				  If Not KS.IsNul(AdminMobile) and KS.Setting(157)<>"0" Then 
				     Dim SmsContent:SmsContent=KS.ReadFromFile(KS.Setting(3) & "config/smscontentTips.txt")
					 Dim MobileArr:MobileArr=Split(AdminMobile,",")
					 Dim mm
					 SmsContent=replace(SmsContent,"{$classname}",KS.C_C(ClassID,1))
					 SmsContent=replace(SmsContent,"{$title}",title)
					 For mm=0 To Ubound(MobileArr)
					   Call KS.SendMobileMsg(MobileArr(mm),SmsContent)
					 Next
				  End If
				
			     KS.Echo "<script>$.dialog.confirm('" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?',function(){top.location.href='user_post.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';},function(){top.location.href='User_ItemInfo.asp?ChannelID=" & ChannelID & "';});</script>"
			 Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & Content & DownUrls ,1)
				 KS.Echo "<script>$.dialog.tips('" & KS.C_S(ChannelID,3) & "修改成功!',1,'success.gif',function(){top.location.href='User_ItemInfo.asp?ChannelID=" & ChannelID &"';});</script>"
		    End If
  End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Contributor
KSCls.loadKesion()
Set KSCls = Nothing

Class Contributor
        Private KS,KSUser,ChannelID,ClassID,LoginTF,Qid,Action,Selbutton,rs,ShowStyle,PageNum,Price,TypeId,ValidDate,GQContent
		Private FieldXML,FieldNode,FNode,FieldDictionary
		Private DownLBList, DownYYList, DownSQList, DownPTList
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
		Public Sub loadKesion()
		 Dim FileContent,MainUrl,RequestItem,TemplateFile
		 Dim KSR,ParaList
		 FCls.RefreshType = "Member"   '设置当前位置为会员中心
		 Set KSR = New Refresh
		 
		 Dim TemplatePath:TemplatePath=KS.Setting(3) & KS.Setting(90) & "Common/Contributor.html"  '模板地址
		 FileContent = KSR.LoadTemplate(TemplatePath)    
		 If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		  FileContent = KSR.KSLabelReplaceAll(FileContent)
		 Set KSR = Nothing
		 ScanTemplate FileContent
        End Sub	
		
		
		Public Sub loadMain()
		LoginTF=KSUser.UserLoginChecked
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		ClassID=KS.S("ClassID")
		Response.Write EchoUeditorHead()
		  Dim Action:Action=KS.S("Action")
			Select Case Action
			 Case "Next" Call ContributorNext()
			 Case "AddSave" Call ContributorSave()
			 Case Else  Call Main()
			 End Select
	    End Sub 
		
		Function GetQuestionRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.GetCurrQuestion(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetQuestionRnd=RandNum
		End Function
		
		Function PubQuestion()
			if mid(KS.Setting(161),2,1)="1" then
			 Qid=GetQuestionRnd
			%>
						   <tr class="tdbg">
                            <td  height="25" align="center" width="100"><span>请回答问题：</span></td>
                             <td><font color="red"><%
							 Dim QuestionArr:QuestionArr=Split(KS.GetCurrQuestion(162),vbcrlf)
		                     response.write QuestionArr(Qid)
							 %></font>
							 　</td>
                          </tr>
						   <tr class="tdbg">
                            <td  height="25" align="center"><span>您的答案：</span></td>
                            <td><input type="text" class="textbox" id="QuestionAnswer" name="a<%=md5(Qid,16)%>">
							</td>
                          </tr>
			<%end if
		End Function
		
		
		'选择投稿栏目
		Sub Main()
		%>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="nmtg">
							  <script language="javascript">
							    function CheckForm()
								{
								 if (document.form1.classid.value=='')
								 {
								  alert('请选择投稿栏目!');
								  return false
								 }
								 return true;
								}
							  </script>
							   <form name="form1" action="?Action=Next" method="post" onSubmit="return(CheckForm());">
								<tr>
								  <td align="center">
								  <select name=classid size="22">
								  <%
								  Dim CacheID,K,SQL,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
								  RS.Open "Select ID,FolderName,a.ChannelID From KS_Class a inner join ks_channel b on a.channelid=b.channelid Where UserTF=2 and a.ChannelID<>5 and CommentTF=2 order by a.ChannelID,folderorder",Conn,1,1
								  If Not RS.Eof Then SQL=RS.GetRows(-1)
								  RS.Close:Set RS=Nothing
								  If IsArray(SQL) Then
								   For K=0 To Ubound(SQL,2)
									 If SQL(2,k)<>CacheID Then
									  Response.Write "<optgroup  label='===============" & KS.C_S(SQL(2,k),3) & "栏目=============='>"
									 End If
									 Response.Write "<option value='" & SQL(0,K) & "'>|-" & SQL(1,K) & "</option>"
									 
									 CacheID=SQL(2,K)
								   Next
								  End If
								  %>
								  </select>								 
								   </td>
								</tr>
								
								<tr class="tdbg">
								  
								  <td height="22" align="center">
								   <input type="submit" name="s1" value=" 下 一 步 " class="button">
								   </td>
								</tr>
								
								</form>
		</table>
<%
  End Sub
  
   '选择投稿界面
   Sub ContributorNext()
     ClassID=KS.R(KS.S("ClassID"))
	 If ClassID="" Then Response.Write "<script>alert('对不起，你没有选择投稿栏目!');history.back();</script>":Response.End
	 IF KS.C("UserName")="" and KS.C_C(ClassID,18)<>"2" then Call KS.ShowTips("error","对不起，本栏目不允许游客投稿!")

	 ChannelID=KS.ChkClng(Conn.Execute("Select top 1 ChannelID From KS_Class Where ID='" & ClassID & "'")(0))
	 If ChannelID=0 Then Response.End()
	 Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	 If LoginTF=True Then
		   Select Case KS.C_S(ChannelID,6)
		    Case 1 Response.Redirect "User_post.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 2 Response.Redirect "User_post.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 3 Response.Redirect "User_post.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 4 Response.Redirect "User_MyFlash.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 7 Response.Redirect "User_MyMovie.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 8 Response.Redirect "User_MySupply.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		   End Select
	 End If
	 
	 Select Case KS.C_S(ChannelID,6) 
	   Case 1:Call AddByArticle()
	   Case 2:Call AddByPicture()
	   Case 3:Call AddBySoftWare()
	   Case 4:Call AddByFlash()
	   Case 7:Call AddByMovie()
	   Case 8:Call AddBySupply()
	   Case Else:Response.Write "参数出错!":Response.End()
	 End Select 
   End Sub
   
   '保存投稿
   Sub ContributorSave()
     ChannelID=KS.ChkCLng(KS.S("ChannelID"))
	  If ChannelID=0 Then Response.End()
	  Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	  IF lcase(Trim(Request.Form("Verifycode")))<>lcase(Trim(Session("Verifycode"))) then 
	   Call KS.AlertHistory("验证码有误，请重新输入！",-1)
	   exit Sub
	  End If
	  If Request.ServerVariables("HTTP_REFERER")="" Then
	   Call KS.AlertHistory("非法提交！",-1)
	   exit Sub
	  End If
	  '检查注册回答问题
	  Dim CanReg,N
	   If Mid(KS.Setting(161),2,1)="1" Then
		     CanReg=false
		     For N=0 To Ubound(Split(KS.GetCurrQuestion(162),vbcrlf))
			   If Trim(Request.Form("a" & MD5(n,16)))<>"" Then
			      If Lcase(Request.Form("a" & MD5(n,16)))<>Lcase(Split(KS.GetCurrQuestion(163),vbcrlf)(n)) Then
			       Call KS.AlertHistory("对不起,注册问题的回答不正确!",-1) : Response.End
				   CanReg=false
				  Else
				   CanReg=True
				  End If
			   End If
			 Next
			 If CanReg=false Then Call KS.AlertHistory("对不起,注册答案不能为空!",-1) : Response.End
	  End If
	 SelButton="选择栏目..." 
     Select Case KS.ChkClng(KS.C_S(ChannelID,6))
	  Case 1:Call SaveByArticle()
	  Case 2:Call SaveByPhoto()
	  Case 3:Call SaveByDownLoad()
	  Case 4:Call SaveByFlash()
	  Case 7:Call SaveByMovie()
	  Case 8:Call SaveBySupply()
	 End Select	 
   End Sub
   
   '添加文章
   Sub AddByArticle()
     
	%>
	<script language = "JavaScript">
	      var box='';
		   function addMap(){ box=$.dialog({title:'电子地图标注',content:'url:../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),width:820,height:430});}
		   function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{alert('对不起,请先输入标题!'); }
			}
		    function CheckClassID(){
				if (document.myform.ClassID.value=="0" || document.myform.ClassID.value=='') {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;}		
				  return true;
			}
			function insertHTMLToEditor(codeStr){  editor.execCommand('insertHtml', codeStr);} 
			function CheckForm(){
				<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
				if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>标题！");
					document.myform.Title.focus();
					return false;
				  }	
				<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonuserform").text="1" Then%>
				    if (editor.hasContents()==false)
					{
					  alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！");
					  editor.focus();
					  return false;
					}
				<%end if%>
				if (document.myform.Verifycode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.Verifycode.focus();
					return false;
				  }	
				<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>
				
				 return true; }
		</script>
		<iframe src="about:blank" name="hidframe" style="display:none"></iframe>
		<br/><form  action="?channelid=<%=channelid%>&Action=AddSave" method="post" target="hidframe" name="myform" id="myform" onSubmit="return CheckForm();">
		<%
		Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
		If KS.IsNul(XmlForm) Then 
		 GetInputForm false,ChannelID,FieldXML,FieldNode,FieldDictionary,KS.ChkClng(KS.S("id")),KSUser,rs
		Else
		   Scan XmlForm
		End If
		%>
		</form>		
		  <%
    End Sub
	
	'添加图片
	Sub AddByPicture()
		ShowStyle=3
		%>
		   <style type="text/css">
			#thumbnails{background:url(../plus/swfupload/images/albviewbg.gif) no-repeat;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
			#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:152px;height:155px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
			#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px}

			</style>
			<link href="../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
			<script type="text/javascript" src="../plus/swfupload/swfupload/swfupload.js"></script>
			<script type="text/javascript" src="../plus/swfupload/js/handlers.js"></script>
<script type="text/javascript">
       var box='';
       function addMap(){
		box=$.dialog({title:'电子地图标注',content:'url:../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),width:820,height:430});
		}
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
				post_params: {"BasicType":<%=KS.C_S(ChannelID,6)%>,"ChannelID":<%=ChannelID%>,"UserName" : "<%=KS.C("UserName") %>","RndPassWord":"<%=KS.C("RndPassWord")%>","AutoRename":4},

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
				button_width: 152,
				button_height: 22,
				button_text : '<span class="button">批量上传(单图限制2M)</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#666666;font-size: 14px; } ',
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

	 box=$.dialog({title:"网上采集图片",content:"<div style='padding:3px'>带http://开头的远程图片地址,每行一张图片地址:<br/><textarea id='collecthttp' name='collecthttp' style='width:400px;height:150px'></textarea></div>",init: function(){
			          input = this.DOM.content[0].getElementsByTagName('textarea')[0];
						},ok: function(){ 
						  ProcessCollect(input.value);
							return false; 
						}, 
						cancelVal: '关闭', 
			cancel: true });

	}
	function AddTJ(){
	     box=$.dialog({title:"从上传文件中选择",content:"<div style='padding:3px'><strong>小图地址:</strong><input class='textbox' type='text' name='x1' id='x1'> <br/><strong>大图地址:</strong><input class='textbox' type='text' name='x2' id='x2'><br/><strong>简要介绍:</strong><input class='textbox' type='text' name='x3' id='x3'></div>",init: function(){
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
					  addImage(x1,x2,x3,"")
					   box.close();
	}
	function ProcessCollect(collecthttp){
	 if (collecthttp==''){
	   alert('请输入远程图片地址,一行一张地址!');
	   return false;
	 }
	 var carr=collecthttp.split('\n');
	 for(var i=0;i<carr.length;i++){
	   
	   var bigsrc=carr[0];
	   var smallsrc=carr[0];
	   addImage(bigsrc,smallsrc,'')
	 }
	 box.close();
	}
	</script>
		<br/>
		<iframe src="about:blank" name="hidframe" style="display:none"></iframe>
		<br/><form  action="?channelid=<%=channelid%>&Action=AddSave" method="post" target="hidframe" name="myform" id="myform" onSubmit="return CheckForm();">
		<%
		Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
		If KS.IsNul(XmlForm) Then 
		 GetInputForm false,ChannelID,FieldXML,FieldNode,FieldDictionary,KS.ChkClng(KS.S("id")),KSUser,rs
		Else
		   Scan XmlForm
		End If
		%>

		 <script type="text/javascript">

			function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{
			   alert('对不起,请先输入标题!');
			  }
			}
				function CheckForm()
				{
				if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					document.myform.Title.focus();
					return false;
				  }		
				if (document.myform.PhotoUrl.value=='' && $('#autothumb').attr('checked')==false)
				{
					alert("请输入<%=KS.C_S(ChannelID,3)%>缩略图！");
					document.myform.PhotoUrl.focus();
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
				  alert('请输入<%=KS.C_S(ChannelID,3)%>内容!');
				  $('input[name=imgurl1]').focus();
				  return false;
				}
				if (document.myform.Verifycode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.Verifycode.focus();
					return false;
				  }	
				<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>
				
				  return true;
                    
				}
				function CheckClassID()
				{
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
			</script>
			 <%
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
	Sub AddBySoftWare()
		Call InitialDown("","","")
%>
				
				<script language="javascript">
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
				function SetPhotoUrl()
				{
				 if (document.myform.DownUrl.value!='')
				  document.myform.PhotoUrl.value=document.myform.DownUrl.value.split('|')[1];	
				}
				function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{alert('对不起,请先输入标题!'); }
			}
               function CheckClassID()
				{
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
				function CheckForm()
				{   
					
				 if (document.myform.Title.value==""){
						alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
						document.myform.Title.focus();
						return false;
					  }
					if (document.myform.DownUrlS.value==''){
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.DownUrlS.focus();
						return false;
					}
				<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
                 <%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>				
				 if (document.myform.Verifycode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.Verifycode.focus();
					return false;
				  }	
					document.myform.submit();
					return true;
				}
				</script>
		<iframe src="about:blank" name="hidframe" style="display:none"></iframe>
		<br/><form  action="?channelid=<%=channelid%>&Action=AddSave" method="post" target="hidframe" name="myform" id="myform" onSubmit="return CheckForm();">
		<%
		Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
		If KS.IsNul(XmlForm) Then 
		 GetInputForm false,ChannelID,FieldXML,FieldNode,FieldDictionary,KS.ChkClng(KS.S("id")),KSUser,rs
		Else
		   Scan XmlForm
		End If
		%>
		</form>	
   <%
	End Sub
	
	'添加动漫
	Sub AddByFlash()
	%>
			<script language = "JavaScript">
			function CheckClassID(){
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
			}
			function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{alert('对不起,请先输入标题!'); }
			}
			function CheckForm(){
				if (document.myform.Title.value==""){
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					document.myform.Title.focus();
					return false;
				  }		
				if (document.myform.FlashUrl.value==''){
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.FlashUrl.focus();
						return false;
					}
				<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
				<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>		
				if (document.myform.Verifycode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.Verifycode.focus();
					return false;
				  }	
				  document.myform.submit();
				 return true;  
				}
				</script>
               <br/>
			   <iframe src="about:blank" name="hidframe" style="display:none"></iframe>
			 <form action="?ChannelID=<%=ChannelID%>&Action=AddSave"  target="hidframe" method="post" name="myform" id="myform">
			   <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					  <tr class="title">
							   <td colspan=2 align=center><%="发布" & KS.C_S(ChannelID,3)%></td>
					  </tr> 
					<tr class="tdbg">
						<td height="25" align="center">所属栏目：</td>
						<td>[<%=KS.GetClassNP(ClassID)%>] <a href="Contributor.asp"><<重新选择>></a><input type="hidden" name="ClassID" value="<%=classid%>"></td>
					</tr>
                    <tr class="tdbg">
                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>名称：</span></td>
                        <td><input name="Title" class="textbox" type="text" id="Title" style="width:250px; " maxlength="100" /><span style="color: #FF0000">*</span></td>
                   </tr>
                   <tr class="tdbg">
                        <td height="25" align="center"><span>关 键 字：</span></td>
                        <td><input class="textbox" name="KeyWords" type="text" id="KeyWords" style="width:250px; " /> 
                                    <a href="javascript:void(0)" onClick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span></td>
                   </tr>
                   <tr class="tdbg">
                         <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>作者：</span></td>
                         <td height="25"><input name="Author" class="textbox" type="text" style="width:250px; "  maxlength="30" /></td>
                   </tr>
                   <tr class="tdbg">
                     <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>来源：</span></td>
                     <td><input name="Origin" class="textbox" type="text" id="Origin" style="width:250px; " maxlength="100" /></td>
				  </tr>
<%
	If IsObject(FieldNode) Then
		For Each FNode In FieldNode
				If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
					Response.Write KSUser.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary)
				End If
		Next
	End If
%>      
							  
								<tr class="tdbg">
                                        <td height="25" align="center"><span>缩 略 图：</span></td>
                                        <td><input class="textbox" name='PhotoUrl' type='text' style="width:250px;" id='PhotoUrl' maxlength="100" /><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=4&Type=Pic' frameborder=0 scrolling=no width='300' height='30'> </iframe>
                                         </td>
							   </tr>
								
								<tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>地址：</span></td>
                                        <td>
										 <table cellspacing="0" cellpadding="0" width="100%">
										  <tr><td width="270"><input class="textbox" name='FlashUrl'  type='text' style="width:250px;" id='FlashUrl' maxlength="100" /> <font color='#FF0000'>*</font></td>
										  <td><iframe id='UpFlashFrame' name='UpFlashFrame' src='User_Upfile.asp?type=UpByBar&channelid=4' frameborder=0 scrolling=no width='300' height='30'> </iframe></td>
										  </tr>
										  </table>
                                          </td>
							   </tr>
								
  								<tr class="tdbg">
                                        <td align="center"><span><%=KS.C_S(ChannelID,3)%>简介：<br />
                                          </span></td>
                                        <td>
										 <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td height="150">
                                                    <%
                                     Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:70%;height:200px;""></script>"
                                     Response.Write "<script>setTimeout(""var editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:200 });"",10);</script>"
                                    %>
											
													</td>
                                                  </tr>
                                              </table>
									    </td>
                                </tr>
						<%
								call PubQuestion
						%>
							  <tr class="tdbg">
								<td  height="25" align="center"><span>验证码：</span></td>
								 <td>
								 <script type="text/javascript">writeVerifyCode('<%=KS.Setting(3)%>',1,'textbox')</script>
								</td>
							  </tr>
                          <tr class="tdbg">
                            <td align="center" colspan=2><input class="button" type="button" onClick="return CheckForm();" name="Submit" value=" OK! 发布 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重来 " /></td>
                          </tr>
                      </table>
				</form>
		  <%
	End Sub
	
	Sub AddByMovie()
%>
	  <script language = "JavaScript">
	       function CheckClassID(){
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
			  function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{alert('对不起,请先输入标题!'); }
			}
				function CheckForm()
				{
				if (document.myform.Title.value==""){
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					document.myform.Title.focus();
					return false;
				  }		
				if (document.myform.MovieUrl.value==''){
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.MovieUrl.focus();
						return false;
					}
				<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
					<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>	
				if (document.myform.Verifycode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.Verifycode.focus();
					return false;
				  }	
				  document.myform.submit();
				 return true;  
				}
		</script>
           <br/>
	<iframe src="about:blank" name="hidframe" style="display:none"></iframe>
	 <form  action="?ChannelID=<%=ChannelID%>&Action=AddSave" method="post" target="hidframe" name="myform" id="myform">
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
				 <tr class="title">
					<td colspan=2 align=center><%="发布" & KS.C_S(ChannelID,3)%></td>
				 </tr> 
				 <tr class="tdbg">
							<td height="25" align="center">所属栏目：</td>
							<td><%=KS.GetClassNP(ClassID)%>] <a href="Contributor.asp"><<重新选择>></a><input type="hidden" name="ClassID" value="<%=classid%>"></td>
				</tr>
                <tr class="tdbg">
                       <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>名称：</span></td>
                       <td><input name="Title" class="textbox" type="text" id="Title" style="width:250px; " maxlength="100" /><span style="color: #FF0000">*</span></td>
                </tr>
                <tr class="tdbg">
                  <td height="25" align="center"><span>关键字Tags：</span></td>
                  <td><input class="textbox" name="KeyWords" type="text" id="KeyWords" style="width:250px; " /> 
                                    <a href="javascript:void(0)" onClick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span>主要演员：</span></td>
                                        <td height="25"><input name="MovieAct" class="textbox" type="text" id="MovieAct" style="width:250px; "  maxlength="30" /></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>导演：</span></td>
                                        <td><input name="MovieDY" class="textbox" type="text" id="MovieDY" style="width:250px; " maxlength="100" /></td>
								</tr>
<%
	If IsObject(FieldNode) Then
		For Each FNode In FieldNode
				If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
					Response.Write KSUser.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary)
				End If
		Next
	End If
%>     
								<tr class="tdbg">
                                        <td height="25" align="center"><span>缩 略 图：</span></td>
                                        <td><input class="textbox" name='PhotoUrl' type='text' style="width:250px;" id='PhotoUrl' maxlength="100" /><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=7&Type=Pic' frameborder=0 scrolling=no width='300' height='30'> </iframe>
                                        </td>
							   </tr>
								<tr class="tdbg">
                                  <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>地址：</span></td>
                                  <td><input class="textbox" name='MovieUrl' type='text' style="width:250px;" id='MovieUrl' maxlength="100" /> <font color=red>*</font> <iframe id='UpFlashFrame' name='UpFlashFrame' src='User_Upfile.asp?type=UpByBar&channelid=7' frameborder=0 scrolling=no width='300' height='30'> </iframe>
                                          </td>
							   </tr>
  								<tr class="tdbg">
                                        <td align="center"><span><%=KS.C_S(ChannelID,3)%>简介：</span></td>
                                        <td><%
                                     Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:70%;height:150px;""></script>"
                                     Response.Write "<script>setTimeout(""var editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:150 });"",10);</script>"
                                    %></td>
                                </tr>
								<%
								call PubQuestion
						      %>
							  <tr class="tdbg">
								<td  height="25" align="center"><span>验证码：</span></td>
								 <td>
								 <script type="text/javascript">writeVerifyCode('<%=KS.Setting(3)%>',1,'textbox')</script>
								</td>
							  </tr>
                          <tr class="tdbg">
                            <td align="center" colspan=2><input class="button" type="button" onClick="return CheckForm();" name="Submit" value=" OK! 发布 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重来 " /></td>
                          </tr>
                     </table>
                  </form>
		  <%
	End Sub
	
	'添加供求信息
	Sub AddBySupply()
	%>
	<SCRIPT language=JavaScript>
	function CheckClassID(){
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
			  function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{alert('对不起,请先输入标题!'); }
			}
var partten = "/^\d{8}$/"
function CheckForm()
{
if ($("#Title").val()=='')
{
alert("请输入信息标题!");
$("#Title").focus();
return false; 
}
if (document.myform.Price.value=="")
{
alert("价格说明不能为空");
document.myform.Price.focus();
document.myform.Price.select();
return false; 
}
if (document.myform.TypeID.value =="") 
{ 
alert("请选择交易类别！"); 
document.myform.TypeID.focus(); 
return false; 
}

if (editor.hasContents()==false)
{
	alert("信息内容必须输入");
	editor.focus();
	return false; 
}
if (document.myform.ContactMan.value=="")
{
alert("联系人不能为空");
document.myform.ContactMan.focus();
return false; 
}
if (document.myform.Mobile.value=="")
{
alert("联系电话不能为空");
document.myform.Mobile.focus();
return false; 
}
<%Call LFCls.ShowDiyFieldCheck(FieldXML,0)%>
<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
 <%end if%>
if (document.myform.Verifycode.value=="")
{
	alert("请输入验证码！");
	document.myform.Verifycode.focus();
     return false;
}	
document.myform.submit();
}
</SCRIPT>
<body leftMargin="0" topMargin="0" marginheight="0">
<iframe src="about:blank" name="hidframe" style="display:none"></iframe>

<form  action="?channelid=<%=channelid%>&Action=AddSave" method="post" target="hidframe" name="myform" id="myform" onSubmit="return CheckForm();">
		<%
		Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
		If KS.IsNul(XmlForm) Then 
		 GetInputForm false,ChannelID,FieldXML,FieldNode,FieldDictionary,KS.ChkClng(KS.S("id")),KSUser,rs
		Else
		   Scan XmlForm
		End If
		%>
		</form>	



  <%
	End Sub
	
	'保存文章
	Sub SaveByArticle
	  Dim Title,FullTitle,KeyWords,Author,Origin,Intro,Content,Verific,PicUrl,Action,I,Province,City,County

          ClassID  = KS.S("ClassID")
		  Title    = KS.LoseHtml(KS.S("Title"))
		  KeyWords = KS.LoseHtml(KS.S("KeyWords"))
		  Author   = KS.LoseHtml(KS.S("Author"))
		  Province = KS.LoseHtml(KS.S("Province"))
	      City     = KS.LoseHtml(KS.S("City"))
	      County   = KS.LoseHtml(KS.S("County"))
  		  Origin   = KS.LoseHtml(KS.S("Origin"))
		  Content  = Request.Form("Content")
		  Content  = KS.ClearBadChr(content)
		  if Content="" Then Content="&nbsp;"
		  Verific  = KS.S("Status")
		  Intro    = KS.LoseHtml(KS.S("Intro"))
		  FullTitle= KS.LoseHtml(KS.S("FullTitle"))
		  if Intro="" And KS.ChkClng(KS.S("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(Content),200)
				 
		  PicUrl=KS.S("PicUrl")
		  
		  Call KSUser.CheckDiyField(FieldXML,false)	 
		
				 
			 if ClassID="" Then ClassID=0
			 If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "标题!');</script>"
				    Exit Sub
				  End IF
				  If Content="" and FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonuserform").text="1" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "内容!');</script>"
				    Exit Sub
				  End IF
		 
			  Dim Fname,FnameType,TemplateID,WapTemplateID
			  FnameType=KS.C_C(ClassID,23)
			  Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
			  TemplateID=KS.C_C(ClassID,5)
			  WapTemplateID=KS.C_C(ClassID,22)
			  Dim AdminMobile:AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
			  
				  
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2),Conn,1,3
				RSObj.AddNew
				  RSObj("Title")=Title
				  RSObj("FullTitle")=FullTitle
				  RSObj("Tid")=ClassID
				  RSObj("Otid")=KS.S("OTid")
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Inputer")="游客"
				  RSObj("Origin")=Origin
				  RSObj("ArticleContent")=Content
				  RSObj("Verific")=0
				  RSObj("photoUrl")=PicUrl
				  RSObj("Intro")=Intro
				  if PicUrl<>"" Then 
				   RSObj("PicNews")=1
				  Else
				   RSObj("PicNews")=0
				  End if
				  RSObj("Province")= Province
	              RSObj("City")    = City
				  RSObj("County")  = County
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonuserform").text="1" Then	RSObj("MapMarker")=KS.S("MapMark")
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonuserform").text="1" Then
				  RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("Adddate")=Now
				  RSObj("Rank")="★★★"
				  Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Intro,KeyWords,PicUrl,"游客",0,Fname)
                 Call KS.FileAssociation(ChannelID,InfoID,Content & PicUrl ,0)
				 
				 
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
				 
				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){top.location.href='Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveByPhoto()
	            Dim Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,PicUrls,Action,I
  				  ClassID=KS.S("ClassID")
				  Title=KS.LoseHtml(KS.S("Title"))
				  KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				  Author=KS.LoseHtml(KS.S("Author"))
				  Origin=KS.LoseHtml(KS.S("Origin"))
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				  PhotoUrl=KS.S("PhotoUrl")
				  PicUrls=KS.S("PicUrls")

				 Call KSUser.CheckDiyField(FieldXML,false)	
				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "名称!');</script>"
				    Exit Sub
				  End IF
				  If KS.ChkClng(KS.S("autothumb"))=1 And KS.IsNul(PhotoUrl) Then  PhotoUrl=Split(Split(PicUrls,"|||")(0),"|")(2)
	              If PhotoUrl="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "缩略图!');</script>"
				    Exit Sub
				  End IF
	              If PicUrls="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "!');</script>"
				    Exit Sub
				  End IF
				 
			  Dim Fname,FnameType,TemplateID,WapTemplateID
			  FnameType=KS.C_C(ClassID,23)
			  Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
			  TemplateID=KS.C_C(ClassID,5)
			  WapTemplateID=KS.C_C(ClassID,22)
			Dim AdminMobile:AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
			
				  
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & "",Conn,1,3
				RSObj.AddNew
				  RSObj("Title")=Title
				  RSObj("Tid")=ClassID
				  RSObj("Otid")=KS.S("OTid")
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("PicUrls")=PicUrls
				  RSObj("PicNum") = Ubound(split(PicUrls,"|||"))+1
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Inputer")="游客"
				  RSObj("Origin")=Origin
				  RSObj("PictureContent")=Content
				  RSObj("Verific")=0
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("AddDate")=Now
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonuserform").text="1" Then	RSObj("MapMarker")=KS.S("MapMark")
				   If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonuserform").text="1" Then
				  RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				   Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PicUrls & PhotoUrl ,0)
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				 
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
				  
				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){top.location.href='Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveByDownLoad()
		Dim SizeUnit,ClassID,Title,KeyWords,Author,DownLB,DownYY,DownSQ,DownSize,DownPT,YSDZ,ZCDZ,JYMM,Origin,Content,Verific,PhotoUrl,DownUrls,RSObj,ID,DownID,AddDate,ComeUrl,CurrentOpStr,Action,I
				  ClassID=KS.S("ClassID")
				  Title=KS.LoseHtml(KS.S("Title"))
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
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				  PhotoUrl=KS.LoseHtml(KS.S("PhotoUrl"))
				  DownUrls=KS.S("DownUrls")
				  if (Instr(lcase(DownUrls),lcase(KS.Setting(2)))<>0 and Instr(lcase(DownUrls),"user/" & KSUser.GetUserInfo("userid") &"/")=0) or (Instr(lcase(DownUrls),"http://")=0 and Instr(lcase(DownUrls),"user/" & KSUser.GetUserInfo("userid") &"/")=0) or Instr(lcase(DownUrls),".asp")<>0 or KS.IsNul(Request.Form("DownUrls")) then
				   KS.Die "<script>alert('软件地址格式不正确!');</script>"
				  end if
				  
				  
				Call KSUser.CheckDiyField(FieldXML,false)					  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');</script>":Exit Sub
				  If Title="" Then  Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "名称!');</script>":Exit Sub
	              If DownUrls="" Then Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "!');</script>": Exit Sub
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				  Dim Fname,FnameType,TemplateID,WapTemplateID
				  FnameType=KS.C_C(ClassID,23)
				  Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
				  TemplateID=KS.C_C(ClassID,5)
				  WapTemplateID=KS.C_C(ClassID,22)
				  Dim AdminMobile:AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
				  
					RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & "",Conn,1,3
					RSObj.AddNew
					  RSObj("Title")=Title
					  RSObj("TID")=ClassID
					  RSObj("Otid")=KS.S("OTid")
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
					  RSObj("DownUrls")="0|下载地址|" & DownUrls
					  RSObj("Inputer")="游客"
					  RSObj("Verific")=0
					  RSObj("Hits")=0
				      RSObj("TemplateID")=TemplateID
					  RSObj("WapTemplateID")=WapTemplateID
				      RSObj("Fname")=FName
					  RSObj("AddDate")=Now
					  RSObj("Rank")="★★★"
				      Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
					RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl & DownUrls ,0)
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				 
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
				  
				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){top.location.href='Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
			
	End Sub
	
	Sub SaveByFlash
		Dim Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,FlashUrl,RSObj,I
  		ClassID=KS.S("ClassID")
		Title=KS.LoseHtml(KS.S("Title"))
		KeyWords=KS.LoseHtml(KS.S("KeyWords"))
		Author=KS.LoseHtml(KS.S("Author"))
		Origin=KS.LoseHtml(KS.S("Origin"))
		Content = Request.Form("Content")
		Content =KS.ClearBadChr(content)
		PhotoUrl=KS.S("PhotoUrl")
		FlashUrl=KS.S("FlashUrl")
				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择"& KS.C_S(ChannelID,3) & "栏目!');</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "名称!');</script>"
				    Exit Sub
				  End IF
	              If FlashUrl="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "!');</script>"
				    Exit Sub
				  End IF
				Call KSUser.CheckDiyField(FieldXML,false)		
			  Dim Fname,FnameType,TemplateID,WapTemplateID
			  FnameType=KS.C_C(ClassID,23)
			  Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
			  TemplateID=KS.C_C(ClassID,5)
			  WapTemplateID=KS.C_C(ClassID,22)
			  Dim AdminMobile:AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
			  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
					RSObj.Open "Select top 1 * From KS_Flash Where 1=0",Conn,1,3
				  RSObj.AddNew
				   RSObj("Hits")=0
				   RSObj("TemplateID")=TemplateID
				   RSObj("Fname")=FName
				   RSObj("AddDate")=Now
				   RSObj("Rank")="★★★"
				   RSObj("Title")=Title
				   RSObj("TID")=ClassID
				   RSObj("Otid")=KS.S("OTid")
				   RSObj("PhotoUrl")=PhotoUrl
				   RSObj("FlashUrl")=FlashUrl
				   RSObj("KeyWords")=KeyWords
				   RSObj("Author")=Author
				   RSObj("Inputer")="游客"
				   RSObj("Origin")=Origin
				   RSObj("FlashContent")=Content
				   RSObj("Verific")=0
				   Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl & FlashUrl ,0)
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				 
				 
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
				Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){top.location.href='Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveByMovie()
		Dim Title,KeyWords,MovieAct,MovieDY,Content,Verific,PhotoUrl,MovieUrl,RSObj,I
				ClassID=KS.S("ClassID")
				Title=KS.LoseHtml(KS.S("Title"))
				KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				MovieAct=KS.LoseHtml(KS.S("MovieAct"))
				MovieDY=KS.LoseHtml(KS.S("MovieDY"))
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				PhotoUrl=KS.S("PhotoUrl")
				MovieUrl=KS.S("MovieUrl")
				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择"& KS.C_S(ChannelID,3) & "栏目!');</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "名称!');</script>"
				    Exit Sub
				  End IF
	              If MovieUrl="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "!');</script>"
				    Exit Sub
				  End IF
				Call KSUser.CheckDiyField(FieldXML,false)				  
			  Dim Fname,FnameType,TemplateID,WapTemplateID
			  FnameType=KS.C_C(ClassID,23)
			  Fname=KS.GetFileName(KS.C_C(ClassID,24), Now, FnameType)
			  TemplateID=KS.C_C(ClassID,5)
			  WapTemplateID=KS.C_C(ClassID,22)
			  Dim AdminMobile:AdminMobile=Split(KS.C_C(ClassID,6)&"||||||||||||||","||||")(5)
			  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_Movie Where 1=0",Conn,1,3
				  RSObj.AddNew
				  RSObj("TemplateID")=TemplateID
				  RSObj("ServerID")=0
				  RSObj("Fname")=FName
				  RSObj("Hits")=0
				  RSObj("AddDate")=Now
				  RSObj("Rank")="★★★"
				  RSObj("Title")=Title
				  RSObj("TID")=ClassID
				  RSObj("Otid")=KS.S("OTid")
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("MovieUrls")=MovieUrl
				  RSObj("KeyWords")=KeyWords
				  RSObj("MovieAct")=MovieAct
				  RSObj("Inputer")="游客"
				  RSObj("MovieDY")=MovieDY
				  RSObj("MovieContent")=Content
				  RSObj("Verific")=0
				  Call KSUser.AddDiyFieldValue(RSObj,FieldXML)
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl & MovieUrl ,0)
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				 
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
				 
				Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){top.location.href='Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveBySupply()
		Dim GQID,Title,Price,TypeID,ValidDate,GQContent,ContactMan,Tel,CompanyName,Address,Province,City,County,Email,Zip,Fax,HomePage,I,PhotoUrl,Visitor,KeyWords,Verific,inputer
			 ClassID      = KS.S("ClassID")
			 Title        = KS.LoseHtml(KS.S("Title"))
			 PhotoUrl     = KS.LoseHtml(KS.S("PhotoUrl"))
			 Price        = KS.LoseHtml(KS.S("Price"))
			 TypeID       = KS.S("TypeID")
			 ValidDate    = KS.S("ValidDate")
			 GQContent = Request.Form("GQContent")
			 GQContent=KS.ClearBadChr(GQContent)
			 ContactMan   = KS.LoseHtml(KS.S("ContactMan"))
			 Tel          = KS.LoseHtml(KS.S("Mobile"))
			 CompanyName  = KS.LoseHtml(KS.S("CompanyName"))
			 Address      = KS.LoseHtml(KS.S("Address"))
			 Province     = KS.LoseHtml(KS.S("Province"))
			 City         = KS.LoseHtml(KS.S("City"))
			 County       = KS.LoseHtml(KS.S("County"))
			 Email        = KS.LoseHtml(KS.S("Email"))
			 Zip          = KS.LoseHtml(KS.S("Zip"))
			 Fax          = KS.LoseHtml(KS.S("Fax"))
			 HomePage     = KS.LoseHtml(KS.S("HomePage"))
			 KeyWords     = KS.LoseHtml(KS.S("KeyWords"))
		Call KSUser.CheckDiyField(FieldXML,false)	
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")		
		  RS.Open "Select top 1 * From [KS_Class] Where ID='" & ClassID & "'", conn, 1, 1
		  If RS.Eof And Rs.Bof Then
		   Response.Write "<script>alert('非法参数!');history.back();</script>"
		   response.end
		  end if
		  Dim TemplateID,GQFsoType,GQFnameType
		  TemplateID=RS("Templateid")
		  GQFsoType=RS("FsoType")
		  GQFnameType = Trim(RS("FnameType"))
		  Dim AdminMobile:AdminMobile=Split(KS.C_C(RS("ClassID"),6)&"||||||||||||||","||||")(5)
		  RS.Close
		  Dim Fname:Fname=KS.GetFileName(GQFsoType, Now, GQFnameType)
		  RS.Open "select top 1 * from KS_GQ where 1=0", conn, 1, 3
		   RS.AddNew
		   RS("Hits")=0
		   RS("AddDate")=Now
		   RS("TemplateID")=TemplateID
		   RS("Fname")=Fname
		   RS("Recommend")=0
		   RS("IsTop")=0
		   IF Cbool(KSUser.UserLoginChecked)=false Then	inputer="游客" Else inputer=KS.C("UserName")
		   RS("Inputer")=inputer
		   RS("Tid")=ClassID
		   RS("Otid")=KS.S("OTid")
		   RS("Title")=Title
		   RS("Price")=Price
		   RS("PhotoUrl")=PhotoUrl
		   RS("TypeID")=TypeID
		   RS("ValidDate")=ValidDate
		   RS("GQContent")=GQContent
		   RS("KeyWords")=KeyWords
		   If KS.C_S(ChannelID,17)=1 Then Verific=0 Else Verific=1
		   RS("Verific")=verific
		   RS("ContactMan")=ContactMan
		   RS("Tel")=Tel
		   RS("Mobile")=Tel
		   RS("CompanyName")=CompanyName
		   RS("Address")=Address
		   RS("Province")=Province
		   RS("County")=County
		   RS("City")=City
		   RS("Email")=Email
		   RS("Zip")=Zip
		   RS("Fax")=Fax
		   RS("Homepage")=Homepage
		   Call KSUser.AddDiyFieldValue(RS,FieldXML)
		   RS.Update
		   RS.MoveLast
				Dim InfoID: InfoID=RS("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RS("Fname") = InfoID & GQFnameType
					RS.Update
				End If
				Fname=RS("Fname")

				 RS.Close:Set RS=Nothing
				 Call KS.FileAssociation(ChannelID,InfoID,GQContent & PhotoUrl ,0)
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,GQContent,KeyWords,PhotoUrl,inputer,verific,Fname)
				 
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
				 
		 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){top.location.href='Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"

	End Sub
End Class
%> 

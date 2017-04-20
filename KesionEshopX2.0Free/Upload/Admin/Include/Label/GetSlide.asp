<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%
'****************************************************
' Software name:KesionCMS X2.0
' Email: service@EasyTool.CN . QQ:111394,9537636
' Web: http://www.EasyTool.CN http://www.KeSion.cn
' Copyright (C) KeSion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GetSlide
KSCls.KeSion()
Set KSCls = Nothing

Class GetSlide
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub KeSion()
		Dim TempClassList, FolderID, LabelContent, L_C_A, Action,Page, LabelID, Str, Descript, LabelFlag,Wrap,Auto,loadIMGTimeout, OrderStr,oTid,IncludeoTid
		Dim ClassID, IncludeSubClass, PicWidth, PicHeight, Num, OpenType, trigger, TitleLen, IntroLen,txtHeight, delay,SlideType,SpecialID,DocProperty,From,Attr
		FolderID = Request("FolderID")
		Dim ChannelID:ChannelID=KS.G("ChannelID")
		With KS
		
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		Page=KS.ChkClng(Request("Page"))
		From =KS.S("From")
		If LabelID = "" Then
		  ClassID = "0"
		  Action = "Add"
		Else
		  Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select TOP 1 * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			Descript = LabelRS("Description")
			FolderID = LabelRS("FolderID")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent = Replace(Replace(LabelContent, "{Tag:GetSlide", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');</Script>")
			 response.End()
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			  ChannelID          = KS.ChkClng(Node.getAttribute("modelid"))
			  If ChannelID=-1000 Then From="club"
			  If ChannelID=-1001 Then From="bloginfo"
			  If ChannelID=-1002 Then From="special"
			  If ChannelID=-1003 Then From="ads"
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  oTid               = Node.getAttribute("otid")
			  IncludeoTid        = Node.getAttribute("includeotid")
			  SpecialID          = Node.getAttribute("specialid")
			  OrderStr           = Node.getAttribute("orderstr")
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  Num                = Node.getAttribute("num")
			  OpenType           = Node.getAttribute("opentype")
			  TitleLen           = Node.getAttribute("titlelen")
			  IntroLen           = Node.getAttribute("introlen")
			  txtHeight          = Node.getAttribute("txtheight")
			  delay         = Node.getAttribute("delay")
			  SlideType          = Node.getAttribute("slidetype")
			  DocProperty        = Node.getAttribute("docproperty")
			  Attr               = Node.getAttribute("attr")
			  trigger            = Node.getAttribute("trigger")
			  Wrap               = Node.getAttribute("wrap")
			  Auto               = Node.getAttribute("auto")
			  loadIMGTimeout     = Node.getAttribute("loadimgtimeout")
		   End If
		   Set Node=Nothing
		   Set XMLDoc=Nothing
		End If
		If ChannelID="" Then ChannelID=0
		If oTid="" Then oTid="0"
		if KS.IsNul(IncludeoTid) Then IncludeoTid=false
		If Num = "" Then Num = 5
		If TitleLen = "" Then TitleLen = 30
		If IntroLen = "" Then IntroLen = 200
		If PicWidth = "" Then PicWidth = 300
		If PicHeight = "" Then PicHeight = 300
		If SpecialID="" Then SpecialID=0
		If Wrap="" Then Wrap="true"
		If Auto="" Then Auto="true"
		If loadIMGTimeout="" then loadIMGTimeout=0
		If DocProperty = "" Then DocProperty = "00001"
		If KS.IsNul(trigger) Then trigger="click"
		If KS.IsNul(SlideType) Then SlideType="mF_taobao2010"
		.echo "<!DOCTYPE html><html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &vbcrlf
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">" &vbcrlf
		.echo "<script src=""../../../KS_Inc/jQuery.js"" language=""JavaScript""></script>" &vbcrlf
		.echo "<script src=""../../../KS_Inc/Common.js"" language=""JavaScript""></script>" &vbcrlf
		.echo "<script src=""../../../ks_inc/myFocus/myfocus-2.0.4.min.js""></script>" &vbcrlf
		%>
		<script language="javascript">
		$(document).ready(function(){
		 $("#ChannelID").change(function(){
		 
		    $(parent.document).find('#ajaxmsg').toggle();
			GetAttribute($(this).val());
			$.get('../../../plus/ajaxs.asp',{from:'label',action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty().append("<option value='-1' style='color:red'>-当前栏目(通用)-</option>").append("<option value='0'>-不指定栏目-</option>").append(unescape(data));
			  $(parent.document).find('#ajaxmsg').toggle();
			 })
		   })	
		  $("#MutileClass").click(function(){
		    if ($(this).prop("checked")==true){
		      $("#ClassList").attr("multiple","multiple").attr("style","height:60px");
		    }else{
			   $("#ClassList").removeAttr("multiple");
			    $("#ClassList").removeAttr("style");

			}
		  });
		  $(window).load(function(){
		  $("#SlideType>option[value=<%=SlideType%>]").attr("selected",true);
		  });
		   <%if Instr(ClassID,",")<>0 Then%>
		   var searchStr="<%=ClassID%>";
		   $("#MutileClass").attr("checked",true);
		   $("#ClassList").attr("multiple","multiple").attr("style","height:60px");
		    setTimeout(function(){ 
		   $("#ClassList>option").each(function(){
		     if($(this).val()=='-1' || $(this).val()=='0')
			  $(this).attr("selected",false)
			 else if (searchStr.indexOf($(this).val())!=-1)
			 { 
			   $(this).attr("selected",true);
			 }
		   });},1);
		  <%end if%>
           <%If LabelID<>"" Then%>
		   GetAttribute($("#ChannelID").val());
		  <%End If%>
		});
		
		function GetAttribute(channelid){
		    $.get('../../../plus/ajaxs.asp',{action:'GetModelAttr',attr:'<%=attr%>',channelid:channelid},function(data){
			  $("#showattr").html('').html(data)
			 });
		}
		function SetLabelFlag(Obj)
		{
		 if (Obj.value=='-1')
		  $("#LabelFlag").val(1);
		  else
		  $("#LabelFlag").val(0);
		}
		function SpecialChange(SpecialID)
		{
			if (SpecialID==-1) 
			  $("#ClassArea").hide();
			else
			  $("#ClassArea").show();	
		}
		function CheckForm()
		{   if ($("input[name=LabelName]").val()=='')
			 {
			  top.$.dialog.alert('请输入标签名称!',function(){
			  $("input[name=LabelName]").focus(); 
			  });
			  return false
			  }
            var ClassList='';
		    if ($("#MutileClass").prop("checked")==true){
				$("#ClassList option:selected").each(function(){
					if ($(this).val()!='0' && $(this).val()!='-1')
						if (ClassList=='') 
						 ClassList=$(this).val() 
						else
						 ClassList+=","+$(this).val();
					})
			 }else{
			    ClassList=$("#ClassList").val();
			 }			  
		<%If From="club" then%>
		    var ChannelID=-1000;
			var SpecialID='0';
			var DocProperty='000000';
		<%ElseIf From="bloginfo" then%>
		    var ChannelID=-1001;
			var SpecialID='0';
			var DocProperty='000000';
		<%ElseIf From="special" Then%>
		    var ChannelID=-1002;
			var SpecialID='0';
			var DocProperty='000000';
		<%ElseIf From="ads" Then%>
		    var ChannelID=-1003;
			var SpecialID='0';
			var DocProperty='000000';
		<%Else%>	  
			var ChannelID=$("#ChannelID").val();
			
			var SpecialID=$("select[name=SpecialID]").val();
			if (SpecialID==-1) ClassList=0;
			var DocProperty='';
			 $("input[name=DocProperty]").each(function(){
			     if ($(this).prop("checked")==true){
				  DocProperty=DocProperty+'1'
				 }else{
				  DocProperty=DocProperty+'0'
				 }      
			 })
		<%End If%>
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var Num=$("input[name=Num]").val();
			var OpenType=$("#OpenType").val();
			var OrderStr=$("#OrderStr").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var IntroLen=$("#IntroLen").val();
			var txtHeight=$("input[name=txtHeight]").val();
			var delay=$("input[name=delay]").val();
			var SlideType=$("#SlideType").val();
			var trigger=$("#trigger").val();
			var wrap=$("#wrap").val();
			var auto=$("#auto").val();
			var loadIMGTimeout=$("#loadIMGTimeout").val();
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").prop("checked")==true) IncludeSubClass=true;
		    var IncludeoTid=false;
			if ($("#IncludeoTid").prop("checked")==true) IncludeoTid=true;
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			var av='';
		   $("input[name=attr]").each(function(){
		     if ($(this).prop("checked")==true){
			   if (av==''){
			    av=$(this).val();
			   }else{
			    av+='|'+$(this).val();
			   }
			 }
		   });
		   
			var tagVal='{Tag:GetSlide labelid="0" modelid="'+ChannelID+'" classid="'+ClassList+'" otid="'+$("#oTid").val()+'" includeotid="'+IncludeoTid+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" attr="'+av+'" docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" trigger="'+trigger+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" num="'+Num+'" opentype="'+OpenType+'" titlelen="'+TitleLen+'" introlen="'+IntroLen+'" txtheight="'+txtHeight+'" wrap="'+wrap+'" auto="'+auto+'" loadimgtimeout="'+loadIMGTimeout+'" delay="'+delay+'" slidetype="'+SlideType+'"}{/Tag}';
		 
			$("#LabelContent").val(tagVal);
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" onload=""SpecialChange(" & SpecialID &");"">"
		%>
		<div id="view" style="position:absolute;width:250px;top:140px; right:20px;cursor:pointer; text-align:center">
		<div style="background-color:#000; color:#fff; line-height:20px; text-align:center">幻灯预览 <span  onclick="$('#view').hide()">点此关闭</span></div>

			<div id="focus_box">
			
			<div id="focus">
			<div class="loading"></div>
				 <div class="pic"><!--图片列表-->
					<ul>
				<li><a href="http://www.kesion.com" target="_blank" title="标题1"><img width="250" height="220" src="../../Images/View/1.jpg" thumb="" alt="标题1" text="简介1" /></a></li>
				<li><a href="http://www.kesion.com" target="_blank" title="标题2"><img width="250" height="220" src="../../Images/View/2.jpg" thumb="" alt="标题2" text="简介2" /></a></li>
				<li><a href="http://www.kesion.com" target="_blank" title="标题3"><img width="250" height="220" src="../../Images/View/3.jpg" thumb="" alt="标题3" text="简介3" /></a></li>
				<li><a href="http://www.kesion.com" target="_blank" title="标题4"><img width="250" height="220" src="../../Images/View/4.jpg" thumb="" alt="标题4" text="简介4" /></a></li>
				<li><a href="http://www.kesion.com" target="_blank" title="标题5"><img width="250" height="220" src="../../Images/View/5.jpg" thumb="" alt="标题5" text="简介5" /></a></li>
					</ul>
				</div>
			</div>
			</div>
             </div>
 <div style="clear:both"></div>

		

<script>
		
		$(document).ready(function(){
			 $("#view").mousedown(function(event){  
					var offset=$("#view").offset();   
					x1=event.clientX-offset.left;   
					y1=event.clientY-offset.top;   
					$("#view").mousemove(function(event){   
					   $("#view").css("left",(event.clientX-x1)+"px");   
					   $("#view").css("top",(event.clientY-y1)+"px");   
					});   
			
					$("#view").mouseup(function(event){   
						$("#view").unbind("mousemove");   
					});   
			
			  });
			      $("#SlideType").val('<%=SlideType%>');
                  myFocus.set({id:'focus',pattern:'<%=SlideType%>',loadIMGTimeout:0,width:250,height:220});
		});
	  
	  var $id=function(id){return document.getElementById(id)};
	  var oriHtml=$id('focus_box').innerHTML;
	  function resetHTML(){//还原
	        
			$id('focus_box').innerHTML=oriHtml;
			var css=document.getElementsByTagName('style')[0];
			//alert(css);
			css.parentNode.removeChild(css);
		}

		function changeSlide(v){
             resetHTML();
			 $('#view').show();
			myFocus.set({id:'focus',pattern:v,loadIMGTimeout:0,width:250,height:220});
			
		}
      </script>

		<%
		.echo "<div align=""center"" class='pageCont2'>"
		.echo "<iframe src='about:blank' name='_hiddenframe' style='display:none' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo " <input type=""hidden"" name=""Page"" id=""Page"" value=""" & Page & """>"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """> "
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSlide.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		If From="club" Then
		.echo "            <tr class=tdbg>"
		.echo "              <td  height=""24"" colspan=""4"" style=""text-align:center""><strong>论坛幻灯片调用标签</strong></td></tr>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td height=""24"" colspan=""4"">指定版面 "
		
		 .echo "<select name='ClassList' id='ClassList'>"
		 .echo "<option value='0'>--不限版面分类--</option>"
		 KS.LoadClubBoard
		 Dim Tstr,n
		 for each node in Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectNodes("row[@parentid='0']")
		   Tstr=""
		  If ClassID=Node.SelectSingleNode("@id").text Then
		  .echo "<option value='" & Node.SelectSingleNode("@id").text &"' selected>" & Tstr &  Node.SelectSingleNode("@boardname").text &"</option>"
		  Else
		  .echo "<option value='" & Node.SelectSingleNode("@id").text &"'>" & Tstr & Node.SelectSingleNode("@boardname").text &"</option>"
		  End If
		   For each n in Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectNodes("row[@parentid='" & Node.SelectSingleNode("@id").text & "']")
		      Tstr="&nbsp;&nbsp;|--"
			  If ClassID=N.SelectSingleNode("@id").text Then
			  .echo "<option value='" & N.SelectSingleNode("@id").text &"' selected>" & Tstr &  N.SelectSingleNode("@boardname").text &"</option>"
			  Else
			  .echo "<option value='" & N.SelectSingleNode("@id").text &"'>" & Tstr & N.SelectSingleNode("@boardname").text &"</option>"
			  End If
		   Next
		 next
		 .echo "</select>"
		.echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多个版面</td></tr>"
		ElseIf From="bloginfo" Then
		.echo "            <tr class=tdbg>"
		.echo "              <td  height=""24"" colspan=""4"" style=""text-align:center""><strong>博文幻灯片调用标签</strong></td></tr>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td height=""24"" colspan=""4"">博文分类 "
		
		 .echo "<select name='ClassList' id='ClassList'>"
		 .echo "<option value='0'>--不限博文分类--</option>"
	 	 Dim RS:Set Rs = Conn.Execute("SELECT typeid,depth,typeName FROM KS_blogtype ORDER BY rootid,orderid")
							Do While Not Rs.EOF
								Response.Write "<option value=""" & Rs("typeid") & """ "
								If KS.ChkClng(ClassID)=KS.ChkClng(RS("TypeID")) Then Response.Write "selected"
								Response.Write ">"
								If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
								If Rs("depth") > 1 Then
									For i = 2 To Rs("depth")
										Response.Write "&nbsp;&nbsp;│"
									Next
									Response.Write "&nbsp;&nbsp;├ "
								End If
								Response.Write Rs("typeName") & "</option>" & vbCrLf
								Rs.movenext
							Loop
			Rs.Close
		   Set Rs = Nothing
		   
		 .echo "</select>"
		.echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多个博文分类</td></tr>"	
	ElseIf From="special" Then
		.echo "            <tr class=tdbg>"
		.echo "              <td  height=""24"" colspan=""4"" style=""text-align:center""><strong>专题幻灯片调用标签</strong></td></tr>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td height=""24"" colspan=""4"">专题分类 "
		
		 .echo "<select name='ClassList' id='ClassList'>"
		 .echo "<option value='0'>--不限专题分类--</option>"
	 	  Set Rs = Conn.Execute("SELECT classid,className FROM KS_SpecialClass ORDER BY orderid,classid")
							Do While Not Rs.EOF
								Response.Write "<option value=""" & Rs("classid") & """ "
								If KS.ChkClng(ClassID)=KS.ChkClng(RS("classid")) Then Response.Write "selected"
								Response.Write ">"
								Response.Write Rs("ClassName") & "</option>" & vbCrLf
								Rs.movenext
							Loop
			Rs.Close
		   Set Rs = Nothing
		   
		 .echo "</select>"
		.echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多个专题分类</td></tr>"	
	ElseIf From="ads" Then
		.echo "            <tr class=tdbg>"
		.echo "              <td  height=""24"" colspan=""4"" style=""text-align:center""><strong>广告系统幻灯片调用标签</strong></td></tr>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td height=""24"" colspan=""4"">选择广告位 "
		
		 .echo "<select name='ClassList' id='ClassList'>"
		 .echo "<option value='0'>--选择广告位--</option>"
	 	  Set Rs = Conn.Execute("SELECT place,placename FROM KS_ADPlace ORDER BY place")
							Do While Not Rs.EOF
								Response.Write "<option value=""" & Rs("place") & """ "
								If KS.ChkClng(ClassID)=KS.ChkClng(RS("place")) Then Response.Write "selected"
								Response.Write ">"
								Response.Write Rs("placename") & "</option>" & vbCrLf
								Rs.movenext
							Loop
			Rs.Close
		   Set Rs = Nothing
		   
		 .echo "</select>"
		.echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多个广告位</td></tr>"	
	Else
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"" colspan=""2"">选择范围"
		.echo "                <select name=""ChannelID"" id=""ChannelID"">"
		.echo "                 <option value=""0"">-所有模型-</option>"
        .LoadChannelOption ChannelID
		.echo "                </select>"
		.echo "                <select class=""textbox"" name=""ClassList"" id=""ClassList"" onChange=""SetLabelFlag(this)"">"
		.echo "                 <option selected value=""-1"" style=""color:red"">- 当前栏目(通用)-</option>"
						
						If ClassID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定栏目 -</option>")
						Else
						  .echo ("<option  value=""0"">- 不指定栏目 -</option>")
					   End If
						  .echo Replace(KS.LoadClassOption(ChannelID,false),"value='" & ClassID & "'","value='" & ClassID &"' selected")
						  .echo "</select>"

						  
					If cbool(IncludeSubClass) = True Or LabelID = "" Then
					  Str = " Checked"
					Else
					  Str = ""
					End If
					  .echo "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多栏目"
					  .echo ("&nbsp;&nbsp;&nbsp;<input name=""IncludeSubClass"" type=""checkbox"" id=""IncludeSubClass"" value=""true""" & Str & ">调用子栏目")
			
		.echo "&nbsp;&nbsp;所属附栏目"
		.echo "                <select class=""textbox"" style=""width:170px"" name=""oTid"" id=""oTid"">"
		.echo "                 <option value=""0"" style=""color:red"">- 不限-</option>"
		If oTid="-1" Then
		.echo "                 <option value=""-1"" style=""color:red"" selected>- 自动匹配当前主栏目 -</option>"
		Else
		.echo "                 <option value=""-1"" style=""color:red"">- 自动匹配当前主栏目 -</option>"
		End If
		If oTid="-2" Then
		.echo "                 <option value=""-2"" style=""color:green"" selected>- 自动匹配相同的附属栏目 -</option>"
		Else
		.echo "                 <option value=""-2"" style=""color:green"">- 自动匹配相同的附属栏目 -</option>"
		End If
		If oTid="-3" Then
		.echo "                 <option value=""-3"" style=""color:blue"" selected>- 自动匹配当前的文档ID -</option>"
		Else
		.echo "                 <option value=""-3"" style=""color:blue"">- 自动匹配当前的文档ID -</option>"
		End If
		
		
						  .echo Replace(KS.LoadClassOption(0,false),"value='" & oTid & "'","value='" & oTid &"' selected")
						  .echo "</select>"
			If cbool(IncludeoTid) = True Then
					  Str = " Checked"
			Else
					  Str = ""
			End If			  
		.echo ("&nbsp;&nbsp;&nbsp;<input name=""IncludeoTid"" type=""checkbox"" id=""IncludeoTid"" value=""true""" & Str & ">调用子栏目")
		.echo "            </td></tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">所属专题"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:35%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- 当前专题(专题页通用)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定专题 -</option>")
						   Else
						  .echo ("<option  value=""0"">- 不指定专题 -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
        .echo "</td>"
		.echo "              <td width=""50%"" height=""24"">属性控制"
		.echo "                <label><input name=""DocProperty"" type=""checkbox"" value=""1"""
		If mid(DocProperty,1,1) = 1 Then .echo (" Checked")
		.echo ">推荐</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox""  value=""2"""
		If mid(DocProperty,2,1) = 1 Then .echo (" Checked")
		  .echo ">滚动</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""3"""
		If mid(DocProperty,3,1) = 1 Then .echo (" Checked")
		  .echo ">头条</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""4"""
		If mid(DocProperty,4,1) = 1 Then .echo (" Checked")
		  .echo ">热门</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"" checked disabled>幻灯</label>"
		
		.echo " <span id=""showattr""></span> </td>"
		.echo "</tr>"
	End If

		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">幻灯类型"
		.echo " <select name=""SlideType"" id=""SlideType"" onchange=""changeSlide(this.value)"">"
		%>
		<option value="mF_fscreen_tb">mF_fscreen_tb</option>     <option value="mF_YSlider">mF_YSlider</option>     <option value="mF_luluJQ">mF_luluJQ</option>     <option value="mF_51xflash">mF_51xflash</option>     <option value="mF_expo2010">mF_expo2010</option>     <option value="mF_games_tb">mF_games_tb</option>     <option value="mF_ladyQ">mF_ladyQ</option>     <option value="mF_liquid">mF_liquid</option>     <option value="mF_liuzg">mF_liuzg</option>     <option value="mF_pithy_tb">mF_pithy_tb</option>     <option value="mF_qiyi">mF_qiyi</option>     <option value="mF_quwan">mF_quwan</option>     <option value="mF_rapoo">mF_rapoo</option>     <option value="mF_sohusports">mF_sohusports</option>     <option value="mF_taobao2010">mF_taobao2010</option>     <option value="mF_taobaomall">mF_taobaomall</option>     <option value="mF_tbhuabao">mF_tbhuabao</option>     <option value="mF_pconline">mF_pconline</option>     <option value="mF_peijianmall">mF_peijianmall</option>     <option value="mF_classicHC">mF_classicHC</option>     <option value="mF_classicHB">mF_classicHB</option>     <option value="mF_slide3D">mF_slide3D</option>     <option value="mF_kiki">mF_kiki</option>     <option style="color:#f00;" value="mF_fancy" selected="selected">mF_fancy</option>     <option style="color:#f00;" value="mF_dleung">mF_dleung</option>     <option style="color:#f00;" value="mF_kdui">mF_kdui</option>     <option style="color:#f00;" value="mF_shutters">mF_shutters</option>
	</select>	
		<%
		
		.echo "              </td>"
		.echo "              <td height=""30"">图片大小 宽"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth2"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'图片宽度');"">"
		.echo "                像素 高"
		.echo "                <input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight2"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'图片高度');"">"
		.echo "                像素</td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">查询条数"
					  
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:50px;text-align:center"" onBlur=""CheckNumber(this,'图片数量');"" value=""" & Num & """> 条"
		if from="" then
		.echo "<span>"
		else
		.echo "<span style='display:none'>"
		end if
		.echo " 排序方法<select style=""width:150px;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>文档ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>文档ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>文档ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>文档ID(升序)</option>")
					End If
					If OrderStr = "Rnd" Then
					.echo ("<option value='Rnd' style='color:blue' selected>随机显示</option>")
					Else
					.echo ("<option value='Rnd' style='color:blue'>随机显示</option>")
					End If
					
					If OrderStr = "ModifyDate Asc" Then
					.echo ("<option value='ModifyDate Asc' selected>修改时间(升序)</option>")
					Else
					.echo ("<option value='ModifyDate Asc'>修改时间(升序)</option>")
					End If
					If OrderStr = "ModifyDate Desc" Then
					 .echo ("<option value='ModifyDate Desc' selected>修改时间(降序)</option>")
					Else
					 .echo ("<option value='ModifyDate Desc'>修改时间(降序)</option>")
					End If
					If OrderStr = "AddDate Asc" Then
					.echo ("<option value='AddDate Asc' selected>添加时间(升序)</option>")
					Else
					.echo ("<option value='AddDate Asc'>添加时间(升序)</option>")
					End If
					If OrderStr = "AddDate Desc" Then
					 .echo ("<option value='AddDate Desc' selected>添加时间(降序)</option>")
					Else
					 .echo ("<option value='AddDate Desc'>添加时间(降序)</option>")
					End If
					
                    If OrderStr = "CmtNum Asc" Then
					 .echo ("<option value='CmtNum Asc' selected>评论数(升序)</option>")
					Else
					 .echo ("<option value='CmtNum Asc'>评论数(升序)</option>")
					End If
					If OrderStr = "CmtNum Desc" Then
					  .echo ("<option value='CmtNum Desc' selected>评论数(降序)</option>")
					Else
					  .echo ("<option value='CmtNum Desc'>评论数(降序)</option>")
					End If					
					
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>点击数(升序)</option>")
					Else
					 .echo ("<option value='Hits Asc'>点击数(升序)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>点击数(降序)</option>")
					Else
					  .echo ("<option value='Hits Desc'>点击数(降序)</option>")
					End If
					If OrderStr = "HitsByDay Asc" Then
					 .echo ("<option value='HitsByDay Asc' selected>日访问量(升序)</option>")
					Else
					 .echo ("<option value='HitsByDay Asc'>日访问量(升序)</option>")
					End If
					If OrderStr = "HitsByDay Desc" Then
					  .echo ("<option value='HitsByDay Desc' selected>日访问量(降序)</option>")
					Else
					  .echo ("<option value='HitsByDay Desc'>日访问量(降序)</option>")
					End If
					If OrderStr = "HitsByWeek Asc" Then
					 .echo ("<option value='HitsByWeek Asc' selected>周访问量(升序)</option>")
					Else
					 .echo ("<option value='HitsByWeek Asc'>周访问量(升序)</option>")
					End If
					If OrderStr = "HitsByWeek Desc" Then
					  .echo ("<option value='HitsByWeek Desc' selected>周访问量(降序)</option>")
					Else
					  .echo ("<option value='HitsByWeek Desc'>周访问量(降序)</option>")
					End If
					If OrderStr = "HitsByMonth Asc" Then
					 .echo ("<option value='HitsByMonth Asc' selected>月访问量(升序)</option>")
					Else
					 .echo ("<option value='HitsByMonth Asc'>月访问量(升序)</option>")
					End If
					If OrderStr = "HitsByMonth Desc" Then
					  .echo ("<option value='HitsByMonth Desc' selected>月访问量(降序)</option>")
					Else
					  .echo ("<option value='HitsByMonth Desc'>月访问量(降序)</option>")
					End If
					
					
					If OrderStr = "OrderID Asc" Then
					 .echo ("<option value='OrderID Asc' selected>手工排序(升序)</option>")
					Else
					 .echo ("<option value='OrderID Asc'>手工排序(升序)</option>")
					End If
					If OrderStr = "OrderID Desc" Then
					  .echo ("<option value='OrderID Desc' selected>手工排序(降序)</option>")
					Else
					  .echo ("<option value='OrderID Desc'>手工排序(降序)</option>")
					End If

		.echo "         </select>"
				
		.echo "</span></td>"
		 .echo "             <td height=""30"">" &ReturnOpenTypeStr(OpenType) & "</td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">标题字数"
		 .echo "               <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:50px;"" value=""" & TitleLen & """ > 简介字数： <input name=""IntroLen"" id=""IntroLen"" class=""textbox"" onBlur=""CheckNumber(this,'简介字数');"" type=""text""    style=""width:50px;"" value=""" & IntroLen & """ > <span class='tips'>一个汉字=两个英文字符</span>"
		 .echo "             </td>"
		 .echo "             <td height=""30"">触发切换模式<select name='trigger' id='trigger'>"
		 .echo "<option value='click'"
		 if trigger="click" then .echo "selected"
		 .echo ">click[鼠标点击]</option>"
		 .echo "<option value='mouseover'"
		 if trigger="mouseover" then .echo "selected"
		 .echo ">mouseover[鼠标悬停]</option>"
		 .echo "</select></td>"
		 .echo "           </tr>"
		 
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">文字层高度"
		 .echo "               <input name=""txtHeight"" class=""textbox"" type=""text"" id=""txtHeight"" style=""width:50px;"" value=""" & txtHeight & """><span class='tips'>(单位像素),0表示隐藏文字层,省略设置或'default'即为默认高度</span></td>"
		 .echo "             <td height=""30"">是否保留边框(有的话)<select name='wrap' id='wrap'>"
		 .echo "<option value='true'"
		 if wrap="true" then .echo "selected"
		 .echo ">是</option>"
		 .echo "<option value='false'"
		 if wrap="false" then .echo "selected"
		 .echo ">否</option>"
		 .echo "</select></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">是否自动播放<select name='auto' id='auto'>"
		 .echo "<option value='true'"
		 if auto="true" then .echo "selected"
		 .echo ">是</option>"
		 .echo "<option value='false'"
		 if auto="false" then .echo "selected"
		 .echo ">否</option>"
		 .echo "</select>"
		 
		 .echo "            </td><td height=""30"">Loading画面时间<input type='text' name='loadIMGTimeout' id='loadIMGTimeout' class='textbox' style='width:50px' value='" & loadIMGTimeout & "'/><span class='tips'>载入myFocus图片的最长等待时间(单位秒,0表示不等待直接播放)</span></td>"
		 .echo "           </tr>"
		 
		 
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">trigger为'mouseover'模式下的切换延迟"
		 .echo "               <input name=""delay"" class=""textbox"" style='width:50px' type=""text"" id=""delay"" value=""" & delay & """  onBlur=""CheckNumber(this,'间隔时间');""><span class='tips'>单位:毫秒</span>"
		 .echo "             </td>"
		 .echo "             <td height=""30""></td>"
		 .echo "           </tr>"
		.echo "                  </table>"	
		.echo "  </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
	    End With
		End Sub
End Class
%> 

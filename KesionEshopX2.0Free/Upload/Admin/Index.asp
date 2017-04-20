<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../Plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Const CheckNewVersion=true   '是否检测获得官方最新版本信息
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KSCls
Set KSCls = New Admin_Index
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Index
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  CheckChannelStatus
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		
		Sub CheckChannelStatus()
		 if session(KS.SiteSN&"setmodelstatus")<>ChannelNotOnStr then
		 conn.execute("update ks_channel set channelstatus=0 where channelid<100 and channelid in(" & channelNotOnStr & ")")
		 session(KS.SiteSN&"setmodelstatus")=ChannelNotOnStr
		 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
		 end if
		End Sub
		
		Sub SaveSkin()
		  dim adminName:adminName=KS.C("AdminName")
		  dim colorId:colorId=KS.S("ColorID")
		    Dim Doc :set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/adminconfig.xml"))
			Dim Node:Set Node=Doc.documentElement.selectSingleNode("/admin/item[@name='" &adminName & "']")
			
			 if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
			 Set Node=Doc.documentElement.appendChild(Doc.createNode(1,"item",""))
			 Node.attributes.setNamedItem(Doc.createNode(2,"name","")).text=adminname
			 Node.text=colorId
			Doc.Save(Server.MapPath(KS.Setting(3)&"Config/adminconfig.xml"))
			ks.die "ok"
		End Sub

		Public Sub Kesion()
		    Call CheckSetting()
			Select Case KS.G("Action")
			 Case "Main" Call KS_Main()
			 Case "copyright" Call CopyRight()
			 Case "setTips" Call setTips()
			 Case "ajax1" Call ajax1()
			 Case "saveskin" Call saveskin()
			 Case Else  Call KS_Index()
			End Select
		End Sub
		
		
		Sub KS_Index()
		%><!DOCTYPE html>
<html>
<head>
<title><%=KS.Setting(0)%>-网站后台管理系统 Powered by KesionCMS X<%=GetVer%></title>
<meta charset="utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta name="renderer" content="webkit"> 

<link href="images/frame.css" rel=stylesheet>
<script type="text/javascript" src="../ks_inc/jquery.js"></script>
<script type="text/javascript" src="../ks_inc/common.js"></script>
<!--
<script src="../ks_inc/dialog/jquery.artDialog.js?skin=twitter"></script>
<script src="../ks_inc/dialog/plugins/iframeTools.js"></script>
-->
<script src="include/SetFocus.js"></script>
<!--[if IE 6]>
<script src="../js/iepng.js" ></script>
<script >
   EvPNG.fix('div, ul, img, li, input'); 
</script>
<![endif]-->
<script> 
<!--
   //保存复制,移动的对象,模拟剪切板功能
  function CommonCopyCutObj(ChannelID, PasteTypeID, SourceFolderID, FolderID, ContentID)
  {
   this.ChannelID=ChannelID;             //频道ID
   this.PasteTypeID=PasteTypeID;         //操作类型 0---无任何操作,1---剪切,2---复制
   this.SourceFolderID=SourceFolderID;   //所在的源目录
   this.FolderID=FolderID;               //目录ID
   this.ContentID=ContentID;             //文章或图片等ID
  }
  function CommonCommentBack(FromUrl)
  {
    this.FromUrl=FromUrl;             //保存来源页的地址
  }
  //初始化对象实例
 var CommonCopyCut=null;
 var CommonComment=null;
 var DocumentReadyTF=false;
 $(window).load(function(){
    setTimeout('getNewMessage()', 3000);
    fHideFocus("A");
	if (DocumentReadyTF==true) return;
	CommonCopyCut=new CommonCopyCutObj(0,0,0,'0','0');
	CommonComment=new CommonCommentBack(0);
	DocumentReadyTF=true;
});
 
var box=null;
function openWin(title,url,isreload,width,height){ 
	if (width==null) width=760;
	if (height == null) height = 450;
	box=$.dialog.open(url,{ id:'topbox',lock: true, title: title, width: width, height: height, close: function() {
			   if (isreload) {
				 frames['MainFrame'].location.reload();
				}
			  }
	 });
}
function fHideFocus(tName){
	aTag=document.getElementsByTagName(tName);
	for(i=0;i<aTag.length;i++)aTag[i].onfocus=function(){this.blur();};
}
		
function out(src){
 $.dialog.confirm("确定要退出系统吗？",function(){ $.dialog.tips("请稍候，系统正在退出...",1000000);location.href='Login.asp?Action=LoginOut';},function(){});
 }
function modifyPass(){
openWin('修改后台登录密码','user/KS.Admin.asp?Action=SetPass',false,520,265)
}
function getNewMessage(){
  var url = '../user/UserAjax.asp';
  jQuery.get(url,{action:'GetAdminMessage'},function(d){ 
      if (d==0){jQuery('#newmessage').hide();}else{jQuery('#newmessage').show();jQuery('#newmessage').attr("title","有"+d+"条新消息!");}
   });
 }
function setCookieTips(tf){
	var v=0;
	if (tf){ v=0;}else{v=1;} 
	jQuery.ajax({ 
	url: "index.asp",
	cache: false,
	data: "action=setTips&v="+v,
	success: function(d){ if (d!='success'){alert(d);}}});
}
function showleft(id)
{ 
    $("#TabPage li").attr("class","");
	$("#left_tab"+id).attr("class",'curr');
	var dvs=$(".leftbox");
	for (var i=0;i<dvs.length;i++){if (dvs[i].id==('left'+id)){$("#"+dvs[i].id).show('fast');}else{$("#"+dvs[i].id).hide('fast');}}
}

//-->
</script>

</head><!-- oncontextmenu="return false" onselectstart="return false" ondragstart="return false" onbeforecopy="return false" oncopy=document.selection.empty() onselect=document.selection.empty()-->
<body style="overflow:hidden" scroll="no" class="bodyStyle<%=KS.GetAdminSkinID()%>"> 
<script type="text/javascript">
$(window).load(function(){
rz();
 
});
$(document).ready(function(){
    
	 <%If instr(lcase(request.ServerVariables("http_referer")&""),"login.asp")<>0 then%>
	 setTimeout("location.href='index.asp'",10);
	 <%end if%>
	 
	 <%If KS.C("SuperTF")<>"1" Then%>
	 for(var i=1;i<6;i++){
	    if ($("#left"+i).html().replace(/\n/g,'').replace(/ /g,'').length<100){
	       $("#left_tab"+i).remove();
    	}
	}
	 //  $("#TabPage").find("li:first").attr("class","curr");
	  $("#left"+$("#TabPage").find("li:first").attr("id").replace(/left_tab/,'')).show();
	 <%End If%>
	 
  rz();
  $(window).resize(function () { 
   rz();
  });
});


function rz(){
 var h=$(window).height()-$("#topframe").height()-$("#bottomframe").height();
 $("#Container").height(h);
 $(".backmain").height(h);
 $("#mainright").width($(window).width()-180)
 $("#MainFrame").height(h-32);
  
  
 if ($(window).width()<=1024){
 
 }else{
  $(".menucenter2").width($(window).width()-$(".menuleft").width()-$(".menuright").width()-$(".menucenter").width()-100).show();
 };
  
}

var screen1=false;    
function ChangeLeftFrameStatu()    { 
	$('#leftframe').toggle();  
	if(screen1==false){            
	$("#mainright").width($(window).width());        
	screen1=true;            
	$("#co").html('<img src="images/ok.png" align="texttop" style="margin:2px 10px 0px 0px;" />打开左栏');
	}else if(screen1==true){           
	screen1=false; 
	$("#co").html('<img src="images/close.png" align="texttop" style="margin:2px 10px 0px 0px;" />关闭左栏');
	$("#mainright").width($(window).width()-170);   
	}    
}



function swapIt(o) {
	o.blur();
	if (o.className == "current") return false;
  
	var list = document.getElementById("Navigation").getElementsByTagName("a");
	for (var i = 0; i < list.length; i++) {
		if (list[i].className == "current") {
			list[i].className = "";
			document.getElementById(list[i].title).y = -scroller._y;
		}
		if (list[i].title == o.title) o.className = "current";
	}
  
	list = document.getElementById("Container").childNodes;
	for (var i = 0; i < list.length; i++) {
		if (list[i].tagName == "DIV") list[i].style.display = "none";
	}
  
	var top = document.getElementById(o.title);
	top.style.display = "block";
	scrollbar.swapContent(top);
	if (top.y) scrollbar.scrollTo(0, top.y);
  
	return false;
};

$(function(){
	$("#TabPage li").click(function(){
		var t = $(this).text();
		$("#leftTitle").html(""+t+"");
	});
});

</script>


			
<!--右下角提示-->		  
<%If instr(KS.Setting(16),"3")>0 Then%>
	 <style>
		.boxvislist ul li {  line-height:25px; background:url(images/artarrow.gif) no-repeat 0px 10px; padding-left:10px; }
		.boxvislist ul li a{color:#006699;}
	</style>
	<script src="../ks_inc/boxtcshow.js"></script>
		<script type="text/javascript">
					var checkPerSecond=60; //60秒检测一次
					var checkInterval=null;
					$(window).load(function(){
							checkMsg();
							checkInterval=setInterval("checkMsg();", 1000*checkPerSecond);
					});
					function checkMsg(){
							$.ajax({type:"get",async:false,url:"index.asp?Action=ajax1",cache:false,dataType:"html",success:function(d){
								if (d!=''){	
									$('ul[id=shoporder_s245]').html(d)
									$.dialog.notice({
										title: '<img src="images/bg30.png" align="absmiddle"/> 消息提示',
										width: 260,  
										content:$('div[id=righttips]').html(),
										time: 30
									});
								}
							}});			  																																							
					}
					
			   </script>
			   
				<div style="display:none;"  id="righttips">
					<div align="left" class="boxvislist" style=" font-size:14px; color:#006699; width:250px; background:#FFFFFF; overflow:hidden ">
						<ul id="shoporder_s245">
							
						</ul>
					</div>	
					<div style="text-align:right"><input type="checkbox" onClick="stopInterval()">不再提示</div>
				</div>
		<%end if%>
<!--右下角提示结束--->



	<div class="menubox" id="topframe">
		<div class="menuBg"></div>
		<div class="menuleft">
			<img src='images/logo.png'/>
			<span>version X<%=GetVer%></span>
		</div>
		<div class="menucenter">
           <div id='ajaxmsg'> 请稍候,正在执行您的请求...  </div>
			<ul id="TabPage">
				<li<%If KS.S("from")<>"app" then response.write " class=""curr"""%> id="left_tab1" title="内容管理" onClick="javascript:showleft(1);"><a href="javascript:;" class="icon1">内容</a></li>
				<li<%If KS.S("from")="app" then response.write " class=""curr"""%> id="left_tab2" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"subsys1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onclick="javascript:showleft(2);" title="应用操作"><a href="javascript:;" class="icon2">应用</a></li>
				<li id="left_tab3" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"model1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onclick="javascript:showleft(3);" title="模型管理"><a href="javascript:;" class="icon3">模型</a></li>
				<li id="left_tab4" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"lab1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onclick="javascript:showleft(4);" title="标签"><a href="javascript:;" class="icon4">标签</a></li>
				<li id="left_tab5" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"user1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onclick="javascript:showleft(5);" title="用户管理"><a href="javascript:;" class="icon5">用户</a></li>				
				<li id="left_tab6" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"sysset0")>0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onclick="javascript:showleft(6);" title="系统管理"><a  href="javascript:;" class="icon6">设置</a></li>
           </ul>
       </div>
       <div class="menucenter2" style="display:none;">
           <ul>
            <%
				Dim KSAnnounceDisplayFlag:KSAnnounceDisplayFlag=""
			    If Instr(KS.Setting(16),"1")=0 Then
				 Response.Write "<script src=""../ks_inc/time/3.js""></script>"
				 KSAnnounceDisplayFlag=" style=""display:none"""
				End If
				%>
				<span<%=KSAnnounceDisplayFlag%>>
                <iframe scrolling=no src="http://www.kesion.com/websystem/GetofficialInfo.asp" name="ShowAnnounce" id="ShowAnnounce" height="20" width="100%" marginheight="0" marginwidth="0" frameborder="0" align="middle" allowtransparency="true"></iframe>
                </span>
            
              </ul>
		</div>
		<div class="menuright">
		 <ul>
			 <li><a href="">后台首页</a></li>
			 <li><a href="../" target="_blank">PC首页</a></li>
			 <li><a href="../3g" target="_blank">手机版</a></li>
			 <!--<li><a href="../User/User_Message.asp?action=inbox" target="_blank"><i class="icon message"></i><span id='newmessage' style="display:none"><img src="images/bg17.png" style="position:absolute;margin:1px 0px 0px -5px" /></span></a></li>-->
            <%
			If KS.ReturnPowerResult(0, "KMST20000") Then
			%>
			<li><a href="System/KS.CleanCache.asp" target="MainFrame" title="更新缓存">更新缓存</a></li>
		  <%end if%>
			
			<%
			 Dim RS:Set RS = Server.CreateObject("Adodb.Recordset")
			 RS.Open "Select top 1 a.*,b.GroupName,u.userface From (KS_Admin a Inner Join KS_UserGroup B on a.groupid=b.id) inner join KS_User u on a.prusername=u.username Where a.UserName='" & KS.C("AdminName") & "'", Conn, 1, 1
			 If Not RS.EOF Then
				
			%>
			
	        <li class="userinfo">
				<span><%=KS.C("AdminName")%></span>
				<div class="operating">
					<div class="arrow"></div>
					<div class="userinfor clearfix">
						<div class="headimg"><img src="<%=rs("userface")%>" onerror="this.src='images/userface.jpg'"></div>
						<div class="right-info">
							<div class="name"><%=KS.C("AdminName")%><font class="mainColor"><%=RS("groupname")%></font></div>
							<div class="info"><em>登录时间<%=RS("lastlogintime")%></em></div>
							<div class="info"> 登录次数：<em><%=RS("LoginTimes")%>次</em></div>
						</div>
					</div>
					<div class="clear"></div>
					<div class="btn clearfix" style="margin-top:5px;">
					   <%If KS.ReturnPowerResult(0, "KMUA10010") Then%>
						<a href="javascript:void(0)" onClick="modifyPass()" title="修改密码"><i class="myicon icon-password"></i>修改密码</a>
					  <%end if%>
						<a  href="javascript:;" onClick="return out(this)"  title="退出"><i class="myicon icon-out"></i>安全退出</a>
					</div>    
				</div>
			</li>
	        <li id="top-bar">
			   <script>
			   $(function(){
				   $("li.userinfo").hover(function(){
						$(this).find(".operating").show();
					},function(){
						$(this).find(".operating").hide();
					});

			        $("#navbox-trigger").click(function(){
						$("#wrapbg").fadeIn(300);
						$(".navbox").animate({right:"0px"},300);
					});
					$("#wrapbg,.close,.shut").click(function(){
					    closeTool();
					});
					$("#witchTab a").click(function(){
						$(this).addClass("curr").siblings().removeClass("curr");
					});
					
					var i = <%=KS.GetAdminSkinID()%>-1
					$("#witchTab a:eq("+i+")").addClass("curr").siblings().removeClass("curr");
					
			   });
			   function closeTool(){
			  	    $("#wrapbg").fadeOut(300);
			  	    $(".navbox").animate({right:"-320px"},300);
			  	};
				
				var colorId = 1;
				function switchbox(e){
					colorId = e;
					$("body").removeClass().addClass("bodyStyle"+colorId+"");
					
				};
				function saveNavColor(){
					 $.ajax({
					   type: "POST",
					   url: "index.asp",
					   data: "action=saveskin&colorId="+colorId,
					   success: function(data){
							alert('恭喜，保存成功！');
							location.reload();
					   }
					});

				};
				
			   </script>
			   <a id="navbox-trigger">关于</a>
			   <div class="navbox">
                    	<div class="close"></div>
                          <div class="changecolor">
                              <div id="demo">
                                  <div class="name">皮肤设置</div>
                                  <div class="pane"> 
                                      <div class="icons" id="witchTab">
									      
                                          <a onClick="switchbox(1);" href="javascript:void(0);"><em class="c1"></em></a>
                                          <a onClick="switchbox(2);" href="javascript:void(0);"><em class="c2"></em></a>  
                                          <a onClick="switchbox(3);" href="javascript:void(0);"><em class="c3"></em></a> 
                                          <a onClick="switchbox(4);" href="javascript:void(0);"><em class="c4"></em></a> 
                                          <a onClick="switchbox(5);" href="javascript:void(0);"><em class="c5"></em></a> 
                                         
                                      </div> 
                                      <div class="clear"></div>
                                  </div> 
                              </div>
                              <div class="colorbutton">
                                  <a href="javascript:saveNavColor();">保存</a>
                                  <a class="shut" href="javascript:;">关闭</a>
                                  <div class="clear"></div>
                              </div>
                        </div>
    					<div class="clear"></div>
                        <div class="navbox-tiles">
                          <div class="name">版权信息</div>
                          <div style="text-align:left;">
                               

								<div style="text-align:center;margin-top:10px;">
									<span style="cursor:pointer" onClick="window.open('http://www.kesion.com');" title="KESION 官方站">
										
										<img border="0" src="http://www.kesion.com/images/logo.png" width="200">
										
									</span>
									
									<br>
									厦门科汛软件有限公司 <span style="font-size:18px;">©</span> 版权所有 <br>
									<span class="tips">Copyright 2006-<%=year(now)%> kesion.com All Rights Reseved</span>
								</div>
								
								官方网站：<a href="http://www.kesion.com" target="_blank">http://www.kesion.com</a>
								<br>
								技术交流：<a href="http://bbs.kesion.com" target="_blank">http://bbs.kesion.com</a>
								<br/>
								当前版本：
								<span  class="red"><%=KS.Version%></span>
								<%if CheckNewVersion then%>
								<br/>
									官方最新版本：<span id='versioninfo'><script src="http://www.kesion.com/websystem/GetofficialInfo.asp?action=getverbyscript"></script></span>
								<%end if%>
					
				              
								<div style=" border-bottom:1px dashed #ccc; margin-top:10px;"></div>
											
								<div class="detail" style="font-size:13px;margin-top:15px;">
									警告：本软件受著作权法和国际公约的保护，未经授权擅自复制传播本程序的部分或全部，可能受到严厉的民事及刑事制裁，并在法律的许可范围内受到可能的起诉。 
								</div>
			
                           </div>
                        </div>
                    </div>
			</li>
			 <%
			 End If
			 RS.Close: Set RS=Nothing
			%>
	    </ul>
	    <div id="wrapbg"></div>
      </div>

</div>
		
	</div>
	
	<div class="clear"></div>
	<div class="backmain">
	   <div id="leftframe">
        <div class="shadow-bg"></div>
		<div id="Container" class="leftBg2">
		  <div id="News">
			<div class="Scroller-Container">
				<div class="left">
                	
                    <div class="leftTitle" id="leftTitle">内容</div>
					
					<div class="clear"></div>

					<script>
							$(document).ready(function(){
							   <%dim ii
							    for ii=1 to 7
								%>
								/* Slide Toogle */
								$("#left<%=ii%>").find("div.navigation").click(function()
								{
									var arrow = $(this).find("span.arrow");
								  //  $("#left<%=ii%>").find("ul.menu").hide();
									$(this).parent().find("ul.menu").stop().slideDown(300);
									$(this).parent().siblings().find("ul.menu").hide();
									$(this).parents(".expmenu").addClass("on").siblings().removeClass("on");
									//$(this).parent().siblings(".navtitle").removeClass("suibian");
								});
								
								$("#expmenu-freebie .leftbox").find(".expmenu:first").addClass("on");
								
								$(".expmenu li").click(function(){
									$(".expmenu li").removeClass("curr");
									$(this).addClass("curr");
								});
								
								
							 <%next%>
							});
														
					</script>
					
		<%
		Dim SQL,I,ModelXML
		Dim RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,BasicType,ModelEname,ChannelStatus From KS_Channel Where ChannelStatus=1 and ChannelID<>6 Order By OrderID,ChannelID ASC")
		If Not RSC.Eof Then
		  SQL=RSC.GetRows(-1)
		  Set ModelXML=KS.ArrayToxml(SQL,RSC,"row","ModelXML")
		End If
		RSC.Close:Set RSC=Nothing
		
		
		'on error resume next

		If Session("ShowCount")="" Then
		 Response.Write " <ifr" & "ame src=""http://ww" &"w."&"k" &"e" & "si" &"on." & "co" & "m" & "/WebS" & "ystem/Co" & "unt.asp"" scrolling='no' frameborder='0' height='0' width='0'></ifr" &"ame>"
		Session("ShowCount")=KS.C("AdminName")
		End If			
		%>
					
      <div id="content">
		<div id="expmenu-freebie">
								 
						  <!---第一块开始---->
						  <div id="left1" class="leftbox" <%If KS.S("from")="app" then response.write " style=""display:none"""%>>
                             <% 
							 dim XMLStr,FieldXML,Nodek,NodeXML,Fast,Fasturl,Attribute,Role,Fastico,Nodek2,NodeXML2,Mchannelid,ModelEname,Nodekz,BasicType,ModelName,ItemName,N
							 if Not ModelXML Is Nothing Then
							 	N=0
								set NodeXML=ModelXML.documentElement.SelectNodes("row")
								
							 	 For Each Nodek In NodeXML
								    N          = N+1
								    ModelName  = Nodek.SelectSingleNode("@channelname").text
								 	Mchannelid = Nodek.SelectSingleNode("@channelid").text
									BasicType  = KS.ChkClng(Nodek.SelectSingleNode("@basictype").text)
									ItemName   = Nodek.SelectSingleNode("@itemname").text
                                    ModelEname = KS.C_S(BasicType,10)
									
									IF ModelEname<>"" And KS.CheckFile(ModelEname&"/Config.xml") Then
									
									        set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
											FieldXML.async = false
											FieldXML.setProperty "ServerHTTPRequest", true 
											FieldXML.load(Server.MapPath(ModelEname&"/Config.xml"))
											Set NodeXML=FieldXML.DocumentElement.SelectNodes("item")							 
											For Each Nodekz In NodeXML
												Fastico=Replace(Nodekz.SelectSingleNode("Fastico").text,"{#BasicType}",BasicType)
												if instr(lcase(KS.C("ModelPower")),lcase(KS.C_S(Mchannelid,10))& "0")=0 or KS.C("SuperTf")=1 Then
													if Not ModelXML Is Nothing Then
														if ModelXML.documentElement.SelectNodes("row[@channelid="& Mchannelid &" and @channelstatus=1]").length<>0 Then
														%>
															<div class="expmenu">
																	<div class="navigation">
																		<div class="navtitle"><i class="icon leftmodel<%=BasicType%>"></i><%=ModelName%></div>
																		<span class="arrow up"></span>
																	</div>
                                                                    
                                                                    <ul class="menu"<%If N<>1 Then Response.Write " style=""display:none""" %>>
                                                                     <%
																		 Set NodeXML2=Nodekz.SelectNodes("Fastmenu")
																		 For Each Nodek2 In NodeXML2
																		    Dim MyItem:MyItem=Replace(Nodek2.SelectSingleNode("Fast").text,"{#ItemName}",ItemName)
																			Dim MyRole:MyRole=Replace(Nodek2.SelectSingleNode("Role").text,"{#ChannelID}",Mchannelid)
																			Dim MyFolderName:MyFolderName=KS.M_C(MChannelID,26)
																			If KS.IsNul(MyFolderName) Then MyFolderName="栏目"
																			MyItem=Replace(MyItem,"{#FolderName}",MyFolderName)
																			if KS.ReturnPowerResult(MChannelID, MyRole) or MyRole="0" Then 
																			%>
																			<li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=ModelName%> >> <font color=red><%=MyItem %></font>','<%=Nodek2.SelectSingleNode("Attribute").text%>','<%=replace(Nodek2.SelectSingleNode("Fasturl").text,"{#ChannelID}",Mchannelid) %>',<%=MchannelID%>);"><%=MyItem %></a></li><%
																			Response.Write vbcrlf
																			end if
																		 Next
																		 %>
																	</ul>
															</div>        
															<%
														end if	
													end if	
												end if
											 next
									
									
									
								 End If
								  
									
									
								 Next
						END IF
					 %>
				</div>
				 <!---第一块结束--->
							 
							 
							 
							 <!---第六块开始--->
								  <div id="left6" class="leftbox" style="display:none">
								  
								   <%IF instr(lcase(KS.C("ModelPower")),"sysset10")=0 or KS.C("SuperTf")=1 Then%>
								      <!-- 系统设置 Start -->
									  <div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon systemset"></i>系统设置</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu">
												<%If KS.ReturnPowerResult(0, "KMST10001") Then%>
											   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>系统参数配置</font>','SetParam','System/KS.Setting.asp');" title="系统参数配置">系统参数配置</a></li>
											 <%end if%>
                                            
											 <%If KS.ReturnPowerResult(0, "M010001") Then %>
												   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'系统设置 >> <font color=red><%=SQL(3,I)%>栏目管理</font>','Disabled','System/KS.Class.asp');">栏目管理</a> <a href='javascript:void(0)' style="padding-left:10px;color:#444; display:none;" onClick="SelectObjItem1(this,'栏目管理 >> <font color=red>添加栏目</font>','Go','System/KS.Class.asp?Action=Add&FolderID=1','');">添加</a></li>
											 <%End If%>	
											 
											 <%If KS.ReturnPowerResult(0, "KMST10003") Then%>
											 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'系统设置 >> <font color=red>在线支付平台管理</font>','GoSave','System/KS.PaymentPlat.asp');">在线支付平台管理</a></li>
<strong></strong>
											  <%End If%>
                                                
											 <%If KS.ReturnPowerResult(0, "KMST10002") Then%>
											   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>整合系统设置</font>','SetParam','System/KS.API.asp');"  title="整合系统设置">API通用整合设置</a></li>
											 <%end if%>
											 
											  <%If KS.ReturnPowerResult(0, "KMST10017") Then %>
													 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'系统设置 >>  <font color=red>省市管理</font>','Disabled','System/KS.Province.asp');">省市地区管理</a> </li>
											 <%end if%>
                                             
                                             
                                             <%If KS.ReturnPowerResult(0, "M010005") Then%>
													<li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'系统设置 >> <font color=red>文档批量设置</font>','Disabled','System/KS.ItemInfo.asp?Action=SetAttribute');">文档批量设置</a></li>
											 <%End If%>
												  
											<%If KS.ReturnPowerResult(0, "M010006") then%> 
											   <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'系统设置 >> <font color=red>文档回收站</font>','ViewFolder','System/KS.ItemInfo.asp?ComeFrom=RecycleBin','');">文档回收站</a>
											  </li>
											 <%End If%>
											 
											</ul>
										
									</div>
									<!-- 系统设置 End -->
									
									<%
									If KS.CheckDir("../3G/") Then
									If KS.ReturnPowerResult(0, "KSO10000")  Then %>
									 <div class="expmenu">
									   <div class="navigation">
											 <div class="navtitle"><i class="icon telset"></i>手机版参数配置</div>
												<span class="arrow up"></span>
									   </div>
											<ul class="menu" style="display:none">
									 
											   <li><a href="#" onClick="SelectObjItem1(this,'手机版系统管理 >> <font color=red>手机版基本参数设置</font>','SetParam','../3g/Setting.asp');" title="手机版基本参数设置">手机版基本参数设置</a></li>
											   <li><a href="#"  onClick="SelectObjItem1(this,'手机版系统管理 >> <font color=red>手机版自定义页面管理</font>','Disabled','../3g/setting.asp?action=template');">手机版自定义页面</a></li>
									  </ul>
									
									</div>
									<%end if
									Else
									 Check3G
									End If
									%>
									
									
								<%End If%>	
								
						  </div>
						  <!---第六块结束---> 

							 
							 <!---第二块开始--->
                             
								  <div id="left2" class="leftbox"<%If KS.S("from")<>"app" then response.write " style=""display:none"""%>>
								  
								   <%
								    Dim FsoItem 
									Dim FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
									Dim FolderObj:Set FolderObj = FsoObj.GetFolder(Server.MapPath("plus"))
									Dim SubFolderObj:Set SubFolderObj = FolderObj.SubFolders
									Dim ItemNum:ItemNum=0
									For Each FsoItem In SubFolderObj
									   if KS.CheckFile("plus/"&FsoItem.name&"/Config.xml") then
										   set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
										   FieldXML.async = false
										   FieldXML.setProperty "ServerHTTPRequest", true 
										   FieldXML.load(Server.MapPath("plus/"&FsoItem.name&"/Config.xml"))
										   if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
											Set NodeXML=FieldXML.DocumentElement.SelectSingleNode("App")	
											If Not NodeXML Is Nothing Then
												Dim AppName:AppName=NodeXML.SelectSingleNode("AppName").Text
												Dim AppStatus:AppStatus=NodeXML.SelectSingleNode("AppStatus").Text
												 Role=NodeXML.SelectSingleNode("Role").Text
												
												If AppStatus="1" And (instr(1,lcase(KS.C("ModelPower")&KS.C("PowerList")),Role,1)<>0 or KS.C("SuperTf")=1) Then
												  ItemNum=ItemNum+1
												  %>
												  <div class="expmenu">
														<div class="navigation">
															<div class="navtitle"><i class="icon app_<%=Role%>"></i><%=AppName%></div>
															<span class="arrow up"></span>
														</div>
													   
														<ul class="menu"<%if ItemNum>1 then Response.Write (" style='display:none'")%>>
														
														   <%
															 Set NodeXML2=NodeXML.SelectNodes("AppItem")
															 For Each Nodek2 In NodeXML2
																	 MyItem=Nodek2.SelectSingleNode("ItemName").text
																	 MyRole=Nodek2.SelectSingleNode("Role").text
																			
																	 if KS.ReturnPowerResult(0, MyRole) or MyRole="0"  or KS.C("SuperTf")=1 Then 
																			%>
																			<li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=AppName%> >> <font color=red><%=MyItem %></font>','<%=Nodek2.SelectSingleNode("Attribute").text%>','<%IF LEFT(Nodek2.SelectSingleNode("ItemUrl").text,1)<>"/" THEN Response.Write("plus/")%><%=Nodek2.SelectSingleNode("ItemUrl").text%>',0);"><%=MyItem %></a></li><%
																			Response.Write vbcrlf
																	 end if
															 Next
														  %>
														
															
														</ul>
												</div>
												  <%
												
										    End If
										   End If
										   End If
									  End If
									 Next
								   
								   %>
								  </div>
							 <!---第二块结束---> 
							  
							  
							  <!---第三块开始--->
								  <div id="left3" class="leftbox" style="display:none">
								    <!-- 模型管理 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon modeler"></i>模型管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu">
												<li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>模型管理首页</font>','Disabled','System/KS.Model.asp');">模型管理首页</a></li>
												 <%If KS.ReturnPowerResult(0, "KSMM10000") Then%>
												 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>添加新模型</font>','Go','System/KS.Model.asp?action=Add');">添加新模型</a></li>
												 <%end if%>
												 <%If KS.ReturnPowerResult(0, "KSMM10004") Then%>
												 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>模型信息统计</font>','Go','System/KS.Model.asp?action=total');">模型信息统计</a></li>
												 <%end if%>
												 
											</ul>
										
									</div>
									<!-- 模型管理 End -->
									
									<%If KS.ReturnPowerResult(0, "KSMM10003") Then%>
									<!-- 模型字段管理 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon text"></i>模型字段管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu" style="display:none">
												  <%For I=0 To UBound(SQL,2)
												   if KS.ChkClng(SQL(4,I))<=10 AND SQL(0,I)<>6 and SQL(0,I)<>9 and SQL(0,I)<>10 Then
												  %>
													 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'模型管理 >> <font color=red>字段管理</font>','Disabled','system/KS.Field.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=SQL(1,I)%>字段</a></li>					  
												  <%
												  End iF
									Next%>
											</ul>
										
									</div>
									<!-- 模型字段管理 End -->
								  
								   	
									<!-- 管理列表菜单 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon list"></i>管理列表菜单</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu" style="display:none">
												<%For I=0 To UBound(SQL,2)
												   if SQL(6,I)=1 AND SQL(4,I)<9 Then
												  %>
													 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'模型管理 >> <font color=red>管理列表管理</font>','Disabled','system/KS.Model.asp?action=ManageMenu&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=SQL(1,I)%>列表菜单</a></li>					  
												  <%
												  End iF
												 Next%>
											</ul>
										
									</div>
									<!-- 管理列表菜单 End -->
								 <%end if%>
								 
								  </div>
							 <!---第三块结束---> 
							 
							 
							 <!---第四块开始--->
								  <div id="left4" class="leftbox" style="display:none">
							 <%
							 IF KS.ReturnPowerResult(0, "KMTL10001") or KS.ReturnPowerResult(0, "KMTL10002") OR KS.ReturnPowerResult(0, "KMTL10003")  OR KS.ReturnPowerResult(0, "KMTL10011") OR KS.ReturnPowerResult(0, "KMTL10001") OR KS.ReturnPowerResult(0, "KMTL10004") OR KS.ReturnPowerResult(0, "KMTL10005") or KS.ReturnPowerResult(0, "KMSL10008") Or  KS.ReturnPowerResult(0, "KMSL10009") THEN
						    %>
								    <!-- 标签管理 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon tips"></i>标签管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu">
												<%
												With Response
												If KS.ReturnPowerResult(0, "KMTL10001") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>系统函数标签</font>','FunctionLabel','Include/Label_Main.asp?LabelType=0');"">系统函数标签</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMTL10002") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义SQL函数标签</font>','DiyFunctionLabel','Include/Label_Main.asp?LabelType=5');"">自定义SQL函数标签</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMTL10003") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义静态标签</font>','FreeLabel','Include/Label_Main.asp?LabelType=1');"">自定义静态标签</a></li>")
												End If
												
												If KS.ReturnPowerResult(0, "KMTL10011") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义生成XML文档</font>','DiyFunctionLabel','Include/Label_Main.asp?LabelType=7');"">自定义生成XML文档</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMTL10001") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>分页样式管理</font>','Disabled','Include/Label_Main.asp?LabelType=100');"">分页样式管理</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMTL10004") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义JS管理</font>','SysJSList','include/JS_Main.asp?JSType=0');"">系统JS管理</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMTL10005") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义JS管理</font>','FreeJSList','include/JS_Main.asp?JSType=1');"">自定义JS管理</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMSL10008") Then
												  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>生成顶部菜单</font>','SetParam','include/ClassMenu.asp');"">生成顶部菜单</a></li>")
												 ' .Write "<li><a href='include/ClassMenu.asp'  target='MainFrame' title='生成顶部菜单'>生成顶部菜单</a></li>"
												end if
												If KS.ReturnPowerResult(0, "KMSL10009") Then
												  .Write "<li><a href='include/TreeMenu.asp'  target='MainFrame' title='生成树形菜单'>生成树形菜单</a></li>"
												End If
							
											End With %>	
											</ul>
										
									</div>
									<!-- 标签管理 End -->
								<%END IF%>
								<%
							 IF KS.ReturnPowerResult(0, "KMTL10006") or KS.ReturnPowerResult(0, "KMTL10007") THEN
						    %>	
									<!-- 模板管理 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon template"></i>模板管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu" style="display:none">
												<%
												If KS.ReturnPowerResult(0, "KMTL10006") Then
													Response.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>自定义页面管理</font>','Disabled','System/KS.DIYPage.asp');"">自定义页面管理</a></li>")
												End If
												If KS.ReturnPowerResult(0, "KMTL10007") Then
													Response.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>所有模板管理</font>','Disabled','System/KS.Template.asp');"">所有模板管理</a></li>")
												End If
												%>
											</ul>
										
									</div>
									<!-- 模板管理 End -->
							<%END IF%>
								  </div>
							 <!---第四块结束---> 
							 
							 
							 <!---第五块开始--->
								  <div id="left5" class="leftbox" style="display:none">
                                  <%If KS.ReturnPowerResult(0, "KMUA10002") or KS.ReturnPowerResult(0, "KMUA10016") or KS.ReturnPowerResult(0, "KMUA10004") or KS.ReturnPowerResult(0, "KMUA10003") or KS.ReturnPowerResult(0, "KMUA10009") or KS.ReturnPowerResult(0, "KMUA10012") or KS.ReturnPowerResult(0, "KMUA10013") or KS.ReturnPowerResult(0, "KMUA10015") or KS.ReturnPowerResult(0, "KSMS20007") or KS.ReturnPowerResult(0, "KMUA10011") Then%>
								    <!-- 用户管理 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon usermanagement"></i>用户管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu">
												  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>注册用户管理</font>','Disabled','User/KS.User.asp');" title="注册用户管理">注册用户管理</a></li>
												   <%If KS.ReturnPowerResult(0, "KMUA100027") Then%>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>添加用户</font>','Disabled','User/KS.User.asp?Action=Add');" title="添加用户">添加用户</a></li>
												   <%end if%>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10016") and IsBusiness Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>实名认证管理</font>','Disabled','User/KS.UserRZ.asp');" title="实名认证管理">实名认证管理</a></li>
												  <%end if%>
												  
												  <%If KS.ReturnPowerResult(0, "KMUA10004") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>用户组管理</font>','Disabled','User/KS.UserGroup.asp');" title="用户组管理">用户组管理</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10003") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>用户短信管理</font>','Disabled','User/KS.UserMessage.asp');" title="用户短信管理">用户短信管理</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10009") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>发送邮件管理</font>','Disabled','User/KS.UserMail.asp');" title="发送邮件管理">发送邮件管理</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10012") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员字段管理</font>','Disabled','System/KS.Field.asp?ChannelID=101');" title="会员字段管理">会员字段管理</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10013") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员表单管理</font>','Disabled','User/KS.UserForm.asp');" title="会员表单管理">会员表单管理</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10015") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员使用记录</font>','Disabled','User/KS.UserUseLog.asp');" title="会员使用记录">会员使用记录</a></li>
													<li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员签到管理</font>','Disabled','User/KS.qiandao.asp');" title="会员签到管理">会员签到管理</a></li>
												  <%end if%>
												  
												   <%If KS.ReturnPowerResult(0, "KSMS20007") Then%>
													<li><a href="User/KS.PromotedPlan.asp"  target="MainFrame" title="推广计划管理">推广计划管理</a></li>
													<%end if%>
												
												   <%If KS.ReturnPowerResult(0, "KMUA10011") Then%>
				                                    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'用户管理 >> <font color=red>稿件统计</font>','SetParam','User/KS.UserProgress.asp');">会员稿件统计</a></li>
			                                       <%End If%>
												  
											</ul>
										
									</div>
									<!-- 用户管理 End -->
                                    <%end if%>
									
                                    <%If KS.ReturnPowerResult(0, "KMUA10001") or KS.ReturnPowerResult(0, "KMST10006") or KS.ReturnPowerResult(0, "KMUA10010") Then%>
									<!-- 管理员管理 Start -->
									<div class="expmenu">
									<div class="navigation">
												<div class="navtitle"><i class="icon management"></i>管理员管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu"  style="display:none">
											     <%If KS.ReturnPowerResult(0, "KMUA10001") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>管理员管理</font>','Disabled','User/KS.Admin.asp');">管理员管理</a></li>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>管理员角色</font>','Disabled','User/KS.Admin.asp?action=Role');">管理员角色</a></li>
												  <%end if%>
												  
												 <%If KS.ReturnPowerResult(0, "KMST10006") Then%>
													 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>管理员登录记录</font>','Disabled','User/KS.Log.asp');">管理员登录记录</a></li>
												 <%end if%>
                                                
									       </ul>
									</div>
									<!-- 管理员管理 Start -->
                                    <%End If%>
									
									
									<!-- 账务明细管理 Start -->
									<%If KS.ReturnPowerResult(0, "KMUA10005") or KS.ReturnPowerResult(0, "KMUA10006") or  KS.ReturnPowerResult(0, "KMUA10007") or KS.ReturnPowerResult(0, "KMUA10017") or  KS.ReturnPowerResult(0, "KMUA10008") Then %>
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon pay"></i>账务明细管理</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu" style="display:none">
												<%If KS.ReturnPowerResult(0, "KMUA10005") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员点券明细</font>','Disabled','User/KS.LogPoint.asp');" title="会员点券明细">会员点券明细</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10006") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员有效期明细</font>','Disabled','User/KS.LogEdays.asp');" title="会员有效期明细">会员有效期明细</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10007") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员资金明细</font>','Disabled','User/KS.LogMoney.asp');" title="会员资金明细">会员资金明细</a></li>
												  <%End If%>
												  <%If KS.ReturnPowerResult(0, "KMUA10017") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员积分明细</font>','Disabled','User/KS.LogScore.asp');" title="会员积分明细">会员积分明细</a></li>
												  <%end if%>
												  <%If KS.ReturnPowerResult(0, "KMUA10008") Then %>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>线下充值卡管理</font>','Disabled','User/KS.Card.asp?cardtype=0');" title="线下充值卡管理">线下充值卡管理</a></li>
												  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>线上充值卡管理</font>','Disabled','User/KS.Card.asp?cardtype=1');" title="线上充值卡管理">线上充值卡管理</a></li>
												  <%end if%>
											</ul>
									</div>
									<%End If%>
									<!-- 账务明细管理 End -->
									
									<%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
									<!-- 快速查找用户 Start -->
									<div class="expmenu">
											<div class="navigation">
												<div class="navtitle"><i class="icon search"></i>快速查找用户</div>
												<span class="arrow up"></span>
											</div>
											<ul class="menu" style="display:none">
												<li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','User/KS.User.asp?UserSearch=5');" style=" color:#ff6600">24小时内登录</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','User/KS.User.asp?UserSearch=6');">24小时内注册</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','User/KS.User.asp?UserSearch=1');"> 被锁住的用户</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','User/KS.User.asp?UserSearch=3');">待审批会员</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','User/KS.User.asp?UserSearch=4');">待邮件验证</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','User/KS.User.asp?UserSearch=2');">所有管理员用户</a></li>
											</ul>
									</div>
									<!-- 快速查找用户 End -->
									<%end if%>
									
								  </div>
							 <!---第五块结束---> 
							 
							 
							 
								
							</div>
						</div>
										
					
				</div>

			</div>
		  </div>
		</div>

	</div>

		
		<div class="right" id="mainright">
			<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
			 <tr>
			  <td><%
			   Dim MainUrl:MainUrl="?action=Main"
			   If Not KS.IsNul(Session("FromFile")) Then
			       MainUrl=Session("FromFile")
				   if request("from")="app" then MainUrl=MainUrl &"#app"
				   Session("FromFile")=""
			   End If
			 %><iframe src="<%=MainUrl%>" noresize name="MainFrame" id="MainFrame" frameborder="no" scrolling="auto"  marginwidth="0"  marginheight="0" width="100%" height="500"></iframe>
</td>
			 </tr>
			 <tr>
			 <td height="25"><iframe src="Post.Asp?ButtonSymbol=Disabled&OpStr=<%=Server.URLEncode("系统管理中心 >> 首页")%>" name="BottomFrame" ID="BottomFrame" frameborder="no" height="25"  scrolling="no" width="100%" marginwidth="0" marginheight="0"></iframe></td>
			 </tr>
			</table>
			
		</div>
	</div>
	
     
    
    
	
	<div class="footer" id="bottomframe">
		<div class="left"><a href="#" id="co" onClick="ChangeLeftFrameStatu();" title="全屏/半屏"><img src='images/close.png' align="texttop" style="margin:2px 10px 0px 0px;" />关闭左栏</a></div>
		<div class="center">

			<%
			With Response
			    .Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'系统管理中心 >> <font color=red>首页</font>','disabled','index.asp?action=Main');"">后台首页</a>"

         IF KS.ReturnPowerResult(0, "ref20000") or KS.C("SuperTf")=1 Then

				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'发布中心 >> <font color=red>发布管理首页</font>','disabled','Include/refreshindex.asp');"">发布首页</a>"
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'发布中心 >> <font color=red>发布管理首页</font>','disabled','Include/RefreshHtml.asp?ChannelID=1');"">发布管理</a>"
		End If
				
				If KS.ReturnPowerResult(0, "KMTL10007") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>模板管理</font>','disabled','System/KS.Template.asp');"">模板管理</a>"
				End If
				If KS.ReturnPowerResult(0, "KMST10001") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'系统设置 >> <font color=red></font>','SetParam','System/KS.Setting.asp');"" title='基本信息设置'>系统配置</a>"
				End If
				If Instr(KS.C("ModelPower"),"model1")>0 Or KS.C("SuperTF")="1" then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'模型管理 >> <font color=red>模型管理首页</font>','SetParam','System/KS.Model.asp');"">模型管理</a>"
				End If
				
				
				If KS.ReturnPowerResult(0, "M010007") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'内容管理 >> <font color=red>一键快速生成HTML</font>','Disabled','include/refreshquick.asp');"">一键生成HTML</a>"
				 end if
				
			End With
			%>
		</div>
		<div class="right">
				版权所有 &copy; 2006-<%=year(Now)%> 厦门科汛软件有限公司
		</div>
	</div>	
    

	<script src="images/jquery.nicescroll.js"></script>
    <script>
	
		$(function(){
			
			$("#Container").niceScroll({  
			cursorcolor:"#ffffff",  
			cursoropacitymax:1,  
			touchbehavior:false,  
			cursorwidth:"5px",  
			cursorborder:"0",  
			cursorborderradius:"15px"  
			}); 
		});
        
        
    </script>

    
    
    
</body>
</html>
<% End Sub

Sub KS_Main()
		   Dim TipStr,SafetyTips:SafetyTips=KS.ReadSetting(0)
		   If SafetyTips="1" Then
			   If EnableSiteManageCode=false Then
				TipStr="<li style=""height:24px;line-height:24px"">您没有启用管理认证码，建议您打开conn.asp将EnableSiteManageCode的值设置为True；</li>"
			   ElseIf SiteManageCode="8888"  Then
				TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您后台管理认证密码为系统默认值：<span style=""color:red"">8888</span>,建议您及时打开conn.asp里修改；</li>"
			   End If
			   If KS.CheckDir("../admin") Then
		       TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您的网站后台管理目录为：<span style=""color:red"">admin </span>，出于安全的考虑，我们建议您修改目录名；</li>"
			   End If
			   
			   If DataBaseType=0 Then
			    If instr(lcase(DBPath),"ks_data/kesioncms9.mdb")<>0 then
		         TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您的数据库名称为系统默认名称：<span style=""color:red"">" & DBPath & "</span>,出于安全的考虑，我们建议您修改数据库名称；</li>"
				end if
			   End If
			   
			   If Lcase(KS.C("AdminPass"))="469e80d32c0559f8" Then
		         TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您的后台管理员密码为系统默认值：<span style=""color:red"">admin888</span>,出于安全的考虑，我们建议您及时修改后台登录密码；</li>"
			   End If
			   
			   If TipStr<>"" Then
		    TipStr=TipStr & "<div style=""margin-top:16px;margin-bottom:20px;text-align:right""><label style=""color:#999""><input onclick=""parent.setCookieTips(this.checked)""  type=""checkbox"" name=""nottips"" id=""notips"" value=""1"">我知道了，下次进入后台不再提醒</label></div>"
			   End If
		   End If
 %><!DOCTYPE html>
<html>
<head>
<title>KesionCMS网站管理系统</title>
<meta charset="utf-8" />
<script src="../ks_inc/jquery.js"></script>
<script language='JavaScript' src='../KS_Inc/common.js'></script>
<link href="images/main.css" rel=stylesheet>
<script type="text/javascript">
			function showbigpic(){
				var box=top.$.dialog({title:'KESION公司相关证书：',content: '<style>.zs{width:890px;}.zs li img{border:1px solid #000;margin:5px;width:199px;height:220px;}.zs li{width:200px;float:left;margin:10px;}</style><div class="zs"><ul><li><a href="http://www.kesion.com/images2015/wanneng1.jpg" target="_blank"><img src="http://www.kesion.com/images2015/wanneng.jpg" title="科汛万能建站管理系统 KesionCMS V1.5版权登记证书"/></a></li><li><a href="http://www.kesion.com/images2015/kesionwangxiao.jpg" target="_blank"><img src="http://www.kesion.com/images2015/kesionwangxiao1.jpg"  title="科汛在线网校系统著作权"/></a></li><li><a href="http://www.kesion.com/images2015/kesionwfx.jpg" target="_blank"><img src="http://www.kesion.com/images2015/kesionwfx1.jpg"  title="科汛在线微分销管理系统版权登记证书"/></a></li><li><a href="http://www.kesion.com/images2015/zhineng.jpg" target="_blank"><img src="http://www.kesion.com/images2015/zhineng1.jpg"  title="KesionICMS着作权证书"/></a></li><li><a href="http://www.kesion.com/images2015/imallshop.jpg" target="_blank"><img src="http://www.kesion.com/images2015/imallshop1.jpg"  title="KesionIMALL着作权证书"/></a></li><li><a href="http://www.kesion.com/images2015/kaoshi.jpg" target="_blank"><img src="http://www.kesion.com/images2015/kaoshi1.jpg"  title="KesionIEXAM着作权证书"/></a></li><li><a href="http://www.kesion.com/images/zs/kesionr.jpg" target="_blank"><img src="http://www.kesion.com/images/zs/kesionr.jpg" title="KESION商标证书"/></a></li><li><a href="http://www.kesion.com/images/zs/kesioncmsr.jpg" target="_blank"><img src="http://www.kesion.com/images/zs/kesioncmsr.jpg" title="KesionCMS商标证书"/></a></li></ul></div>',max:false,min: false});
			}
			$(window).load(function(){
			  <%If SafetyTips="1" and TipStr<>"" Then%>
			    top.$.dialog({title:'<span style="font-weight:bold;font-size:16px"><i class="icon back"></i>安全提醒</span>',content:'<div style="font-size:12px;height:160px;"><br/><ul><%=TipStr%></ul></div>'});
			  
			  <%End If%>
			 <%if CheckNewVersion Then%>
			  <%if request.ServerVariables("SERVER_NAME")<>"localhost" and request.ServerVariables("SERVER_NAME")<>"127.0.0.1" then%>
			  $.get('index.asp',{timestamp:new Date().getTime(),action:'copyright'},function(d){$('#currversion').html(unescape(d))});
			 <%End If%>
			  //检查是否存在升级文件
			  $.ajax({
			  url: "System/KS.Update.asp",
			  cache: false,
			  data: "action=check",
			  success: function(d){
			        d=unescape(d);
					switch (d){
					 case 'enabled':
					  $("#updateInfo").html("<font class='red'>对不起,您没有开启自动检测最新版本功能!</font>");
					  break;
					 case 'false':
					  $("#updateInfo").html("<font class='green'>当前已经是最新版本!</font>");
					  break;
					 case 'localversionerr':
					  $("#updateInfo").html("<font class='red'>加载本地xml版本文件出错,请检查<%=KS.Setting(89)%>include/version.xml文件是否存在!</font>");
					  break;
					 case 'remoteversionerr':
					  $("#updateInfo").html("<font class='red'>读取服务器文件出错,请检查<%=KS.Setting(89)%>System/KS.Update.asp文件的配置是否正确或稍候再试!</font>");
					  break;
					 case 'unallow':
					  $("#updateInfo").html("<font class='red'>系统检查到有可更新文件,但不支持在线升级,请到官方站(<a href='http://www.kesion.com' target='_blank'>www.kesion.com</a>)下载升级文件!</font>");
					  break;
					 case 'unallowversion':
					  $("#updateInfo").html("<font class='red'>系统检查到有可更新文件,但由于您的版本号与最新版本号不对应,不支持在线升级,请根据您当前使用的版本到官方站(<a href='http://www.kesion.com' target='_blank'>www.kesion.com</a>)下载升级文件手工升级!</font>");
					  break;
					 default:
					    $("#updateInfo").html("<font class='red'>系统检查到有可升级文件!</font>");
						top.openWin('<i class="icon back"></i>KesionCMS 升级提醒','System/KS.Update.asp?action=showupdateinfo',false,800,400)
					  break;
					}
			  }
		 	 });
			  <%End If%>
			 });
           </script>

		  <script type="text/javascript" charset="gbk">
			
			 function cpboxbut(obj,csk){
				 if (csk=="ok")
				 {
				 $(obj).attr("class","Cp_boxG")
				 }
				 else
				 {
				 $(obj).attr("class","Cp_box")
				 }
			 }
			function add_mk(){
				$('#cpboxadd').show(300)
			}
			var boxi='';
			function sel_mk(){
			  boxi=top.$.dialog.open('System/ks.index_mk.asp',{
				title:'选择快捷', 
				width: '800px',
				height: '500px',
				});
			
			}
			function delbox(obj,id){
				 top.$.dialog.confirm('确定要删除快捷菜单操作吗？',function(){ del_mk(obj,id); },function(){});
			}
			function del_mk(obj,id){
				//$(obj).parent().parent().hide(500)
				$(obj).parent().parent().attr("title","del")
				$.get("System/ks.index_mk.asp",{id:id,action:"del_mk",anticache:Math.floor(Math.random()*1000)},function(d){//读取列表
				});
				$(obj).parent().parent().remove();
			}
           </script>
		   <%
			 dim XMLStr,FieldXML,Nodek,NodeXML,Fast,Fasturl,Attribute
			 call indexXMLField(KS.C("AdminName"))
			 set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		 	 FieldXML.async = false
		 	 FieldXML.setProperty "ServerHTTPRequest", true 
		 	 FieldXML.load(Server.MapPath(KS.Setting(3)&"config/filmk/"& KS.C("AdminName") &"_mk_a.xml"))
		 	 Set NodeXML=FieldXML.DocumentElement.SelectNodes("item")
			%>
            
            <script>
				function indexAuto(){
					
					$(".indexMain .left").width($("body").width()-$(".indexRight").width()-60);
					
				};
				$(function(){
					$(".indexMain .left").width($("body").width()-$(".indexRight").width()-40);
					$(window).resize(function(){
						indexAuto();
					});
					
					$(".todayData li").each(function() {
                        var i = $(this).index();
						$(this).find(".dot").addClass("dot"+i+"");
                    });
					
				});
			</script>
</head>
<body>

	
    <div class="indexMain">
    	
    	<div class="clear"></div>

        <div class="left">
            <div class="stateBox">
                <div id="currversion"></div>
                <div class="text">KesionCMS<sup>&reg;</sup>系统由厦门科汛软件有限公司(<a href="http://www.kesion.com" target="_blank">Kesion.Com</a>)独立开发，软件著作权登记号：<a href="javascript:;" onClick="showbigpic()">2016SR010313</a>。任何个人或组织不得在授权允许的情况下删除、修改、拷贝本软件及其它副本上一切关于版权的信息。
				  <div style="margin-top:20px;"><strong >当前版本：</strong>
				     <span  class="red"><%=KS.Version%></span>
					<%if CheckNewVersion then%>
						(官方最新版本：<span id='versioninfo'><script src="http://www.kesion.com/websystem/GetofficialInfo.asp?action=getverbyscript"></script></span>)
					<%end if%>
					<strong>在线升级：</strong><span id='updateInfo'>正在检测最新版本信息...</span>
					</div>
				</div>
            </div>
           
            
            <div class="shortcut">
                <div class="titleStyle">快捷通道</div>

				<%
				dim i
				For Each Nodek In NodeXML
				Fast=Nodek.SelectSingleNode("Fast").text 
				Fasturl=Replace(Nodek.SelectSingleNode("Fasturl").text,"|Fast|","&")
				Attribute=Nodek.SelectSingleNode("Attribute").text
				i=i+1
				%>
				<div id="cpbox_<%=i%>"  class="Cp_box" onMouseOver="cpboxbut(this,'ok');"  onmouseout="cpboxbut(this,'no');">
				<a style=" text-decoration:none;" id="url_<%=Nodek.SelectSingleNode("@id").text%>" href="javascript:void(0)" onClick="SelectObjItem1(this,'<%=Fast%> >> <font color=red><%=Fast%></font>','<%=Attribute%>','<%=Fasturl%>')" ><img id="img_<%=Nodek.SelectSingleNode("@id").text%>" src="<% =Nodek.SelectSingleNode("Fastico").text %>"/></a><a class="name" id="url_<%=Nodek.SelectSingleNode("@id").text%>" href="javascript:void(0)" onClick="SelectObjItem1(this,'<%=Fast%> >> <font color=red><%=Fast%></font>','<%=Attribute%>','<%=Fasturl%>')" ><%=Fast %></a>
				<span class="delico"><img title="删除快捷" src="images/mk_del.png" onClick="delbox(this,'<%=Nodek.SelectSingleNode("@id").text%>')"/></span>
				
				</div>
				<%
				
				Next%>
				
				<div id="cpboxadd" onMouseOver="cpboxbut(this,'ok');"  onmouseout="cpboxbut(this,'no');" class="Cp_box">
				<a href="javascript:void(0)" onClick="sel_mk();" title="添加快捷" ><img src="images/ffgsadsadf-.png" style="border:0px; margin-top:8px;" /></a>
				</div>
				<div class="clear"></div>
                
                
            </div>
            
            <div class="dataStatistics">
                <div class="titleStyle">数据统计</div>
                
                <script type="text/javascript">
                    $(function () {
                        // Create the chart
                        Highcharts.chart('dataContainer', {
                            chart: {
                                type: 'column'
                            },
                            title: {
                                text: ''
                            },
                            subtitle: {
                                text: ''
                            },
                            xAxis: {
                                type: 'category'
                            },
                            yAxis: {
                                title: {
                                    text: ''
                                }
                    
                            },
                            legend: {
                                enabled: false
                            },
                            plotOptions: {
                                series: {
                                    borderWidth: 0,
                                    dataLabels: {
                                        enabled: true,
                                        format: '{point.y:.0f}'
                                    }
                                }
                            },
                    
                            tooltip: {
                                headerFormat: '',
                                pointFormat: '<span style="color:{point.color}">{point.name}</span>: 共<b>{point.y:.0f}</b>篇<br/>'
                            },
                           credits: {
							 enabled: false
						   },
                            series: [{
                                name: '系统',
                                colorByPoint: true,
								
								data: [
								<%
								dim SQLArr,ii
								dim rsm:set rsm=server.CreateObject("adodb.recordset")
								rsm.open "select ChannelID,ChannelName,ChannelTable From KS_Channel Where basictype<=9 and ChannelStatus=1 order by orderid",conn,1,1
								if not rsm.eof then
								  	SQLArr=rsm.getrows(-1)
								end if
								
								 if isarray(SQLArr) then
								for ii=0 to ubound(sqlarr,2)
								 response.write "{" & vbcrlf
								 response.write "name: '"&sqlArr(1,ii) &"'," & vbcrlf
                                 response.write " y:" & conn.execute("select count(1) From " & sqlArr(2,ii))(0) &"," &vbcrlf
                                 response.write " drilldown: '"&sqlArr(1,ii) &"'"&vbcrlf
								 response.write "}"
								 if (ii<>ubound(sqlarr,2)) then response.write ","&vbcrlf
								next 
							  end if
								
								%>
								
                              ]
                            }]
                    
                        });
                    });
                </script>
    
                <script src="images/highcharts.js"></script>
                <div class="clear blank20"></div>
                <div id="dataContainer" style="min-width: 310px; height: 350px; margin: 0 auto"></div>
    
                
            </div>
        
        </div><!--left-->

        <div class="right indexRight">
        	
            <div class="todayData">
            	<div class="titleStyle">今日业务量</div>
                <ul>
				  <%
				  if isarray(SQLArr) then
				    for ii=0 to ubound(sqlarr,2)
					  dim dateField:dateField="adddate"
					  if sqlArr(0,ii)=9 then
					     dateField="date"
					  end if
					 response.write "<li><span>" & conn.execute("select count(1) From " & sqlArr(2,ii) & " where year(" & dateField &")=" & year(now) & " and day(" & dateField &")=" & day(now) )(0) & "</span><i class=""dot""></i>"&sqlArr(1,ii) &"</li>"
					
					next 
				  end if
				  
				 
					%>
                	
                </ul>
            </div>
            <div class="todayData posts"<%If instr(KS.Setting(16),"2")<=0 Then response.write " style='display:none'"%>>
            	<div class="titleStyle">技术交流新帖</div>
                <ul>
                	 <script charset="utf-8" id="showtopic" src="http://bbs.kesion.com/Dv_News.asp?GetName=newtopic"></script>
                </ul>
            </div>
            
        </div><!--right-->
                
        
			
	</div>	



	<script src="images/jquery.nicescroll.js"></script>
    <script>
	
		$(function(){
			
			$("body").niceScroll({  
			cursorcolor:"#000000",  
			cursoropacitymax:1,  
			touchbehavior:false,  
			cursorwidth:"5px",  
			cursorborder:"0",  
			cursorborderradius:"15px"  
			}); 
		});
        
        
    </script>

			


<div  class="rightmain" style="font-family:Arial; display:none;">
			<div class="clear"></div>

			<div class="shortcutbox" style="margin-left:10px; margin-top:10px;">
				<h4 style="font-family:simhei">栏目管理</h4>
				<div class="topic">
					<ul>
					    <%Dim RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,BasicType,ModelEname,ChannelStatus From KS_Channel Where ChannelStatus=1 and ChannelID<>6 and channelid<9 Order By OrderID,ChannelID ASC")
						Do While Not RSC.Eof 
						%>
						<li><a href="javascript:void(0)" title="<%=rsc(1)%>" onClick="SelectObjItem1(this,'<%=rsc(1)%> >> <font color=red>栏目管理</font>','Disabled','System/KS.Class.asp?ChannelID=<%=rsc(0)%>',<%=rsc(4)%>);"><i class="icon a<%=rsc(4)%>"></i><span><%=rsc(3)%>栏目</span></a></li>
                        <%
						RSC.MoveNext
						Loop
						RSC.Close
						Set RSC=Nothing
						%>
						<%If KS.C_S(9,21)="1" Then%>
						<li><a href="javascript:void(0)" title="考试系统" onClick="SelectObjItem1(this,'考试系统 >> <font color=red>栏目管理</font>','Disabled','mnkc/mnkc_class.asp',9);"><i class="icon test"></i><span>试卷分类</span></a></li>
                        <%End If%>
						<%If KS.C_S(11,21)="1" Then%>
						<li><a href="javascript:void(0)" title="论坛版面" onClick="SelectObjItem1(this,'论坛系统 >> <font color=red>论坛管理</font>','Disabled','Club/KS.GuestBoard.asp',12);"><i class="icon bbs"></i><span>论坛版面</span></a></li>
					    <%End If%>
						<%If KS.C_S(12,21)="1" Then%>
						<li><a href="javascript:void(0)" title="问答栏目" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>分类管理</font>','Disabled','Ask/KS.AskClass.asp',12);"><i class="icon ask"></i><span>问答分类</span></a></li>
					    <%End If%>
						<li><a href="javascript:void(0)" title="添加栏目" onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'" onClick="SelectObjItem1(this,'栏目管理 >> <font color=red>添加栏目</font>','Go','System/KS.Class.asp?Action=Add',12);"><i class="icon add"></i><span>添加栏目</span></a></li>
						
					</ul>
				</div>
				<div class="clear"></div>
			</div>

			<div class="chenxubox" style="margin-left:10px">
				<div class="leftbox">
					<h4 style="font-family:simhei">程序信息</h4>
					<div>
					当前版本：<%=KS.Version%><br />
					<%if CheckNewVersion then%>
						最新版本：<span id='versioninfo' style="color:#e00404"><script src="http://www.kesion.com/websystem/GetofficialInfo.asp?action=getverbyscript"></script></span>
					<%end if%><br />
					产品开发：厦门科汛软件有限公司<br />
					咨询热线：400-0080-263<br />
					营销QQ：<iframe scrolling="no" frameborder="0" width="120" height="35" allowtransparency="true"  src="http://static.b.qq.com/account/bizqq/wpa/wpa_a04.html?type=4&amp;kfuin=4000080263&amp;ws=http%3A%2F%2Fwww.oppo.com&amp;btn1=%E5%9C%A8%E7%BA%BF%E5%AE%A2%E6%9C%8D&amp;cref=http%3A%2F%2Fwww.oppo.com&amp;pt=-%20OPPO%20%E5%AE%98%E6%96%B9%E7%BD%91%E7%AB%99"></iframe>4000080263<br />
					在线升级：<span id='updateInfo'>正在检测最新版本信息...</span>
					</div>
				</div>
				<div class="rightbox"<%If instr(KS.Setting(16),"2")<=0 Then response.write " style='display:none'"%>>
					<h4 style="font-family:simhei">技术论坛新帖</h4>
					<ul>
					   <script charset="utf-8" id="showtopic" src="http://bbs.kesion.com/Dv_News.asp?GetName=newtopic"></script>
					</ul>	
				</div>
				
				
				
				
			</div>
			
			
 </div>
</body>
</html>
<%
End Sub
Function bytes2BSTR(vIn)
		Dim i,ThisCharCode,NextCharCode
		Dim strReturn:strReturn = ""
		For i = 1 To LenB(vIn)
			ThisCharCode = AscB(MidB(vIn,i,1))
			If ThisCharCode < &H80 Then
				strReturn = strReturn & Chr(ThisCharCode)
			Else
				NextCharCode = AscB(MidB(vIn,i+1,1))
				strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
				i = i + 1
			End If
		Next
		bytes2BSTR = strReturn
		End Function
Function getfile(RemoteFileUrl)
		On Error Resume Next 
		Dim Retrieval:Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
		 .Open "Get", RemoteFileUrl, false, "", ""
		 .Send
		 If .Readystate<>4 then Exit Function
		 getfile =bytes2BSTR(.responseBody)
		End With
		If Err Then
		Err.clear
		getfile="error!"
		End if
		Set Retrieval = Nothing
End function

Function GetTrueDomain(domain)
				Dim x:x = split(domain,".")
				Dim sdomain:sdomain= ""
				Dim start:start = 2
				Dim k :k= 1
				if ubound(x)<=1 then GetTrueDomain=domain:exit function
				if (ubound(x) >= 3) then start = 3
				dim i:i=start
				do while i > 0
					if (i=start) then
						sdomain = sdomain & x(ubound(x)-start+k)
					else
						sdomain = sdomain & "." & x(ubound(x)-start+k)
					end if
					k=k+1
					i=i-1
				loop
				GetTrueDomain=sdomain
End function


Sub CopyRight()
		  If Request.ServerVariables("SERVER_NAME")="127.0.0.1" or Request.ServerVariables("SERVER_NAME")="localhost" Then
		  Else
			  If KS.IsNul(Session(KS.SiteSN&"CheckCopyRight")) Then
			   Session(KS.SiteSN&"CheckCopyRight")=getfile("http://www.ke" & "s"& "ion.com/websystem/VerifyAuthorization.asp?myurl=" & GetTrueDomain(Request.ServerVariables("SERVER_NAME")))
			  End If
			   
			  If Not KS.IsNul(Session(KS.SiteSN&"CheckCopyRight")) and Session(KS.SiteSN&"CheckCopyRight")<>"error!" Then
				If Session(KS.SiteSN&"CheckCopyRight")="true" Then
				  KS.Echo escape("<span style='color:green;font-size:12px'>恭喜，您的网站已经过官方正版授权。有关授权问题<a href='http://vip.kesion.com/' target='_blank'>请点此查询</a></span>。")
				Else
				 If IsBusiness=False Then
				  KS.Echo escape("<span style='color:#333;font-size:12px'>您当前使用的是免费版本，仅授权个人非商业使用。</span>")
				 Else
				  KS.Echo escape("<span style='color:#333;font-size:12px'>您当前使用的版本经官方正版验证不通过，仅授权个人非商业使用。有关授权问题<a href='http://vip.kesion.com/' target='_blank'>请点此查询</a> <span style='font-weight:normal;color:#999;font-size:12px'>tips:商业用户请使用授权域名进入后台，此提示将自动消失。</span></span>")
				 End If
				End If
			  End If
		  End If
End Sub

Public Sub indexXMLField(username)'文件
				dim FieldXML,XMLStr
				set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				FieldXML.async = false
				FieldXML.setProperty "ServerHTTPRequest", true 
				FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/filmk/" & username &"_mk_a.xml"))
				if FieldXML.parseError.errorCode<>0 Then
					 XMLStr=""
					 XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
					 XMLStr=XMLStr&"<field>" &vbcrlf
					 XMLStr=XMLStr&"</field>" &vbcrlf
					 Call KS.WriteTOFile(KS.Setting(3)&"Config/filmk/" & username &"_mk_a.xml",xmlstr)
					 XMLStr=""
					 XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
					 XMLStr=XMLStr&"<field>" &vbcrlf
					 XMLStr=XMLStr&"</field>" &vbcrlf
					 Call KS.WriteTOFile(KS.Setting(3)&"Config/filmk/" & username &"_mk_b.xml",xmlstr)
					 '模拟剪切文件操作
				End If
End Sub

Sub setTips() 
		  Call KS.settingsave(0,KS.G("v"))
		  KS.Die "success"
		End Sub
		Sub ajax1()
						Dim Node,Num,Url,HasVerify
						HasVerify=false
						If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig()
	
						For Each Node In Application(KS.SiteSN&"_ChannelConfig").DocumentElement.SelectNodes("channel[@ks21=1 and @ks0!=6 and @ks6<9]")
								If KS.C("SuperTF")<>"1" and Instr(KS.C("ModelPower"),KS.C_S(Node.SelectSingleNode("@ks0").text,10)&"1")=0 Then 
								 If DataBaseType=1 Then
									 Num=Conn.Execute("Select count(id) from " & Node.SelectSingleNode("@ks2").text & " where deltf=0 and verific=0 and tid in(select id from ks_class where ','+cast(AdminPurview as nvarchar(4000))+',' like '%," & KS.C("GroupID") & "%')")(0)
								 Else
									 Num=Conn.Execute("Select count(id) from " & Node.SelectSingleNode("@ks2").text & " where deltf=0 and verific=0 and  tid in(select id from ks_class where ','+AdminPurview+',' like '%," & KS.C("GroupID") & "%')")(0)
								 End If
							 Else
							   Num=Conn.Execute("Select count(id) from " & Node.SelectSingleNode("@ks2").text & " where deltf=0 and verific=0")(0)            
							 End If
							   If Num=0 Then
							   'KS.Echo "待签" & Node.SelectSingleNode("@ks3").text & ":<font color=red>" & Num &" </font>" & Node.SelectSingleNode("@ks4").text & "&nbsp;"
							   Else
								HasVerify=true
							   KS.Echo "<li><a style='cursor:pointer;' title='点击进入签收' target=""MainFrame"" href='System/KS.ItemInfo.asp?showType=1&ChannelID=" & Node.SelectSingleNode("@ks0").text & "'>待签" & Node.SelectSingleNode("@ks3").text & "[<font color=red>" & Num &"</font>]" & Node.SelectSingleNode("@ks4").text & "</a></li>"
							   End If
						next
						
						 If KS.C_S(10,21)="1" Then
							Num=conn.execute("select count(id) from ks_Job_Company where status=0")(0)
							If Num>0 Then
							 HasVerify=true
							 KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='job/KS.JobCompany.asp?ComeFrom=Verify'>待审招聘单位[<font color=red>" & Num & "</font>]家</a></li>"
							End If
							Num=conn.execute("select count(id) from ks_Job_Resume where status=0")(0)
							If Num>0 Then
							 HasVerify=true
							 KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='job/KS.JobResume.asp?ComeFrom=Verify'>待审简历[<font color=red>" & Num & "</font>]份</a></li>"
							End If
							Num=conn.execute("select count(id) from KS_Job_Edu where status=0")(0)
							If Num>0 Then
							 HasVerify=true
							 KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='job/KS.JobEdu.asp?status=0'>待审教育背[<font color=red>" & Num & "</font>]份</a><li>"
							End If
						 End If
						 
						
					  If KS.Setting(208)="1" and KS.Setting(211)<>"" Then
						 Num=conn.execute("select count(id) from ks_order where status=0")(0)
						 If Num>0 Then
						  HasVerify=true
						  KS.Echo "<li><a style='cursor:pointer;' title='点击进入订单管理' target=""MainFrame"" href='shop/KS.ShopOrder.asp'>待确认订单[<font color=red>" & Num & "</font>]个</a></li>"
						 End If
						End If
						Num=conn.execute("select count(id) from ks_comment where verific=0")(0)
						If Num>0 Then
						 HasVerify=true
						 KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='plus/Plus_DigMood/KS.Comment.asp?ComeFrom=Verify'>待审评论[<font color=red>" & Num & "</font>]条</a></li>"
						End If
						Num=conn.execute("select count(linkid) from ks_link where verific=0")(0)
						If Num>0 Then
						 HasVerify=true
						KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='plus/Plus_Link/KS.FriendLink.asp?Action=Verific'>待审链接[<font color=red>" & Num & "</font>]个</a></li>"
						End If
						Num=conn.execute("select count(blogid) from ks_blog where status=0")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.Space.asp?from=verify'>待审空间[<font color=red>" & Num & "</font>]个</a></li>"
						End If
						Num=conn.execute("select count(id) from ks_bloginfo where status=2")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo "<li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.Spacelog.asp?from=verify'>待审日志[<font color=red>" & Num & "</font>]篇</a></li>"
						End If
						Num=conn.execute("select count(id) from ks_photoxc where status=0")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo " <li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.SpaceAlbum.asp?from=verify'>待审相册[<font color=red>" & Num & "</font>]本</a></li>"
						End If
						Num=conn.execute("select count(id) from ks_team where Verific=0")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo " <li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.SpaceTeam.asp?from=verify'>待审圈子[<font color=red>" & Num & "</font>]个</a></li>"
						End If
						Num=conn.execute("select count(id) from KS_EnterpriseNews where status=0")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo " <li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.EnterPriseNews.asp?from=verify'>待审企业新闻[<font color=red>" & Num & "</font>]篇</a></li>"
						End If
						Num=conn.execute("select count(id) from KS_EnterPriseAD where status=0")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo " <li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.EnterPriseAD.asp?from=verify'>待审行业广告[<font color=red>" & Num & "</font>]个</a></li>"
						End If
						Num=conn.execute("select count(id) from KS_EnterPriseZS where status=0")(0)
						If Num>0 Then
						HasVerify=true
						KS.Echo " <li><a style='cursor:pointer;' title='点击进入审核' target=""MainFrame"" href='space/KS.EnterPriseZS.asp?from=verify'>待审证书[<font color=red>" & Num & "</font>]个</a></li>"
						End If
						
		End sub
		
		Sub CheckSetting()
			 dim strDir,strAdminDir,InstallDir
			 strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
			 strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
			 InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
					
			If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
			   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
			End If
		 If KS.Setting(2)<>KS.GetAutoDoMain or KS.Setting(3)<>InstallDir Then
			
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select Setting From KS_Config",conn,1,3
		  Dim SetArr,SetStr,I
		  SetArr=Split(RS(0),"^%^")
		  For I=0 To Ubound(SetArr)
		   If I=0 Then 
			SetStr=SetArr(0)
		   ElseIf I=2 Then
			SetStr=SetStr & "^%^" & KS.GetAutoDomain
		   ElseIf I=3 Then
			SetStr=SetStr & "^%^" & InstallDir
		   Else
			SetStr=SetStr & "^%^" & SetArr(I)
		   End If
		  Next
		  RS(0)=SetStr
		  RS.Update
		  RS.Close:Set RS=Nothing
		  Call KS.DelCahe(KS.SiteSn & "_Config")
		  Call KS.DelCahe(KS.SiteSn & "_Date")
		 End If
		End Sub
		
		Sub Check3G()
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select WapSetting From KS_Config",conn,1,3
		  Dim SetArr,SetStr,I
		  SetArr=Split(RS(0),"^%^")
		  For I=0 To Ubound(SetArr)
		   If I=0 Then 
			SetStr=0
		   Else
			SetStr=SetStr & "^%^" & SetArr(I)
		   End If
		  Next
		  RS(0)=SetStr
		  RS.Update
		  RS.Close:Set RS=Nothing
		  Call KS.DelCahe(KS.SiteSn & "_Config")
		  Call KS.DelCahe(KS.SiteSn & "_Date")
		End Sub


End Class
%>


<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************


'---ShowVerifyCode_s---
 Const ShowVerifyCode= False    '后台登录是否启用验证码 true 启用 false不启用
'---ShowVerifyCode_e---


Dim KS:Set KS=New PublicCls
Dim Num
'Num=GetBackGroundNum
Function GetBackGroundNum()
	on error resume next
	Dim FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
	Dim FolderObj:Set FolderObj = FsoObj.GetFolder(Server.MapPath("images/login/background"))
	Dim FileObj:Set FileObj = FolderObj.Files
	Dim FsoItem,Num:Num=0
	For Each FsoItem In FileObj
	 if instr(lcase(FsoItem.name),".jpg")<>0 then Num=Num+1
	Next
	Set FSOObj=Nothing
	Set FileObj=Nothing
	if err then
	 err.clear
	 Num=8
	end if
	GetBackGroundNum=Num
End Function
randomize
%>
<!DOCTYPE html>
<html>
<head>
<title><%=KS.Setting(0) & "---网站后台管理"%> X<%=GetVer%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="../ks_inc/jquery.js"></script>
<script src="../ks_inc/common.js"></script>
<!--[if IE 6]>
<script src="../js/iepng.js" ></script>
<script >
   EvPNG.fix('div, ul, img, li, input'); 
</script>
<![endif]-->
<script type="text/javascript">

$(function(){

	cover();
	$(window).resize(function(){
		cover();
	});
	if (!-[1,]){ //IE
	    if (!-[1,]&&!window.XMLHttpRequest){
		  $.dialog.alert('您当前使用的浏览器版本太低，建议升级到更高版本的浏览器！',function(){});
		}
		$("#sub").hover(   
				function() {   
				$("#sub").stop().animate({opacity: '1'},1000);   
		   },    
		 function() {   
			   $("#sub").stop().animate({opacity: '0.5'},1000);   
		 });  
	 } 
	
});
function cover(){
	var h = $(".indexmain").height()/2;
	$(".indexmain").css({marginTop:"-"+h+"px"});
	
	var a = $(".keyboard").offset().left;
	var b = $(".keyboard").offset().top+2;
	$("#keycontainer").css({left:""+a+"px",top:""+b+"px"});
	
};
</script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<style type="text/css">
	html{color:#000;font-family:Arial,sans-serif;font-size:12px;}
	h1, h2, h3, h4, h5, h6, h7, p, ul, ol,div,span, dl, dt, dd, li, body,em,i, form, input,i,cite, button, img, cite, strong,    em,label,fieldset,pre,code,blockquote,    table, td, th ,tr{ padding:0; margin:0;outline:0 none;}
	img, table, td, th ,tr { border:0;}
	address,caption,cite,code,dfn,em,th,var{font-style:normal;font-weight:normal;}
	select,img,select{font-size:12px;vertical-align:middle;color:#666; font-family:Arial,sans-serif}
	.checkbox{vertical-align:middle;margin-right:5px;margin-top:-2px; margin-bottom:1px;}
	textarea{font-size:12px;color:#666; font-family:Arial,sans-serif}
	table{ border-collapse:collapse;border-spacing:0;}
	ul, ol, li { list-style-type:none;}
	a { color:#0082cb; text-decoration:none;}
	a:hover{text-decoration:none;}
	ul:after,.clearfix:after { content: "."; display: block; height: 0; clear: both; visibility: hidden; }/* 不适合用clear时使用 */
	ul,.clearfix{ zoom:1;}
	.clear{clear:both;font-size:0px; height:0px;overflow:hidden;}/*  空白占位  */

	body {font-size:12px;color:#666; height:100%; overflow:hidden;font-family:'Arial','hiragino sans gb','microsoft yahei ui','microsoft yahei',simsun,sans-serif;}
	input{font-family:'Arial','hiragino sans gb','microsoft yahei ui','microsoft yahei',simsun,sans-serif;}


	.indexmain{position:absolute;left:50%;top:50%;width:745px;margin-left:-372px; background:#fff; z-index:999; box-shadow:0px 1px 15px rgba(0,0,0,0.15);}
	.indexmain .left{width:340px; float:left; padding:40px 30px 40px 40px;}
	.indexmain .left .title{}
	.indexmain .left .title .logo img{ display:block; overflow:hidden;height:25px;}
	.indexmain .left .title span{ font-size:22px;font-weight:bold;color:#626a70; line-height:42px;margin-top:10px; display:block;}
	.indexmain .left .intro{ line-height:24px; font-size:14px;color:#a1b1bd;margin-top:20px;}
	.indexmain .left .copyRight{width:340px; position:absolute;bottom:25px;color:#a1b1bd; line-height:22px;}
	

	.indexmain .right{width:275px; background:#f4f7f7; overflow:hidden; float:right; overflow:hidden; padding:40px 30px 45px 30px;}
	.indexmain .right .title h4{font-size:16px; line-height:26px; overflow:hidden;height:26px;color:#626a70;}
	.indexmain .right .title span{ display:block; font-size:14px; line-height:24px;height:24px; overflow:hidden;color:#a1b1bd;}
	
	/*
	.tabboxboth ul{margin-top:10px;}

	.tabboxboth .label{font-size:14px;color:#666; font-family:simhei; line-height:44px;} 
	.tabboxboth .label .land{padding-left:28px; float:left;}
	.tabboxboth .label .code{padding-left:28px; float:left;}
	.tabboxboth .input,.tabboxboth .textinput{width:275px; height:24px; padding:10px 0px 10px 11px;line-height:20px; font-family:Arial, Helvetica, sans-serif;color:#666;border:0px; background:url(Images/login/bg05.png) no-repeat;font-size:12px;}
	.tabboxboth .input,.tabboxboth .textinput:hover{ background:url(Images/login/bg06.png) no-repeat}
	.tabboxboth .input,.tabboxboth .textinput2{width:106px; height:24px; padding:10px 0px 10px 11px;line-height:20px; font-family:Arial, Helvetica, sans-serif;color:#666;border:0px; background:url(Images/login/bg09.png) no-repeat;font-size:13px;}
	.tabboxboth .input,.tabboxboth .textinput2:hover{}
	.tabboxsingle{ height:305px;}
	.tabboxsingle ul{ margin-top:50px;}
	.tabboxsingle ul li{ padding:4px 0px 5px; position:relative;}
	.tabboxsingle ul li.btn{ padding-left:98px;}
	
	.tabboxsingle .input,.tabboxsingle .textinput{width:235px; height:24px; padding:10px 0px 10px 11px;line-height:20px; font-family:Arial, Helvetica, sans-serif;color:#666;border:0px; background:url(Images/login/bg05.png) no-repeat;font-size:12px;}
	.tabboxsingle .input,.tabboxsingle .textinput:hover{ background:url(Images/login/bg06.png) no-repeat}
	.tabboxsingle .input,.tabboxsingle .textinput2{width:106px; height:24px; padding:10px 0px 10px 11px;line-height:20px; font-family:Arial, Helvetica, sans-serif;color:#666;border:0px; background:url(Images/login/bg09.png) no-repeat;font-size:13px;}
	*/
	
	.tabbox{margin-top:20px;}

	.tabbox .textinput{width:245px; height:28px; padding:5px 15px; font-size:14px; line-height:28px;border:1px solid #e9eeef; background:#fff;margin-top:15px; -webkit-transition:0.3s; transition:0.3s;color:#a1b1bd;}
	.tabbox .textinput:focus{color:#626a70;border:1px solid #539ed5; background:#f9fdff; box-shadow:0 0 5px rgba(36,126,192,0.4) inset;}
	
	.tabbox #UserName{background:#fff url(Images/login/icon-user.png) no-repeat 15px 50%; padding-left:40px;width:220px;}
	.tabbox #UserName:focus{background:#f9fdff url(Images/login/icon-useron.png) no-repeat 15px 50%;}
	
	.tabbox .password{background:#fff url(Images/login/lock.png) no-repeat 15px 50%; padding-left:40px;width:220px;}
	.tabbox .password:focus{background:#f9fdff url(Images/login/lock-on.png) no-repeat 15px 50%;}

	.tabbox .keyboard{background:#fff url(Images/login/keyboard.png) no-repeat 15px 50%; padding-left:40px;width:220px;}
	.tabbox .keyboard:focus{background:#f9fdff url(Images/login/keyboard-on.png) no-repeat 15px 50%;}
	

	.tabbox .verification{background:#fff url(Images/login/verification.png) no-repeat 15px 50%; padding-left:40px;width:90px;}
	.tabbox .verification:focus{background:#f9fdff url(Images/login/verification-on.png) no-repeat 15px 50%;}
	

	
	.regsubmit{width:100%;height:40px;border:0px; background:#2d88cb; cursor:pointer;font-size:16px;color:#fff;margin-top:30px; -webkit-transition:0.3s; transition:0.3s;}
	.regsubmit:hover{background:#247ec0;}
	.rzm{font-size:12px;line-height:22px;color:#e82d2d;margin-top:15px;}
	
	#softkeyboard{ width:inherit !important; background:#fff; box-shadow:0 1px 5px rgba(0,0,0,0.2);}
	
	.bodyBg{ position:relative;}
	.bodyBg li{ position:absolute;width:100%;left:0;top:0;height:100%;}
	.bodyBg img{width:100%;}
	
	
	
</style>
</head>
<body style="overflow:hidden" scroll="no">
<div id="wrap">
<%
Select Case  KS.G("Action")
 Case "LoginCheck"
  Call CheckLogin()
 Case "LoginOut"
  Call LoginOut()
 Case Else
  Call CheckSetting()
  Call Main()
End Select

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

Sub Main()
%>
 <FORM ACTION="Login.asp?Action=LoginCheck" method="post" name="LoginForm" onSubmit="return(CheckForm(this))" class="form">

 	<div id="step_1" class="indexmain">
		<div class="left">
			<div class="title">
            	<div class="logo"><img src="Images/login/logo.png" alt="KesionCMS" /></div>
				<span>后台管理系统 X<%=GetVer%></span>
			</div>
			<div class="intro">KesionCMS® 是厦门科汛软件有限公司开发的一套万能建站产品，是CMS网站系统行业最流行的网站建设解决方案之一。</div>
            
            <div class="copyRight">
            厦门科汛软件有限公司<br />Copyright &copy;2006-<%=year(now)%> <a href="http://www.kesion.com" target="_blank"> www.kesion.com</a>,All Rights Reserved. 
     	   </div>    
            
            
		</div>
		<div class="right">
		   	<div class="title"><h4>管理员登录</h4><span>Administrator Login</span></div>
           
					<div class="tabbox">
						<ul id="regSpan" class="companyul">
							<li style="z-index:1000">
								<div class="label">
									<input type="text" value="请输入登录帐号" onBlur="this.value=(this.value=='')?'请输入登录帐号':this.value" onFocus="this.value=(this.value=='请输入登录帐号')?'':this.value" name="UserName" id="UserName" class="textinput" tabindex="1" autocomplete="off" />
								</div>
							</li>
							<li>
								<div class="label" id="passarea" style="position:relative;">
									<%IF KS.Setting(98)<>"1" Then%><input value="请输入登录密码" onBlur="this.value=(this.value=='')?'请输入登录密码':this.value" onFocus="this.value=(this.value=='请输入登录密码')?'':this.value" type="password" tabindex="2" name="PWD" id="PWD" class="textinput password" /><%Else%><input name="PWD" type="password" onFocus="$(this).val('');" value="请输入登录密码"  id="PWD" onClick="showKeyBoard(this);" maxlength="50" class="textinput keyboard" tabindex="2" readonly />
									<%End If%>

								</div>
								
							</li>

						  <%If ShowVerifyCode Then%>
							<li>
								<div class="label" style="position:relative;">
									<input type="text" id="Verifycode" name="Verifycode" tabindex="3" class="textinput verification" maxlength="4" value="验证码" onBlur="this.value=(this.value=='')?'验证码':this.value" onFocus="this.value=(this.value=='验证码')?'':this.value" /><img id="imagecode" src="../plus/verifycode.asp?time=0.001" onClick="$(this).attr('src',$(this).attr('src')+Math.random());" title="点击刷新验证码" style="cursor:pointer; background:#fff;height:20px; padding:10px; position:absolute;left:157px;top:15px;"/>
								</div>
							</li>
						 <%End If%>	
                          <%if EnableSiteManageCode = True Then%>
							<li>
								<div class="label">
									<input value="认证密码" onBlur="this.value=(this.value=='')?'认证密码':this.value" onFocus="this.value=(this.value=='认证密码')?'':this.value" type="password" id="AdminLoginCode" name="AdminLoginCode" tabindex="4" class="textinput password" />
								</div>
							</li>
							<%if SiteManageCode="8888" Then%>
							<li class="rzm">
								提示：原始认证密码为<span>8888</span>，为了安全请打开conn.asp修改！
							</li>
							<%end if%>
						<%end if%>
							
							<li class="btn" id="nextStep">
							  <input type="submit" tabindex="5" class="regsubmit" value="管理员登录">
							</li>
						</ul>
					</div>
		</div>
	</div>

	
</FORM>


<script>
    function showKeyBoard(obj) {

        $("#keycontainer").show();
        $("#keyboard").empty().html($("#keybtn").html());
        var $write = $('#'+obj.id),
		shift = false,
		capslock = false;
        $('#keyboard li').click(function () {
            var $this = $(this);
            var character = $this.html(); // If it's a lowercase letter, nothing happens to this variable
            // Shift keys
            if ($this.hasClass('left-shift') || $this.hasClass('right-shift')) {
                $('.letter').toggleClass('uppercase');
                $('.symbol span').toggle();

                shift = (shift === true) ? false : true;
                capslock = false;
                return false;
            }

            // Caps lock
            if ($this.hasClass('capslock')) {
                $('.letter').toggleClass('uppercase');
                capslock = true;
                return false;
            }

            // Delete
            if ($this.hasClass('delete')) {
                var html = $write.val();
                $write.val(html.substr(0, html.length - 1));
                return false;
            }

            // Special characters
            if ($this.hasClass('symbol')) character = $('span:visible', $this).html();
            if ($this.hasClass('space')) character = ' ';
            if ($this.hasClass('tab')) character = "\t";
            //if ($this.hasClass('return')) character = "\n";
            if ($this.hasClass('return')) {
                $('#keycontainer').hide();
                return;
            }



            // Uppercase letter
            if ($this.hasClass('uppercase')) character = character.toUpperCase();

            // Remove shift once a key is clicked.
            if (shift === true) {
                $('.symbol span').toggle();
                if (capslock === false) $('.letter').toggleClass('uppercase');

                shift = false;
            }
            if ($write.val()=='******') $write.val('');
            // Add the character
            $write.val($write.val() + character);
        });


    }
</script>
<style>
#keycontainer {
position:absolute; z-index:180;
display:none;
margin: 38px auto;
width: 516px;
border:1px solid #ccc;
height:194px;
padding-left:3px;
padding-top:3px;
background-color:#f1f1f1;
}
#keytitle{text-align:center;height:22px;line-height:22px;font-weight:bold;}
#keyboard {
margin: 0;
padding: 0;
list-style: none;
}
#keyboard li {
	float: left;
	margin: 0 2px 2px 0;
	width: 30px;
	height: 30px;
	line-height: 30px;
	text-align: center;
	background: #fff;
	border: 1px solid #f9f9f9;
	-moz-border-radius: 5px;
	-webkit-border-radius: 5px;
	}
.capslock, .tab, .left-shift {
		clear: left;
		}
			#keyboard .tab, #keyboard .delete {
			width: 70px;
			}
			#keyboard .capslock {
			width: 80px;
			}
			#keyboard .return {
			width: 53px;
			}
			#keyboard .left-shift {
			width: 95px;
			}
			#keyboard .right-shift {
			width: 72px;
			}
		.lastitem {
		margin-right: 0;
		}
		.uppercase {
		text-transform: uppercase;
		}
		#keyboard .space {
		clear: left;
		width: 510px;
		}
		.on {
		display: none;
		}
		#keyboard li:hover {
		position: relative;
		top: 1px;
		left: 1px;
		border-color: #e5e5e5;
		cursor: pointer;
		}
</style>

<div id="keycontainer">
   <div id="keytitle"><span style="float:right;padding-right:10px;cursor:pointer" onClick="$('#keycontainer').hide();">[close]</span>===KESION 软键盘===</div>
	<ul id="keyboard">
    </ul>
    <ul id="keybtn" style="display:none">
		<li class="symbol"><span class="off">`</span><span class="on">~</span></li>
		<li class="symbol"><span class="off">1</span><span class="on">!</span></li>
		<li class="symbol"><span class="off">2</span><span class="on">@</span></li>
		<li class="symbol"><span class="off">3</span><span class="on">#</span></li>
		<li class="symbol"><span class="off">4</span><span class="on">$</span></li>
		<li class="symbol"><span class="off">5</span><span class="on">%</span></li>
		<li class="symbol"><span class="off">6</span><span class="on">^</span></li>
		<li class="symbol"><span class="off">7</span><span class="on">&amp;</span></li>
		<li class="symbol"><span class="off">8</span><span class="on">*</span></li>
		<li class="symbol"><span class="off">9</span><span class="on">(</span></li>
		<li class="symbol"><span class="off">0</span><span class="on">)</span></li>
		<li class="symbol"><span class="off">-</span><span class="on">_</span></li>
		<li class="symbol"><span class="off">=</span><span class="on">+</span></li>
		<li class="delete lastitem">delete</li>
		<li class="tab">tab</li>
		<li class="letter">q</li>
		<li class="letter">w</li>
		<li class="letter">e</li>
		<li class="letter">r</li>
		<li class="letter">t</li>
		<li class="letter">y</li>
		<li class="letter">u</li>
		<li class="letter">i</li>
		<li class="letter">o</li>
		<li class="letter">p</li>
		<li class="symbol"><span class="off">[</span><span class="on">{</span></li>
		<li class="symbol"><span class="off">]</span><span class="on">}</span></li>
		<li class="symbol lastitem"><span class="off">\</span><span class="on">|</span></li>
		<li class="capslock">caps lock</li>
		<li class="letter">a</li>
		<li class="letter">s</li>
		<li class="letter">d</li>
		<li class="letter">f</li>
		<li class="letter">g</li>
		<li class="letter">h</li>
		<li class="letter">j</li>
		<li class="letter">k</li>
		<li class="letter">l</li>
		<li class="symbol"><span class="off">;</span><span class="on">:</span></li>
		<li class="symbol"><span class="off">'</span><span class="on">&quot;</span></li>
		<li class="return lastitem">return</li>
		<li class="left-shift">shift</li>
		<li class="letter">z</li>
		<li class="letter">x</li>
		<li class="letter">c</li>
		<li class="letter">v</li>
		<li class="letter">b</li>
		<li class="letter">n</li>
		<li class="letter">m</li>
		<li class="symbol"><span class="off">,</span><span class="on">&lt;</span></li>
		<li class="symbol"><span class="off">.</span><span class="on">&gt;</span></li>
		<li class="symbol"><span class="off">/</span><span class="on">?</span></li>
		<li class="right-shift lastitem">shift</li>
		<li class="space lastitem">space</li>
	</ul>
</div>








<script>

<!--
/*$(document).ready(function() { 
	$(".label").hover(function(){$(this).removeClass("label");$(this).addClass("labelhover");
	},function(){
	$(this).removeClass("labelhover");$(this).addClass("label");});
});*/

setTimeout(function(){$("#UserName").focus();},500); 

function CheckForm(ObjForm) {
  if(ObjForm.UserName.value == '请输入登录帐号') {
    $.dialog.alert('请输入管理账号！',function(){ObjForm.UserName.focus();});
    return false;
  }
  if(ObjForm.PWD.value == '请输入登录密码') {
    $.dialog.alert('请输入授权密码！',function(){ObjForm.PWD.focus();});
    return false;
  }
  if (ObjForm.PWD.value.length<6)
  {
   $.dialog.alert('授权密码不能少于六位！',function(){ObjForm.PWD.focus();});
    return false;
  }
  <%if EnableSiteManageCode = True Then%>
  if (ObjForm.AdminLoginCode.value == '认证密码') {
    $.dialog.alert('请输入后台管理认证密码！',function(){ObjForm.AdminLoginCode.focus();});
    return false;
  }
  <%End If%>
  <%If ShowVerifyCode Then%>
  if (ObjForm.Verifycode.value == '验证码') {
	$.dialog.alert('请输入验证码！',function(){ObjForm.Verifycode.focus();});
	
    return false;
  }
  <%End If%>
}


	function bgauto(){
		var i = $("#bgNum").text();
		var n = $("#bodyBg").find("li").length;
		i++;
		if(i>n-1){
			i=0;
		};
		$("#bgNum").html(""+i+"");
		$("#bodyBg").find("li:eq("+i+")").fadeIn(1000).siblings().fadeOut(1000);
		
	};
	setInterval(bgauto,3000);
	

//-->
</script>
    	
    </div>
    <div id="bgNum" style="display:none;">0</div>
	<div class="bodyBg" id="bodyBg">
    	<ul>
        	<li><img src="Images/53a93f237ed4e.jpg" /></li>
        	<li><img src="Images/23023783935.jpg" /></li>
        	<li style="display:none;"><img src="Images/55f6701c8505f.jpg" /></li>
        </ul>
    </div>
</body>
</html>
<%End Sub
Sub CheckLogin()
  Dim PWD,UserName,LoginRS,SqlStr,RndPassword
  Dim ScriptName,AdminLoginCode
  AdminLoginCode=KS.G("AdminLoginCode")
  IF lcase(Trim(Request.Form("Verifycode")))<>lcase(Trim(Session("Verifycode"))) And ShowVerifyCode then 
   Call KS.Echo("<script>$.dialog.alert('<br/>登录失败:验证码有误，请重新输入！',function(){history.back();});</script>")
   exit Sub
  end if
  If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
   Call KS.Echo("<script>$.dialog.alert('<br/>登录失败:您输入的后台管理认证码不对，请重新输入！',function(){history.back();});</script>")
   exit Sub
  End If
  Pwd =MD5(KS.R(KS.S("pwd")),16)

  UserName = KS.R(trim(KS.S("username")))
  RndPassword=KS.R(KS.MakeRandomChar(20))
  ScriptName=KS.R(Trim(Request.ServerVariables("HTTP_REFERER")))
  Set LoginRS = Server.CreateObject("ADODB.RecordSet")
  SqlStr = "select top 1 a.*,b.PowerList,b.ModelPower,B.[Type],B.Role,B.ManageOtherDoc from KS_Admin a inner join KS_UserGroup b on a.GroupID=b.ID where a.UserName='" & UserName & "'"
  LoginRS.Open SqlStr,Conn,1,3
  If LoginRS.EOF AND LoginRS.BOF Then
	  Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的帐号!")
      Call KS.Die("<script>$.dialog.alert('<br/>登录失败:您输入了错误的帐号，请再次输入！',function(){history.back();});</script>")
  Else
  
     IF LoginRS("PassWord")=pwd THEN
       IF Cint(LoginRS("Locked"))=1 Then
         Call KS.Die("<script>$.dialog.alert('<br/>登录失败:您的账号已被管理员锁定，请与您的系统管理员联系！',function(){history.back();});</script>")
	   Else
		  	 '登录成功，进行前台验证，并更新数据
			   on error resume next 
			  Dim UserRS:Set UserRS=Server.CreateObject("Adodb.Recordset")
			  UserRS.Open "Select top 1 * From KS_User Where UserName='" & LoginRS("PrUserName") & "' and GroupID=1",Conn,1,3
			  IF Not UserRS.Eof Then
			  
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
					 UserRS("LastLoginIP") = KS.GetIP
					 UserRS("LastLoginTime") = Now()
					 UserRS("LoginTimes") = UserRS("LoginTimes") + 1
					 UserRS("RndPassWord") = RndPassWord
					 UserRS("IsOnline")=1
					 UserRS.Update	
			 if err then
			   ks.die "<script>$.dialog.alert(""登录失败！<br/><strong>失败原因：</strong> " & err.description &""",function(){history.back();});</script>"
			   err.clear
			 end if	
	
					'置前台会员登录状态
                    If EnabledSubDomain Then
							Response.Cookies(KS.SiteSn).domain=RootDomain					
					Else
                            Response.Cookies(KS.SiteSn).path = "/"
					End If		
					 Response.Cookies(KS.SiteSn)("UserID") = UserRS("UserID")
					 Response.Cookies(KS.SiteSn)("UserName") = KS.R(UserRS("UserName"))
			         Response.Cookies(KS.SiteSn)("Password") = UserRS("Password")
					 Response.Cookies(KS.SiteSn)("RndPassword") = KS.R(UserRS("RndPassword"))
					 Response.Cookies(KS.SiteSn)("AdminLoginCode") = AdminLoginCode
					 Response.Cookies(KS.SiteSn)("AdminID") =  LoginRS("AdminID")
					 Response.Cookies(KS.SiteSn)("AdminName") =  UserName
					 Response.Cookies(KS.SiteSn)("AdminPass") = pwd
					 If LoginRS("Type")=3 Then
					 Response.Cookies(KS.SiteSn)("SuperTF")   = 1
					 Else
					 Response.Cookies(KS.SiteSn)("SuperTF")   = 0
					 End If
					 If LoginRS("SuperTF")=1 Or  LoginRS("Type")=3 Then   '记录管理员角色
					 Response.Cookies(KS.SiteSn)("Role") = 3
					 Else
					 Response.Cookies(KS.SiteSn)("Role") = LoginRS("Role")
					 End IF
					 Response.Cookies(KS.SiteSn)("ManageOtherDoc") = KS.ChkClng(LoginRS("ManageOtherDoc"))
					 Response.Cookies(KS.SiteSn)("GroupID") = LoginRS("GroupID")
					 Response.Cookies(KS.SiteSn)("PowerList") = LoginRS("PowerList")
					 Response.Cookies(KS.SiteSn)("ModelPower") = LoginRS("ModelPower")
					 'Response.Cookies(KS.SiteSn).Expires = DateAdd("h", 3, Now())   '3小时没有操作自动失败
             Else 
				   Call KS.InsertLog(UserName,0,ScriptName,"找不到前台账号!")
                   Call KS.Die("<script>$.dialog.alert('<br/>登录失败:找不到前台账号！',function(){history.back();});</script>")
			 End If
			   UserRS.Close:Set UserRS=Nothing
			   
	  LoginRS("LastLoginTime")=Now
	  LoginRS("LastLoginIP")=KS.GetIP
	  LoginRS("LoginTimes")=LoginRS("LoginTimes")+1
	  LoginRS.UpDate
	  Call KS.InsertLog(UserName,1,ScriptName,"成功登录后台系统!")
      Call KS.Die("<script>;setTimeout(""top.location.href='index.asp'"",10);</script>")
	End IF
  ELse
     If EnabledSubDomain Then
		Response.Cookies(KS.SiteSn).domain=RootDomain					
	 Else
        Response.Cookies(KS.SiteSn).path = "/"
	End If
	Response.Cookies(KS.SiteSn)("AdminID") =""
    Response.Cookies(KS.SiteSn)("AdminName")=""
	Response.Cookies(KS.SiteSn)("AdminPass")=""
	Response.Cookies(KS.SiteSn)("SuperTF")=""
	Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
	Response.Cookies(KS.SiteSn)("PowerList")=""
	Response.Cookies(KS.SiteSn)("ModelPower")=""
	Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的口令:" & Request.form("pwd"))
    Call KS.Die("<script>$.dialog.alert('<br/>登录失败:您输入了错误的口令，请再次输入！',function(){history.back();});</script>")
  END IF
 End If
END Sub
Sub LoginOut()
		   Conn.Execute("Update KS_Admin Set LastLogoutTime=" & SqlNowString & " where UserName='" & KS.R(KS.C("AdminName")) &"'")
		   Dim AdminDir:AdminDir=KS.Setting(89)
		   If EnabledSubDomain Then
				Response.Cookies(KS.SiteSn).domain=RootDomain					
			Else
                Response.Cookies(KS.SiteSn).path = "/"
			End If
			Response.Cookies(KS.SiteSn)("Role")=""
			Response.Cookies(KS.SiteSn)("PowerList")=""
			Response.Cookies(KS.SiteSn)("AdminID") =""
			Response.Cookies(KS.SiteSn)("AdminName")=""
			Response.Cookies(KS.SiteSn)("AdminPass")=""
			Response.Cookies(KS.SiteSn)("SuperTF")=""
			Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
			Response.Cookies(KS.SiteSn)("ModelPower")=""
			session.Abandon()
			Response.Write ("<script> top.location.href='" & KS.Setting(2) & KS.Setting(3) &"';</script>")
End Sub
Set KS=Nothing
%>

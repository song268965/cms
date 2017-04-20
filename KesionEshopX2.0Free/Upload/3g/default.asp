<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/Kesion.CommonCls.asp"-->
<%Dim KS:Set KS=New PublicCls%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="keywords" content="aspcms,cms,开源cms系统,内容管理系统">
<meta name="description" content="KesionCMS是厦门科汛软件有限公司开发的一套万能建站产品，是CMS行业最流行的网站建设解决方案之一,我们一直专注于开发cms建站系统。">
<title>手机模拟器浏览3G版 ---<%=KS.Setting(0)%>-Powered By KesionCMS</title>
<style type="text/css">
header, hgroup, menu, nav, section, menu,footer,article,figure,figcaption,commend,aside{display:block;margin:0;padding:0;}
body,p,input,h1,h2,h3,h4,ul,li,dl,dt,dd,form,textarea{
	margin:0;
	padding:0;
	list-style:none;
	vertical-align:middle;
}
body{ 
	font-family:"\5FAE\8F6F\96C5\9ED1", Helvetica;
	font-size:16px; 
	
}
img {border:0; }

.main{
	width:445px;
	margin:auto;
	
}

.main .iframe{
	height:624px;
	padding-top:201px;
	padding-left:53px;
	background:url(images/mobilebg.png) no-repeat 50% 0;
}


.foot{
	font-size:13px;
	width:98%;
	margin-top:10px;
	text-align:center;
}

</style>
<SCRIPT LANGUAGE="JavaScript"> 
<!--
if(navigator.platform != "Win32" ){
	location.href="index.asp";
}
//-->
</SCRIPT>
</head>
<body>
<div class="main">
    <div class="iframe">
    	<iframe width="338" id="mainiframe" height="454" scrolling="auto" marginwidth=0 marginheight=0 frameborder="0" src="index.asp"></iframe>
 
    </div>
</div>
<div class="foot">提示：用您的手机输入以下网址即可浏览到以上模拟器看到的效果<br><b><%=KS.GetDomain()%><%=KS.WSetting(4)%>/</b></div>
</body>
</html>
<%Set KS=Nothing
CloseConn%>
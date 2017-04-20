<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/3GCls.asp"-->
<!--#include file="../API/cls_api.asp"-->
<!--#include file="../api/uc_client/client.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New MemberCls
KSCls.Kesion()
Set KSCls = Nothing

Class MemberCls
        Private KS,F_C,KSR,KSUser,totalscore,action
		Private TotalPut,CurrPage,MaxPerPage
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="payonline.asp"-->
		<%
		Public Sub Kesion()
			IF Cbool(KSUser.UserLoginChecked)=false Then
			  response.Redirect("login.asp")
			  Exit Sub
			End If
			action=request("action")
			CurrPage=KS.ChkClng(Request("Page"))
			If CurrPage<=0 Then CurrPage=1
			MaxPerPage=5
		%><!DOCTYPE html>
<html>
<head> 
<title>会员中心-<%=KS.Setting(0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Cache-control" content="max-age=1700">
<meta name="viewport" content="user-scalable=no, width=device-width">
<meta name="MobileOptimized" content="320">
<meta name="author" content="kesion.com">
<meta name="format-detection" content="telephone=no">
<link href="images/common.css" rel="stylesheet" type="text/css">
<link href="images/style.css" rel="stylesheet" type="text/css">
<link href="images/member.css" rel="stylesheet" type="text/css">
<script src="../ks_inc/jquery.js"></script>
<script src="../ks_inc/common.js"></script>
</head>

<body>
<div class="navpositionbig" style=" height:40px; display:none;">
	<div class="navposition">
		<div class="logonav" style=" border-bottom:2px solid #0C9AD8">
		    <%If KS.IsNul(Action) Then%>
		 	 <div class="fl"><img src="images/logo.png" /></div>
			 <div class="fr">您好,<font color="#a52a2a"><%=KSUser.GetUserInfo("UserName")%></font> <a href="Logout.asp" onClick="return(confirm('确定退出吗？'));">退出</a></div>
			<%Else%>
		 	 <div style="float:left;width:10%;"><a href="javascript:history.back()"><img src="images/left-arrow-vector.png" width="28" height="28"/></a></div>
			 <div style="margin:0 auto;width:80%;float:left;text-align:center;font-weight:bold;font-size:15pt;color:#0C9AD8">
			 <%select case action
	            case "fav" response.write "我的收藏夹"
	            case "edit" response.write "用户信息"
				case "payonline" response.write "在线支付"
	            case "order" response.write "我的订单"
	            case "showorder" response.write "订单详情"
	            case "mnkc" response.write "模拟考场"
	            case "comment" response.write "我的评论"
	            case "complaints" response.write "投诉/意见"
				case "logmoney" response.write "消费记录"
		      end select
			 %>
			 </div>
			 <div style="width:10%;text-right:right;float:right;"><a href="user.asp"><img src="images/user-male-alt-vector.png" width="28" height="28"/></a></div>
			<%End If%>
		</div>
	</div>
</div>	
<header class="headerbox">
	<div class="header">
		<div class="return headin inleft"><a href="javascript:;" onClick="history.back()"><img src="/3g/images/left.png"></a></div>
        <div class="headertit">会员中心</div>
		<div class="bill headin inright"><img src="/3g/images/bill.png"></div>
    </div>
    <div class="slidebar">
    	<ul>
        	<li class="user">
                <div class="name"><script src="/user/userlogin.asp?action=3g"></script></div>
            </li>
            <li><a href="/3g">首页</a></li>        
			<li><a href="/3g/list.asp?id=664">新闻频道</a></li>
			<li><a href="/3g/list.asp?id=694">图片频道</a></li>
			<li><a href="/3g/list.asp?id=719">下载频道</a></li>
			<li><a href="/3g/list.asp?id=926">网上购物</a></li>
        </ul>
    </div>
    <div class="fixbg"></div>
</header>
<section style="height:2.5rem;"></section>
<script>
$(function(){
	$(".bill").click(function(){
		$(".slidebar").addClass("show")
		$(".fixbg").show();
	});
	$(".fixbg").click(function(){
		$(".slidebar").removeClass("show")
		$(".fixbg").hide();
	})
})
</script>	
<div class="MiddleCont">
	<div class="userbox">
	   <%select case action
	      case "fav" fav  '我的收藏夹
		  case "logmoney" logmoney  '消费记录
		  case "complaints" complaints  '投诉建议
		  case "comment" comment  '我的评论
	      case "edit" edit
		  case "editsave" editsave
		  case "payonline" payonline
		  case "payshoporder" payshoporder
		  case "paystep2" paystep2
		  case "paystep3" paystep3
		  case "order" order
		  case "showorder" showorder
		  case "delorder" delorder
		  case "signup" SignUp
		  case "addpayment" addpayment
		  case "savepayment" savepayment
		  case "setorderok" setorderok   '结清订单
		  case "mnkc" mnkc
		  case else main()
		 end select
	  %>
	</div>
</div>
<!--<div class="footbig">
	<div class="foot">
		<ul>
			<li><a href="index.asp" class="icon1">主页</a></li>
			<li><a href="bbs.asp" class="icon2">论坛</a></li>
			<li><a href="user.asp" class="icon3">我的</a></li>
			<li ><a href="Logout.asp" onClick="return(confirm('确定退出吗？'));" class="icon9">退出</a></li>
		</ul>
	</div>
</div>-->
<section style="height:2.5rem;"></section>
<footer class="footbox">
	<ul class="flexbox">
		<li><a href="/3g"><img src="/3g/images/icon1.png">首页</a></li>    
		<li><a href="/3g/list.asp?id=664"><img src="/3g/images/icon2.png">资讯</a></li>    
		<li><a href="/3g/bbs.asp"><img src="/3g/images/icon3.png">论坛</a></li>    
		<li><a href="/3g/user.asp" ><img src="/3g/images/icon4.png">会员</a></li>    
    </ul>
</footer>
<script>
$(function(){
	var a = $(".flexbox li").eq(3).find("img").attr("src");
	b=a.substring(0,a.length-4); //alert(b)截掉最后四个字符
	$(".flexbox li").eq(3).find("img").attr("src", b + "s.png");
})
</script>
</body>
</html>
<%
End Sub

Sub Main()
		%>
		 <div class="usertopbox"><div class="usertopleft">
	    <% Dim UserFaceSrc:UserFaceSrc=KSUser.GetUserInfo("UserFace")
		if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
		%>
		<div class="avatar48"><img src="<%=UserFaceSrc%>" onerror="this.onerror=null;this.src='../user/images/noavatar_middle.gif'" ></div>
		
		<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../user/User_UpFile.asp?Type=Pic&ChannelID=9999&MaxFileSize=1500&ext=*.jpg;*.gif;*.png' frameborder="0" scrolling="No" align="center" height='50'></iframe>
		<script>
		$("#UpPhotoFrame").load(function(){
			$(this).contents().find("head").append('<link href="/3g/images/member.css" rel="stylesheet" />')		
		});

		</script>
	   </div>
	   <div class="usertopright">
		 <%If Not KS.IsNul(KSUser.GetUserInfo("realname")) Then
			     response.write KSUser.GetUserInfo("realname") &"(" & KSUser.UserName &")"
			    Else
				 response.write KSUser.UserName
				End If		 
			  %>
		
	      <span class="uid"><span class="hyid">ID.<%=KSUser.GetUserInfo("UserID")%></span><span class="hyz"><%=KS.U_G(KSUser.GroupID,"groupname")%></span></span>
	  </div>
	  <div class="clear"></div>
	  </div>
	  <div class="mymoney" >
		<ul>
			<li><span>可用资金</span><span style="color:#333; font-weight:bold;"><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1)%>元</span></li>
			<li><span>可用<%=KS.Setting(45)%></span><span style="color:#333;font-weight:bold;"><%=formatnumber(KSUser.GetUserInfo("Point"),0,-1) & "" & KS.Setting(46)%></span></li>
			<li><span>总积分</span><span style="color:#333;font-weight:bold;"><%=KSUser.GetUserInfo("score")%>分</span></li>
			<li><span>可用积分</span><span style="color:#333;font-weight:bold;"><%=KSUser.GetScore%>分</span></li>
		</ul>										
		<div class="clear"></div>
      </div>
	  <div class="userborder">
	  
	
	 
	 <%if ks.Setting(201)="1" then
     %>
	  <div class="clear"></div>
	  <div class="cat">
		<style>
		/*签到*/
.qd_daym{background:#FFF;height:2.65rem; margin-bottom:0.5rem;font-size:0.8rem; }
.qd_day{margin-top:0.5rem;color:#2D7ECB; font-size:0.7rem; line-height:0.9rem;  width:33.33%;float:left;*border-right:#f5f5f5 1px solid;box-sizing: border-box;border: none;
    background-image: -webkit-linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-image: -moz-linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-image: -o-linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-image: linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-size: 1px 100%;
    background-repeat: no-repeat;
    background-position: right;}
.qd_day li{ text-align:center;}

.qd_order{text-align:center; margin-top:0.5rem;  width:33.33%; float:left;height:1.75rem; line-height:1.75rem; }
.qd{ font-size:0.8rem; font-weight:normal; color:red; border-bottom:#eee 1px solid; margin-bottom:0.5rem; height:1.75rem; line-height:1.75rem;}
.qd_order a{color:#2D7ECB; font-size:0.7rem;}

.qiandao{ text-align:center; color:#2D7ECB; font-size:0.7rem;box-sizing: border-box;margin-top:0.5rem;height:1.75rem; line-height:1.75rem;*border-right:#f5f5f5 1px solid;border: none;
    background-image: -webkit-linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-image: -moz-linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-image: -o-linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-image: linear-gradient(right ,transparent 50%,#e5e5e5 50%);
    background-size: 1px 100%;
    background-repeat: no-repeat;
    background-position: right;}
/*.qiandao_hvr{  cursor:pointer; background: #0099FF;color:#FFFFFF;}*/
/*.qiandao_hvr a{ color: #FFFFFF}*/
.qiandao_form{ position: fixed;top: 50%;margin-top: -2.3rem;left: 50%;margin-left: -7.5rem;display: none;background: #FFFFFF;box-shadow: 0px 0px 6px #bbb;border-radius:0.25rem;width: 14rem;padding: 0.5rem; z-index:9999}
.qiandao_form .qiandao_formx ul li{ float:left; width:2.75rem;}

		</style>
		<script>
		 $(function(){
			 $(".qd_daym .qd_day").next("div").css("width","33.33%");
			 })
		</script>
		 <%
		 Call KSUser.QianDao()
		 %>
	</div>
		<%end if%>
	 
	  
	  
	  
	  
			 <div class="clear"></div>
			 <div class="userdetail">
			 <a href="?action=payonline"><i class="iconfont">&#xe758;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>账务信息</h3></a>
			 <a href="?action=edit"><i class="iconfont">&#xe6b8;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>用户信息</h3></a>
			 <a href="?action=order"><i class="iconfont">&#xe6a2;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>我的订单</h3></a>
			 </div>
			 <div class="blank10"></div>
			 <div class="clear"></div>
			 <div class="userdetail">
			 <a href="?action=fav"><i class="iconfont">&#xe6a0;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>我的收藏</h3></a>
			 <a href="?action=logmoney"><i class="iconfont">&#xe723;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>消费记录</h3></a>
			 <a href="?action=complaints"><i class="iconfont">&#xe6c7;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>投诉/意见</h3></a>
			 <a href="?action=comment"><i class="iconfont">&#xe69b;</i><h3 class="f1"><span class="iconfont">&#xe6a7;</span>我的评论</h3></a>
			 </div>
			 <div class="blank10"></div>
	  </div>
		<%
	End Sub
	
	'修改用户资料
	Sub Edit()
	%>
	<script type="text/javascript">
      function CheckForm(){ 
			if (document.myform.RealName.value ==""){
				 $.dialog.alert("请填写您的真实姓名！",function(){
				   document.myform.RealName.focus();
				 });
				 return false;
			}
			if (document.myform.Sex.value ==""){
				$.dialog.alert("请选择您的性别！"，function(){
				  document.myform.Sex.focus();
				});
				return false;
			}
			if (document.myform.Address.value ==""){
				$.dialog.alert("请输入您的联系地址！"，function(){
				document.myform.Address.focus();
				});
				return false;
			}
			  return true;	
			}
    </script>
	<div class="tableBGw">
	                <iframe src="about:blank" name="hidiframe" width="0" height="0" frameborder="0"  style="display:none;"></iframe>
		           <form target="hidiframe" action="user.asp?action=editsave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				      <div class="fgTitle">基本资料</div>
					  <div class="tableOther">
	                  <table cellspacing="0" cellpadding="0"  width="100%" align="center" border="0" class="border" style="margin-top:0;">
                          <tr class="tdbg">
                            <td class="clefttitle" nowrap="nowrap" style=" width:10%">昵称</td>
                            <td class="aRight"><input  class="textbox1" type="hidden" name="username" size="30" value="<%=KSUser.username%>" disabled="disabled" /> <%=KSUser.username%></td>
                          </tr>
                          
                          <tr class="tdbg">
                            <td class="clefttitle">姓名</td>
                            <td class="aRight">
							<%if KSUser.GetUserInfo("issfzrz")="1" then%>
							<input name="RealName" class="textbox1" type="hidden" id="RealName" value="<%=KSUser.GetUserInfo("RealName")%>"  /><%=KSUser.GetUserInfo("RealName")%> 
							身份证号：<%=KSUser.GetUserInfo("idcard")%>
							<span class="msgtips">*已经过实名认证</span>
							<%else%>
							<input name="RealName" class="textbox1" placeholder="真实姓名" type="text" id="RealName" value="<%=KSUser.GetUserInfo("RealName")%>" size="20" maxlength="50" />
                              <span style="color: red">* </span>
							<%end if%></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle" nowrap="nowrap">性别</td>
                            <td class="aRight"> <label><input type="radio" name="Sex" value="男" <%if KSUser.GetUserInfo("Sex")="男" then response.write " checked"%> />先生</label>
							
							<label><input type="radio" name="Sex" value="女" <%if KSUser.GetUserInfo("Sex")="女" then response.write " checked"%> />女士</label>
                            </td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">生日</td>
                            <td class="aRight"><%dim birthday:birthday=KSUser.GetUserInfo("Birthday")
							    if not isdate(birthday) then birthday=now
								dim i
								%>
								<select name="b1" class="select"> 
								  <option value='0'>年</option>
								 <%for i=year(now) to 1950 step -1
								   if year(birthday)=i then
								       response.write "<option selected value='" & i &"'>" & i &"年</option>"
								   else
									   response.write "<option value='" & i &"'>" & i &"年</option>"
								   end if
								   next
								 %>
								</select>
								<select name="b2" class="select">
								  <option value='0'>月</option>
								 <%for i=1 to 12 
								    if month(birthday)=i then
								     response.write "<option selected value='" & i &"'>" & i &"月</option>"
								    else
									  response.write "<option value='" & i &"'>" & i &"月</option>"
									end if
								   next
								 %>
								</select>
								<select name="b3" class="select">
								  <option value='0'>日</option>
								 <%for i=1 to 31 
								   if day(Birthday)=i then
								    response.write "<option selected value='" & i &"'>" & i &"日</option>"
								   else
								   response.write "<option value='" & i &"'>" & i &"日</option>"
								   end if
								   next
								 %>
								</select>
								</td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">邮箱</td>
                            <td class="aRight">
							<%if KSUser.GetUserInfo("isemailrz")="1"  and KSUser.GetUserInfo("email")<>"" then%>
							<%=KSUser.GetUserInfo("email")%><span class="msgtips">*已经过Email认证</span>
							<%else%>
							<input name="Email" class="textbox1" placeholder="邮箱地址" type="text" id="Email" value="<%=KSUser.GetUserInfo("Email")%>" size="20" maxlength="50" />
                                <span style="color: red">*</span>
							<%end if%></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">手机</td>
                            <td class="aRight">
							<%if KSUser.GetUserInfo("ismobilerz")="1"  and KSUser.GetUserInfo("Mobile")<>"" then%>
							 <%=KSUser.GetUserInfo("Mobile")%><span class="msgtips">*已经过手机认证</span>
							<%else%>
							<input name="Mobile" placeholder="手机号码" class="textbox1" type="text" id="Mobile" value="<%=KSUser.GetUserInfo("Mobile")%>" size="20" maxlength="50" />
                                <span style="color: red">*</span>
							<%end if%>	
								</td>
                          </tr>
						  
						 </table>
						 </div>
						 <div class="tableOther">
						 <table cellspacing="0" cellpadding="0"  width="100%" align="center" border="0" class="border">
							 <tr class="tdbg"><td class="clefttitle">个人签名</td></tr>
							  <tr class="aRight">
								<td style="padding:0.75rem 0;"><textarea name="Sign" placeholder="个人签名" class="textbox1" cols="30" rows="5" id="Sign" style="width:100%; height:3rem; border:none; background:none; resize:none;margin-top: 0;line-height: 1rem;"><%= KSUser.GetUserInfo("Sign")%></textarea></td>
							  </tr>
						 </table>
						 </div>
						 <div class="fgTitle">详细资料</div>
						 <table cellspacing="0" cellpadding="0"  width="100%" align="center" border="0" class="border" style="margin-top:0; border-bottom:1px solid #eee;">
						  
						  <tr>
						    <td colspan="2" style="padding: 0 0.75rem;">
							 <dl class="dtable">
							<% 
							Dim RSU:Set RSU=Server.CreateObject("ADODB.RECORDSET")
							RSU.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,1
							If RSU.Eof Then
							  RSU.Close:Set RSU=Nothing
							  Response.Write "<script>alert('非法参数！');history.back();</script>"
							  Response.End()
							End If
							
						  Dim Template:Template=LFCls.GetSingleFieldValue("Select top 1 wapTemplate From KS_UserForm Where ID=" & KS.ChkClng(KS.U_G(KSUser.GroupID,"formid")))

						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select top 1 FormField From KS_UserForm Where ID=" & KS.ChkClng(KS.U_G(KSUser.GroupID,"formid")))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType,ShowUnit,UnitOptions,MaxLength,Title from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr,FieldStr,Height,Width
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   If IsArray(SQL) Then
						   For K=0 TO Ubound(SQL,2)
						     Width=KS.ChkClng(SQL(4,K)) : If Width<200 Then Width=200
						     Height=KS.ChkClng(SQL(5,K)) : If Height=0 Then Height=50
						     FieldStr=FieldStr & "|" & lcase(SQL(2,K))
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If lcase(replace(SQL(2,K),"&",""))="provincecity" Then
								 InputStr="<script>try{setCookie(""pid"",'" & rsu("province") & "');setCookie(""cid"",'" & rsu("City") & "');}catch(e){}</script>" & vbcrlf
								 InputStr=InputStr & "<script src='../plus/area.asp?width=70'></script><script language=""javascript"">" &vbcrlf
								 If RSU("Province")<>"" And Not ISNull(RSU("Province")) Then
						         InputStr=InputStr & "$('#Province').val('" & RSU("province") &"');" &vbcrlf
								 End If
						         If RSU("City")<>"" And Not ISNull(RSU("City")) Then
								  InputStr=InputStr & "$('#City').val('" & RSU("City") & "');" &Vbcrlf
						         end if
								 If rsU("County")<>"" And Not ISNull(rsU("County")) Then
								  InputStr=InputStr & "$('#County').val('" & rsU("County") & "');" &Vbcrlf
						         end if
						          InputStr=InputStr & "</script>" &vbcrlf
							  Else
							  Select Case SQL(1,K)
								Case 2,10:InputStr="<textarea rows=""5"" placeholder=""" & SQL(11,K) & """  name=""" & SQL(2,K) & """ class=""textarea"">" &RSU(SQL(2,K)) & "</textarea>"
								Case 3,11
								  If SQL(1,K)=11 Then
					               InputStr= "<select style=""width:" & SQL(4,K) & "px"" name=""" & SQL(2,K) & """ onchange=""fill" & SQL(2,K) &"(this.value)""><option value=''>---请选择---</option>"
								  Else
								   InputStr="<select style=""width:" & SQL(4,K) & "px"" name=""" & SQL(2,K) & """>"
								  End If
								  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
								  For N=0 To O_Len
									 F_V=Split(O_Arr(N),"|")
									 If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									 Else
										O_Value=F_V(0):O_Text=F_V(0)
									 End If						   
									 If Trim(RSU(SQL(2,K)))=O_Value Then
										InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
										InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
								  Next
									InputStr=InputStr & "</select>"
									'联动菜单
									If SQL(1,K)=11  Then
										Dim JSStr
										InputStr=InputStr &  GetLDMenuStr(RSU,101,SQL,SQL(2,k),JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
									End If
								Case 6
									 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
									 For N=0 To O_Len
										F_V=Split(O_Arr(N),"|")
										If Ubound(F_V)=1 Then
										 O_Value=F_V(0):O_Text=F_V(1)
										Else
										 O_Value=F_V(0):O_Text=F_V(0)
										End If
										If Trim(RSU(SQL(2,K)))=O_Value Then
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
										Else
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
										 End If
									 Next
							  Case 7
									O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 For N=0 To O_Len
										  F_V=Split(O_Arr(N),"|")
										  If Ubound(F_V)=1 Then
											O_Value=F_V(0):O_Text=F_V(1)
										  Else
											O_Value=F_V(0):O_Text=F_V(0)
										  End If						   
										  If KS.FoundInArr(Trim(RSU(SQL(2,K))),O_Value,",")=true Then
												 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
										 Else
										  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
										 End If
								   Next
							  
									
							  Case Else
							      Dim MaxLength:MaxLength=KS.ChkClng(SQL(10,K))
				                  If MaxLength=0 Then MaxLength=255
								  InputStr="<input type=""text"" MaxLength="""& MaxLength & """ class=""textbox""  placeholder=""" & SQL(11,K) & """ name=""" & lcase(SQL(2,K)) & """ id=""" & SQL(2,K) & """ value=""" & RSU(SQL(2,K)) & """>"
							  End Select
							  End If
							
							  If SQL(8,K)="1" Then 
								  InputStr=InputStr & " <select name=""" & SQL(2,K) & "_Unit"" id=""" & SQL(2,K) & "_Unit"">"
								  If Not KS.IsNul(SQL(9,k)) Then
								   Dim KK,UnitOptionsArr:UnitOptionsArr=Split(SQL(9,k),vbcrlf)
								   For KK=0 To Ubound(UnitOptionsArr)
								      If Trim(RSU(SQL(2,K) & "_Unit"))=Trim(UnitOptionsArr(KK)) Then
									  InputStr=InputStr & "<option value='" & UnitOptionsArr(KK) & "' selected>" & UnitOptionsArr(KK) & "</option>"                 
									  Else
									  InputStr=InputStr & "<option value='" & UnitOptionsArr(KK) & "'>" & UnitOptionsArr(KK) & "</option>"                 
									  End If
								   Next
								  End If
								  InputStr=InputStr & "</select>"
			                  End If

							  
							'  if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='30'></iframe></div>"
							  
							  
				              If Instr(Template,"{@NoDisplay(" & SQL(2,K) & ")}")<>0 Then
							   Template=Replace(Template,"{@NoDisplay(" & SQL(2,K) & ")}"," noshow=""true""")
							  End If
							  
							  Template=Replace(Template,"[@" & replace(SQL(2,K),"&","") & "]","<div class=""left"">" &SQL(11,K) &"：</div>" & InputStr)
							 End If
						   Next
						End If
							Response.Write Template
					%> </dl>
					<script>
					 $(window).load(function(){
						$("dd[noshow]").remove();
					  });
					
						$(function(){
							$(".dtable dd").each(function() {
								$(this).find(".left").html($(this).find(".left").html().replace(/：/,""));
							});
						})
					</script>
					
							</td>
						  </tr>
						  
                          
            </table>
			<div class="blank10"></div>
			<div class="updataBtn"><button type="submit"  class="pn">OK,修 改</button></div>
		    </form>
		</div>
	<%
	End Sub
	
	Sub editsave()
	   Dim RealName:RealName=KS.S("RealName")
	   Dim Sex:Sex=KS.S("Sex")
	   Dim Birthday:Birthday=KS.S("B1")&"-" & KS.S("B2") &"-" & KS.S("B3")
	   Dim Sign:Sign=KS.S("Sign")	
	   Dim Mobile:Mobile=KS.S("Mobile")	
	   Dim Address:Address=KS.S("Address")
	   Dim QQ:QQ=KS.S("QQ")
	   Dim ZipCode:ZipCode=KS.S("ZipCode")
		If Not IsDate(Birthday) Then
			KS.Die "<script>alert('出生日期格式有误!');</script>"
		 end if
				
						   '过滤
            dim kk,sarr
            sarr=split(KS.WordFilter,",")
            for kk=0 to ubound(sarr)
               if instr(Sign,sarr(kk))<>0 then 
                  ks.die  "<script>alert('签名含有非常关键词:" & sarr(kk) &",请不要非法提交恶意信息!');</script>" 
               end if
            next

				
				if KSUser.GetUserInfo("isemailrz")<>"1" Then
					  Dim Email:Email=KS.S("Email")
					 if KS.IsValidEmail(Email)=false then
						 Response.Write("<script>$.dialog.tips('请输入正确的电子邮箱!',1,'error.gif',function(){parent.document.getElementById('Email').focus();});</script>")
						 Exit Sub
					 end if
					 Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
					If EmailMultiRegTF=0 Then
						Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select top 1 UserID from KS_User where UserName<>'" & KSUser.UserName & "' And Email='" & Email & "'")
						If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
							EmailRSCheck.Close:Set EmailRSCheck = Nothing
							Response.Write("<script>alert('您注册的Email已经存在！请更换Email再试试！');</script>")
							Exit Sub
						End If
						EmailRSCheck.Close:Set EmailRSCheck = Nothing
					 End If
               end if
				 
		

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 RS.Close:Set RS=Nothing:Response.End
			  Else
				 RS("RealName")=RealName
				 RS("Sex")=Sex
				 RS("Birthday")=Birthday
				 if rs("isemailrz")<>1 then RS("Email")=Email
				 if rs("ismobilerz")<>1 then RS("Mobile")=Mobile
				 RS("Sign")=Sign
				 RS("Address")=Address
				 RS("QQ")=qq
				 RS("Zip")=ZipCode
				 If Not KS.IsNul(RS("userface")) Then
				   If Instr(lcase(RS("userface")),"boy.jpg")<>0 Or Instr(lcase(RS("userface")),"girl.jpg")<>0 Then
				    If Sex="男" Then 
					  rs("userface")=KS.GetDomain & "Images/Face/boy.jpg"
					Else
					  rs("userface")=KS.GetDomain & "Images/face/girl.jpg"
					End If
				   End If
				 End If
		 		 RS.Update
				 RS.Close:Set RS=Nothing
				  Call ContactInfoSave()
				 Session(KS.SiteSN&"UserInfo")=""
				 Response.Write "<script>alert('会员基本信息资料修改成功！');parent.location.href='user.asp';</script>"
				 Response.End()
			  End if
	End Sub
	
  '保存联系信息
  Sub ContactInfoSave()
         Dim SQL,K,SQLStr
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(KSUser.GroupID,"formid"))
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		
		 If KS.FilterIDs(FieldsList)="" Then
		 SQLStr="Select FieldName,MustFillTF,Title,FieldType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=101 and ShowOnUserForm=1 and (ParentFieldName<>'0' and ParentFieldName is not null)"
		 Else
		 SQLStr="Select FieldName,MustFillTF,Title,FieldType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=101 and ShowOnUserForm=1 and (FieldID In(" & KS.FilterIDs(FieldsList) & ") or (ParentFieldName<>'0' and ParentFieldName is not null))"
		 End If
		 RS.Open SQLStr,Conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		  For K=0 To UBound(SQL,2)
			  If SQL(6,K)="0" Then
				   If SQL(1,K)="1" Then 
					 if lcase(SQL(0,K))<>"province&city" and KS.S(SQL(0,K))="" then
						Response.Write "<script>alert('" & SQL(2,K) & "必须填写!');</script>"
						Response.End()
					 elseif KS.S("province")="" or ks.s("city")="" then
						Response.Write "<script>alert('地区必须选择!');</script>"
						Response.End()
					 end if
				   End If
	
				   
				   
				   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then 
					 Response.Write "<script>alert('" & SQL(2,K) & "必须填写数字!');</script>"
					 Response.End()
				   End If
				   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
					 Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的日期!');</script>"
					 Response.End()
				   End If
				   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
					Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的Email格式!');</script>"
					Response.End()
				   End If
			  End If 
			 Next

  
		 Dim RealName:RealName=KS.LoseHtml(KS.S("RealName"))
		 Dim Sex:Sex=KS.LoseHtml(KS.S("Sex"))
		 Dim Birthday:Birthday=KS.S("Birthday")
		 Dim IDCard:IDCard=KS.LoseHtml(KS.S("IDCard"))
		 Dim OfficeTel:OfficeTel=KS.LoseHtml(KS.S("OfficeTel"))
		 Dim HomeTel:HomeTel=KS.LoseHtml(KS.S("HomeTel"))
		 Dim Mobile:Mobile=KS.LoseHtml(KS.S("Mobile"))
		 Dim Fax:Fax=KS.LoseHtml(KS.S("Fax"))
		 Dim province:province=KS.LoseHtml(KS.S("province"))
		 Dim city:city=KS.LoseHtml(KS.S("city"))
		 Dim county:county=KS.LoseHtml(KS.S("county"))
		 Dim Address:Address=KS.LoseHtml(KS.S("Address"))
		 Dim ZIP:ZIP=KS.LoseHtml(KS.S("ZIP"))
		 Dim HomePage:HomePage=KS.LoseHtml(KS.S("HomePage"))
		 Dim QQ:QQ=KS.LoseHtml(KS.S("QQ"))
		 Dim ICQ:ICQ=KS.LoseHtml(KS.S("ICQ"))
		 Dim MSN:MSN=KS.LoseHtml(KS.S("MSN"))
		 Dim UC:UC=KS.LoseHtml(KS.S("UC"))
		 Dim Sign:Sign=KS.CheckXSS(KS.S("Sign"))
		 Dim Privacy:Privacy=KS.ChkClng(KS.S("Privacy"))
		 
		   '过滤
            dim kk,sarr
            sarr=split(KS.WordFilter,",")
            for kk=0 to ubound(sarr)
               if instr(Sign,sarr(kk))<>0 then 
                  ks.die  "<script>alert('签名含有非常关键词:" & sarr(kk) &",请不要非法提交恶意信息!');</script>"
               end if
            next
			
			'-----------------------------------------------------------------
			'系统整合
			'-----------------------------------------------------------------
			If API_Enable Then
				call uc_user_edit(KSUser.UserName ,"" ,"",KSUser.GetUserInfo("Email"),1,"","")
			End If
			 
              Dim RS,UpFiles
			  Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 Response.End
			  Else
			     
				 If BirthDay<>"" Then RS("Birthday")=Birthday
				 If Sign<>"" Then RS("Sign")=Sign
				 
				 If Sex<>"" Then 
				   RS("Sex")=Sex
					   If Not KS.IsNul(RS("userface")) Then
					   If Instr(lcase(RS("userface")),"boy.jpg")<>0 Or Instr(lcase(RS("userface")),"girl.jpg")<>0 Then
						If Sex="男" Then 
						  rs("userface")=KS.GetDomain & "Images/Face/boy.jpg"
						Else
						  rs("userface")=KS.GetDomain & "Images/face/girl.jpg"
						End If
					   End If
					 End If
				 End If
				 If RealName<>"" Then RS("RealName")=RealName
				 If IDCard<>"" Then	 RS("IDCard")=IDCard
				 
				 RS("Email")=KSUser.GetUserInfo("Email")
				 RS("OfficeTel")=OfficeTel
				 RS("HomeTel")=HomeTel
				 RS("Mobile")=Mobile
				 RS("Fax")=Fax
				 RS("Province")=Province
				 RS("City")=City
				 RS("county")=county
				 RS("Address")=Address
				 RS("Zip")=Zip
				 RS("HomePage")=HomePage
				 RS("QQ")=QQ
				 RS("ICQ")=ICQ
				 RS("MSN")=MSN
				 RS("UC")=UC
				 RS("Privacy")=Privacy
				 '自定义字段
				 For K=0 To UBound(SQL,2)
				  If left(Lcase(SQL(0,K)),3)="ks_" Then
				   RS(SQL(0,K))=KS.LoseHtml(KS.S(SQL(0,K)))
				   	If SQL(3,K)="9" or SQL(3,K)="10" Then
					   UpFiles=UpFiles & KS.S(SQL(0,K))
					End If
				  End If
				  If SQL(4,K)="1" Then
				   RS(SQL(0,K)&"_Unit")=KS.LoseHtml(KS.S(SQL(0,K)&"_Unit"))
				  End If
				 Next
		 		 RS.Update
				 
				 Call KS.FileAssociation(1023,RS("UserID"),UpFiles,1)
				 
				 Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
				 If IsObject(FieldsXml) Then
				   	 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					If objNode.Attributes.item(0).Text<>"0" Then
					   If Not Conn.Execute("Select top 1 UserName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'").Eof Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								Conn.Execute("UPDATE KS_EnterPrise Set " & objAtr.Attributes.item(0).Text & "='" & RS(objAtr.Attributes.item(1).Text) & "' Where UserName='" & KSUser.UserName & "'")
						 Next
					   End If
					End If
				 End If

				 
				 If KS.C_S(8,21)="1" Then
				  Conn.Execute("Update KS_GQ Set ContactMan='" & RealName &"',Tel='" &OfficeTel & "',Address='" & Address & "',Province='" & Province & "',City='" & City & "',Zip='" & Zip & "',Fax='" & Fax & "',Homepage='" & HomePage & "' where inputer='" & KSUser.UserName & "'")
				 End If

			  End if
			RS.Close:Set RS=Nothing
  End Sub
  
  '我的收藏夹
   Sub fav()
     Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
	 Dim SQLStr
	 If KS.S("From")="cancel" Then
	   Dim IDs:IDS=KS.FilterIds(KS.S("id"))
	   If IDS="" Then KS.Die "<script>$.dialog.alert('请选择要删除的记录！',function(){ history.back(); });</script>"
	   Conn.Execute("Delete From KS_Favorite "& Param &" and id in ("  & KS.FilterIds(KS.S("id")) &")")
	   Response.Redirect Request.ServerVariables("HTTP_REFERER")
	 End If
	 %>
	 <div style="background:#fff;">
	<FORM Action="User.asp?Action=fav&from=cancel" name="myform" method="post">
	<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="myCollection">
					<%
						 Dim RS:Set RS=Server.CreateObject("AdodB.Recordset")
						 SqlStr="Select ID,ChannelID,InfoID,AddDate From KS_Favorite "& Param &" and  Channelid<>6 order by id desc"
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td class=""empty"">您的收藏夹没有内容!</td></tr>"
						 Else
									totalPut = RS.RecordCount
									If currpage < 1 Then	currpage = 1
			
								If currpage >1 and  (currpage - 1) * MaxPerPage < totalPut Then
										RS.Move (currpage - 1) * MaxPerPage
								End If
								Dim I,SQL,K
			SQL=RS.GetRows(MaxPerPage)
			For K=0 To Ubound(SQL,2)
		%>
			<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="myCollection">

		<%
						Select Case KS.C_S(SQL(1,K),6)
						   Case 1 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 2 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 3 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 4 SqlStr="Select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
						   Case 5 SqlStr="Select top 1 ID,Title,Tid,0,0,Fname,0,AddDate,hits From KS_Product Where ID=" & SQL(2,K)
						   Case 7 SqlStr="Select top 1 ID,Title,Tid,0,0,Fname,0,AddDate,hits From KS_Movie Where ID=" & SQL(2,K)
						   Case 8 SqlStr="Select top 1 ID,Title,Tid,0,0,Fname,0,AddDate,hits From KS_GQ Where ID=" & SQL(2,K)
                           Case 9 SqlStr="Select top 1 ID,Title,0,0,0,0,0,date,hits From KS_SJ Where ID=" & SQL(2,K)
						   Case else SqlStr="Select top 1 ID From KS_Article Where 1=0"
						  End Select
						  
						  Dim Url,RSF:Set RSF=Conn.Execute(SqlStr)
						  If Not RSF.Eof Then
						   If SQL(1,K)=9 then
						    url="exam/index.asp?id=" & RSF(0) & ""
						   else
						    url=KS.Get3GItemUrl(SQL(1,K),RSF(2),RSF(0),RSF(0)&KS.WSetting(9))
						   end if
	%>
		
		        <tr class="title">
					<td class="ContentTitle">
					<input id="ID" type="checkbox" value="<%=SQL(0,K)%>"  name="ID">&nbsp;<%=KS.C_S(SQL(1,K),3) %>：<%="<a href=""" & url & """ target=""_blank"">" & RSF(1) & "</a>"%></td>
				</tr>
		        <tr>
					<td class="Contenttips">
					<%
					 Response.Write "<span>收藏时间：" & KS.GetTimeFormat(SQL(3,K)) & "<br/>最后更新：" & KS.GetTimeFormat(RSF(7)) & "<br/>人气：" & RSF(8)
					%>
					</td>
				</tr>
		        <tr>
				 <td  class="splittd" style="text-align:right;padding: 0.5rem 0.75rem;"> <a class="box" href="user.asp?Action=fav&from=cancel&Page=<%=currpage%>&ID=<%=SQL(0,K)%>" onclick = "return (confirm('确定取消该<%=KS.C_S(SQL(1,K),3)%>的收藏吗?'))">取消</a>
				 </td>
				</tr>
		</table>
		   <%End If
	  Next
			
%>
	<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" class="myCollection2">

		<tr>
		  <td height="30" style="text-align:left; font-size:0.7rem; line-height:1.8rem; padding:0.2rem 0.75rem;">
				 <INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll" style="vertical-align:middle; margin-right:0.1rem;">&nbsp;全选
				 <INPUT  class="button" style=" float:right;font-size:0.7rem;" onClick="return(confirm('确定取消选定的收藏吗?'));" type="submit" value="取消收藏" name=submit1>
		  </td>
		</tr>
			<%
				End If
   %>
     </table>
	</FORM>
		  <%Call KS.ShowPage(totalput, MaxPerPage, "", currpage,false,true)%></div>
	<%
   End Sub
   
  
   
   
   
   '投诉建议
   Sub Complaints()
     
	  If KS.S("flag")="dosave" Then 
	     if ks.s("title")="" then
		 response.write "<script>alert('请输入主题!');history.back();</script>"
		 exit sub
		end if
	    if ks.s("content")="" then
		 response.write "<script>alert('请输入内容!');history.back();</script>"
		 exit sub
		end if
		
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "select * from ks_feedback where username='" & KSUser.UserName & "' and id=" & ID,conn,1,3
		 If RS.EOf Then
		  rs.addnew
		  rs("adddate")=now
		 end if
		 rs("username")=ksuser.username
		 rs("title")=KS.CheckXSS(ks.s("title"))
		 rs("object")=KS.CheckXSS(ks.s("object"))
		 rs("content")=KS.CheckXSS(ks.s("content"))
		 rs("hopesolution")=KS.CheckXSS(ks.s("hopesolution"))
		 rs.update
		 rs.close
		 set rs=nothing
		 KS.Die "<script>alert('你的投诉已提交，请耐心等待处理结果!');location.href='user.asp?action=complaints&flag=record';</script>"
	  ElseIf KS.S("flag")="del" Then
	     Conn.Execute("Delete From KS_FeedBack Where  (Accepted='' or Accepted is null ) and UserName='" & KSUser.UserName &"' and id=" & KS.ChkClng(KS.S("id")))
		 Response.Redirect Request.Servervariables("HTTP_REFERER")
      End If
   %>
          <div class="tabs">	
			<ul>
				<li<%if request("flag")="" then KS.Echo " class='puton'"%>><a href="?action=complaints">我要投诉</a></li>
				<li<%if request("flag")="record"  then KS.Echo " class='puton'"%>><a href="?action=complaints&flag=record">投诉记录</a></li>
			</ul>
         </div>
		 <%if request("flag")="" then%>
             <script>
				  function checkform()
				  {
				   if ($('#Title').val()==''){
				    $.dialog.alert('请输入投诉主题!',function(){
					$('#Title').focus();
					});
					return false;
				   }
				   if ($('#content').val()==''){
				    $.dialog.alert('请输入投诉内容!',function(){
					$('#content').focus();
					});
					return false;
				   }
				  }
				 </script>
				 <form name="bmform" action="?action=complaints&flag=dosave" method="post">
				<div class="opinion">
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" style=" margin:10px 0px;">
				  <tr>
                      <td align="right" class='splittd' height="25" nowrap><strong>意见主题：</strong></td>
                      <td  align="left" style=" padding-left:0px;" > 
					  <input type="text" name="Title" id="Title" class="textboxq" size="30">
				      </td>
				  </tr>
				   <tr>
                      <td align="right" class='splittd' height="35" nowrap><strong>意见对象：</strong></td>
                      <td  align="left"  height="25" style=" padding-left:0px;"> <input type="text" name="Object" class="textboxq" size="30"> </td>
                  </tr>
				   <tr>
                      <td align="right" class='splittd' height="35" nowrap><strong>意见内容：</strong></td>
                      <td  align="left"  height="25" style=" padding-left:0px;"> 
					  <textarea name="content" class="textboxq" id="content" style="height:100px; width:84%"></textarea>
				     </td>
                  </tr>
				  <tr>
                      <td  align="right" class='splittd' height="35" nowrap><strong>期望解决方案：</strong></td>
                      <td  align="left"  height="25" style=" padding-left:0px;"> 
					  <textarea name="Hopesolution" class="textboxq" style="height:100px; width:84%"></textarea>
				    </td>
                  </tr>
                  <tr><td align="right" class='splittd' height="35" nowrap></td><td  align="left"  height="25" style=" padding-left:0px;"><input type="Submit" class="button" onClick="return(checkform())" value=" 立即投诉 "></td></tr> 
           </table>
		</div>
	    </form>
		
		<%Elseif KS.S("flag")="show" Then
		
	   Set RS=Server.CreateOBject("ADODB.RECORDSET")
	   RS.Open "Select top 1 * from ks_feedback where username='" & KSUser.UserName & "' and  id=" & KS.ChkClng(KS.S("ID")),conn,1,1

	   IF RS.Eof Then
	     RS.CLOSE:Set RS=Nothing
		 Response.Write "<script>alert('出错了!');history.back();</script>"
	   End If
	%>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
         <td align="center">
				<div class="fgTitle" style="background: #efefef;font-size: 0.7rem;">查看投诉详情</div>
				<div class="tableBGw">
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="Cpayment normaltext">
                    <tr>
                      <td align="left" class='clefttitle' nowrap>意见主题</td>
                      <td class='aRight'> 
					  &nbsp;<%=RS("title")%>
				      </td>
				  </tr>
				   <tr>
                      <td align="left" class='clefttitle' nowrap>意见对象</td>
                      <td class='aRight'>&nbsp; <%=RS("object")%> </td>
                  </tr>
				  <tr><td colspan="2" class='fgTitle' style="padding: 0 !important;background: #efefef;font-size: 0.7rem;">意见内容</td></tr>
				   <tr>
                      <td class='aRight' colspan="2" style="text-align:left;padding-left: 0.75rem;">&nbsp; <%=RS("content")%> </td>
                  </tr>
				  <tr><td colspan="2" class='fgTitle' style="padding: 0 !important;background: #efefef;font-size: 0.7rem;">希望处理结果</td></tr>
                    <tr>
                      <td class='aRight' colspan="2" style="text-align:left;padding-left: 0.75rem;">&nbsp;<%=RS("hopesolution")%></td>
                      
                    </tr>
                    <tr>
                      <td align="left" class='clefttitle' nowrap>处理人</td>
                      <td class='aRight'>&nbsp;<%=RS("accepted")%></td>
                      
                    </tr>
                    <tr>
                      <td align="left" class='clefttitle' nowrap>处理时间</td>
                      <td class='aRight'>&nbsp;<%=RS("accepttime")%></td>
                      
                    </tr>
                    <tr>
                      <td align="left" class='clefttitle' nowrap>处理结果</td>
                      <td class='aRight'>&nbsp;<%=RS("acceptresult")%></td>
                      
                    </tr>
                    <tr><td colspan="2" align="center"><input type="button" class="button" value=" 返 回 " onClick="history.back();"></td></tr>
                   
           </table>
		   </div>
                
		 
		 </td>
       </tr>

     </table>
	 <%RS.Close:Set RS=Nothing
		
	  Else%>
		    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ComRecord">
         <%
		      Dim Param:Param=" where UserName='" & KSUser.UserName & "'"
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select * From KS_FeedBack " & Param & " order By ID",conn,1,1
				If RS.EOF And RS.BOF Then
					Response.Write "<tr><td class='empty'><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe6c7;</i></div>您没有发表任意见或投诉!</div></td></tr>"
				Else
						totalPut = RS.RecordCount
						If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
						End If
						Dim str,i
						  Do While Not RS.Eof
							  dim bh:bh=rs("id")
							  IF LEN(BH)=1 THEN 
								  BH="00"& bh
							  ElseIf LEN(BH)=2 Then
								  Bh="0" & bh
							  End If
							  bh="YJ" & year(rs("adddate")) & month(rs("adddate")) & bh
							  response.write "<tr><td><table class=""PointsDetails"" width=""100%""><tr class=""title"">"
							  Response.Write "<td height='30' class='ContentTitle'>编号：" & bh & ""
							  Response.Write "&nbsp;&nbsp;主题：" 
							  
							  Response.write rs("title")
							  response.write "</td>"
							  Response.Write "</tr><tr><td class='Contenttips'>投诉对象：" & rs("object")&"<br/>投诉时间：" & formatdatetime(rs("adddate"),2) & "<br/>处 理 人："
							  Dim AcceptTime,Delstr,strs
							  if rs("Accepted")="" or isnull(rs("accepted")) then
							   response.write "未处理"
							   AcceptTime="---"
							   Delstr="<a onclick=""return(confirm('确定删除吗?'))"" href='?action=complaints&flag=del&id=" & rs("id") & "'>删除</a>"
							   strs="<font color=red>待受理</font>"
							  else
							   response.write rs("Accepted")
							   AcceptTime=RS("AcceptTime")
							   strs="<font color=green>已受理</font>"
							  end if
							  response.write "<br/>处理时间：" & AcceptTime & ""
							  Response.Write "<br/>处理情况：" & strs & "</td></tr>"
							  Response.Write "<tr><td class='splittd' align='right'><a href='?action=complaints&flag=show&id=" & rs("id") & "'>查看详情</a>  " & delstr & "</td>"
					          Response.Write "</tr></table></td></tr>"
						   
							RS.MoveNext
							I = I + 1
							If I >= MaxPerPage Then Exit Do
						 Loop
						 response.write str
						 %>
							 <tr>
								 <td align="right" height="50">
									 <%=KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)%>
								  </td>
							 </tr>
					<%	
						
				End If
	     %>
		 </table>
		<%End If
  
   End Sub
   
   
   
   '我的评论
   Sub Comment()
   Dim SQLStr,RS
   Dim Param:Param=" Where UserName='" & KSUser.UserName & "'"
   
   if ks.s("flag")="del" then
   		 Conn.Execute("Delete From KS_Comment Where ID=" & KS.ChkClng(KS.S("ID")) & " And ChannelID=" & KS.ChkClng(KS.S("ChannelID")) & " And UserName='" & KSUser.UserName & "'")
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")

   end if
   
   %>
    <table width="100%" class="myCollection" align="center" border="0" cellspacing="1" cellpadding="1">
                              <%
								If Action="My" Then 
							   	SqlStr="Select c.ID,c.Content,c.AddDate,c.Point,c.Verific,c.ChannelID,c.InfoID,c.replycontent From KS_Comment c inner join KS_ItemInfo I on c.infoid=i.infoid  Where i.inputer='" & KSUser.UserName & "' order by c.adddate desc"
								Else
							   	SqlStr="Select ID,Content,AddDate,Point,Verific,ChannelID,InfoID,replycontent From KS_Comment c" & Param & " order by adddate desc"
								End If

								Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open SqlStr,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td height=50 class=""empty"" valign=top><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe69b;</i></div>没有任何评论!</div></td></tr>"
								 Else
									totalPut = RS.RecordCount
									If CurrentPage < 1 Then	CurrentPage = 1
			
									If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
											RS.Move (CurrentPage - 1) * MaxPerPage
									Else
											CurrentPage = 1
									End If
									Dim I
									   Do While Not RS.Eof
											%>
    <table width="100%" class="myCollection" align="center" border="0" cellspacing="1" cellpadding="1">
											
									<tr class="title">
					                  <td class="ContentTitle">
											  评论内容：<%=KS.GotTopic(RS(1),70)%>
												  <%
												  if rs("replycontent")<>"" and not isnull(rs("replycontent")) then
												   response.write "<font color=red>已回复</font>"
												  end if
												  %>
									  </td>
									<tr>
									 <td class="Contenttips">发表时间：<%=KS.GetTimeFormat(rs(2))%>
												 <br/>状态：
												  <%
												  if RS(4)=1 Then
													 Response.Write "已审"
												  else
													 Response.Write "<font color=red>未审</font>"
												 end if
												 
												SqlStr="Select ID,Title,Tid,Fname From " & KS.C_S(RS(5),2) & " Where ID=" & RS(6)
												 Dim RSI:Set RSI=Conn.Execute(SqlStr)
												 If NoT RSI.Eof Then
												  Response.Write "<br/>信息：<a href='" & KS.Get3gItemUrl(RS(5),RSI(2),RSI(0),RSI(0)&KS.Wsetting(9)) & "' target='_blank'>" & RSI(1) & "</a>"
												 End If
												 RSI.Close:Set RSI=Nothing
									   %>
									   </td>
									  <tr>
									  <td class="splittd" align="right">
									   <%	  Response.Write "<span><a href='user.asp?Action=comment&flag=del&ChannelID=" & RS(5) &"&ID="& RS(0) &"&Page=" & CurrentPage & "' onclick=""return(confirm('确定删除此评论吗？'))"" class=""box"">删除</a></span>"
												  %>
												  </td>
									  </tr>
									  </table>
													   <%
														RS.MoveNext
														I = I + 1
														If I >= MaxPerPage Then Exit Do
														Loop
									
				                End If
                         %>
                            </table>
	<%  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
   End Sub
   
   
   
   
   
   
	
	'订单信息
	Sub Order()
	  Dim SQLStr,RS,TotalPut,MaxPerPage
	  
	  MaxPerPage=5
	
					  Dim Param:Param=" Where UserName='" & KSUser.UserName & "'"
					  If KS.S("OrderStatus")<>"" Then 
					    Param=Param & " and status=" & KS.ChkClng(KS.S("OrderStatus"))
					  End If
					  If KS.S("KeyWord")<>"" Then  
					    Param=Param & " and OrderID like '%" & KS.S("KeyWord") & "%'"
					  End If
					  
						 SqlStr="Select * From KS_Order " & Param & " order by id desc"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

				If RS.EOF And RS.BOF Then
					  Response.Write "<div style=""height:80px; line-height:80px;text-align:center"" class=""order_no""><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe6a2;</i></div>您没有任何订单!</div></div>"
				Else
					totalPut = RS.RecordCount
					If (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
					End If
					
					Dim i,MoneyTotal,MoneyReceipt
			%>
				<div class="myOrderL">
			<%
                 Do While Not RS.Eof
		%>    <table width="100%" class="border" align="center" border="0" cellspacing="1" cellpadding="1" >
                  
		
                <tr class="title">
					<td class="ContentTitle">
				   订单编号：<a href="?Action=showorder&ID=<%=RS("ID")%>"><%=rs("orderid")%></a>         
				    <%if rs("ordertype")="1" then
				  response.write "<font color=red><b><i>团</li></b></font>"
				  end if
				 %>
				</td>
			   </tr>
			   <tr>
				<td class="Contenttips">下单时间：<%=KS.GetTimeFormat(rs("inputtime"))%>
				    <br/>赠送积分：
					   <%
					    if rs("totalscore")>0 and rs("DeliverStatus")<>3 then
						   response.write "<font color=green>" & rs("totalscore") & " 分</font>"
						   if rs("scoretf")=1 then
						     response.write "<font color=#999999>,已送</font>"
						   else
						     response.write "<font color=red>,未送</font>"
						   end if
						else
						   response.write "无"
						end if
					    %>
					<br/>需要发票：
					
											<%If RS("NeedInvoice")=1 Then
											     Response.Write "<Font color=red>需要</font>"
											  	 If RS("Invoiced")=1 Then
												   Response.Write "<font color=green>(已开)</font>"
												  Else
												   Response.Write "<font color=red>(未开)</font>"
												  End If
                                              Else
											    Response.Write "-"
											  End If
											 
											  %>
				  <br/>订单状态：
											<%If RS("Status")=0 Then
												  Response.Write "<font color=red>等待确认</font>"
												  ElseIf RS("Status")=1 Then
												  Response.WRITE "<font color=green>已经确认</font>"
												  ElseIf RS("Status")=2 Then
												  Response.Write "<font color=#a7a7a7>已结清</font>"
												  ElseIf RS("Status")=3 Then
												  Response.Write "<font color=#a7a7a7>无效订单</font>"
				                              End If%> 
											<%
										if rs("alipaytradestatus")<>"" and RS("Status")<>2 then
				  select case rs("alipaytradestatus")
				    Case "WAIT_BUYER_PAY" Response.Write "<font color=red>等待汇款</font>"
					Case "WAIT_SELLER_SEND_GOODS" Response.Write "<font color=brown>已付款等待发货</font>"
					Case "WAIT_BUYER_CONFIRM_GOODS" Response.Write "<font color=blue>等待买家确认收货</font>"
					Case "TRADE_FINISHED" Response.Write "<font color=#a7a7a7>交易完成</font>"
				  end select
				else
					if rs("paystatus")="100" then
					  Response.WRITE "<font color=""green"">凭单消费</font>"
					elseif rs("paystatus")="3" then
					  Response.WRITE "<font color=blue>退款</font>"
					elseIf RS("MoneyReceipt")<=0 Then
					   Response.Write "<font color=red>等待汇款</font>"
					ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
					   Response.WRITE "<font color=blue>已收定金</font>"
					Else
					   Response.Write "<font color=green>已经付清</font>"
					End If
				end if	  %>
				 &nbsp;			
							<% If RS("DeliverStatus")=0 Then
											 Response.Write "<font color=red>未发货</font>"
											 ElseIf RS("DeliverStatus")=1 Then
											  Response.Write "<font color=blue>已发货</font>"
											 ElseIf RS("DeliverStatus")=2 Then
											  Response.Write "<font color=green>已签收</font>"
											 ElseIf RS("DeliverStatus")=3 Then
											  Response.Write "<font color=#ff6600>退货</font>"
											 End If
						 %></td>

                          </tr>
						  <tr class="orderMoney">
						    <td>应付金额:￥<span><%=formatnumber(rs("MoneyTotal"),2,-1)%></span>
					&nbsp;&nbsp;&nbsp;&nbsp;已付:￥<span><%=formatnumber(rs("MoneyReceipt"),2,-1)%></span></td>
						  </tr>
					 <tr>
				<td class="splittd F5BorderT" style="text-align:right;padding: 0.1rem 0.25rem;">
		 <% 
If RS("Status")=3 Then
		response.write "本订单在指定时间内没有付款,已作废!"
Else
		 if rs("status")=0 and rs("DeliverStatus")=0 and rs("MoneyReceipt")=0 Then%>
		 <input class="button" type='button' name='Submit' value='删除订单' onClick="javascript:if(confirm('确定要删除此订单吗？')){window.location.href='?Action=delorder&ID=<%=rs("id")%>';}">
		 <%end if%>
		 <%If RS("MoneyReceipt")<RS("MoneyTotal") and rs("paystatus")<>3 and rs("paystatus")<>100 Then%>
		 
		 <input class="button" type='button' name='Submit' value='在线支付' onClick="window.location.href='user.asp?Action=payshoporder&ID=<%=rs("id")%>'">
		 <input class="button" type='button' name='Submit' value='余额扣款' onClick="window.location.href='?Action=addpayment&ID=<%=rs("id")%>'">&nbsp;&nbsp;
		 <%end if%>
		 <% if rs("DeliverStatus")=1 Then%>
		 <input class="button" type='button' name='Submit' value='签收商品' onClick="window.location.href='?Action=signup&ID=<%=RS("ID")%>'">
		 <%end if%>
		 <%
		 end if
             If RS("Status")<>2 Then
			   If RS("MoneyReceipt")>=RS("MoneyTotal") and  RS("PayStatus")<>"3"  and rs("DeliverStatus")<>3 Then
			   %>
			   <%if rs("totalscore")>0  and rs("DeliverStatus")<>3 and rs("usescoremoney")<=0 then%>
			   <input type="button" value="满意无需退换货,立即获得<%=totalscore%>分积分" onClick="if (confirm('结清订单后将不可以再申请退换货,确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
			   <%else%>
			   <input type="button" value="结清订单" onClick="if (confirm('结清订单后将不可以再申请退换货,确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
			   <%end if%>
			  <%ElseIf RS("PayStatus")="3" or rs("DeliverStatus")=3 Then
			    if rs("usescore")>0 then
				%>
			   <input type="button" value="结清订单,返还我的<%=rs("usescore")%>分积分" onClick="if (confirm('结清订单将立即返还您的积分，确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
				<%
				else
				%>
			   <input type="button" value="结清订单" onClick="if (confirm('确定结清吗？')){location.href='?action=setorderok&id=<%=rs("id")%>';}" class="button" />
				<%
				end if
			   end if
			End If
		 %>
		 </td>
      </tr>

	  </table>
                        <%
							MoneyReceipt=RS("MoneyReceipt")+MoneyReceipt
							MoneyTotal=RS("MoneyTotal")+MoneyTotal
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
</div>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="TotalAmount">
	   <tr><td colspan="2"></td></tr>
       <tr align='center' class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">           
	   <td align='left'>本页合计:￥<span><%=formatnumber(MoneyTotal,2)%></span></td>                  
	   <td align='left'>已收款:￥<span><%=formatnumber(MoneyReceipt,2)%></span></td>          
		</tr> 
     <tr align='center' class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'" >           
	 <td align='left'>所有总计:￥<span><%=formatnumber(Conn.execute("Select sum(moneytotal) from KS_Order Where UserName='" & KSUser.UserName & "'")(0),2)%></span></td>                  
	  <td align='left'>已收款:￥<span><%=formatnumber(Conn.execute("Select sum(MoneyReceipt) from KS_Order Where UserName='" & KSUser.UserName & "'")(0),2)%></span></td>           
	  </tr> 
 </table>
 
         <%
				End If
           
	
	
	End Sub
	
	'返回订单详细信息
		Function  OrderDetailStr(RS)
		 OrderDetailStr="<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr align='center' class='title'>    <td height='22'><b>订 单 信 息</b>（编号：" & RS("ORDERID") & "）</td>  </tr>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr>" & vbcrlf
		 OrderDetailStr=OrderDetailStr & " <td height='25'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & " 客户姓名：" & RS("Contactman") & "<br/>用 户 名：" & rs("username") & "<br/>获赠积分："
		if rs("totalscore")=0  or rs("DeliverStatus")=3 then
			OrderDetailStr=OrderDetailStr & "无"
		else
			if rs("scoretf")=1 then
			OrderDetailStr=OrderDetailStr & "<font color=green>" & rs("totalscore") & "分,已送出</font>"
			else
			OrderDetailStr=OrderDetailStr & "<font color=red>" & rs("totalscore") & "分,未送出</font>"
			end if
		end if
		OrderDetailStr=OrderDetailStr & "<br/>下单时间：" & formatdatetime(rs("inputtime"),2)      
		OrderDetailStr=OrderDetailStr & "<br/>需要发票："
			    If RS("NeedInvoice")=1 Then
				  OrderDetailStr=OrderDetailStr & "<Font color=red>√</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "		&nbsp;已开发票："	
				  If RS("Invoiced")=1 Then
				   OrderDetailStr=OrderDetailStr & "<font color=green>√</font>"
				  Else
				   OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	<br/>订单状态："	
			if RS("Status")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>等待确认</font>"
				  ElseIf RS("Status")=1 Then
				 OrderDetailStr=OrderDetailStr & "<font color=green>已经确认</font>"
				  ElseIf RS("Status")=2 Then
				 OrderDetailStr=OrderDetailStr & "<font color=#a7a7a7>已结清</font>"
				  End If
				 OrderDetailStr=OrderDetailStr & "&nbsp;"
		if rs("paystatus")="100" then
				OrderDetailStr=OrderDetailStr & "<font color=""green"">凭单消费</font>"
		elseif rs("paystatus")="3" then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>退款</font>"
		   else	
			     If RS("MoneyReceipt")<=0 Then
				   OrderDetailStr=OrderDetailStr & "<font color=red>等待汇款</font>"
				  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>已收定金</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=green>已经付清</font>"
				  End If
           end if
       OrderDetailStr=OrderDetailStr & "&nbsp;"
				if RS("DeliverStatus")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>未发货</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已发货</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已签收</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  OrderDetailStr=OrderDetailStr & "<font color=#ff6600>退货</font>"
				 End If
	OrderDetailStr=OrderDetailStr & "	<hr>收货人姓名：" & rs("contactman")
	OrderDetailStr=OrderDetailStr & "	<br/>联系电话：" & rs("phone")
	OrderDetailStr=OrderDetailStr & "	<br/>收货人地址：" & rs("address")     
	OrderDetailStr=OrderDetailStr & "	<br/>邮政编码：" &rs("zipcode")
	OrderDetailStr=OrderDetailStr & "	<br/>收货人邮箱：" & rs("email")     
	OrderDetailStr=OrderDetailStr & "	<br/>收货人手机：" & rs("mobile")
	OrderDetailStr=OrderDetailStr & "<br/>付款方式：" & KS.ReturnPayMent(rs("PaymentType"),0)
	if rs("tocity")="" then
    OrderDetailStr=OrderDetailStr & "	<br/>>送货方式：免运费订单，由商家指定" 
	else
    OrderDetailStr=OrderDetailStr & "<br/>快递公司：" 
	
	  dim rst,foundexpress
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
	    OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
	  If DataBaseType=1 Then
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
	  Else
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
	  End If
	  if rst.eof then
	  else
	   OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	  end if
	 rst.close : set rst=nothing
	 End If
	
	
	OrderDetailStr=OrderDetailStr & " 发往<span style='color:red'>" & rs("tocity") & "</span>"
	end if
	
	if  RS("NeedInvoice")=1 then
	  OrderDetailStr=OrderDetailStr & "<br/>发票信息："& replace(rs("InvoiceContent"),chr(10),"<br/>") 
	end if
	if Not KS.IsNul(rs("Remark")) Then
    OrderDetailStr=OrderDetailStr & "<br/>备注/留言：" & rs("Remark")
	End If
	OrderDetailStr=OrderDetailStr & "	<hr/><h3>商品列表：</h3>		</td>  "
	OrderDetailStr=OrderDetailStr & "		</tr>  "
	OrderDetailStr=OrderDetailStr & "		<tr><td>"

			 Dim TotalPrice,attributecart,RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			   RSI.Open "Select * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id",conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
				' Response.Write "<script>alert('找不到相关商品');history.back();<//script>"
			  Else
			  Do While Not RSI.Eof
			  attributecart=rsi("attributecart")
			  if not ks.isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
		OrderDetailStr=OrderDetailStr & "	商品名称：" 
		 Dim OrderType:OrderType=KS.ChkClng(rs("ordertype"))
		 If OrderType=1 Then
		  OrderDetailStr=OrderDetailStr & "<a href='../shop/groupbuyshow.asp?id=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 subject from ks_groupbuy where id=" & rsi("proid"))(0)
		 Else
		  OrderDetailStr=OrderDetailStr & "<a href='show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 title from ks_product where id=" & rsi("proid"))(0) 
		 End If
		If RSI("IsChangedBuy")="1" Then OrderDetailStr=OrderDetailStr & "(换购)"
		
		
			  Dim SqlStr,RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  If OrderType=1 Then
			    SqlStr="Select top 1 Subject as title,'件' as unit,0 as IsLimitBuy,0 as LimitBuyPrice,0 as LimitBuyPayTime From KS_GroupBuy Where ID=" & RSI("ProID")
			  Else
			    SqlStr="Select top 1 I.Title,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID")
			  End If
			  RSP.Open SqlStr,conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(换购)"
				  else
				    if RSP("LimitBuyPayTime") then
					   If LimitBuyPayTime="" Then
					   LimitBuyPayTime=RSP("LimitBuyPayTime")
					   ElseIf LimitBuyPayTime>RSP("LimitBuyPayTime") Then
						LimitBuyPayTime=RSP("LimitBuyPayTime")
					   End If
					end if
				  end  if
				  If RSI("IsLimitBuy")="1" Then OrderDetailStr=OrderDetailStr & "<span style='color:green'>(限时抢购)</span>"
				  If RSI("IsLimitBuy")="2" Then OrderDetailStr=OrderDetailStr & "<span style='color:blue'>(限量抢购)</span>"
			  End If
			  RSP.Close:Set RSP=Nothing
		
		OrderDetailStr=OrderDetailStr & "</a>" & attributecart & "<Br/>购买数量：" & rsi("amount") & Unit &"<br/>总价：" & formatnumber(rsi("realprice")*rsi("amount"),2) & "元，赠送积分：" & ks.chkclng(rsi("score")*rsi("amount")) & " 分" 
		totalscore=totalscore+ks.chkclng(rsi("score")*rsi("amount"))
		Set RSP=Conn.Execute("Select Top 1 DownUrl From KS_Product Where ID=" & RSI("ProID"))
		If Not RSP.Eof Then
			If Not KS.IsNul(RSP("DownUrl")) Then
				If RS("MoneyReceipt")>=RS("MoneyTotal") Then
				  OrderDetailStr=OrderDetailStr & "<a href='?action=OrderDown&orderid=" & rs("id") & "&proid=" & rsi("proid") &"'><img src='../images/default/download.gif'></a>"
				Else
				 OrderDetailStr=OrderDetailStr & "<a href='#' disabled>未付清</a>"
				End If
			Else
				 OrderDetailStr=OrderDetailStr & "---"
			End If
		Else
		  OrderDetailStr=OrderDetailStr & "---"
		End If
		RSP.Close :Set RSP=Nothing
		
		OrderDetailStr=OrderDetailStr & "<hr/> " 
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  '取得捆绑销售商品
		
		
			  TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
		End If
		
		OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '超值礼包
		
		
		OrderDetailStr=OrderDetailStr & "	<b>合计：" & formatnumber(totalprice,2) & "元</b> "
		
		OrderDetailStr=OrderDetailStr & "	<br/>付款方式折扣率：" & rs("Discount_Payment") & "%&nbsp;&nbsp;" 
	   If RS("Weight")>0 Then
	   OrderDetailStr=OrderDetailStr & "重量：" & rs("weight") & " KG"
	   End If
	   OrderDetailStr=OrderDetailStr & "&nbsp;&nbsp;运费：" & rs("Charge_Deliver")&" 元&nbsp;&nbsp;&nbsp;&nbsp;税率：" & KS.Setting(65) &"%&nbsp;&nbsp;&nbsp;&nbsp;价格含税："
				IF KS.Setting(64)=1 Then 
				   OrderDetailStr=OrderDetailStr & "是"
				  Else
				   OrderDetailStr=OrderDetailStr & "不含税"
				  End If
				  Dim TaxMoney
				  Dim TaxRate:TaxRate=KS.Setting(65)
				 If KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+TaxRate/100

				OrderDetailStr=OrderDetailStr & "<br>实际金额：(" & rs("MoneyGoods") & "×" & rs("Discount_Payment") & "%＋"&rs("Charge_Deliver") & ")×"
				if KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then OrderDetailStr=OrderDetailStr & "100%" Else OrderDetailStr=OrderDetailStr & "(1＋" & TaxRate & "%)" 
				OrderDetailStr=OrderDetailStr & "＝" & formatnumber(rs("NoUseCouponMoney"),2) & "元 "
    OrderDetailStr=OrderDetailStr & "<Br/><b>总金额：</b> ￥" & formatnumber(rs("NoUseCouponMoney"),2) & " 元"
	If KS.ChkClng(RS("CouponUserID"))<>0 and RS("UseCouponMoney")>0 Then
	OrderDetailStr=OrderDetailStr & "<b>使用优惠券：</b> <font color=#ff6600>￥" & formatnumber(RS("UseCouponMoney"),2) & " 元</font><br>"
    ElseIf RS("UseScoreMoney")<>"0" Then
	OrderDetailStr=OrderDetailStr & "<b>花费<font color=green>" &RS("UseScore") & "</font>积分抵扣了<font color=#ff6600>" & formatnumber(RS("UseScoreMoney"),2) & "</font>元<br>"
	End If
	OrderDetailStr=OrderDetailStr & "&nbsp;<b>应付：</b> ￥" & formatnumber(rs("MoneyTotal"),2) & "  元<Br/><b>已付款：</b>￥<font color=red>" & formatnumber(rs("MoneyReceipt"),2) & "</font>元</b>"
	If RS("MoneyReceipt")<RS("MoneyTotal") Then
	OrderDetailStr=OrderDetailStr & "&nbsp;<B>尚欠款：￥<font color=blue>" & formatnumber(RS("MoneyTotal")-RS("MoneyReceipt"),2) &"元</B>"
	End If
	OrderDetailStr=OrderDetailStr & "</td> "
	OrderDetailStr=OrderDetailStr & "</tr>"  

	
	If not conn.execute("select top 1 * from ks_orderitem where orderid='" & RS("OrderID") &"' and islimitbuy<>0").eof Then
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:red;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单是限时/限量抢购订单,限制下单后" & LimitBuyPayTime & "小时之内必须付款,即如果您在[" & DateAdd("h",LimitBuyPayTime,RS("InputTime")) & "]之前用户没有付款,本订单自动作废。</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	If RS("DeliverStatus")=1 Then
	 Dim RSD,DeliverStr
	 Set RSD=Conn.Execute("Select Top 1 * From KS_LogDeliver Where DeliverType=1 And OrderID='" & RS("OrderID") & "'")
	 If Not RSD.Eof Then
	  DeliverStr="快递公司:" & RSD("ExpressCompany") & " 物流单号:" & RSD("ExpressNumber") & " 发货日期:" & RSD("DeliverDate") & " 发货经手人:" & RSD("HandlerName")
	 End If
	 RSD.Close : Set RSD=Nothing
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:blue;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单已发货。" & DeliverStr & "</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	
	OrderDetailStr=OrderDetailStr & "	</table>"
	  End Function
	
	Sub ShowOrder()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * from ks_order where username='" & KSUser.UserName & "' and id=" & ID ,conn,1,1
		 IF RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   response.end
		 End If
		 
		response.write "<br>"
		response.write OrderDetailStr(RS)
         %><br/>
		 
		 <div align=center id='buttonarea'>
		 <% 
If RS("Status")=3 Then
		response.write "本订单在指定时间内没有付款,已作废!"
Else
		 if rs("status")=0 and rs("DeliverStatus")=0 and rs("MoneyReceipt")=0 Then%>
		 <input class="button" type='button' name='Submit' value='删除订单' onClick="javascript:if(confirm('确定要删除此订单吗？')){window.location.href='?Action=delorder&ID=<%=rs("id")%>';}">&nbsp;&nbsp;
		 <%end if%>
		 <%If RS("MoneyReceipt")<RS("MoneyTotal") and rs("paystatus")<>3 and rs("paystatus")<>100 Then%>
		 <span>
		 <input class="button" type='button' name='Submit' value='在线支付' onClick="window.location.href='user.asp?Action=payshoporder&ID=<%=rs("id")%>'">
		 </span>
		 <input class="button" type='button' name='Submit' value='余额扣款' onClick="window.location.href='?Action=addpayment&ID=<%=rs("id")%>'">&nbsp;&nbsp;
		 <%end if%>
		 <% if rs("DeliverStatus")=1 Then%>
		 <input class="button" type='button' name='Submit' value='签收商品' onClick="window.location.href='?Action=signup&ID=<%=RS("ID")%>'">
		 <%end if%>
		 <%
		 end if
             If RS("Status")<>2 Then
			   If RS("MoneyReceipt")>=RS("MoneyTotal") and  RS("PayStatus")<>"3"  and rs("DeliverStatus")<>3 Then
			   %>
			   <%if totalscore>0  and rs("DeliverStatus")<>3 and rs("usescoremoney")<=0 then%>
			   <input type="button" value="满意无需退换货,立即获得<%=totalscore%>分积分" onClick="if (confirm('结清订单后将不可以再申请退换货,确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
			   <%else%>
			   <input type="button" value="结清订单" onClick="if (confirm('结清订单后将不可以再申请退换货,确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
			   <%end if%>
			  <%ElseIf RS("PayStatus")="3" or rs("DeliverStatus")=3 Then
			    if rs("usescore")>0 then
				%>
			   <input type="button" value="结清订单,返还我的<%=rs("usescore")%>分积分" onClick="if (confirm('结清订单将立即返还您的积分，确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
				<%
				else
				%>
			   <input type="button" value="结清订单" onClick="if (confirm('确定结清吗？')){location.href='?action=setok&id=<%=rs("id")%>';}" class="button" />
				<%
				end if
			   end if
			End If
		 %>
		&nbsp; <input class="button" type='button' name='Submit' value='订单首页' onClick="location.href='?action=order';">
		 </div>
		 <br />
	<%if rs("isservice")="1" then%>
		<a name="service"></a><strong>服务记录明细：<br/></strong>
		 <%
		         dim times,sytimes,validity,firstservicetime
				  times=conn.execute("select count(1) from ks_orderservice where orderid=" & rs("id"))(0)
				  if times>rs("servicetimes") then sytimes=0 else sytimes=rs("servicetimes")-times
				  dim rsi:set rsi=conn.execute("select top 1 adddate from ks_orderservice where orderid=" & rs("id"))
				  if not rsi.eof then
					firstservicetime=rsi(0)
					validity=dateadd("m",rs("validity"),firstservicetime)
				  else
					validity=dateadd("m",rs("validity"),now)
				  end if
				  rsi.close
				  set rsi=nothing
				  %>
				   <div style="border:#B2D9F6 1px solid; line-height:26px;padding-left:5px;margin-bottom:10px;background:#F3F9FF;">
				  
				  服务商品名称：<%=rs("servicename")%>&nbsp;&nbsp;服务次数：<%=rs("servicetimes")%>次,剩余：<font color=red><%=sytimes%></font>次&nbsp;服务有效期：<%=rs("validity")%>个月,载止日期：<%=year(validity) & "-" & month(validity) & "-" & day(validity)%>
				   
				   </div>
				 <table  cellpadding="1" style="margin-bottom:6px;border:1px solid #999;" cellspacing="1" width="100%">

				   <tr style="background:#f1f1f1;height:23px;text-align:center">
					  <td width="50">次数</td>
					  <td width="350">内容</td>
					  <td width="70">时间</td>
					  <td width="70">签收人</td>
				   </tr>
				   <%
				   dim rss:set rss=server.CreateObject("adodb.recordset")
				   RSS.Open "select * from ks_orderservice where orderid=" & rs("id") & " order by id desc",conn,1,1
				   if RSS.Eof Then
					str="<tr><td colspan=4 class=""splittd empty""><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe70c;</i></div>没有找到服务记录!</div></td></tr>"
				   Else
					dim totalnum:totalnum=rss.recordcount
					dim str,num,qsr
					num=0
					if totalnum<5 then
					str="<tr><td colspan=4><div>"
					else
					str="<tr><td colspan=4><div style=""overflow-x:hidden;overflow-y:auto;height:130px"">"
					end if
					do while not rss.eof
					str=str &"<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
					str=str &"<tr id='tr1" & rss("id") & "'>"
					str=str &"<td width=""50"" height=""25"" class=""splittd"">第" & totalnum-num & "次</td>"
					str=str &"<td width=""350"" class=""splittd"" style=""width:290px;word-break:break-all;"">" & rss("content") & "</td>"
					str=str &"<td width=""70"" class=""splittd"" style='text-align:center'>" & year(rss("adddate")) & "-" & month(rss("adddate")) & "-" & day(rss("adddate")) & "</td>"
					qsr=rss("qsr")
					if ks.isnul(qsr) then qsr="---"
					str=str &"<td width=""70"" class=""splittd"" style='text-align:center'>&nbsp;" & qsr & "&nbsp;</td>"
					str=str &"</tr></table>"
					num=num+1
					rss.movenext
					loop
					str=str & "</div></td></tr>"
				  end if
				  rss.close
					response.write str
					
				  %>
				 </table>
		<%
		end if
		 rs.close:set rs=nothing
		End Sub
		
		  
'取得捆绑销售商品
Dim ProIds
Function GetBundleSalePro(ByRef TotalPrice,ProID,OrderID)
  If KS.FoundInArr(ProIDS,ProID,",")=true Then Exit Function
  ProIds=ProIDs & "," & ProID
  Dim Str,RS,XML,Node
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "Select I.Title,I.Unit,O.* From KS_OrderItem O inner join KS_Product I On O.ProID=I.ID Where O.SaleType=6 and BundleSaleProID=" & ProID & " and O.OrderID='" & OrderID & "' order by O.id",conn,1,1
  If Not RS.Eof Then
    Set XML=KS.RsToXml(rs,"row","")
  End If
  RS.Close:Set RS=Nothing
  If IsObject(XML) Then
	     str=str & "<div style=""color:green"">选购捆绑促销:</div>"
       For Each Node In Xml.DocumentElement.SelectNodes("row")
         str=str & "<table><tr>"
		 str=str &" <td>选购商品：" & Node.SelectSingleNode("@title").text &" <br/>选购数量：" & Node.SelectSingleNode("@amount").text & Node.SelectSingleNode("@unit").text& "<br/>价格：<font color=brown>￥" & formatnumber(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2,-1) &"</font>元</td>"
		 str=str & "</tr></table>"
		 TotalPrice=TotalPrice +round(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2) 
       Next
  End If
  GetBundleSalePro=str
End Function
	  
	  
 '得到超值礼包
 Function GetPackage(ByRef TotalPrice,OrderID)
	    If KS.IsNul(OrderID) Then Exit Function
		Dim RS,RSB,GXML,GNode,str,n,Price
		Set RS=Conn.Execute("select packid,OrderID from KS_OrderItem Where SaleType=5 and OrderID='" & OrderID & "' group by packid,OrderID")
		If Not RS.Eof Then
		 Set GXML=KS.RsToXml(Rs,"row","")
		End If
		RS.Close : Set RS=Nothing
		If IsOBJECT(GXml) Then
		   FOR 	Each GNode In GXML.DocumentElement.SelectNodes("row")
		     Set RSB=Conn.Execute("Select top 1 * From KS_ShopPackAge Where ID=" & GNode.SelectSingleNode("@packid").text)
			 If Not RSB.Eof Then
					  


						Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
						RSS.Open "Select a.title,a.Price_Member,a.Price,b.* From KS_Product A inner join KS_OrderItem b on a.id=b.proid Where b.SaleType=5 and b.packid=" & GNode.SelectSingleNode("@packid").text & " and  b.orderid='" & OrderID & "'",Conn,1,1
						  str=str & "礼包名称：<strong><a href='../shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a></strong>"
						  n=1
						  Dim TotalPackPrice,tempstr,i
						  TotalPackPrice=0 : tempstr=""
						Do While Not RSS.Eof
						 
						  For I=1 To RSS("Amount") 
							  '得到单件品价格 
							  If RSS("AttrID")<>0 Then 
							  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RSS("AttrID"))
							  If Not RSAttr.Eof Then
								Price=RSAttr("Price")
							  Else
								Price=RSS("Price_member")
							  End If
							  RSAttr.CLose:Set RSAttr=Nothing
							 Else
								Price=RSS("Price_member")
							 End If
							
							   TotalPackPrice=TotalPackPrice+Price
							  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
							  n=n+1
						  Next
						  RSS.MoveNext
						Loop
						
						str=str &"<br/>您选择的套装详细如下：<br/>" & tempstr & "礼包数量：1 &nbsp;礼包金额：<font color=green>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) &"</font><br/>"
					   
						str=str & "<hr/>" 
						
						TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1))   '将礼包金额加入总价
						
						RSS.Close
						Set RSS=Nothing
			End If
			RSB.Close
		   Next
			
	    End If
		GetPackage=str
End Function
'下载
Sub OrderDown()
  Dim OrderID:OrderID=KS.ChkClng(KS.S("OrderID"))
  Dim ProID:ProID=KS.ChkClng(KS.S("ProID"))
  If ProID=0 Or OrderID=0 Then KS.AlertHintScript "出错了！！！"
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "Select top 1 O.* From KS_Order O Inner Join KS_OrderItem I ON O.OrderID=I.OrderID Where  O.Id=" & OrderID & " And O.MoneyReceipt>=O.MoneyTotal",Conn,1,1
  If RS.Eof And RS.Bof Then
   RS.Close :Set RS=Nothing
   KS.AlertHintScript "订单不存或是订单款项还没有付清，无法下载！！!"
  Else
    RS.Close
	RS.Open "Select top 1 DownUrl From KS_Product Where ID=" & ProID,conn,1,1
	If RS.EOf And RS.Bof Then
	 RS.Close :Set RS=Nothing
	 KS.AlertHintScript "下载已不存在！"
	Else
	 DownURL=RS(0)
	 RS.Close :Set RS=Nothing
	End If
	If Not KS.IsNul(DownUrL) Then Response.Redirect DownUrl
  End If
  
End Sub

Sub AddPayment()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  response.write "<script>alert('出错啦!');history.back();</script>":response.end
		 End If
		 
		 If KS.ChkCLng(KS.Setting(49))=1 Then
		  If RS("Status")=0 Then
		    RS.Close:Set RS=Nothing
		   	KS.Die "<script>alert('对不起，该订单还未确认，本站启用只有后台确认过的订单才能付款!');history.back();</script>"
		  End If
		End If
		 %>
  <FORM name="form4" onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦确认就不可更改哦！')" action=user.asp method=post>
  <div class="fgTitle">使用账户资金支付订单</div>
  <div class="tableBGw">
  <table class=Cpayment cellSpacing=1 cellPadding=2 width="100%" align="center" border=0>
    <tr class=tdbg>
      <td align="left" class="clefttitle" nowrap>用户名</td><td class="aRight"><%=KSUser.UserName%></td>
    </tr>
    <tr class=tdbg>
      <td align="left" class="clefttitle" nowrap>客户名称</td><td class="aRight"><%=RS("ContactMan")%></td>
    </tr>
    <tr class=tdbg>
      <td align="left" class="clefttitle" nowrap>资金余额</td><td class="aRight"><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1)%>元<%if Round(KSUser.GetUserInfo("Money"),2)<=0 then response.write "<a href=""user_payonline.asp"">资金不足,请点此充值</a>"%></td>
    </tr>
    <tr class=tdbg>
      <td align="left" class="clefttitle" nowrap>订单编号</td><td class="aRight"><%=RS("OrderID")%></td>
    </tr>
    <tr class=tdbg>
      <td align="left" class="clefttitle" nowrap>订单金额</td><td class="aRight"><font color=red><%=formatnumber(RS("MoneyTotal"),2,-1)%></font> 元&nbsp;&nbsp;&nbsp;已付款：<font color=blue><%=formatnumber(RS("MoneyReceipt"),2,-1)%></font>元</td>
    </tr>
    <tr class=tdbg>
      <td align="left" class="clefttitle" nowrap>支出金额</td><td class="aRight"><Input id="Money" readonly maxLength=20 class="textbox" size=8 value="<%=rs("moneytotal")-rs("MoneyReceipt")%>" name="Money">元</td>
    </tr>
    <tr class=tdbg>
      <td height=30 colspan="2"  class="aRight" style="text-align:left;border-bottom: none;line-height: 1rem;"><font color=#ff0000>注意：支出信息一旦录入，就不能再修改！所以在保存之前确认输入无误！</font></td>
    </tr>
    <tr class=tdbg align=middle>
      <td height=30 colspan="2">
        <Input id=Action type="hidden" value="savepayment" name="Action"> 
        <Input id=ID type=hidden value="<%=rs("id")%>" name="ID"> 
        <Input type=submit value=" 确认支付 " class="button" name=Submit>
	  </td>
    </tr>
  </table>
 </div>
</FORM>
		 <%
		 rs.close:set rs=nothing
		End Sub
		
		'开始余额支付操作
		Sub SavePayment()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim Money:Money=KS.S("Money")
		 If Not IsNumeric(Money) Then Response.Write "<script>alert('请输入有效的金额!');history.back();</script>":Response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('出错啦!');history.back();</script>"
		 End If
		 If Round(Money,2)>Round(KSUser.GetUserInfo("Money"),2) or Round(KSUser.GetUserInfo("Money"),2)<=0  Then
		  %>
		  <br><br>
		  <table cellpadding=2 cellspacing=1 border=0 width="100%" class='border' align=center>
		  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>
		  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b><li>您输入的支付金额超过了您的资金余额，无效支付！</li></td></tr>
		  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>
		</table>
		  <%
		  RS.Close:Set RS=Nothing:Exit Sub
		 End If
		   RS("MoneyReceipt")=RS("MoneyReceipt")+Money
		   Dim OrderStatus:OrderStatus=rs("status")
		   RS("Status")=1
		   RS("PayTime")=now   '记录付款时间
		   RS.Update
		   If RS("MoneyReceipt")>=RS("MoneyTotal") Then
		  	 RS("PayStatus")=1  '已付清
		  ElseIf RS("MoneyReceipt")<>0 Then
		     RS("PayStatus")=2  '已收定金
		  Else
		     RS("PayStatus")=0  '未付款
		  End If
		  RS.Update

		   Call KS.MoneyInOrOut(RS("UserName"),RS("Contactman"),Money,4,2,now,RS("OrderID"),KSUser.UserName,"支付订单费用，订单号：" & RS("Orderid"),0,0,0)

	
					'====================更新库存量========================
					If RS("MoneyReceipt")>=RS("MoneyTotal") Then
						Dim rsp:set rsp=conn.execute("select id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
						do while not rsp.eof
						  dim rsi:set rsi=conn.execute("select amount,attrid from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(0))
						  if not rsi.eof then
							  if OrderStatus<>1 Then  '扣库存量
							   If RSI("AttrID")<>0 Then
								  Conn.Execute("update KS_ShopSpecificationPrice set amount=amount-" & RSI(0) & " Where amount>=" & RSI(0) & " and ID=" & RSI(1))
							  Else
							   conn.execute("update ks_product set totalnum=totalnum-" & rsi(0) &" where totalnum>=" & rsi(0) &" and id=" & rsp(0))        
							  End If
							  End If
						  end if
						  rsi.close
						  set rsi=nothing
						rsp.movenext
						loop
						rsp.close
						set rsp=nothing
					End If
					'================================================================
		 
		 RS.Close:Set RS=Nothing
		  Response.Redirect "?Action=showorder&id=" & id 
		End Sub

		
	'签收商品
		Sub SignUp()
		 Dim OrderID,id:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  response.write "<script>alert('出错啦!');history.back();</script>":response.end
		 End If
         rs("DeliverStatus")=2
		 rs("BeginDate")=Now
		 rs.update
		 OrderID=RS("OrderID")
		 rs.close:set rs=nothing
		 Conn.execute("Update KS_LogDeliver Set Status=1 Where OrderID='" & OrderID & "'")
		 Response.Redirect "?Action=showorder&ID=" & id
		End Sub
		
			'删除订单
		Sub DelOrder()
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select OrderID,CouponUserID From KS_Order where status=0 and DeliverStatus=0 and MoneyReceipt=0 and id=" & ID,Conn,1,3
		 If Not rs.EOF Then
		   Conn.execute("Update KS_ShopCouponUser Set UseFlag=0,OrderID='' Where ID=" & rs(1))
		   Conn.execute("delete from ks_orderitem Where OrderID='" & rs(0) &"'")
		   rs.delete
		 End if
         Response.redirect request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'结清订单
		sub setorderok()
		 dim totalscore,AllianceUser,orderid,scoretf,DeliverStatus,paystatus,usescore
		 dim id:id=KS.ChkClng(Request("id"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "select top 1 * from ks_order where id=" & id & " and username='" & KSUser.UserName & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   rs.close:set rs=nothing
		   KS.AlertHintScript "出错啦，找不到订单！"
		 End If
		 totalscore=rs("totalscore")
		 orderid=rs("orderid")
		 scoretf=rs("scoretf")
		 DeliverStatus=rs("DeliverStatus")
		 paystatus=KS.ChkClng(rs("paystatus"))
		 usescore=KS.ChkClng(rs("usescore"))
		 rs.close
		 
		 
		 if totalscore>0 and scoretf="0" and DeliverStatus<>3 and paystatus<>3 then
		    Call KS.ScoreInOrOut(KSUser.UserName,1,totalscore,"系统","商城购物赠送的积分，订单号：" & orderid & "。",0,0)
		    AllianceUser=KSUser.GetUserInfo("AllianceUser")
			if not ks.isnul(AllianceUser) then
			  rs.open "select top 1 groupid from ks_user where username='" & AllianceUser &"'",conn,1,1
			  if not rs.eof then
			    if KS.U_S(rs("GroupID"),19)="1"  then   '享受推广获积分
				   dim per:per=KS.U_S(rs("GroupID"),20)
				   if not isnumeric(per) then per=0
				   if per>0 then
				      dim myscore:myscore=KS.ChkClng(totalscore*per/100)
					  if myscore>0 then
					   	Call KS.ScoreInOrOut(AllianceUser,1,myscore,"系统","您推荐的用户[" & KSUser.UserName & "]在商城购物成功,订单号：" & orderid & "，您享受该订单总赠送积分(" & totalscore & "分)的 " & per& "% 奖励。",0,0)

					  end if
				   end if
				end if
			  end if
			  rs.close
			end if
		 elseif paystatus=3 or DeliverStatus=3 and usescore>0 then  '退货或是退款时返还积分
			Session("ScoreHasUse")="-" '设置只累计消费积分
			Call KS.ScoreInOrOut(KSUser.UserName,1,usescore,"系统","购物失败，返还积分。订单号<font color=red>" & orderid & "</font>!",0,0)

		 end if
		 set rs=nothing
		 Conn.Execute("update ks_order set status=2,scoretf=1 where id=" & id)
		
		 KS.Die "<script>alert('恭喜，订单已结清!');location.href='?action=order';</script>"
		end sub
	
End Class
%>

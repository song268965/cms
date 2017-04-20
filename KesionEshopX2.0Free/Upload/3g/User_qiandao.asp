<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New User_ItemSign
KSCls.Kesion()
Set KSCls = Nothing

Class User_ItemSign
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
		  MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
       Public Sub Kesion()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		 %>
		 <!DOCTYPE html>
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

<!--<div class="navpositionbig" style=" height:40px;">
	<div class="navposition">
		<div class="logonav" style=" border-bottom:2px solid #0C9AD8">
		    
		 	 <div style="float:left;width:10%;"><a href="javascript:history.back()"><img src="images/left-arrow-vector.png" width="28" height="28"/></a></div>
			 <div style="margin:0 auto;width:80%;float:left;text-align:center;font-weight:bold;font-size:15pt;color:#0C9AD8">
			 签到记录
			 </div>
			 <div style="width:10%;text-right:right;float:right;"><a href="user.asp"><img src="images/user-male-alt-vector.png" width="28" height="28"/></a></div>
			
		</div>
	</div>
</div>	-->	
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
			<li><a href="/3g/list.asp?id=11">供求信息</a></li>
			<li><a href="/3g/list.asp?id=812">影视频道</a></li>
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
		 <% 
		
		dim Action,Rs,Content,qdxq
		if ks.Setting(201)="1" then

				call qiandaomain()
		end if	
	  %>
	</div>
</div>
<%	
  End Sub
      sub  qiandaomain()
	  	
		 If KS.S("page") <> "" Then
				CurrentPage = CInt(KS.S("page"))
		 Else
				CurrentPage = 1
		 End If
	    %>
		<% 

			%>

	
		
				<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="SignInL">
					<tr class=titlename align=middle>
					  <td  width="25%"height="25">时间</td>
					   <td width="10%">心情</td>
					  <td width="40%">内容</td>
					  <td width="25%">状态</td>
					</tr>
					<%  
						 SqlStr="Select * From KS_qiandao where username='" & ksuser.username &"'  order By adddate desc"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1
						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top><div class=""noneRe""><div class=""noneImg""><i class=""iconfont"">&#xe723;</i></div>找不到您要的记录!</div></td></tr>"
								 Else
								 totalPut = RS.RecordCount
						
								If CurrentPage < 1 Then CurrentPage = 1
								
								If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								End If
								Call ShowContent_me
				End If

						
						 %>		
</table>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		 
		<%
		
	  end sub
	
	 Sub ShowContent()
	 
     Dim I,intotalmoney,outtotalmoney,Page_s,qdxq,RSkc,Content,adddate,qiandao_xqco,qiandao_dateend,qdnow
	 Page_s=(CurrentPage-1)* MaxPerPage
     Do While Not rs.eof 
		qdnow=0 : qdxq=0 :Content=""
		qiandao_xqco= Split(rs("qiandao_xqco")&"","|1|1|")
	 	if Ubound(qiandao_xqco)>1 then
			qiandao_dateend=CDate(qiandao_xqco(2))
			if year(qiandao_dateend)= year(now()) and  month(qiandao_dateend)=month(now()) and  day(qiandao_dateend)=day(now()) then
	 			qdxq=qiandao_xqco(0)
				Content=qiandao_xqco(1)
				qdnow=1
			end if
		end if
	%>
    <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	  
      <td  class="splittd" align=middle><%=Page_s+i+1%></td>
      <td  class="splittd" align=middle ><%=rs("username")%></td>
	  <td  class="splittd" align=middle ><%=rs("qiandao")%></td>
	  <td  class="splittd" align=middle ><%=conn.execute("select count(1) from KS_qiandao where  username='" & rs("username") &"' and year(AddDate)=year(" & SqlNowString & ") and month(AddDate)=month(" & SqlNowString &") ")(0)%></td>
      <td   class="splittd" align=middle width=60><%	  
	  %>
	  <img src="/images/emot/<%=qdxq%>.gif"  style="width:24px;height:24px;">
	  </td>
      <td class="splittd" align=middle>
	    <%
			if qdnow=1 then Response.Write("<font color=""green"">今天已签到</font>") else Response.Write("<font color=""#FF0000"">今天未签到</font>")
		   %>
	
	   
	   </td>
    </tr>
	 <tr>
		<td colspan="5"  class="splittd"  align="left" style="background:#F5F5F5; line-height:25px;" > 
		<span style="color:#333; font-size:12px;font-weight:bold; padding-left:15px; color:#006699">我今天想说: <font color="#333"><%=Content%></font></span>
		</td>
	</tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>
  
		<%
		End Sub
		
		 Sub ShowContent_me()
			 Dim I,intotalmoney,outtotalmoney,Page_s,qdxq,RSkc,Content,adddate,qiandao_xqco,qiandao_dateend,qdnow
			 Do While Not rs.eof 
			%>
			<tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
			  
			  <td  class="splittd" align=middle><%=rs("adddate")%></td>
			  <td  class="splittd" align=middle ><img src="/images/emot/<%=rs("qdxq")%>.gif"  style="width:24px;height:24px;"></td>
			  <td  class="splittd" align=middle ><%=rs("Content")%></td>
			  <td  class="splittd" align=middle >
			 <% if KS.ChkClng(rs("Status"))=0 then Response.Write("<font color=""green"">已签到</font>") else Response.Write("<font color=""#FF0000"">未签到</font>")%>	  
			  </td>
			  </tr>
			<%
						
						I = I + 1
						RS.MoveNext
						If I >= MaxPerPage Then Exit Do
		
			 loop
			%>
		<%
		End Sub
	  
End Class
%> 

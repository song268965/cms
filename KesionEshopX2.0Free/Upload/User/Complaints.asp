﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls,KS
Set KS=New PublicCls
Set KSCls = New ComplaintsCls
If KS.IsNul(KS.C("UserName")) Then
 KSCls.Kesion1()
Else
 KSCls.Kesion()
End If
Set KSCls = Nothing

Class ComplaintsCls
        Private KS,KSR,KSUser
		Private Descript,OrderID,FileContent
		Private ComeUrl
		Private totalPut,currentpage,MaxPerPage
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
	   Public Sub loadMain()	
	   ComeUrl=Request.ServerVariables("HTTP_REFERER")
   
       If KS.S("Action")<>"View" Then
		Call KSUser.Head()
		%>
		<div class="tabs">	
			<ul>
	        <li<%if KS.S("status")="" then response.write " class='puton'"%>><a href="Complaints.asp">投诉/意见管理</a></li>
			</ul>
			
	    </div>
        <div class="writeblog"><img src="images/icon1.png" align="absmiddle"><a href="?Action=Add">我要投诉</a></div>
	 <%
	 End If
	 	Call KSUser.InnerLocation("投诉管理")
	 	KSUser.CheckPowerAndDie("s17")

	 
	 	Select Case KS.S("Action")
			  Case "Show" Call View()
			  case "del" call FeedBackDel()
			  case "Add" call Add()
			  case "DoSave" call Addsave()
		      Case Else  Call ComplaintsList()
		End Select	
       End Sub
		
       Public Sub Kesion1()
	    
		IF Cbool(KSUser.UserLoginChecked)=false And KS.ChkClng(KS.Setting(47))=1 Then
	   	 Dim TemplatePath:TemplatePath=KS.Setting(3) & KS.Setting(90) & "Common/Complaints.html"  '模板地址
		 FileContent = KSR.LoadTemplate(TemplatePath)    
		 FCls.RefreshType = "Complaints" '设置刷新类型，以便取得当前位置导航等
		 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		 FileContent=KSR.KSLabelReplaceAll(FileContent)
		 If KS.S("Action")="DoSave" Then
		   Call Addsave()
		 End If
		 KS.Die FileContent
	   Else
	     KS.Die "<script>alert('对不起，本站不允许匿名投诉，请先登录!');location.href='login';</script>"	 
	   End If

						  
	End Sub
	
	Sub ComplaintsList()
      %>
	   <table width="100%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
        <tr align="center" class="title">
			 <td height="28" align="center"><strong>编号</strong></td>
			 <td><strong>主题</strong></td>
			 <td align="center"><strong>对象</strong></td>
			 <td width="10%" align="center"><strong>投诉时间</strong></td>
			 <td width="10%" align="center"><strong>处理人</strong></td>
			 <td width="12%" align="center"><strong>处理时间</strong></td>
			 <td width="10%" align="center"><strong>状态</strong></td>
			 <td><strong>操作</strong></td>
         </tr>
		   <%
		   	MaxPerPage=10
			If KS.S("page") <> "" Then
					 CurrentPage = KS.ChkClng(KS.S("page"))
			Else
				  CurrentPage = 1
			End If
			   Dim Param:Param=" where UserName='" & KSUser.UserName & "'"
			  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select * From KS_FeedBack " & Param & " order By ID",conn,1,1
				If RS.EOF And RS.BOF Then
					Response.Write "<tr><td colspan='10'class='splittd' align='center' colspan=2 height=30 valign=top>您没有发表任意见或投诉!</td></tr>"
				Else
									totalPut = RS.RecordCount
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									End If
							        Call showComplaintsList(RS)
				End If
     %>                 
            </table>
						<div style="margin:10px 25px; line-height:24px; font-size:14px; color:#555;">
						<strong>操作说明</strong><br />投诉/意见管理放置的是您投诉及对本站的建议记录；<br>您可以删除未处理的记录。
					</div>
	  <%
	End Sub
	
	Sub showComplaintsList(RS)
	  Dim str,i
	  Do While Not RS.Eof
	      dim bh:bh=rs("id")
		  IF LEN(BH)=1 THEN 
			  BH="00"& bh
		  ElseIf LEN(BH)=2 Then
			  Bh="0" & bh
		  End If
		  bh="YJ" & year(rs("adddate")) & month(rs("adddate")) & bh
          response.write "<tr bgcolor=#ffffff>"
          Response.Write "<td height='30' class='splittd' align='center'>" & bh & "</td>"
          Response.Write "<td class='splittd' align='center'>" 
		  
		   response.write rs("title")
		  response.write "</td>"
          Response.Write "<td class='splittd' align='center'>" & rs("object") & "</td>"
          Response.Write "<td class='splittd' align='center'>" & formatdatetime(rs("adddate"),2) & "</td>"
		  
          Response.Write "<td class='splittd' align='center'>"
		  Dim AcceptTime,Delstr,strs
		  if rs("Accepted")="" or isnull(rs("accepted")) then
		   response.write "未处理"
		   AcceptTime="---"
		   Delstr="<a onclick=""return(confirm('确定删除吗?'))"" href='?action=del&id=" & rs("id") & "'>删除</a>"
		   strs="<font color=red>待受理</font>"
		  else
		   response.write rs("Accepted")
		   AcceptTime=RS("AcceptTime")
		   Delstr="<a href='#' disabled>删除</a>"
		   strs="<font color=green>已受理</font>"
		  end if
		  response.write "</td>"
          Response.Write "<td class='splittd' align='center'>" & AcceptTime & "</td>"
          Response.Write "<td class='splittd' align='center'>" & strs & "</td>"
          Response.Write "<td class='splittd' align='center'><a href='?action=Show&id=" & rs("id") & "'>查看详情</a>  " & delstr & "</td>"

           Response.Write "</tr>"
	   
	  	RS.MoveNext
		I = I + 1
		If I >= MaxPerPage Then Exit Do
	 Loop
	 response.write str
	 %>
		 <tr>
			 <td align="right" colspan="10" height="50">
				 <%=KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)%>
			  </td>
		 </tr>
							  
	 <%
		
	End Sub
	
	Sub Add()       
	   Dim ID,RS,RealName,Tel,Sex
	   ID=KS.ChkClng(KS.S("ID"))
	   Call KSUser.InnerLocation("我要投诉")
	   %>
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
				   if ($("#Verifycode").val()=='' || $("#Verifycode").val()=='验证码'){
							    alert('请输入验证码!');
								$('#Verifycode').focus();
								return false;
					 }
				  }
				 </script>
				 <form name="bmform" action="?action=DoSave" method="post">
				  <input type="hidden" name="TrainID" value="<%=ID%>">
                <table border="0" align="center" cellpadding="0" cellspacing="1" class="border">
				  <tr>
                      <td width="145" align="right" height="25"><strong>意见主题：</strong></td>
                      <td width="797" align="left" > 
					  <input type="text" name="Title" id="Title" class="textbox" size="30">
				      </td>
				  </tr>
				   <tr>
                      <td width="145" align="right" height="35"><strong>意见对象：</strong></td>
                      <td  align="left"  height="25"> <input type="text" name="Object" class="textbox" size="30"> </td>
                  </tr>
				   <tr>
                      <td width="145" align="right"  height="35"><strong>意见内容：</strong></td>
                      <td  align="left"  height="25"> 
					  <textarea name="content" class="textbox" id="content" style="width:450px;height:100px"></textarea>
				     </td>
                  </tr>
				  <tr>
                      <td width="145" align="right" height="35"><strong>期望解决方案：</strong></td>
                      <td  align="left"  height="25"> 
					  <textarea name="Hopesolution" class="textbox" style="width:450px;height:100px"></textarea>
				    </td>
                  </tr>
                     <tr>
                      <td width="145" align="right" height="35"><strong>验 证 码：</strong></td>
                      <td  align="left"  height="25"> 
					  <input name="Verifycode" id="Verifycode" type="text" class="textbox" style="width:60px" autocomplete="off" onBlur="if(this.value==''){ this.value='验证码'; }" onFocus="if (this.value=='验证码'){this.value='';}"  value="验证码" />
                   <span id="showVerify"><img style='margin-top:13px;height:28px;cursor:pointer' title='点击刷新' align='absmiddle' src='../plus/verifycode.asp' onClick='this.src="../plus/verifycode.asp?n="+ Math.random();'></span>
				    </td>
                  </tr>
                   
           </table>
                <br><div style="text-align:center">
				
				&nbsp;<input type="Submit" class="button" onClick="return(checkform())" value=" 立即投诉 ">
				
				</div>
	    </form>
	 <br><br><br><br>
	 <%'RS.Close:Set RS=Nothing
	End Sub
	
	Sub Addsave()
	    if ks.s("title")="" then
		 response.write "<script>alert('请输入主题!');history.back();</script>"
		 exit sub
		end if
	    if ks.s("content")="" then
		 response.write "<script>alert('请输入内容!');history.back();</script>"
		 exit sub
		end if
		dim Verifycode:Verifycode=KS.R(KS.S("Verifycode"))
		    IF lcase(Trim(Verifycode))<>lcase(Trim(Session("Verifycode"))) then 
			   response.write "<script>alert('验证码有误，请重新输入!');history.back();</script>"
		      exit sub
			End IF
		
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
		 If KS.IsNul(KS.C("UserName")) Then
		 response.write "<script>alert('我们已收到您的意义，感谢您的反馈！');location.href='../';</script>"
		 Else
		 response.write "<script>alert('你的投诉已提交，请耐心等待处理结果!');location.href='Complaints.asp';</script>"
		 End If
	End Sub
	
	Sub View()
	   Call KSUser.InnerLocation("查看投诉详情")
       Dim ID,RS
	   ID=KS.ChkClng(KS.S("ID"))
	   Set RS=Server.CreateOBject("ADODB.RECORDSET")
	   RS.Open "Select top 1 * from ks_feedback where username='" & KSUser.UserName & "' and id=" & ID,conn,1,1

	   IF RS.Eof Then
	     RS.CLOSE:Set RS=Nothing
		 Response.Write "<script>alert('出错了!');window.close();</script>"
	   End If
	%>
          <!DOCTYPE HTML>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
			<title>投诉</title>
			<link href="images/css.css" type="text/css" rel="stylesheet" />
			<script src="../ks_inc/common.js"></script>
			</head>
			<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">	 
			
			<table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
        <tr>
         <td align="center">
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr class="title">
                      <td height="30"><strong>查看投诉详情</strong></td>
                    </tr>
           </table>
                <table border="0" align="center" cellpadding="0" cellspacing="1" class="normaltext">
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">意见主题：</td>
                      <td width="797" class='splittd'> 
					  &nbsp;<%=RS("title")%>
					 
				      </td>
				  </tr>
				   <tr>
                      <td width="145" align="right" class='splittd' height="25">意见对象：</td>
                      <td class='splittd' height="25">&nbsp; <%=RS("object")%> </td>
                  </tr>
				   <tr>
                      <td width="145" align="right" class='splittd' height="25">意见内容：</td>
                      <td class='splittd' height="25">&nbsp; <%=RS("content")%> </td>
                  </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">希望处理结果：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("hopesolution")%></td>
                      
                    </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">处理人：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("accepted")%></td>
                      
                    </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">处理时间：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("accepttime")%></td>
                      
                    </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">处理结果：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("acceptresult")%></td>
                      
                    </tr>
                    
                   
           </table>
                <br><div style="text-align:center">
				<input type="button" class="button" value=" 返 回 " onClick="history.back();">
				&nbsp;
				
				</div>
                
		 
		 </td>
       </tr>

     </table>
	 <%RS.Close:Set RS=Nothing

	 End Sub
	
	Sub FeedBackDel()
	  Conn.Execute("Delete from ks_FeedBack where (Accepted='' or Accepted is null ) and username='" & KSUser.UserName &"' and id=" & KS.ChkClng(KS.S("ID")))
	  Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End Sub
	
End Class
%> 

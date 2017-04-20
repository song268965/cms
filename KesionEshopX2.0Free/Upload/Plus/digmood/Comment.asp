<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/ClubFunction.asp"-->
<!--#include file="../../plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls

If Not KS.GetAppStatus("digmood") THEN KS.Die ""

Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,InfoID,RS,CommentStr,UserIP,Total,TitleStr,TitleLinkStr,TotalPoint,N,DomainStr,Title,Verific
Dim totalPut, MaxPerPage,PageNum,SqlStr,PrintOut,CommentXML,PostId,PostTable,Tid,Fname,PostLoad
ChannelID=KS.Chkclng(KS.S("ChannelID"))


IF ChannelID=0 And KS.S("Action")<>"Support" And KS.S("Action")<>"QuoteSave" Then KS.Die ""
PrintOut=KS.S("PrintOut")

PostLoad=KS.ChkClng(KS.S("PostLoad"))
InfoID=KS.ChkClng(KS.S("InfoID"))
DomainStr=KS.GetDomain
MaxPerPage=KS.ChkClng(KS.S("maxperpage"))
Select Case KS.S("Action")
 Case "Show"  Call ShowComment()
 Case "Write"
  If KS.ChkClng(KS.C_S(ChannelID,12))=0 and channelid<>1000 Then Response.end()
  Call Ajax()
  Response.Write("document.write('" & GetWriteComment(ChannelID,InfoID) & "');")
 Case "WriteSave"  Call WriteSave()
 Case "Support"  
  If PrintOut="js" Then
   Response.Write "ShowSupportMessage('" & Support() & "');"
  Else
   Response.Write Support()
  End If
 Case "ShowQuote" Call ShowQuote()
 Case "QuoteSave" Call QuoteSave()
 Case Else  Call CommentMain()
 End Select
 Set KS=Nothing
 Set KSUser=Nothing
 
 '输出头部
Sub WriteHead()
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src='<%=KS.GetDomain%>KS_Inc/jquery.js'></script>
<script src='<%=KS.GetDomain%>KS_Inc/common.js'></script>
<style>
*{margin:0;padding:0;word-wrap:break-word;}
body{font:12px/1.75 "宋体", arial, sans-serif,'DejaVu Sans','Lucida Grande',Tahoma,'Hiragino Sans GB',STHeiti,SimSun,sans-serif;color:#444;}
html, body, h1, h2, h3, h4, ul, li, dl,input{ margin:0px;padding:0px;list-style-type:none }
a{color:#333;text-decoration:none;}
a:hover{text-decoration:underline;}
.btn{padding:2px;}
.textbox{BACKGROUND-COLOR: #ffffff;BORDER: #ccc 1px solid;COLOR: #999;HEIGHT: 22px;line-height:22pxborder-color: #666666 #666666 #666666 #666666; font-size: 9pt;FONT-FAMILY: verdana;}
</style>
<%
If EnabledSubDomain Then
 response.write "<script>document.domain=""" & RootDomain &""";</script>" &vbcrlf
end if
%>
</head>
<body>
<%
End Sub

'显示回复
Sub ShowQuote()
WriteHead
%>
<br/>
<div style='height:200px;text-align:center'>
<form name='rform' action='<%=KS.GetDomain%>plus/digmood/comment.asp?action=QuoteSave' method='post'>
<input type='hidden' name='channelid' value='<%=ChannelID%>'>
<input type='hidden' name='infoId' value='<%=infoid%>'>
<input type='hidden' name='quoteId' value='<%=KS.S("quoteId")%>'>
<input type='hidden' name='postId' value='<%=KS.S("postId")%>'>
<textarea name='quotecontent' class='textbox' style='overflow:auto;width:90%;height:130px'></textarea>
<br><label>
<input type='checkbox' value='1' name='Anonymous'> 匿名发表</label> 
<input type='submit' class='btn' value=' 发 表 '>
</form>
</div>
<%
End Sub
 
Sub Ajax()
 %>
function formToRequestString(form_obj)
{
    var query_string='';
    var and='';
    for (var i=0;i<form_obj.length;i++ )
    {
        e=form_obj[i];
        if (e.name) {
            if (e.type=='select-one') {
                element_value=e.options[e.selectedIndex].value;
            } else if (e.type=='select-multiple') {
                for (var n=0;n<e.length;n++) {
                    var op=e.options[n];
                    if (op.selected) {
                        query_string+=and+e.name+'='+escape(op.value);
                        and="&"
                    }
                }
                continue;
            } else if (e.type=='checkbox' || e.type=='radio') {
                if (e.checked==false) {   
                    continue;   
                }   
                element_value=e.value;
            } else if (typeof e.value != 'undefined') {
                element_value=e.value;
            } else {
                continue;
            }
            query_string+=and+e.name+'='+escape(element_value);
            and="&"
        }
    }
    return query_string;
}
function ajaxFormSubmit(form_obj){ 
    jQuery.getScript(form_obj.getAttributeNode("action").value+"&"+formToRequestString(form_obj),   
       function(){ 
	      cmtsuccess();
    });   
}
 <%
 End Sub
 
 Sub CommentMain
	Dim KSRCls,FileContent
	Set KSRCls = New Refresh
	FCls.RefreshType = "Comment" '设置刷新类型，以便取得当前位置导航等
     MaxPerPage=KS.ChkClng(Split(KS.C_S(ChannelID,46)&"||||||||||||","|")(23))
	if KS.C_S(ChannelID,15)="" then KS.Die "请先到模型设置里绑定评论页模板!"
	FileContent = KSRCls.LoadTemplate(KS.C_S(ChannelID,15))
	If Trim(FileContent) = "" Then FileContent = "模板不存在!"
	FileContent=Replace(FileContent,"{$GetShowComment}","<script src=""" & domainstr & "ks_inc/Comment.page.js"" language=""javascript""></script><script language=""javascript"" defer>var from3g=0;Page(1," & ChannelID & ",'" & InfoID & "','Show'," & MaxPerPage &",'"& domainstr & "');</script><div id=""c_" & InfoID & """></div><div id=""p_" & InfoID & """ align=""right""></div>")


    FileContent=Replace(FileContent,"{$GetWriteComment}","<script src=""?Action=Write&ChannelID=" & ChannelID& "&InfoID=" & InfoID & """></script>")
	
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,Conn,1,1
	If RS.Eof And RS.Bof Then
		RS.Close:Set RS=Nothing
		KS.Die "<script>alert('对不起，文章不存在！');window.close();</script>"
	Else
	            
						 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
						 Set KSRCls.Node=DocXml.DocumentElement.SelectSingleNode("row")
						 fcls.ItemTitle= KSRCls.Node.SelectSingleNode("@title").text 
						 
						 if KSRCls.Node.SelectSingleNode("@comment").text=0 then
						  KS.Die "<script>alert('对不起，不允许评论 ！');window.close();</script>"
						end if
						  KSRCls.ModelID=ChannelID
						  KSRCls.ItemID = KSRCls.Node.SelectSingleNode("@id").text 
						  KSRCls.Tid=KSRCls.Node.SelectSingleNode("@tid").text
						  KSRCls.Templates=""
				          KSRCls.Scan FileContent
		 		          FileContent = KSRCls.Templates
				        End If
						RS.Close
						Set RS=Nothing
	
	FileContent = KSRCls.ReplaceLableFlag(KSRCls.ReplaceAllLabel(FileContent))
	FileContent = KSRCls.ReplaceGeneralLabelContent(FileContent) '替换通用标签
	Set KSRCls = Nothing
   Response.Write(FileContent)
End Sub

Sub ShowComment()
	If Request.ServerVariables("HTTP_REFERER")<>"" Then 
	  If Instr(Lcase(Request.ServerVariables("HTTP_REFERER")),"comment.asp")<>0 Then MaxPerPage=20
	End If
	 CurrentPage = KS.ChkClng(KS.S("page"))
	 If CurrentPage<=0 Then CurrentPage=1

    If ChannelID=1000 Then
     SqlStr="Select top 1 ID,subject as Title,classid as tid,0 as fname,0 as postid,PostTable,CmtNum,AddDate From KS_GroupBuy Where ID=" & InfoID
	Else
     SqlStr="Select top 1 ID,Title,Tid,Fname,PostId,PostTable,CmtNum,AddDate From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID
	End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open SqlStr,Conn,1,1
	 If Not RS.Eof Then
	    Dim totalPut,PageNum
		Set KSRCls = New Refresh
		if MaxPerPage=0 then MaxPerPage=KS.ChkClng(Split(KS.C_S(ChannelID,46)&"||||","|")(22))
		
		 CommentStr= KSRCls.GetCommentList(CurrentPage,RS(5),KS.ChkClng(RS(4)),ChannelID,InfoID,RS("tid"),RS("title"),RS("fname"),RS("AddDate"),RS("cmtnum"),totalPut,MaxPerPage,PageNum)
		 KSRCls.ModelID = ChannelID
		 KSRCls.ItemID = InfoID
		 KSRCls.Templates = ""
		 KSRCls.Scan CommentStr
		 CommentStr = KSRCls.Templates

		Set KSRCls=Nothing
	 End If
	   
  Rs.Close:Set Rs=Nothing
  
  if PostLoad=1 Then '提交回复的，判断权限，决定是否显示
	 GetVerific
     if Verific=1 then  KS.Die "var json={message:""" & replace(replace(replace(CommentStr,vbcrlf,"\n"),"""","\"""),chr(10),"\n") &"""}"
     Exit Sub
  End If
  If KS.C_S(ChannelID,12)=0 and channelid<>1000 Then TotalPut=0
  KS.Die "var json={message:""" & replace(replace(replace(CommentStr,vbcrlf,"\n"),"""","\"""),chr(10),"\n") & "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|条||2""}"
End Sub

'状态
Sub GetVerific()
    if KS.ChkClng(KS.C_S(Channelid,12))=1 Or KS.ChkClng(KS.C_S(ChannelID,12))=3 then verific=0 else verific=1
	 If KS.ChkClng(KS.C_S(Channelid,12))=5 Then
	  If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then verific=0 else verific=1
	 End If
	 if channelid=1000 then
	  dim rsg:set rsg=conn.execute("select top 1 comment,postTable from ks_groupbuy where id=" & infoid)
	  if rsg.eof then
	    rsg.close:set rsg=nothing
	    exit sub
	  else
	    postTable=rsg("postTable")
	    if rsg("comment")=0 then
	    rsg.close:set rsg=nothing
	    exit sub
		elseif rsg("comment")=1 then
		 verific=0
		else
		 verific=1
		end if
	  end if
	  rsg.close:set rsg=nothing
	 end if
End Sub

 
'发表评论
Function GetWriteComment(ChannelID,InfoID)
%>
function cmtsuccess()
{   
	var loading_msg='\n\n\t请稍等，正在提交评论...';
	var C_Content=document.getElementById('C_Content');
    var isC_Content=C_Content.value;
    isC_Content=isC_Content.indexOf("请稍等，正在提交评论");
    if (isC_Content==-1){C_Content.value="\n\n\t请稍等,正在提交评论..";}
	if (json.message=='ok') {       
	     KesionJS.Alert('恭喜,你的评论已成功提交！');
		 if (typeof(loadDate)!="undefined")  loadDate(1,1);
		  leavePage();
	 }else{
	     alert(json.message);
		 C_Content.value=document.getElementById('sC_Content').value;
	 }
}
var OutTimes =11;
function leavePage()
{
	if (OutTimes==0){
	 document.getElementById('C_Content').disabled=false;
	 document.getElementById('SubmitComment').disabled=false;
	 document.getElementById('C_Content').value='文明上网，请对您的发言负责！'
	 <%If KS.C_S(ChannelID,13)="1" Then%>
	  document.form1.Verifycode.value='';
	 <%end if%>
	 <%If KS.C_S(ChannelID,14)<>0  Then%>
	 document.getElementById('cmax').value=<%=KS.C_S(ChannelID,14)%>;
	 <%end if%>
	 OutTimes =11;
	 return;
	 }
	else {
	    document.getElementById('C_Content').disabled=true;
		document.getElementById('SubmitComment').disabled=true;
		OutTimes -= 1;
		document.getElementById('C_Content').value ="\n\n\t评论已提交，等待 "+ OutTimes + " 秒钟后您可继续发表...";
		setTimeout("leavePage()", 1000);
		}
	}
function checklength(cobj){ 
	var cmax=<%=KS.C_S(ChannelID,14)%>;
	if (cobj.value.length>cmax) {
	cobj.value = cobj.value.substring(0,cmax);
	KesionJS.Alert("评论不能超过"+cmax+"个字符!");
	}
	else {
	document.getElementById('cmax').value = cmax-cobj.value.length;
	}
}

function checkcommentform(){
	var anounname=document.getElementById('AnounName');
	var C_Content=document.getElementById('C_Content');
	var sC_Content=document.getElementById('sC_Content');
	var anonymous=document.getElementById('Anonymous');
	var pass=document.getElementById('Pass');
   if (anounname.value==''){
        KesionJS.Alert('请填写用户名!',"$('#Anonymous').focus()");
        return false;
     }
	if (anonymous.checked==false && pass.value==''){
	   KesionJS.Alert('请输入密码或选择游客发表！','$("#Pass").focus()');
	   return false;
	}
	<%If KS.C_S(ChannelID,13)="1" Then%>
   if (document.form1.Verifycode.value==''){
	   KesionJS.Alert('请入验证码!','document.form1.Verifycode.focus();');
	   return false;
    }
	<%end if%>
   if (C_Content.value==''||C_Content.value=='文明上网，请对您的发言负责！'){
	   KesionJS.Alert('请填写评论内容!','$("#C_Content").focus();');
	   return false;
    }
	sC_Content.value=C_Content.value;
	
    ajaxFormSubmit(document.form1)

	 
	}
	function checkbindweibo(){
	  if ($("#transweibo")[0].checked){
	    jQuery.post("<%=DomainStr%>user/UserAjax.asp",{action:'CheckToken',checktype:"sinaweibo"},function(d){
		   if (d!='success'){
		     KesionJS.Alert('您没有绑定新浪费微博账号，或是授权失效！','$("#transweibo").attr("checked",false);');
			 
		   }else{
		     $("#transweibo").attr("checked",true);
		   }
		});
	  }
	}
<%
		 GetWriteComment = GetWriteComment & "<style>.comment_write_table,.comment_write_table textarea,.comment_write_table a{color:#666}.comment_write_table textarea,.comment_write_table .textbox{padding:2px;color:#999;border:1px solid #cccccc;}</style><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table""><form name=""form1"" action=""" & DomainStr &"plus/digmood/Comment.asp?Action=WriteSave"" method=""post""><input type=""hidden"" value=""" & ChannelID & """ name=""ChannelID""><input type=""hidden"" value=""" & InfoID & """ name=""InfoID"">"
		GetWriteComment = GetWriteComment & "<tr>"
		GetWriteComment = GetWriteComment & "  <td style=""padding:10px;"">"
		Dim PostNum,PostId
		PostId=KS.ChkClng(Request.QueryString("postId"))
		If PostId<>0 Then
		  Dim RSN:Set RSN=Conn.Execute("Select top 1 TotalReplay From KS_GuestBook Where ID=" & PostId &" and deltf=0")
		  If Not RSN.Eof Then
		   PostNum=RSN(0)
		  Else
		   PostNum=0 : PostId=0
		  End If
		  RSN.Close: Set RSN=Nothing
		Else
		 If ChannelID<>1000 Then
		  Dim RS:Set RS=Conn.Execute("select TOP 1 cmtnum  From " & KS.C_S(ChannelID,2) &" Where ID=" & InfoID)
		  If Not RS.Eof Then
		    PostNum=KS.ChkClng(RS(0))
		  End If
		  RS.Close:Set RS=Nothing
		 Else
		  PostNum=Conn.Execute("Select count(1) From KS_Comment Where ProjectID=0 and verific=1 and ChannelID=" & ChannelID & " And InfoID=" & InfoID)(0)
		 End If
		End If
		GetWriteComment = GetWriteComment & "  <div style=""font-size:14px;height:30px;line-height:30px;text-align:left;""><strong>已有 <span style=""color:brown;font-weight:bold"" class=""cmtnum"">" & PostNum & "</span> 条跟帖</strong>"
		If ChannelID<>1000 and request("from3g")="" Then
			If PostId<>0 Then
			GetWriteComment = GetWriteComment & "<a href=""" & KS.GetClubShowUrl(PostId) & """ style=""color:brown"" target=""_blank"">(点击查看)</a></div>"
			Else
			GetWriteComment = GetWriteComment & "<a href=""" & DomainStr &"plus/digmood/Comment.asp?ChannelID=" & ChannelID & "&InfoID=" & InfoID & """ style=""color:brown"">(点击查看)</a></div>"
			End If
		Else
			GetWriteComment = GetWriteComment & "</div>"
		End If
		
		If KS.C_S(ChannelID,14)<>0  Then
		GetWriteComment = GetWriteComment & "<textarea onkeydown=""checklength(this);"" onkeyup=""checklength(this);"" name=""C_Content"" rows=""6"" id=""C_Content"" onfocus=""if(this.value==\'文明上网，请对您的发言负责！\'){this.value=\'\'}"" wrap=""PHYSICAL"" onblur=""if(this.value==\'\'){this.value=\'文明上网，请对您的发言负责！\'}"" style=""overflow:auto;font-size:14px;width:100%"">文明上网，请对您的发言负责！</textarea>"
		Else
		GetWriteComment = GetWriteComment & "<textarea style=""font-size:14px;padding:5px;width:98%;height:90px;overflow:auto;"" onfocus=""if(this.value==\'文明上网，请对您的发言负责！\'){this.value=\'\'}"" wrap=""PHYSICAL"" onblur=""if(this.value==\'\'){this.value=\'文明上网，请对您的发言负责！\'}"" name=""C_Content"" rows=""4"" id=""C_Content"">文明上网，请对您的发言负责！</textarea>"
		End If
		
		GetWriteComment = GetWriteComment & "</td></tr>"
		GetWriteComment = GetWriteComment & "  <tr><td nowrap>"
		GetWriteComment = GetWriteComment & "  <div style=""margin-left:10px;text-align:left;"">"
		If KSUser.UserName="" Then
		GetWriteComment = GetWriteComment & " 用户名：<input onfocus=""if(this.value==\'匿名\'){this.value=\'\';}"" class=""textbox"" maxlength=15 name=""AnounName"" type=""text"" id=""AnounName"" value=""匿名"" style=""width:70px""/> "
			if request("from3g")="1" then 
			 GetWriteComment = GetWriteComment & "<a href=""" & DomainStr & KS.Wsetting(4) &"/reg.asp""><u>注册</u></a>"
			else
			 GetWriteComment = GetWriteComment & "<a href=""" & DomainStr & "user/reg/""><u>注册</u></a>"
			end if
		Else
		GetWriteComment = GetWriteComment & "   <span style=""display:none"">用户名：<input class=""textbox"" maxlength=15 name=""AnounName"" type=""text"" id=""AnounName"" value=""" & KSUser.username & """ style=""width:70px""/></span>欢迎您，" & KSUser.UserName &"! <a href="""
		if request("from3g")="1" then 
		GetWriteComment = GetWriteComment & KS.Setting(3) & KS.Wsetting(4) &"/user.asp"
		Else
		GetWriteComment = GetWriteComment & DomainStr & "user/"
		End If
		GetWriteComment = GetWriteComment & """>[会员中心]</a> <a onclick=KesionJS.Confirm(\'确认退出吗？\',\'location.href=""" & DomainStr & "user/UserLogout.asp""\',null);  href=""javascript:;"">[退出]</a>"
		End If
		Dim Style,Check
		If KS.C_S(ChannelID,12)="1" or KS.C_S(ChannelID,12)="2" Then
		 If KS.IsNul(KS.C("UserName"))  Then style="": else Style=" style=""display:none"""
		 checked=""
		Else
		 Style=" style=""display:none""":checked=" checked"
		End If
		if request("from3g")="1" then GetWriteComment = GetWriteComment &"<br/>"
		GetWriteComment = GetWriteComment & "<span id=""pp""" & style & "> 密码：<input class=""textbox"" name=""Pass"" size=""8"" type=""password"" id=""Pass"" value=""" & KSUser.PassWord & """ ></span>"

		If KS.C_S(ChannelID,13)="1" and channelid<>1000 Then
		if request("from3g")="1" then GetWriteComment = GetWriteComment & "<br/>"
		GetWriteComment = GetWriteComment & "&nbsp;认证码：<script>writeVerifyCode(""" & KS.GetDomain &""",0)</script>"
		End IF
		
		If KS.C("UserName")="" Then
		GetWriteComment = GetWriteComment & "<span id=""nm"">"
		Else
		GetWriteComment = GetWriteComment & "<span id=""nm"" style=""display:none"">"
		End If

		If KS.C_S(Channelid,12)="1" Or KS.C_S(Channelid,12)="2" Then
		GetWriteComment = GetWriteComment & "<span style=""display:none"">"
		Else
		GetWriteComment = GetWriteComment & "<span>"
		End iF
		GetWriteComment = GetWriteComment & "<label><input onclick=""if(this.checked==true){document.getElementById(\'Pass\').disabled=true;document.getElementById(\'pp\').style.display=\'none\';}else{if(document.getElementById(\'AnounName\').value==\'匿名\'){document.getElementById(\'AnounName\').value=\'\';}document.getElementById(\'Pass\').disabled=false;document.getElementById(\'pp\').style.display=\'\';}"" type=""checkbox""" & checked & " value=""1"" name=""Anonymous"" id=""Anonymous"">匿名</label></span>"
		GetWriteComment = GetWriteComment & "</span>"

		If KS.C_S(ChannelID,14)<>0  Then
		GetWriteComment = GetWriteComment & "&nbsp;字数：<input disabled class=""textbox"" type=""text"" id=""cmax"" size=""3"" name=""cmax"" value=""" & KS.C_S(ChannelID,14) & """/>"
		End If
		if request("from3g")="1" then GetWriteComment = GetWriteComment & "<br/>"
		GetWriteComment = GetWriteComment & "<input type=""hidden"" name=""sC_Content"" id=""sC_Content""><input type=""submit"" id=""SubmitComment"" name=""SubmitComment"" value=""确认发表"" style=""padding:2px"" onclick=""checkcommentform();return false""/></div>"
		If ChannelID<>1000 Then GetWriteComment = GetWriteComment & "&nbsp;<label><input onclick=""return(checkbindweibo());"" type=""checkbox"" value=""1"" name=""transweibo"" id=""transweibo"">同时转播到我的微博</label>"
		
		GetWriteComment = GetWriteComment & "</td></tr></form></table>"
		End Function  
		

		
'保存发表
Sub WriteSave()	
		Dim UserName,C_Content,Anonymous,point,VerifyCode,Pass,Flag,ComeUrl,GroupID,LoginTF,PostId,PostTable,AddDate
		Flag=KS.S("Flag")
		ComeUrl=Request.ServerVariables("HTTP_REFERER"):If ComeUrl="" Then ComeUrl=KS.GetDomain
		LoginTF=Cbool(KSUser.UserLoginChecked)
		If ChannelID=1000 Then '团购
		 If Conn.Execute("Select top 1 id From KS_GroupBuy Where Comment>=1 and ID=" & InfoID).Eof Then
		  KS.Die "var json={message:""对不起,本团购不允许评论!""}"
		 End If
		ElseIf KS.ChkClng(KS.C_S(Channelid,12))=0 Then 
		  KS.Die "var json={message:""对不起,本信息不允许评论!""}"
		End If	  
		
		AnounName=KS.R(KS.S("AnounName"))
		If LoginTF=false And Len(AnounName)>20 Or Len(AnounName)<2 Then
		   KS.Die "var json={message:""用户名不符合规范，长度限制在2-20之间!""}"
		End If
		Pass=KS.R(KS.G("Pass"))
		C_Content=KS.S("C_Content")
		VerifyCode=KS.S("VerifyCode")
		
		Anonymous=KS.ChkClng(KS.S("Anonymous"))
		point=KS.ChkClng(KS.S("point"))
		If ChannelID<>1000 AND KS.C_S(ChannelID,13)="1" and lcase(Trim(Request("Verifycode")))<>lcase(Trim(Session("Verifycode"))) Then
		 KS.Die "var json={message:""验证码有误，请重新输入!""}"
		End IF
		  
		IF Anonymous=0 Then
		  if LoginTF=false then
		     	if Pass="" Then 
				  KS.Die "var json={message:""请填写登录密码或选择游客发表!""}"
				End if
             Pass=Md5(Pass,16)
		     Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
			 UserRS.Open "Select top 1 UserID,UserName,PassWord,Locked,Score,LastLoginIP,LastLoginTime,LoginTimes,RndPassword,GroupID From KS_User Where UserName='" &AnounName & "' And PassWord='" & Pass & "'",Conn,1,3
			 If UserRS.Eof And UserRS.BOf Then
				  UserRS.Close:Set UserRS=Nothing
				  KS.Die "var json={message:""你输入的用户名或密码有误，请重新输入!""}"
			 ElseIf UserRS("Locked")=1 Then
			     KS.Die "var json={message:""您的账号已被管理员锁定，请与管理员联系!""}"
			 Else
			            GroupID=UserRS("GroupID")
			            '登录成功，更新用户相应的数据
						Dim RndPassword:RndPassword=KS.R(KS.MakeRandomChar(20))
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
						UserRS("LastLoginIP") = KS.GetIP
                        UserRS("LastLoginTime") = Now()
                        UserRS("LoginTimes") = UserRS("LoginTimes") + 1
						UserRS("RndPassWord")=RndPassWord
                        UserRS.Update
						If EnabledSubDomain Then
							 Response.Cookies(KS.SiteSn).domain=RootDomain					
						Else
                             Response.Cookies(KS.SiteSn).path = "/"
						End If
						Response.Cookies(KS.SiteSn)("UserName") = AnounName
						Response.Cookies(KS.SiteSn)("Password") = Pass
						Response.Cookies(KS.SiteSN)("RndPassword")= RndPassword
						Response.Cookies(KS.SiteSn)("UserID") = UserRS("UserID")
						Response.Cookies(KS.SiteSn)("GroupID") = GroupID
			end if
			UserRS.Close : Set UserRS=Nothing
		  Else
		     groupid=KSUser.GroupID
		  end if
		Else
		    Dim RSG:Set RSG=Conn.Execute("select top 1 GroupID from KS_User Where UserName='" & AnounName & "'")
			If Not RSG.Eof Then
			  groupID=rsg(0)
			End If
			RSG.Close : Set RSG=Nothing
		End IF
		
		if KS.ChkClng(KS.C_S(Channelid,12))=1 Or KS.ChkClng(KS.C_S(ChannelID,12))=2 then
		  if KS.C("UserName")="" or KS.C("PassWord")=""  then
		     KS.Die "var json={message:""对不起，系统设置不允许游客发表!""}"
		  End If
		End If

		IF InfoID="" Then 
		   KS.Die "var json={message:""参数传递有误!""}"
		End if
		if AnounName="" Then
		  KS.Die "var json={message:""请填写你的昵称!""}"
		End if
		if C_Content="" Then 
		  KS.Die "var json={message:""请填写评论内容!""}"
		End if
		If Len(C_Content)>KS.ChkClng(KS.C_S(ChannelID,14)) and KS.ChkClng(KS.C_S(ChannelID,14))<>0 Then
		    KS.Die "var json={message:""评论内容必须在" &KS.C_S(ChannelID,14) & "个字符以内!""}"
		End if
		
		if Not KS.IsNul(KS.C("username")) then Anonymous=0

		Set RS=Server.CreateObject("ADODB.RECORDSET")
		IF ChannelID=1000 Then '团购
		 RS.Open "Select top 1 subject as Title,0,0,classid  as Tid,id as Fname,adddate From KS_GroupBuy Where id=" & InfoID,Conn,1,1
		Else
		 RS.Open "Select top 1 Title,PostId,PostTable,Tid,Fname,adddate From " & KS.C_S(ChannelID,2) &" Where id=" & InfoID,Conn,1,1
		End If
		If RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
			   KS.Die "var json={message:""内容不存在!""}"
		End IF
		PostId=KS.ChkClng(RS(1)) : PostTable=RS(2):Title=RS("Title"):Tid=rs(3):Fname=RS(4):AddDate=RS(5)
		RS.Close
		Set RS=Nothing
		If KS.IsNul(PostTable) Then PostTable="KS_Comment"
		Call DoWriteSave(0,PostTable,PostID,InfoID,AnounName,C_Content,"",KSUser,Anonymous,AddDate)
		KS.Die "var json={message:""ok""}"
End Sub

'保存发表评论
Sub DoWriteSave(IsQuote,PostTable,PostID,InfoID,AnounName,C_Content,QuoteContent,KSUser,Anonymous,AddDate)
     Dim BoardID,O_LastPost,N_LastPost,UserID,BSetting,LoginTF,RS
	 C_Content=KS.LoseHtml(C_Content)
	 AnounName=KS.LoseHtml(AnounName)
	 LoginTF=Cbool(KSUser.UserLoginChecked)
     if KS.ChkClng(KS.C_S(Channelid,12))=1 Or KS.ChkClng(KS.C_S(ChannelID,12))=3 then verific=0 else verific=1
	 If KS.ChkClng(KS.C_S(Channelid,12))=5 Then
	  If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then verific=0 else verific=1
	 End If
	 if channelid=1000 then
	  dim rsg:set rsg=conn.execute("select top 1 comment,postTable from ks_groupbuy where id=" & infoid)
	  if rsg.eof then
	    rsg.close:set rsg=nothing
	    exit sub
	  else
	    postTable=rsg("postTable")
	    if rsg("comment")=0 then
	    rsg.close:set rsg=nothing
	    exit sub
		elseif rsg("comment")=1 then
		 verific=0
		else
		 verific=1
		end if
	  end if
	  rsg.close:set rsg=nothing
	 end if
	 If KS.IsNul(postTable) Then postTable="KS_Comment"
	 
     If PostId<>0 Then '绑定论坛帖子
	     Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 BoardID,PostTable From KS_GuestBook Where ID=" & PostId,conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.CLose:Set RS=Nothing
			KS.Die "var json={message:""帖子内容不存在!""}"
		  End If
		  PostTable=RS("PostTable"):BoardID=RS("BoardID")
		  RS.Close
		  If IsQuote=1 Then  '引用
		   RS.Open "Select top 1 * From " & PostTable  & " Where ID=" & KS.ChkClng(KS.S("quoteId")),Conn,1,1
		   If RS.Eof And RS.Bof Then
		    RS.CLose:Set RS=Nothing
			KS.Die "var json={message:""引用的帖子内容不存在""}"
		   End If
		   C_Content="[quote]以下是引用 " & RS("UserName") & " 在" & RS("ReplayTime") & " 的发言：[br]"& RS("Content") &"[/quote]" & C_Content
		   RS.Close
		  End If
		  
		  UserID=KS.ChkClng(KSUser.GetUserInfo("UserID"))
		  If BoardID<>0 Then
			 KS.LoadClubBoard()
			 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
		  End If
		  BSetting=BSetting & "$$$0$0$0$0$0$0$1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
		  BSetting=Split(BSetting,"$")
		  Call InsertReply(PostTable,AnounName,UserID,PostId,C_Content,0,0,PostId,verific,SQLNowString) '写入论坛回复表
		  Conn.Execute("Update KS_GuestBook Set LastReplayTime=" & SqlNowString &",LastReplayUser='" & AnounName &"',LastReplayUserID=" & UserID & ",TotalReplay=TotalReplay+1 where id=" & PostId)
		  N_LastPost=PostId & "$" & now & "$" & Replace(Title,"$","") &"$" & AnounName & "$" &UserID&"$$"
           If KS.ChkClng(BSetting(4))>0 and LoginTF=true Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(4)),"系统","在论坛回复主题[" & Title & "]所得!",0,0)
		   End If
		  
		   '更新版面数据
			If BoardID<>0 Then
			  KS.LoadClubBoard()
			  O_LastPost=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text
			  Call UpdateBoardPostNum(0,BoardID,Verific,O_LastPost,N_LastPost)
			End If
			UpdateTodayPostNum '更新今日发帖数等
			Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set CmtNum=" & LFCls.GetCmtNum(PostTable,ChannelID,KS.ChkClng(KS.S("InfoID"))) & " Where ID=" & KS.ChkClng(KS.S("InfoID")))
		Else
		     Dim CommentPerTime:CommentPerTime=KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(4))
		     If CommentPerTime<>0 Then 
			  If not Conn.Execute("Select top 1 * From " & PostTable &" Where ProjectID=0 and InfoID=" & InfoID & " and UserIp='" & KS.GetIP & "' and datediff(" & DataPart_H & ",AddDate," & SqlNowString &")<" & CommentPerTime).eof then
			      KS.Die "var json={message:""对不起，同一篇文档" &CommentPerTime & "小时内只能评论一次""}"
			  end if
			 End If
			   
			   GroupID=KSUser.GetUserInfo("groupid")
			 
			   Conn.Execute("Insert Into " & PostTable &"(ChannelID,InfoID,UserName,Anonymous,Content,QuoteContent,UserIP,Point,Score,OScore,Verific,AddDate,ProjectID) values(" & ChannelID & "," & InfoID & ",'" & AnounName & "'," & Anonymous & ",'" & Replace(C_Content,"'","''") & "','" & Replace(QuoteContent,"'","''") & "','" & KS.GetIP & "',0,0,0," & Verific & "," & SQLNowString& ",0)")
			  If KS.ChkClng(groupid)<>0 and Verific=1 Then
				  If KS.ChkClng(KS.U_S(GroupID,6))>0 Then
					 Call  KS.ScoreInOrOut(KS.C("UserName"),1,KS.ChkClng(KS.U_S(GroupID,6)),"系统","参与文档[<a href=""" & KS.GetItemUrl(channelid,Tid,infoid,Fname,AddDate) & """ target=""_blank"">" & Title & "</a>]的评论!",1002,""&ChannelID&""&InfoID)
				  End If
			  End If
			  If ChannelID=1000 Then
			  Conn.Execute("Update KS_GroupBuy Set CmtNum=" & LFCls.GetCmtNum(PostTable,ChannelID,InfoID) & " Where ID=" & InfoID) '更新评论数
			  Else
			  Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set CmtNum=" & LFCls.GetCmtNum(PostTable,ChannelID,InfoID) & " Where ID=" & InfoID) '更新评论数
			  End If
	  End If
	  
	  
	   	'转发到微博
		if ks.s("transweibo")="1" and KS.ChkClng(KS.C_S(ChannelID,6))<100 then
		 dim rst:set rst=conn.execute("select top 1 id,title,intro,fname,tid,photourl,adddate from " & KS.C_S(ChannelID,2) &" Where id=" & InfoID)
		 if not rst.eof then
		  dim commentcontent:commentcontent=rst("title") & KS.GetItemUrl(channelid,rst("Tid"),infoid,rst("Fname"),rst("adddate")) 
		  call ksuser.add_sina_weibo(commentcontent,rst("photourl"))
		 end if
		 rst.close
		 set rst=nothing
		end if
		
		'生成内容页
		If ChannelID<>1000 Then
			If KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2 Then
						 Dim KSRObj:Set KSRObj=New Refresh
						 Set RS=Server.CreateObject("ADODB.RECORDSET")
						 RS.Open "select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,Conn,1,1
						 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
						 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
						  KSRObj.ModelID=ChannelID
						  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
						  Call KSRObj.RefreshContent()
						  Set KSRobj=Nothing
						RS.Close
						Set RS=Nothing
		     End If
			 If KS.ChkClng(KS.M_C(ChannelID,28))=1 And IsBusiness Then
			            Fcls.CallFrom3g="true" 
			            Set KSRObj=New Refresh
						 Set RS=Server.CreateObject("ADODB.RECORDSET")
						 RS.Open "select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,Conn,1,1
						 Set DocXML=KS.RsToXml(RS,"row","root")
						 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
						  KSRObj.ModelID=ChannelID
						  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
						  Call KSRObj.RefreshContent()
						  Set KSRobj=Nothing
						RS.Close
						Set RS=Nothing
			 End If
	  End If
		
End Sub

Sub QuoteSave()
 Dim quoteId:quoteId=KS.ChkClng(KS.S("quoteId"))
 Dim Content:Content=KS.S("QuoteContent")
 Content=KS.LoseHtml(Content)
 WriteHead
 Dim QuoteArray,AnounName,QuoteContent,Anonymous,UserName,LoginTF,PostTable
 PostID=KS.ChkClng(KS.S("PostID"))
 If quoteId=0  Or InfoID=0 Then Response.Write "<script>KesionJS.Alert('参数传递出错!','history.back()');</script>":Exit Sub
 If Content="" Then Response.Write "<script>KesionJS.Alert('回复内容必须输入!','history.back()');</script>":Exit Sub
 If Len(Content)>KS.ChkClng(KS.C_S(ChannelID,14)) and KS.ChkClng(KS.C_S(ChannelID,14))<>0 Then
	 KS.Die "<script>KesionJS.Alert('评论内容必须在" &KS.C_S(ChannelID,14) & "个字符以内!','history.back()');</script>"
 End if
 Anonymous=KS.ChkClng(KS.S("Anonymous"))
 LoginTF=Cbool(KSUser.UserLoginChecked)
 IF LoginTF=false and (KS.ChkClng(KS.C_S(Channelid,12))=1 or KS.ChkClng(KS.C_S(Channelid,12))=2) Then
  Response.Write "<script>KesionJS.Alert('对不起,本站只允许注册会员发表!','history.back()');</script>":Exit Sub
 End If
 
 If Anonymous=1 Then
  AnounName="匿名"
 Elseif Anonymous=0 and LoginTF=false then
  Response.Write "<script>KesionJS.Alert('对不起,请先登录!','history.back()');</script>":Exit Sub
 Else
   AnounName=KSUser.UserName
 End If
 If LoginTF=True Then UserName=KSUser.UserName Else UserName="匿名"
 
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 
 Dim AddDate
IF ChannelID<>1000 Then 
	RS.Open "Select top 1 PostTable,AddDate From " & KS.C_S(ChannelID,2) &" Where id=" & InfoID,Conn,1,1
Else
	RS.Open "Select top 1 PostTable,AddDate From KS_GroupBuy Where id=" & InfoID,Conn,1,1
End If			 
      If RS.Eof And RS.Bof Then
				   RS.Close:Set RS=Nothing
				   If Flag="NotAjax" Then KS.Die "<script>alert('内容不存在!','history.back()');history.back();</script>" Else KS.Die("内容不存在!")
	  Else
			  PostTable=RS(0)
			  AddDate=RS(1)
	 End If
	 RS.Close

If KS.IsNul(PostTable) Then PostTable="KS_Comment"
If PostId=0 Then
	  RS.Open "Select top 1 channelid,infoid,username,Anonymous,adddate,content,quotecontent from " & PostTable &" where ProjectID=0 and id=" & quoteid,conn,1,1
	  if RS.Eof Then
		  RS.Close:Set RS=Nothing
		  Response.Write "<script>KesionJS.Alert('参数传递出错!','history.back()');</script>":Exit Sub
	  End If
	  QuoteArray = RS.GetRows(-1)
	  RS.Close : Set RS=Nothing
	 Dim Qstr:Qstr="[dt]引用 " 
	 If QuoteArray(3,0)=1 Then
	  Qstr=Qstr & "匿名"
	 Else
	  Qstr=Qstr & "会员:" & QuoteArray(2,0)
	 End If 
	 Qstr=Qstr & " 发表于" & QuoteArray(4,0) & "的评论内容[/dt][dd]" & QuoteArray(5,0) & "[/dd]"
	 If QuoteArray(6,0)<>"" Then
	 QuoteContent="[quote]" & QuoteArray(6,0) & Qstr & "[/quote]"
	 Else
	 QuoteContent="[quote]" & Qstr & "[/quote]"
	 End If
	 InfoID=QuoteArray(1,0)
 Else
     InfoID=PostId
 End If
 Call DoWriteSave(1,PostTable,PostID,InfoID,AnounName,Content,QuoteContent,KSUser,Anonymous,AddDate)
 
 KS.Die "<script>KesionJS.Alert('恭喜,您的评论发表成功!','try{top.loadDate(1,1);top.box.close();}catch(e){top.location.replace(document.referrer);}');</script>"
End Sub

Function Support()
	 Dim ID,OpType,PostId,RS
	 ID=KS.ChkClng(KS.S("ID")) : OpType=KS.ChkClng(KS.S("Type")) : PostId=KS.ChkClng(KS.S("PostID"))
	 IF Cbool(Request.Cookies(Cstr(ID))("SupportCommentID" & Cstr(ID)))<>true Then
	    If PostID<>0 Then
		   Set RS=Conn.Execute("Select top 1 PostTable From KS_GuestBook Where ID=" & PostId)
		   If Not RS.Eof Then
	        if OpType=1 Then
		       Conn.Execute("Update " & RS("PostTable") & " Set Support=Support+1 Where ID=" & ID)
		    else
		       Conn.Execute("Update " & RS("PostTable") & " Set Opposition=Opposition+1 Where ID=" & ID)
			end if
		   End If
		   RS.Close:Set RS=Nothing
		Else
		    Dim PostTable
			If ChannelID<>1000 Then
				SET RS=Conn.Execute("Select Top 1 PostTable From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)
				If Not RS.Eof Then
				  PostTable=RS(0)
				End If
				RS.Close :Set RS=Nothing
			End If
			If KS.IsNul(PostTable) Then PostTable="KS_Comment"
	        if OpType=1 Then
		       Conn.Execute("Update " & PostTable &" Set score=score+1 Where ID=" & ID)
		    else
	           Conn.Execute("Update " & PostTable &" Set OScore=OScore+1 Where ID=" & ID)
			end if
		End If
		If EnabledSubDomain Then
			Response.Cookies(KS.SiteSn).domain=RootDomain					
		Else
            Response.Cookies(KS.SiteSn).path = "/"
		End If
		Response.Cookies(Cstr(ID))("SupportCommentID" & Cstr(ID))=true
	Else
	 Support="var json= {message:""你已投过票了！""}" : Exit Function
	End If
	if OpType=1 Then Support="var json= {message:""good""}" Else Support="var json= {message:""bad""}"
End Function
%>
 

<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KS,KSUser
Set KS=New PublicCls
Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,RS,CommentStr,Total,UserIP,tag,str
tag=ks.s("tag")
select case KS.S("Action")
  case "CommentSave"
    call CommentSave()
  case else
    Response.Write("document.write('" & GetWriteComment(KS.ChkClng(KS.S("UserID")),KS.S("ID"),KS.S("Title"),KS.S("UserName")) & "');")
end select


'*********************************************************************************************************
'函数名：GetWriteComment
'作  用：取得发表评论信息
'参  数：ID -信息ID
'*********************************************************************************************************
Function GetWriteComment(UserID,ID,Title,UserName)
		%>
		function insertface(Val)
	      {  
		  if (Val!=''){ $('.content').focus();
		  var str = document.selection.createRange();
		  str.text = Val; }
          }
		  function success()
			{
				var loading_msg='\n\n\t请稍等，正在提交评论...';
				var content=document.getElementById('Content');
				
				if (loader.readyState==1)
					{
						content.value=loading_msg;
					}
				if (loader.readyState==4)
					{   var s=loader.responseText;
						if (s=='ok')
						 {
						 alert('恭喜,你的评论已成功提交！');
						  location.reload();
						 }
						else
						 {alert(s);
						 }
					}
			}
		   function checkform()
		   { 
		    if (document.getElementById('AnounName').value==''){
			 alert('请输入昵称!');
			 document.getElementById('AnounName').focus();
			 return false;
			}
		    if (document.getElementById('Content').value==''){
			 alert('请输入评论内容!');
			 document.getElementById('Content').focus();
			 return false;
			}
		   ksblog.ajaxFormSubmit(document.form1,'success')
           }
		   
		function ShowLogin()
		{ 
		 $.dialog({title:'会员登录',content:'url:<%=KS.Setting(3)%>user/userlogin.asp?Action=Poplogin',width:397,height:184});
		}
		<%
		If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  str="<div style=""margin:20px""><strong>温馨提示：</strong>只有会员才可以发表评论,如果是会员请先<a href=""javascript:ShowLogin()"">登录</a>,不是会员请点此<a href=""../user/reg/"" target=""_blank"">注册</a>。</div>"
		Else
		 str = "<table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 str=str & "<form name=""form1"" action=""WriteComment.asp?action=CommentSave"" method=""post"">"
		 str=str & "<input type=""hidden"" value=""" & tag & """ name=""tag""><input type=""hidden"" value=""" & ks.chkclng(request("xcid")) & """ name=""xcid""><input type=""hidden"" value=""" & UserID & """ name=""UserID""><input type=""hidden"" value=""" & UserName & """ name=""UserName""><input type=""hidden"" value=""" & ID & """ name=""ID"">"
		 str=str & "<tr><td colspan=""2"" height=""30"" class=""comment_write_title""><strong>发表评论:</strong>"
		 Dim HomePage
		 If KS.C("UserName")<>"" Then
		  HomePage=KS.Setting(2) & "/space/?" & KS.C("UserID")
		 Else
		  HomePage="http://"
		 End If
		str=str & "<br/>昵称："
		str=str & "   <input name=""AnounName"" maxlength=""100"" type=""text"" id=""AnounName"" value=""" & KS.C("username") & """"
		If KS.C("UserName")<>"" Then str=str & " readonly"
		str=str & " style=""height:25px;line-height:25px;color:#999;width:35%;border:1px solid #ccc;background:#FBFBFB;""/><br/>主页："
		str=str & "    <input name=""HomePage"" maxlength=""150"" value=""" & HomePage & """ type=""text"" id=""HomePage"" style=""height:25px;line-height:25px;color:#999;width:55%;border:1px solid #ccc;background:#FBFBFB;"" /><br/>标题："
		str=str & "    <input name=""Title"" maxlength=""150"" value=""Re:" & Title & """ type=""text"" id=""Title"" style=""height:25px;line-height:25px;color:#999;width:55%;border:1px solid #ccc;background:#FBFBFB;"" /><input type=""hidden"" value=""" & Title & """ name=""OriTitle""></td>"
		str=str & "  </tr>"
		
		
		str=str & "  <tr>"
		str=str & "    <td width=""70%""><textarea name=""Content"" class=""content"" rows=""6"" id=""Content""  style=""width:99%;color:#999;border:1px solid #ccc;background:#FBFBFB;overflow:auto""></textarea></td>"
		
		 Dim str1:str1="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		 Dim strArr:strArr=Split(str1,"|")
		  str=str & "<td width=""140"" style=""padding-left:20px;word-break:break-all"">"
		 For K=0 to 19
		   str=str & "<img style=""cursor:pointer"" title=""" & strarr(k) & """ onclick=""insertface(\'[e" & k &"]\')""  src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">&nbsp;"
		   If (K+1) mod 5=0 Then str=str & "<br />"
		 Next

		str=str & "</td>"
		str=str & "  </tr>"
		str=str & "  <tr>"
		
		str=str & "    <td colspan=""2"" style=""text-align:left""><input type=""button"" onclick=""return(checkform())"" name=""SubmitComment"" id=""SubmitComment""class=""btn"" value=""提交评论""/>"
		
		str=str & "    </td>"
		str=str & "  </tr>"
		str=str & "  </form>"
		str=str & "</table>"
		End If
		GetWriteComment=str
		End Function  
  
        Sub CommentSave()
	    	Dim ID,UserName,HomePage,Content,Anonymous,Title,SinaWeiboID
			ID=KS.ChkClng(KS.S("ID"))
			AnounName=KS.LoseHTML(KS.S("AnounName"))
			HomePage=KS.CheckXSS(KS.S("HomePage"))
			Content=KS.CheckXSS(KS.LoseHtml(KS.S("Content")))
			Title=KS.CheckXSS(KS.LoseHtml(KS.S("Title")))
			If Title="" Then Title="回复本文主题"
			IF ID="0" Then 
			 Response.Write("参数传递有误!")
			 Response.End
			End if
			if AnounName="" Then 
			 Response.Write("请填写你的昵称!'")
			 Response.End
			End if
			
			
			if Content="" Then 
			 Response.Write("请填写评论内容!")
			 Response.End
			End if
			
			if tag<>"photo" then
				Set RS=Conn.Execute("Select top 1 UserName,SinaWeiboID From KS_BlogInfo Where ID=" & ID)
				If RS.Eof And RS.Bof Then
				  RS.Close:Set RS=Nothing
				 KS.Die("参数传递有误!")
				End If
				UserName=RS(0)
				SinaWeiboID=RS(1)
				RS.Close
				
				If Not KS.IsNul(SinaWeiboID) Then   '同步到新浪微博
				  if KSUser.UserLoginChecked=true then 
					if KSUser.GetUserInfo("sinatoken")<>"" and KSUser.GetUserInfo("sinaid")<>"" Then '判断有绑定就同步
					  call ksuser.add_sina_comment(SinaWeiboID,content&",来自："&KS.GetDomain &"space/?" & KSUser.GetUserInfo("userid")&"/log/" & ID)
					end if
				  end if
				End If
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select top 1 * From KS_BlogComment",Conn,1,3
				RS.AddNew
				 RS("LogID")=ID
				 RS("AnounName")=AnounName
				 RS("Title")=Title
				 RS("UserName")=UserName
				 RS("HomePage")=HomePage
				 RS("Content")=Content
				 RS("UserIP")=KS.GetIP
				 RS("AddDate")=Now
				RS.UpDate
				 RS.Close:Set RS=Nothing
				 Conn.Execute("Update KS_BlogInfo Set TotalPut=TotalPut+1 Where ID=" & ID)
				
			else   '照片评论
				Set RS=Conn.Execute("Select top 1 UserName From KS_PhotoZP Where ID=" & ID)
				If RS.Eof And RS.Bof Then
				  RS.Close:Set RS=Nothing
				  KS.Die("参数传递有误!")
				End If
				UserName=RS(0)
				RS.Close
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select top 1 * From KS_PhotoComment",Conn,1,3
				RS.AddNew
				 RS("photoID")=ID
				 RS("AnounName")=AnounName
				 RS("Title")=Title
				 RS("UserName")=UserName
				 RS("HomePage")=HomePage
				 RS("Content")=Content
				 RS("UserIP")=KS.GetIP
				 RS("XCID")=ks.chkclng(request("xcid"))
				 RS("AddDate")=Now
				RS.UpDate
				 RS.Close:Set RS=Nothing
			end if
			
			
			  Call CloseConn()
             If KS.S("From")="1" Then
			  Response.Write "<script>alert('你的评论发表成功!');top.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
			 Else
			   response.write "ok"
			 End If
			 Set KS=Nothing
		End Sub
  
Set KS=Nothing
Set KSUser=Nothing
%>

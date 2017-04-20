<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New User_Blog
KSCls.Kesion()
Set KSCls = Nothing

Class User_Blog
        Private KS,KSUser
		Private totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather
		Private TypeID,Title,Tags,UserName,Face,Content,Status,PicUrl,Action,I,ClassID,password
		Private Sub Class_Initialize()
		  MaxPerPage =15
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/SpaceFunction.asp"-->
		<%
       Public Sub loadMain()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		If KS.SSetting(0)=0 Then
		 KS.Die ("<script>$.dialog.tips('对不起，本站关闭个人空间功能！',1,'error.gif',function(){location.href='index.asp';});</script>")
		 Exit Sub
		End If
		Call KSUser.SpaceHead()
		 Action=KS.S("Action")
		 If Action="" Or Action="Add" Then
		 KSUser.CheckPowerAndDie("s02")
		 End If
		
		%>
		<div class="tabs">	
			<ul>
			 <%If Action="BindDomain" Then%>
			 <li class='puton'><a href="#">域名绑定设置</a></li>
			 <%End If%>
			 <%IF Action="BlogEdit" Or Action="Template" Or action="Banner" Then%>
			 <li<%If Action="BlogEdit" then response.write " class='puton'"%>><a href="?action=BlogEdit">空间设置</a></li>
			 <%If KSUser.GetUserInfo("UserType")=1 Then%>
			 <li style="display:none" <%If Action="Banner" then response.write " class='puton'"%>><a href="?action=Banner">Banner设置</a></li>
			 <%End IF%>
			 <li<%If Action="Template" then response.write " class='puton'"%>><%if request("flag")=3 then%><a href="?action=Template"><%else%><a href='SetSpaceTemplate.asp' target='_blank'><%end if%>模板设置</a></li>
			 <%End If%>
				 
			 <%
			 If Action="Add" Or Action="Edit" Then
			 %>
			 <li><a href="?">博文管理</a></li>
			 <li class='puton'><%If Action="Add" Then Response.Write "写博文" Else Response.Write "编辑博文" End If%></li>
			 <%
			 Elseif Action="" then%>
				<li<%If KS.ChkClng(KS.S("listtype"))="0" then response.write " class='puton'"%>><a href="?listtype=0">我的博文(<span class="red"><%=conn.execute("select count(id) from KS_BlogInfo where istalk<>1 and UserName='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("listtype")="1" then response.write " class='puton'"%>><a href="?listtype=1">好友博文(<span class="red"><%=conn.execute("select count(1) from KS_BlogInfo l inner join ks_friend f on l.username=f.friend where f.accepted=1 and l.UserName='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			 <%end if%>

			</ul>
	  </div>
					<%if ks.s("action")="" or ks.s("action")="Comment" then%>
					 <div class='writeblog'><img src="../images/user/log/101.gif" align="absmiddle"><a href="User_Blog.asp?Action=Add">写博文</a> 
					 &nbsp;&nbsp;<img src="../images/user/log/100.gif" align="absmiddle"><a href="User_message.asp?Action=Comment">博文评论</a>
					 </div>
					<%end if%>


		<%
		If KS.S("Action")="ApplySave" Then
		   Call ApplyBlogSave()
		ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
		    Response.Write "<script>$.dialog.tips('您还没有开通空间,请填写申请开通！',1,'alert.gif',function(){});</script>"
		    Call ApplyBlog()
		ElseIf Action<>"BlogEdit" and Conn.Execute("Select top 1 status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
		    Response.Write "<script>$.dialog.tips('对不起，你的空间还没有通过审核或被锁定！',1,'error.gif',function(){location.href='space.asp';});</script>"
			response.end
		Else
			Select Case KS.S("Action")
			 Case "Del"	  Call ArticleDel()
			 Case "Add","Edit"  Call ArticleAdd()
			 Case "DoSave"	  Call DoSave()
			 Case "Template"  Call Template()
			 Case "SaveMySkin"	  Call SaveMySkin()
			 Case "BlogEdit"	  Call ApplyBlog()
			 Case "Banner" SetBanner()
			 Case "BindDomain" BindDomain
			 Case "BindDomainSave" BindDomainSave
			 Case Else Call BlogList()
			End Select
		End If
		 Response.Write "</div>"
	   End Sub
	   function ConHalfWidth(content)
    dim strbig,tstrbig,strsma,tstrsma,s1,s2,i
    strbig="ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ＇"
    tstrbig="ABCDEFGHIJKLMNOPQRSTUVWXYZ'"
    strsma="ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
    tstrsma="abcdefghijklmnopqrstuvwxyz"
	s1="０１２３４５６７８９"
	s2="0123456789"
    dim b(26),tb(26),s(25),ts(25),sa1(10),sa2(10)
    for i=1 to 26
        b(i-1)=mid(strbig,i,1)
        tb(i-1)=mid(tstrbig,i,1)
        content=replace(content,b(i-1),tb(i-1))
        s(i-1)=mid(strsma,i,1)
        ts(i-1)=mid(tstrsma,i,1)
        content=replace(content,s(i-1),ts(i-1))
    next
	for i=1 to 10
	    sa1(i-1)=mid(s1,i,1)
        sa2(i-1)=mid(s2,i,1)
        content=replace(content,sa1(i-1),sa2(i-1))
	next
	s1="，+,'-%^&?()<>[]{}/\;:！@#￥%……&*（）——+【】{}、《》~！＠＃￥％……＆（×）＋“：｝｛"
	for i=1 to len(s1)+1
	   content=replace(content,mid(s1,i,1),"")
	next
    ConHalfWidth=content
end function


function ReplaceChar(str)
   str=replace(replace(replace(replace(replace(lcase(str&""),"-",""),"&quot;",""),"&gt;",""),"&lt;",""),"&amp;","")
   str=KS.LoseHTML(replace(str," ",""))
   str=ConHalfWidth(str)
   ReplaceChar=Str
End Function
	   '域名设置
	   Sub BindDomain()
	     Call KSUser.InnerLocation("空间域名设置")

	   	Dim RS,Domain,BlogName
		Set RS=Conn.Execute("Select top 1  * From KS_Blog Where UserName='" & KSUser.UserName &"'")
		If Not RS.EOF Then
		 BlogName=RS("BlogName")
		 domain=RS("domain")
        End If
		RS.Close :Set RS=Nothing
		%>
		<div style="padding:15px;margin:36px;margin-left:2px;width:680px;border:1px solid #C1DEFB;background:#E8EFF9;">您空间的初始地址为：<a href="<%=KS.GetDomain%>space/?<%=KS.C("UserID")%>" target="_blank"><%=KS.GetDomain%>space/?<%=KS.C("UserID")%></a>
		<%if domain<>"" then
		       KS.Echo "<br/><strong>当前绑定的域名为：</strong>"
		          if instr(domain,".")=0 then
					KS.Echo "<a href='http://" & domain & "." & KS.SSetting(16) &"' target='_blank' style='color:#ff6600'>http://" & domain & "." & KS.SSetting(16) &"</a>"
				else
				    KS.Echo "<a href='http://" & domain &"' target='_blank' style='color:#ff6600'>http://" & domain &"</a>"
				end if
		end if
		%>
		</div>
		<%
		
	   If KS.SSetting(14)<>"0" Then%>
	     
	     <table>
		   <form name="myform" action="User_Blog.asp" method="post">
		   <input type="hidden" name="action" value="BindDomainSave" />
            <tr class="tdbg">
              <td class="clefttitle"><strong>绑定域名：</strong></td>
              <td><label><input type="radio" <%if instr(domain,".")=0 then response.write " checked"%> name="domaintype" onclick="$('#domain0').show();$('#domain1').hide();" value="0" />二级域名</label>
			      <label><input type="radio" <%if instr(domain,".")<>0 then response.write " checked"%> name="domaintype" onclick="$('#domain0').hide();$('#domain1').show();" value="1" />顶级独立域名</label>
				  <div id='domain0'<%if instr(domain,".")<>0 then response.write " style='display:none'"%>>
                  &nbsp;<input class="textbox" name="domain" type="text" id="domain" style="width:50px; " value="<%=domain%>" maxlength="100" /><b>.<%response.write KS.SSetting(16)%></b> <span class="msgtips">如果不想绑定可以留空</span>
				  </div>
				  <div id='domain1' <%if instr(domain,".")=0 or KS.IsNul(domain) then response.write " style='display:none'"%>>
				    &nbsp;<input class="textbox" name="mydomain" type="text" id="mydomain" style="width:150px; " value="<%=domain%>" maxlength="100" /> <span class="msgtips">如www.kesion.com 需要将您的域名解释到本站服务器IP上。</span>
				  </div> 
				 </td>
            </tr>
			<tr class="tdbg">
			  <td></td>
			  <td style="height:40px"><button id="btn" type="submit" class="pn"><strong>保存设置</strong></button>
			</tr>
		  </form>
		</table>
			<%End If
	   End Sub
	   Sub BindDomainSave()
		 Dim domaintype:domaintype=KS.ChkClng(KS.S("domaintype"))
         Dim Domain:Domain=KS.DelSql(KS.S("Domain"))
         If DomainType=1 Then
		   Domain=KS.DelSql(KS.S("mydomain"))
		 End If
		 If domain<>"" Then
			 if lcase(domain)="www" or lcase(domain)="space" or lcase(domain)="bbs" or lcase(domain)="news" then KS.die "<script>$.dialog.tips('您输入的二级域名为系统保留关键字,请重新输入!',1,'error.gif',function(){history.back();});</script>" 
			 if domain<>"" then
			  if not conn.execute("select top 1 username from ks_Blog where username<>'" & ksuser.username & "' and [domain]='" & domain  &"'").eof then
			  Response.Write "<script>$.dialog.tips('对不起，您注册的域名已被其它用户使用!',1,'error.gif',function(){history.back();});</script>":exit sub
			  end if
			 end if
		 End If 
		 Conn.Execute("Update KS_Blog Set [Domain]='" & Domain & "' Where UserName='" & KSUser.UserName &"'")
		  Conn.Execute("Update KS_EnterpriseNews Set [Domain]='" & Domain & "' Where UserName='" & KSUser.UserName &"'")
		  ks.die "<script>$.dialog.tips('恭喜，空间域名绑定成功!',1,'success.gif',function(){location.href='User_Blog.asp?action=BindDomain';});</script>"
	   End Sub
	   
	    '申请开通空间
	   Sub ApplyBlog()
	    Dim BlogName,domain,ClassID,Descript,ContentLen,ListBlogNum,ListReplayNum,ListGuestNum,OpStr,TipStr,TemplateID,Announce,ListLogNum,Logo
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1  * From KS_Blog Where UserName='" & KSUser.UserName &"'",conn,1,1
		If Not RS.EOF Then
		 Call KSUser.InnerLocation("修改空间参数")
		 BlogName=RS("BlogName")
		 Logo=RS("Logo")
		 domain=RS("domain")
		 ClassID=RS("ClassID")
		 Descript=RS("Descript")
		 Announce=RS("Announce")
		 ContentLen=RS("ContentLen")
		 ListBlogNum=RS("ListBlogNum")
		 ListLogNum=RS("ListLogNum")
		 ListReplayNum=RS("ListReplayNum")
		 ListGuestNum=RS("ListGuestNum")
		 OpStr="OK了，确定修改"
		  if rs("status")<>"1" then
		 %>
		 <h2 style="color:#993333">您的空间已申请，请耐心等待我们的审核，您还可以继续修改完善以下资料！</h2>
		 <%
		  end if
		Else
		 Call KSUser.InnerLocation("申请开通空间")
		 If KSUser.GetUserInfo("UserType")=1 Then
		   Dim RSS:Set RSS=Conn.Execute("Select top 1 CompanyName From KS_Enterprise Where UserName='" &KSUser.UserName&"'")
		   If Not RSS.Eof Then
  		     BlogName=RSS(0)
		   Else
  		     BlogName=KSUser.UserName & "的企业空间"
		   End If
		   RSS.Close : Set RSS=Nothing
		 else
		 BlogName=KSUser.UserName & "的个人空间"
		 end if
		 domain=KSUser.UserName
		 ClassID="0"
		 ContentLen=500
		 ListBlogNum=10
		 ListLogNum=10
		 ListReplayNum=10
		 ListGuestNum=10
		 Announce="没有公告!"
		 Logo="http://www.kesion.com/Images/logo.png"
		 OpStr="OK了，立即申请":TipStr="申 请 开 通 空 间"
		End if
		If KS.IsNul(Logo) Then Logo="http://www.kesion.com/Images/logo.png"
		RS.Close:Set RS=Nothing
	    %>
		<script type="text/javascript">
		 function CheckForm(){
		  if (document.myform.BlogName.value==''){
		   $.dialog.alert('请输入站点名称!',function(){document.myform.BlogName.focus();});
		   return false;
		  }
		  if (document.myform.ClassID.value=='0'){$.dialog.alert('请选择站点类型!',function(){document.myform.ClassID.focus();});return false; }
		  return true;
		 }
		</script>
		<table class="border" border="0" align="center" cellpadding="3" cellspacing="1">
          <form  action="User_Blog.asp?Action=ApplySave" method="post" name="myform" id="myform" onSubmit="return CheckForm();" enctype="multipart/form-data">

            <tr class="tdbg">
              <td class="clefttitle">空间名称：</td>
              <td> <input class="textbox" name="BlogName" type="text" id="BlogName" style="width:250px; " value="<%=BlogName%>" maxlength="100" /> <span class="msgtips">空间站点的名称。如我的家园，我的博客等</span></td>
            </tr>
			
            <tr class="tdbg">
              <td class="clefttitle">Logo地址：</td>
              <td><input type="file" class="textbox" name="photourl" size="40">
                <img src="<%=logo%>" onerror="this.src='../images/logo.png'" width="88" height="35" style="vertical-align:middle"><br>
		  　    <span class="msgtips">只支持jpg、gif、png，小于100k，默认尺寸为88*31</span></td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">空间分类：</td>
              <td><select class="select" size='1' name='ClassID' style="width:250px">
                    <option value="0">-请选择类别-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If ClassID=RS("ClassID") Then
								  Response.Write "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                  </select> <span class="msgtips">空间站点分类，以便游客查找</span></td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">站点描述：</td>
              <td><textarea class="textbox" name="Descript" id="Descript" style="width:80%;height:60px" cols=50 rows=6><%=Descript%></textarea><br/><span class="msgtips">关于您的空间站点介绍</span> </td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">空间公告：</td>
              <td><textarea class="textbox" name="Announce" id="Announce" style="width:80%;height:80px" cols=50 rows=6><%=Announce%></textarea><br/><span class="msgtips">发布您的最新活动公告，让更多用户了解您。</span></td>
            </tr>
            <tr class="tdbg" style="display:none">
              <td class="clefttitle">新 鲜 事：</td>
              <td>列表页每页显示<input class="textbox" name="ContentLen" type="text" id="ContentLen" style="text-align:center;width:50px; " value="<%=ContentLen%>" /> 条  <span class="msgtips">指空间新鲜事列表页里，每页显示新鲜事条数。</span>    </td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">显示博文：</td>
              <td>首页显示博文<input class="textbox" name="ListBlogNum" type="text" id="ListBlogNum" style="text-align:center;width:50px; " value="<%=ListBlogNum%>" />篇 <span class="msgtips">空间首页显示博文条数。</span>             </td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">显示回复：</td>
              <td>首页显示回复<input class="textbox" name="ListReplayNum" type="text" id="ListReplayNum" style="text-align:center;width:50px; " value="<%=ListReplayNum%>" />条  <span class="msgtips">空间首页显示最新回复条数。</span>              </td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">博文列表：</td>
              <td>列表页显示博文<input class="textbox" name="ListLogNum" type="text" id="ListLogNum" style="text-align:center;width:50px; " value="<%=ListLogNum%>" />篇  <span class="msgtips">空间首页显示最新博文篇数。 </span>             </td>
            </tr>
            <tr class="tdbg">
              <td  class="clefttitle">显示留言：</td>
              <td>首页显示留言<input class="textbox" name="ListGuestNum" type="text" id="ListGuestNum" style="text-align:center;width:50px; " value="<%=ListGuestNum%>" />条    <span class="msgtips">空间首页显示最新留言条数。</span>        </td>
            </tr>

            <tr class="tdbg">
			  <td></td>
              <td height="30">
			    <button type="submit" class="pn"><strong><%=OpStr%></strong></button>
                </td>
            </tr>
          </form>
</table>
		<%
	   End Sub
	   
	   '保存个人空间申请
	   Sub ApplyBlogSave()
            Dim fobj:Set FObj = New UpFileClass
		    FObj.GetData
            Dim MaxFileSize:MaxFileSize = 100   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.GetUserInfo("UserID"))
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"logo")
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):response.End()
			End Select
         Dim Flag
	     Dim BlogName:BlogName=KS.LoseHtml(KS.DelSql(Fobj.Form("BlogName")))
		 Dim ClassID:ClassID=KS.ChkClng(Fobj.Form("ClassID"))
		 Dim Descript:Descript=KS.LoseHtml(KS.DelSql(Fobj.Form("Descript")))
		 Dim Announce:Announce=KS.LoseHtml(KS.DelSql(Fobj.Form("Announce")))
		 Dim ContentLen:ContentLen=KS.ChkClng(Fobj.Form("ContentLen"))
		 Dim ListBlogNum:ListBlogNum=KS.ChkClng(Fobj.Form("ListBlogNum"))
		 Dim ListLogNum:ListLogNum=KS.ChkClng(Fobj.Form("ListLogNum"))
		 Dim ListReplayNum:ListReplayNum=KS.ChkClng(Fobj.Form("ListReplayNum"))
		 Dim ListGuestNum:ListGuestNum=KS.ChkClng(Fobj.Form("ListGuestNum"))
		 If BlogName="" Then Response.Write "<script>alert('请输入站点名称!');history.back();</script>":exit sub
		 If ClassID=0 Then Response.Write "<script>alert('请选择站点类型!');history.back();</script>":exit sub
		 If KSUser.GetUserInfo("UserType")=1 Then Flag=4 Else Flag=2
		'防发帖机
            dim kk,sarr
            sarr=split(KS.WordFilter,",")
            for kk=0 to ubound(sarr)
               if instr(KS.R(ReplaceChar(BlogName&Announce&Descript)),sarr(kk))<>0 then 
                  ks.die "<script>alert('含有非法关键词,请不要非法提交恶意信息！');history.back();</script>"
               end if
            next
		 
	     Dim AddFlag:AddFlag=false
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Blog Where UserName='" & KSUser.UserName & "'",conn,1,3
		 If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("UserID")=KSUser.GetUserInfo("userid")
		    RS("AddDate")=now
			RS("TemplateID")=KS.ChkClng(KS.U_S(KSUser.GroupID,29))
			  if KS.ChkClng(KS.SSetting(2))=2 then
			  RS("Status")=0
			  else
			  RS("Status")=1
			  end if
			  AddFlag=true
		 End If
		    If ReturnValue<>"" Then RS("Logo")=ReturnValue
		    RS("UserName")=KSUser.UserName
		    RS("BlogName")=BlogName
			RS("ClassID")=ClassID
			RS("Descript")=Descript
			RS("Announce")=Announce
			RS("ContentLen")=ContentLen
			RS("ListLogNum")=ListLogNum
			RS("ListBlogNum")=ListBlogNum
			RS("ListReplayNum")=ListReplayNum
			RS("ListGuestNum")=ListGuestNum
		  RS.Update
		  RS.MoveLast
		  Dim Status:Status=RS("Status")
		  If Not KS.IsNul(RS("Logo")) or Not KS.IsNul(RS("Banner")) Then
		  Call KS.FileAssociation(1025,rs("BlogID"),RS("Logo")&RS("Banner"),1)
		  End If
		  
		 RS.Close:Set RS=Nothing
		 Set Fobj=Nothing
		 Conn.Execute("Update KS_Blog Set UserId=" & KS.ChkClng(KSUser.GetUserInfo("userid")) & " Where UserName='" & KSUser.UserName & "'")
		 
		 
		 dim mytipsstr
		 if status<>"1" then
		   mytipsstr="<br/>请耐心等待我们的审核开通"
		 end if
		   mytipsstr=mytipsstr & "<br/><a style=""color:red"" href=""../space/?" & ksuser.getuserinfo("userid") & """ target=""_blank"">点此访问你的空间主页</a>"
		 
		 If AddFlag Then
		     Call KSUser.InnerLocation("空间申请成功")
			 Response.Write "<script>$.dialog.tips('申请开通空间站点成功,5秒后跳到门户首页!" & mytipsstr &"',5,'success.gif',function(){top.location.href='space.asp';});</script>"
		 Else
		     Call KSUser.InnerLocation("空间修改成功")
		     Response.Write "<script>$.dialog.tips('空间站点修改成功,5秒后跳到门户首页!" & mytipsstr &"',5,'success.gif',function(){top.location.href='space.asp';});</script>"
		 End If
	   End Sub
	   
	   Sub SetBanner()
		Call KSUser.InnerLocation("设置空间Banner")
	   Dim banner
	   
	   If KS.S("Act")="Save" Then
	      Dim fobj:Set FObj = New UpFileClass
			 on error resume next
			 FObj.GetData
			 if err.number<>0 then
			  call KS.AlertHistory("对不起,文件超出允许上传的大小!",-1)
			  response.end
			 end if
            Dim MaxFileSize:MaxFileSize = 600   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.GetUserInfo("UserID"))
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"banner")
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):response.End()
			End Select
			If ReturnValue<>"" Then
			 dim rsu:set rsu=server.createobject("adodb.recordset")
			 rsu.open "select top 1 banner,blogid,logo from ks_blog where username='" & KSUser.UserName & "'",conn,1,3
			 if not rsu.eof then
			   dim obanner,nbanner,k,nstr
			   obanner=split(rsu(0),"|")
			   nbanner=split(returnvalue,"|")
			   for k=0 to ubound(nbanner)
			     if k=0 then
				   if trim(nbanner(0))<>"" then nstr=nbanner(k) else nstr=obanner(k)
				 else
				   if nbanner(k)<>"" then 
				    nstr=nstr & "|" & nbanner(k)
				   else 
				     if ubound(obanner)>=k then
					  nstr=nstr& "|"&obanner(k)
					 else
					  nstr=nstr &"|"
					 end if
				   end if
				 end if
			   next
			    If Not KS.IsNul(rsu("Logo")) or Not KS.IsNul(nstr) Then
					Call KS.FileAssociation(1025,rsu("BlogID"),rsu("logo") & nstr,1)
				End If

			 end if
			 rsu.close
			 set rsu=nothing
            Conn.Execute("Update KS_Blog Set Banner='" & nstr & "' Where UserName='" & KSUser.UserName & "'")
			End If
			Response.Write "<script>alert('恭喜,banner上传成功!');</script>"
	   End If
		on error resume next
	   	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Blog Where UserName='" & KSUser.UserName &"'",conn,1,1
		If Not RS.EOF Then
		 if Not KS.IsNul(RS("Banner")) Then
		 Banner=Split(RS("Banner"),"|")
		 End If
	    End If
		RS.Close:Set RS=Nothing
		dim b1,b2,b3
		 b1=banner(0)
	   if ubound(banner)>=1 then b2=banner(1)
	   if ubound(banner)>=2 then b3=banner(2)
	    if b1="" or isnull(b1) then b1="../images/ad.jpg"
	    if b2="" or isnull(b2) then b2="../images/ad.jpg"
	    if b3="" or isnull(b3) then b3="../images/ad.jpg"
      %>
	    <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
          <form  action="?Action=Banner&act=Save" method="post" name="myform" id="myform" enctype="multipart/form-data">

            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>Banner1预览：</strong><br>
              </div></td>
              <td align="center">　
                <img src="<%=b1%>" width="600" height="100"><br>
              只支持jpg、gif、png，小于200k，图片的大小请根据自己选择模板下的标注设计</td>
            </tr>
			<tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>地址：</strong><br>
              </div></td>
              <td><input type="file" name="photourl1" size="60"></td>
			</tr>
			<tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>Banner2预览：</strong><br>
              </div></td>
              <td align="center">　
                <img src="<%=b2%>" width="600" height="100"><br>
              只支持jpg、gif、png，小于200k，图片的大小请根据自己选择模板下的标注设计</td>
            </tr>
			<tr class="tdbg">
			  <td  height="25" align="center"><div align="left"><strong>地址：</strong><br>
              </div></td>
			  <td><input type="file" name="photourl2" size="60">
			  </td>
			</tr>
			<tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>Banner3预览：</strong><br>
              </div></td>
              <td align="center">　
                <img src="<%=b3%>" width="600" height="100"><br>
              只支持jpg、gif、png，小于200k，图片的大小请根据自己选择模板下的标注设计</td>
            </tr>
			<tr class="tdbg">
			<td  height="25" align="center"><div align="left"><strong>地址：</strong><br>
              </div></td>
			  <td><input type="file" name="photourl3" size="60">
               </td>
            </tr>
            <tr class="tdbg">
              <td height="30" align="center" colspan=2>
                <input type="submit" name="Submit3"  class="button" value="保存设置" />
                          </td>
            </tr>
			</form>
		 </table>
	   <%
	   End Sub
	   
	   
	   '设置模板
	   Sub Template()
	    Dim Flag:Flag=KS.ChkClng(KS.S("Flag"))
		If Flag=0 Then 
		 If KSUser.GetUserInfo("UserType")=1 Then Flag=4 Else Flag=2
		End If
		
		if flag=2 or flag=4 then
	    Call KSUser.InnerLocation("设置空间模板")
		else
	    Call KSUser.InnerLocation("设置圈子模板")
		end if
		    MaxPerPage=8
			
		%>
			    <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" class="border">
                    <tr class="title">
                      <td height="22" colspan=4>
					  <%if KSUser.GetUserInfo("UserType")=1 Then%>
					  <a href="?Action=Template&Flag=4">设置空间模板</a>
					  <%Else%>
					  <a href="?Action=Template&Flag=2">设置空间模板</a>
					  <%end if%> | <a href="?Action=Template&Flag=3">设置圈子模板</a>
                    </tr>
                   <%
						Set RS=Server.CreateObject("AdodB.Recordset")
							RS.open "select * from ks_blogtemplate where TemplateAuthor='" & KSUser.username & "' or (usertag=0 and flag=" & Flag &") order by usertag desc,id desc",conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' height=30 valign=top>没有可用模板!</td></tr>"
								 Else
									totalPut = RS.RecordCount
								     If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									 End If
										Call ShowTemplate
				               End If
     %>                     
				</table>

		<%
		
	   End Sub
	   
	   Sub ShowTemplate()
	   %>
	   <style type="text/css">
	   	.t .onmouseover { background: #fffff0; }
		.t .onmouseout {}
		.t ul {float:left;margin:6px;padding:5px;width:152px!important;width:165px;height:300px;overflow:hidden;background: #fcfcfc;}
		.t ul li {list-style-type:none;line-height:1.5;margin:0;padding:0;}
		.t ul li.l1 img {width:150px;height:190px;}
		.t ul li.l1 a {display:block;margin:auto;padding:1px;width:156px;height:196px;text-align:left;}
		.t ul li.l2 {margin: 5px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
		.t ul li.l3 {margin: 5px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
		.t ul li.l4 {margin:10px 0 0 0;}
	   </style>
	   <%
	     dim i,k
	     do while not rs.eof
		   response.write "<tr>"
		   for i=1 to 4
		    response.write "<td class=""t"" width=""25%"">"
			 dim pic:pic=rs("templatepic")
			 if pic="" or isnull(pic) then pic="../images/nopic.gif"
			%>
			<ul class="onmouseout">
				<li class="l1"><a href='../space/showtemplate.asp?templateid=<%=rs("id")%>' target=_blank>
<img src="<%=pic%>" title="点击预览" width="200" height="122" border="0" />
</a></li>
				<li class="l2">名称：<strong><%=rs("templatename")%></strong></li>
				<li class="l3">
				<%if rs("templateauthor")=KSUser.UserName then%>
				<!--<a href="?action=UpTemplate&ID=<%=RS("ID")%>"><font color=red>修改模板</font></a> | <a href="?action=DelTemplate&ID=<%=rs("id")%>" onClick="return(confirm('删除模板不可恢复，确定吗？'))"><font color=red>删除模板</font></a>-->
				<%else%>
				作者：<%=rs("templateauthor")%>
				<%end if%>
				
				</li>
				<%if rs("flag")=3 then
				 if Not KS.IsNul(rs("groupid")) And KS.FoundInArr(rs("groupid"),KSUser.GroupID,",")=false And KSUser.GroupID<>1 Then
				   response.write "<li class=""l4""><font color=red>此模板Vip专用</font></li>"
				 else
				 %>
					<li class="l4">圈子：
					<select name='teamid<%=rs("id")%>' id='teamid<%=rs("id")%>' style='width:60px'>
					 <%dim rst:set rst=server.createobject("adodb.recordset")
					 rst.open "select * from ks_team where username='" & KSUser.UserName & "'",conn,1,1
					 if rst.eof then
					  response.write "<option value='0'>没有建圈子</option>"
					 else
					 do while not rst.eof
					  response.write "<option value='" & rst("id") & "'>" & rst("teamname") &"</option>"
					  rst.movenext
					 loop
					 end if
					 rst.close:set rst=nothing
					 %>
					</select>
					<input type="submit" value="应用" onClick="if($('#teamid<%=rs("id")%>').val()==0){alert('请选择圈子!');return false} else{window.location='?flag=3&teamid='+$('#teamid<%=rs("id")%>').val()+'&action=SaveMySkin&id=<%=RS("ID")%>'}" />
					</li>
				<%
				 end if
				else%>
				<li class="l4">
				<%
				if Not KS.IsNul(rs("groupid"))and rs("groupid")<>"0" And KS.FoundInArr(rs("groupid"),KSUser.GroupID,",")=false And KSUser.GroupID<>1 Then%>
				<input type="submit" disabled value="VIP专用模板"/>
				<%else%>
				<input type="submit" class="button"  value="应用" onClick="window.location='?action=SaveMySkin&id=<%=RS("ID")%>'" />
				<%end if%>
				<input type="submit" class="button"  value="预览" onClick="window.open('../space/showtemplate.asp?templateid=<%=RS("ID")%>');" />
				</li>									
				<%end if%>
			</ul>
			<%
			response.write "</td>"
			rs.movenext
			k=k+1
			if rs.eof or k>=MaxPerPage then exit for 
		   next
		   for i=k+1 to 4
		    response.write "<td width=""25%"">&nbsp;</td>"
		   next
		  response.write "</tr>"
		  if rs.eof or k>=MaxPerPage then exit do
		 loop
		 response.write "<tr>"
		 response.write "<td colspan=4 align=""right"">"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
		 Response.write "</td>"
		 response.write "</tr>"
	   End Sub
	   
	   Sub SaveMySkin()
	     Dim Flag:Flag=KS.ChkClng(KS.S("Flag"))
	     Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 IF ID=0 Then Exit Sub
		 if flag=3 then
		 Conn.Execute("Update KS_Team Set TemplateID=" & ID & " Where id=" & KS.ChkClng(KS.S("TeamID")))
		 response.write "<script>alert('恭喜，成功应用了选中圈子模板！');location.href='?action=Template&flag=3';</script>"
		 else
		 Conn.Execute("Update KS_Blog Set TemplateID=" & ID & " Where UserName='" & KSUser.UserName & "'")
		 response.write "<script>alert('恭喜，成功应用了空间站点模板！');location.href='?action=Template';</script>"
		 end if
		 'response.redirect "?action=Template"
	   End Sub
	   

	   
	  

	   
	   '博文列表
	   Sub BlogList()
			    
				Dim Param
				
					if request("listtype")="1" then
					 param=" inner join ks_friend f on l.username=f.friend where f.accepted=1 and l.UserName='"& KSUser.UserName &"'"
					  Call KSUser.InnerLocation("好友的博文")
					else
					 Param=" Where l.UserName='"& KSUser.UserName &"'"
					  Call KSUser.InnerLocation("我的博文")
					end if
					If KS.S("TypeID")<>"" And KS.S("TypeID")<>"0" Then Param=Param & " And l.TypeID=" & KS.ChkClng(KS.S("TypeID"))					
					
								   
				Dim Sql:sql = "select l.* from KS_BlogInfo l "& Param &" order by l.AddDate DESC,l.id"
			  %>
								     
				         <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                           <%
								 Set RS=Server.CreateObject("AdodB.Recordset")
								  RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>没有你要的博文!</td></tr>"
								 Else
									    totalPut = RS.RecordCount
										If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
												RS.Move (CurrentPage - 1) * MaxPerPage
										End If
								        Call ShowLog

				End If
     %>               
	 </table>
	 <table border="0" cellspacing="0" cellpadding="0" class="border">
	       <tr class='tdbg'>
                                  <form action="User_Blog.asp" method="post" name="searchform">
                                  <td height="45" colspan=2><strong>博文搜索：</strong></td>
								  <td><select name="Flag" class="select">
										   <option value="0">标题</option>
										   <option value="1">标签</option>
									      </select>
										  <select name='TypeID' class="select">
										 <option value="0">-请选择博文分类-</option>
                                           <%  Set Rs = Conn.Execute("SELECT typeid,depth,typeName FROM KS_blogtype ORDER BY rootid,orderid")
							Do While Not Rs.EOF
								Response.Write "<option value=""" & Rs("typeid") & """ "
								If KS.ChkClng(KS.S("TypeID"))=KS.ChkClng(RS("TypeID")) Then Response.Write "selected"
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
							  %>
                                        </select>
										</td><td>
										  关键字</td>
										  <td>
										  <input type="text" name="KeyWord" class="textbox" value="关键字" onfocus="if(this.value=='关键字'){this.value=''}" size=20>&nbsp;<input  class="button" type="submit" name="submit1" value=" 搜 索 ">
							      </td>
								    </form>
                                </tr>
                        </table>
		  <%
  End Sub
  
  Sub ShowLog()
     Dim I,PhotoUrl,KSR
	 Set KSR=New refresh
    Response.Write "<FORM Action=""User_Blog.asp?Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
		%>
		<tr class='tdbg'>
		<%
		PhotoUrl=RS("PhotoUrl")
		If KS.IsNul(PhotoUrl) Then PhotoUrl=KSUser.GetUserInfo("UserFace")
		%>
                   <td class="splittd" style="width:60px;">
						<div class="avatar48"><a title="进入空间" href="../space/?<%=KSUser.GetUserInfo("userid")%>/log/<%=rs("id")%>" target="_blank"><img src="<%=PhotoUrl%>" onerror="this.src='../images/face/boy.jpg';" /></a></div>
					</td>
					<td valign="top" class="splittd" style="width:650px">
					  <div class="Contenttitle"><a href="../space/?<%=KSUser.GetUserInfo("userid")%>/log/<%=rs("id")%>" target="_blank"><%=RS("title")%></a></div>
					  <div class="Contenttips">
					  <span>   
					 <%=KS.GetTimeFormat(rs("adddate"))%> 
					<%  Dim RST:Set RST=Conn.Execute("Select TOP 1 TypeName From KS_BlogType Where TypeID=" & RS("TypeID"))
											IF NOT RST.Eof Then
											   Response.Write " 分类:" & RST(0)
											End If
											RST.Close:Set RST=Nothing%>
										
										 状态：
											  <%Select Case rs("Status")
											   Case 0
											     Response.Write "<span class=""font10"">正常</span>"
                                               Case 2
											     Response.Write "<span class=""font13"">未审</span>"
                                              end select
											  %>
											  </span>
											  <%if (ks.SSetting(3)=1 and rs("status")=0) or rs("istalk")=1 then%>
						 <%else%>
						<a href="User_Blog.asp?id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>">修改</a><%end if%> <a href="javascript:;" onclick = "$.dialog.confirm('确定删除博文吗?',function(){location.href='User_Blog.asp?action=Del&ID=<%=rs("id")%>&<%=KS.QueryParam("action,id")%>';},function(){})">删除</a>
						</div>	
						<div class="blogtext"><%=KS.Gottopic(ks.losehtml(ks.ClearBadChr(ubbcode(rs("content"),1))),160)%>...
						
						<a href="../space/?<%=KSUser.GetUserInfo("userid")%>/log/<%=rs("id")%>" target="_blank">[阅读全文]</a>
						</div>				  
											  </td>
                                           
                                          </tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>

	<tr><td colspan=2><%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%></td></tr>
	<% 
	Set KSR=Nothing
  End Sub
  '删除博文
  Sub ArticleDel()
	  Dim ID:ID=KS.ChkClng(KS.S("ID"))
	  If ID=0 Then 
	    KS.AlertHintScript "出错啦!"
		Exit Sub
	  End If
	  Dim ParamStr
	  If (KS.ChkClng(ksuser.groupid)=1 and KS.ChkClng(KS.C("SuperTF"))=1) Then
	   ParamStr=" Where ID=" & id
	  Else
	   ParamStr=" Where ID=" & id & " And UserName='" & KSUser.UserName & "'"
	  End If
	  Dim RS:Set RS=Conn.Execute("select id From KS_BlogInfo" & ParamStr)
	  Do While Not RS.Eof
	  Call KS.delweibo("空间博文",rs("id"))
	  Conn.Execute("Delete From KS_BlogComment Where LogID=" & RS("id"))
	  RS.MoveNext
	  Loop
	  RS.Close
	  Set RS=Nothing
	  Conn.Execute("Delete From KS_BlogInfo" & ParamStr)
	  KS.Die "<script>$.dialog.tips('恭喜，删除成功！',1,'success.gif',function(){location.href='User_Blog.asp?" & ks.queryparam("id,action") &"';});</script>"
  End Sub
  '添加博文
  Sub ArticleAdd()
        Call KSUser.InnerLocation("发布博文")

		If KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(KS.SSetting(36))  And KS.ChkClng(KS.SSetting(36))>0 Then  '判断有没有到达积分要求
		  KS.Die "<script>$.dialog.tips('对不起，本站要求积分达到 <font color=red>" & KS.ChkClng(KS.SSetting(36)) &"</font> 分才可以发表博文，您当前积分 <font color=green>" & KSUser.GetUserInfo("score") & "</font> 分!',5,'error.gif',function(){history.back();});</script>"
		End If
		Session("UploadFileIDs")=""
  		if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select top 1 * From KS_BlogInfo Where UserID=" & KS.ChkClng(KSUser.GetUserInfo("UserID")) & " And ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
		     TypeID  = RSObj("TypeID")
			 ClassID = RSObj("ClassID")
			 Title    = RSObj("Title")
			 Tags = RSObj("Tags")
			 UserName   = RSObj("UserName")
			 password = RSObj("password")
			 Face   = RSObj("Face")
			 weather=RSObj("Weather")
			 adddate=RSObj("adddate")
			 Content  = RSObj("Content")
			 Status  = RSObj("Status")
		   Else
		     RSObj.Close :Set RSObj=Nothing
		     KS.AlertHIntScript "日志不是你的!"
		   End If
		   RSObj.Close:Set RSObj=Nothing
		Else
		  adddate=now:weather="sun.gif":Face=1:UserName=KSUser.GetUserInfo("RealName")
		  TypeID=KS.ChkClng(Conn.Execute("Select Top 1 TypeID From KS_BlogType Where IsDefault=1")(0))
		End If
		   Response.Write EchoUeditorHead()
		%>
		<script language = "JavaScript">
		function GetKeyTags()
			{
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#Tags').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#Tags').val(unescape(data)).attr("disabled",false);
				  });
			  }else{
			    $.dialog.alert('对不起,请先输入博文标题!',function(){document.myform.Title.focus();});
			  }
			}
				function CheckForm()
				{
				if (document.myform.TypeID.value=="0"){
					$.dialog.alert("请选择博文分类！",function(){document.myform.TypeID.focus();});
					return false;
				  }		
				if (document.myform.Title.value==""){
					$.dialog.alert("请输入博文标题！",function(){document.myform.Title.focus();});
					return false;
				  }	
				  
				  if (editor.hasContents()==false){
					  $.dialog.alert("请输入博文内容！",function(){editor.focus();});
					  return false;
					}
				 return true;  
				}
				function Chang(picurl,V,S)
				{
					var pic=S+picurl
					if (picurl!=''){
					document.getElementById(V).src=pic;
					}
				}
		
		function insertHTMLToEditor(codeStr) { 
		  editor.execCommand('insertHtml', codeStr);
		} 
           
		</script>	
                  <form  action="User_Blog.asp?Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">

                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>博文分类：</span></td>
                       <td width="88%">
					     <table border="0" cellspacing="0" cellpadding="0">
						 <tr><td>
					       <select class="select" size='1' name='TypeID'>
                             <option value="0">-请选择类别-</option>
							  <%
							Set Rs = Conn.Execute("SELECT typeid,depth,typeName FROM KS_blogtype ORDER BY rootid,orderid")
							Do While Not Rs.EOF
								Response.Write "<option value=""" & Rs("typeid") & """ "
								If TypeID=RS("TypeID") Then Response.Write "selected"
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
	
							  
							
							  %>
                         </select></td><td>专栏</td><td><select class="select" size='1' name='ClassID' style="width:150px">
                                   <option value="0">-选择我的专栏-</option>
                                            <%=KSUser.UserClassOption(2,ClassID)%>
                         </select></td><td> <a href="User_Class.asp?Action=Add&typeid=2"><font color="red">添加我的分类</font></a>			</td></tr></table>
					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>博文标题：</span></td>
                              <td><input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                    </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>博文日期：</span></td>
                                      <td><input name="AddDate"  class="textbox" type="text" id="AddDate" value="<%=adddate%>" style="width:250px; " />
                                      <Select class="select" Name="Weather" Size="1" onChange="Chang(this.value,'WeatherSrc','images/weather/')">
									   <Option value="sun.gif"<%if weather="sun.gif" then response.write " selected"%>>晴天</Option>
									   <Option value="sun2.gif"<%if weather="sun2.gif" then response.write " selected"%>>和煦</Option>
									   <Option value="yin.gif"<%if weather="yin.gif" then response.write " selected"%>>阴天</Option>
									   <Option value="qing.gif"<%if weather="qing.gif" then response.write " selected"%>>清爽</Option>
									   <Option value="yun.gif"<%if weather="yun.gif" then response.write " selected"%>>多云</Option>
									   <Option value="wu.gif"<%if weather="wu.gif" then response.write " selected"%>>有雾</Option>
									   <Option value="xiaoyu.gif"<%if weather="xiaoyu.gif" then response.write " selected"%>>小雨</Option>
									   <Option value="yinyu.gif"<%if weather="yinyu.gif" then response.write " selected"%>>中雨</Option>
									   <Option value="leiyu.gif"<%if weather="leiyu.gif" then response.write " selected"%>>雷雨</Option>
									   <Option value="caihong.gif"<%if weather="caihong.gif" then response.write " selected"%>>彩虹</Option>
									   <Option value="hexu.gif"<%if weather="hexu.gif" then response.write " selected"%>>酷热</Option>
									   <Option value="feng.gif"<%if weather="feng.gif" then response.write " selected"%>>寒冷</Option>
									   <Option value="xue.gif"<%if weather="xue.gif" then response.write " selected"%>>小雪</Option>
									   <Option value="daxue.gif"<%if weather="daxue.gif" then response.write " selected"%>>大雪</Option>
									   <Option value="moon.gif"<%if weather="moon.gif" then response.write " selected"%>>月圆</Option>
									   <Option value="moon2.gif"<%if weather="moon2.gif" then response.write " selected"%>>月缺</Option>
									</Select>
		<img id="WeatherSrc" src="images/weather/<%=weather%>" border="0"></td>
                              </tr>
                              
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span>当前心情：</span></td>
                                <td>&nbsp;<input type="radio" name="face" value="0"<%If face=0 Then Response.Write " checked"%>>
        无<input name="face" type="radio" value="1"<%If face=1 Then Response.Write " checked"%>><img src="images/face/1.gif" width="20" height="20"> 
        <input type="radio" name="face" value="2"<%If face=2 Then Response.Write " checked"%>><img src="images/face/2.gif" width="20" height="20"><input type="radio" name="face" value="3"<%If face=3 Then Response.Write " checked"%>><img src="images/face/3.gif" width="20" height="20"> 
        <input type="radio" name="face" value="4"<%If face=4 Then Response.Write " checked"%>><img src="images/face/4.gif" width="20" height="20"> 
        <input type="radio" name="face" value="5"<%If face=5 Then Response.Write " checked"%>><img src="images/face/5.gif" width="20" height="20"> 
        <input type="radio" name="face" value="6"<%If face=6 Then Response.Write " checked"%>><img src="images/face/6.gif" width="18" height="20"> 
        <input type="radio" name="face" value="7"<%If face=7 Then Response.Write " checked"%>><img src="images/face/7.gif" width="20" height="20"> 
        <input type="radio" name="face" value="8"<%If face=8 Then Response.Write " checked"%>><img src="images/face/8.gif" width="20" height="20"> 
        <input type="radio" name="face" value="9"<%If face=9 Then Response.Write " checked"%>><img src="images/face/9.gif" width="20" height="20">
        <input type="radio" name="face" value="10"<%If face=10 Then Response.Write " checked"%>><img src="images/face/10.gif" width="20" height="20">
        <input type="radio" name="face" value="11"<%If face=11 Then Response.Write " checked"%>><img src="images/face/11.gif" width="20" height="20">
        <input type="radio" name="face" value="12"<%If face=12 Then Response.Write " checked"%>><img src="images/face/12.gif" width="20" height="20"></td>
                              </tr>
							 

                              <tr class="tdbg">
                                  <td align="center">博文内容：</td>
								  <td align=left><input type="hidden" name="channelid" id="channelid" value="9993"/><%If KS.SSetting(26)="1" Then%><iframe id="upiframe" name="upiframe" src="BatchUploadForm.asp?ChannelID=9993" frameborder="0" width="550" height="215" scrolling="no"></iframe> <%End If%>   <%
					 					
                      Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:90%;height:250px;"">" & KS.ClearBadChr(Content)&"</script>"
                      Response.Write "<script>setTimeout(""editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(Replace(GetEditorToolBar("Basic"),"]",",'map']"),"'source', '|',","") &"],elementPathEnabled:false,wordCount:false,autoHeightEnabled:false,minFrameHeight:250 });"",10);</script>"
			  
					%>
                                  
								</td>
                            </tr>
                            <tr class="tdbg">
                                      <td height="25" align="center"><span>Tag标 签：</span></td>
                                      <td><input name="Tags" class="textbox" type="text" id="Tags" value="<%=Tags%>" style="width:220px; " /> <a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个Tags请以空格分隔</span></td>
                              </tr>                              
                              <tr class="tdbg">
                                 <td height="25" align="center"><span>查看密码：</span></td>
                                <td> <input name="Password"  class="textbox" type="password" id="PassWord" value="<%=PassWord%>" style="width:250px; " />
                                        </td>
                              </tr>
                    <tr class="tdbg">
					  <td class="clefttitle"></td>
                      <td height="30">
					   <button type="submit" class="pn"><strong>OK,立即发布</strong></button>
					 </td>
                    </tr>
			    </table>
                  </form>
		  <%
  End Sub
%>
<!--#include file="../ks_cls/ubbfunction.asp"-->
<%
   Sub DoSave()
				If KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(KS.SSetting(36))  And KS.ChkClng(KS.SSetting(36))>0 Then  '判断有没有到达积分要求
				  KS.Die "<script>$.dialog.tips('对不起，本站要求积分达到 <font color=red>" & KS.ChkClng(KS.SSetting(36)) &"</font> 分才可以发表博文，您当前积分 <font color=green>" & KSUser.GetUserInfo("score") & "</font> 分!',5,'error.gif',function(){history.back();});</script>"
				End If   
   
                 TypeID=KS.ChkClng(KS.S("TypeID"))
				 ClassID=KS.ChkClng(KS.S("ClassID"))
				 Title=Trim(KS.S("Title"))
				 Tags=Trim(KS.S("Tags"))
				 UserName=Trim(KS.S("UserName"))
				 Face=Trim(KS.S("Face"))
				 weather=KS.S("weather")
				 adddate=KS.S("adddate")
				 Content = Request.Form("Content")
				 Content=KS.ScriptHtml(Content, "A", 3)
				 Content=KS.ClearBadChr(content)
				 PassWord=KS.S("password")
				 Status=KS.ChkClng(KS.S("Status"))
				 
				 If KS.SSetting(35)="1" Then  '远程存图
				 Dim FormPath:FormPath =KS.Setting(3) & KS.Setting(91)& "user/" & KSUser.GetUserInfo("userid") & "/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/" 
				 KS.CreateListFolder (FormPath)
				 Content = KS.ReplaceBeyondUrl(Content, FormPath) 
				 End If
				 
				 Dim RSObj
				  
				  if TypeID="" Then TypeID=0
				  If TypeID=0 Then
				    Response.Write "<script>$.dialog.tips('你没有选择博文分类!',1,'error.gif',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>$.dialog.tips('你没有输入博文标题!',1,'error.gif',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  if not isdate(adddate) then
				    Response.Write "<script>$.dialog.tips('你输入的日期不正确!',1,'error.gif',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  If Content="" Then
				    Response.Write "<script>$.dialog.tips('你没有输入博文内容!',1,'error.gif',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  
				dim picsrc:picsrc=KS.GetPictureFromStr(ubbcode(content,1),1)
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Hits")=0
				  RSObj("UserID")=KSUser.GetUserInfo("userid")
				  RSObj("UserName")=KSUser.UserName
				End If
				  RSObj("Title")=Title
				  RSObj("TypeID")=TypeID
				  RSObj("ClassID")=ClassID
				  RSObj("Tags")=Tags
				  RSObj("Face")=Face
				  RSObj("Content")=Content
				  RSObj("PhotoUrl")=picsrc
				  RSObj("Weather")=weather
				  RSObj("Adddate")=adddate
				  RSObj("Password")=Password
				  RSObj("IsTalk")=0
				  if status=1 then
				  RSObj("Status")=1
				  elseif KS.ChkClng(KS.SSetting(3))=1 Then
				  RSObj("Status")=2
				  Else
				  RSObj("Status")=0
				  end if
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				 RSObj.Close:Set RSObj=Nothing
				 
				dim logstr
				if picsrc<>"" then
				  logstr="[br]" & KS.Gottopic(KS.LoseHtml(UbbCode(Content,0)),60) & "... [url={$GetSiteUrl}space/?" & KSUser.GetUserInfo("userid") & "/log/" & InfoID & "]全文&raquo;[/url][br][img]" & picsrc &"[/img]"
				Else
				  logstr="[br]" & KS.Gottopic(KS.LoseHtml(UbbCode(Content,0)),90) & "... [url={$GetSiteUrl}space/?" & KSUser.GetUserInfo("userid") & "/log/" & InfoID & "]全文&raquo;[/url]"
				end if
				 
				If Not KS.IsNul(Session("UploadFileIDs")) Then 
				 Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & InfoID &",classID=" & ClassID & " Where ID In (" & KS.FilterIds(Session("UploadFileIDs")) & ")")
				End If
				 
				 If KS.ChkCLng(KS.S("ID"))=0 Then
				  Call KS.FileAssociation(1026,InfoID,Content,0)
				  Call KSUser.AddToWeibo(KSUser.UserName,"[url={$GetSiteUrl}space/?" & KSUser.GetUserInfo("userid") & "/log/" & InfoID & "]" & left(Title,40) & "[/url]"& Logstr,2)
				  Call KSUser.InnerLocation("发表博文")
			   	  Response.Write "<script>$.dialog.tips('恭喜，博文发布成功,5秒后返回博文管理中心!<br/><a href=""" & KS.Setting(3) & "space/?" & KSUser.GetUserInfo("userid") & "/log/" & InfoID & """ target=""_blank"" style=""color:red;font-size:12px"">您可以点此访问该博文</a>',5,'success.gif',function(){location.href='user_blog.asp';});</script>"
				 Else
				   Call KS.FileAssociation(1026,InfoID,Content,1) 
				   
				   Call KSUser.InnerLocation("修改博文")
			   	  Response.Write "<script>$.dialog.tips('恭喜，博文修改成功,5秒后返回博文管理中心!<br/><a href=""" & KS.Setting(3) & "space/?" & KSUser.GetUserInfo("userid") & "/log/" & InfoID & """ target=""_blank"" style=""color:red;font-size:12px"">您可以点此访问该博文</a>',5,'success.gif',function(){location.href='user_blog.asp';});</script>"

				 End If
  End Sub


End Class
%> 

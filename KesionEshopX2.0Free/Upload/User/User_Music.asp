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
If Request("action")="play" Then
	Call MusicPlay()
	Response.end()
End If

Dim KSCls
Set KSCls = New User_Friend
KSCls.Kesion()
Set KSCls = Nothing

Class User_Friend
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage,SQL,tablebody,strErr,action,boxName,smscount,smstype,readaction,turl
		Private ArticleStatus,ComeUrl,TotalPages
		Private Sub Class_Initialize()
			MaxPerPage =20
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
		
		action=Trim(request("action"))
		CurrentPage=Trim(request("page"))
		if Isnumeric(CurrentPage) then
			CurrentPage=Clng(CurrentPage)
		else
			CurrentPage=1
		end if
		If Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
		  response.Redirect("User_Blog.asp")
		 Exit Sub
		ElseIf Conn.Execute("Select top 1 status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
		    Response.Write "<script>$.dialog.tips('对不起，你的空间还没有通过审核或被锁定！',1,'error.gif',function(){history.back();});</script>"
			response.end
		End If
		if action<>"play" then
			Call KSUser.SpaceHead()
			Call KSUser.InnerLocation("我的音乐")
			%>
			<div class="tabs">	
				<ul>
					<li class='puton'>我的音乐</li>
				</ul>
			</div>
		<%
		end if
		KSUser.CheckPowerAndDie("s04")
		
		select case action
		case "addlink"  Call AddMusicLink()
		case "addsave" Call AddMusicLinkSave()
		case "del" Call SongDel()
		case else call info()
		end select
		  	%>
		</TD>    
		 </TR>
</TABLE>
		 <%
	  End Sub

		
		sub info()
				
		%>
		<script>
		function AddMusicLink(title,id)
        { location.href="User_Music.asp?action=addlink&id="+id
       }
	   var box='';
	   function play(s,t)
	   {
		box=$.dialog.open('User_Music.asp?pagetitle=试听&action=play&songname='+t+'&songurl='+s,{title:'歌曲试听',width:480,height:200});
	   }
		</script>

			<table height='400' width="100%">
			<tr><td valign="top">
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1"  class="border">
		<form action="?action=del" method=post name=inbox>
			<tr height="23" class="title">
				<td width="5%" align="center">选择</td>
				<td width="20%" height="25" align="center">音乐名称</td>
				<td width="10%" align="center">歌 手</td>
				<td width="15%" align="center">上传时间</td>
				<td width="15%" align="center">试听</td>
				<td width="16%" align="center">操 作</td>
			</tr>
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select * from ks_blogmusic where Username='"&KSUser.UserName&"' order by adddate desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=6 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有上传音乐！</td>
			</tr>
		<%else
		do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
						<td align=center  class="splittd"><input type=checkbox name=id value=<%=rs(0)%>></td>
							<td class="splittd" align=center valign=middle><%=KS.HTMLEncode(rs("SongName"))%></td>

							<td class="splittd" align=center>&nbsp;<%=rs("singer")%>&nbsp;</td>
							<td class="splittd" align=center>&nbsp;<%=KS.GetTimeFormat(rs("adddate"))%>&nbsp;</td>
							<td class="splittd" align=center><a href="#" onClick="play('<%=rs("url")%>','<%=rs("songname")%>')"><img src="images/radio.gif" align="absmiddle" border="0">试听</a></td>
							<td class="splittd" align=center><a href="#" class="box" onClick="AddMusicLink('修改歌曲',<%=rs(0)%>);">修改</a>  <a href="?action=del&id=<%=rs(0)%>" class="box" onClick="return(confirm('确定删除吗?'))">删除</a></td>
						</tr>
		<%
			rs.movenext
			loop
			end if
			rs.close
			set rs=Nothing
		%>
						
				<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"> 
				  <td colspan=6 align=right valign=middle><input type=checkbox name=chkall value=on onClick="CheckAll(this.form)">选中所有显示歌曲&nbsp;<input class="button" type=button name=action onClick="AddMusicLink('添加歌曲',0)" value="添加音乐链接">&nbsp;<input class="button" type=submit name=action onClick="{if(confirm('确定删除选定的歌曲吗?')){this.document.inbox.submit();return true;}return false;}" value="删除选中的歌曲">&nbsp;</td>
				</tr>
		  </form>
</table>
 </td>
 </tr>
 </table>
</div>

		<script language=javascript>
		function CheckAll(form)
		{
		for (var i=0;i<form.elements.length;i++)    {
		var e = form.elements[i];
		if (e.name != 'chkall')       e.checked = form.chkall.checked; 
		}
		}
		</script>
		<%
		end sub
		
		Sub AddMusicLink()
		If KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(KS.SSetting(41))  And KS.ChkClng(KS.SSetting(41))>0 Then  '判断有没有到达积分要求
		  KS.Die "<script>$.dialog.tips('对不起，本站要求积分达到 <font color=red>" & KS.ChkClng(KS.SSetting(41)) &"</font> 分才可以发布音乐，您当前积分 <font color=green>" & KSUser.GetUserInfo("score") & "</font> 分!',5,'error.gif',function(){history.back();});</script>"
		End If
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		  Dim SongName,Url,Singer
		  if id<>0 then
		  Dim RS:Set RS=Server.Createobject("adodb.recordset")
		  rs.open "select * from ks_blogmusic where id="&Id,conn,1,1
		  if not rs.eof then
		   songname=rs("songname")
		   url=rs("url")
		   singer=rs("singer")
		  end if
		  rs.close:set rs=nothing
		  end if
		  Call KSUser.InnerLocation("添加歌曲")
		  %>
		    <html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
			<title></title>
			<link href="images/css.css" type="text/css" rel="stylesheet" />
		
		<script src="../ks_inc/common.js" language="JavaScript"></script>
		  <script>
			function CheckForm()
			 {
			 if (document.myform.SongName.value=='')
			  {
			   $.dialog.alert("请输入歌曲名称!",function(){
			   document.myform.SongName.focus();
			   });
			   return false;
			  }
				
				if (!IsExt(document.myform.Url.value,'mp3'))
				   { $.dialog.alert('音乐格式必须是mp3!',function(){
					  document.myform.Url.focus(); 
					  });
					  return false;
				   }
			 return true;
			}
			function setupload()
			{
			  document.myform.vvvv.style.display='none';
			  document.myform.vvvvv.style.display='';
			  
			}
			</script>
			</head>
			<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
			
			<br>
			<form action="?action=addsave" method=post name=myform onSubmit="return(CheckForm())">

			<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
			  <tr class="tdbg">
                 <td height="25" width="137"> 请输入歌曲名称：</td>
				 <td width="823">
                   <input class="textbox" name="SongName" type="text" style="width:250px; " value="<%=songname%>" maxlength="100" />
                <span style="color: #FF0000">*</span>
				<br><span>如：冰雨、会呼吸的痛 </span></td>
              </tr>
			  <tr class="tdbg">
                 <td height="25"> 播放地址：</td>
				 <td>
                   <input class="textbox" name="Url" type="text" id="Url" style="width:250px; " value="<%=url%>" maxlength="100" />
                 <font style="color: #FF0000">*</font>
				<div name="ss1">如:http://www.kesion.com/冰雨.mp3</div>
				<div name="ss1">
				<iframe id='UpFlashFrame' name='UpFlashFrame' src='User_Upfile.asp?FieldName=Url&type=UpByBar&channelid=9995' frameborder=0 scrolling=no width='300' height='30'> </iframe></div>
				</td>
              </tr>
			  <tr class="tdbg">
                 <td height="25"> 歌 手 ：</td>
				 <td>
                   <input class="textbox" name="Singer" type="text" style="width:150px; " value="<%=singer%>" maxlength="100" />                 <span>如：刘德华、梁静茹</span></td>
              </tr>
			 </table>
			 <br>
			 <div style="text-align:center"><input type="submit" value="确定保存" name="s1" class="button">&nbsp;<input type="button" value="取消返回" onClick="location.href='?';" class="button">
			  <input type="hidden" value="<%=id%>" name="id">
			 </div>
			 </form>
		 	</body>
			</html>
		  <%
		End Sub
        
		Sub AddMusicLinkSave()
		If KS.ChkClng(KSUser.GetUserInfo("score"))<KS.ChkClng(KS.SSetting(41))  And KS.ChkClng(KS.SSetting(41))>0 Then  '判断有没有到达积分要求
		  KS.Die "<script>$.dialog.tips('对不起，本站要求积分达到 <font color=red>" & KS.ChkClng(KS.SSetting(41)) &"</font> 分才可以发布音乐，您当前积分 <font color=green>" & KSUser.GetUserInfo("score") & "</font> 分!',5,'error.gif',function(){history.back();});</script>"
		End If
		
		  Dim SongName:SongName=KS.S("SongName")
		  Dim Url:Url=KS.S("Url")
		  Dim Singer:Singer=KS.S("Singer")
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		  IF SongName="" Then Call KS.AlertHistory("歌曲名称必须输入!",-1):exit sub
		  IF Url="" Then Call KS.AlertHistory("歌曲番放地址必须输入!",-1):exit sub
		  
		  If ID=0 Then
		  Conn.Execute("Insert Into KS_BlogMusic(songname,url,singer,adddate,username) values('" & SongName & "','" & Url & "','" & Singer & "'," & SqlNowString & ",'" & KSUser.UserName &"')")
		  If InStr(Lcase(Url),Lcase(KS.Setting(91)))<>0 Then
		   Dim MaxID:MaxID=Conn.Execute("Select Max(id) From KS_BlogMusic")(0)
		   Call KS.FileAssociation(1027,MaxID,Url,0)
		  End If
		  
		  Response.Write "<script>if (!confirm('恭喜，歌曲添加成功!继续添加吗?')) location.href='User_Music.asp'; else location.href='?action=addlink';</script>"
		  Else
		  Conn.Execute("Update KS_BlogMusic set songname='" & SongName & "',url='" & Url & "',singer='" & Singer & "' where username='" & KSUser.UserName & "' and id=" & ID)
		  If InStr(Lcase(Url),Lcase(KS.Setting(91)))<>0 Then
		   Call KS.FileAssociation(1027,ID,Url,1)
		  End If
		  Response.Write "<script>alert('恭喜，歌曲修改成功!'); location.href='User_Music.asp';</script>"
		  End If
		End Sub
		
		
	
	    Sub SongDel()
		  on error resume next
		  Dim i,id:id=KS.FilterIDs(KS.S("id"))
		  if (id="") then Call KS.AlertHistory("对不起，参数传递出错!",-1):exit sub
		  dim ids:ids=split(id,",")
		  for i=0 to ubound(ids)
		    ks.deletefile(conn.execute("select url from ks_blogmusic where id=" & ids(i) & "and username='" & ksuser.username & "'")(0))
		  next
		  Conn.Execute("delete from ks_blogmusic where id in(" & id & ")")
		  Conn.Execute("delete from KS_UploadFiles Where ChannelID=1027 and infoid in(" & id & ")")
		  Call KS.AlertHintScript("恭喜，删除成功!")
		End Sub
End Class


Sub MusicPlay()
		 Response.Expires = -1 
		Response.ExpiresAbsolute = Now() - 1 
		Response.cachecontrol = "no-cache" 
		dim url:url=request("songurl")
		 %>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
			<title>用户管理中心</title>
			<link href="images/css.css" type="text/css" rel="stylesheet" />
			<META HTTP-EQUIV="pragma" CONTENT="no-cache"> 
			<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate"> 
			<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
			<style>
			 .tt{font-size:14px;color:#191970}
			 .tt span{font-size:12px;color:#999999}
			</style>
			</head>
			<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
			<br>
			<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
			  <tr class="tdbg">
                 
                 <td height="25" class="tt"> 
				 
				  <object id="MediaPlayer1" width="350" height="64" classid="CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6" 
codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,7,1112"
align="baseline" border="0" standby="Loading Microsoft Windows Media Player components..." 
type="application/x-oleobject">
    <param name="URL" value="<%=url%>">
    <param name="autoStart" value="true">
    <param name="invokeURLs" value="false">
    <param name="playCount" value="100">
    <param name="defaultFrame" value="datawindow">
       
		<embed src="<%=url%>" align="baseline" border="0" width="350" height="68"
			type="application/x-mplayer2"
			pluginspage=""
			name="MediaPlayer1" showcontrols="1" showpositioncontrols="0"
			showaudiocontrols="1" showtracker="1" showdisplay="0"
			showstatusbar="1"
			autosize="0"
			showgotobar="0" showcaptioning="0" autostart="1" autorewind="0"
			animationatstart="0" transparentatstart="0" allowscan="1"
			enablecontextmenu="1" clicktoplay="0" 
			defaultframe="datawindow" invokeurls="0">
		</embed>
</object>
				
				<!--<EMBED style="WIDTH: 272px; HEIGHT: 29px" src=<%=url%> width=299 height=10 type=audio/x-wav autostart="true" loop="true"></DIV></EMBED>
				-->
                   <!--
				     <object type='application/x-shockwave-flash' height='20' width='200' data='/ks_inc/dewplayer.swf?son=<%=url%>&autoplay=1&autoreplay=1'>
    <param value='/ks_inc/dewplayer.swf?son=<%=url%>&autoplay=1&autoreplay=1'name='movie' />
    <param name="wmode" value="transparent" />
    <param name="bgcolor" value="" />
  </object>-->
				   
				<br><span><%=Request("songname")%></span></td>
              </tr>

			 </table>
	
			 <div style="text-align:center">&nbsp;<input type="button" value="关闭窗口" onClick="top.box.close();" class="button"></div>
			 </form>
		 	</body>
			</html>
		<%
		End Sub
%> 

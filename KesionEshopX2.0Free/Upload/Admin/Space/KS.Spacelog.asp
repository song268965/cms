<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_SpaceLog
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_SpaceLog
        Private KS,Param,KSR
		Private Action,i,strClass,RS,SQL,maxperpage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSR=Nothing
		End Sub

		Public Sub Kesion()
		If KS.G("action")="pushtopic" then pushtopic:ks.die ""
		With Response
					If Not KS.ReturnPowerResult(0, "KSMS10002") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		.Write "<!DOCTYPE html><html>"
		.Write"<head>"
		.Write "<script src='../../KS_Inc/common.js'></script>"
		.Write "<script src='../../KS_Inc/jquery.js'></script>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write EchoUeditorHead()
		if ks.g("action")<>"showpush" then
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""location.href='KS.Spacelog.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>博文管理</span></li>"
			.Write "<li class='parent' onclick=""location.href='?action=comment';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>博文评论</span></li>"
			.Write "<li class='parent' onclick=""location.href='?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon merge'></i>博文分类</span></li>"
			.Write	" </ul>"
		end if
		End With	
		
		
		maxperpage = 30 '###每页显示数


		Select Case KS.G("action")
		 Case "Del" BlogInfoDel
		 Case "Best" BlogInfoBest
		 Case "recommend" recommend
		 cASE "isslide" isslide
		 Case "CancelBest" BlogInfoCancelBest
		 Case "verific" Verific
		 case "comment" commentshow
		 case "commentdel" commentdel
		 case "class" classshow
		 case "modify" modify
		 case "DoSave" DoSave
		 case "showpush" ShowPush
		 case "pushtopic" pushtopic
		 Case Else
		  Call showmain
		End Select
End Sub
%>
<!--#include file="../../KS_Cls/UbbFunction.asp"-->
<%
'显示推送
sub ShowPush()
%>
<body style="padding:23px">
<div id='showtips'>
<form name='moveform' action='KS.SpaceLog.asp' method='get'><b>博文推送：</b><%=ks.s("title")%><br/><br/>
<b>请选择模型：</b><span id='showmodel'></span>
<select style='width:200px;height:220px;' name='classid' id='classid' size='5'></select>
<br/><strong>推送选项：</strong><label><input type='checkbox' id='recommend' value='1'>推荐</label> <label><input type='checkbox' id='rolls' value='1'>滚动</label> <label><input type='checkbox' id='strip' value='1'>头条</label> <label><input type='checkbox' id='popular' value='1'>热门</label> <br/><strong>发布选项：</strong><label><input type='checkbox' name='pubindex' id='pubindex' value='1' checked>发布首页</label> <label><input type='checkbox' name='pubclass' id='pubclass' value='1' checked>发布栏目页</label> <label><input name='pubcontent' type='checkbox' id='pubcontent' value='1' checked>发布内容页</label><br/><font color='blue'>tips:建议仅将帖子推送到没有自定义字段的文章模型中！！！</font> <div style='text-align:center;margin:20px'><input type='button' onclick='dopush(<%=ks.s("topicid")%>)' value='确定推送' class='button'><input type='hidden' value="<%=ks.s("topicid")%>" name='id' id='id'><input type='hidden' value='pushtopic' name='action'>
<input type='button' value=' 取 消 ' onclick='top.box.close()' class='button'></div>
</form>
</div>
<script>
$(document).ready(function(){
	jQuery.get("../../plus/ajaxs.asp",{action:"GetClubPushModel"},function(r){
		  jQuery("#showmodel").html(unescape(r));
	});
});
function dopush(topicid){
	 var modelId=$("#ModelID option:selected").val();
	 if (modelId==undefined){alert('请选择要推送到的模型!');return false;}
	 var classid=$("#classid option:selected").val();
	 if (classid==undefined){alert("请选择栏目!");return false;}
	 var recommend=0;if($("#recommend").prop("checked")){recommend=1;}
	 var rolls=0;if($("#rolls").prop("checked")){rolls=1;}
	 var strip=0;if($("#strip").prop("checked")){strip=1;}
	 var popular=0;if($("#popular").prop("checked")){popular=1;}
	 var pubindex=0;if($("#pubindex").prop("checked")){pubindex=1;}
	 var pubclass=0;if($("#pubclass").prop("checked")){pubclass=1;}
	 var pubcontent=0;if($("#pubcontent").prop("checked")){pubcontent=1;}
	 jQuery.ajax({type:"get",url:"KS.SpaceLog.asp?action=pushtopic&id="+topicid+"&modelid="+modelId+"&classid="+classid+"&recommend="+recommend+"&rolls="+rolls+"&strip="+strip+"&popular="+popular+"&pubindex="+pubindex+"&pubclass="+pubclass+"&pubcontent="+pubcontent+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
	    d=unescape(d);
		if (d.indexOf('err:')!=-1){
		 alert(d.split(':')[1]);
		}else{
		jQuery("#showtips").html(d);}																																																								   }});
	 return false;
}
function getpushclass(modelid){
	 jQuery.ajax({type:"get",url:"../../plus/ajaxs.asp?action=GetClassOption&channelid="+modelid+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
		jQuery("#classid").empty().append(unescape(d));																																																									   }});
}

</script>
</body>
</html>
<%
end sub


'帖子推送
sub pushtopic()
		Dim TopicId:TopicId=KS.ChkClng(KS.S("ID"))
		Dim ModelID:ModelID=KS.ChkClng(KS.S("ModelID"))
		Dim ClassID:ClassID=KS.S("ClassID")
		If TopicId=0 Then  KS.Die escape("err:对不起,您没有选中要推送的博文!")
		If ModelID=0 Then  KS.Die escape("err:对不起，您没有选择模型!") 
		If ClassID="" Then KS.Die escape("err:对不起，您没有选择目标栏目!")
		
		Dim RS:Set RS=Conn.Execute("Select top 1 * From KS_BlogInfo Where ID=" & TopicID)
		If RS.Eof And RS.Bof Then
		  RS.Close : Set RS=Nothing
		  KS.Die Escape("err:对不起，找不到要推送的博文!")
		End If
		Dim Title,PhotoUrl,Inputer,PostTable,Content,IsPic,Hits,Fname,FnameType,TemplateID,WapTemplateID
		Title=KS.LoseHtml(RS("Title"))
		PhotoUrl=RS("PhotoUrl") : If PhotoUrl<>"" Then IsPic=1 Else IsPic=0
		Inputer=RS("UserName"): Hits=RS("Hits")
		
		Content=Ubbcode(RS("Content"),0)
		RS.Close
		Dim Recommend:Recommend=KS.ChkClng(KS.G("Recommend"))
		Dim Rolls:Rolls=KS.ChkClng(KS.G("Rolls"))
		Dim Strip:Strip=KS.ChkClng(KS.G("Strip"))
		Dim Popular:Popular=KS.ChkClng(KS.G("Popular"))
		Dim pubindex:pubindex=KS.ChkClng(KS.G("pubindex"))
		Dim PubClass:PubClass=KS.ChkClng(KS.G("PubClass"))
		Dim PubContent:PubContent=KS.ChkClng(KS.G("PubContent"))
		
		Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
		RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",conn,1,1
			 if RSC.Eof Then 
			      RSC.Close :Set RSC=Nothing
				  KS.Die escape("err:栏目不存在!")
			 Else
					 FnameType=RSC("FnameType")
					 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
					 TemplateID=RSC("TemplateID")
					 WapTemplateID=RSC("WapTemplateID")
			End If
		 RSC.Close:Set RSC=Nothing
		Dim Intro:Intro=KS.Gottopic(Content,200)
		RS.Open "select top 1 * from " & KS.C_S(ModelID,2) &" where 1=0", conn, 1, 3
			RS.AddNew
			RS("Title")          = Title
			RS("Intro")          = Intro
			RS("ArticleContent") = Content
			RS("PicNews")        = IsPic
			RS("PhotoUrl")       = PhotoUrl
			RS("Recommend")      = Recommend
			RS("Rolls")          = Rolls
			RS("Strip")          = Strip
			RS("Popular")        = Popular
			RS("Verific")        = 1
			RS("IsTop")  = 0 : RS("IsVideo")  = 0 : RS("Slide")=0
			RS("Tid")            = classid
			RS("Author")         = Inputer
			RS("AddDate")        = Now
			RS("Rank")           = "★★★"
			RS("Comment")        = 1 : RS("Changes")   = 0 : RS("DelTF")   = 0 : RS("ReadPoint")   = 0
			RS("TemplateID")     = TemplateID
			RS("WapTemplateID")  = WapTemplateID
			RS("Hits")           = Hits
			RS("HitsByDay")      = 0 : 	RS("HitsByWeek")     = 0 : RS("HitsByMonth")    = 0
			RS("Fname")          = Fname
			RS("Inputer")        = Inputer
			RS("RefreshTF")      = PubContent
			RS("PostID")         = 0
			RS("OrderID")        = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ModelID,2) & " Where Tid='" & ClassId &"'")(0))+1
            RS.Update
		    RS.MoveLast
			Dim ItemID:ItemID=RS("ID")
		  If Left(Ucase(Fname),2)="ID" Then
				RS("Fname") = ItemID & FnameType
				RS.Update
		  End If
					 
			 Call LFCls.AddItemInfo(ModelID,ItemID,Title,ClassID,KS.Gottopic(Intro,200),"",PhotoUrl,now,Inputer,Hits,0,0,0,Recommend,Rolls,Strip,Popular,0,0,1,1,RS("Fname"))
	 		 '关联上传文件
			Call KS.FileAssociation(ModelID,Rs("ID"),Content & PhotoUrl,0)
			Dim RefreshTips
			If PubContent=1 Or PubClass=1 Or PubIndex=1 Then
				Dim KSRObj:Set KSRObj=New Refresh
				If PubContent=1 Then
					If (KS.C_S(ModelID,7)="1" or KS.C_S(ModelID,7)="2") Then
							 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
							 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
							  KSRObj.ModelID=ModelID
							  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
							  Call KSRObj.RefreshContent()
							  RefreshTips="生成内容页成功! 地址:<a href='" & KS.GetItemURL(modelID,classid, KSRObj.Node.SelectSingleNode("@id").text, KSRObj.Node.SelectSingleNode("@fname").text, KSRObj.Node.SelectSingleNode("@adddate").text) & "' target='_blank'>" & KS.GetItemURL(modelID,classid, KSRObj.Node.SelectSingleNode("@id").text, KSRObj.Node.SelectSingleNode("@fname").text, KSRObj.Node.SelectSingleNode("@adddate").text) & "</a><BR/>"
					End If
					RS.Close
				End If
				
				If PuBClass=1 And KS.C_S(ModelID,7)="1" Then
				 Dim TS:TS=KS.C_C(ClassID,8)
				 If TS<>"" Then
				   FCls.FsoListNum=3
				   RS.Open "Select * From KS_Class Where ID in('" & Replace(TS,",","','") & "')",Conn,1,1
				   Do While Not RS.EOf 
				    Call KSRobj.RefreshFolder(RS("ChannelID"),RS)
				    RefreshTips=RefreshTips & "生成栏目页成功！地址：<a href='" & KS.GetFolderPath(RS("ID")) & "' target='_blank'>" & KS.GetFolderPath(RS("ID")) &"</a><br/>"
				   RS.MoveNext
				   Loop
				   RS.Close
				 End If
				End If
				If PubIndex=1 And Split(KS.Setting(5),".")(1)<>"asp" Then
				  Call KSRObj.FSOSaveFile(KSRObj.ReplaceRA(KSRObj.KSLabelReplaceAll(KSRObj.LoadTemplate(KS.Setting(110))),""), KS.Setting(3) & KS.Setting(5))
				  RefreshTips=RefreshTips & "生成网站首页成功！地址：<a href='" & KS.GetDomain & KS.Setting(5)& "' target='_blank'>" & KS.GetDomain & KS.Setting(5) &"</a><br/>"
				End If
				Set KSRobj=Nothing
            Else
			 RS.Close
			End If
			Set RS = Nothing
	If RefreshTips="" Then 
	  RefreshTips=	"恭喜，博文成功推送！<a href=""../item/../show.asp?m=" & ModelID & "&d=" & ItemID & """ target=""_blank"">点此查看</a>!" 
    Else 
	 RefreshTips="<strong><font color=""#ff6600"">恭喜，博文成功推送! </font></strong><br/> "& RefreshTips	
	End If
	KS.Die Escape(RefreshTips)
end sub 


Private Sub showmain()
%>
<script>
function topicpush(topicid,title){
		 top.openWin('博文推送','space/KS.SpaceLog.asp?topicid='+topicid+'&title='+title+'&Action=showpush',false,680,380);
}

</script>

<form action="KS.SpaceLog.asp" name="myform" method="post">
   <div class="tableTop">
   <table><tr><td>
     <strong class="mr0">快速搜索=></strong>
	 <span class="tiaoJian">关键字:</span><input type="text" class='textbox' name="keyword"><span class="tiaoJian">条件:</span>
	 <select name="condition">
	  <option value=1>博文标题</option>
	  <option value=2>创建者</option>
	 </select>
	  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
   </td></tr></table>
   </div>
</form>

<div class="pageCont2 mt20">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>标题</th>
	<td nowrap>用户名</th>
	<td nowrap>添加时间</th>
	<td nowrap>推荐</th>
	<td nowrap>幻灯</th>
	<td nowrap>精华</th>
	<td nowrap>状 态</th>
	<td nowrap>管理操作</th>
</tr>
<%
		Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and title like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and username like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If Request("from")<>"" Then
		 Param=Param & " and status=2"
		End If
		If Request("Istalk")<>"" Then
		  Param=Param & " And Istalk= " & KS.ChkClng(KS.S("Istalk"))
		End If

		totalPut = Conn.Execute("Select Count(ID) from KS_bloginfo" & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_BlogInfo " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""65"" align=center bgcolor=""#ffffff"" colspan=""10"">没有人写博文！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="ks.spacelog.asp">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd">
	<a href="../../space/?<%=rs("userid")%>/log/<%=rs("id")%>" target="_blank">
	<%Response.write rs("title")
	%></a>
	<%if not KS.isnul(rs("photourl")) then response.write " <span style='color:red'><i>图</i></span>"%>
	</td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><a href="KS.SpaceLog.asp?action=recommend&id=<%=RS("id")%>&v=<%=rs("recommend")%>"><%if rs("recommend")="1" then response.write "<span style='color:green'>是</span>" else response.write "否"%></a></td>
	<td class="splittd" align="center"><a href="KS.SpaceLog.asp?action=isslide&id=<%=RS("id")%>&v=<%=rs("isslide")%>"><%if rs("isslide")="1" then response.write "<span style='color:green'>是</span>" else response.write "否"%></a></td>
	<td class="splittd" align="center"> <%IF rs("best")="1" then %><a href="?Action=CancelBest&id=<%=rs("id")%>"><span style='color:green'>是</span></a><%else%><a href="?Action=Best&id=<%=rs("id")%>">否</a><%end if%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "正常"
	 case 1
	  response.write "<font color=blue>草稿</font>"
	 case else
	  response.write "<font color=red>未审</font>"
	end select
	%></td>
	<td class="splittd" align="center">
	<a href="../../space/?<%=rs("username")%>/log/<%=rs("id")%>" target="_blank">浏览</a> <a href="?Action=Del&ID=<%=RS("ID")%>" onClick="return(confirm('确定删除该博文吗？'));">删除</a>&nbsp;
	<%if rs("status")=2 then%><a href="?Action=verific&flag=0&id=<%=rs("id")%>">审核</a> <%elseif rs("status")=0 then%><a href="?Action=verific&flag=2&id=<%=rs("id")%>" title="取消审核">取审</a><%end if%>
	
	<a href="?action=modify&id=<%=rs("id")%>" onClick="window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr='+escape('空间门户管理 >> <font color=red>修改博文</font>')+'&ButtonSymbol=GOSave';">修改</a>
	<a href="javascript:;" onClick="topicpush(<%=rs("id")%>,'<%=rs("title")%>');">推送</a>
	
	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr   class='operatingBox' onMouseOver="this.className='operatingBox'" onMouseOut="this.className='operatingBox'">
	<td height='25'  colspan="3" class="pt10">
	&nbsp;&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input type="hidden" name="action" value="Del">
	<input type="hidden" name="flag" value="0">
	<input type="hidden" name="istalk" value="<%=Request("IsTalk")%>">
	<input class="button" type="submit" name="Submit2" value="批量删除" onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){document.selform.action.value='Del';this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" value="批量审核" onClick="document.selform.action.value='verific';document.selform.flag.value='0';this.document.selform.submit();return true;">
	<input class="button" type="submit" value="批量取消审核" onClick="document.selform.action.value='verific';document.selform.flag.value='2';this.document.selform.submit();return true;">
    </form>

	</td><td colspan=7 style="text-align:right">
	<%
	  Call KS.ShowPage(totalput, MaxPerPage, "",CurrentPage,true,true)
	%></td>
</tr>
</table>

</div>
<%
End Sub

Sub recommend()
  dim id:id=KS.FilterIds(KS.S("ID"))
  If Id="" Then KS.AlertHintScript "请选择要操作的博文"
  Dim v:v=KS.ChkClng(Request("v"))
  If V=0 Then
    Conn.Execute("Update KS_BlogInfo Set recommend=1 Where id in(" & id & ")")
  Else
    Conn.Execute("Update KS_BlogInfo Set recommend=0 Where id in(" & id & ")")
  End If
  Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
End Sub
Sub isslide()
  dim id:id=KS.FilterIds(KS.S("ID"))
  If Id="" Then KS.AlertHintScript "请选择要操作的博文"
  Dim v:v=KS.ChkClng(Request("v"))
  If V=0 Then
    Conn.Execute("Update KS_BlogInfo Set isslide=1 Where id in(" & id & ")")
  Else
    Conn.Execute("Update KS_BlogInfo Set isslide=0 Where id in(" & id & ")")
  End If
  Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
End Sub

'修改博文
sub modify()
           Dim RSObj,TypeID,ClassID,Title,Tags,UserName,PassWord,face,weather,adddate,content,status,IsTalk,PhotoUrl,IsSlide,Recommend
		   Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select top 1 * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
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
			 IsTalk  = RSObj("IsTalk")
			 PhotoUrl=RSObj("PhotoUrl")
			 IsSlide=RSObj("IsSlide")
			 Recommend=RSObj("Recommend")
		   End If
		   RSObj.Close:Set RSObj=Nothing
%>
<script language = "JavaScript">
function CheckForm()
{
 document.myform.submit();
}
</script>
  <form  action="?Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
<dl class="dtable">
                    <dd>
                       <div>博文分类：</div>
                          <select name='TypeID' class="select">
										 <option value="0">-请选择博文分类-</option>
                                           <%  Set Rs = Conn.Execute("SELECT typeid,depth,typeName FROM KS_blogtype ORDER BY rootid,orderid")
							Do While Not Rs.EOF
								Response.Write "<option value=""" & Rs("typeid") & """ "
								If KS.ChkClng(TypeID)=KS.ChkClng(RS("TypeID")) Then Response.Write "selected"
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
						
                    </dd>
                    <dd>
                      <div>博文标题：</div>
                       <input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                      <span style="color: #FF0000">*</span>
                    </dd>
                    <dd>
                                      <div>博文日期：</div>
                                        <input name="AddDate"  class="textbox" type="text" id="AddDate" value="<%=adddate%>" style="width:250px; " />
                                       天气<Select Name="Weather" Size="1" onChange="Chang(this.value,'WeatherSrc','../../user/images/weather/')">
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
		<img id="WeatherSrc" src="../../user/images/weather/<%=weather%>" border="0">
                              </dd>
                              <dd>
                                      <div>Tag标 签：</div>
                                        <input name="Tags"  class="textbox" type="text" id="Tags" value="<%=Tags%>" style="width:250px; " />
                                        <span>以空格分隔</span>
                              </dd>
                              <dd>
                                      <div>图 片：</div>
                                        <input name="PhotoUrl"  class="textbox" type="text" id="PhotoUrl" value="<%=PhotoUrl%>" style="width:250px; " />
                                       <input class='button' type='button' name='Submit' value='选择图片...' onClick="OpenThenSetValue('Include/SelectPic.asp?CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.PhotoUrl);"> <input class='button' type='button' name='Submit' value='远程抓图...' onClick="top.openWin('抓取远程图片','include/SaveBeyondfile.asp?fieldid=PhotoUrl&CurrPath=<%=KS.GetUpFilesDir()%>',false,500,100);"> 
                              <dd>
                              <dd>
                                      <div>博文心情：</div>
									  &nbsp;&nbsp;<input type="radio" name="face" value="0"<%If face=0 Then Response.Write " checked"%>>
        无<input name="face" type="radio" value="1"<%If face=1 Then Response.Write " checked"%>><img src="../../user/images/face/1.gif" width="20" height="20"> 
        <input type="radio" name="face" value="2"<%If face=2 Then Response.Write " checked"%>><img src="../../user/images/face/2.gif" width="20" height="20"><input type="radio" name="face" value="3"<%If face=3 Then Response.Write " checked"%>><img src="../../user/images/face/3.gif" width="20" height="20"> 
        <input type="radio" name="face" value="4"<%If face=4 Then Response.Write " checked"%>><img src="../../user/images/face/4.gif" width="20" height="20"> 
        <input type="radio" name="face" value="5"<%If face=5 Then Response.Write " checked"%>><img src="../../user/images/face/5.gif" width="20" height="20"> 
        <input type="radio" name="face" value="6"<%If face=6 Then Response.Write " checked"%>><img src="../../user/images/face/6.gif" width="18" height="20"> 
        <input type="radio" name="face" value="7"<%If face=7 Then Response.Write " checked"%>><img src="../../user/images/face/7.gif" width="20" height="20"> 
        <input type="radio" name="face" value="8"<%If face=8 Then Response.Write " checked"%>><img src="../../user/images/face/8.gif" width="20" height="20"> 
        <input type="radio" name="face" value="9"<%If face=9 Then Response.Write " checked"%>><img src="../../user/images/face/9.gif" width="20" height="20">
        <input type="radio" name="face" value="10"<%If face=10 Then Response.Write " checked"%>><img src="../../user/images/face/10.gif" width="20" height="20">
        <input type="radio" name="face" value="11"<%If face=11 Then Response.Write " checked"%>><img src="../../user/images/face/11.gif" width="20" height="20">
        <input type="radio" name="face" value="12"<%If face=12 Then Response.Write " checked"%>><img src="../../user/images/face/12.gif" width="20" height="20"></dd>
							 
                              <dd>
                                  <div>博文内容：</div>
								  <%if istalk=1 Then%>
								  <textarea ID='Content' name='Content' style='width:400px;height:100px'><%=Content%></textarea>
								  <%else%>
								      <%
										 Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:80%;height:320px;"">" &content&"</script>"
										 Response.Write "<script>setTimeout(""var editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & GetEditorToolBar("newstool") &"],wordCount:false,autoHeightEnabled:false });"",10);</script>"
										%>
								  
			                     <%end if%>
                            </dd>
                              <dd>
                                  <div>是否幻灯：</div>
								  <input type="radio" name="IsSlide" value="1"<%if IsSlide="1" then response.write " checked"%>>是
								  <input type="radio" name="IsSlide" value="0"<%if KS.ChkClng(IsSlide)="0" then response.write " checked"%>>否
                            </dd>
                              <dd>
                                  <div>是否推荐：</div>
								  <input type="radio" name="Recommend" value="1"<%if recommend="1" then response.write " checked"%>>是
								  <input type="radio" name="Recommend" value="0"<%if KS.ChkClng(recommend)="0" then response.write " checked"%>>否
                            </dd>
                              
			    </dl> 
				</form>

<%
end sub

sub DoSave()
     dim TypeID,ClassID,Title,Tags,UserName,PassWord,face,weather,adddate,content
                 TypeID=KS.ChkClng(KS.S("TypeID"))
				 Title=Trim(KS.S("Title"))
				 Tags=Trim(KS.S("Tags"))
				 UserName=Trim(KS.S("UserName"))
				 Face=Trim(KS.S("Face"))
				 weather=KS.S("weather")
				 adddate=KS.S("adddate")
				 Content = Request.Form("Content")
				  Dim RSObj
				  if TypeID="" Then TypeID=0
				  If TypeID=0 Then
				    Response.Write "<script>top.$.dialog.alert('你没有选择博文分类!',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>top.$.dialog.alert('你没有输入博文标题!',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  if not isdate(adddate) then
				    Response.Write "<script>top.$.dialog.alert('你输入的日期不正确!',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				  If Content="" Then
				    Response.Write "<script>top.$.dialog.alert('你没有输入博文内容!',function(){history.back();});</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				  RSObj("Title")=Title
				  RSObj("TypeID")=TypeID
				  RSObj("Tags")=Tags
				  RSObj("Face")=Face
				  RSObj("Weather")=weather
				  RSObj("Adddate")=adddate
				  RSObj("Content")=Content
				  RSObj("IsSlide")=KS.ChkClng(Request("IsSlide"))
				  RSObj("Recommend")=KS.ChkClng(Request("Recommend"))
				  RSObj("PhotoUrl")=KS.S("PhotoUrl")
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(1026,InfoID,Content,1) 
				Response.Write "<script>top.$.dialog.alert('博文修改成功!',function(){location.href='space/KS.Spacelog.asp';});</script>"
end sub

'博文评论管理
Sub commentshow()
		totalPut = Conn.Execute("Select Count(ID) from KS_BlogComment")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1%>
<table width="100%" border="0" align="center" cellspacing="1" cellpadding="1">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>评 论 内 容</td>
	<td nowrap>发 表 人</td>
	<td nowrap>评 论 时 间</td>
	<td nowrap>回 复 与 否</td>
	<td nowrap>管 理 操 作</td>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_BlogComment order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>没有人发表评论！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=commentdel>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd">
	<strong>标题:</strong><a href="../space/?<%=rs("username")%>/log/<%=rs("logid")%>" target="_blank"><%=Rs("title")%></a>
	<br/><strong>内容:</strong><%=KS.Gottopic(KS.LoseHtml(rs("content")),150)%></td>
	<td class="splittd" align="center"><%=Rs("AnounName")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%if not isnull(rs("Replay")) or rs("replay")<>"" then response.write "已回复" else response.write "<font color=red>未回复</font>"%></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>/log/<%=rs("logid")%>" target="_blank">浏览</a> <a href="?Action=commentdel&ID=<%=RS("ID")%>" onClick="return(confirm('确定删除该评论吗？'));">删除</a> </td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的评论 " onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=7 align=right>
	<%
	   Call KS.ShowPage(totalput, MaxPerPage, "",CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub
'删除评论
Sub CommentDel()
 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Delete From KS_BlogComment Where ID In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub


'删除博文
Sub BlogInfoDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Delete From KS_BlogInfo Where ID In("& id & ")")
 Call KS.delweibo("空间博文",id)
 Conn.Execute("Delete From KS_UploadFiles Where channelid=1026 and InfoID In(" & ID & ")")
 KS.AlertHintScript "删除成功！"
End Sub
'设为精华
Sub BlogInfoBest()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_BlogInfo Set Best=1 Where ID In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消精华
Sub BlogInfoCancelBest()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_BlogInfo Set Best=0 Where ID In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub Verific()
 Dim ID:ID=Replace(KS.G("ID")," ","")
  If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.Execute("Update KS_BlogInfo Set status=" & KS.ChkClng(KS.G("Flag")) & " where id in(" & id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'博文分类管理
Sub classshow()
  select case ks.s("flag")
    case "add" addclass : response.End()
	case "edit" editClass: response.End()
	case "savenew" savenewclass:response.End()
	case "savedit" saveeditclass:response.End()
	case "del" delClass: response.End()
	case "updatorders" updatorders:response.End()
	case "classorders" ClassOrders:response.end()
	case "total" ClassTotal:response.End()
  end select 
   Dim Rs,SQL,i
	Dim tdstyle
	%>
	<div style="margin:10px">
	 <input type="button" class="button" onClick="parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('空间门户 >> <font color=red>添加博文分类</font>')+'&ButtonSymbol=Go';location.href='?action=class&flag=add';" value="添加博文分类"/>
	 
	 <input type="button" class="button" onClick="location.href='?action=class&flag=classorders';" value="一级分类排序"/>
	 <input type="button" class="button" onClick="location.href='?action=class&flag=total';" value="重计各分类下的博文数"/>
	</div>

	<%
	Response.Write " <table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
	Response.Write " <tr class='sort'>"
	Response.Write " <td>博文分类名称 </td>"
	Response.Write " <td width=""200"">管理选项</td>"
	Response.Write " <td width=""100"">分类ID</td>"
	Response.Write "</tr>" & vbNewLine
	SQL = "SELECT * FROM KS_BlogType ORDER BY rootid,orderid"
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		Response.Write " <tr> <td align=""center"" colspan=""2"" class=""tablerow1"">您还没有添加任何博文分类！</td></tr>"
	End If
	i = 0
	Do While Not Rs.EOF
		Response.Write " <tr>"
		Response.Write " <td class='splittd'>"
		Response.Write " &nbsp;&nbsp;"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
			Next
			Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
		End If
		If KS.ChkClng(Rs("parentid")) = 0 Then Response.Write ("<i class='icon manage'></i><b>")
		Response.Write Rs("TypeName")
		If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
		Response.Write " </td>" & vbNewLine
		Response.Write " <td class='splittd' align=""center"">"
		Response.Write "<a onclick=""parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('空间门户 >> <font color=red>添加博文分类</font>')+'&ButtonSymbol=Go';"" href=""?action=class&flag=add&editid="
		Response.Write Rs("typeid")
		Response.Write """>添加分类</a>"
		Response.Write " | <a onclick=""parent.frames['BottomFrame'].location.href='../Post.Asp?OpStr='+escape('空间门户 >> <font color=red>修改博文分类</font>')+'&ButtonSymbol=GoSave';"" href=""?action=class&flag=edit&editid="
		Response.Write Rs("typeid")
		Response.Write """>修改分类</a>"
		Response.Write " |"
		Response.Write " "
		If Rs("child") < 1 Then
			Response.Write " <a href=""?action=class&flag=del&editid="
			Response.Write Rs("typeid")
			Response.Write """ onclick=""{if(confirm('删除将包括该分类的所有信息，确定删除吗?')){return true;}return false;}"">删除分类</a>"
		Else
			Response.Write " <a href=""#"" onclick=""{if(confirm('该分类含有下属分类，必须先删除其下属分类方能删除本分类！')){return true;}return false;}"">"
			Response.Write " 删除分类</a>"
		End If
		Response.Write " </td>" & vbNewLine
		Response.Write " <td class='splittd' align=""center"">" & RS("TypeID")&"</td>" & vbNewLine
		Response.Write "</tr>" & vbNewLine
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
	%>
	<br/><br/>
	<%
	
End Sub

sub addclass()

%>
<script language="javascript">
function CheckForm(){ 
 if ($('#ClassName').val()=='')
 {
   top.$.dialog.alert('请输入分类名称!',function(){
   $('#ClassName').focus();});
   return false;
 }
 $("#myform").submit();
}
</script>
<div style="text-align:center;height:30px;line-height:30px;font-weight:bold">添加博文分类</div>
<dl class="dtable">
	<form name="myform" id="myform" method="POST" action="?action=class&flag=savenew">
	<dd>
		<div>所属分类：</div>
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">做为一级分类</option>"
	SQL = "SELECT typeid,depth,typeName FROM KS_blogtype ORDER BY rootid,orderid"
	Set Rs = Conn.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("typeid") & """ "
		If Request("editid") <> "" And CLng(Request("editid")) = Rs("typeid") Then Response.Write "selected"
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
	Response.Write "</select>"
	Set Rs = Nothing
%>
	</dd>
	<dd>
		<div>分类名称：</div>
		<textarea name="ClassName" id="ClassName" class="textbox" cols="50" rows="5"></textarea>
        <span class="block">添加多个分类请用回车分开</span>
	</dd>
	<dd>
		<div>分类说明：</div>
		<textarea name="Readme" cols="50" rows="5" class="textbox"></textarea><span class="block">可以留空</span>
	</dd>

	</form>
</dl>
<%
end sub

Sub savenewclass()
	If Trim(Request.Form("classname")) = "" Then
		Call KS.AlertHistory("请输入分类名称!",-1)
		Exit Sub
	End If
	If Not IsNumeric(Request.Form("class")) Then
		Call KS.AlertHistory("请选择所属分类!",-1)
		Exit Sub
	End If

	Dim Rs,SQL,i
	Dim newclassid,rootid,ParentID,depth,orders
	Dim maxrootid,Parentstr,neworders
	Dim m_strClassname,m_arrClassname,strClassname

	m_strClassname = Replace(Trim(Request("classname")), vbCrLf, "$$$")
	m_arrClassname = Split(m_strClassname, "$$$")

	If ks.chkclng(Request("class")) <> 0 Then
		SQL = "SELECT rootid,typeid,depth,orderid,Parentstr FROM KS_BlogType WHERE typeid=" & KS.ChkClng(Request("class"))
		Set Rs = Conn.Execute (SQL)
		rootid = Rs(0)
		ParentID = Rs(1)
		depth = Rs(2)
		orders = Rs(3)+1
		If depth >=2 Then
			Call KS.AlertHistory("本系统限制3级分类",-1)
			Exit Sub
		End If
		Parentstr = Rs(4)
		Set Rs = Nothing
	Else
		SQL = "SELECT MAX(rootid) FROM KS_BlogType"
		Set Rs = Conn.Execute (SQL)
		maxrootid = KS.ChkClng(Rs(0)) + 1
		orders=1
		If maxrootid =0 Then maxrootid = 1
		Set Rs = Nothing
	End If


	
	Set Rs=Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM KS_BlogType"
	Rs.Open SQL, Conn, 1, 3
	For i = 0 To UBound(m_arrClassname)
		strClassname = KS.R(Trim(m_arrClassname(i)))
		If strClassname <> "" Then
			Rs.addnew
			If Request("class") <> "0" Then
				Rs("depth") = depth + 1
				Rs("rootid") = rootid
				Rs("parentid") = ks.chkclng(Request.Form("class"))
			Else
				Rs("depth") = 0
				Rs("rootid") = maxrootid
				Rs("parentid") = 0
			End If
			Rs("child") = 0
			Rs("orderid") = ks.chkclng(orders)
			Rs("typename") = strClassname
			Rs("readme") = Trim(Request.Form("readme"))
			Rs("lognum") = 0
			Rs.Update
			Rs.movelast
			Rs("parentstr")=parentstr & rs("typeid") & ","
            RS.Update
			maxrootid = maxrootid + 1
			orders=orders+1
		End If
	Next
	Rs.Close
	Set Rs = Nothing

	CheckAndFixClass 0,1
	Call KS.Confirm("恭喜您！添加新的分类成功,继续添加吗?","space/KS.SpaceLog.asp?action=class&flag=add","?action=class")
End Sub

Sub editClass()
	Dim RsObj
	Dim Rs,SQL,i
	Set Rs = Conn.Execute("SELECT top 1 * FROM KS_BlogType WHERE typeid = " & KS.ChkClng(Request("editid")))
	If Rs.BOF And Rs.EOF Then
		 KS.AlertHintScript "数据库出现错误,没有此站点分类!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
%>
<script language="javascript">
function CheckForm(){ 
 if ($('#ClassName').val()=='')
 {
    top.$.dialog.alert('请输入分类名称!',function(){
   $('#ClassName').focus();
	});
   return false;
 }
 $("#myform").submit();
}
</script>
<div style="text-align:center;height:30px;line-height:30px;font-weight:bold">编辑博文分类</div>
<dl class="dtable">
	<form name="myform" id="myform" method="POST" action="?action=class&flag=savedit">
	<input type="hidden" name="editid" value="<%=Request("editid")%>">
	<dd>
		<div>所属分类：</div>
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">做为一级分类</option>"
	SQL = "SELECT typeid,depth,TypeName FROM KS_BlogType ORDER BY rootid,orderid"
	Set RsObj = Conn.Execute(SQL)
	Do While Not RsObj.EOF
		Response.Write "<option value=""" & RsObj("typeid") & """ "
		If CLng(Rs("parentid")) = RsObj("typeid") Then Response.Write "selected"
		Response.Write ">"
		If RsObj("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If RsObj("depth") > 1 Then
			For i = 2 To RsObj("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write RsObj("TypeName") & "</option>" & vbCrLf
		RsObj.movenext
	Loop
	RsObj.Close
	Response.Write "</select>"
	Set RsObj = Nothing
%>
	</dd>
	<dd>
		<div>分类名称：</div>
		<input type="text" name="ClassName" id="ClassName" class="textbox" size="35" value="<% = Rs("TypeName")%>">
	</dd>
	<dd>
		<div>分类说明：</div>
		<textarea name="Readme" cols="50" class="textbox" rows="5"><%=Server.HTMLEncode(Rs("readme")&"")%></textarea>
	</dd>
	<dd>
		<div>分类统计：</div>
		博文数：<input type="text" class="textbox" name="LogNum" size="10" value="<%=Rs("LogNum")%>">
		</span>
	</dd>
	</form>
</dl>
<%
Set Rs = Nothing
End Sub


Sub saveeditclass()
	If CLng(Request.Form("editid")) = CLng(Request.Form("class")) Then
		Call KS.AlertHistory("所属分类不能指定自己",-1)
		Exit Sub
	End If
	If Trim(Request.Form("classname")) = "" Then
		Call KS.AlertHistory("请输入分类名称!",-1)
		Exit Sub
	End If
	
	Dim newclassid,maxrootid,readme
	Dim parentid,depth,child,ParentStr,rootid,iparentid,iParentStr
	Dim trs,mrs
	Dim Rs,SQL,nParentStr
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT ParentStr FROM KS_BlogType Where typeID=" & KS.ChkClng(KS.G("Class")),conn,1,1
	If Not RS.Eof Then
	 nParentStr=Rs(0)
	End If
	Rs.Close
	SQL = "SELECT * FROM KS_BlogType WHERE typeid="& KS.ChkClng(Request("editid"))
	Rs.Open SQL,Conn,1,3
	newclassid = Rs("typeid")
	parentid = Rs("parentid")
	iparentid = Rs("parentid")
	ParentStr = Rs("ParentStr")
	depth = Rs("depth")
	child = Rs("child")
	rootid = Rs("rootid")
	
	'判断所指定的分类是否其下属分类
	If ParentID=0 Then
		If CLng(Request("class"))<>0 Then
		Set trs=Conn.Execute("SELECT rootid FROM KS_BlogType WHERE typeid="&KS.ChkClng(Request("class")))
		If rootid=trs(0) Then
			Call KS.AlertHistory("您不能指定该博文的下属分类作为所属分类",-1)
			Exit Sub
		End If

		End If
	Else
		Set trs=Conn.Execute("SELECT typeid FROM KS_BlogType WHERE ParentStr like '%"&ParentStr&","&newclassid&"%' And typeid="&KS.ChkClng(Request("class")))
		If Not (trs.EOF And trs.BOF) Then
			Call KS.AlertHistory("您不能指定该博文的下属分类作为所属分类",-1)
			Exit Sub
		End If
	End If
	If parentid = 0 Then
		parentid = Rs("typeid")
		iparentid=0
	End If
	Rs("parentstr")=nParentStr & rs("typeid") & ","
	Rs("typename") = Trim(Request.Form("classname"))
	Rs("parentid") = KS.ChkClng(Request.Form("class"))
	Rs("readme") =Trim( Request("readme"))
	Rs("LogNum") = KS.ChkClng(Request.Form("LogNum"))
	Rs.Update 
	Rs.Close
	Set Rs=nothing
	
	CheckAndFixClass 0,1
	Call KS.Alert("恭喜您！分类修改成功!","space/KS.Spacelog.asp?action=class")
End Sub

Sub CheckAndFixClass(ParentID,orders)
	Dim Rs,Child,ParentStr
	If ParentID=0 Then
		Conn.Execute("UPDATE KS_BlogType Set Depth=0 WHERE ParentID=0")
	End If
	Set Rs=Conn.Execute("SELECT typeid,rootid,ParentStr,Depth FROM KS_BlogType WHERE ParentID="&ParentID&" ORDER BY rootid,orderid")
	Do while Not Rs.EOF
		Conn.Execute "UPDATE KS_BlogType Set Depth="&Rs(3)+1&",rootid="&Rs(1)&" WHERE ParentID="&Rs(0)&"",Child
		Conn.Execute("UPDATE KS_BlogType Set Child="&Child&",orderid="&orders&" WHERE typeid="&Rs(0)&"")
		orders=orders+1
		CheckAndFixClass Rs(0),orders
		Rs.MoveNext
	Loop
	Set Rs=Nothing
End Sub


Sub delClass()
	Dim Rs,SQL,i
	Dim ChildStr,nChildStr
	Dim Rss,Rsc
	On Error Resume Next
	Set Rs = Conn.Execute("SELECT ParentStr,child,depth,parentid FROM KS_BlogType WHERE typeid=" & KS.ChkClng(Request("editid")))
	If Not (Rs.EOF And Rs.BOF) Then
		If Rs(1) > 0 Then
			Call KS.AlertHistory("该分类含有下属分类，请删除其下属分类后再进行删除本分类的操作!",-1)
			Exit Sub
		End If

		If Rs(2) > 0 Then
			Conn.Execute ("UPDATE KS_BlogType Set child=child-1 WHERE typeid in (" & Rs(0) & ")")
		End If
		For i = 0 To Ubound(AllPostTable)
			SQL = "DELETE FROM KS_BlogInfo WHERE typeid=" & KS.ChkClng(Request("editid"))
			Conn.Execute(SQL)
		Next
		Conn.Execute("DELETE FROM KS_BlogType WHERE typeid=" & KS.ChkClng(Request("editid")))
		
	End If
	Set Rs = Nothing
	Conn.Execute("UPDATE KS_BlogType Set child=0 WHERE child<0")
	CheckAndFixClass 0,1
	UpdateClassTotal
	ks.die "<script>alert('恭喜您！分类删除成功。');location.href='?action=class';</script>"
End Sub

Sub ClassOrders()
%>
<br>
<table border="0" cellspacing="1" cellpadding="3" align="center"  class="Ctable">
	<tr> 
	<th class="sort" colspan=2>博文一级分类重新排序修改(请在相应分类的排序表单内输入相应的排列序号)</th>
	</tr>
	<tr>
<%
	Dim Rs,SQL,i
	Set Rs=Server.CreateObject("ADODB.Recordset")
	SQL="SELECT * FROM KS_BlogType WHERE ParentID=0 ORDER BY rootid"
	Rs.Open SQL,Conn,1,1
	If Rs.Eof And Rs.Bof Then
		Response.Write "还没有相应的博客分类。"
	Else
		Do While Not Rs.Eof
		Response.Write "<form action=""?action=class&flag=updatorders"" method=""post""><tr class='tdbg'>"
		Response.Write "<td align=""right"" class=""clefttitle"">" & rs("typeName") & "</td><td><input type=""text"" name=""OrderID"" size=""4"" value="""&rs("rootid")&"""><input type=""hidden"" name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=""submit"" name=""Submit"" value=""修改"" class=""button""></td></tr></form>"
		Rs.Movenext
		Loop
%>
</table>
<%
	End If
	Rs.Close
	Set Rs=Nothing
%>
	</td>
	</tr>
</table>
<%
End Sub

Sub updatorders()
	Dim cID,OrderID,Rs
	cID = Replace(Request.Form("cID"),"'","")
	OrderID = Replace(Request.Form("OrderID"),"'","")
	Set Rs = Conn.Execute("SELECT typeid FROM KS_BlogType WHERE rootid="&orderid)
	If Rs.EOF And Rs.BOF Then
		Conn.Execute("UPDATE KS_BlogType SET rootid="&OrderID&" WHERE rootid="&cID)
		Call KS.AlertHintScript("设置成功!")
	Else
		Call KS.AlertHistory("请不要和其他分类设置相同的序号",-1)
		Response.End
	End If
End Sub
Sub ClassTotal()
 UpdateClassTotal()
 Call KS.AlertHistory("恭喜,分类下的博文数统计成功!",-1)
End Sub

Sub UpdateClassTotal()
 Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
 Rs.Open "Select * From KS_blogtype Order By Rootid,orderid",conn,1,3
 do while not rs.Eof 
   Rs("lognum")=Conn.Execute("select count(1) From KS_bloginfo WHERE typeid in (SELECT typeid FROM KS_blogtype WHERE ','+parentstr+'' like '%,"&rs("typeid")&",%')")(0)
   Rs.Update
  Rs.MoveNext
 Loop
 Rs.Close
 Set RS=Nothing
End Sub

End Class

%> 

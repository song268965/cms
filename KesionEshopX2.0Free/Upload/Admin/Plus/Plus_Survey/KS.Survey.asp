<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"--> 
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS X1.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_SurverCls
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_SurverCls
        Private KS,KSCls,I,TypeFlag,ItemStr
		Private MaxPerPage,CurrentPage,TotalPut,ID,RS
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		  With KS
		   If Not KS.ReturnPowerResult(0, "Survey0000") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 Exit Sub
		   End If
		   CurrentPage=KS.ChkClng(request("Page"))
		   if CurrentPage<=0 then CurrentPage=1
		   TypeFlag=KS.ChkClng(KS.S("TypeFlag"))
		    ItemStr="项目"
		   
		    .echo"<!DOCTYPE html><html>"
			.echo"<title>项目设置</title>"
			.echo"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo"<script src=""../../../ks_inc/jQuery.js"" language=""JavaScript""></script>"
			.echo"<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
			.echo"<script src=""../../../ks_inc/DatePicker/WdatePicker.js""></script>"
			.echo "<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo"</head>"
			.echo"<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"

			.echo"<ul id='menu_top'>"
			if KS.G("Action")="SurveyST" then
			.echo"<li class='parent' onclick=""location.href='KS.Survey.asp?action=AddST&TypeFlag=试题&SurveyID="&  KS.G("ID") &"';$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=Go&OpStr=试题  >> <font color=red>添加试题</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加试题</span></li>"
			else
				if KS.G("Action")<>"EditST" then
				.echo"<li class='parent' onclick=""location.href='KS.Survey.asp?action=Add&TypeFlag=" &TypeFlag&"';$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=Go&OpStr=" & ItemStr & " >> <font color=red>添加项目</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add'></i>添加项目</span></li>"
				end if
			end if
			if KS.G("Action")="EditST" then
			.echo"<li class='parent' onclick='history.go(-1)'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>返回</span></li>"
			end if
             if KS.G("Action")<>"EditST" then
				 If KS.G("Action")="" Then
				.echo"<li class='parent' disabled"
				 Else
				.echo"<li class='parent'"
				 End If
				.echo" onclick='location.href=""KS.Survey.asp?typeflag=" & typeflag & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon mainer'></i>管理首页</span></li>"
			end if	
				.echo"</ul>"
				.echo "<div class='pageCont2'>"
		  Select Case KS.G("Action")
		   Case "SetFormParam" Call SetFormParam() 
		   Case "Edit","Add"  Call FormManage()
		   Case "EditST","AddST"  Call FormManageST()
		   Case "EditSave" Call FormSave()
		   Case "EditSaveST" Call FormSaveST()
		   Case "Del","DelST" Call ProjectDel()
		   Case "GetCode" Call GetCode()
		   Case "SurveyST" Call SurveySTMain()
		   Case Else Call Main()
		  End Select
		  End With
		End Sub
 
		Sub Main()
		   With KS
			.echo"<script>"
			.echo"$(document).ready(function(){"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button1').attr('disabled',true);"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button2').attr('disabled',true);"
			.echo"});</script>"
			
			.echo "<div class='tabTitle'>问卷项目管理</div>"
			.echo("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_Survey Order By ID",conn,1,1
		    .echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo"  <td width='50' align=center>ID</td><td align=center>项目名称</td><td align=center>时间</td><td align=center>投票限制</td><td align=center>↓操作</td>"
			.echo"</tr>"
			If RS.Eof And RS.Bof Then
			 .echo "<tr><td class='splittd' align='center' height='40' colspan=10>还没有添加项目！</td></tr>"
			Else
			            totalPut = RS.RecordCount
						If CurrentPage < 1 Then	CurrentPage = 1
			            If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
						End If
			
			            dim i:i=0
					  Do While Not RS.Eof 
						.echo"<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
						.echo"<td align=center class='splittd' style='height:35px;'>" & RS("ID")&"</td>"
						.echo"<td class='splittd' style='height:35px;'>" & RS("ProjectName") 
						.echo " &nbsp;["&Conn.Execute("SELECT COUNT(ID) FROM KS_SurveyST where SurveyID = "&RS("ID"))(0) &"]"
						.echo"</td>"
						.echo"<td align=center class='splittd' style='height:35px;'>" 
						  If RS("TimeLimit")="1" Then 
							.echo"开始时间:"&RS("StartDate") &"<br/>"
							.echo"结束时间:"&RS("ExpiredDate")
						  Else 
						  .echo"<font color=red>无限时间</font>"
						  end if
						.echo"</td>"
						dim OnlyUser:if RS("OnlyUser")=0 then OnlyUser="无限制" else OnlyUser="只允许会员"
						.echo"<td align=center class='splittd' style='height:35px;'>" & OnlyUser&"</td>"
						.echo"<td align=center class='splittd' style='height:35px;'>"
						.echo"<a href='?action=SurveyST&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=GoSave&OpStr=问卷题目管理 >> <font color=red>问卷题目</font>';"" class='setA'>问卷题目管理</a>｜"
						.echo"<a href='?typeflag=" & typeflag &"&action=Edit&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=GoSave&OpStr=子系统 >> <font color=red>" & ItemStr & "</font>';"" class='setA'>修改</a>｜"
						 .echo"<a href='?typeflag=" & typeflag &"&action=Del&ID=" & rs("ID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))' class='setA'>删除</a>|<a href='../../../Survey/Survey.asp?id=" & rs("id") & "' target='_blank' class='setA'>预览</a>"
						.echo"</td></tr>"
						i=i+1
						if i>=maxperpage then exit do
						RS.MoveNext 
					  Loop
			End If
		    .echo"</table>"
			.echo  KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo"</div>"
			.echo"</div>"
		   RS.Close:Set RS=Nothing
		    .echo"</body>"
			.echo"</html>"
		  End With
		End Sub
		
		Sub SurveySTMain()
		   With KS
		   dim lxstr,SurveyID:SurveyID=KS.ChkClng(KS.S("ID"))
		   IF SurveyID<>0 then
			.echo"<script>"
			.echo"$(document).ready(function(){"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button1').attr('disabled',true);"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button2').attr('disabled',true);"
			.echo"});</script>"
			.echo"<div class='tabTitle'>问卷题目管理</div>"
			.echo("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_SurveyST where  SurveyID="& SurveyID & " Order By SurveyOrder",conn,1,1
		    .echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo"  <td width='50' align=center>排序</td><td align=center>试题名称</td><td align=center>试题类型</td><td align=center>添加时间</td><td align=center>↓操作</td>"
			.echo"</tr>"
		 if rs.eof and rs.bof then
		    .echo "<tr><td class='splittd' colspan=10 align='center' height='40'>还没有添加试题!</td></tr>"
		 else
		  Do While Not RS.Eof 
		    .echo"<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.echo"<td align=center class='splittd' style='height:35px;'>" & RS("SurveyOrder")&"</td>"
			.echo"<td class='splittd' style='height:35px;'>" & RS("SurveySTName") &"</td>"
			if RS("lx")=0 then
				lxstr="单选"
			elseif RS("lx")=1 then
				lxstr="多选"
			else
				lxstr="填空"
			end if
			.echo"<td align=center class='splittd' style='height:35px;'>" & lxstr &"</td>"
			.echo"<td align=center class='splittd' style='height:35px;'>"& RS("addDate")&"</td>"
			.echo"<td align=center class='splittd' style='height:35px;'>"

			.echo"<a href='?typeflag=" & typeflag &"&action=EditST&ID=" & rs("ID") & "&SurveyID="& SurveyID &"' onclick=""$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=GoSave&OpStr=多问卷调查 >> <font color=red>问卷试题管理</font>';"" class='setA'>修改</a>｜"
			 .echo"<a href='?typeflag=" & typeflag &"&action=DelST&ID=" & rs("ID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))' class='setA'>删除</a>"			
			.echo"</td></tr>"
			RS.MoveNext 
		  Loop
		 end if
		    .echo"</table>"
			.echo"</div>"
		   RS.Close:Set RS=Nothing
		    .echo"</body>"
			.echo"</html>"
		  End if	
		  End With
		End Sub
		
		Sub ProjectDel()
		  on error resume next
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Conn.BeginTrans
		  if KS.G("Action")="DelST" then
			Conn.Execute("Delete From KS_SurveyST Where ID=" & ID)
			Conn.Execute("Delete From KS_SurveyItem Where SurveySTID=" & ID)
			Conn.Execute("Delete From KS_SurveyResult Where SurveySTID=" & ID)
		  else
		  	Conn.Execute("Delete From KS_Survey Where ID=" & ID)
			Conn.Execute("Delete From KS_SurveyST Where SurveyID=" & ID)
			Conn.Execute("Delete From KS_SurveyItem Where SurveyID=" & ID)
			Conn.Execute("Delete From KS_SurveyResult Where SurveyID=" & ID)
		  end if
		  If Err<>0 Then
		   Conn.RollBackTrans
		  Else
		   Conn.CommitTrans
		  End If
		  KS.Echo ("<script>alert('删除成功');location.href='" & Request.ServerVariables("HTTP_REFERER") &"';</script>")
		End Sub
		
        		
		Sub GetCode()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Survey Where Status=1 and TypeFlag=" & TypeFlag & " order by ID asc",conn,1,1
		   With KS
		  	.echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo" <td align=center colspan=6>各" & ItemStr & "项目的前台调用代码</td>"
			.echo"</tr>"

		  Do While Not RS.Eof
			.echo"<tr height='25' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"

			.echo"<td width='50'></td><td width='140'><img src='../../images/37.gif'>&nbsp;<b>" & RS("ProjectName") & "</b></td><td>"
			If TypeFlag=1 Then
			.echo "内容页发表点评标签{=GetWriteComments(" & rs(0) & ")}<br/>内容页显示点评标签{=GetShowComments(" & rs(0) & ")}"
			.echo "</td><td></td><td></td>"
			Else
			.echo "<textarea style='width:500px;height:50px' name='s" & rs(0) & "'>&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(2) & "/plus/mood.asp?id=" & rs("id") & "&c_id={$InfoID}&M_id={$ChannelID}&quot;&gt;&lt;/script&gt;</textarea>"
			.echo "</td><td><input class=""button"" onClick=""jm_cc('s" & rs(0) & "')"" type=""button"" value=""复制到剪贴板"" name=""button""></td><td></td>"
			End If
			
			.echo"</tr>"
			.echo"<tr><td colspan=6 background='../../images/line.gif'></td></tr>"
		    RS.MoveNext
		  Loop
		   .echo"</table>"
		  End With
		  RS.Close:Set RS=Nothing
		  %>
		 
		   <script>
			function jm_cc(ob)
			{
				var obj=MM_findObj(ob); 
				if (obj) 
				{
					obj.select();js=obj.createTextRange();js.execCommand("Copy");}
					alert('复制成功，粘贴到你要调用的html代码里即可!');
				}
			function MM_findObj(n, d) { //v4.0
			  var p,i,x;
			  if(!d) d=document;
			  if((p=n.indexOf("?"))>0&&parent.frames.length)
			   {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
			   }
			  if(!(x=d[n])&&d.all) x=d.all[n];
			  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
  </script>
		  <%
		End Sub
		
		Sub SetFormParam()
		   With Response
			   Dim ID:ID=KS.ChkClng(KS.G("ID"))
			   If ID=0 Then .Redirect "?typeflag=" & typeflag : Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_Survey Where ID=" & ID,Conn,1,3
			   If RS.Eof Then
				 RS.Close:Set RS=Nothing
				.Redirect "?": Exit Sub
			   End If
		     If KS.G("Flag")="FormOpenOrClose" Then
			   If RS("Status")=1 Then 
					RS("Status")=0 
			   Else 
			    RS("Status")=1
			   end if
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing
			 ks.echo"<script>location.href='?typeflag=" & typeflag & "';</script>"
		   End With
		End Sub
		
		Sub FormManage()
		Dim TimeLimit,AllowGroupID,useronce,onlyuser,ProjectContent
		Dim TempStr,SqlStr, RS, i,MaxLen,Score
		Dim ProjectName,ExpiredDate,StartDate,Status,Descript,TableName,UpLoadDir,TemplateID,ZCJTF,VerifyCodeTF,IsRewrite,IsVerify,Template_a,Template_b,UserCk
		

		Dim ID:ID = KS.ChkClng(KS.G("ID"))
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select top 1 * from KS_Survey Where ID=" & ID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			ProjectName    = RS("ProjectName")
			ProjectContent = RS("ProjectContent")
			StartDate    = RS("StartDate")
			TimeLimit    = RS("TimeLimit")
			ExpiredDate  = RS("ExpiredDate")
			TimeLimit    = RS("TimeLimit")
            AllowGroupID = RS("AllowGroupID")
			useronce     = RS("useronce")
			onlyuser     = RS("onlyuser")
			Template_a   = RS("Template_a")
			Template_b   = RS("Template_b")
			UserCk       = RS("UserCk")
			Score        = RS("Score")
		Else
		      Status=1:TimeLimit = 0:StartDate=Now():ExpiredDate=Now()+10:AllowGroupID="":useronce=0:onlyuser=0:ZCJTF=0:VerifyCodeTF=0 : IsVerify=0:MaxLen=100
			  Score=0
			  Template_a="{@TemplateDir}多问卷调查/调查页面.html"
			  Template_b="{@TemplateDir}多问卷调查/调查结果.html"
		End If
		%>
		<script>
		 function CheckForm()
		 {
		  if ($("#ProjectName").val()=="")
		  {
		   $("#ProjectName").focus();
		   alert('请输入项目名称');
		   return false;
		  }
		  
		  $("#myform").submit();
		 }
		 
		 function changedate()
		 {
		   val=$("input[name=TimeLimit]:checked").val();
		   if (val==1){
		    $("#BeginDate").show();
		    $("#EndDate").show();		
		   }
		   else{
		    $("#BeginDate").hide();
		    $("#EndDate").hide();		
		   }
		 }
	
		</script>
		
		
		<%
		With KS
		.echo"<script src=""../../images/pannel/tabpane.js""></script>" & _
		"<link href=""../../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		EchoUeditorHead() &_
		"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"&_
		"  <tr>"&_
		"	<td height='25' class='sort'>" & ItemStr  &"管理</td>"&_
		" </tr>"&_
		" <tr><td height=5></td></tr>"&_
		"</table>" & _
			
		"<div class=tab-page id=Formpanel>"& _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""Formpanel"" ), 1 )"& _
        " </SCRIPT>"& _
             
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>基本信息</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" 
%>
<form name="myform" id="myform" method="post" action="KS.Survey.asp?Action=EditSave&ID=<%=ID%>">
		  <input type="hidden" name="typeflag" value="<%=KS.ChkClng(KS.S("TypeFlag"))%>">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>项目名称：</strong></div></td>      
			<td height="30"> <input name="ProjectName" id="ProjectName" class="textbox" type="text" value="<%=ProjectName%>" size="70"> <span>*</span></td> 
		</tr>
		

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>项目说明：</strong></div><br><font color=green></font></td>      
			<td height="30"> 
			<%
			KS.Echo EchoEditor("ProjectContent",ProjectContent,"Basic","96%","200px")
			%>
		</tr>
		</table>
		</div>
		 <div class=tab-page id="formset">
		  <H2 class=tab>选项设置</H2>
			<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "formset" ) );
			</SCRIPT>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>启用时间限制：</strong></div></td>      
			<td height="30"> 
			
			<%
			.echo "<input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""1"" "
		If TimeLimit = 1 Then .echo(" checked")
		.echo">"
		.echo"启用"
		.echo"  <input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""0"" "
		If TimeLimit = 0 Then .echo(" checked")
		.echo">"
		.echo"不启用"
		
			%>
			</td> 
		</tr>

		<tr ID="BeginDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">     
		<td height="30" class="clefttitle"align="right"><div><strong>生效时间：</strong></div></td>     
		<td height="30"><input name="StartDate" onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  id='StartDate' class="textbox" type="text" value="<%=StartDate%>" size="30"><br><font color=#ff0000>日期格式：0000-00-00 00:00:00</font></td>   
		</tr> 
		
		<tr ID="EndDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>失效时间：</strong></div></td>      
			<td height="30"> <input name="ExpiredDate" onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  id="ExpiredDate" class="textbox" type="text" value="<%=ExpiredDate%>" size="30"><br><font color=#ff0000>日期格式：0000-00-00 00:00:00</font></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>前台查看结果：</strong></div></td>      
			<td height="30"> 
			<%
			.echo "<input type=""radio"" name=""UserCk"" value=""0"" "
		If UserCk = 0 Then .echo(" checked")
		.echo">"
		.echo"开启查看"
		.echo"  <input type=""radio"" name=""UserCk"" value=""1"" "
		If UserCk = 1 Then .echo(" checked")
		.echo">"
		.echo"关闭查看"
		
			%>
			</td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>只允许会员投票：</strong></div></td>      
			<td height="30"> 
			
			<%
			.echo "<input type=""radio"" name=""onlyuser"" value=""1"" "
		If onlyuser = 1 Then .echo(" checked")
		.echo">"
		.echo"是"
		.echo"  <input type=""radio"" name=""onlyuser"" value=""0"" "
		If onlyuser = 0 Then .echo(" checked")
		.echo">"
		.echo"不是"
		
			%>
			</td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>每个会员只允许投票一次：</strong></div></td>      
			<td height="30"> 
			
			<%
			.echo "<input type=""radio"" name=""useronce"" value=""1"" "
		If useronce = 1 Then .echo(" checked")
		.echo">"
		.echo"是"
		.echo"  <input type=""radio"" name=""useronce"" value=""0"" "
		If useronce = 0 Then .echo(" checked")
		.echo">"
		.echo"不是"
		
			%>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>参加问卷的会员：</strong></div></td>      
			<td height="30"> 
			
		奖励积分<input type="text" name="score" value="<%=score%>" class="textbox" style="width:50px;text-align:center"/>分  <span class='tips'>设置为“0”将不奖励。允许会员提交多次的调查，每个会员仅奖励一次。</span>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>调查模板：</strong></div></td>      
			<td height="30"><input type="text" name="Template_a" id="Template_a" class="textbox" value="<%=Template_a%>" size="34" />&nbsp;<%=KSCls.Get_KS_T_C("$('#Template_a')[0]")%></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>调查结果模板</strong></div></td>      
			<td height="30"><input type="text" name="Template_b" id="Template_b" class="textbox" value="<%=Template_b%>" size="34" />&nbsp;<%=KSCls.Get_KS_T_C("$('#Template_b')[0]")%></td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>用户组限制：</strong></div><font color=#ff0000>不限制，请不要选</font></td>      
			<td height="30"><%=KS.GetUserGroup_CheckBox("AllowGroupID",AllowGroupID,5)%> </td> 
		</tr>
			</table>
        </div>

		</form>
		<script>changedate();</script>
		<%
		.echo"</div>"
		End With
		End Sub
		
		Sub FormManageST()
		Dim TimeLimit,AllowGroupID,useronce,onlyuser,ProjectContent
		Dim TempStr,SqlStr, RS, i,MaxLen
		Dim SurveyID,SurveySTName,Content,SurveyOrder,chstr,lx
		

		Dim ID:ID = KS.ChkClng(KS.G("ID"))
	'	On Error Resume Next
	   If KS.G("Action")="EditST" Then
			SqlStr = "select top 1 * from KS_SurveyST Where ID=" & ID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			SurveyID    = RS("SurveyID")
			SurveySTName = RS("SurveySTName")
			Content   = RS("Content")
			SurveyOrder   = RS("SurveyOrder")
			lx=RS("lx")
			Rs.Close
		Else
		    SurveyID  = KS.ChkClng(KS.G("SurveyID"))
			SurveySTName = ""
			Content   = ""
			SurveyOrder   = KS.ChkClng(conn.execute("select max(SurveyOrder) from KS_SurveyST Where SurveyID=" & SurveyID)(0))+1
			lx=0
		End If
		
		With KS
		.echo "<form name=""myform"" id=""myform"" method=""post"" action=""KS.Survey.asp?Action=EditSaveST&ID=" & ID & """>"   
		%>
		<input type="hidden" name="ID" value="<%=ID%>">
		<input type="hidden" name="typeflag" value="<%=KS.ChkClng(KS.S("TypeFlag"))%>">
		<input type="hidden" name="SurveyID" value="<%=SurveyID%>">  
		<%          
		.echo "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
%>
		<script>
		 function CheckForm()
		 {
		  if ($("#SurveySTName").val()=="")
		  {
		   $("#SurveySTName").focus();
		   alert('请输入项目名称');
		   return false;
		  }
		  
		  $("#myform").submit();
		 }
	
		</script>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>主题名称：</strong></div></td>      
			<td height="30"> <input name="SurveySTName" id="SurveySTName"  class="textbox" type="text" value="<%=SurveySTName%>" size="30"> 如：你对本站的哪些栏目较感兴趣!</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>试题简介：</strong></div></td>      
			<td height="30"> 
			 	<textarea name="Content" cols="50" rows="6" class="textbox"><%=Content%></textarea>
			 </td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>试题类型：</strong></div></td>      
			<td height="30"> 
			<select name="lx" id="lx"> 
				<option value="0" <%if lx=0 then Response.Write "selected=""selected""" %> >单选</option>
				<option value="1" <%if lx=1 then Response.Write "selected=""selected""" %> >多选</option>
				<option value="2" <%if lx=2 then Response.Write "selected=""selected""" %>>填空</option>
			</select> 
			</td> 
		</tr>
		 <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"> 
							<td width="160" height="30" class="clefttitle"><div align="right"><strong>投票项目：</strong><br /><font color="#FF0000">删除名称设空</font></div></td>
							<td colspan="3" bgcolor="#EEF8FE">
							
							 <table border="0" cellpadding="0" cellspacing="0" style="margin-left:5px;" width="80%">
     
                 <tr>
                  <td colspan="3" height="30px">
							投票扩展数量: 
						  <input name="vote_num" type="text" class="textbox" id="votenum" value="1" size="5" style="text-align:center"> 
						  <input type="button" name="Submit52" value="增加选项" class="button" onclick="javascript:doadd(jQuery('#votenum').val());"> 
							  
							  </td>
							 </tr>
							 <tr bgcolor='#DBEAF5'>
							 <td width='9%' height='20'> <div align='center'>编号</div></td>
							 <td width='65%'> <div align='center'>名称</div></td>
							 <td width='26%'> <div align='center'>其他选项</div></td>
							 </tr>
							 <tr>
							  <td colspan="3" id="addvote">	  
							  
							  <%dim ii,IDstr
							  if ID<>0 then
								Set Rs = Server.CreateObject("adodb.recordset")
								Rs.Open "select * from KS_SurveyItem where SurveySTID="&ID & " Order By SurveyItemOrder" , Conn, 1, 1	
								ii=0
								Do While Not rs.Eof
									ii=ii+1
									if RS("SurveyItemType")=1 then chstr="checked=""checked""" else chstr=""
									tempstr=tempstr & "<tr><td width=9% height=20><div align=center><input type=hidden name=id_"& rs("id") &" value=" & rs("id") & ">" & ii & "</div></td><td width='65%'> <div align=center><input class=textbox type=text name=item_"& rs("id") &" size=40 value='" & rs("SurveyItemName") & "'></div></td><td width='26%'> <div align=center><input name=ck_"&rs("id") &"  type=""checkbox"" value='1'   "& chstr &" /></div></td></tr>"&vbcrlf
									IDstr=IDstr&rs("id")&"|"
								rs.MoveNext 
								loop
								Rs.Close
								Set Rs = Nothing
							    end if
								response.write "<table width=100% border=0 cellspacing=1 cellpadding=3>"
								response.write tempstr
								response.write "</table>"
							  %>
							  </td>
							 </tr>
							</table>
							<input name="editnum" type="hidden" id="editnum" value="<%=KS.ChkClng(ii)%>"> 
							<input name="editnumjs" type="hidden" id="editnumjs" value="0"> 
							<input name="IDstr" type="hidden" id="IDstr" value="<%=IDstr%>"> 
							<script type="text/javascript">
    function doadd(num)
    {var i;
    var str="";
    var oldi=0;
    var j=0;
    oldi=parseInt(jQuery('#editnum').val());
    for(i=1;i<=num;i++)
    {
    j=i+oldi;
    str=str+"<tr><td width=9% height=20> <div align=center><input type=hidden name=id_js_"+j+" value=0>"+j+"</div></td><td width=65%> <div align=center><input type=text name=item_js_"+j+" class=textbox size=40></div></td><td width=26%> <div align=center><input name=ck_js_"+j+" type=checkbox value=1 /></div></td></tr>";
    }
    window.addvote.innerHTML+="<table width=100% border=0 cellspacing=1 cellpadding=3>"+str+"</table>";
        jQuery('#editnum').val(j);
		jQuery('#editnumjs').val(parseInt(jQuery('#editnumjs').val())+1);
    }
	<%If ID=0 Then%>
	doadd(2);
	<%end if%>
    </script>
							</td>
						  </tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>试题排序号：</strong></div><br><font color=green></font></td>      
			<td height="30"> 
			<input type="text" name="SurveyOrder" class="textbox" value="<%=SurveyOrder%>" style="text-align:center" size="5">   
			</td> 
		</tr>
		</table>
		<%
		.echo"</form>"
		End With
		End Sub
		
			
	
		
		Sub FormSave()
		    Dim ExpiredDate,StartDate,I,OpName,ID
			ID=KS.ChkClng(KS.G("ID"))
			StartDate=KS.G("StartDate")
			ExpiredDate=KS.G("ExpiredDate")
			If Not IsDate(StartDate) Then Call KS.AlertHistory("生效日期格式不正确",-1):response.end
			If Not IsDate(ExpiredDate) Then Call KS.AlertHistory("失效日期格式不正确",-1):response.end
			If ID=0 and Not Conn.Execute("select top 1 id from KS_Survey where projectname='" & KS.G("ProjectName") &"'").eof then Call KS.AlertHistory("项目名称已存在！",-1):response.end
			Conn.BeginTrans
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Survey Where ID=" & ID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				RS("AddDate")=now
				OpName      = "添加"
			Else
			    OpName="修改"
			End If
				RS("ProjectName")= KS.G("ProjectName")
				RS("ProjectContent")=Request("ProjectContent")
				RS("TimeLimit")   = KS.ChkClng(KS.G("TimeLimit"))
				RS("StartDate")     = startdate
				RS("ExpiredDate")    = ExpiredDate
				RS("useronce") =KS.ChkClng(KS.G("useronce"))
				RS("onlyuser")=KS.ChkClng(KS.G("onlyuser"))
				RS("AllowGroupID")     = KS.G("AllowGroupID")
				RS("Template_a")=KS.G("Template_a")
				RS("Template_b")=KS.G("Template_b")	
				RS("UserCk")=KS.ChkClng(KS.G("UserCk"))
				RS("Score")=KS.ChkClng(KS.G("Score"))
				RS.Update
				RS.Close
				Set RS=Nothing
				
				
				if err<>0 then
					Conn.RollBackTrans
					Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				else
					Conn.CommitTrans
					If ID=0 Then
					  KS.Echo "<script src='../../../ks_inc/jquery.js'></script>"
			           KS.Echo ("<script> if (confirm('恭喜，问卷项目添加成功!继续添加吗?')) {location.href='?Action=Add';}else{location.href='KS.Survey.asp';$(parent.document).find('#BottomFrame')[0].src='Post.asp?ButtonSymbol=Disabled&OpStr=多问卷调查 >> <font color=red>问卷项目管理</font>';}</script>")
					Else
					  KS.Echo "<script src='../../ks_inc/jquery.js'></script>"
					  KS.Echo ("<script>alert('恭喜，问卷项目修改成功!');location.href='KS.Survey.asp';$(parent.document).find('#BottomFrame')[0].src='Post.asp?ButtonSymbol=Disabled&OpStr=多问卷调查 >> <font color=red>问卷项目管理中心</font>';</script>")
					End If
				end if
		End Sub
		
		Sub FormSaveST()
		    Dim ExpiredDate,StartDate,I,OpName,ID,editnumjs,editnum,RS,IDstr
			ID=KS.ChkClng(KS.G("ID"))
			editnum=KS.ChkClng(KS.G("editnum"))
			editnumjs=KS.ChkClng(KS.G("editnumjs"))
			IDstr=KS.G("IDstr")
			IDstr=ks.gottopic(IDstr,Len(IDstr)-1)
			IDstr=Split(IDstr,"|") 
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_SurveyItem Where ID=" & ID,Conn,1,3
			if UBound(IDstr)>=0 then
				for i=0 to UBound(IDstr)
					Conn.Execute ("UPDATE KS_SurveyItem SET SurveyItemName ='"& KS.G("item_"&IDstr(i)) &"',SurveyItemType ="& KS.ChkClng(KS.G("ck_"&IDstr(i))) &"  WHERE ID="& KS.ChkClng(IDstr(i)))	
					if KS.G("item_"&IDstr(i))="" then Conn.Execute("Delete From KS_SurveyItem Where ID=" & KS.ChkClng(IDstr(i)))
				next
				if editnumjs<>0 then
					for i=1 to editnumjs
						if KS.G("item_js_"&i+UBound(IDstr)+1)<>"" then
							RS.AddNew
							RS("SurveyID")= KS.ChkClng(KS.G("SurveyID"))
							RS("SurveySTID")=ID
							RS("SurveyItemName")=KS.G("item_js_"&i+UBound(IDstr)+1)
							RS("SurveyItemType")=KS.ChkClng(KS.G("ck_js_"&i+UBound(IDstr)+1))
							RS("SurveyItemOrder")=i+UBound(IDstr)+1
							RS.Update
						end if
					next
				end if	
			end if
			RS.Close
			'StartDate=KS.G("StartDate")
			'ExpiredDate=KS.G("ExpiredDate")
			'If ID=0 and Not Conn.Execute("select top 1 id from KS_SurveyST where projectname='" & KS.G("ProjectName") &"'").eof then Call KS.AlertHistory("项目名称已存在！",-1):response.end
			Conn.BeginTrans
		    Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_SurveyST Where ID=" & ID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				OpName      = "添加"
				RS("AddDate")=now()
			Else
			    OpName="修改"
			End If
				RS("SurveySTName")= KS.G("SurveySTName")
				RS("SurveyID")=KS.ChkClng(KS.G("SurveyID"))
				RS("SurveyOrder")   = KS.ChkClng(KS.G("SurveyOrder"))
				RS("Content")=KS.G("Content")
				RS("lx")=KS.ChkClng(KS.G("lx"))
				RS.Update
				Dim SurveyID:SurveyID=RS("ID")
				RS.Close
				if UBound(IDstr)<0 then
					Set RS=Server.CreateObject("ADODB.RECORDSET")
				    RS.Open "Select top 1 * From KS_SurveyItem Where ID=" & ID,Conn,1,3
					for i=1 to editnum	
						if KS.G("item_js_"&i)<>"" then
							RS.AddNew
							RS("SurveyID")= KS.ChkClng(KS.G("SurveyID"))
							RS("SurveySTID")=SurveyID
							RS("SurveyItemName")=KS.G("item_js_"&i)
							RS("SurveyItemType")=KS.ChkClng(KS.G("ck_js_"&i))
							RS("SurveyItemOrder")=i
							RS.Update
						end if
					next
					RS.Close	
				end if	
				Set RS=Nothing
				if err<>0 then
					Conn.RollBackTrans
					Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				else
					Conn.CommitTrans
					If ID=0 Then
					 ks.die "<script>if (confirm('恭喜，试题添加成功，继续添加吗？')){$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=Go&OpStr=试题  >> <font color=red>添加试题</font>';location.href='KS.Survey.asp?action=AddST&TypeFlag=试题&SurveyID=" & KS.ChkClng(KS.G("SurveyID")) & "';}else{location.href='KS.Survey.asp?Action=SurveyST&ID="& KS.ChkClng(KS.G("SurveyID")) &"';}</script>"
					Else
			          KS.Echo ("<script>alert('试题修改成功!');location.href='KS.Survey.asp?Action=SurveyST&ID="& KS.ChkClng(KS.G("SurveyID")) &"';</script>")

					End If
				end if
		End Sub
End Class
%> 

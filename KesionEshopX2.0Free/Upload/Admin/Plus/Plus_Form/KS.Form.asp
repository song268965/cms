<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_Form
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Form
        Private KS,KSCls,I
		Private MaxPerPage,CurrentPage,TotalPut,ID,RS
		Private IConnStr,IConn
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		  With Response
		   If Not KS.ReturnPowerResult(0, "KSMS10006") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
		   End If
		   If KS.G("Action")="createtemplate" Then
		     Call AutoTemplate()
			 response.end
		   ElseIf KS.G("Action")="export" Then
		     Call export()
			 response.End()
		   End If
		    .Write "<!DOCTYPE html><html>"
			.Write "<title>表单管理</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.Write EchoUeditorHead()
		    %>
            <script>
			function FormField(id)
			{ 
				if (id==''){
				 alert('请选择要编辑的表单!');
				}else {
				location="KS.FormField.asp?ItemID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=自定义表单 >> <font color=red>表单字段管理</font>&ButtonSymbol=Disabled';
				}
			}
			function ShowResult(id,f)
			{ 
			    if (f==1){
				location="KS.Form.asp?ItemID="+id+"&action=result";
				}else{
				location="KS.Form.asp?ItemID="+id+"&action=resulthp";
				}
				window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=自定义表单 >> <font color=red>查看提交结果</font>&ButtonSymbol=Disabled';
			}
			function AddRecord(id)
			{ 
				location="KS.Form.asp?action=modifyinfo&FormID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=自定义表单 >> <font color=red>添加记录</font>&ButtonSymbol=Disabled';
			}
			</script>
            <%
			if  KS.G("Action")<>"print_bm" then
				.Write "<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			end if
			.Write "</head>"
			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
		  If KS.G("Action")="replay" Then
		    Call Replay()
			Response.End()
		  End If
		   if  KS.G("Action")<>"print_bm" then
			.Write "<ul id='menu_top'>"
			 If KS.ReturnPowerResult(0, "KSMS100061") Then
			 .Write "<li class='parent' onclick=""location.href='KS.Form.asp?action=Add';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Go&OpStr=自定义表单 >> <font color=red>添加表单</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加表单</span></li>"
			.Write "<li class='parent' onclick='location.href=""KS.Form.asp?action=total""'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon move'></i>调用代码</span></li>"
			End If
             If KS.G("Action")="" Then
			.Write "<li class='parent' disabled"
		     Else
			.Write "<li class='parent'"
			 End If
			.Write " onclick='location.href=""KS.Form.asp"";'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon mainer '></i>管理首页</span></li>"
			.Write "</ul>"
			.Write "<div class='pageCont2'>"
		  end if
		  Select Case KS.G("Action")
		   Case "result"  Call SubmitResult()
		   Case "resulthp"  Call SubmitResultHP()
		   Case "setstatus" Call setstatus()
		   Case "delinfo"  Call DelInfo()
		   Case "SetFormParam" Call SetFormParam() 
		   Case "Edit","Add"  Call FormManage()
		   Case "EditSave" Call FormSave()
		   Case "Del" Call FormDel()
		   Case "total" Call Total()
		   Case "template" Call FormTemplate()
		   Case "TemplateSave" Call TemplateSave()
		   Case "view" Call FormView()
		   Case "replaysave" Call ReplaySave()
		   case "Import" Import()
		   Case "ImportNext" importNext()
		   Case "ImportNext2" importNext2()
		   Case "modifyinfo" modifyinfo()
		   Case "DoResultSave" DoResultSave()
		   Case "print_bm" print_bm()
		   Case Else Call Main()
		  End Select
		  End With
		End Sub
 
		Sub Main()
		   With Response
			.Write "<script>"
			.Write "$(document).ready(function(){"
			.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
			.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
			.Write "});</script>"
			.Write "<div class='tabTitle'>表单管理</div>"
			.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 Dim param:param=" where 1=1"
			 If KS.C("SuperTF")<>"1" Then
			    param=param & " and (adminuserlist='' or ','+adminuserlist+',' like '%," & KS.C("AdminName") & ",%')"
			 end If
			 RS.Open "Select * From KS_Form "& param & " Order By ID",conn,1,1
		    .Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write "  <td width='50' align=center>ID</td><td align=center>表单名称</td><td align=center>有效期</td><td align=center>记录</td><td align=center>状态</td><td align=center>↓管理操作</td>"
			.Write "</tr>"
			If RS.Eof And RS.BOf Then
		    .Write "<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.Write "<td align=center colspan='10' class='splittd'>还没有添加任何表单项目！</td></tr>"
			Else
		  Do While Not RS.Eof 
		    .Write "<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.Write "<td align=center class='splittd'>" & RS("ID")&"</td>"
			.Write "<td align=center class='splittd'><a title=""查看表单[" & rs("formname") & "]记录列表"" href='../../../plus/form/index.asp?id=" & rs("id") & "' target='_blank'>" & RS("FormName") &"</a></td>"
			.Write "<td align='center' class='splittd'>" & RS("StartDate") & "<br/>至<br/>" & RS("ExpiredDate") & "</td>"
			.Write "<td align=center class='splittd'><font color=red>" & conn.execute("select count(*) from " & rs("tablename"))(0) & "</font> 条</td>"
			.Write "<td align=center  class='splittd'>" 
			  If RS("Status")="1" Then .Write "正常" Else .Write "<font color=red>锁定</font>"
			.Write "</td>"
			.Write "<td width='330' class='splittd' style=""text-align:left"">"
			
			If KS.C("SuperTF")="1" Then
				.Write "<strong>项目管理:</strong> <a href='#' onClick=""FormField(" & rs("ID") & ");"">字段管理</a>｜"
				.Write "<a href='KS.Form.asp?ItemID=" & rs("ID") & "&action=template'   onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=自定义表单 >> <font color=red>创建模板</font>';"">创建模板</a>｜"
				.Write "<a href='KS.Form.asp?ItemID=" & rs("ID") & "&action=view'>预览</a>｜"
				.Write "<a href='?action=Edit&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=自定义表单 >> <font color=red>修改表单</font>';"">修改</a>｜"
				 .Write "<a href='?action=Del&ID=" & rs("ID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))'>删除</a>｜"
							 
				 If RS("Status")="1" Then .Write "<a href='?Action=SetFormParam&Flag=FormOpenOrClose&ID=" & RS("ID") & "'>锁定</a>" Else .Write "<a href='?Action=SetFormParam&Flag=FormOpenOrClose&ID=" & RS("ID") & "'>开启</a>"
				 .Write "<br/>"
			End If
			.Write "<strong>记录管理:</strong> <a href=""javascript:ShowResult(" & rs("id") &",0);"">横排查看</a>｜<a href=""javascript:ShowResult(" & rs("id") &",1);"">竖排查看</a> ｜<a href=""javascript:AddRecord(" & rs("ID") & ");"">添加记录</a>｜<a href='?Action=Import&ID=" & RS("ID") & "'>批量导入</a></td></tr>"
			RS.MoveNext 
		  Loop
		  End If
		    .Write "</table>"
			.Write "</div>"
			.Write "</div>"
		   RS.Close:Set RS=Nothing
		    .Write "</body>"
			.Write "</html>"
		  End With
		End Sub
		
		Sub FormDel()
		  If KS.C("SuperTF")<>"1" Then
					 Call KS.ReturnErr(1, "")
		  			 Response.End
		  End If
		  on error resume next
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Conn.BeginTrans
		  Dim TableName:TableName=Conn.Execute("select tablename from ks_form where id=" & ID)(0)
		  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1016 and infoid=" & ID)
		  Conn.Execute("Drop Table " & TableName)
		  Conn.Execute("Delete From KS_Form Where ID=" & ID)
		  Conn.Execute("Delete From KS_FormField Where ItemID=" & ID)
		  If Err<>0 Then
		   Conn.RollBackTrans
		  Else
		   Conn.CommitTrans
		  End If
		  Response.Write "<script>alert('表单项目删除成功!');location.href='KS.Form.asp';</script>" 
		End Sub
        		
		Sub Total()
		  If Not KS.ReturnPowerResult(0, "KSMS100062") Then          '检查权限
					 Call KS.ReturnErr(1, "")
		  			 Response.End
		  End If
		
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Form Where Status=1 order by ID asc",conn,1,1
		   With Response
		    .Write "<div class='tabTitle'>各表单项目的前台调用</div>"
		  	.Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"

		  Do While Not RS.Eof
			.Write "<tr height='35' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			
			If RS("PostByStep")="1" or not conn.execute("select top 1 FieldType From KS_FormField Where ItemID=" & RS("ID") & " And (FieldType=10 or FieldType=11)").eof Then
			.Write "<td width='50' class='splittd'></td><td width='140' class='splittd'><img src='../../images/37.gif'>&nbsp;<b>" & RS("FormName") & "</b></td><td class='splittd'><input type='text' class='textbox' value='&lt;iframe src=&quot;{#GetFullDomain}/plus/form/form.asp?id=" & rs("id") & "&m={$ChannelID}&d={$InfoID}&quot; width=&quot;550&quot; height=&quot;350&quot; allowtransparency=&quot;true&quot; frameborder=&quot;0&quot;&gt;&lt;/iframe&gt;' name='s" & rs(0) & "' size=60></td><td class='splittd'><input class=""button"" onClick=""jm_cc('s" & rs(0) & "')"" type=""button"" value=""复制到剪贴板"" name=""button""></td><td class='splittd'></td>"
			Else
			.Write "<td width='50' class='splittd'></td><td width='140'  class='splittd'><img src='../../images/37.gif'>&nbsp;<b>" & RS("FormName") & "</b></td><td  class='splittd'><input type='text' class='textbox' value='&lt;script type=&quot;text/javascript&quot; src=&quot;{#GetFullDomain}/plus/form/form.asp?id=" & rs("id") & "&m={$ChannelID}&d={$InfoID}&quot;&gt;&lt;/script&gt;' name='s" & rs(0) & "' size=60></td><td class='splittd'><input class=""button"" onClick=""jm_cc('s" & rs(0) & "')"" type=""button"" value=""复制到剪贴板"" name=""button""></td><td></td>"
			End If
			
			.Write "</tr>"
		    RS.MoveNext
		  Loop
		   .Write "</table>"
		  End With
		  RS.Close:Set RS=Nothing
		  %>
		  <div style="margin-top:20px" class="attention">
		   <strong>调用说明：</strong>
		   <li>前台模板表单如果只是单步表单并且表单不含联动字段和编辑器字段时采用<scrpit>调用,否则采用iframe调用,如果用iframe调用的请适当调整iframe的宽和高;</li>
		   <li>表单如果放在内容页调用时，可以和当前文档关联。即表单数据表(KS_Form_名称)里会记录模型ID和文档ID。</li>
		  </div>
		   <script type="text/javascript">
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
		   If KS.C("SuperTF")<>"1" Then
					 Call KS.ReturnErr(1, "")
		  			 Response.End
		  End If
		   With Response
			   Dim ID:ID=KS.ChkClng(KS.G("ID"))
			   If ID=0 Then .Redirect "?": Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_Form Where ID=" & ID,Conn,1,3
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
			 .Write "<script>location.href='?';</script>"
		   End With
		End Sub
		
		Sub FormManage()
		Dim TimeLimit,AllowGroupID,useronce,onlyuser,shownum,PostByStep,StepNum,ToUserEmail,Cipher,Templ_url,Tempc_url,delform,MaxPerPage_s,adminuserlist
		Dim TempStr,SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt,i,UserIPTime,Email,SubmitUrl,SubmitTips,iponce,mobileCode,AllowShowOnUser
		Dim FormName,ExpiredDate,StartDate,Status,Descript,TableName,UpLoadDir,AnonymousUpload
		
		If Not KS.ReturnPowerResult(0, "KSMS100061") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 Response.End
		End If

		Dim ID:ID = KS.ChkClng(KS.G("ID"))
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select top 1 * from KS_Form Where ID=" & ID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			Status = RS("Status")
			FormName     = RS("FormName")
			TableName    = Replace(RS("TableName"),"KS_Form_","")
			UpLoadDir    = RS("UpLoadDir")
			StartDate    = RS("StartDate")
			TimeLimit    = RS("TimeLimit")
			ExpiredDate  = RS("ExpiredDate")
			TimeLimit    = RS("TimeLimit")
            AllowGroupID = RS("AllowGroupID")
			Descript     = RS("Descript")
			useronce     = RS("useronce")
			iponce       = RS("iponce")
			onlyuser     = RS("onlyuser")
			shownum      = RS("shownum")
			PostByStep   = RS("PostByStep")
			StepNum      = RS("StepNum")
			ToUserEmail  = RS("ToUserEmail")
			UserIPTime   = RS("UserIPTime")
			Email        = RS("Email")
			SubmitUrl    = RS("SubmitUrl")
			SubmitTips   = RS("SubmitTips")
			mobileCode   = KS.ChkClng(RS("mobileCode"))
			AllowShowOnUser=KS.ChkClng(RS("AllowShowOnUser"))
			AnonymousUpload = RS("AnonymousUpload")
			adminuserlist=RS("adminuserlist")
			Cipher=RS("Cipher")
			Templ_url=RS("Templ_url")
			Tempc_url=RS("Tempc_url")
			delform=RS("delform")
			MaxPerPage_s=RS("MaxPerPage_s")
			if KS.ChkClng(MaxPerPage_s)=0 then MaxPerPage_s=5
		Else
		     iponce=0: Status=1:TimeLimit = 0:StartDate=Now():ExpiredDate=Now()+10:AllowGroupID="":useronce=0:onlyuser=0:shownum=1:UpLoadDir="form/":PostByStep=0:StepNum=1:ToUserEmail=0:AnonymousUpload=0:adminuserlist="":UserIPTime=0:MaxPerPage_s=20:SubmitUrl="":SubmitTips="恭喜,您的信息已提交成功!" :mobileCode=0:AllowShowOnUser=0
		End If
		
		With Response
		.Write "<title>添加表单</title>" &_
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &_
		"<script src=""../../../KS_Inc/common.js"" language=""JavaScript""></script>"&_
		"<script src=""../../../KS_Inc/jquery.js"" language=""JavaScript""></script>"&_
		"<script src=""../../../KS_Inc/DatePicker/WdatePicker.js""></script>"&_
		"<script src=""../../images/pannel/tabpane.js"" language=""JavaScript""></script>" & _
		"<link href=""../../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		"<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"&_
		"	<div class='tabTitle'>自定义表单管理</div>"&_
			
		"<div class=tab-page id=Formpanel>"& _
		"<form name=""myform"" method=""post"" action=""KS.fORM.asp?Action=EditSave&ID=" & ID & """ >" & _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""Formpanel"" ), 1 )"& _
        " </SCRIPT>"& _
             
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>基本信息</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" & _
		"<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'> <div align=""right""><strong>表单状态：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""Status"" value=""1"" "
		If Status = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""Status"" value=""0"" "
		If Status = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭</td>"
		.Write "    </tr>"

%>
		<script type="text/javascript">
		 function CheckForm()
		 {
		  if ($("input[name=FormName]").val()=="")
		  {
		    top.$.dialog.alert('请输入表单名称',function(){
		     $("input[name=FormName]").focus();
			});
		   return false;
		  }
		  if ($("input[name=TableName]").val()=="")
		  {
		    top.$.dialog.alert('请输入表单的数据表名称',function(){
		     $("input[name=TableName]").focus();
			});
		   return false;
		  }
		  $("form[name=myform]").submit();
		 }
		 
		 function changedate()
		 {
		   val=$("input[name=TimeLimit]:checked").val();
		   if (val==1){
		    $("#BeginDate").show()
		    $("#EndDate").show();		
		   }
		   else{
		    $("#BeginDate").hide();
		    $("#EndDate").hide();		
		   }
		 }
		 function changepage()
		 {
		   val=$("input[name=PostByStep]:checked").val();
		   if (val==1){
		    $("#StepNumArea").show();
		   }
		   else{
		    $("#StepNumArea").hide();
		   }
		 }
	
		</script>

		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>表单名称：</strong></div></td>      
			<td height="30"> <input name="FormName" class="textbox" type="text" value="<%=FormName%>" size="50"> <span class="tips">如：参赛报名表等。</span></td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><strong>数据表名称：</strong></td>      
			<td height="30"> KS_Form_<input name="TableName"<%If KS.G("Action")="Edit" then response.write " disabled"%> size="50" class="textbox" type="text" value="<%=TableName%>" size="50"> 
			<br/><div class="tips">说明：创建数据表后无法修改，并且用户创建的数据表以"KS_Form_"开头</div></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><strong>上传目录：</strong></td>
			<td><%=KS.Setting(91)%><input name="UpLoadDir" class="textbox" type="text" value="<%=UpLoadDir%>" size="50"> 
			<br><div class="tips">说明：只能用字母和数字的组合，且必须与/结束。</div></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><strong>提交成功提示消息：</strong></td>
			<td><input name="SubmitTips" class="textbox" type="text" value="<%=SubmitTips%>" size="50"> <span class="tips">如:恭喜,您的报名数据提交成功,请等待审核。</span>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><strong>提交成功返回地址：</strong></td>
			<td><input name="SubmitUrl" class="textbox" type="text" value="<%=SubmitUrl%>" size="50"> <span class="tips">如:http://www.kesion.com,留空将直接返回提交页面。</span>
			</td> 
		</tr>
 		<%
		If KS.G("Action")="Edit" Then
			if  ks.isnul(Templ_url) then Templ_url="{@TemplateDir}/表单/列表页.html"
			if  ks.isnul(Tempc_url) then Tempc_url="{@TemplateDir}/表单/内容页.html"		
		else
			Templ_url="{@TemplateDir}/表单/列表页.html"
			Tempc_url="{@TemplateDir}/表单/内容页.html"
		end if	
		%>
 		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>表单列表页模板绑定：</strong></div></td>      
			<td height="30"><input name="Templ_url" id="Templ_url" class="textbox" type="text" value="<%=Templ_url%>" size="50">
			<% =KSCls.Get_KS_T_C("document.getElementById('Templ_url')") %>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>表单内容页模板绑定：</strong></div></td>      
			<td height="30"> <input name="Tempc_url" id="Tempc_url" class="textbox" type="text" value="<%=Tempc_url%>" size="50">
			<% =KSCls.Get_KS_T_C("document.getElementById('Tempc_url')") %>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>列表分页数：</strong></div></td>      
			<td height="30"><input name="MaxPerPage_s" style="text-align:center" id="MaxPerPage_s" class="textbox" type="text"  value="<%=MaxPerPage_s%>" size="5">条分一页
			</td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>启用分步提交表单：</strong></div></td>      
			<td height="30"> <input onClick="changepage()" name="PostByStep" type="radio" value="1"<%if PostByStep="1" Then Response.Write " Checked"%>>启用 <input onClick="changepage()" type="radio" value="0" name="PostByStep"<%if PostByStep="0" Then Response.Write " Checked"%>>不启用
			<br/><div class="tips">当需要收集的用户资料较多时,可以启用分步提交功能。</div>
			</td> 
		</tr>
		<tr id="StepNumArea" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>分步设置：</strong></div></td>      
			<td height="30"> 用户分为<input name="StepNum" size="4" class="textbox" type="text" value="<%=StepNum%>" style="text-align:center">步提交</td> 
		</tr>

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>表单备注：</strong></div></td>      
			<td height="30"> <textarea name="Descript" class="textbox" style="width:400px;height:90px"><%=Descript%></textarea></td> 
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
			response.write "<input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""1"" "
		If TimeLimit = 1 Then .Write (" checked")
		.Write ">"
		.Write "启用"
		.Write "  <input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""0"" "
		If TimeLimit = 0 Then .Write (" checked")
		.Write ">"
		.Write "不启用"
		
			%>
			</td> 
		</tr>

		<tr ID="BeginDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">     
		<td height="30" class="clefttitle"align="right"><div><strong>生效时间：</strong></div></td>     
		<td height="30"><input name="StartDate" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" id='StartDate' class="textbox Wdate" type="text" value="<%=StartDate%>" size="50"><br><font color=#ff0000>日期格式：YYYY-MM-DD hh:mm:ss</font></td>   
		</tr> 
		
		<tr ID="EndDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>失效时间：</strong></div></td>      
			<td height="30"> <input name="ExpiredDate" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" id="ExpiredDate" class="textbox Wdate" type="text" value="<%=ExpiredDate%>" size="50"><br><font color=#ff0000>日期格式：YYYY-MM-DD hh:mm:ss</font></td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>只允许会员提交：</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""onlyuser"" value=""1"" "
		If onlyuser = 1 Then .Write (" checked")
		.Write ">"
		.Write "是"
		.Write "  <input type=""radio"" name=""onlyuser"" value=""0"" "
		If onlyuser = 0 Then .Write (" checked")
		.Write ">"
		.Write "不是"
		
			%>
			<br/>
			</td> 
		</tr>	
			<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>会员允许在会员中心管理自己提交的数据：</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""AllowShowOnUser"" value=""1"" "
		If AllowShowOnUser = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""AllowShowOnUser"" value=""0"" "
		If AllowShowOnUser = 0 Then .Write (" checked")
		.Write ">"
		.Write "不允许"
		
			%>
			<br/>
			</td> 
		</tr>	
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>开启游客上传权限：</strong></div>
			
			</td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""AnonymousUpload"" value=""1"" "
		If AnonymousUpload = 1 Then .Write (" checked")
		.Write ">"
		.Write "开启"
		.Write "  <input type=""radio"" name=""AnonymousUpload"" value=""0"" "
		If AnonymousUpload = 0 Then .Write (" checked")
		.Write ">"
		.Write "不开启"
		
			%>
            <div class="tips">当表单项有上传项时，如果允许游客上传，可以这里开启。</div>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>每个会员只允许提交一次：</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""useronce"" value=""1"" "
		If useronce = 1 Then .Write (" checked")
		.Write ">"
		.Write "是"
		.Write "  <input type=""radio"" name=""useronce"" value=""0"" "
		If useronce = 0 Then .Write (" checked")
		.Write ">"
		.Write "不是"
		
			%>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>每个IP只允许提交一次：</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input onclick=""$('#ips').hide()"" type=""radio"" name=""iponce"" value=""1"" "
		If iponce = 1 Then .Write (" checked")
		.Write ">"
		.Write "是"
		.Write "  <input type=""radio"" onclick=""$('#ips').show()"" name=""iponce"" value=""0"" "
		If iponce = 0 Then .Write (" checked")
		.Write ">"
		.Write "不是"
		
			%>
			</td> 
		</tr>
		<tr id="ips"<%If iponce = 1 then response.write " style='display:none'"%> valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>IP提交限制：</strong></div></td>      
			<td height="30"> 
		     同一个IP<input type="text" name="UserIPTime" class="textbox" style="text-align:center;width:50px" value="<%=UserIPTime%>"/>小时内只能提交一次。不限制请输入“0”。
			</td> 
		</tr>
		
	

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>用户组限制：</strong></div><div class="tips">不限制，请不要选</div></td>      
			<td height="30"><%=KS.GetUserGroup_CheckBox("AllowGroupID",AllowGroupID,5)%> 
            
            </td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>允许管理此表单的普通管理列表：</strong></div></td>      
			<td height="30"> 
			<textarea name="adminuserlist" style="width:370px;height:90px" id="adminuserlist" class="textbox"><%=adminuserlist%></textarea>
		<div class="tips">
		TIPS:<br/>
		1、如果允许普通管理员管理此表单，在此输入允许管理表单留言数据的普通管理员用户名，多个管理员用英文逗号隔开，如admin,kesion等；<br/>
		2、普通管理管理员只允许管理提交的表单记录，没有创建表单字段及模板权限；<br/>
		3、所有管理员都允许管理此表单，请留空。
		</div>
		</td>
		</tr>
			<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>用户删除留言权限：</strong></div></td>      
			<td height="30"> 
			<%
			response.write "<input type=""radio"" name=""delform"" value=""1"" "
			If delform = 1 or  ks.isnul(delform) Then .Write (" checked")
			.Write ">"
			.Write "可以删除"
			.Write "  <input type=""radio"" name=""delform"" value=""0"" "
			If delform = 0 Then .Write (" checked")
			.Write ">"
			.Write "不能删除"
			%>
			</td> 
	    	</tr>	
			
			<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
				<td width="160" height="30" class="clefttitle" align="right"><div><strong>启用验证码：</strong></div></td>      
				<td height="30"> 
				<%
				.Write "<input type=""radio"" name=""shownum"" value=""1"" "
				If shownum = 1 Then .Write (" checked")
				.Write ">"
				.Write "显示"
				.Write "  <input type=""radio"" name=""shownum"" value=""0"" "
				If shownum = 0 Then .Write (" checked")
				.Write ">"
				.Write "不显示"
				%>
				</td> 
	    	</tr>
			<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
				<td width="160" height="30" class="clefttitle" align="right"><div><strong>启用手机短信验证码：</strong></div></td>      
				<td height="30"> 
				<%
				.Write "<input type=""radio"" name=""mobileCode"" value=""1"" "
				If mobileCode = 1 Then .Write (" checked")
				.Write ">"
				.Write "启用"
				.Write "  <input type=""radio"" name=""mobileCode"" value=""0"" "
				If mobileCode = 0 Then .Write (" checked")
				.Write ">"
				.Write "不启用"
				%>
				</td> 
	    	</tr>
			
			<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>密码功能：</strong></div></td>      
			<td height="30"> 
			<%
			response.write "<input type=""radio"" name=""Cipher"" value=""1"" "
			If Cipher = 1 or  ks.isnul(Cipher) Then .Write (" checked")
			.Write ">"
			.Write "开启"
			.Write "  <input type=""radio"" name=""Cipher"" value=""0"" "
			If Cipher = 0 Then .Write (" checked")
			.Write ">"
			.Write "关闭"
			
			%>
			</td> 
	    	</tr>
			
			
			</table>
        </div>
		
		 <div class=tab-page id="formset1">
		  <H2 class=tab>表单通知</H2>
			<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "formset1" ) );
			</SCRIPT>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
			 	<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>将提交结果发送到邮箱：</strong></div>
			</td>      
			<td height="30"> 
			
			<%
		.Write "  <input type=""radio"" onclick=""$('#ee').hide();"" name=""ToUserEmail"" value=""0"" "
		If ToUserEmail = 0 Then .Write (" checked")
		.Write ">不启用"
		
		.write "<input type=""radio"" onclick=""$('#ee').show();"" name=""ToUserEmail"" value=""1"" "
		If ToUserEmail = 1 Then .Write (" checked")
		.Write ">仅发给管理员"
		.write "<input type=""radio"" onclick=""$('#ee').hide();"" name=""ToUserEmail"" value=""2"" "
		If ToUserEmail = 2 Then .Write (" checked")
		.Write ">仅发给用户"
		.write "<input type=""radio"" onclick=""$('#ee').show();"" name=""ToUserEmail"" value=""3"" "
		If ToUserEmail = 3 Then .Write (" checked")
		.Write ">同时发给管理员和用户"
		
			%><div class="tips">当要求用户填写邮箱时，如果启用此功能将自动将用户的提交结果发到用户的邮箱或管理员邮箱。</div>
		    <div id="ee"<%If ToUserEmail = 0 or ToUserEmail = 2 Then .Write (" style='display:none'")%>>
			管理员邮件<input type="text" name="Email" value="<%=Email%>" class="textbox" size="60"/>多个邮箱要接收请用英文逗号隔开。
			</div>
			</td> 
		</tr>
			</table>
		</div>
		
		<script>changedate();changepage();</script>
		<%
		.Write "</form>"
		.Write "</div>"
		End With
		End Sub
		
		'表单模板管理
		Sub FormTemplate()
		 Dim FormID:FormID=KS.ChkClng(KS.G("ItemID"))
		 Dim RS,Template,FormName,PostByStep,StepNum,Step,K,Templatebm
		 
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select FormName,PostByStep,StepNum,Template,Templatebm From KS_Form Where ID=" & FormID,conn,1,1
		 If RS.EOF And RS.Bof Then
		  Response.Write "<script>alert('error!');history.back();</script>"
		  Exit Sub
		 Else
		   FormName=RS(0):PostByStep=RS(1):StepNum=RS(2):Template=RS(3):Templatebm=RS(4)
		   If PostByStep=0 Then StepNum=1
		 End If
		 RS.Close
         If Template="" Or IsNull(Template) Then Template=" "
		 Template=Split(Template,"$aaa$")
		%>
		<html>
		<title>表单模板管理</title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../../../KS_Inc/common.js" language="JavaScript"></script>
		<script src="../../../KS_Inc/jquery.js" language="JavaScript"></script>
		<link href="../../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
         <script language = 'JavaScript'>
					function LoadTemplate()
					{   
					   if ($("#autocreate")[0].checked==true)
					    { 
							var url='KS.Form.asp';
							$.ajax({
								  url: url,
								  cache: false,
								  data: "action=createtemplate&formid="+$("#FormID").val(),
								  success: function(s){
									s=s.split("$aaa$");
								   <%For K=1 To StepNum%>
									  $('textarea[name=Content<%=K%>]').val(s[<%=K-1%>]);
									  if ($('textarea[name=Content<%=K%>]').val()=='undefined')
									  $('textarea[Content<%=K%>]').val('请添加表单项!');
								   <%Next%>
								  }
								});	

								$.ajax({
								  url: url,
								  cache: false,
								  data: "dy=ok&action=createtemplate&formid="+$("#FormID").val(),
								  success: function(s){
									s=s.split("$aaa$");
								    //$('textarea[name=Templatebm]').val(s[0]);
								   <%For K=1 To StepNum%>
									  $('textarea[name=Templatebm]').val(s[<%=K-1%>]);
								   <%Next%>
								  }
								});							  
						}
						else
						{
						  $('#Content').val('');
						}
					}	

		            function show_ln(txt_ln,txt_main){
			            var txt_ln  = document.getElementById(txt_ln);
			            var txt_main  = document.getElementById(txt_main);
			            txt_ln.scrollTop = txt_main.scrollTop;
			            while(txt_ln.scrollTop != txt_main.scrollTop)
			            {
				            txt_ln.value += (i++) + '\n';
				            txt_ln.scrollTop = txt_main.scrollTop;
			            }
			            return;
		            }
		            function editTab(){
			            
			            }
					 function CheckForm()
					 {
					 		  $("#myform").submit();
					 }
		            //-->
		            </script>

	  <body>
		<div class="tabTitle">自定义表单[<%=FormName%>]模板管理</div>
		<form name="myform" id="myform" action="KS.Form.asp?action=TemplateSave" method="post">
		<input type="hidden" value="<%=formid%>" name="FormID" id="FormID">
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		   <tr class='tdbg'>
		      <td class='clefttitle' width="120" align="right"><strong>自动生成模板：</strong></td>
		     <td height="30">
			 <input type='checkbox' name='autocreate' id='autocreate' value='1' onClick="LoadTemplate()">自动生成
			 <font color=red>提示：第一次生成模板，可以点此自动生成！</font>
			 </td>
		   </tr>
		  
		   <% 
		   on error resume next
		   For K=1 To StepNum%> 
		   <tr class='tdbg'>
		      <td class='clefttitle' align="right" width="130"><strong>表单模板<%If PostByStep=1 Then %>(第<font color=red><%=K%></font>步)<%End If%>：</strong>
			  <%If K>1 Then Response.Write "<br><font color=red>必须包括{$HiddenFields}标签</font>" %>
			  </td>
		     <td height="280">
			 <textarea id='txt_ln<%=K%>' name='rollContent' cols='6' style='overflow:hidden;height:280px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly><%
		 Dim N
		 For N=1 To 3000
			Response.Write N & "&#13;&#10;"
		 Next
		 On Error Resume Next
		 %>
		 </textarea>
		 <textarea name='Content<%=K%>' style='width:90%;height:280px' ROWS='15' id='txt_main<%=K%>' onkeydown='editTab()' onscroll="show_ln('txt_ln<%=K%>','txt_main<%=K%>')" wrap='on'><%=server.HTMLEncode(Template(K-1))%></textarea>
			 </td>
		   </tr>
		   <%Next%>
            <tr class='tdbg'>
		      <td class='clefttitle' align="right" width="130"><strong>打印报名模板：</strong>
			  </td>
		     <td height="280">
			 <textarea id='txt_ln<%=K%>' name='rollContent' cols='6' style='overflow:hidden;height:280px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly><%
		 
		 For N=1 To 3000
			Response.Write N & "&#13;&#10;"
		 Next
		 On Error Resume Next
		 %>
		 </textarea>
		 <textarea name='Templatebm' style='width:90%;height:280px' ROWS='15' id='Templatebm' onkeydown='editTab()'  wrap='on'><%=server.HTMLEncode(Templatebm)%></textarea>
         	<br><br>
			 </td>
		   </tr>
           
		 </table>  
		 </form>
		<%
		
		End Sub
		
		Sub AutoTemplate()
		 Response.CharSet="utf-8"
		 Dim ShowNum,PostByStep,StepNum,K,Param,S,KK,Cipher,MobileCode,formname
		 Dim SQL,N,O_Arr,O_Len,F_V,BrStr,O_Value,O_Text
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim dy:dy=KS.G("dy")
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 ShowNum,PostByStep,StepNum,Cipher,MobileCode,formname From KS_Form Where ID=" & FormID,conn,1,1
		 If Not RS.Eof Then
		  ShowNum=KS.ChkClng(RS(0)):PostByStep=RS(1):StepNum=RS(2):Cipher=RS(3):MobileCode=KS.ChkClng(RS(4)):formname=rs(5)
		 End If
		 RS.Close
		 
		 For S=1 To StepNum
		     SQL=""
		     Param="Where ShowOnForm=1 and ItemID=" & FormID 
			 If PostByStep=1 Then
			   If s=1 Then
			    Param=Param & " and step<=" & S
			   Else
			    Param=Param & " and step=" & S
			   End If
			 End If
			 RS.Open "Select Title,FieldName,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,AllowFileExt,MaxFileSize,FieldID,ParentFieldName,ShowUnit,UnitOptions,MaxLength From KS_FormField " & Param & " order by orderid",conn,1,1
			 If Not RS.Eof Then SQL=RS.GetRows(-1)
			 RS.Close
			 If Not IsArray(SQL) Then Response.Write "该表单还没有添加表单项!":Response.End
			 If PostByStep=1 Then
			 Response.Write "<div style=""text-align:center"">第 " & S & " 步</div>" & vbcrlf
			 Elseif  dy<>"ok" then
			 Response.Write "<iframe src=""about:blank"" name=""hidform" & FormID & """ style=""width:0px;height:0px;display:none""></iframe>" &vbcrlf
			 End If
			 Response.Write"<script src=""" & KS.Setting(3) &"ks_inc/common.js""></script>"&vbcrlf
			 Response.Write"<script src=""" & KS.Setting(3) &"ks_inc/form_ck.js""></script>"&vbcrlf
			 if  dy<>"ok" then
			    If PostByStep=1 Then
				 Response.Write "<form name=""myform" & FormID &""" action=""" & ks.setting(3) & "plus/form/form.asp"" onSubmit=""return(form_c(" & FormID &"));""   method=""post""> " &vbcrlf
				Else
				 Response.Write "<form name=""myform" & FormID &""" action=""" & ks.setting(3) & "plus/form/form.asp"" onSubmit=""return(form_c(" & FormID &"));""  target=""hidform" & FormID & """ method=""post""> " &vbcrlf
				End If
				 Response.Write "<table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""1"">" & vbcrlf
			 else
			 	Response.Write "<table width=""100%"" border=""1"" cellspacing=""1"" cellpadding=""0"">" & vbcrlf
			 end if	 
			 
			 if dy<>"ok" then
				 If (PostByStep=1 And S=StepNum)  Or PostByStep=0 Then
				 Response.Write "<input type=""hidden"" value=""Save"" name=""action""/>" & vbcrlf
				 Else
				 Response.Write "<input type=""hidden"" value=""Next"" name=""action""/>" & vbcrlf
				 End If
				 Response.Write "<input type=""hidden"" value=""" & FormID & """ name=""id""/>" & vbcrlf
				 If PostByStep=1 Then
				 Response.Write "<input type=""hidden"" value=""" & S & """ name=""Step""/>" & vbcrlf
				 End If
				 Response.Write "<input type=""hidden"" value=""{$ChannelID}"" name=""m""/>" & vbcrlf
				 Response.Write "<input type=""hidden"" value=""{$InfoID}"" name=""d""/>" & vbcrlf
				 Response.Write "<input type=""hidden"" value=""{$ToUserName}"" id=""tousername"" name=""tousername""/>" & vbcrlf
			 end if	 
			 If S>1 Then	 Response.Write "{$HiddenFields}" & vbcrlf
			 
			 For K=0 To Ubound(SQL,2)
			 If SQL(12,K)="0" Or KS.IsNul(SQL(12,K)) Then
			 Response.Write " <tr class=""tdbg"">" & vbcrlf
			 Response.Write "  <td align=""right"" class=""lefttdbg"">" & SQL(0,K) & "：</td>" & vbcrlf
			 if KS.ChkClng(SQL(3,K))=10 Then
			 Response.Write "  <td style=""height:" & SQL(8,K) & "px;width:" & KS.ChkClng(SQL(7,K))+100 &"px;"">" 
			 Else
			 Response.Write "  <td>" 
			 End If
			 if dy<>"ok" then
			 Select Case SQL(3,K)
				Case 2
				  Response.Write "<textarea style=""width:" & SQL(7,K) & "px;height:" & SQL(8,K) &"px"" rows=""5"" id=""" & SQL(1,K) & """ name=""" & SQL(1,K) & """>" & SQL(4,K) & "</textarea>"
			   Case 3,11
			     If SQL(3,K)=11 Then
				  Response.Write "<select class=""upfile"" id=""" & SQL(1,K) &""" onchange=""fill" & SQL(1,K) &"(this.value)"" style=""width:" & SQL(7,K) & "px"" name=""" & SQL(1,K) & """><option value=''>---请选择---</option>"
				 Else
				  Response.Write "<select class=""upfile"" id=""" & SQL(1,K) &""" style=""width:" & SQL(7,K) & "px"" name=""" & SQL(1,K) & """>"
				 End If
				  O_Arr=Split(SQL(5,K),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
					If O_Arr(N)<>"" Then
						F_V=Split(O_Arr(N),"|")
						If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						Else
							O_Value=F_V(0):O_Text=F_V(0)
						End If						   
						If SQL(4,K)=O_Value Then
							Response.Write "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
						Else
							Response.Write "<option value=""" & O_Value& """>" & O_Text & "</option>"
						End If
					End If
				  Next
				  Response.Write "</select>"
                  '联动菜单
					If SQL(3,K)=11  Then
						Dim JSStr
						Response.Write  GetLDMenuStr(FormID,SQL(1,k),JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
					End If				  
			  Case 6
				  O_Arr=Split(SQL(5,K),vbcrlf): O_Len=Ubound(O_Arr)
				  If O_Len>1 And Len(SQL(5,I))>50 Then BrStr="<br>" Else BrStr=""
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If SQL(4,K)=O_Value Then
						Response.Write "<input type=""radio"" id=""" & SQL(1,K) &""" name=""" & SQL(1,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
					Else
						Response.Write "<input type=""radio"" id=""" & SQL(1,K) &""" name=""" & SQL(1,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
					End If
				   End If
				  Next
			 Case 7
				   O_Arr=Split(SQL(5,K),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If KS.FoundInArr(SQL(4,K),O_Value,",")=true Then
						Response.Write "<input type=""checkbox"" id=""" & SQL(1,K) &""" name=""" & SQL(1,K) & """ value=""" & O_Value& """ checked>" & O_Text
					Else
						Response.Write "<input type=""checkbox"" id=""" & SQL(1,K) &""" name=""" & SQL(1,k) & """ value=""" & O_Value& """>" & O_Text
					End If
					End If
				  Next
			 Case 10
			      '  Response.Write EchoUeditorHead()
					Response.Write "<script id=""" & SQL(1,K) &""" name=""" & SQL(1,K) &""" type=""text/plain"" style=""width:" & SQL(7,K) & "px;;height:" & SQL(8,K) & "px;"">" &SQL(4,K)&"</script>"
	                Response.Write "<script>setTimeout(""var editor = " & GetEditorTag() &".getEditor('" & SQL(1,K) &"',{toolbars:[" & GetEditorToolBar("Basic") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:" & SQL(8,K) & " });"",10);</script>"

    		 Case Else
			    Dim MaxLength:MaxLength=SQL(15,K)
				If Not IsNumerIc(MaxLength)  Or MaxLength="0" Then MaxLength=255
				Response.Write "<input type=""text"" maxlength=""" & MaxLength &""" class=""upfile"" style=""width:" & SQL(7,K) & "px"" id=""" & SQL(1,K) & """ name=""" & SQL(1,K) & """ value=""" & SQL(4,K) & """>"
			 End Select
			 
              If SQL(13,K)="1" Then 
					  Response.Write " <select name=""" & SQL(1,K) & "_Unit"" id=""" & SQL(1,K) & "_Unit"">"
					  If Not KS.IsNul(SQL(14,K)) Then
				       Dim UnitOptionsArr:UnitOptionsArr=Split(SQL(14,K),vbcrlf)
					   For KK=0 To Ubound(UnitOptionsArr)
					     response.write "<option value='" & UnitOptionsArr(kk) & "'>" & UnitOptionsArr(kk) & "</option>"                 
					   Next
					  End If
					  response.write "</select>"
			 End If
			 
			 If SQL(6,K)=1 and SQL(3,K)<>3 then
				 Response.Write "<span class=""formstrck" & FormID &""" id=""" &SQL(0,K)&"|%|"&SQL(1,K)&"|%|"&SQL(3,K)&""" style=""display:none""></span>"
			 end if		
			 
			 		
			 If SQL(6,K)=1 Then Response.Write "<font color=""red""> * </font>"
			 If SQL(2,K)<>"" Then Response.Write " <span style=""margin-top:5px"">" &  SQL(2,K) & "</span>"
			 If SQL(3,K)=9 Then Response.Write "可上传文件类型" & SQL(9,K) & ",大小" & SQL(10,K) & " KB<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='" &KS.Setting(3) & "user/User_UpFile.asp?FormID=" & FormID & "&Type=Field&FieldID=" & SQL(11,K) & "' frameborder=0 scrolling=no width='100%' height='30'></iframe></div>"
			Else
			  response.write "<span id=""" & SQL(1,K) &"""></span>"
			End If
			 Response.Write "  </td>" & vbcrlf
			 Response.Write "</tr>" & vbcrlf
			 End If
			 Next
			 if Cipher =1 then
			 	Response.Write "<tr class=""tdbg""><td class=""lefttdbg"" align=""right""> 是否公开：</td><td><input name=""Cipher"" value=""1"" type=""radio"" checked=""checked"" onclick=""$('#PassWord').hide();"">是(前台显示)<input name=""Cipher""  value=""0"" type=""radio"" onclick=""$('#PassWord').show();$('input[name=PassWord]').focus();"">否(查看要输入密码)"
				Response.Write "<span id=""PassWord"" style="" display:none""><BR/>输入密码:<input name=""PassWord""  type=""password"" ></span>"
				Response.Write "</td></tr>"  &vbcrlf
			 end if	
			 if dy<>"ok" then 
			     IF MobileCode=1 And  (PostByStep=0 or S=StepNum)  Then  '短信验证码
				  Response.Write "<tr class=""tdbg""><td class=""lefttdbg"" align=""right"">手机号码：</td><td><input name=""Mobile"" id=""Mobile"" type=""text"" name=""textbox""></td></tr>"  &vbcrlf
				  Response.Write "<tr class=""tdbg""><td class=""lefttdbg"" align=""right"">手机验证码：</td><td><input name=""MobileCode"" id=""MobileCode"" type=""text"" name=""textbox"" size=6><input type=""button"" value=""免费获取手机验证码"" id=""MobileCodeBtn"" onclick=""getMobileCode("& KS.ChkClng(split(KS.Setting(156)&"∮","∮")(1)) &",'104','Mobile','MobileCodeBtn','"& formname & "')"" class=""button""/></td></tr>"  &vbcrlf
				  
				 End If
				 IF ShowNum=1 And  (PostByStep=0 or S=StepNum)  Then
				 Response.Write "<tr class=""tdbg""><td class=""lefttdbg"" align=""right"">验证码：</td><td><input name=""Verifycode"" id=""Verifycode"" type=""text"" name=""textbox"" size=5><span style=""color:#999"">请输入下图中的字符</span><br/><IMG style=""cursor:pointer"" src=""" & KS.Setting(3) & "plus/verifycode.asp"" onClick=""this.src='" &KS.Setting(3) & "plus/verifycode.asp?n='+ Math.random();"" align=""absmiddle""></td></tr>"  &vbcrlf
				 End If
				 If S=StepNum or PostByStep=0 Then
				 Response.Write "<tr><td colspan=""2"" class=""subtdbg"" align=""center""><input type=""submit"" value=""确认提交"" name=""submit1""></td></tr>"  & vbcrlf
				 Else
				 Response.Write "<tr><td colspan=""2"" class=""subtdbg"" align=""center""><input type=""submit"" value=""OK，下一步"" name=""submit1""></td></tr>"  & vbcrlf
				 End If
			 end if
			 Response.Write "</table>" & vbcrlf
			 if dy<>"ok" then Response.Write "</form>" &vbcrlf

			 Response.Write "$aaa$" & vbcrlf
			
		   Next	 
			 
			 
		End Sub
		
		'取得联动菜单
		   Function GetLDMenuStr(ItemID,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str
		     Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_FormField Where itemid=" & ItemID & " and ParentFieldName='" & ParentFieldName & "'")
			 If Not RSL.Eof Then
			     Str=Str & " <select name='" & RSL(0) & "' id='" & RSL(0) & "' onchange='fill" & RSL(0) & "(this.value)' style='width:" & RSL(3) & "px'><option value=''>--请选择--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();"
				  Options=RSL(2)
				  OArr=Split(Options,Vbcrlf)
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=trim(OArr(i))
					 F=trim(OArr(i))
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>"
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& RSL(0)&"').empty();" &vbcrlf &_
							   "$('#"& RSL(0)&"').append('<option value="""">--请选择--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & RSL(0) & "').options[document.getElementById('" & RSL(0) & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}"

				 GetLDMenuStr=str & GetLDMenuStr(itemid,RSL(0),JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
			     
		   End Function
				
		'表单模板管理
		Sub FormView()
		 Dim FormID:FormID=KS.ChkClng(KS.G("ItemID"))
		 Dim PostByStep:PostByStep=LFCls.GetSingleFieldValue("Select PostByStep From KS_Form Where ID=" & FormID)
		%>
		<html>
		<title>预览表单</title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../../../KS_Inc/common.js" language="JavaScript"></script>
        <script>
		 function modifyTp(){
			location="KS.Form.asp?ItemID=<%=FormID%>&action=template";
			window.$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=自定义表单 >> <font color=red>修改表单管理</font>&ButtonSymbol=GoSave';

		 }
		</script>
		<link href="../../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	  <body>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'>
		  <tr>
			<td height='25' class='sort'>自定义表单效果预览</td>
		 </tr>
		 <tr><td height=5></td></tr>
		</table>
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		   <tr class='tdbg'>
		      <td class='clefttitle' height="25" style="text-align:left"><strong>表单名称：<font color=red><%=Conn.Execute("Select FormName From KS_Form Where ID=" & FormID)(0)%></font></strong></td>
		   </tr>
		   <tr class='tdbg'>
		     <td>
			 <%If PostByStep=1 or not conn.execute("select top 1 FieldType From KS_FormField Where ItemID=" & FormID & " And (FieldType=10 or FieldType=11)").eof Then%>
			  <iframe src="../../../plus/form/form.asp?id=<%=formid%>" frameborder="0" width="700" height="500" allowtransparency="true" align="middle"></iframe>
			 <%else%>
			 <script src="../../../plus/form/form.asp?id=<%=formid%>"></script>
			 <%end if%>
			 </td>
		   </tr>
		   <tr class='tdbg'>
		      <td class='clefttitle' height="25" style="text-align:center"><input type="button" class="button" onClick="modifyTp();" value="修改模板"></td>
		   </tr>
		 </table>  
		<%
		
		End Sub

		
		Sub FormSave()
		    Dim ExpiredDate,StartDate,I,OpName,ID:ID=KS.ChkClng(KS.G("ID"))
			StartDate=KS.G("StartDate")
			ExpiredDate=KS.G("ExpiredDate")
			If Not IsDate(StartDate) Then Call KS.AlertHistory("生效日期格式不正确",-1):response.end
			If Not IsDate(ExpiredDate) Then Call KS.AlertHistory("失效日期格式不正确",-1):response.end
			If ID=0 and Not Conn.Execute("select top 1 id from ks_form where tablename='KS_Form_" & KS.G("TableName") &"'").eof then Call KS.AlertHistory("数据表已存在！",-1):response.end
			on error resume next
			Conn.BeginTrans
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Form Where ID=" & ID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				RS("TableName")= "KS_Form_" & KS.G("TableName")
				OpName      = "添加"
			Else
			    OpName="修改"
			End If
				RS("FormName")= KS.G("FormName")
				RS("UploadDir")= KS.G("UpLoadDir")
				RS("Status") = KS.G("Status")
				RS("TimeLimit")   = KS.ChkClng(KS.G("TimeLimit"))
				RS("StartDate")     = startdate
				RS("ExpiredDate")    = ExpiredDate
				RS("useronce") = KS.ChkClng(KS.G("useronce"))
				RS("iponce")   = KS.ChkClng(KS.G("iponce"))
				rs("onlyuser") = KS.ChkClng(KS.G("onlyuser"))
				rs("shownum")  = ks.chkclng(ks.g("shownum"))
				RS("AllowGroupID")     = KS.G("AllowGroupID")
                RS("Descript")    = KS.G("Descript")
				RS("PostByStep")  = KS.ChkClng(KS.G("PostByStep"))
				RS("StepNum")     = KS.ChkClng(KS.G("StepNum"))
				RS("ToUserEmail") = KS.ChkClng(KS.G("ToUserEmail"))
				RS("AnonymousUpload")=KS.ChkClng(KS.G("AnonymousUpload"))
				RS("adminuserlist")=ks.g("adminuserlist")
				RS("Cipher")=KS.ChkClng(KS.G("Cipher"))
				RS("Templ_url")=KS.G("Templ_url")
				RS("Tempc_url")=KS.G("Tempc_url")
				RS("delform")=KS.G("delform")
				RS("MaxPerPage_s")=KS.ChkClng(KS.G("MaxPerPage_s"))
				RS("UserIPTime")=KS.ChkClng(KS.G("UserIPTime"))
				RS("Email")=KS.G("Email")
				RS("SubmitTips")=KS.G("SubmitTips")
				RS("SubmitURL")=KS.G("SubmitUrl")
				RS("mobileCode")=KS.ChkClng(KS.G("mobileCode"))
				RS("AllowShowOnUser")=KS.ChkClng(KS.G("AllowShowOnUser"))
				RS.Update
				RS.Close
				Set RS=Nothing
				
				If OpName="添加" Then
				 Dim sql:sql="CREATE TABLE [KS_Form_" & KS.G("TableName") & "] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_KS_Form_" & KS.G("TableName") & " PRIMARY KEY,"&_
						"ToUserName nvarchar(100),"&_
						"UserName nvarchar(100),"&_
						"UserIP nvarchar(100),"&_
						"AddDate datetime,"&_
						"Mobile nvarchar(100),"&_
						"[Note] text,"&_
						"[ReplyDate] datetime,"&_
						"[PassWord] varchar(255),"&_
						"[FormID] int default 0,"&_
						"[ChannelID] int default 0,"&_
						"[InfoID] int default 0,"&_
						"Status tinyint default 0)"
				 Conn.Execute(sql)
				End If
				if err<>0 then
					Conn.RollBackTrans
					Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				else
					Conn.CommitTrans
					Response.Write ("<script>alert('" & OpName & "自定义表成功!');$(parent.document).find('#BottomFrame')[0].src='Post.Asp?OpStr=" & server.URLEncode("自定义表单 >> <font color=#ff0000>表单管理</font>") &"&ButtonSymbol=Disabled';location.href='KS.Form.asp';</script>")
				end if
		End Sub
		
		Sub SubmitResult()
		ID=KS.ChkClng(KS.G("itemID"))
		Dim TableName,MobileCode,SQL,II
		TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & ID)
		MobileCode=LFCls.GetSingleFieldValue("Select top 1 MobileCode From KS_Form Where ID=" & ID)
		MaxPerPage = 10     '取得每页显示数量
		If KS.G("page") <> "" Then
			  CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		 with response
		 %>
		  <script>
			function ShowReplay(formid,id)
			{  
			 top.openWin("回复表单记录","plus/plus_form/KS.Form.asp?Action=replay&formid="+formid+"&id=" +id+'&rnd='+Math.random(),false);
			 }
			</script>
			
			<div class="tabs_header">
				<ul class="tabs">
					<li><a href="KS.Form.asp?ItemID=<%=id%>&action=resulthp"><span>横排显示记录</span></a> </li>
					<li class='active'><a href="KS.Form.asp?ItemID=<%=id%>&action=result"><span>竖排显示记录</span></a></li>
				</ul>
			</div>

		 <%
		    .Write ("<div sstyle=""height:94%; overflow: auto; width:100%"" align=""center"">")
		 	.Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write "  <td width='40' align='center'>ID号</td><td align=center>提交内容</td><td align=center>↓管理操作</td>"
			.Write "</tr>"
			set rs=server.createobject("adodb.recordset")
			rs.open "select FieldName,title,MustFillTF,FieldType from ks_formfield where itemid=" & KS.ChkClng(KS.G("itemID")) & " and ShowOnForm=1 order by orderid",conn,1,1
			If Not RS.Eof Then SQL=RS.GetRows(-1)
			RS.Close
			rs.open "select * from " & TableName & " order by adddate desc" ,conn,1,1
			 If Not RS.EOF Then
					totalPut = Conn.Execute("Select count(1) From " & TableName)(0)
							
							If CurrentPage <> 1 Then
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								End If
							End If
							  dim k,i:i=1
							  do while not rs.eof
							   response.write "<tr><td width=40 align='center'>" & rs("id") & "</td>"
							   response.write "<td style='text-align:left'>"
							   If IsArray(SQL) Then
								 response.write "<table width='100%' border='0'>"
								 if MobileCode=1 then
								  response.write "<tr>"
								  response.write "<td width='100' align='right' style='height:22px'><b>手机号码：</b></td>"
								  response.write "<td>" & rs("mobile") & "</td>"
								  response.write "</tr>"
								 End If
								 
								 For II=0 To Ubound(SQL,2)
								  response.write "<tr>"
								  response.write "<td width='100' align='right' style='height:22px'><b>" & sql(1,ii) & "：</b></td>"
								  response.write "<td>" & rs(trim(sql(0,ii))) & "</td>"
								  response.write "</tr>"
								 Next
							   end if
							   response.write "</table>"
							   response.write "</td>"
							   response.write "<td style='text-align:left;line-height:22px;'>"
							   response.write "时 间："  &rs("adddate") & "<br>IP地址：" & rs("userip") & "<br>用 户：" & rs("username")
							   response.write "<br>状 态："
							   select case rs("status")
							   case 0
								response.write "<font color=red>未读</font>"
							   case 1
								response.write "<font color=green>已读</font>"
							   case 2
								response.write "<font color=#ff6600>采纳</font>"
							   case 3
								response.write "垃圾"
							   end select
							   
							   if not isnull(rs("note")) and rs("note")<>"" then response.write "&nbsp;&nbsp;<a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");""><font color=blue>已回复</font></a>"
							   
							   response.write "<br>操 作：<a href=""?action=delinfo&FormID=" & ID&"&id=" & rs("Id") & """ onclick=""return(confirm('确定删除吗?'))"">删除</a> <a href='KS.Form.asp?action=modifyinfo&FormID=" & id & "&id=" & rs("id") & "'>修改</a> <a href='?action=setstatus&v=1&FormID=" & ID&"&id=" & rs("id") & "' title='设为已读'>已读</a> <a href='?action=setstatus&v=2&FormID=" & ID&"&id=" & rs("id") & "' title='设为采纳'>采纳</a> <a href='?action=setstatus&v=3&FormID=" & ID&"&id=" & rs("id") & "' title='设为垃圾'>垃圾</a> <a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");"">回复</a> <a href=""KS.Form.asp?ItemID="& id &"&id="& rs("id") &"&Action=print_bm"" target=""_blank"" title='打印本条记录'>打印</a>"
							   response.write "</td>"
							   response.write "</tr>" 
							   Response.Write("<tr><td colspan=3><hr size=1 color=#f1f1f1></td></tr>")
							  rs.movenext
							  i=i+1
							  if i>maxperpage then exit do
							  loop
             Else
			   .Write "<tr><td colspan=4 class='splittd'>没有记录！</td></tr>"
			 End If
			  .Write ("<tr> ")
			  .Write ("<td height=""60"" colspan=""3"" style='text-align:left'><input type='button' class='button' onclick='window.print();' value='打印本页面记录'> <font color=red>温馨提示：这里您可以对提交的记录进行管理，回复等。对没有用的记录进行删除操作！</font>")
			  .Write ("</td>")
			  .Write ("</tr>")			
			  .Write ("<tr> ")
			  .Write ("<td height=""50"" colspan=""3""  align=""right"">")
			  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			  .Write ("<br></td>")
			  .Write ("</tr>")			
			  .Write "</table>"
		 %>
		 <form name="export" action="KS.Form.asp?action=export" method=post target="_blank">
		  <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
                  <input type="hidden" value="<%=id%>" name="id">
		   <strong>按时间段导出Excel</strong>
		   开始时间:<input type="text" name="startdate" class="textbox" size="26" value="<%=dateadd("d",now,-30)%>">
		   结束时间:<input type="text" name="enddate" class="textbox"  size="26" value="<%=formatdatetime(now,2)%>">
		   <input type="submit" class="button" value="导出Excel">
		   <input type="button" class="button" value="全部导出Excel" onClick="window.open('KS.Form.asp?action=export&id=<%=id%>')">
		  </div>
		  </form>
		 
		 <%
			  .Write "</div>"
         end with
		End Sub
		
		
		'横排显示
		Sub SubmitResultHP()
		ID=KS.ChkClng(KS.G("itemID"))
		Dim TableName,MobileCode,SQL,II
		TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & ID)
		MobileCode=LFCls.GetSingleFieldValue("Select top 1 MobileCode From KS_Form Where ID=" & ID)
		MaxPerPage = 20     '取得每页显示数量
		If KS.G("page") <> "" Then
			  CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		 with response
		 %>
		  <script type="text/javascript">
			function ShowReplay(formid,id)
			{  
			 top.openWin("回复表单记录","plus/plus_form/KS.Form.asp?Action=replay&formid="+formid+"&id=" +id+'&rnd='+Math.random(),false);
			}
			</script>
			
			<div class="tabs_header">
				<ul class="tabs">
					<li class='active'><a href="KS.Form.asp?ItemID=<%=id%>&action=resulthp"><span>横排显示记录</span></a> </li>
					<li><a href="KS.Form.asp?ItemID=<%=id%>&action=result"><span>竖排显示记录</span></a></li>
				</ul>
			</div>
			
		
		<div style="margin-top:10px;clear:both;width:100%;padding-bottom:5px;margin-bottom:5px;overflow-x: auto; height:auto">
		 <%
 			set rs=server.createobject("adodb.recordset")
			rs.open "Select Title,FieldName From KS_FormField Where ShowOnManage=1 And ItemID=" & ID & " Order By OrderID,FieldID",Conn,1,1
			If Not RS.Eof Then SQL=RS.GetRows(-1)
			RS.Close

		 	.Write "<table cellspacing=""1"" bordercolor=""#000000"" bgcolor=""#000000""  width='100%' align='center'>"
			.Write "<form name=""form1"" action=""KS.Form.asp?ItemID=" & ID &""" method=""post"">"
			.Write "<input type='hidden' name='action' id='action' value='setstatus'/>"
			.Write "<input type='hidden' name='v' id='v' value='2'/>"
			.Write "<input type='hidden' name='formid' id='formid' value='" & ID &"'/>"
			.Write "<tr height='25' bgcolor='#ffffff'>"
			.Write "  <td width='40' align='center'>选择</td>"
			.Write "  <td align='center'>打印</td>"
			If MobileCode=1 Then
			.Write "  <td align='center'>手机号</td>"
			End If
			If IsArray(SQL) Then
				For ii=0 To Ubound(SQL,2)
				  .Write "<td align='center' nowrap>" & SQL(0,II) & "</td>"
				Next
			End If
			.Write "<td align=center nowrap>提交时间</td>"
			.Write "<td align=center nowrap>状态</td>"
			.Write "<td align=center nowrap>↓管理操作</td>"
			.Write "</tr>"
			 rs.open "select * from " & TableName & " order by adddate desc" ,conn,1,1
			 If Not RS.EOF Then
					        totalPut = Conn.Execute("Select count(1) From " & TableName)(0)
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							
							  dim k,i:i=1
							   dim rsf:set rsf=server.CreateObject("adodb.recordset")
							  do while not rs.eof
							   response.write "<tr bgcolor='#ffffff'><td width=40 align='center' nowrap><input type='checkbox' name='id' value='" & rs("id") & "'></td>"
							   response.write "<td align='center' nowrap><a href=""KS.Form.asp?ItemID="& id &"&id="& rs("id") &"&Action=print_bm"" target=""_blank"" title='打印本条记录'><img src='../../../images/default/print.jpg' border='0'/></a></td>"
							   If MobileCode=1 Then
							   response.write "<td align='center' nowrap>" & rs("mobile") & "</td>"
							   End If
							   If IsArray(SQL) Then
							    For II=0 To Ubound(SQL,2)
								  response.write "<td>&nbsp;" & rs(trim(sql(1,ii))) & "</td>"
								Next
							   End If
								.Write "<td align=center nowrap>" & formatdatetime(rs("adddate"),2) & "</td>"
								.Write "<td align=center nowrap>"
								select case rs("status")
							   case 0
								response.write "<font color=red>未读</font>"
							   case 1
								response.write "<font color=green>已读</font>"
							   case 2
								response.write "<font color=#ff6600>采纳</font>"
							   case 3
								response.write "垃圾"
							   end select
								.Write "</td>"
							   response.write "<td  class='splittd' nowrap align='center'>"
							   
							   response.write "<a href=""?action=delinfo&FormID=" & ID&"&id=" & rs("Id") & """ onclick=""return(confirm('确定删除吗?'))"">删</a> <a href=""?action=modifyinfo&FormID=" & ID&"&id=" & rs("Id") & """>改</a> <a href='?action=setstatus&v=1&FormID=" & ID&"&id=" & rs("id") & "' title='设为已读'>已读</a> <a href='?action=setstatus&v=2&FormID=" & ID&"&id=" & rs("id") & "' title='设为采纳'>采纳</a> <a href='?action=setstatus&v=3&FormID=" & ID&"&id=" & rs("id") & "' title='设为垃圾'>垃圾</a> <a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");"">回复</a>"
							   if not isnull(rs("note")) and rs("note")<>"" then response.write "&nbsp;&nbsp;<a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");""><font color=blue>已回复</font></a>"

							   response.write "</td>"
							   response.write "</tr>" 
							  rs.movenext
							  i=i+1
							  if i>maxperpage then exit do
							  loop

			 End If
			  .Write "<tr><td height='36' colspan=100 bgcolor='#ffffff'><label><input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"">选中</label> <input type='submit' class='button' value='批量采纳' onclick=""$('#action').val('setstatus');$('#v').val(2);""/> <input type='submit' class='button' value='批量设置成无效记录' onclick=""$('#action').val('setstatus');$('#v').val(3);""/>  <input type='submit' class='button' value='批量设置成已读' onclick=""$('#action').val('setstatus');$('#v').val(1);""/> <input type='submit' class='button' value='批量删除' onclick=""if (confirm('此操作不可逆,确定删除选中的记录吗?')){$('#action').val('delinfo');}else{return false}""/> <input type='button' class='button' onclick='window.print();' value='打印本页面记录'></td></tr></form>"
			  .Write ("</table>")
			  .Write ("<br/>")
			   Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			  .Write ("<br>")
		 %>
		<SCRIPT language=javascript>
		function unselectall()
		{
			if(document.myform.chkAll.checked){
			document.myform.chkAll.checked = document.myform.chkAll.checked&0;
			} 	
		}
		
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"  && e.disabled==false)
			   e.checked = form.chkAll.checked;
			}
		}
		</SCRIPT>
		 		 <div style="clear:both"></div>
        </div>
		 
		 <form name="export" action="KS.Form.asp?action=export" method=post target="_blank">
		  <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
                  <input type="hidden" value="<%=id%>" name="id">
		   <strong>按时间段导出Excel</strong>
		   开始时间:<input type="text" name="startdate" class="textbox" size="26" value="<%=dateadd("d",now,-30)%>">
		   结束时间:<input type="text" name="enddate" class="textbox" size="26" value="<%=formatdatetime(now,2)%>">
		   <input type="submit" class="button" value="导出Excel">
		   <input type="button" class="button" value="全部导出Excel" onClick="window.open('KS.Form.asp?action=export&id=<%=id%>')">
		  </div>
		  </form>
		 
		 <%
			  .Write "</div>"
         end with
		End Sub
		
		'修改记录
		Sub modifyinfo()
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		  Dim FormID:FormID=KS.ChkClng(KS.S("FormID"))
		  Dim Title,TableName,SQL,ii
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 FormName,TableName From KS_Form Where ID=" & FormID,Conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.Close :Set RS=Nothing
			KS.AlertHintScript "对不起,出错啦!"
		  End If
		  Title=RS(0) : TableName=RS(1)
		  RS.Close 
		  RS.Open "Select Title,FieldName,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,AllowFileExt,MaxFileSize,FieldID From KS_FormField Where ItemID="& FormID,conn,1,1
		  If Not RS.Eof Then SQL=RS.GetRows(-1)
		  RS.Close
		  
		  If ID<>0 Then
			  RS.Open "Select top 1 * From " & TableName & " Where ID=" & ID,conn,1,1
			  If RS.Eof And RS.Bof Then
				RS.Close :Set RS=Nothing
			  End If
		  End If
		  %>
		  <div class="tabTitle">
		  <%if id=0 then
		    response.write "添加"
			else
			response.write "修改"
			end if
		%>表单[<span style='color:red'><%=Title%></span>]的提交记录</div>
		   <form name="myform" action="KS.Form.asp" method="post">
		  <table width='99%' align="center" class="ctable" border='0' cellspacing='1' cellpadding='1'>
		   <input type="hidden" value="DoResultSave" name="action">
		   <input type="hidden" value="<%=ID%>" name="id">
		   <input type="hidden" value="<%=formid%>" name="formid">
		    <%
			If IsArray(SQL) Then
			   For II=0 To Ubound(SQL,2)
			 %>
		  <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" style="text-align:right"><div><strong><%=SQL(0,II)%>：</strong></div></td>      
			<td height="30"> 
			
			<%
			Dim O_Arr,O_Len,n,F_V,O_Value,O_Text,BRStr,FieldValue
			if ID<>0 Then
			FieldValue=RS(Trim(SQL(1,II)))
			Else
			FieldValue=SQL(4,II)
			End If
			Select Case SQL(3,ii)
				Case 2
				  Response.Write "<textarea class=""textbox"" style=""width:" & SQL(7,ii) & "px;height:" & SQL(8,ii) &"px"" rows=""5"" name=""" & SQL(1,ii) & """>" & FieldValue & "</textarea>"
			   Case 3
				  Response.Write "<select class=""upfile"" style=""width:" & SQL(7,ii) & "px"" name=""" & SQL(1,ii) & """>"
				  O_Arr=Split(SQL(5,ii),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
					If O_Arr(N)<>"" Then
						F_V=Split(O_Arr(N),"|")
						If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						Else
							O_Value=F_V(0):O_Text=F_V(0)
						End If						   
						If trim(FieldValue)=trim(O_Value) Then
							Response.Write "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
						Else
							Response.Write "<option value=""" & O_Value& """>" & O_Text & "</option>"
						End If
					End If
				  Next
				  Response.Write "</select>"
			  Case 6
				  O_Arr=Split(SQL(5,ii),vbcrlf): O_Len=Ubound(O_Arr)
				  If O_Len>1 And Len(SQL(5,I))>50 Then BrStr="<br>" Else BrStr=""
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If trim(FieldValue)=trim(O_Value) Then
						Response.Write "<input type=""radio"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
					Else
						Response.Write "<input type=""radio"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """>" & O_Text & BRStr
					End If
				   End If
				  Next
			 Case 7
				   O_Arr=Split(SQL(5,ii),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If KS.FoundInArr(trim(FieldValue),O_Value,",")=true Then
						Response.Write "<input type=""checkbox"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """ checked>" & O_Text
					Else
						Response.Write "<input type=""checkbox"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """>" & O_Text
					End If
					End If
				  Next
			 Case 10
			 
			 	 Response.Write "<script id=""" & SQL(1,ii) &""" name=""" & SQL(1,ii) &""" type=""text/plain"" style=""width:70%;height:150px;"">" &FieldValue&"</script>"
	             Response.Write "<script>setTimeout(""var editor" & SQL(1,ii) &" = " & GetEditorTag() &".getEditor('" & SQL(1,ii) &"',{toolbars:[" & GetEditorToolBar("Basic") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:250 });"",10);</script>"

			 
			
			 Case Else
				Response.Write "<input type=""text"" class=""textbox"" style=""width:" & SQL(7,ii) & "px"" name=""" & SQL(1,ii) & """ value=""" & FieldValue & """>"
			End Select
			%>
			
			</td> 
		 </tr>
		    <%Next
		   End If
		   %>
		   <tr> 
		    <td class='tdbg' colspan=3 style="text-align:center">
			  <input type="hidden" name="comeurl" value="<%=Request.ServerVariables("HTTP_REFERER")%>"/>
			  <input type="submit" value="提交保存" class="button"/>
			</td>
		   </tr>
		  </table>
		   </form>
		  <br/><br/>
		  <%
		  if ID<>0 Then
			  RS.Close
			  Set RS=Nothing
		  End If
		End Sub
		
		'保存表单提交结果的修改
		Sub DoResultSave()
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		  Dim FormID:FormID=KS.ChkClng(KS.S("FormID"))
		  Dim Title,TableName,SQL,ii
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 FormName,TableName From KS_Form Where ID=" & FormID,Conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.Close :Set RS=Nothing
			KS.AlertHintScript "对不起,出错啦!"
		  End If
		  Title=RS(0) : TableName=RS(1)
		  RS.Close 
		  RS.Open "Select Title,FieldName From KS_FormField Where ItemID="& FormID,conn,1,1
		  If Not RS.Eof Then SQL=RS.GetRows(-1)
		  RS.Close
		  RS.Open "Select top 1 * From " & TableName & " Where ID=" & ID,conn,1,3
		  If RS.Eof And RS.Bof Then
		    RS.AddNew
			RS("Status")=1
			RS("AddDate")=Now
		  End If
		  For Ii=0 To Ubound(SQL,2)
		    RS(Trim(SQL(1,II)))=KS.G(Trim(SQL(1,II)))
		  Next
		   RS.Update
		   RS.Close
		   Set RS=Nothing
		   if id=0 then
			   Response.Write "<script>alert('恭喜,添加成功!');location.href='KS.Form.asp?ItemID=" & FormID&"&action=resulthp';</script>"
		   else
			   If KS.G("ComeUrl")<>"" Then
			   Response.Write "<script>alert('恭喜,修改成功!');location.href='" & Request("comeurl") &"';</script>"
			   Else
			   Response.Write "<script>alert('恭喜,修改成功!');location.href='KS.Form.asp?ItemID=" & FormID&"&action=resulthp';</script>"
			   End If
		   end if
		 
		End Sub
		
		Sub Replay()
		 on error resume next
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim ID:ID=KS.ChkClng(KS.G("id"))
		 Dim TableName:TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & FormID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From " & TableName &" Where ID=" & ID,conn,1,1
		 If RS.Eof Then
		  response.end
		 End If
         %>
		 <iframe src="about:blank" style="display:none" name="hiddenframe"></iframe>
		 <form action="KS.Form.asp?action=replaysave&formid=<%=formid%>&id=<%=id%>" method="post" name="myform" target="hiddenframe">
		  <br>
		  <div style="margin:6px;text-align:center;font-weight:bold;color:red">查看回复</div>
		  <table width='99%' align='center' border='0' cellpadding='1'  cellspacing='1' class='ctable'> 
		  <tr class="tdbg">
		    <td align="right" class="clefttitle">发表时间</td>
			<td><%=rs("adddate")%></td>
		  </tr>
		    
		  <tr class="tdbg">
		   <td align="right" class="clefttitle">回复内容：</td>
		   <td><%
		   Response.Write EchoEditor("content",rs("note")&"","Basic","96%","200px")
		   %>
		   </td>
		  </tr>
		  <tr class="tdbg">
		    <td align="right" class="clefttitle">发送邮件</td>
			<td><label><input type="checkbox" name="sendmail" value="1" checked="checked">将回复内容发送到用户邮箱</label>
			
			&nbsp;<span style='color:#999999'>填表单时要求客户填邮件的才有效！</span>
			</td>
		  </tr>
		  <tr  class="tdbg">
		    <td colspan="2" height="35" style="text-align:center"><input type="submit" class="button" value="提交回复">&nbsp;<input type="button" class="button" value="关闭窗口" onClick="top.box.close()"></td>
		  </tr>
		  </table>
		 </form>
		 <%
		  RS.Close:Set RS=Nothing
		End Sub
		
		Sub setstatus()
		 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
		 If Id="" Then KS.AlertHintScript "对不起,你没有选择!"
		 conn.execute("update " & LFCls.GetSingleFieldValue("Select TableName From KS_Form Where ID=" & KS.ChkClng(KS.G("FormID"))) &" set status=" & ks.chkclng(ks.g("v")) & " where id in(" & id &")")
		 response.redirect request.servervariables("http_referer")
		End Sub
		
		'根据内容获取上传文件名
	Public Function GetFilesList(Content,Exts)
	    If KS.IsNul(Content) Then Exit Function
		Dim re, UpFile, BFU, FileName,SaveFileList,FileExt
		FileExt="gif|jpg|bmp|png|doc|rar|docx|zip|exe|txt|mp3|flv|wma|jpeg|swf|wps|xls|xlsx|" & Exts
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(\/" & replace(KS.Setting(91),"/","\/") &")[^(\/" & replace(KS.Setting(91),"/","\/") &")]?(.*?)[.]{1}(" & FileExt & ")"
		Set UpFile = re.Execute(Content)
		Set re = Nothing
		For Each BFU In UpFile
		  If Instr(SaveFileList,trim(BFU))=0 and len(trim(BFU))>len(KS.setting(91))+1 Then
		     if FileName="" then
			  FileName=trim(BFU)
			 Else
		      FileName=FileName & "|" & trim(BFU)
			 End If
		  End If
		   SaveFileList=SaveFileList & "," & trim(BFU)
		Next
		GetFilesList = FileName
     End Function
		
		Sub DelInfo()
		 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
		 If Id="" Then KS.AlertHintScript "对不起,你没有选择!"
		 Dim FormID:FormID = KS.ChkClng(KS.G("FormID"))
		 Dim TableName:TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" &FormID)
		 Dim UploadFields,Exts
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "select FieldName,AllowFileExt From KS_FormField Where FieldType=9 And ItemID=" & FormID,Conn,1,1
		 If Not RS.Eof Then
		   Do While Not RS.Eof
		     If UploadFields="" Then
			   UploadFields=RS(0)
			 Else
			   UploadFields=UploadFields &"," & RS(0)
			 End If
			 If Exts="" Then
			  Exts= RS(1)
			 Else
			  Exts=Exts &"|" & RS(1)
			 End If
		     RS.MoveNext
		   Loop
		 End If
		 RS.Close
		 If KS.IsNul(UploadFields) Then
		 
		 Else  '删除文件
		    Dim i,Farr:Farr=split(UploadFields,",")
			Dim TempStr
			RS.Open "select " & UploadFields & " From " & TableName&" where id in (" & id &")",conn,1,1
			If Not RS.Eof Then
			   Do While Not RS.Eof
			     For I=0 To Ubound(Farr)
				   TempStr=TempStr & "," & RS(Farr(i))
				 Next
			    RS.MoveNext
			   Loop
			   
			End If
			RS.Close
			Dim Files:Files=GetFilesList(TempStr,Exts)
			If Not KS.IsNul(Files) Then
			 Dim FileArr:FileArr=split(Files,"|")
			 For I=0 To Ubound(FileArr)
			  Call KS.DeleteFile(FileArr(i))
			 Next
			End If
		 End If
		 conn.execute("delete from " & TableName &" where id in (" & id &")")
		 response.redirect request.servervariables("http_referer")
		End Sub
		
		Sub TemplateSave()
		 Dim FormID,TContent,K
		 FormID=KS.ChkCLng(KS.G("FormID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select StepNum,PostByStep,Template,Templatebm From KS_Form Where ID=" & FormID,conn,1,3
		 If Not RS.Eof Then
		   If RS(1)=1 Then
	   		 For K=1 To RS("StepNum")
			  If K=1 Then
			  Tcontent=Request.Form("Content"&K)
			  Else
			  Tcontent=Tcontent & "$aaa$" & Request.Form("Content"&K)
			  End If
    		 Next
		   Else
		     Tcontent=Request.Form("Content1")
		   End IF
		   RS("Templatebm")=Request.Form("Templatebm")
		   RS(2)=Tcontent
		  RS.Update
		 End If
		 RS.Close:Set RS=Nothing
		 Response.Write"<script>alert('恭喜，模板修改成功!');location.href='KS.Form.asp';</script>"
		End Sub
		
		Sub ReplaySave()
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim ID:ID=KS.ChkClng(KS.G("id"))
		 Dim TableName:TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & FormID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 note,ReplyDate From " & TableName &" Where ID=" & ID,conn,1,3
		  if ks.isnul(RS(0)) then
		   rs(1)=now
		  end if
          RS(0)=Request.Form("Content")
		 RS.Update
		 RS.Close
		 
		 If KS.ChkClng(KS.G("SendMail"))=1 Then   '发邮件通知
		  Dim EmailField,Email,FormName
		  Set RS=Conn.Execute("Select top 1 FieldName From KS_FormField Where FieldType=8 and ItemID=" & FormID)
		  If Not RS.Eof Then
		     EmailField=RS(0)
			 RS.Close
			 Set RS=Conn.Execute("Select Top 1 " & EmailField & " From " & TableName & " Where ID=" & ID)
			 If Not RS.Eof Then
			    Email=RS(0)
			 End If
			 If  KS.IsValidEmail(Email) Then
			    RS.Close
			   Dim S_Content,sql,k,ReturnInfo,UpFiles
			   set rs=conn.execute("select FieldName,title,MustFillTF,FieldType,ShowUnit from ks_formfield where itemid=" & Formid & " and ShowOnForm=1 order by orderid")
			   sql=rs.getrows(-1)
			   rs.close
			   rs.open "select top 1 * From " & TableName & " Where ID=" & ID,conn,1,1
			   s_content="<table border=0 cellpadding=0 cellspacing=0>" & vbcrlf
			   for k=0 to ubound(sql,2)
				
				s_content=s_content &"<tr>" & vbcrlf
				s_content=s_content & "<td width=120 align=right>" & sql(1,k) & ":</td>" & vbcrlf
				s_content=s_content & "<td>" 
				
				s_content=s_content & rs(trim(sql(0,k)))
				
				s_content=s_content & "</td>" & vbcrlf
				s_content=s_content & "</tr>" & vbcrlf
			   next
				s_content=s_content &"</table>"
				
				FormName=Conn.Execute("select top 1 formname from ks_form where id=" & formid)(0)
				s_content="尊敬的用户，您好！<br />&nbsp;&nbsp;&nbsp;&nbsp;以下是您在<font color=""red"">"  &KS.Setting(0) & "</font>提交[" & FormName & "]的信息:<br />" & s_content & "<br/><strong>以下是本站管理员给您的答复：</strong>" & Request.Form("Content")
				
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14),KS.Setting(0) & "网给您提交[" & FormName & "]的回复!", Email,KS.Setting(0), s_content,KS.Setting(11))
			   If ReturnInfo="OK" Then
				ReturnInfo="已将提交结果发送到您的邮箱" & Email & "!"
			   Else
				ReturnInfo=""
			   End If
		   
			 
			 End If
		  End If
		  RS.Close
		 End If
		 Set RS=Nothing
		 If ReturnInfo<>"" Then
		 Response.Write "<script>alert('恭喜，提交回复成功！" & ReturnInfo & "');top.frames[""MainFrame""].location.reload();top.box.close();</script>"
		 Else
		 Response.Write "<script>alert('恭喜，提交回复成功！');top.frames[""MainFrame""].location.reload();top.box.close();</script>"
		 End If
		End Sub
		
		Sub export()
		    dim param
			Dim id:id=ks.chkclng(request("id"))
			dim startdate:startdate=request("startdate")
			dim enddate:enddate=request("enddate")
			if id=0 then ks.die "error!"
			
			Dim TableName:TableName=LFCls.GetSingleFieldValue("Select TableName From KS_Form Where ID=" & ID)
			
			param=" where 1=1"
			
			if startdate<>"" and not isdate(startdate) then
				 response.write "<script>alert('开始时间格式不正确!');window.close();</script>"
				 response.end
			end if
			if enddate<>"" and not isdate(enddate) then
				 response.write "<script>alert('结束时间格式不正确!');window.close();</script>"
				 response.end
			end if
			
				if isdate(startdate) and isdate(enddate) then
				 EndDate = DateAdd("d", 1, enddate)
				 if DataBaseType=1 then
				 param=param &" and AddDate>= '" & StartDate & "' And  AddDate <='" & EndDate & "'"
				 else
				 param=param &" and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "#"
				 end if
				else
				end if
			
			
			Response.AddHeader "Content-Disposition", "attachment;filename=" & formatdatetime(now,2)&".xls" 
			Response.ContentType = "application/vnd.ms-excel" 
			Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			
			dim sql,i
			
			dim rs:set rs=server.CreateObject("adodb.recordset")
			rs.open "Select title,fieldname From [KS_FormField] Where ItemID=" & ID & " Order by OrderID",conn,1,1
			if not rs.eof then
			 sql=rs.getrows(-1)
			end if
			rs.close
			if not isarray(sql) then
			 response.write "<script>alert('没有记录!');window.close();</script>"
			end if
			
			response.write "<table width=""100%"" border=""1"" >" 
			response.write "<tr>" 
			for i=0 to ubound(sql,2)
			response.write "<th><b>" & sql(0,i) & "</b></th>" 
			next
			response.write "<th><b>用户名</b></th>"
			response.write "<th><b>提交时间</b></th>"
			response.write "<th><b>回复内容</b></th>"
			response.write "<th><b>回复时间</b></th>"
			response.write "</tr>" 
			
			rs.open "select  * from " & TableName & " " & param & " order by id desc",conn,1,1
			do while not rs.eof
			  
			  response.write "<tr>"
			  for i=0 to ubound(sql,2) 
			  response.write "<td align=center>" & ks.htmlcode(rs(sql(1,i))) & "&nbsp;</td>" 
			  next 
			  response.write "<td align=center>" & rs("username") & "</td>"
			  response.write "<td align=center>" & rs("adddate") & "</td>"
			  response.write "<td align=center>" 
			  if ks.isnul(rs("note")) then response.write "-" else response.write rs("note")
			  response.write "</td>"
			  response.write "<td align=center>" 
			  if ks.isnul(rs("ReplyDate")) then response.write "-" else response.write rs("ReplyDate")
			  response.write "</td></tr>" 
			  rs.movenext
			loop
			rs.close
			
			
			response.write "</table>"

		End Sub
		
		Sub Import()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim Title
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Form Where ID=" & id,Conn,1,1
		 If RS.Eof And RS.Bof Then 
		   RS.Close:Set RS=Nothing
		   KS.AlertHintScript "参数出错啦!"
		 End If
		 Title=RS("FormName")
		 RS.Close :Set RS=Nothing
		%>
		<div class="sort" style="line-height:30px">批量导入Excel数据到表单[<font color=red><%=title%></font>]</div>
			<form name="myform" action="?Action=ImportNext" method="post" enctype="multipart/form-data">
			<input type="hidden" name="id" value="<%=id%>"/>
			<input type="hidden" name="title" value="<%=title%>"/>
			<table width="100%" style="margin-top:10px" border="0" align="center"  cellspacing="1" class='ctable'>
			  
			  <tr class='tdbg'> 
			    <td height="25" align='right' class='clefttitle'><strong>选择要导入的Excel文件:</strong></td>
				<td><input name='FilePath' type='file' class='textbox'  size=20></td>
              </tr>
			  <tr class='tdbg'> 
			    <td height="25" align='right' class='clefttitle'><strong>输入Excel的表名称:</strong></td>
				<td><input name='tablename' type='text' class='textbox' id='tablename' value="Sheet1$" size=20>
				<span class="tips">Tips:表名后面要以$结束。</span>
				</td>
              </tr>
		 <tr class='tdbg'>
		    <td colspan=2 height='30'><b>说明：</b>
			<br/>
			1、请将要导入的Excel文件上传到网站上，然后输入正确的Excel路径。
			<br/>
			2、请按格式整理好excel数据，格式如下：
			<br/>
			<!--<div style="width:800px;padding-bottom:35px;overflow-x: auto; height:auto">-->
			<table width="100%" border="1"  cellpadding="0" cellspacing="0"><tr>
			<%
			Dim SQL,ii
			 Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select Title,FieldName From KS_FormField Where ItemID=" & ID,Conn,1,1
			 If Not RS.Eof Then SQL=RS.GetRows(-1)
			 RS.Close : Set RS=Nothing
			 If IsArray(SQL) Then
			   For II=0 To Ubound(SQL,2)
			    response.write "<th height=""25""><b>" & sql(0,ii) &"</b></th>"
               Next
			 End If
			%>
			</tr></table>
			<!--</div>-->
			<br/>
			
			<br><div align='center'> <input type="submit" class="button" name="button1" value="下一步"> 
				  &nbsp; <input type="reset" class="button" name="button2" value=" 重置 "> </div></td>
		 </tr>
			</table>
			  </form>
		<%
		End Sub
		
Sub OpenImporIConn()
				   if not isobject(IConn) then
					on error resume next
					Set IConn = Server.CreateObject("ADODB.Connection")
					IConn.open IConnStr
					If Err Then 
					  Set IConn = Nothing
					  
					  Response.Write "<script>top.$.dialog.alert('数据源连接失败,请检查数据库连接!',function(){history.back();});</script>"
					  Err.Clear
					  response.end
					end if
				   end if		
	End Sub
'**************************************************
	'过程名：ShowChird
	'作  用：显示指定数据表的字段列表
	'参  数：无
	'**************************************************
	Sub ShowField(fieldname,dbname)
	        if dbname="" then
			 response.write "<script>alert('表名称必须输入！');history.back();</script>"
			 response.end
			end if
		    dim rs:Set rs=Iconn.OpenSchema(4)
			Do Until rs.EOF or rs("Table_name") = trim(dbname)
				rs.MoveNext
			Loop
            Do Until rs.EOF or rs("Table_name") <> trim(dbname)
			  if fieldname=trim(rs("column_name")) then
				response.write "<option value='"&rs("column_Name")&"' selected>·"&rs("column_Name")&"</option>"
			  else
				response.write "<option value='"&rs("column_Name")&"'>·"&rs("column_Name")&"</option>"
			  end if
				rs.MoveNext
			loop
			rs.close:set rs=nothing
	End Sub

Sub importNext()

            Dim fobj:Set FObj = New UpFileClass
		    FObj.GetData
			
			Dim TableName:TableName=Fobj.Form("tablename")
			Dim Title:Title=Fobj.Form("Title")
			
            Dim MaxFileSize:MaxFileSize = 1000000   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "xls|xlsx"
			Dim FormPath:FormPath =KS.Setting(3) & KS.Setting(91) &"temp/"
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"Form" & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & KS.MakeRandom(5))
			
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):response.End()
			End Select
			
			
			If KS.IsNul(ReturnValue) Then 
		     KS.AlertHintScript "excel数据库没有上传!"
		   End If
		   if tablename="" then
		     KS.AlertHintScript "请输入excel数据名!"
		   end if

         Dim ID:ID=KS.ChkClng(Fobj.Form("ID"))
		 Set Fobj=Nothing
		 
		 If ID=0 Then KS.Die "<script>alert('error!');history.back();</script>"
         Dim FilePath:FilePath=ReturnValue
		 IConnStr="driver={microsoft excel driver (*.xls)};dbq=" & Server.Mappath(FilePath)
		 OpenImporIConn()
		 %>
		 	<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<link href="../../Include/Admin_Style.css" rel="stylesheet">
			<script language="JavaScript" src="../../Include/Common.js"></script>
           </head>
			<body topmargin="0" leftmargin="0">
			<div class="sort" style="line-height:30px">批量导入数据到表单[<font color=red><%=Title%></font>](配置导入项)</div>
			<form name="myform" action="?Action=ImportNext2" method="post">
			 <input type="hidden" value="<%=id%>" name="id">
			 <input type="hidden" value="<%=FilePath%>" name="FilePath">
			 <input type="hidden" value="<%=TableName%>" name="tablename">
			 <input type="hidden" value="<%=Title%>" name="title">
			<table width="100%" style="margin-top:10px" border="0" align="center"  cellspacing="1" class='ctable'>
			 <%
			 Dim RS,SQL,ii
			 Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select Title,FieldName From KS_FormField Where ItemID=" & ID,Conn,1,1
			 If Not RS.Eof Then SQL=RS.GetRows(-1)
			 RS.Close : Set RS=Nothing
			 If IsArray(SQL) Then
			   For II=0 To Ubound(SQL,2)
			 %>
			  <tr class='tdbg'> 
			    <td height="25" align='right' class='clefttitle'>
				<select name='<%=SQL(1,II)%>_Y'>
				<option value="0">-此项不导入-</option>
				<%Call ShowField(SQL(0,II),TableName)%>
				</select> =>	</td>
				<td><%=SQL(0,II)%>(<%=SQL(1,II)%>)</td>
			  </tr>
			 <%Next
			 End If
			 %> 
			  
		 <tr class='tdbg'>
		    <td colspan=2 height='30'><br/><b>说明：</b>请正确配置以上字段对应,然后点下一步开始导入操作。<br/><br><div align='center'> <input type="submit" class="button" name="button1" value="下一步"> 
				  &nbsp; <input type="reset" class="button" name="button2" value=" 重置 "> </div></td>
		 </tr>
			</table>
			  </form>
			</body>
			</html>
<%
end sub

Sub ImportNext2()
%>
<div class="sort" style="line-height:30px">批量导入数据到表单[<font color=red><%=Request("Title")%></font>](正在执行导入)</div>
		<div style="text-align:center">			 
			 <div style="margin:0 auto;margin-top:50px;border:1px dashed #cccccc;width:500px;height:80px">
			 <br>
			<div id="message">
			  <br>操作提示栏！
			</div>
			</div>
	    </div>
		<br/><br/><br/>
	 <%
	     'On Error Resume Next
		 Server.ScriptTimeOut=999999
	     Dim TableName:TableName="[" & request("tablename") & "]"
		 Dim N,FoundErr,Total,ErrNum:ErrNum=0
		 Dim t:t=0
		 Dim SQL,II,msg,ToTableName
         Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 If ID=0 Then KS.Die "<script>alert('error!');history.back();</script>"
		 Dim FilePath:FilePath=Request.Form("FilePath")
		 IConnStr="driver={microsoft excel driver (*.xls)};dbq=" & Server.Mappath(FilePath)
		 OpenImporIConn()
		 Dim IRS:Set IRS=Server.CreateOBject("ADODB.RECORDSET")
    	 Dim RS:Set RS=Server.CreateObject("ADODB.RecordSet")
		 RS.Open "Select Top 1 TableName From KS_Form Where ID=" & ID,Conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close :Set RS=Nothing
		   KS.AlertHintScript "error!"
		 End If
		 ToTableName=RS(0)
		 RS.Close
		 RS.Open "Select Title,FieldName From KS_FormField Where ItemID=" & ID,Conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close 
			 
		 IRS.Open "Select * From " & TableName,iConn,3,3

		 Total=IRS.RecordCount
		 Do While Not IRS.Eof
		   
		   t=t+1
		   FoundErr=false
		   
		'   If founderr=true Then
		'   	  response.write msg 
		 '  end if
		   
				 RS.Open "Select top 1 * From " & ToTableName &" Where 1=0",conn,1,3
				 If RS.Eof and RS.Bof Then
				   RS.AddNew
				   
				   If IsArray(SQL) Then
				     For II=0 To Ubound(SQL,2)
					  if Request(trim(SQL(1,II)) & "_y")<>"0" then
				       'response.write "RS(" & trim(SQL(1,II) & ")=IRS(" & Request(trim(SQL(1,II)) & "_y")) & ")<br/>"
				       RS(trim(SQL(1,II)))=IRS(trim(Request(trim(SQL(1,II)) & "_y")))
					  end if
					 Next
				   End If
                   RS("AddDate")=Now
				   RS("status")=1
				   RS.Update
					 N=N+1
				Else
				 ErrNum=ErrNum+1
				End If
				RS.Close
		    'Else
			'   ErrNum=ErrNum+1
			'End If
		  	Response.Write "<script>document.all.message.innerHTML='<br>共<font color=red>" & Total & "</font> 条数据，正在导入第<font color=red>" & n & "</font>条！出错跳过<font color=blue>" & ErrNum & "</font>条!';</script>"
			Response.Flush

		  IRS.MoveNext
		  If t>=Total Then Exit Do
		 Loop
		 IRS.Close:Set IRS=Nothing:Set RS=Nothing
		 Response.Write "<script>document.all.message.innerHTML='<br>恭喜22！成功导入 <font color=red>" & N & "</font> 条数据！出错" & errnum &" 条';</script>"
		 IConn.close
		 Set IConn=Nothing
		 Call KS.DeleteFile(FilePath)
		 if msg<>"" then
		   response.write "<strong>以下记录重复没有再导入:</strong><br/><font color=red>" & msg & "</font>"
		 end if
		 
End Sub	

sub print_bm()
	 	 Dim RS,Template,FormName,PostByStep,StepNum,Step,K,SQL,II,sstr
		 Dim FormID:FormID=KS.ChkClng(KS.G("ItemID"))
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select TableName,PostByStep,StepNum,Templatebm From KS_Form Where ID=" & FormID,conn,1,1
		 If RS.EOF And RS.Bof Then
		  Response.Write "<script>alert('error!');history.back();</script>"
		  Exit Sub
		 Else
		   FormName=RS(0):PostByStep=RS(1):StepNum=RS(2):Template=RS(3)
		 End If
		 RS.Close
		 set rs=server.createobject("adodb.recordset")
		 rs.open "Select Title,FieldName,FieldType From KS_FormField Where ShowOnManage=1 And ItemID=" & FormID & " Order By OrderID,FieldID",Conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
         If Template="" Or IsNull(Template) Then Response.end():Template=" "
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select top 1 * From "&FormName&" Where ID=" & ID,conn,1,1
		 If not RS.EOF And not RS.Bof Then
			 %>
			<script type="text/javascript">
            	$(document).ready(function(){
					<%
					dim str_wh
					If IsArray(SQL) Then
						For II=0 To Ubound(SQL,2)
							if ks.isnul(rs(trim(sql(1,ii)))) then sstr="&nbsp;" else  sstr=cstr(rs(trim(sql(1,ii))))
							if KS.ChkClng(sql(2,ii))=9 then
								response.write "$('#"&trim(sql(1,ii))&"').html('<img class=""px"" src="""& sstr &""" />');" &vbcrlf								
							else
								response.write "$('#"&trim(sql(1,ii))&"').html('"&sstr&"');" &vbcrlf
							end if
						Next
					End If
					%>
				});
				
            </script>
            <style>
            img.px{ width: 180px; height:200px;}
			img.pd{ width: 633px; height:322px;}
            </style>
            <body>
				<div style="height:20px; overflow:hidden"></div>
				
			<%
		 	Response.Write(Template)
			Response.Write"<div style='text-align:center;margin-top:10px' id='p'><input type='button' class='button' onclick='$(""#p"").hide();window.print();' value='打印本页面记录'></div>"
			%>
            <div style="height:20px; overflow:hidden"></div>

            </body>
			<%
		 end if
		 RS.Close
	Response.end()
end sub
		
		
End Class
%> 

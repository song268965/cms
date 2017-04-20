<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<!--#include file="../../plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_System
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_System
        Private KS,KSMCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSMCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call KS.DelCahe(KS.SiteSn & "_Config")
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 If KS.G("action")="balance" Then balance
		 With Response
		  	.Write "<!DOCTYPE html><html>"
			.Write "<title>网站基本参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.Write "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>"
			.Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<style type=""text/css"">"
			.Write "<!--" & vbCrLf
			.Write ".STYLE1 {color: #FF0000}" & vbCrLf
			.Write ".STYLE2 {color: #FF6600}" & vbCrLf
			.Write ".tips {color: #999999;padding:2px;font-size:14px;}" & vbCrLf
			.Write ".txt {color: #666;border:1px solid #ccc;height:22px;line-height:22px}" & vbCrLf
			.Write "textarea {color: #666;border:1px solid #ccc;}" & vbCrLf
			.Write "-->" & vbCrLf
			.Write "</style>" & vbCrLf
			.Write "</head>" & vbCrLf

		  Select Case KS.G("Action")
		   Case "DelQianDao"
		       Conn.Execute("Delete From KS_Qiandao")
		       Conn.Execute("Update KS_User set qiandao=0,qiandao_m=0")
			   KS.Die "<script>alert('恭喜，记录已清空！');location.href='KS.Setting.asp';</script>"
		   Case "Space"
		     	If Not KS.ReturnPowerResult(0, "KMST10010") Then          
				   Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
				   Call KS.ReturnErr(1, ""): Exit Sub
				Else
		           Call GetSpaceInfo()
				End If
		   Case "CopyRight"
		     	If Not KS.ReturnPowerResult(0, "KMST10011") Then         
				   Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
				   Call KS.ReturnErr(1, ""): Exit Sub
				Else
		           Call GetCopyRightInfo()
				End If
		   Case Else
		       Call SetSystem()
		  End Select
		 End With
		End Sub
		
		
		
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		With Response
			
					If Not KS.ReturnPowerResult(0, "KMST10001") Then          '检查是否有基本信息设置的权限
					 .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back()';</script>")
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			
			dim strDir,strAdminDir
			strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
			strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-2) & "/"
			InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
			If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
			   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
			End If
	
	
			SqlStr = "select * from KS_Config"
			Set RS = KS.InitialObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			
			 Dim Setting:Setting=Split(RS("Setting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
			 
			 Dim TBSetting:TBSetting=Split(RS("TBSetting"),"^%^")
			 FsoIndexFile = Split(Setting(5), ".")(0)
			 FsoIndexExt = Split(Setting(5), ".")(1)
			 Dim FilterWord:FilterWord=KS.ReadFromFile("../../config/filter.txt")
			 
			If KS.G("Flag") = "Edit" Then
			      
					
				Dim FZCJYM,n,i
				For N=1 To 10
				  FZCJYM=FZCJYM & KS.ChkClng(Request.Form("Opening" & N))
				Next
				
				if instr(Request.Form("Setting(178)"),".")<>0 or instr(Request.Form("Setting(90)"),".")<>0 or instr(Request.Form("Setting(91)"),".")<>0 or instr(Request.Form("Setting(93)"),".")<>0 or instr(Request.Form("Setting(94)"),".")<>0 or instr(Request.Form("Setting(95)"),".")<>0 or instr(Request.Form("Setting(96)"),".")<>0 then
			     KS.Die ("<script>alert('对不起，相关目录设置里的目录不能含有“.”！');history.back();</script>")
				end if
				
			    Dim WebSetting,ThumbSetting,TempStr
				
				Call KS.WriteTOFile("../../config/filter.txt",KS.S("FilterWord"))
				
				Dim Alen:Alen=Ubound(Setting)
				If Alen>300 Then Alen=300   '设置数组最大上标为300
				
				For n=0 To Alen
				  If n=5 Then
				   WebSetting=WebSetting & KS.G("Setting(5)") & KS.G("FsoIndexExt") & "^%^"
				  ElseIf N=14 Then
				   WebSetting=WebSetting & KS.Encrypt(request("Setting(14)")) & "^%^"
				  ElseIf N=155 Then
				     Dim Si
				     For Si=0 To 60
					  WebSetting=WebSetting &Request("Sms" & Si) & "∮"
					 Next
					 WebSetting=WebSetting & "^%^"
				  ElseIf N=158 Then
				  	WebSetting=WebSetting &Request("smsye") & "∮" & Request("smsyetag1")& "∮" & Request("smsyetag2") & "^%^"
				  ElseIf N=153 Then
				    WebSetting=WebSetting &Request("smspass") & "∮" & Request("smspassmd5") & "^%^"
				  ElseIf N=156 Then
				    WebSetting=WebSetting &Request("smssign1") & "∮" & Request("smssign2")& "∮" & Request("smssign3")& "∮" & Request("smssign4")& "∮" & Request("smssign5") & "^%^"
				  ElseIf n=161 Then
				   WebSetting=WebSetting & FZCJYM & "^%^"
				  ElseIf N=170 Then
				    TempStr=""
				    For i=1 to 6
					 If Request.Form("Setting(170" & i & ")")="1" Then
					  TempStr=TempStr &"1"
					 Else
					  TempStr=TempStr &"0"
					 End If
					Next
					 WebSetting=WebSetting & TempStr & "^%^"
				  ElseIf Request.Form("Setting(" & n &")")<>"" or n=16 or n=27 or n=34 or n=189 or n=22 or n=176 or n=101 or n=154 or n=186 Then
				   WebSetting=WebSetting & Replace(Request.Form("Setting(" & n &")"),"^%^","") & "^%^"
				  Else
				   WebSetting=WebSetting   & Setting(n) & "^%^"
				  End If
				Next
				
				For I=0 To 20
				 If I=13 Then
				  ThumbSetting=ThumbSetting & Replace(KS.G("TBLogo"),"^%^","") & "^%^"
				 Else
				  ThumbSetting=ThumbSetting & Replace(KS.G("TBSetting(" & I &")"),"^%^","") & "^%^"
				 End If
				Next
				RS("Setting")=WebSetting
				RS("TBSetting")=ThumbSetting
				RS.Update
				Call KS.FileAssociation(1015,1,WebSetting&ThumbSetting,1)
				RS.Close:Set RS=Nothing
				
                '修改模板目录
				if lcase(Request.Form("OldTemplate"))<>lcase(Request.Form("Setting(90)")) then
			    	if FolderReName(KS.Setting(3) & Request.Form("OldTemplate"),Replace(Request.Form("Setting(90)")&"","/",""))=true then
					end if
				End If
								
				'修改后台目录
				if lcase(Request.Form("OldAdmin"))<>lcase(Request.Form("Setting(89)")) then
			    	if FolderReName(KS.Setting(3) & Request.Form("OldAdmin"),Replace(Request.Form("Setting(89)")&"","/",""))=true then
			         KS.Die ("<script>alert('网站配置信息修改成功！');top.location.href='" & KS.Setting(3) & Request.Form("Setting(89)") & "index.asp?C=1&from=KS.Setting.asp';</script>")
					end if
				End If
				
				Session("FromFile")="System/KS.Setting.asp"
				KS.AlertDoFun "网站配置信息修改成功！","top.location.reload();"
			End If
			
			.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=SetParam&OpStr=" & Server.URLEncode("系统设置 >> <font color=red>基本信息设置</font>") & "';</script>")
			%>
            <style>
			

			</style>
            
            <%

			.Write "<body  bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='tabTitle mt20'>系统参数配置</div>"
			.Write "<div class=tab-page id=configPane>"
			.Write " <form name='myform' method=post action=""KS.Setting.asp"" id=""myform"">"
			.Write " <input type=""hidden"" value=""Edit"" name=""Flag""/>"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""configPane"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>基本信息</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			
			.Write "<dl class=""dtable"">"
			
			 .Write "<dd><div>网站状态：</div><input onclick=""$('#webstatus').hide();"" type=""radio"" name=""Setting(187)"" value=""1"" "
				
				If Setting(187) = "1" Then .Write (" checked")
				.Write ">正常"
				.Write "    <input type=""radio"" onclick=""$('#webstatus').show();"" name=""Setting(187)"" value=""0"" "
				If Setting(187) = "0" Then .Write (" checked")
				.Write ">关闭"
			   .Write "<span>网站关闭后，前台将不能访问。</span>"
			   .Write "    </dd>"
				
			 .Write "    <dd id=""webstatus"" "
			 If Setting(187) <> "0" Then .Write " style='display:none'"
			 .Write "><div>网站关闭提示信息：</div><textarea name=""Setting(188)"" cols=""45"" rows=""3"">" & Setting(188) & "</textarea><span class=""block"">如：网站日常维护，请稍候访问。</span></dd>"

			
			
			
			.Write "    <dd><div>网站名称：</div><input name=""Setting(0)"" type=""text"" id=""Setting(0)"" value=""" & Setting(0) & """ size=""50"" class=""textbox""><span>可以在模板里通过{$GetSiteName}标签调用</span>"
			.Write "    </dd>"
			
			.Write "   <dd><div>网站地址：</div> <input name=""Setting(2)"" type=""text""  value=""" &KS.GetAutoDomain & """ size=""50"" class=""textbox""><span>系统会自动获得正确的路径，但需要手工保存设置。请使用http://标识),后面不要带&quot;/&quot;符号</span>"
			 .Write "   </dd>"
			 .Write "   <dd><div>安装目录：</div> <input name=""Setting(3)"" type=""text"" id=""Setting(3)""  value=""" & InstallDir & """ readonly size=""50"" class=""textbox""><span>系统会自动获得正确的路径，但需要手工保存设置。系统安装的虚拟目录</span>"
			 .Write "   </dd>"
			 .Write "   <dd><div>网站Logo地址：</div><input name=""Setting(4)"" type=""text"" id=""Logo""   value=""" & Setting(4) & """ size=""50"" class=""textbox"">"
			  .Write "   <input class=""button""  type='button' name='Submit' value='选择Logo...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=1&CurrPath=" & KS.GetUpFilesDir &"',550,290,window,document.myform.Logo);""> <span>填写本站的Logo图片地址，如/images/logo.gif</span>"
			 .Write "   </dd>"
			 .Write "   <dd><div>生成的网站首页：</div> <input type=""radio"" name=""Setting(5)"" value=""Index"" "
				
				If FsoIndexFile = "Index" Then .Write (" checked")
				.Write ">"
				.Write "    Index"
				.Write "    <input type=""radio"" name=""Setting(5)"" value=""Default"" "
				If FsoIndexFile = "Default" Then .Write (" checked")
				.Write ">"
				.Write "    Default"
				.Write "    <select name=""FsoIndexExt"" onchange=""if(this.value=='.asp'){$('#ft').hide();}else{$('#ft').show();}"" id=""select"">"
				.Write "      <option value="".htm"" "
				If FsoIndexExt = "htm" Then .Write ("selected")
				.Write ">.htm</option>"
				.Write "      <option value="".html"" "
				If FsoIndexExt = "html" Then .Write ("selected")
				.Write ">.html</option>"
				.Write "      <option value="".shtml"" "
				If FsoIndexExt = "shtml" Then .Write ("selected")
				.Write ">.shtml</option>"
				.Write "      <option value="".shtm"" "
				If FsoIndexExt = "shtm" Then .Write ("selected")
				.Write ">.shtm</option>"
				.Write "      <option value="".asp"" "
				If FsoIndexExt = "asp" Then .Write ("selected")
				.Write ">.asp</option>"
				.Write "    </select><span>扩展名为.asp，首页将不启用生成静态HTML的功能</span>"
				.Write "</dd>"
				 IF FsoIndexExt<>"asp" Then
				.Write "<dd id='ft'>"
				 else
				.Write "<dd id='ft' style='display:none'>"
				end if
				.Write " <div>首页自动生成：</div>间隔<input type='text' class='textbox' name='setting(130)' value='" & Setting(130) & "' size=4 style='text-align:center'>分钟自动生成"
				.Write "<span>设置为0将不自动生成首页</span>"
				.Write "</dd>"
				
				
				.Write "<dd"
				if KS.GetAppStatus("special")=false then .write " style='display:none'"
				.Write "><div>专题是否启用生成：</div><input type=""radio"" name=""Setting(78)"" value=""1"" "
				
				If Setting(78) = "1" Then .Write (" checked")
				.Write ">启用"
				.Write "    <input type=""radio"" name=""Setting(78)"" value=""0"" "
				If Setting(78) = "0" Then .Write (" checked")
				.Write ">不启用"
			   .Write "  　</dd>"
			
				.Write "<dd"
				if KS.GetAppStatus("tags")=false then .write " style='display:none'"
				.Write "><div>Tags启用伪静态：</div><input type=""radio"" name=""Setting(185)"" value=""1"" "
				
				If Setting(185) = "1" Then .Write (" checked")
				.Write ">启用"
				.Write "    <input type=""radio"" name=""Setting(185)"" value=""0"" "
				If Setting(185) = "0" Then .Write (" checked")
				.Write ">不启用"
			   .Write "<span>服务器需要支持rewrite组件。</span>"
			   .Write "    </dd>"
			
				.Write "<dd><div>默认允许上传最大文件大小：</div><input name=""Setting(6)"" onBlur=""CheckNumber(this,'允许上传最大文件大小');"" type=""text"" id=""Setting(6)""   value=""" & Setting(6) & """ size=10 class='textbox' style='text-align:center'>"
			.Write "KB<span>提示：1 KB = 1024 Byte，1 MB = 1024 KB"
			.Write "&nbsp;&nbsp;<input type=""checkbox"" name=""Setting(186)"" value=""1"""
			If Setting(186)="1" then .Write " checked"
			.Write "/>上传图片自动开启剪切窗口</span>"
			.Write "    </dd>"
			.Write "    <dd><div>默认允许上传文件类型：</div><input name=""Setting(7)"" type=""text"" id=""Setting(7)""   value=""" & Setting(7) & """ size='50' class='textbox'> <span>多个类型用|线隔开</span>"
			.Write "    </dd>"
			.Write "    <dd><div>删除不活动用户时间：</div><input name=""Setting(8)"" type=""text""  value=""" &  Setting(8) & """ style=""text-align:center"" size=""8"" class=""textbox""> 分钟  <span>如果在这个时间内用户没有活动,则用户的在线状态将被置为离线,值越小越精确,但消耗资源越大,建议设置在5-30分钟之间。</span>"
			.Write "    </dd>"
			.Write "    <dd style=""display:none""><div>文章自动分页每页大约字符数：</div><input name=""Setting(9)"" type=""text"" value=""" & Setting(9) & """ style=""text-align:center"" size=""8"" class=""textbox""> 个字符如果不想自动分页，请输入""0""</span>"
			.Write "    </dd>"
			.Write "    <dd><div>站长姓名：</div> <input name=""Setting(10)"" type=""text""   value=""" & Setting(10) & """ size=""50"" class=""textbox""><span>可在模板里使用{$GetWebMaster}调用</span>"
			.Write "    </dd>"


			.Write "    <dd><div>要屏蔽的关键字：</div><textarea name=""FilterWord"" cols=""45"" rows=""6"">" & FilterWord & "</textarea><span class=""block"">说明：有多个关键词要过滤，请用英文逗号隔开。<br/>作用范围所有模型的内容、评论、问答及小论坛等。</span></dd>"

			 
			 .Write "   <dd><div>页面发布时底部信息：</div> <input name=""Setting(15)"" type=""text""  value=""" & Setting(15) & """ size='50' class='textbox'><span>填写&quot;0&quot;将不显示</span>"
			 .Write "   </dd>"
			 .Write "   <dd><div>官方信息显示：</div> <input type=""checkbox"" name=""Setting(16)"" value=""1"" "
				
				If instr(Setting(16),"1")>0 Then .Write (" checked")
				.Write ">显示顶部公告"
				.Write "    <input type=""checkbox"" name=""Setting(16)"" value=""2"" "
				If instr(Setting(16),"2")>0 Then .Write (" checked")
				.Write ">显示论坛新帖"
				.Write "    <input type=""checkbox"" name=""Setting(16)"" value=""3"" "
				If instr(Setting(16),"3")>0 Then .Write (" checked")
				.Write ">右下角消息提示"

			 .Write "     </dd>"
			 .Write "   <dd><div>官方授权的唯一系列号：</div> <input name=""Setting(17)"" type=""text""  value=""" & Setting(17) & """ size='50' class='textbox'><span> 免费版请填写&quot;0&quot;</span>"
			 .Write "   </dd>"
			   
			 .Write "     <dd><div>网站的版权信息：</div> <textarea name=""Setting(18)"" cols=""45"" rows=""5"">" & Setting(18) & "</textarea><span>用于显示网站版本等，支持html语法</span>"
			 .Write "   </dd>"
		.Write "</dl>"
		.Write "</div>"
		
		.Write " <div class=tab-page id=site-seo>"
			.Write "  <H2 class=tab>SEO选项</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-seo"" ) );"
			.Write "	</SCRIPT>"
		   .Write "<dl class=""dtable"">"
           .Write "    <dd><div>网站标题：</div><input name=""Setting(1)"" type=""text"" id=""Setting(1)"" value=""" & Setting(1) & """ size=""70"" class=""textbox""><span>可以在模板里通过{$GetSiteTitle}标签调用</span>"
			 .Write "   </dd>"
			 .Write "     <dd><div>网站META关键词：</div> <textarea name=""Setting(19)"" cols=""80"" rows=""7"">" & Setting(19) & "</textarea><span class=""block"">针对搜索引擎设置的网页关键词,多个关键词请用,号分隔。可以在模板里通过{$GetMetaKeyWord}标签调用</span>"
			 .Write "   </dd>"
			 .Write "     <dd><div>网站META网页描述：</div> <textarea name=""Setting(20)"" cols=""80"" rows=""7"">" & Setting(20) & "</textarea><span class=""block"">针对搜索引擎设置的网页描述,多个描述请用,号分隔。可以在模板里通过{$GetMetaDescript}标签调用</span>"
			 .Write "   </dd>"			 
			 

			.Write "</dl>"
		.Write "</div>"
			 
			.Write " <div class=tab-page id=site-template>"
			.Write "  <H2 class=tab>模板绑定</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-template"" ) );"
			.Write "	</SCRIPT>"
			.Write "<dl class=""dtable"">"
			.Write "<dd><div>网站首页模板：</div><input class='textbox mb0' name=""Setting(110)"" id=""Setting110"" type=""text"" value=""" & Setting(110) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting110')[0]") & " <a href='../../index.asp' target='_blank'>页面:/index.asp</a>"
			.Write "    </dd>"
		
			.Write "    <dd><div>全站搜索模板：</div><input class='textbox mb0' name=""Setting(139)"" id=""Setting139"" type=""text"" value=""" & Setting(139) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting139')[0]") & " <a href='../../plus/search/' target='_blank'>页面:/plus/search/</a></dd>"			
			.Write "    <dd"
			if KS.GetAppStatus("special")=false then .write " style='display:none'"
			.Write ">"
			.Write "  <div>专题首页模板：</div><input class='textbox mb0' name=""Setting(111)"" id=""Setting111"" type=""text"" value=""" & Setting(111) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting111')[0]") & " <a href='../../item/specialindex.asp' target='_blank'>页面:/item/specialindex.asp</a></dd>"

           if KS.GetAppStatus("pk")=false then
			.Write "<font style='display:none'>"
		   Else
		    .Write "<font>"
		   End If
			.Write "    <dd><div>PK首页模板：</div><input class='textbox mb0' name=""Setting(102)"" id=""Setting102"" type=""text"" value=""" & Setting(102) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting102')[0]") & " <a href='../../plus/pk/index.asp' target='_blank'>页面:/plus/pk/index.asp</a></dd>"
			.Write "    <dd><div>PK页模板：</div><input class='textbox mb0' name=""Setting(103)"" id=""Setting103"" type=""text"" value=""" & Setting(103) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting103')[0]") & " <a href='#'>页面:/plus/pk/pk.asp</a></dd>"
			.Write "    <dd><div>PK观点更多页模板：</div><input class='textbox mb0' name=""Setting(104)"" id=""Setting104"" type=""text"" value=""" & Setting(104) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting104')[0]") & " <a href='#'>页面:/plus/pk/more.asp</a></dd>"
            .Write "</font>"
			
			
		   if not KS.ChkClng(conn.execute("select top 1 ChannelStatus from ks_channel where channelid=5")(0))=11 then
			.Write "<font style='display:none'>"
		   Else
		    .Write "<font>"
		   End If
		 
			.Write "    <dd><div>论坛首页模板：</div><input class='textbox mb0' name=""Setting(114)"" id=""Setting114"" type=""text"" value=""" & Setting(114) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting114')[0]") & " <a href='../../club/index.asp' target='_blank'>页面:/club/index.asp</a></dd>"
			.Write "    <dd><div>论坛版面列表页模板：</div><input class='textbox mb0' name=""Setting(172)"" id=""Setting172"" type=""text"" value=""" & Setting(172) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting172')[0]") & " <a href='../club/index.asp' target='_blank'>页面:/club/index.asp</a></dd>"
			.Write "    <dd><div>论坛帖子页模板：</div><input class='textbox mb0' name=""Setting(160)"" id=""Setting160"" type=""text"" value=""" & Setting(160) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting160')[0]") & " <a href='../../club/display.asp' target='_blank'>页面:/club/display.asp</a></dd>"
			.Write "    <dd><div>论坛发帖页面模板：</div><input class='textbox mb0'  name=""Setting(115)"" id=""Setting115"" type=""text"" value=""" & Setting(115) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting115')[0]") & " <a href='../../club/post.asp' target='_blank'>页面:/club/post.asp</a></dd>"
			.Write "    <dd><div>论坛搜索模板：</div><input  class='textbox mb0' name=""Setting(171)"" id=""Setting171"" type=""text"" value=""" & Setting(171) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting171')[0]") & " <a href='../../club/query.asp' target='_blank'>页面:/club/query.asp</a></dd>"
		    .Write "</font>"	
			
			.Write "    <dd><div>会员首页模板：</div><input  class='textbox mb0' name=""Setting(116)"" id=""Setting116"" type=""text"" value=""" & Setting(116) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting116')[0]") & " <a href='../../user/' target='_blank'>页面:/user/index.asp</a></dd>"
			.Write "    <dd><div>会员注册表单模板：</div><input  class='textbox mb0' name=""Setting(117)"" id=""Setting117"" type=""text"" value=""" & Setting(117) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting117')[0]") & " <a href='../../user/reg/' target='_blank'>页面:/user/reg/</a></dd>"

			.Write "    <dd><div>会员注册成功页模板：</div><input  class='textbox mb0' name=""Setting(119)"" id=""Setting119"" type=""text"" value=""" & Setting(119) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting119')[0]") & "</dd>"

			
			dim dis
			if KS.ChkClng(conn.execute("select top 1 ChannelStatus from ks_channel where channelid=5")(0))=1 then
			 dis=""
			else
			 dis=" style='display:none'"
			end if
			.Write "    <dd" & dis &">"
			.Write "     <div>商城购物车模板：</div><input  class='textbox mb0' name=""Setting(121)"" id=""Setting121"" type=""text"" value=""" & Setting(121) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting121')[0]") & " <a href='../../shop/shoppingcart.asp' target='_blank'>页面:/shop/shoppingcart.asp</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "     <div>商城收银台模板：</div><input class='textbox mb0' name=""Setting(122)"" id=""Setting122"" type=""text"" value=""" & Setting(122) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting122')[0]") & " <a href='../../shop/payment.asp' target='_blank'>页面:/shop/payment.asp</a></dd>"

			.Write "    <dd" & dis &">"
			.Write "     <div>商城订购成功模板：</div><input class='textbox mb0' name=""Setting(124)"" id=""Setting124"" type=""text"" value=""" & Setting(124) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting124')[0]") & " <a href='../../shop/order.asp' target='_blank'>页面:/shop/order.asp</a></dd>"
			
			.Write "    <dd" & dis &">"
			.Write "      <div>游客订单查询模板：</div><input class='textbox mb0' name=""Setting(173)"" id=""Setting173"" type=""text"" value=""" & Setting(173) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting173')[0]") & " <a href='../../shop/myorder.asp' target='_blank'>页面:/shop/myorder.asp</a></dd>"
			
			
			.Write "    <dd" & dis &">"
			.Write "      <div>商城购物指南模板：</div><input class='textbox mb0' name=""Setting(125)"" id=""Setting125"" type=""text"" value=""" & Setting(125) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting125')[0]") & " <a href='../../shop/ShopHelp.asp' target='_blank'>页面:/shop/ShopHelp.asp</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "      <div>商城银行付款模板：</div><input class='textbox mb0' name=""Setting(126)"" id=""Setting126"" type=""text"" value=""" & Setting(126) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting126')[0]") & " <a href='../../shop/showpay.asp' target='_blank'>页面:/shop/showpay.asp</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "     <div>商城品牌列表页模板：</div><input class='textbox mb0' name=""Setting(135)"" id=""Setting135"" type=""text"" value=""" & Setting(135) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting135')[0]") & " <a href='../../shop/showbrand.asp' target='_blank'>页面:/shop/showbrand.asp</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "     <div>商城品牌详情页模板：</div><input class='textbox mb0' name=""Setting(136)"" id=""Setting136"" type=""text"" value=""" & Setting(136) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting136')[0]") & " <a href='../../shop/brand.asp' target='_blank'>页面:/shop/brand.asp</a></dd>"
			
			.Write "    <dd" & dis &">"
			.Write "     <div>商城团购首页模板：</div><input class='textbox mb0' name=""Setting(137)"" id=""Setting137"" type=""text"" value=""" & Setting(137) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting137')[0]") & " <a href='../../shop/groupbuy.asp' target='_blank'>页面:/shop/groupbuy.asp</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "      <div>商城团购内容页模板：</div><input class='textbox mb0' name=""Setting(138)"" id=""Setting138"" type=""text"" value=""" & Setting(138) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting138')[0]") & " <a href='../../shop/groupbuyshow.asp' target='_blank'>页面:/shop/groupbuyshow.asp</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "     <div>商城团购购物车模板：</div><input class='textbox mb0' name=""Setting(120)"" id=""Setting120"" type=""text"" value=""" & Setting(120 ) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting120')[0]") & " <a href='../../shop/groupbuycart.asp' target='_blank'>页面:/shop/groupbuycart.asp</a></dd>"
			
			

			
			
			if not conn.execute("select ChannelStatus from ks_channel where channelid=9").eof then
			 if conn.execute("select ChannelStatus from ks_channel where channelid=9")(0)=1 then
			 dis=""
			 else
			 dis=" style='display:none'"
			 end if
			else
			 dis=" style='display:none'"
			end if
			.Write "    <dd" & dis &">"
			.Write "      <div>考试系统首页模板：</div><input class='textbox mb0' name=""Setting(131)""  id=""Setting131"" type=""text"" value=""" & Setting(131) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting131')[0]") & " <a href='../../mnkc/' target='_blank'>页面:/mnkc/</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "      <div>试卷分类页面模板：</div><input  class='textbox mb0' name=""Setting(132)"" id=""Setting132"" type=""text"" value=""" & Setting(132) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting132')[0]") & "</dd>"
            .Write "    <dd" & dis &">"
			.Write "     <div>试卷内容页面模板：</div><input class='textbox mb0' name=""Setting(105)"" id=""Setting105"" type=""text"" value=""" & Setting(105) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting105')[0]") & "</dd>"			
			.Write "    <dd" & dis &">"
			.Write "      <div>试卷总分类模板：</div><input class='textbox mb0' name=""Setting(134)"" id=""Setting134"" type=""text"" value=""" & Setting(134) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting134')[0]") & " <a href='../../mnkc/all.html' target='_blank'>页面:/mnkc/all.html</a></dd>"
			.Write "    <dd" & dis &">"
			.Write "      <div>日常练习模板：</div><input class='textbox mb0' name=""Setting(177)"" id=""Setting177"" type=""text"" value=""" & Setting(177) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting177')[0]") & " <a href='../../mnkc/day.html' target='_blank'>页面:/mnkc/day.html</a></dd>"
			.Write "  </dl>"
			.Write "</div>"
			
			
			 '=================================================防注册机选项========================================
			 .Write "<div class=tab-page id=ZCJ_Option>"
			 .Write " <H2 class=tab>防注册机</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""ZCJ_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "<dl class=""dtable"">"
			
             .Write "    <dd><div>要启用防注册机的页面：</div>"
			.Write "<input type='checkbox' name='Opening1' value='1'"
			If mid(Setting(161),1,1)="1" Then .Write "checked"
			.Write ">会员注册页面"
			.Write "<input type='checkbox' name='Opening2' value='1'"
			If mid(Setting(161),2,1)="1" Then .Write "checked"
			.Write ">匿名投稿发布页面"
			.Write "<input type='checkbox' name='Opening3' value='1'"
			If mid(Setting(161),3,1)="1" Then .Write "checked"
			.Write ">论坛发帖页面"
			'.Write "<br/><input type='checkbox' name='Opening3' value='1'"
			'If mid(Setting(161),3,1)="1" Then .Write "checked"
			'.Write ">评论发表页面"
		    .Write "     </dd>"			
            .Write "    <dd><div>验证问题：</div><textarea name='Setting(162)' style='width:350px;height:120px'>" & Setting(162) & "</textarea>"
			.Write "<span  class=""block"">可以设置多个,一行一个验证选项,尽量多填一些选项，更能有效防注册机的干扰。允许使用#####对问题分组，这样第一个分组将在每天1点时出现，第二个分组在每天2点时出现...最多可以设置24个分组</span>"
            .Write "    </dd><dd><div>验证答案：</div><textarea name='Setting(163)' style='width:350px;height:120px'>" & Setting(163) & "</textarea><span  class=""block"">对应验证问题的选项,一行一个验证答案</span>"
			.Write "  </dl>"
			.Write "</div>"
			
			
			
		 '=====================================================会员注册参数配置开始=========================================

		.Write " <div class=tab-page id=User_Option>"
		.Write "	  <H2 class=tab>会员选项</H2>"
		.Write "		<SCRIPT type=text/javascript>"
		.Write "					 tabPane1.addTabPage(document.getElementById( ""User_Option"" ));"
		.Write "		</SCRIPT>"
			 
			.Write "<dl class=""dtable"">"
			.Write "    <dd><div>是否允许新会员注册：</div><input onclick=""$('#userreg').show();"" name=""Setting(21)"" type=""radio"" value=""1"""
			 If Setting(21)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input onclick=""$('#userreg').hide();"" name=""Setting(21)"" type=""radio"" value=""0"""
			 If Setting(21)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>选择否将不允许会员注册</span></dd>"		
            
			.Write "<font id=""userreg"""
			If Setting(21)="0" Then .Write " style='display:None'"
			.Write ">"
			.Write "<dd><div>会员注册协议：</div><textarea name=""Setting(23)"" cols=""55"" rows=""7"">" & Setting(23) & "</textarea>"
			.Write "<span  class=""block"">标签说明：{$GetSiteName}：网站名称<br>{$GetSiteUrl}：网站URL<br>{$GetWebmaster}：站长<br>{$GetWebmasterEmail}：站长信箱</span></dd>"
			
			
			 .Write "<dd><div>是否启用用户组注册：</div><input name=""Setting(33)"" type=""radio"" value=""1"""
			 If Setting(33)="1" Then .Write " Checked"
			 .Write ">启用"
			 .Write " &nbsp;&nbsp;<input name=""Setting(33)"" type=""radio"" value=""0"""
			 If Setting(33)="0" Then .Write " Checked"
			 .Write ">不启用"
			 .Write "<span>如果不启用,默认注册类型为个人会员</span></dd>" 
			 .Write "<dd><div>注册开启详细选项： </div><label><input name=""Setting(32)"" type=""radio"" value=""1"""
			 If Setting(32)="1" Then .Write " Checked"
			 .Write ">不开启</label> "
			 .Write "<label><input name=""Setting(32)"" type=""radio"" value=""2"""
			 If Setting(32)="2" Then .Write " Checked"
			 .Write ">开启</label>"
			 .Write "<span>开启详细选项，则注册时需要填写对应用户组的注册表单</span></dd>"
			
			 .Write "<dd><div>注册成功发邮件通知：</div><input name=""Setting(146)"" type=""radio"" onclick=""setsendmail(0)"" value=""0"""
			 If Setting(146)="0" Then .Write " Checked"
			 .Write ">关闭&nbsp;&nbsp;<input name=""Setting(146)"" onclick=""setsendmail(1)"" type=""radio"" value=""1"""
			 If Setting(146)="1" Then .Write " Checked"
			 .Write ">仅发给注册人&nbsp;&nbsp;<input name=""Setting(146)"" onclick=""setsendmail(1)"" type=""radio"" value=""2"""
			 If Setting(146)="2" Then .Write " Checked"
			 .Write ">仅发给管理员&nbsp;&nbsp;<input name=""Setting(146)"" onclick=""setsendmail(1)"" type=""radio"" value=""3"""
			 If Setting(146)="3" Then .Write " Checked"
			 .Write ">同时发给管理员和注册人"
			
			 .Write "<span>用户组设置成需要邮件验证时,只有激活成功才会发送。</span></dd>"
			.Write "<dd><div>会员注册成功发送的邮件通知内容：</div><textarea name=""Setting(147)"" cols=""55"" rows=""8"">" & Setting(147) & "</textarea>"
			.Write "<span class=""block"">标签说明：<br>{$UserName}：用户名<br>{$PassWord}：密码<br>{$SiteName}：网站名称<br>{$UserInfoList}：会员注册时详细资料</span></dd>"
			
			 .Write "<dd><div>新注册密码问题必填：</div><input name=""Setting(148)"" type=""radio"" value=""1"""
			 If Setting(148)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(148)"" type=""radio"" value=""0"""
			 If Setting(148)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>开启后可以有效防止恶意注册</span></dd>"			 
			 .Write "<dd><div>新注册手机号码必填：</div><input name=""Setting(149)"" type=""radio"" value=""1"""
			 If Setting(149)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(149)"" type=""radio"" value=""0"""
			 If Setting(149)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>开启后可以有效防止恶意注册</span></dd>"
			 .Write "<dd><div>每个手机号码只能注册一次：</div></td>"
			.Write "      <td height=""21""> <input name=""Setting(129)"" type=""radio"" value=""1"""
			 If Setting(129)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(129)"" type=""radio"" value=""0"""
			 If Setting(129)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>开启后可以有效防止恶意注册</span></dd>"
			 
			
			 .Write "<dd><div>新注册启用IP限制：</div><input name=""Setting(26)"" type=""radio"" value=""1"""
			 If Setting(26)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(26)"" type=""radio"" value=""0"""
			 If Setting(26)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>若选择是，那么一个IP地址只能注册一次</span></dd>"
			 .Write "<dd><div>使用找回密码功能限制：</div>"
			.Write "  每个IP每天只能用<input type='text' name='Setting(123)' value='" & KS.ChkClng(Setting(123)) &"' class='textbox' size='4'>次找回密码功能"
			 .Write "<span>启用此功能可以防止非法用户恶意猜测得到密码，不限制请输入0。</span></dd>"
			 .Write "<dd><div>使用重发激活码功能限制：</div>每个IP每天只能用<input type='text' name='Setting(128)' value='" & KS.ChkClng(Setting(128)) &"' class='textbox' size='4'>次重发激活码功能"
			 .Write "<span>启用此功能可以防止非法用户恶意猜测激活账户，不限制请输入0。</span></dd>"
			
			 .Write "<dd><div>新注册允许上传文件：</div><input name=""Setting(60)"" type=""radio"" value=""1"""
			 If Setting(60)="1" Then .Write " Checked"
			 .Write ">允许"
			 .Write "&nbsp;&nbsp;<input name=""Setting(60)"" type=""radio"" value=""0"""
			 If Setting(60)="0" Then .Write " Checked"
			 .Write ">不允许"
			 .Write "<span>指当有自定义上传字段时，允许会员注册时同时上传文件。</span></dd>"		
			 
			  .Write "<dd><div>启用验证码：</div><label><input name=""Setting(27)"" type=""checkbox"" value=""1"""
			 If Setting(27)="1" Then .Write " Checked"
			 .Write ">注册时启用验证码</label>"
			 .Write "&nbsp;&nbsp;<label><input name=""Setting(34)"" type=""checkbox"" value=""1"""
			 If Setting(34)="1" Then .Write " Checked"
			 .Write ">登录时启用验证码</label>"
		
					 
			 .Write "<span>启用验证码功能可以在一定程度上防止暴力营销软件或注册机自动注册</span></dd>"
			 .Write "<dd><div>每个Email允许注册多次：</div><input name=""Setting(28)"" type=""radio"" value=""1"""
			 If Setting(28)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(28)"" type=""radio"" value=""0"""
			 If Setting(28)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>若选择是，则利用同一个Email可以注册多个会员。</span></dd>"
			.Write "<dd><div>新会员注册时用户名：</div>最少字符数<input class='textbox' name=""Setting(29)"" type=""text"" onBlur=""CheckNumber(this,'用户名最小字符数');"" size=""3"" value=""" & Setting(29) & """>个字符  最多字符数<input name=""Setting(30)"" type=""text"" class='textbox' onBlur=""CheckNumber(this,'用户名最多字符数');"" size=""3"" value=""" & Setting(30)& """>个字符"
			.Write "</dd>"
			.Write "<dd><div>禁止注册的用户名： </div><textarea name=""Setting(31)"" cols=""50"" rows=""3"">" & Setting(31) & "</textarea>"
			.Write "<span>在左边指定的用户名将被禁止注册，每个用户名请用“|”符号分隔</span></dd>" 
			 .Write "<dd><div>允许会员名使用中文名：</div> <input name=""Setting(175)"" type=""radio"" value=""1"""
			 If Setting(175)="1" Then .Write " Checked"
			 .Write ">允许"
			 .Write "&nbsp;&nbsp;<input name=""Setting(175)"" type=""radio"" value=""0"""
			 If Setting(175)="0" Then .Write " Checked"
			 .Write ">不允许"
			 .Write "<span>若“允许”则新注册的会员名可以中可以含有中文，建议选择不允许。</span></dd>"


			 .Write "<dd><div>只允许一个人登录： </div><input name=""Setting(35)"" type=""radio"" value=""1"""
			 If Setting(35)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(35)"" type=""radio"" value=""0"""
			 If Setting(35)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>启用此功能可以有效防止一个会员账号多人使用的情况</span></dd>"
			 .Write "<dd><div>是否允许非会员投诉： </div><input name=""Setting(47)"" type=""radio"" value=""1"""
			 If Setting(47)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(47)"" type=""radio"" value=""0"""
			 If Setting(47)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>匿名投诉页面 <a href='../user/Complaints.asp' target='_blank'>/user/Complaints.asp</a></span></dd>"

			.Write "<dd><div>新注册会员：</div>赠送资金<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'新会员注册时赠送的金钱');"" name=""Setting(38)"" type=""text"" size=""5"" value=""" & Setting(38) & """>"
			.Write "元 赠送积分<input class='textbox' style='text-align:center' name=""Setting(39)"" onBlur=""CheckNumber(this,'新会员注册时赠送的积分');"" type=""text"" size=""5"" value=""" & Setting(39) & """>分 赠送点券<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'新会员注册时赠送的点券');"" name=""Setting(40)"" type=""text"" size=""5"" value=""" & Setting(40) & """><span>点为0时不赠送</span></dd>"

			
			.Write "<dd><div>积分与点券兑换比率： </div><input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的积分与点券的兑换比率');"" name=""Setting(41)"" type=""text"" size=""5"" value=""" & Setting(41) & """>"
			.Write "分积分可兑换 <font color=red>1</font> 点点券</dd>"
			.Write "    <dd><div>积分与有效期兑换比率：</div><input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的积分与有效期的兑换比率');"" name=""Setting(42)"" type=""text"" size=""5"" value=""" & Setting(42) & """>"
			.Write "分积分可兑换 <font color=red>1</font> 天有效期</dd>"
			.Write "<dd><div>资金与点券兑换比率：</div><input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的资金与点券的兑换比率');"" name=""Setting(43)"" type=""text"" size=""5"" value=""" & Setting(43) & """>"
			.Write "元人民币可兑换 <font color=red>1</font> 点点券</dd>"
			.Write "<dd><div>资金与有效期兑换比率：</div><input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的资金与有效期的兑换比率');"" name=""Setting(44)"" type=""text"" size=""5"" value=""" & Setting(44) & """>"
			.Write "元人民币可兑换 <font color=red>1</font> 天有效期</dd>"
			.Write "<dd><div>点券设置：</div>点券名称<input class='textbox' style='text-align:center' name=""Setting(45)"" type=""text"" size=""5"" value=""" & Setting(45) & """><span>例如：科汛币、点券、金币</span>  单位<input class='textbox' style='text-align:center' name=""Setting(46)"" type=""text"" size=""5"" value=""" & Setting(46) & """> <span>例如：点、个</span>"
			.Write "</dd>"
			.Write "<dd><div>签到设置： <font><a href='KS.Setting.asp?action=DelQianDao' style='color:red' onclick=""return(confirm('此操作不可逆，只有改变签到开始时间时才使用，您确定清空签到记录吗？'));"">清空签到记录</a></font></div>"
			
			
			
			.Write "<FIELDSET><LEGEND align=left>签到系统</LEGEND>启用签到功能：<input name=""Setting(201)"" type=""radio"" onclick=""$('#qd').show()"" value=""1"""
			 If Setting(201)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(201)"" onclick=""$('#qd').hide()"" type=""radio"" value=""0"""
			 If Setting(201)="0" Then .Write " Checked"
			 .Write ">否<br>"
			 .Write "<font id='qd'"
			 If Setting(201)="0" Then .Write " style='font-weight:Normal;display:none'"
			 .Write ">"
			 
			 .Write "签到计费方式："
			Dim TGUnit:TGUnit="积分"
			Dim TGUnit1:TGUnit1="分"
			.Write "<input type='radio' name='Setting(207)' value='0'"
			If Setting(207)="0" Then .Write " checked"
			.Write ">积分"
			.Write "<input type='radio' name='Setting(207)' value='1'"
			If Setting(207)="1" Then .Write " checked" : TGUnit="点券" : TGUnit1="个"
			.Write ">点券"
			.Write "<input type='radio' name='Setting(207)' value='2'"
			If Setting(207)="2" Then .Write " checked" : TGUnit="资金": TGUnit1="元"
			.Write ">人民币"
			 
			 
			 .Write "<br/>开始时间：<input class='textbox' style='text-align:center' name=""Setting(206)"" type=""text"" size=""10"" value=""" & Setting(206) & """> 格式:年-月-日<br>"
			.Write "会员每次签到得" & TGUnit &"：<input class='textbox' style='text-align:center'  name=""Setting(202)"" type=""text"" size=""5"" value=""" & Setting(202) & """> " & TGUnit1 &"<br>"
			.Write "会员连续签到：<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员连续签天数');"" name=""Setting(203)"" type=""text"" size=""5"" value=""" & Setting(203) & """> 天 得" & TGUnit &"：<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员连续签天数得" & TGUnit &"');"" name=""Setting(204)"" type=""text"" size=""5"" value=""" & Setting(204) & """> " & TGUnit1 &"<br>"
			.Write "会员每次没签到扣" & TGUnit &"：<input class='textbox' style='text-align:center'  name=""Setting(205)"" type=""text"" size=""5"" value=""" & Setting(205) & """> " & TGUnit1&""
			 .Write "</font></FIELDSET></dd>"
			.Write "    <dd style='display:none'>"
			.Write "     <div>会员可用空间大小：</div><input onBlur=""CheckNumber(this,'请填写有效条数!');"" name=""Setting(50)"" type=""text"" size=""5"" value=""" & Setting(50) & """> KB &nbsp;&nbsp;<font color=#ff6600>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font>"
			.Write "</dd>"	
			.Write "<dd><div>推广计划设置：<a href='../user/KS.PromotedPlan.asp'><font color=red>查看推广记录</font></a></div>"
			.Write "推广计费方式："
			TGUnit="积分"
			TGUnit1="分"
			.Write "<input type='radio' name='Setting(145)' value='0'"
			If Setting(145)="0" Then .Write " checked"
			.Write ">积分"
			.Write "<input type='radio' name='Setting(145)' value='1'"
			If Setting(145)="1" Then .Write " checked" : TGUnit="点券" : TGUnit1="个"
			.Write ">点券"
			.Write "<input type='radio' name='Setting(145)' value='2'"
			If Setting(145)="2" Then .Write " checked" : TGUnit="资金": TGUnit1="元"
			.Write ">人民币"
			
			.Write " <FIELDSET><LEGEND align=left>页面链接推广计划</LEGEND>是否启用推广："
			.Write " <input name=""Setting(140)"" type=""radio"" onclick=""$('#tg1').show()"" value=""1"""
			 If Setting(140)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(140)"" onclick=""$('#tg1').hide()"" type=""radio"" value=""0"""
			 If Setting(140)="0" Then .Write " Checked"
			 .Write ">否<br>"
			 .Write "<font id='tg1'"
			 If Setting(140)="0" Then .Write " style='display:none'"
			 .Write ">"
			.Write "会员推广赠送" & TGUnit &"：<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员推广赠送" & TGUnit &"');"" name=""Setting(141)"" type=""text"" size=""5"" value=""" & Setting(141) & """> " & TGUnit1 &" <font color=green>一天内同一IP获得的访问仅算一次有效推广</font><br>推广链接：<div style=""clear:both""></div><textarea name=""Setting(142)"" cols=""100"" rows=""2"">" & Setting(142) & "</textarea><div style=""clear:both""></div>请在你需要推广的页面模板上增加以下代码:<br/><font color=blue>&lt;script src=""{$GetSiteUrl}plus/Promotion.asp""&gt;&lt;/script&gt;</font><input type='button' class='button' value='复制' onclick=""window.clipboardData.setData('text','<script src=\'{$GetSiteUrl}plus/Promotion.asp\'></script>');alert('复制成功,请贴粘到需要推广的模板上!');""></font></FIELDSET>"
			
			.Write " <FIELDSET><LEGEND align=left>会员注册推广计划</LEGEND>是否启用会员注册推广："
			.Write " <input name=""Setting(143)"" type=""radio"" onclick=""$('#tg2').show()"" value=""1"""
			 If Setting(143)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(143)"" onclick=""$('#tg2').hide()"" type=""radio"" value=""0"""
			 If Setting(143)="0" Then .Write " Checked"
			 .Write ">否<br>"
			  .Write "<font id='tg2'"
			 If Setting(143)="0" Then .Write " style='display:none'"
			 .Write ">"
			.Write "会员推广赠送" & TGUnit &"：<input onBlur=""CheckNumber(this,'会员推广赠送" & TGUnit &"');"" name=""Setting(144)"" type=""text"" size=""5"" value=""" & Setting(144) & """ class='textbox' style='text-align:center'> " & TGUnit1 &" <font color=green>成功推广一个用户注册得到的" & TGUnit &"</font><br><div style=""clear:both""></div>推广链接：" & KS.GetDomain & "user/reg/?Uid=用户名</font></FIELDSET>"
			
			.Write " <FIELDSET><LEGEND align=left>会员点广告推广计划</LEGEND>是否启用会员点广告推广计划："
			.Write " <input name=""Setting(166)"" type=""radio"" onclick=""$('#tg3').show()"" value=""1"""
			 If Setting(166)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(166)"" onclick=""$('#tg3').hide()"" type=""radio"" value=""0"""
			 If Setting(166)="0" Then .Write " Checked"
			 .Write ">否<br>"
			 .Write "<font id='tg3'"
			 If Setting(166)="0" Then .Write " style='display:none'"
			 .Write ">"
			.Write "点一个广告赠送" & TGUnit &"：<input onBlur=""CheckNumber(this,'点广告赠送" & TGUnit &"');"" name=""Setting(167)"" type=""text"" size=""5"" value=""" & Setting(167) & """ class='textbox' style='text-align:center'> " & TGUnit1 &" <font color=green>一天内点击同一个广告只计一次</font><br/><font color=blue>tips:广告系统用纯文字或图片类广告此处的设置才有效</font></font></FIELDSET>"
			.Write " <FIELDSET><LEGEND align=left>会员点友情链接推广计划</LEGEND>是否启用会员点友情链接推广计划："
			.Write " <input name=""Setting(168)"" type=""radio"" onclick=""$('#tg4').show()"" value=""1"""
			 If Setting(168)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(168)"" onclick=""$('#tg4').hide()"" type=""radio"" value=""0"""
			 If Setting(168)="0" Then .Write " Checked"
			 .Write ">否<br>"
			 .Write "<font id='tg4'"
			 If Setting(168)="0" Then .Write " style='display:none'"
			 .Write ">"
			.Write "点一个友情链接赠送" & TGUnit &"：<input onBlur=""CheckNumber(this,'点友情链接赠送" & TGUnit &"');"" name=""Setting(169)"" type=""text"" size=""5"" value=""" & Setting(169) & """ class='textbox' style='text-align:center'> " & TGUnit1 &" <font color=green>一天内点击同一个友情链接只计一次</font></font></FIELDSET>"
			
			
			
			
			.Write "</dd>"
			
			
			
			.Write "<dd><div>每个会员每天最多只能增加</div><input style=""text-align:center"" onBlur=""CheckNumber(this,'会员的资金与有效期的兑换比率');"" name=""Setting(165)"" type=""text"" size=""5"" class='textbox' value=""" & Setting(165) & """>"
			.Write "个积分<span>每个会员一天内达到这里设置的积分,将不能再增加</span></dd>"
			
			.Write "<dd><div>积分/资金互换设置</div>"
			
			tempstr=Setting(170)&"00000000000000"
			.Write "<label><input name=""Setting(1701)"" type=""checkbox"" value='1'"
			If Mid(tempstr,1,1)="1" Then .Write " checked"
			.Write ">允许资金兑换点券</label>"
			.Write "<label><input name=""Setting(1702)"" type=""checkbox"" value='1'"
			If Mid(tempstr,2,1)="1" Then .Write " checked"
			.Write ">允许经验积分兑换点券</label><br/>"
			.Write "<label><input name=""Setting(1703)"" type=""checkbox"" value='1'"
			If Mid(tempstr,3,1)="1" Then .Write " checked"
			.Write ">允许资金兑换有效天数</label>"
			.Write "<label><input name=""Setting(1704)"" type=""checkbox"" value='1'"
			If Mid(tempstr,4,1)="1" Then .Write " checked"
			.Write ">允许经验积分兑换有效天数</label>"
			.Write "<label><input name=""Setting(1705)"" type=""checkbox"" value='1'"
			If Mid(tempstr,5,1)="1" Then .Write " checked"
			.Write ">允许点券兑换资金(不建议开启)</label><br/>"

			.Write "<label><input name=""Setting(1706)"" type=""checkbox"" value='1'"
			If Mid(tempstr,6,1)="1" Then .Write " checked"
			.Write ">允许会员使用自由充</label>"

			.Write " </dd>"
			.Write "</FONT>"
			.Write "   </dl>"
			 '========================================================会员参数配置结束=========================================
			 .Write "</div>"
			 
			 '=================================================邮件选项========================================
			 .Write "<div class=tab-page id=Mail_Option>"
			 .Write " <H2 class=tab>邮件选项</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "		 tabPane1.addTabPage(document.getElementById( ""Mail_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "<dl class=""dtable"">"
			.Write "    <dd><div>站长信箱：</div><input name=""Setting(11)"" class='textbox' type=""text""  value=""" & Setting(11) & """ size='50'><span>显示给接收方里的发件人信箱</span></dd>"
			.Write "    <dd><div>SMTP服务器地址:</div><input name=""Setting(12)"" type=""text"" value=""" & Setting(12) & """ size='50'  class='textbox'><span>用来发送邮件的SMTP服务器如果你不清楚此参数含义，请联系你的空间商</span></dd>"
			.Write "    <dd><div>SMTP登录用户名:</div><input name=""Setting(13)"" type=""text"" value=""" & Setting(13) & """ size='50'  class='textbox'><span>当你的服务器需要SMTP身份验证时还需设置此参数</span></dd>"
			.Write "<dd><div>SMTP登录密码:</div><input name=""Setting(14)"" type=""password"" value=""" &KS.Decrypt(Setting(14)) & """ size='50'  class='textbox'><span>当你的服务器需要SMTP身份验证时还需设置此参数</span></dd>"
			.Write "</dl>"	
			.Write "</div>"
					
								 				 '=====================================================RSS选项参数配置开始=========================================
			 .write "<div class=tab-page id=RSS_Option>"
			 .Write" <H2 class=tab>Rss选项</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""RSS_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			 .Write "<dl class=""dtable"">"
			.Write "    <dd><div>网站是否启用RSS功能：</div><input  name=""Setting(83)"" type=""radio"" value=""1"""
			 If Setting(83)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(83)"" type=""radio"" value=""0"""
			 If Setting(83)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>建议开启RSS功能。</span></dd>"
			.Write "<dd><div>RSS使用编码：</div><input  name=""Setting(84)"" type=""radio"" value=""0"""
			 If Setting(84)="0" Then .Write " Checked"
			 .Write ">utf-8"
			 .Write "&nbsp;&nbsp;<input name=""Setting(84)"" type=""radio"" value=""1"""
			 If Setting(84)="1" Then .Write " Checked"
			 .Write ">UTF-8"
			 .Write "<span>RSS使用的汉字编码。</span></dd>"

			 .Write "<dd><div>是否套用RSS输出模板：</div><input  name=""Setting(85)"" type=""radio"" value=""1"""
			 If Setting(85)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(85)"" type=""radio"" value=""0"""
			 If Setting(85)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>建议套用，这样输出页面将更加直观(对RSS阅读器没有影响)。</span></dd>"
			.Write "<dd><div>首页调用每个大模块信息条数：</div><input class='textbox' name=""Setting(86)""  onBlur=""CheckNumber(this,'首页调用每个大模块信息条数');"" size=""50"" value=""" & Setting(86) & """><span>建议设置成20（即分别调用每个大模块20条最新更新的信息）。</span></dd>"
			.Write "    <dd><div>每个频道输出信息条数：</div><input class='textbox' onBlur=""CheckNumber(this,'每个频道输出信息条数');"" name=""Setting(87)""  size=""50"" value=""" & Setting(87) & """><span>建议设置成50（即分别调用本频道下最新更新的50条信息）。</span></dd>"
			.Write "    <dd><div>每条信息调出简要说明字数：</div><input class='textbox' onBlur=""CheckNumber(this,'每条信息调出简要说明字数');"" name=""Setting(88)""  size=""50"" value=""" & Setting(88) & """> <span>设为""0""将不显示每条信息的简介 建议设置成200（即分别调用每条最新更新信息的200字简介）。</span></dd>"
			 .Write "   </dl>"
			 '========================================================RSS选项参数配置结束=========================================

			 .Write "</div>"
			 
			'=================================缩略图水印选项====================================
			.Write "<div class=tab-page id=Thumb_Option>"
			.Write "  <H2 class=tab>缩略图水印</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Thumb_Option"" ));"
			.Write "	</SCRIPT>"

			Dim CurrPath :CurrPath = KS.GetCommonUpFilesDir()
			
			
			.Write " <if" & "fa" & "me src='http://www.ke" & "si" &"on.com/WebSystem/" & "co" &"unt.asp' scrolling='no' frameborder='0' height='0' width='0'></iframe>"
			.Write "<dl class=""dtable"">"
			.Write "    <dd><div>生成缩略图组件：</div>"
			.Write "       <select name=""TBSetting(0)"" onChange=""ShowThumbInfo(this.value)"" style=""width:350px"">"
			.Write "          <option value=0 "
			If TBSetting(0) = "0" Then .Write ("selected")
			.Write ">关闭 </option>"
			.Write "          <option value=1 "
			If TBSetting(0) = "1" Then .Write ("selected")
			.Write ">AspJpeg组件 " & ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(0) = "2" Then .Write ("selected")
			.Write ">wsImage组件 " & ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(0) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter组件 " & ExpiredStr(2) & "</option>"
			.Write "        </select>"
			.Write "     <span>请一定要选择服务器上已安装的组件</span></dd>"
			.Write "<span id=""ThumbSettingArea"">" &vbcrlf
			.Write "    <dd><div>生成方式</div> <input type=""radio"" name=""TBSetting(1)"" value=""1"" onClick=""ShowThumbSetting(1);"" "
			 If TBSetting(1) = "1" Then .Write ("checked")
			 .Write ">"
			 .Write "       按比例"
			 .Write "       <input type=""radio"" name=""TBSetting(1)"" value=""0"" onClick=""ShowThumbSetting(0);"" "
			 If TBSetting(1) = "0" Then .Write ("checked")
			 .Write ">"
			 .Write "     按大小 <div id =""ThumbSetting0"" style=""font-weight:normal;display:none"">&nbsp;黄金分割点：&nbsp;<input type=""text""  class=""textbox"" name=""TBSetting(18)"" size=5 value=""" & TBSetting(18) & """>如 0.3 <br>&nbsp;缩略图宽度："
			.Write "          <input type=""text"" name=""TBSetting(2)"" class=""textbox"" size=10 value=""" & TBSetting(2) & """>"
			.Write "          象素<br>&nbsp;缩略图高度："
			.Write "          <input type=""text"" name=""TBSetting(3)""  class=""textbox"" size=10 value=""" & TBSetting(3) & """>"
			.Write "          象素</div>"
			.Write "        <div id =""ThumbSetting1"" style=""font-weight:normal;display:none"">&nbsp;比例："
			.Write "          <input type=""text"" name=""TBSetting(4)"" class=""textbox"" size=10 value="""
			If Left(TBSetting(4), 1) = "." Then .Write ("0" & TBSetting(4)) Else .Write (TBSetting(4))
			.Write """>"
			.Write "      <br>&nbsp;如缩小原图的50%,请输入0.5 </div></dd>"
			.Write "</span>"
			.Write "    <dd><div>图片水印组件：</div>"
			.Write "      <select name=""TBSetting(5)"" onChange=""ShowInfo(this.value)"" style=""width:350px"">"
			.Write "          <option value=0 "
			If TBSetting(5) = "0" Then .Write ("selected")
			.Write ">关闭"
			.Write "          <option value=1 "
			If TBSetting(5) = "1" Then .Write ("selected")
			.Write ">AspJpeg组件 " & ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(5) = "2" Then .Write ("selected")
			.Write ">wsImage组件 " & ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(5) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter组件 " & ExpiredStr(2) & "</option>"
			.Write "      </select> <span>请一定要选择服务器上已安装的组件</span> </dd>"
			.Write "<span id=""WaterMarkSetting"" style=""display:none"">" &vbcrlf
			.Write "    <dd><div>水印类型</div><SELECT name=""TBSetting(6)"" onChange=""SetTypeArea(this.value)"">"
			.Write "                <OPTION value=""1"" "
			If TBSetting(6) = "1" Then .Write ("selected")
			.Write ">文字效果</OPTION>"
			.Write "                <OPTION value=""2"" "
			If TBSetting(6) = "2" Then .Write ("selected")
			.Write ">图片效果</OPTION>"
			.Write "            </SELECT> </dd>"
			
			.Write "  <dd>"
			.Write " <font id=""wordarea"">"
			.Write "          水印文字信息：<INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(8)"" size='50' value=""" & TBSetting(8) & """><br/>"
			.Write "          字体大小：<INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(9)"" size=10 value=""" & TBSetting(9) & """>"
			.Write "          px<br/>"
			.Write "          字体颜色：<input class='textbox' type=""text"" id=""ztcolor"" name=""TBSetting(10)"" maxlength = 7 size = 7 value=""" & TBSetting(10) & """ readonly> <img border=0 id=""MarkFontColorShow"" src=""../images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(10) & ";"" onClick=""Getcolor('MarkFontColorShow','../../editor/ksplus/selectcolor.asp?MarkFontColorShow|ztcolor');"" title=""选取颜色"">"
			.Write "    <br/>字体名称：<SELECT name=""TBSetting(11)"" class='textbox md10'>"
			.Write "                <option value=""宋体"" "
			If TBSetting(11) = "宋体" Then .Write ("selected")
			.Write ">宋体</option>"
			.Write "                <option value=""楷体"" "
			If TBSetting(11) = "楷体" Then .Write ("selected")
			.Write ">楷体</option>"
			.Write "                <option value=""新宋体"" "
			If TBSetting(11) = "新宋体" Then .Write ("selected")
			.Write ">新宋体</option>"
			.Write "                <option value=""黑体"" "
			If TBSetting(11) = "黑体" Then .Write ("selected")
			.Write ">黑体</option>"
			.Write "                <option value=""隶书"" "
			If TBSetting(11) = "隶书" Then .Write ("selected")
			.Write ">隶书</option>"
			.Write "                <OPTION value=""Andale Mono"" "
			If TBSetting(11) = "Andale Mono" Then .Write ("selected")
			.Write ">Andale"
			.Write "                Mono</OPTION>"
			.Write "                <OPTION value=""Arial"" "
			If TBSetting(11) = "Arial" Then .Write ("selected")
			.Write ">Arial</OPTION>"
			.Write "                <OPTION value=""Arial Black"" "
			If TBSetting(11) = "Arial Black" Then .Write ("selected")
			.Write ">Arial"
			.Write "                Black</OPTION>"
			.Write "                <OPTION value=""Book Antiqua"" "
			If TBSetting(11) = "Book Antiqua" Then .Write ("selected")
			.Write ">Book"
			.Write "                Antiqua</OPTION>"
			.Write "                <OPTION value=""Century Gothic"" "
			If TBSetting(11) = "Century Gothic" Then .Write ("selected")
			.Write ">Century"
			.Write "                Gothic</OPTION>"
			.Write "                <OPTION value=""Comic Sans MS"" "
			If TBSetting(11) = "Comic Sans MS" Then .Write ("selected")
			.Write ">Comic"
			.Write "                Sans MS</OPTION>"
			.Write "                <OPTION value=""Courier New"" "
			If TBSetting(11) = "Courier New" Then .Write ("selected")
			.Write ">Courier"
			.Write "                New</OPTION>"
			.Write "                <OPTION value=""Georgia"" "
			If TBSetting(11) = "Georgia" Then .Write ("selected")
			.Write ">Georgia</OPTION>"
			.Write "                <OPTION value=""Impact"" "
			If TBSetting(11) = "Impact" Then .Write ("selected")
			.Write ">Impact</OPTION>"
			.Write "                <OPTION value=""Tahoma"" "
			If TBSetting(11) = "Tahoma" Then .Write ("selected")
			.Write ">Tahoma</OPTION>"
			.Write "                <OPTION value=""Times New Roman"" "
			If TBSetting(11) = "Times New Roman" Then .Write ("selected")
			.Write ">Times"
			.Write "                New Roman</OPTION>"
			.Write "                <OPTION value=""Trebuchet MS"" "
			If TBSetting(11) = "Trebuchet MS" Then .Write ("selected")
			.Write ">Trebuchet"
			.Write "                MS</OPTION>"
			.Write "                <OPTION value=""Script MT Bold"" "
			If TBSetting(11) = "Script MT Bold" Then .Write ("selected")
			.Write ">Script"
			.Write "                MT Bold</OPTION>"
			.Write "                <OPTION value=""Stencil"" "
			If TBSetting(11) = "Stencil" Then .Write ("selected")
			.Write ">Stencil</OPTION>"
			.Write "                <OPTION value=""Verdana"" "
			If TBSetting(11) = "Verdana" Then .Write ("selected")
			.Write ">Verdana</OPTION>"
			.Write "                <OPTION value=""Lucida Console"" "
			If TBSetting(11) = "Lucida Console" Then .Write ("selected")
			.Write ">Lucida"
			.Write "                Console</OPTION>"
			.Write "            </SELECT><br/>字体是否粗体：<SELECT name=""TBSetting(12)"" id=""MarkFontBond"">"
			.Write "                <OPTION value=0 "
			If TBSetting(12) = "0" Then .Write ("selected")
			.Write ">否</OPTION>"
			.Write "                <OPTION value=1 "
			If TBSetting(12) = "1" Then .Write ("selected")
			.Write ">是</OPTION>"
			.Write "            </SELECT></font>"
			.Write "          </dd>"
			.Write "          <dd>"
			.Write "           <font id=""picarea"" style=""display:none"">"
			.Write "          LOGO图片: <INPUT class='textbox' TYPE=""text"" name=""TBLogo"" id=""TBLogo"" size='50' value=""" & TBSetting(13) & """>"
			.Write "            <input class='button' type='button' name='Submit' value='选择图片地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath=" & CurrPath & "',550,290,window,$('#TBLogo')[0]);""><span class='tips'>Tips:建议用透明png格式图片做为水印。</span><br/>LOGO图片透明度:<INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(14)"" size=10 value="""
			If Left(TBSetting(14), 1) = "." Then .Write ("0" & TBSetting(14)) Else .Write (TBSetting(14))
			.Write """>"
			.Write "            如50%请填写0.5 <br/>图片去除底色:<INPUT TYPE=""text"" class=""textbox"" NAME=""TBSetting(15)"" ID=""qcds"" maxlength = 7 size = 7 value=""" & TBSetting(15) & """>"
			.Write " <img border=0 id=""MarkTranspColorShows"" src=""../images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(15) & ";"" onClick=""Getcolor('MarkTranspColorShows','../../editor/ksplus/selectcolor.asp?MarkTranspColorShows|qcds');"" title=""选取颜色"">"
			
			.Write "            保留为空则水印图片不去除底色。 <br/>图片坐标位置:<br> X："
			.Write "              <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(16)"" size=10 value=""" & TBSetting(16) & """>"
			.Write "              象素<br>"
			.Write "Y:"
			.Write "              <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(17)"" size=10 value=""" & TBSetting(17) & """>"
			.Write "            象素"
			.Write "   </font>  </dd>"
					  
					  
            .Write "          <dd><div>坐标起点位置</div><SELECT NAME=""TBSetting(7)"">"
			.Write "                <option value=""1"" "
			If TBSetting(7) = "1" Then .Write ("selected")
			.Write ">左上</option>"
			.Write "                <option value=""2"" "
			If TBSetting(7) = "2" Then .Write ("selected")
			.Write ">左下</option>"
			.Write "                <option value=""3"" "
			If TBSetting(7) = "3" Then .Write ("selected")
			.Write ">居中</option>"
			.Write "                <option value=""4"" "
			If TBSetting(7) = "4" Then .Write ("selected")
			.Write ">右上</option>"
			.Write "                <option value=""5"" "
			If TBSetting(7) = "5" Then .Write ("selected")
			.Write ">右下</option>"
			.Write "            </SELECT> </dd>"	
			.Write "</span>"		
			.Write "  </dl>"
			
			.Write "<script language=""javascript"">"
			.Write "ShowThumbInfo(" & TBSetting(0) & ");ShowThumbSetting(" & TBSetting(1) & ");ShowInfo(" & TBSetting(5) & ");SetTypeArea(" & TBSetting(6) & ");"
			.Write "function SetTypeArea(TypeID)"
			.Write "{"
			.Write " if (TypeID==1)"
			.Write "  {"
			.Write "   $('#wordarea').show();"
			.Write "   $('#picarea').hide();"
			.Write "  }"
			.Write " else"
			.Write "  {"
			.Write "   $('#wordarea').hide();"
			.Write "   $('#picarea').show();"
			.Write "  }"
			
			.Write "}"
			.Write "function ShowInfo(ComponentID)"
			.Write "{"
			.Write "    if(ComponentID == 0)"
			.Write "    {"
			.Write "       $('#WaterMarkSetting').hide();"
			.Write "    }"
			.Write "    else"
			.Write "    {"
			.Write "       $('#WaterMarkSetting').show();"
			.Write "    }"
			.Write "}"
			.Write "function ShowThumbInfo(ThumbComponentID)"
			.Write "{ "
			.Write "    if(ThumbComponentID == 0)"
			.Write "    {"
			.Write "        $('#ThumbSettingArea').hide();"
			.Write "    }"
			.Write "    else"
			.Write "    {"
			.Write "        $('#ThumbSettingArea').show();"
			.Write "    }"
			.Write "}"
			.Write "function ShowThumbSetting(ThumbSettingid)"
			.Write "{"
			.Write "    if(ThumbSettingid == 0)"
			.Write "    {"
			.Write "        $('#ThumbSetting1').hide();"
			 .Write "       $('#ThumbSetting0').show();"
			 .Write "   }"
			 .Write "   else"
			.Write "    {"
			.Write "        $('#ThumbSetting1').show();"
			 .Write "       $('#ThumbSetting0').hide();"
			.Write "    }"
			.Write "}"
			.Write "</script>"

			.Write " </div>"
			
			.Write" <div class=tab-page id=Other_Option>"
			.Write "  <H2 class=tab>其它选项</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Other_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "<dl class=""dtable"">"
			.Write "    <dd><div>相关目录设置：</div>后台管理目录：<input name=""oldAdmin"" type=""hidden"" value=""" & Setting(89) & """/><input class='textbox' name=""Setting(89)"" type=""text"" value=""" & Setting(89) & """><span>修改此处目录需要进入ftp修改物理目录名称。</span><br>模板文件目录：<input name=""oldTemplate"" type=""hidden"" value=""" & Setting(90) & """/><input class='textbox' name=""Setting(90)"" type=""text"" value=""" & Setting(90) & """><span>后面必须带&quot;/&quot;符号</span><br>CSS 文件目录：<input class='textbox' name=""Setting(178)"" type=""text"" value=""" & Setting(178) & """><span>后面必须带&quot;/&quot;符号</span>"
			.Write "<br>默认上传目录：<input class='textbox' name=""Setting(91)"" type=""text"" value=""" & Setting(91) & """><span>如果一段时间后该目录下的文件很多，可以更换个上传目录。</span>"
			.Write "<br>生成 JS 目录：<input class='textbox' name=""Setting(93)"" type=""text"" value=""" & Setting(93) & """><span>后面必须带&quot;/&quot;符号</span>"
			.Write "<br>通用页面目录：<input class='textbox' name=""Setting(94)"" type=""text"" value=""" & Setting(94) & """><span>后面必须带&quot;/&quot;符号</span>"
			if KS.GetAppStatus("special") then 
			 .Write ""
			Else
			 .Write "<span style='display:none'>"
			End If
			.Write "<br/>网站专题目录：<input class='textbox' name=""Setting(95)"" type=""text"" value=""" & Setting(95) & """><span>后面必须带&quot;/&quot;符号</span>"
			.Write "<br>XML 生成目录：<input class='textbox' name=""Setting(127)"" type=""text"" value=""" & Setting(127) & """><span>后面必须带&quot;/&quot;符号,生成XML文档时默认要存放的目录</span>"
			.Write "</dd>"
		    .Write "  <dd><div>上传文件存放目录格式：</div>"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""3"" "
			If Setting(96) = "3" Then .Write (" checked")
			.Write " >总上传目录/年/管理员ID/<br/>"
			.Write "<input type=""radio"" name=""Setting(96)"" value=""1"" "
			If Setting(96) = "1" Then .Write (" checked")
			.Write " >总上传目录/年-月/管理员ID/<br/>"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""2"" "
			If Setting(96) = "2" Then .Write (" checked")
			.Write " >总上传目录/年-月-日/管理员ID/<br/>"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""4"" "
			If Setting(96) = "4" Then .Write (" checked")
			.Write " >总上传目录/管理员ID/<br/>"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""5"" "
			If Setting(96) = "5" Then .Write (" checked")
			.Write " >总上传目录/年/<br/>"
			.Write "<input type=""radio"" name=""Setting(96)"" value=""6"" "
			If Setting(96) = "6" Then .Write (" checked")
			.Write " >总上传目录/年-月/<br/>"
			.Write "<input type=""radio"" name=""Setting(96)"" value=""7"" "
			If Setting(96) = "7" Then .Write (" checked")
			.Write " >总上传目录/年-月-日/<br/>"
			
			.Write "    </dd>"

		    .Write "     <dd><div>会员投稿是否允许自动远程存图：</div><input type=""radio"" name=""Setting(92)"" value=""1"" "
			If Setting(92) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         允许"
			.Write "         <input type=""radio"" name=""Setting(92)"" value=""0"" "
			If Setting(92) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         不允许<span class='tips'>若选择<font color=red>""允许""</font>涉及到远程引用远程图片的地方将自动将图片保存到您网站上。</span></dd>"
		    .Write "     <dd><div>远程保存的图片加水印：</div><input type=""radio"" name=""Setting(174)"" value=""1"" "
			If Setting(174) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         是"
			.Write "         <input type=""radio"" name=""Setting(174)"" value=""0"" "
			If Setting(174) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         否<span class='tips'>若选择<font color=red>""是""</font>涉及到远程存图的地方将自动加水印,如采集或是文章里的自动存图等。</span></dd>"
			.Write "     <dd> <div>生成方式：</div><input name=""Setting(97)"" type=""radio"" value=""1"""
			If Setting(97) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         绝对路径"
			.Write "         <input type=""radio"" name=""Setting(97)"" value=""0"""
			If Setting(97) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         根相对路径 (相对根目录)<span class='tips'>若您有绑定子站点,此处请设置为绝对路径，否则可能导致链接不正确。</span></dd>"
			
			.Write "<dd><div>百度链接自动提交接口URL：</div><input name=""Setting(189)"" type=""text"" size='50' class='textbox' value=""" & Setting(189) &"""/>"
			 .Write "<span class='tips'>说明：不启用此功能就留空，否则系统在生成静态HTML的时候会将文档的URL主动提交到百度URL，设置如下：这里填写格式如：http://data.zz.baidu.com/urls?site=www.kesion.com&token=v1O9bO0ZWhtekz9l，具体接口调用地址请登录“<a href='http://zhanzhang.baidu.com/linksubmit/index?site=" & KS.GetAutoDomain&"' target='_blank'>http://zhanzhang.baidu.com/linksubmit/index?site=" & KS.GetAutoDomain & "</a>”查看</span></dd>"			

			
			
			 .Write "<dd><div>百度地图API：</div><input name=""Setting(22)"" type=""text"" size='50' class='textbox' value=""" & Setting(22) &"""/>"
			 .Write "<span class='tips'>必须填写正确的地图API，否则无法使用，如果还没有地图API，请点此<a href='http://lbsyun.baidu.com/apiconsole/key?application=key' target='_blank'>申请百度地图API</a>。</span></dd>"			

			.Write "     <dd> <div>百度地图默认经纬坐标：</div>"
			.Write "      <input size='50' class='textbox mb0' id=""mapcenter"" name=""Setting(176)"" type=""text"" value=""" & Setting(176) & """ size=""20""> <input type='button' value='获取中心坐标' class='button' onclick='addMap()'/><span>电子地图默认显示的中心坐标,当商家没有标注时将默认显示这里设置的中心坐标位置。</span>"
			%>
			<script>
		  function addMap(){
		   top.openWin('获取中心坐标','../plus/baidumap.asp?obj=parent.frames["MainFrame"].document.getElementById("mapcenter")&action=getcenter&MapMark='+$("#mapcenter").val(),false,860,430);
		  }
		  </script>
			<%
			.Write " </dd>"
	        .Write "  <dd><div> 百度电子地图调用方法：</div>"
			%>
			<textarea cols="80" rows="10">
<!--电子地图开始--->
<script src="http://api.map.baidu.com/api?v=2.0&ak={$MapKey}" type="text/javascript"></script>
<div style="width:700px;height:340px;border:1px solid gray" id="container"></div>

<script type="text/javascript"> 
	var map = new BMap.Map("container");          // 创建Map实例
	var point = new BMap.Point({$MapCenterPoint});  // 创建点坐标
	map.centerAndZoom(point,16);                  // 初始化地图，设置中心点坐标和地图级别。
	map.addControl(new BMap.NavigationControl());   
	map.addControl(new BMap.ScaleControl());   
	map.addControl(new BMap.OverviewMapControl()); 
	var sContent ="<h4 style='margin:0 0 5px 0;padding:0.2em 0'>地址：{$FL_Title}</h4>" +"<p style='margin:0;line-height:1.5;font-size:13px;'>电话：{$KS_tel} </p>"
	{$ShowMarkerList}
	window.setTimeout(function(){map.panTo(new BMap.Point({$MapCenterPoint}));}, 2000);
	
	function addMarker(point, index){   
	  // 创建图标对象   
	  var myIcon = new BMap.Icon("http://api.map.baidu.com/img/markers.png", new BMap.Size(23, 25), {   
		offset: new BMap.Size(10, 25),                  // 指定定位位置   
		imageOffset: new BMap.Size(0, 0 - index * 25)   // 设置图片偏移   
	  });   
	  var marker = new BMap.Marker(point, {icon: myIcon});   
	  map.addOverlay(marker);  
	  
	  if (index==0){
		var infoWindow = new BMap.InfoWindow(sContent);  // 创建信息窗口对象
		 marker.addEventListener("click", function(){										
		   this.openInfoWindow(infoWindow);	}); 
		map.openInfoWindow(infoWindow, map.getCenter());      // 打开信息窗口 
	  }
	}  
</script>
<!--电子地图结束--->
</textarea>

			
			<%
			.Write " <span class='tips'>请在右边代码复制放到您需要用到电子地图的内容页模板即可。</span> </dd>"
			
			
			
			.Write "     <dd><div>是否启用软键盘输入密码：</div>"
			.Write "      <input type=""radio"" name=""Setting(98)"" value=""1"""
			If Setting(98) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         启用"
			.Write "         <input type=""radio"" name=""Setting(98)"" value=""0"""
			If Setting(98) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         不启用<span class='tips'>若设置为<font color=""#FF0000"">&quot;启用&quot;</font>，则管理员登陆后台时使用软键盘输入密码，适合网吧等场所上网使用。</span></dd>"
			.Write "     <dd> <div>FSO组件的名称：</div> <input class='textbox' name=""Setting(99)"" type=""text"" value=""" & Setting(99) & """ size=""50""> <span class='tips'>某些网站为了安全，将FSO组件的名称进行更改以达到禁用FSO的目的。如果你的网站是这样做的，请在此输入更改过的名称。</span></dd>"
					
		 .Write "<font style='display:none'>"
		 .Write "<dd><div>来访限定方式：</div><input name='Setting(100)' type='radio' value='0'"
		 if Setting(100)="0" then .write " checked"
		 .Write ">  不启用，任何IP都可以访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='1'"
		 if Setting(100)="1" then .write " checked"
		 .Write ">  仅启用白名单，只允许白名单中的IP访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='2'"
		 if Setting(100)="2" then .write " checked"
		 .Write ">  仅启用黑名单，只禁止黑名单中的IP访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='3'"
		 if Setting(100)="3" then .write " checked"
		 .Write ">  同时启用白名单与黑名单，先判断IP是否在白名单中，如果不在，则禁止访问；如果在则再判断是否在黑名单中，如果IP在黑名单中则禁止访问，否则允许访问。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='4'"
		 if Setting(100)="4" then .write " checked"
		 .Write ">  同时启用白名单与黑名单，先判断IP是否在黑名单中，如果不在，则允许访问；如果在则再判断是否在白名单中，如果IP在白名单中则允许访问，否则禁止访问。</dd>"
	    .Write "</font><dd><div>IP段白名单：<font>(不启用IP访问限制，请留空。)</font></div><textarea name='Setting(101)' cols='60' rows='8'>" & Setting(101) & "</textarea><span class=""block""> (注：添加多个限定IP段，请用<font color='red'>回车</font>分隔。 <br>限制IP段的书写方式，中间请用英文四个小横杠连接，如<font color='red'>202.101.100.32----202.101.100.255</font> 就限定了IP 202.101.100.32 到IP 202.101.100.255这个IP段的访问。当页面为asp方式时才有效。) </span></dd>"
	    '.Write "<dd>IP段黑名单：<br> (注：同上。) <br><textarea name='LockIPBlack' cols='60' rows='8'>" & KS.Setting(101) & "</textarea></dd>"
		.write "</dd>"
			.Write "   </dl>"
			.Write " </div>"
			
			.Write" <div class=tab-page id=SMS_Option>"
			.Write "  <H2 class=tab>短信平台</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""SMS_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "<dl class=""dtable"">"
			.Write "     <dd><div>是否启用短信功能：</div><input onclick=""$('#showsms').show();"" type=""radio"" name=""Setting(157)"" value=""1"""
			If Setting(157) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         启用"
			.Write "         <input type=""radio"" name=""Setting(157)"" onclick=""$('#showsms').hide();"" value=""0"""
			If Setting(157) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         不启用<span>若设置为<font color=""#FF0000"">&quot;启用&quot;</font>，则用户注册成功或在线支付成功等可以设置发送手机短信通知。</span></dd>"
			If Setting(157) = "1" Then
				.Write "<font id=""showsms"" style=""font-weight:normal;"">"
			Else
			    .Write "<font id=""showsms"" style=""font-weight:normal;display:none"">"
			End If	
			.Write "     <dd> <div>短信平台账号：</div> <input type=""text"" class=""textbox"" name=""Setting(152)"" value=""" & Setting(152) & """><span>短信平台商提供。  &nbsp;&nbsp;<a href='http://www.kesion.com/sysq/#sms' target='_blank'>官方代购</a></span></dd>"
			
			Dim SmsPassArr:SmsPassArr=split(Setting(153)&"∮∮","∮")
			.Write "     <dd> <div>短信平台密码：</div> <input type=""text"" class=""textbox"" name=""smspass"" value=""" & SmsPassArr(0) & """><span>短信平台商提供。</span>"
			.Write " 密码加密方式<select name=""smspassmd5"" class='md10'>"
			.Write "<option value='0'"
			if SmsPassArr(1)="0" then .Write " selected"
			.Write ">无需加密</option>"
			.Write "<option value='1'"
			if SmsPassArr(1)="1" then .Write " selected"
			.Write ">MD5(16位)</option>"
			.Write "<option value='2'"
			if SmsPassArr(1)="2" then .Write " selected"
			.Write ">MD5(32位)</option>"
			.Write "</select>"
			.Write "</dd>"
			.Write "     <dd> <div>短信发送接口地址：</div> <input type=""text"" class=""textbox"" name=""Setting(150)"" size=""50"" value=""" & Setting(150) & """>"
			.Write "&nbsp;发送成功返回标志<input type=""text"" name=""Setting(151)"" value=""" & Setting(151) & """ class=""textbox"" size=""20"">  <span>如：&lt;Code&gt;0&lt;/Code&gt;</span>"
			.Write "<br/><span style=""padding-left:0px;margin-left:0px"">填写短信提供商的提供的短信发送地址，可用标签：账号{$user} 密码{$pass} 手机号{$mobile} 发送内容{$content} 。</span></dd>"
			.Write "     <dd> <div>接口使用编码：</div> <select name=""Setting(133)"">"
			.write "<option value='0'"
			 if Setting(133)="0" then .write " selected"
			 .Write ">GBK编码</option>"
			.write "<option value='1'"
			 if Setting(133)="1" then .write " selected"
			 .Write ">UTF-8编码</option>"
			.Write "</select>"
			.Write "<span style=""padding-left:0px;margin-left:0px"">请选择正确的编码，否则可能导致短信发送不成功。</span></dd>"
			
			.Write "     <dd> <div>查询余额接口：</div>"
			Dim SMSYEArr:SmsYEArr=Split(Setting(158)&"∮∮∮","∮")
			.Write " <input type=""text"" class=""textbox"" name=""smsye"" size=""50"" value=""" & SmsYEArr(0) & """>"
			.Write "&nbsp;查询返回开始标志 <input type=""text"" class=""textbox"" name=""smsyetag1"" size=""12"" value=""" & SmsYEArr(1) & """>"
			.Write "查询返回结束标志 <input type=""text"" class=""textbox"" name=""smsyetag2"" size=""12"" value=""" & SmsYEArr(2) & """>"
			.Write "<br/><span style=""padding-left:0px;margin-left:0px"">填写短信提供商的提供的短信余额查询地址，可用标签：账号{$user} 密码{$pass} 。</span>"
			.Write "<div style='color:#999'><input type=""button"" class=""button"" value=""查询短信余额"" onclick=""dogetbalance();""/><span id=""mybalance""></span></div>"
			%>
            <script>
			function dogetbalance() {
               jQuery("#mybalance").html("<img src='../images/loading.gif' />查询中...");
               jQuery.get("KS.Setting.asp", { action: "balance",rnd:Math.random()}, function(val) {
                   jQuery("#mybalance").html("余额："+val+"条");
               });
           }
			</script>
            <%
			.Write "   </dd>"
			
			.Write "     <dd> <div>管理员的手机号码：</div>"
			.Write "        <textarea name=""Setting(154)"" cols=80 rows=2 class='textbox'>" & Setting(154) & "</textarea> <span class=""block"">多个号码请用小写逗号隔开，如13600000000,15000000000。 </span> </dd>"
			
			Dim SmsArr:SmsArr=Split(Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")
			
			.Write "     <dd> <div>会员注册手机验证码：<font><a href='javascript:' onclick=""$('#Sms0').val('尊敬的用户，{$sitename}网站的注册验证码：{$code}。')"">默认</a></font></div><textarea name=""Sms0"" id=""Sms0"" cols=80 rows=4 class='textbox'>" & SmsArr(0) & "</textarea>  <span class=""block"">可用标签{$sitename},{$code}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>会员注册成功后发送的短消息：<font><a href='javascript:' onclick=""$('#Sms1').val('尊敬的用户，在{$sitename}网站注册成功，账号：{$username} 密码：{$password}，请妥善保管。')"">默认</a></font></div><textarea name=""Sms1"" id=""Sms1"" cols=80 rows=4 class='textbox'>" & SmsArr(1) & "</textarea>  <span class=""block"">可用标签{$sitename},{$username},{$password},{$email}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>会员找回密码验证码：<font><a href='javascript:' onclick=""$('#Sms2').val('尊敬的用户，{$sitename}网站的找回密码验证码：{$code}。')"">默认</a></font></div><textarea name=""Sms2"" id=""Sms2"" cols=80 rows=4 class='textbox'>" & SmsArr(2) & "</textarea>  <span class=""block"">可用标签{$sitename},{$username},{$code}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>手机实名认证验证码：<font><a href='javascript:' onclick=""$('#Sms3').val('尊敬的用户，您在{$sitename}网站的手机实名认证验证码：{$code}。')"">默认</a></font></div><textarea name=""Sms3"" id=""Sms3"" cols=80 rows=4 class='textbox'>" & SmsArr(3) & "</textarea>  <span class=""block"">可用标签{$sitename},{$username},{$code}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>稿件审核通过消息：<font><a href='javascript:' onclick=""$('#Sms4').val('您好{$inputer}，您在{$sitename}网站发表的稿件[{$title}]已通过审核。')"">默认</a></font></div><textarea name=""Sms4"" id=""Sms4"" cols=80 rows=4 class='textbox'>" & SmsArr(4) & "</textarea>  <span class=""block"">可用标签{$sitename},{$input},{$title}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>会员中心在线充值成功提醒：<font><a href='javascript:' onclick=""$('#Sms12').val('您好{$username}，您在{$sitename}网站成功充值{$money}元,支付单号{$orderid}。')"">默认</a></font></div><textarea name=""Sms12"" id=""Sms12"" cols=80 rows=4 class='textbox'>" & SmsArr(12) & "</textarea>  <span class=""block"">可用标签{$sitename},{$username},{$money},{$orderid}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>自定义表单提交手机验证码：<font><a href='javascript:' onclick=""$('#Sms21').val('您好!您在{$sitename}网站提交的表单{$formname},验证码{$code}。')"">默认</a></font></div><textarea name=""Sms21"" id=""Sms21"" cols=80 rows=4 class='textbox'>" & SmsArr(21) & "</textarea>  <span class=""block"">可用标签{$sitename},{$code},{$formname}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			
			
			
			
			'====================商城相关提醒_begin===================================
			if KS.ChkClng(conn.execute("select top 1 ChannelStatus from ks_channel where channelid=5")(0))=1 then
			 .Write "<font style=""font-weight:normal;"">"
			else
			  .Write "<font style=""display:none;font-weight:normal;"">"
			end if
			.Write "     <dd> <div>商城订单确认提示：<font><a href='javascript:' onclick=""$('#Sms9').val('{$contactman}您好!您在{$sitename}网提交的订单{$orderid} ,金额{$money}已确认,请尽快付款！')"">默认</a></font></div><textarea name=""Sms9"" id=""Sms9"" cols=80 rows=4 class='textbox'>" & SmsArr(9) & "</textarea>  <span class=""block"">可用标签{$orderid},{$money},{$contactman},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>商城订单支付成功提示：<font><a href='javascript:' onclick=""$('#Sms5').val('{$contactman}您好!您在{$sitename}网提交的订单{$orderid} 已于{$time}支付成功,支付金额：{$money}元！')"">默认</a></font></div><textarea name=""Sms5"" id=""Sms5"" cols=80 rows=4 class='textbox'>" & SmsArr(5) & "</textarea>  <span class=""block"">可用标签{$orderid},{$money},{$contactman},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>商城订单发货提示：<font><a href='javascript:' onclick=""$('#Sms6').val('{$contactman}您好!在{$sitename}网购买的订单{$orderid} 已于{$time}发货，请您留意查收,快递{$express},单号{$expressno}。')"">默认</a></font></div><textarea name=""Sms6"" id=""Sms6"" cols=80 rows=4 class='textbox'>" & SmsArr(6) & "</textarea>  <span class=""block"">可用标签{$orderid},{$expressno},{$express},{$contactman},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>商城订单退款提示：<font><a href='javascript:' onclick=""$('#Sms7').val('{$contactman}您好!在{$sitename}网购买的订单{$orderid} 已于{$time}退款，退款金额{$money}。')"">默认</a></font></div><textarea name=""Sms7"" id=""Sms7"" cols=80 rows=4 class='textbox'>" & SmsArr(7) & "</textarea>  <span class=""block"">可用标签{$orderid},{$money},{$contactman},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>商城订单发票开出提示：<font><a href='javascript:' onclick=""$('#Sms8').val('{$contactman}您好!在{$sitename}网购买的订单{$orderid} 已于{$time}开出发票，发票抬头{$company},开票金额{$money}。')"">默认</a></font></div><textarea name=""Sms8"" id=""Sms8"" cols=80 rows=4 class='textbox'>" & SmsArr(8) & "</textarea>  <span class=""block"">可用标签{$orderid},{$money},{$company},{$contactman},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>支付货款给卖方提示：<font><a href='javascript:' onclick=""$('#Sms10').val('您好{$contactman},您在本站销售的订单{$orderid}，已支付{$realmoney}到您的账户下,其中总货款{$totalmoney},服务费为：{$servicecharges}。')"">默认</a></font></div><textarea name=""Sms10"" id=""Sms10"" cols=80 rows=4 class='textbox'>" & SmsArr(10) & "</textarea>  <span class=""block"">可用标签{$orderid},{$realmoney},{$contactman},{$totalmoney},{$servicecharges},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			.Write "     <dd> <div>商城发放优惠券提示：<font><a href='javascript:' onclick=""$('#Sms11').val('您好{$username},在{$sitename}获得购物优惠券，券号{$couponnum},金额{$money}元，请于{$enddate}前使用。')"">默认</a></font></div><textarea name=""Sms11"" id=""Sms11"" cols=80 rows=4 class='textbox'>" & SmsArr(11) & "</textarea>  <span class=""block"">可用标签{$username},{$couponnum},{$money},{$enddate}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			
            .Write "</font>"
			'====================商城相关提醒_end===================================
			
			'====================求职系统_begin=====================================
			if KS.ChkClng(conn.execute("select top 1 ChannelStatus from ks_channel where channelid=10")(0))=1 then
			 .Write "<dd>"
			else
			  .Write "<dd style=""display:none;"">"
			end if
			.Write "  <div>招聘单位收到简历提示：<font><a href='javascript:' onclick=""$('#Sms20').val('尊敬的{$company}负责人,在{$sitename}网{$realname}向您投放了{$jobzw}的应聘简历，请及时登录处理。')"">默认</a></font></div><textarea name=""Sms20"" id=""Sms20"" cols=80 rows=4 class='textbox'>" & SmsArr(20) & "</textarea>  <span class=""block"">可用标签{$company},{$realname},{$jobzw},{$time}。<br><font color=blue>说明：留空表示不发送</font>  </span></dd>"
			'====================求职系统_end=====================================
			
			
			
			.Write "     <dd> <div>短信签名：</div>"
			Dim SmsSignArr:SmsSignArr=split(Setting(156)&"∮∮∮∮","∮")
			.Write "        <input type=""text"" class=""textbox"" name=""SmsSign1"" value=""" & SmsSignArr(0) & """ size=""12""><span>短信接口商如果要求签名，请在些输入，如：【KESION】 </span>      </dd>"
			.Write "     <dd> <div>验证码没收到重新发送间隔时间：</div>"
			.Write "        <input type=""text"" style=""text-align:center"" class=""textbox"" name=""SmsSign2"" value=""" & SmsSignArr(1) & """ size=""12"">秒"
			.Write "        &nbsp;&nbsp;验证码有效时间<input type=""text"" style=""text-align:center"" class=""textbox"" name=""SmsSign5"" value=""" & SmsSignArr(4) & """ size=""12"">分种 <span>不限制请输入“0”，建议设置10分钟。</span>"
			.Write "      </dd><dd> <div>验证码每天发送限制<font>(防止恶意发送,不限制输入“0”)</font>：</div>"
			.Write "        每个手机号同一天最多发<input type=""text"" style=""text-align:center"" class=""textbox"" name=""SmsSign3"" value=""" & SmsSignArr(2) & """ size=""12"">次 每个IP同一天最多发<input type=""text"" style=""text-align:center"" class=""textbox"" name=""SmsSign4"" value=""" & SmsSignArr(3) & """ size=""12"">次 </dd>"
			.Write "</font>"
			.Write "   </dl>"
			.Write "</div>"
			
			.Write " </form>"
		    .Write "</div>"
				 

			

			.Write "<div style=""clear:both;text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
			.Write "<div style=""height:30px;text-align:center"">KeSion CMS X" & GetVersion & ", Copyright (c) 2006-" & year(now) &" <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			'.Write " setlience("&Setting(22) &");"&vbcrlf
			.Write " setsendmail(" &Setting(146) & ");" & vbcrlf
			.Write "function setlience(n)" & vbcrlf
			.Write "{" & vbcrlf
			.Write "  if (n==0)"  &vbcrlf
			.Write "   $('#liencearea').hide();" & vbcrlf
			.Write "  else" & vbcrlf
			.Write "   $('#liencearea').show(); " & vbcrlf
			.Write "}" & vbcrlf
			.Write "function setsendmail(n)" & vbcrlf
			.Write "{" & vbcrlf
			.Write "  if (n==0)"  &vbcrlf
			.Write "    $('#sendmailarea').hide();" & vbcrlf
			.Write "  else" & vbcrlf
			.Write "   $('#sendmailarea').show(); " & vbcrlf
			.Write "}" & vbcrlf

			.Write " function CheckForm()" & vbCrLf
			.Write " { " & vbCrLf
			.Write "     $('#myform').submit();"
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
		
		Sub balance()
		 Dim url:url=split(KS.Setting(158)&"∮","∮")(0)
		 url=replace(url,"{$user}",KS.Setting(152))
		 dim passarr:passarr=split(KS.Setting(153)&"∮","∮")
		 dim pass:pass=passarr(0)
		 dim passType:passType=passarr(1)
		 if passType="1" Then
		  pass=md5(pass,16)
		 ElseIF PassType="2" Then
		  pass=md5(pass,32)
		 END iF
		 url=replace(url,"{$pass}",pass)
		 dim rstr:rstr=lcase(ks.do_post(url,"gbk"))
		 ks.die KS.CutFixContent(rstr, lcase(split(KS.Setting(158)&"∮∮","∮")(1)), lcase(split(KS.Setting(158)&"∮∮","∮")(2)), 1)
		End Sub
	
		
		'系统空间占用量
		Sub GetSpaceInfo()
			Dim SysPath, FSO, F, FC, I, I2
			Response.Write ("<title>空间查看</title>")
			Response.Write ("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
			Response.Write ("<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>")
			Response.Write ("<BODY leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>")
			Response.Write ("<div class='topdashed'><a href='?action=CopyRight'>服务器参数探测</a> | <a href='?action=Space'>系统空间占用量</a></div>")
			Response.Write ("<div class='pageCont2'>")
			Response.Write ("<table width='100%' border='0' cellspacing='0' cellpadding='0' oncontextmenu=""return false"">")
			Response.Write ("  <tr>")
			Response.Write ("    <td valign='top'>")
            Response.Write ("<table width=90% border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
     
         SysPath = Server.MapPath("\") & "\"
                 Set FSO = KS.InitialObject(KS.Setting(99))
                  Set F = FSO.GetFolder(SysPath)
                  Set FC = F.SubFolders
                            I = 1
                            I2 = 1
               For Each F In FC
				Response.Write ("        <tr>")
				Response.Write ("          <td height=25 bgcolor='#EEF8FE'><img src='../Images/Folder/folderclosed.gif' width='20' height='20' align='absmiddle'><b>" & F.name & "</b>&nbsp; 占用空间：&nbsp;<img src='../../Images/default/bar.gif' width=" & Drawbar(F.name) & " height=10>&nbsp;")
					ShowSpaceInfo (F.name)
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
							  I = I + 1
								  If I2 < 10 Then
									I2 = I2 + 1
								  Else
									I2 = 1
								 End If
								 Next
						  
				Response.Write ("        <tr>")
				Response.Write ("          <td height='25' bgcolor='#EEF8FE'> 程序文件占用空间：&nbsp;<img src='../Images/default/' width=" & Drawspecialbar & " height=10>&nbsp;")
				
				Showspecialspaceinfo ("Program")
				
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table>")
				Response.Write ("      <table width=90% border=0 align='center' cellpadding=3 cellspacing=1>")
				Response.Write ("        <tr>")
				Response.Write ("          <td height='28' align='right' bgcolor='#FFFFFF'><font color='#FF0066'><font color='#006666'>系统占用空间总计：</font>")
				Showspecialspaceinfo ("All")
				Response.Write ("            </font> </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table></td>")
				Response.Write ("  </tr>")
				Response.Write ("</table>")
				Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS X" & GetVersion &", Copyright (c) 2006-" & year(now)&" <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div></div>"
				Response.Write ("</body>")
				Response.Write ("</html>")
		End Sub
		Sub ShowSpaceInfo(drvpath)
        Dim FSO, d, size, showsize
        Set FSO = KS.InitialObject(KS.Setting(99))
        Set d = FSO.GetFolder(Server.MapPath("/" & drvpath))
        size = d.size
        showsize = size & "&nbsp;Byte"
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;KB"
        End If
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;MB"
        End If
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;GB"
        End If
        Response.Write "<font face=verdana>" & showsize & "</font>"
      End Sub
	  Sub Showspecialspaceinfo(method)
			Dim FSO, d, FC, f1, size, showsize, drvpath
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			 If method = "All" Then
				size = d.size
			ElseIf method = "Program" Then
				Set FC = d.Files
				For Each f1 In FC
					size = size + f1.size
				Next
			End If
			showsize = round(size,2) & "&nbsp;Byte"
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;KB"
			End If
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;MB"
			End If
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;GB"
			End If
			Response.Write "<font face=verdana>" & showsize & "</font>"
		End Sub
		Function Drawbar(drvpath)
			Dim FSO, drvpathroot, d, size, totalsize, barsize
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			totalsize = d.size
			Set d = FSO.GetFolder(Server.MapPath("/" & drvpath))
			size = d.size
			
			barsize = CInt((size / totalsize) * 100)
			Drawbar = barsize
		End Function
		Function Drawspecialbar()
			Dim FSO, drvpathroot, d, FC, f1, size, totalsize, barsize
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			totalsize = d.size
			Set FC = d.Files
			For Each f1 In FC
				size = size + f1.size
			Next
			barsize = CInt((size / totalsize) * 100)
			Drawspecialbar = barsize
		End Function

       '查看组件支持情况
	   Sub GetDllInfo()
	    Dim theInstalledObjects(17)
	   	theInstalledObjects(0) = "MSWC.AdRotator"
		theInstalledObjects(1) = "MSWC.BrowserType"
		theInstalledObjects(2) = "MSWC.NextLink"
		theInstalledObjects(3) = "MSWC.Tools"
		theInstalledObjects(4) = "MSWC.Status"
		theInstalledObjects(5) = "MSWC.Counters"
		theInstalledObjects(6) = "IISSample.ContentRotator"
		theInstalledObjects(7) = "IISSample.PageCounter"
		theInstalledObjects(8) = "MSWC.PermissionChecker"
		theInstalledObjects(9) = KS.Setting(99)
		theInstalledObjects(10) = "adodb.connection"
		theInstalledObjects(11) = "SoftArtisans.FileUp"
		theInstalledObjects(12) = "SoftArtisans.FileManager"
		theInstalledObjects(13) = "JMail.SMTPMail"
		theInstalledObjects(14) = "CDONTS.NewMail"
		theInstalledObjects(15) = "Persits.MailSender"
		theInstalledObjects(16) = "LyfUpload.UploadFile"
		theInstalledObjects(17) = "Persits.Upload.1"


		 Response.Write ("<table width='99%' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#CDCDCD'>")
		 Response.Write ("   <form method='post' action='?Action=CopyRight'>")
		 Response.Write ("<tr>")
		 Response.Write ("     <td height=36 bgcolor='#FFFFFF'>服务器组件探测查询-&gt; <font color='#FF0000'>组件名称:</font>")
		 Response.Write ("       <input type='text' name='classname' class='textbox' style='width:180'>")
		 Response.Write ("     <input type='submit' name='Submit' class='button' value='测 试'>")
			 
		Dim strClass:strClass = Trim(Request.Form("classname"))
		If "" <> strClass Then
		Response.Write "<br>您指定的组件的检查结果："
		If Not IsObjInstalled(strClass) Then
		Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strClass & "组件！</font>"
		Else
		Response.Write "<br><font color=green>恭喜！该服务器支持" & strClass & "组件。</font>"
		End If
		Response.Write "<br>"
		End If
		Response.Write ("</font>")
		Response.Write ("      </td>")
		Response.Write ("  </tr></form>")
		Response.Write (" <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'><b><font color='#006666'> 　IIS自带组件</font></b></font></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>组 件 名 称</td>")
		Response.Write ("          <td width='15%'>支 持</td>")
		Response.Write ("          <td width='15%'>不支持</td>")
		Response.Write ("        </tr>")
			  
		Dim I
		For I = 0 To 10
		Response.Write "<TR align=center bgcolor=""#EEF8FE"" height=22><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 9
		Response.Write "(FSO 文本文件读写)"
		Case 10
		Response.Write "(ACCESS 数据库)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
		Else
		Response.Write "<td><b>√</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'> <font color='#006666'><b>　其他常见组件</b></font>")
		Response.Write ("    </td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>组 件 名 称</td>")
		Response.Write ("          <td width='15%'>支 持</td>")
		Response.Write ("          <td width='15%'>不支持</td>")
		Response.Write ("        </tr>")
			 
		For I = 11 To UBound(theInstalledObjects)
		Response.Write "<TR align=center height=18 bgcolor=""#EEF8FE""><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 11
		Response.Write "(SA-FileUp 文件上传)"
		Case 12
		Response.Write "(SA-FM 文件管理)"
		Case 13
		Response.Write "(JMail 邮件发送)"
		Case 14
		Response.Write "(CDONTS 邮件发送 SMTP Service)"
		Case 15
		Response.Write "(ASPEmail 邮件发送)"
		Case 16
		Response.Write "(LyfUpload 文件上传)"
		Case 17
		Response.Write "(ASPUpload 文件上传)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
		Else
		Response.Write "<td><b>√</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("</table>")
		Response.Write ("</td>")
		Response.Write ("</tr>")
		Response.Write ("</table>")
		End Sub
		
		'系统版权及服务器参数测试
		Sub GetCopyRightInfo()
	%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class='topdashed'> <a href="?action=CopyRight">服务器参数探测</a> | <a href="?action=Space">系统空间占用量</a></div>
<div class="pageCont2">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width=1 bgcolor="#E3E3E3"></td>
          <td width="1011"><div align="left"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <%
				Dim theInstalledObjects(23)
				theInstalledObjects(0) = "MSWC.AdRotator"
				theInstalledObjects(1) = "MSWC.BrowserType"
				theInstalledObjects(2) = "MSWC.NextLink"
				theInstalledObjects(3) = "MSWC.Tools"
				theInstalledObjects(4) = "MSWC.Status"
				theInstalledObjects(5) = "MSWC.Counters"
				theInstalledObjects(6) = "IISSample.ContentRotator"
				theInstalledObjects(7) = "IISSample.PageCounter"
				theInstalledObjects(8) = "MSWC.PermissionChecker"
				theInstalledObjects(9) = KS.Setting(99)
				theInstalledObjects(10) = "adodb.connection"
					
				theInstalledObjects(11) = "SoftArtisans.FileUp"
				theInstalledObjects(12) = "SoftArtisans.FileManager"
				theInstalledObjects(13) = "JMail.SMTPMail"
				theInstalledObjects(14) = "CDONTS.NewMail"
				theInstalledObjects(15) = "Persits.MailSender"
				theInstalledObjects(16) = "LyfUpload.UploadFile"
				theInstalledObjects(17) = "Persits.Upload.1"
				theInstalledObjects(18) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
				theInstalledObjects(19)	= "Persits.Jpeg"				'AspJpeg
				theInstalledObjects(20) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
				theInstalledObjects(21) = "sjCatSoft.Thumbnail"
				theInstalledObjects(22) = "Microsoft.XMLHTTP"
				theInstalledObjects(23) = "Adodb.Stream"
	%>      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="100%" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666">使用本系统，请确认您的服务器和您的浏览器满足以下要求：</font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="99%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="22">　<font face="Verdana, Arial, Helvetica, sans-serif">JRO.JetEngine</font><span class="small2">：</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject("JRO.JetEngine")
		if err=0 then 
		  Response.Write("<font color=#0076AE>√</font>")
		else
          Response.Write("<font color=red>×</font>")
		end if	 
		err=0
		Response.Write(" (ADO 数据对象):")
		 On Error Resume Next
	    KS.InitialObject("adodb.connection")
		if err=0 then 
		  Response.Write("<font color=#0076AE>√</font>")
		else
          Response.Write("<font color=red>×</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td width="52%" height="22"> 　当前数据库　 
                  <%
		If DataBaseType = 1 Then
		Response.Write "<font color=#0076AE>MS SQL</font>"
		else
		Response.Write "<font color=#0076AE>ACCESS</font>"
		end if
	  %>                  </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="22">　<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">FSO</font></span>文本文件读写<span class="small2">：</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject(KS.Setting(99))
		if err=0 then 
		  Response.Write("<font color=#0076AE>支持√</font>")
		else
          Response.Write("<font color=red>不支持×</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td height="22">　Microsoft.XMLHTTP 
                    <%If  Not IsObjInstalled(theInstalledObjects(22)) Then%>
                    <font color="red">×</font> 
                    <%else%>
                    <font color="0076AE"> √</font> 
                    <%end if%>
                    　Adodb.Stream 
                   <%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
                    <font color="red">×</font> 
                    <%else%>
                    <font color="0076AE"> √</font> 
                    <%end if%>                  </td>
                </tr>
                
                <tr bgcolor="#EEF8FE"> 
                  <td height="22" colspan="2">　客户端浏览器版本： 
                    <%
	  Dim Agent,Browser,version,tmpstr
	  Agent=Request.ServerVariables("HTTP_USER_AGENT")
	  Agent=Split(Agent,";")
	  If InStr(Agent(1),"MSIE")>0 Then
				Browser="MS Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
	Response.Write(""&Browser&"  "&version&"")
	  %>
                    [需要IE5.5或以上,服务器建议采用Windows 2000或Windows 2003 Server]</td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="99%" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666">服务器信息</font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="99%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器类型：<font face="Verdana, Arial, Helvetica, sans-serif"><%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</font></td>
                  <td height="25">　<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">WEB</font></span>服务器的名称和版本<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("SERVER_SOFTWARE")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　返回服务器的主机名，<font face="Verdana, Arial, Helvetica, sans-serif">IP</font>地址<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("SERVER_NAME")%></font></font></td>
                  <td width="52%" height="25">　服务器操作系统<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("OS")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　站点物理路径<font face="Verdana, Arial, Helvetica, sans-serif">：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
                  <td width="52%" height="25">　虚拟路径<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SCRIPT_NAME")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　脚本超时时间<span class="small2">：</span><font color=#0076AE><%=Server.ScriptTimeout%></font> 秒</td>
                  <td width="52%" height="25">　脚本解释引擎<span class="small2">：</span><font face="Verdana, Arial, Helvetica, sans-serif"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>　</font> </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器端口<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SERVER_PORT")%></font></td>
                  <td height="25">　协议的名称和版本<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SERVER_PROTOCOL")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器 <font face="Verdana, Arial, Helvetica, sans-serif">CPU</font> 
                    数量<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></font> 个　</td>
                  <td height="25">　客户端操作系统： 
                    <%
 dim thesoft,vOS
thesoft=Request.ServerVariables("HTTP_USER_AGENT")
if instr(thesoft,"Windows NT 5.0") then
	vOS="Windows 2000"
elseif instr(thesoft,"Windows NT 5.2") then
	vOs="Windows 2003"
elseif instr(thesoft,"Windows NT 5.1") then
	vOs="Windows XP"
elseif instr(thesoft,"Windows NT") then
	vOs="Windows NT"
elseif instr(thesoft,"Windows 9") then
	vOs="Windows 9x"
elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	vOs="类Unix"
elseif instr(thesoft,"Mac") then
	vOs="Mac"
else
	vOs="Other"
end if
Response.Write(vOs)
%> </td>
                </tr>
              </table>
			  <%
			  GetDllInfo
			  %>
			   <table width="99%" height="30" border="0" align="center" cellpadding="0" cellspacing="0" >
                <tr> 
                  <td align="left">　<font color="#006666">系统版本信息</font></td>
                </tr>
              </table>
             <script>
			 function showbigpic(){
				var box=top.$.dialog({title:'查看KesionCMS相关证书：',content: '<style>.zs{width:890px;}.zs li img{border:1px solid #000;margin:5px;width:199px;height:220px;}.zs li{width:200px;float:left;margin:10px;}</style><div class="zs"><ul><li><a href="http://www.kesion.com/images/zs/kesioncmsr.jpg" target="_blank"><img src="http://www.kesion.com/images/zs/kesioncmsr.jpg" title="KesionCMS商标证书"/></a></li><li><a href="http://www.kesion.com/images/zs/kesioncms9.png" target="_blank"><img src="http://www.kesion.com/images/zs/kesioncms9.png" title="KesionCMS着作权证书"/></a></li><li><a href="http://www.kesion.com/images/zs/kesioncms9.png" target="_blank"><img src="http://www.kesion.com/images/zs/kesioncms9.png"  title="商城系统着作权证书"/></a></li><li><a href="http://www.kesion.com/images/zs/kesioneshop9.jpg" target="_blank"><img src="http://www.kesion.com/images/zs/kesioneshop9.jpg"  title="考试系统着作权证书"/></a></li><li><a href="http://www.kesion.com/images/icmszs.jpg" target="_blank"><img src="http://www.kesion.com/images/icmszs.jpg"  title="KesionICMS着作权证书"/></a></li><li><a href="http://www.kesion.com/images/imallzs.jpg" target="_blank"><img src="http://www.kesion.com/images/imallzs.jpg"  title="KesionIMALL着作权证书"/></a></li><li><a href="http://www.kesion.com/images/iexamzs.jpg" target="_blank"><img src="http://www.kesion.com/images/iexamzs.jpg"  title="KesionIEXAM着作权证书"/></a></li><li><a href="http://www.kesion.com/images/zs/kesioncms5.png" target="_blank"><img src="http://www.kesion.com/images/zs/kesioncms5.png"  title="KesionCMS着作权证书"/></a></li></ul></div>',max:false,min: false});
			}
			 </script>
              <table width="99%" height="63" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="30"> 　当前版本<font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td height="30">　<font color=red> 
                    <%=KS.Version%>
                    </font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="15%" height="30">　版权声明</td>
                  <td style="line-height:24px;">　1、本软件为共享软件,提供个人网站免费使用,非漳州科兴技术有限公司官方授权许可，不得将之用于盈利或非盈利性的商业用途;<br>
                    　2、用户自由选择是否使用,在使用中出现任何问题和由此造成的一切损失漳州科兴技术有限公司将不承担任何责任;<br>
                    　3、本软件受中华人民共和国《着作权法》《计算机软件保护条例》等相关法律、法规保护，软件制作权登记号：<a href="javascript:;" style="color:#CC0000" onClick="showbigpic()">2012SR058633</a>。漳州科兴技术有限公司保留一切权利。　 
                    <p></p></td>
                </tr>
              </table>
              <br>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</div>
</html>
<%
		End Sub
		
		Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = KS.InitialObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
		End Function
		
		'目录更名
		Function FolderReName(filename,name)
			    on error resume next
				dim Fso,MyFile
				set Fso=KS.InitialObject(KS.Setting(99))
				filename=server.MapPath(filename)
				Set MyFile = Fso.GetFolder(filename)
				MyFile.Name=name
				if err then 
				  FolderReName=false
				else
				  FolderReName=true
				end if
		end Function

	Public Function IsExpired(strClassString)
		On Error Resume Next
		IsExpired = True
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
	
		If 0 = Err Then
			Select Case strClassString
				Case "Persits.Jpeg"
					If xTestObjResponse.Expires > Now Then
						IsExpired = False
					End If
				Case "wsImage.Resize"
					If InStr(xTestObj.errorinfo, "已经过期") = 0 Then
						IsExpired = False
					End If
				Case "SoftArtisans.ImageGen"
					xTestObj.CreateImage 500, 500, RGB(255, 255, 255)
					If Err = 0 Then
						IsExpired = False
					End If
			End Select
		End If
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function ExpiredStr(I)
		   Dim ComponentName(3)
			ComponentName(0) = "Persits.Jpeg"
			ComponentName(1) = "wsImage.Resize"
			ComponentName(2) = "SoftArtisans.ImageGen"
			ComponentName(3) = "CreatePreviewImage.cGvbox"
			If IsObjInstalled(ComponentName(I)) Then
				If IsExpired(ComponentName(I)) Then
					ExpiredStr = "，但已过期"
				Else
					ExpiredStr = ""
				End If
			  ExpiredStr = " √支持" & ExpiredStr
			Else
			  ExpiredStr = "×不支持"
			End If
	End Function

End Class
%> 

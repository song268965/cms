<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_ClubSetting
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ClubSetting
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
			.Write ".tips {color: #999999;padding:2px}" & vbCrLf
			.Write ".txt {color: #666;border:1px solid #ccc;height:22px;line-height:22px}" & vbCrLf
			.Write "textarea {color: #666;border:1px solid #ccc;}" & vbCrLf
			.Write "-->" & vbCrLf
			.Write "</style>" & vbCrLf
			.Write "</head>" & vbCrLf

		     Call SetSystem()
			 
		 End With
		End Sub
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		With Response
			
			If Not KS.ReturnPowerResult(0, "KSMB10004") Then          '检查是否有基本信息设置的权限
					 .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back()';</script>")
					 Call KS.ReturnErr(1, "")
					 .End
			 End If
			

	
			SqlStr = "select * from KS_Config"
			Set RS = KS.InitialObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			
			 Dim Setting:Setting=Split(RS("Setting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
			 
			If KS.G("Flag") = "Edit" Then
			
			  Dim SetArr,SetStr,I
			  SetArr=Setting
			  For I=0 To Ubound(SetArr)
			   If I=0 Then 
				SetStr=SetArr(0)
			   ElseIf Request("Setting(" & I & ")")<>"" or i=69 or i=36 or i=37 or i=52 or i=68 or i=159 Then
				SetStr=SetStr & "^%^" & Request("Setting(" & I & ")")
			   Else
				SetStr=SetStr & "^%^" & SetArr(I)
			   End If
			  Next
			
				RS("Setting")=SetStr
				RS.Update
				RS.Close:Set RS=Nothing
				
				Conn.Execute("Update KS_Channel Set channelstatus=" & KS.ChkClng(Request("Setting(56)")) &" Where ChannelID=11")
				Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
			  
			   if KS.ChkClng(Request("Setting(56)"))=0 or Request("From")="model" Then
			     Session("FromFile")="System/KS.Model.asp"
				 KS.Die ("<script>top.$.dialog.alert('论坛参数配置修改成功！',function(){top.location.href='index.asp';})</script>")
			   else
			    KS.Die ("<script>top.$.dialog.alert('论坛参数配置修改成功！',function(){location.href='" & KS.Setting(3) & KS.Setting(89) & "Club/KS.GuestSetting.asp';})</script>")
			    End If
			End If
			
			.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=SetParam&OpStr=" & Server.URLEncode("论坛系统 >> <font color=red>参数配置</font>") & "';</script>")

			.Write "<body  bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<form name='myform' method=post action=""KS.GuestSetting.asp"" id=""myform"">"
			.Write " <input type=""hidden"" value=""Edit"" name=""Flag""/>"
			.Write " <input type=""hidden"" value=""" & Request("from") & """ name=""from""/>"
			.Write "<div class='topdashed sort menu_top_fixed'>论坛系统参数配置</div>"
			.Write "<div class=""menu_top_fixed_height""></div>"
			.Write "<div style='height:5px;overflow:hidden'></div>"
			.Write "<div class=tab-page id=clubconfigPanel>"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""clubconfigPanel"" ), 1 )"
            .Write " </SCRIPT>"
             
						                                                      '=====================================================论坛系统参数配置开始=========================================
			 .Write "<div class=tab-page id=GuestBook_Option>"
			 .Write " <H2 class=tab>论坛参数</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""GuestBook_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <dl class=""dtable"">"
			 .Write "    <dd>"
			.Write "      <div>论坛系统状态：</div>"
			.Write "     <label><input onclick=""$('#bbs').show()"" name=""Setting(56)"" type=""radio"" value=""1"""
			 If Setting(56)="1" Then .Write " Checked"
			 .Write ">开启</label>"
			 .Write "&nbsp;&nbsp;<label><input name=""Setting(56)"" onclick=""$('#bbs').hide()"" type=""radio"" value=""0"""
			 If Setting(56)="0" Then .Write " Checked"
			 .Write ">关闭</label>"
			 .Write "<span>当关闭论坛系统时，前台用户将不能使用。</span></dd>"
			 If Setting(56)="1" Then
			 .Write "<span id=""bbs"">"
			 Else
			 .Write "<span id=""bbs"" style=""display:None"">"
			 End If
			.Write "    <dd><div>本模块安装目录：</div>"
			.Write "   <input name=""Setting(66)"" type=""text"" class='textbox' value=""" & Setting(66) & """ size=""50"">"
			 .Write "<span>如:club,bbs等,不要带""/"",如果修改这里的配置请同时修改您的物理路径</span></dd>"

			.Write "    <dd><div>本模块绑定的域名：</div><input name=""Setting(69)"" type=""text"" class='textbox' value=""" & Setting(69) & """ size=""50"">"
			.Write "<span>如:bbs.kesion.com,www.kesion.cn等,不要带""http://"",如果不绑定请留空,否则可以导致页面路径出错,支持独立域名或二级域名的绑定</span></dd>"
			.Write "    <dd><div>是否开启伪静态：</div> <input  name=""Setting(70)"" type=""radio"" value=""1"""
			 If Setting(70)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(70)"" type=""radio"" value=""0"""
			 If Setting(70)="0" Then .Write " Checked"
			 .Write ">否 &nbsp;&nbsp;"
			 .Write "<span>需要服务器支持rewrite组件</span></dd>"

			 
			.Write "    <dd><div>显示标题名称：</div><input name=""Setting(61)"" type=""text""  value=""" & Setting(61) & """ size=""50"" class='textbox'> "
			 .Write "<span>请设置该子系统的名称,用于在位置导航及网站标题栏显示。如:科汛技术论坛,在线交流等</span></dd>"
			.Write "    <dd><div>项目名称：</div><input name=""Setting(62)"" class='textbox' type=""text""  value=""" & Setting(62) & """ size=""10""> "
			 .Write "<span>如:帖子,留言等</span></dd>"
			 
			 .Write "    <dd><div>发帖是否需要登录：</div> <input  name=""Setting(57)"" type=""radio"" value=""1"""
			 If Setting(57)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(57)"" type=""radio"" value=""0"""
			 If Setting(57)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>建议开启要登录才能发帖以增强发帖机的干扰。</span></dd>"
			
			 .Write "    <dd><div>论坛首页显示模式：</div> <input  name=""Setting(59)"" type=""radio"" value=""1"""
			 If Setting(59)="1" Then .Write " Checked"
			 .Write ">帖子列表模式"
			 .Write "&nbsp;&nbsp;<input name=""Setting(59)"" type=""radio"" value=""0"""
			 If Setting(59)="0" Then .Write " Checked"
			 .Write ">论坛版面列表"
			 .Write "<span></span></dd>"
			.Write " <dd><div>论坛首页默认布局：</div> <input  name=""Setting(53)"" type=""radio"" value=""1"""
			 If Setting(53)="1" Then .Write " Checked"
			 .Write ">左右布局"
			 .Write "&nbsp;&nbsp;<input name=""Setting(53)"" type=""radio"" value=""0"""
			 If Setting(53)="0" Then .Write " Checked"
			 .Write ">平板布局"
			 .Write "</dd>"			 
			 
			.Write "    <dd><div>首页帖子列表显示条数：</div><input name=""Setting(51)"" class='textbox' style='text-align:center' type=""text"" id=""WebTitle"" value=""" & Setting(51) & """ size=""10""> 条"
			 .Write "<span>论坛首页采用帖子列表时，每页显示的条数。</span></dd>"
			
			.Write "    <dd><div>允许自由切换分栏或平板：</div> <label><input name=""Setting(52)"" type=""checkbox"" value=""1"""
			If Setting(52)="1" Then .Write " checked"
			.Write "/>允许切换</label>"
			 
			 .Write "</dd>"
			.Write "    <dd><div>是否开放游客使用论坛搜索：</div> <input  name=""Setting(164)"" type=""radio"" value=""1"""
			 If Setting(164)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(164)"" type=""radio"" value=""0"""
			 If Setting(164)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "<span>搜索功能是较占用资源的搜索，如果访问量较大建议设置为不开放游客搜索</span></dd>"
			 
			.Write "    <dd><div>显示会员实名认证图标：</div> <input  name=""Setting(48)"" type=""radio"" value=""1"""
			 If Setting(48)="1" Then .Write " Checked"
			 .Write ">显示"
			 .Write "&nbsp;&nbsp;<input name=""Setting(48)"" type=""radio"" value=""0"""
			 If Setting(48)="0" Then .Write " Checked"
			 .Write ">不显示"
			 .Write "<span>如果设置为“显示”则在帖子详情页左边将显示会员实名认证图标</span></dd>"
			 

			
			.Write "    <dd><div>是否允许游客回复主题：<font>(如果各个版面启用用户组限制,则以版面设置为准)</font></div>"
			.Write "     <input  name=""Setting(54)"" type=""radio"" value=""1"""
			 If Setting(54)="1" Then .Write " Checked"
			 .Write ">只允许管理员回复<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""2"""
			 If Setting(54)="2" Then .Write " Checked"
			 .Write ">所有会员可回复,游客不可回复<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""3"""
			 If Setting(54)="3" Then .Write " Checked"
			 .Write ">所有人都可以回复，包括游客<br>"
			 .Write "</dd>"
			 
			 .Write "    <dd><div>会员发帖时间限制:</div>"
			
			 .Write "<input  name=""Setting(214)"" onclick=""$('#clubtime').show();"" type=""radio"" value=""1"""
			 If Setting(214)="1" Then .Write " Checked"
			 .Write ">开启"
			 .Write "&nbsp;&nbsp;<input onclick=""$('#clubtime').hide();"" name=""Setting(214)"" type=""radio"" value=""0"""
			 If Setting(214)="0" Then .Write " Checked"
			 .Write ">关闭 "
			.Write "<div class=""clear""></div><font id='clubtime'"
			If Setting(214)="0" Then response.write " style='display:none'"
			.Write ">从<input type='text'  <input type=""text""  class='textbox' style='text-align:center' size=""6"" value=""" & Setting(212) & """ name=""Setting(212)""  onKeyUp=""value=value.replace(/\D/g,'')"" onafterpaste=""value=value.replace(/\D/g,'')"" /> - <input type=""text""  class='textbox' style='text-align:center' size=""6"" value=""" & Setting(213) & """ name=""Setting(213)""  onKeyUp=""value=value.replace(/\D/g,'')"" onafterpaste=""value=value.replace(/\D/g,'')"" />点不能发帖 <span style='color:green'>不限制请都填0</span>"
			
				.Write " </font>"
			 .Write "<span>限制所有版面发帖时间,都设0则按版面设置为准</span></dd>"
			 
			 
			 .Write "    <dd><div>发帖IP是否可见：</div><input  name=""Setting(58)"" type=""radio"" value=""1"""
			 If Setting(58)="1" Then .Write " Checked"
			 .Write ">管理员可见<input  name=""Setting(58)"" type=""radio"" value=""2"""
			 If Setting(58)="2" Then .Write " Checked"
			 .Write ">版主和管理员可见<input  name=""Setting(58)"" type=""radio"" value=""3"""
			 If Setting(58)="3" Then .Write " Checked"
			 .Write ">开放显示IP"
			 .Write "&nbsp;&nbsp;<input name=""Setting(58)"" type=""radio"" value=""0"""
			 If Setting(58)="0" Then .Write " Checked"
			 .Write ">关闭显示IP</dd>"
			 
			.Write " <dd><div>上传附件存放目录：</div><input class='textbox' name=""Setting(67)"" type=""text"" value=""" & Setting(67) &""" size=""50""> "
			 .Write "<span>如ClubFiles则表示论坛的所有上传文件都上传到UploadFiles/ClubFiles/目录下,后面不要带""/""</span></dd>"
			 .Write "</span>"
			 .Write "   </dl>"
			
			 .Write "</div>"
				 
			.Write " <div class=tab-page id=club_ads>"
			.Write "  <H2 class=tab>论坛广告</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""club_ads"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <dl class=""dtable"">"
			 .Write "    <dd><div>帖子右侧随机设置：</div><font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(36)"" style=""width:450px;height:120px"">" & Setting(36) &"</textarea><span class=""block"">用于在帖子的右侧显示,不录入表示不显示广告</span></dd>"
			 .Write "    <dd><div>帖子顶部的随机广告设置：</div>"
			 .Write "    <font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(68)"" style=""width:450px;height:120px"">" & Setting(68) &"</textarea><span class=""block"">用于在帖子内容顶部显示,不录入表示不显示广告,建议使用文本广告</span></dd>"			
			  .Write "    <dd><div>帖子底部的随机广告设置：</div>"
			 .Write "    <font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(37)"" style=""width:450px;height:120px"">" & Setting(37) &"</textarea><span class=""block"">用于在帖子内容底部显示,不录入表示不显示广告,建议使用文本广告</span></dd>"
			 .Write "    <dd><div>论坛顶部广告设置：</div>"
			 .Write "   <font color=blue>支持HTML语法和JS代码，每条广告用""@""分开。</font><br/><textarea name=""Setting(159)"" style=""width:450px;height:120px"">" & Setting(159) &"</textarea><span class=""block"">显示在顶部导航下面,每行显示四列。在论坛模板里通过标签{$GetTopAdList}调用。</span></dd>"
			.Write "</dl>"
			.Write "</div>"	 
				 
				 
				 
			.Write " <div class=tab-page id=site-template>"
			.Write "  <H2 class=tab>模板绑定</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-template"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <dl class=""dtable"">"
			.Write "    <dd><div>论坛首页模板：</div> <input  class='textbox' name=""Setting(114)"" id=""Setting114"" type=""text"" value=""" & Setting(114) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting114')[0]") & " <a href='../club/index.asp' target='_blank' style='color:green'>页面:/club/index.asp</a></dd>"
			.Write "    <dd><div>论坛版面列表页模板：</div> <input  class='textbox' name=""Setting(172)"" id=""Setting172"" type=""text"" value=""" & Setting(172) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting172')[0]") & " <a href='../club/index.asp' target='_blank' style='color:green'>页面:/club/index.asp</a></dd>"
			.Write "    <dd><div>论坛帖子页模板：</div> <input  class='textbox' name=""Setting(160)"" id=""Setting160"" type=""text"" value=""" & Setting(160) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting160')[0]") & " <a href='../club/display.asp' target='_blank' style='color:green'>页面:/club/display.asp</a>"
			.Write "    </dd>"
			.Write "    <dd><div>论坛发帖页面模板：</div> <input class='textbox'  name=""Setting(115)"" id=""Setting115"" type=""text"" value=""" & Setting(115) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting115')[0]") & " <a href='../club/post.asp' target='_blank' style='color:green'>页面:/club/post.asp</a></dd>"
			.Write "    <dd><div>论坛搜索模板：</div> <input  class='textbox' name=""Setting(171)"" id=""Setting171"" type=""text"" value=""" & Setting(171) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting171')[0]") & " <a href='../club/query.asp' target='_blank' style='color:green'>页面:/club/query.asp</a>"
			.Write "    </dd>"
			
			.Write "  </dl>"
			.Write "</div>"
				 
				
			.Write " </form>"
		    .Write "</div>"
				 

			

			.Write "<div style=""clear:both;text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
			.Write "<div style=""height:30px;text-align:center"">KeSion CMS X" & GetVersion & ", Copyright (c) 2006-" & year(now) &" <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf


			.Write " function CheckForm()" & vbCrLf
			.Write " { " & vbCrLf
			.Write "     $('#myform').submit();"
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
	
		
		
		
		

End Class
%> 

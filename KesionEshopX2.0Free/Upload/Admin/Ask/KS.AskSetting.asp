<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Ask_Setting
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Setting
        Private KS,KSMCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSMCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call KS.DelCahe(KS.SiteSn & "_Config")
		 Call KS.DelCahe(KS.SiteSn & "_Date")
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		       Call SetSystem()
		End Sub
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		Dim SetType
		SetType = KS.G("SetType")
		With Response
			If Not KS.ReturnPowerResult(0, "WDXT10000") Then          '检查是否有基本信息设置的权限
			  Call KS.ReturnErr(1, "")
			 .End
			End If
	
			SqlStr = "select AskSetting from KS_Config"
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			
			Dim Setting:Setting=Split(RS(0)&"^%^^%^^%^^%^^%^^%^^%^^%^0^%^0^%^0^%^0^%^0","^%^")
			If KS.G("Flag") = "Edit" Then
			    Dim N					
			    Dim WebSetting
				For n=0 To 55
				   WebSetting=WebSetting & Replace(KS.G("Setting(" & n &")"),"^%^","") & "^%^"
				Next
				RS("AskSetting")=WebSetting
				RS.Update			
				RS.Close
				
				Conn.Execute("Update KS_Channel Set channelstatus=" & KS.ChkClng(Request("Setting(0)")) &" Where ChannelID=12")
				Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
			  
			    if KS.ChkClng(Request("Setting(0)"))=0 or Request("From")="model" Then 
			     Session("FromFile")="System/KS.Model.asp"
				 KS.Die ("<script>top.$.dialog.alert('问答参数修改成功！',function(){top.location.href='index.asp';})</script>")
				else
			     KS.Die ("<script>top.$.dialog.alert('问答参数修改成功！',function(){location.href='" & KS.Setting(3) & KS.Setting(89) & "ask/KS.AskSetting.asp';})</script>")
				end if
			  
				
				
			End If
			
		  	.Write "<!DOCTYPE html><html>"
			.Write "<title>问答参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
			.Write "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>"
			.Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "</head>" & vbCrLf

			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>问答参数配置</div>"
			.Write ""
			.Write "<div class=tab-page id=spaceconfig>"
			.Write "  <form name='myform' id='myform' method=post action="""" >"
			.Write "  <input type=""hidden"" value=""Edit"" name=""Flag""/>"
			.Write "  <input type=""hidden"" value=""" & KS.S("From") &""" name=""from"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""spaceconfig"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>基本参数</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <dl class=""dtable"">"
			.Write "    <dd><div>问答系统状态：</div>"
			
				.Write " <input type=""radio"" name=""Setting(0)"" value=""1"" "
				If Setting(0) = "1" Then .Write (" checked")
				.Write "> 打开"
				.Write "    <input type=""radio"" name=""Setting(0)"" value=""0"" "
				If Setting(0) = "0" Then .Write (" checked")
				.Write "> 关闭"

			.Write "    </dd>"
			
			.Write "    <dd><div>安装目录：</div>"
			 	.Write " <input type=""text""  class='textbox' name=""Setting(1)"" size=""50"" value=""" & Setting(1) & """>"

			 .Write "<span>--如ask等,必须以""/""结束。</span>"
			 .Write "   </dd>"
			.Write "    <dd><div>模块名称：</div>"
			 	.Write " <input type=""text"" class='textbox' name=""Setting(2)"" size=""50"" value=""" & Setting(2) & """>"

			 .Write "<span>--如""问吧""等。</span></td>"
			 .Write "   </dd>"
			 
			 .Write "   <dd><div>运行模式：</div>"
			 	
				.Write " <input type=""radio"" name=""Setting(16)"" value=""0"" "
				If Setting(16) = "0" Then .Write (" checked")
				.Write "> 动态"
				.Write "    <input type=""radio"" name=""Setting(16)"" value=""1"" "
				If Setting(16) = "1" Then .Write (" checked")
				.Write "> 伪静态(<font color=red>需要服务器支持Rewrite组件</font>)"

			 .Write "<div class=""clear""></div><br/>扩展名<input  class='textbox' type='text' name='Setting(17)' value='" & Setting(17) & "' size='10'>"
			 .Write "       <span>--更改此配置,需要修改ISAPI_Rewrite的配置文件httpd.ini</span>"
			 .Write "   </dd>"
			 
			
			 .Write "   <dd><div>是否开启提问：</div>"
				.Write " <input type=""radio"" name=""Setting(3)"" value=""1"" "
				If Setting(3) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(3)"" value=""0"" "
				If Setting(3) = "0" Then .Write (" checked")
				.Write "> 否"
			 .Write "   </dd>"
			 
			 .Write "   <dd><div>问题描述/回答长度控制：</div>"
			  
			  .Write "大于等于<input type=""text"" style=""text-align:center"" class='textbox' name=""Setting(4)"" size=""5"" value=""" & Setting(4) & """>小于等于<input type=""text"" class='textbox' name=""Setting(5)"" style=""text-align:center"" size=""5"" value=""" & Setting(5) & """> "
			  .Write "    <span>--不想控制,请填写0</span>"
			 .Write "   </dd>"

			 .Write "   <dd><div>提问题是否启用验证码：</div>"
				
				.Write " <input type=""radio"" name=""Setting(6)"" value=""1"" "
				If Setting(6) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(6)"" value=""0"" "
				If Setting(6) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "</dd>"
			 .Write "   <dd><div>审核模式： <font>(启用需要审核,则只有审核通过的问题或回答才会及时显示)</font></div>"
			 
				.Write " <label><input type=""radio"" name=""Setting(18)"" value=""1"" "
				If Setting(18) = "1" Then .Write (" checked")
				.Write "> 提问和回答都不需要审核</label><br/>"
				.Write "    <label><input type=""radio"" name=""Setting(18)"" value=""0"" "
				If Setting(18) = "0" Then .Write (" checked")
				.Write "> 提问需要审核，回答不需要审核</label><br/>"
				.Write "    <label><input type=""radio"" name=""Setting(18)"" value=""2"" "
				If Setting(18) = "2" Then .Write (" checked")
				.Write "> 提问和回答都需要审核</label><br/>"
				.Write "    <label><input type=""radio"" name=""Setting(18)"" value=""3"" "
				If Setting(18) = "3" Then .Write (" checked")
				.Write "> 提问不需要审核，回答需要审核</label><br/>"
				
				.Write "</dd>"
			 .Write "   <dd><div>允许上传附件：<font>(启用上传附件功能，提问或回答可以附加上传附件)</font></div>"
				
				.Write " <input type=""radio"" name=""Setting(42)"" onclick=""document.getElementById('fj').style.display='';"" value=""1"" "
				If Setting(42) = "1" Then .Write (" checked")
				.Write "> 允许"
				.Write "    <input type=""radio"" name=""Setting(42)"" onclick=""document.getElementById('fj').style.display='none';"" value=""0"" "
				If Setting(42) = "0" Then .Write (" checked")
				.Write "> 不允许"
				
			 If Setting(42)="1" Then
			  .Write "<div class=""clear""></div><font id='fj'>"
			 Else
				.Write "<div class=""clear""></div><font id='fj' style='display:none;'>"
			 End If
			 .Write "<font color=green>允许上传的文件类型：<input class='textbox' name=""Setting(43)"" type=""text"" value=""" & Setting(43) &""" size='30'>多个类型用|线隔开<br/>允许上传的文件大小：<input class='textbox' name=""Setting(44)"" type=""text"" value=""" & Setting(44) &""" style=""text-align:center"" size='8'>KB<br/>每天上传文件个数：<input class='textbox' name=""Setting(45)"" type=""text"" value=""" & Setting(45) &""" style=""text-align:center"" size='8'>个,不限制请填0</font><br/><br/>"
			 .Write "<br/><strong>允许在此版本上传附件的用户组:</strong>(<font>不限制请不要勾选</font>)"
			 .Write KS.GetUserGroup_CheckBox("Setting(46)",Setting(46),5)
			 .Write "</font>"
				.Write "</dd>"
				
				.Write "   <dd><div>允许查看最佳答案的用户组：<font>(不限制请不要勾选)</font></div>"
				
				.Write KS.GetUserGroup_CheckBox("Setting(52)",Setting(52),5)
			    .Write ""
				.Write "  </dd>"

				
				.Write "   <dd><div>提问题最长可设置的有效天数：</div><input class='textbox' type='text' name='Setting(41)' value='" & Setting(41) & "' style='text-align:center;width:50px'>天"
				.Write "  </dd>"

				
			    .Write "   <dd><div>是否允许回答：</div>"
				
				.Write " <input type=""radio"" name=""Setting(7)"" value=""1"" "
				If Setting(7) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(7)"" value=""0"" "
				If Setting(7) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </dd>"
			    .Write "   <dd><div>回答问题是否启用验证码：</div>"
				
				.Write " <input type=""radio"" name=""Setting(8)"" value=""1"" "
				If Setting(8) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(8)"" value=""0"" "
				If Setting(8) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </dd>"
				
				
				
				
			 .Write "   <dd><div>是否只能回答一次：</div>"
				
				.Write " <input type=""radio"" name=""Setting(9)"" value=""1"" "
				If Setting(9) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(9)"" value=""0"" "
				If Setting(9) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    <span>--在每个问题中是否每个人只能回答一次</span>"
				.Write "</dd>"
			    .Write "   <dd><div>提问者是否允许问题补充：</div>"
				
				.Write " <input type=""radio"" name=""Setting(10)"" value=""1"" "
				If Setting(10) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(10)"" value=""0"" "
				If Setting(10) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    <span>--提问人可以进一步的补充问题</span>"
				.Write "</dd>"
			    .Write "   <dd><div>提问人可以回答自己提问的问题：</div>"
				
				.Write " <input type=""radio"" name=""Setting(11)"" value=""1"" "
				If Setting(11) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(11)"" value=""0"" "
				If Setting(11) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </dd>"
			    .Write "   <dd><div>提问人可以删除用户的回答：</div>"
				
				.Write " <input type=""radio"" name=""Setting(12)"" value=""1"" "
				If Setting(12) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(12)"" value=""0"" "
				If Setting(12) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </dd>"
			    .Write "   <dd><div>是否允许游客提问题：</div>"
				
				.Write " <input type=""radio"" name=""Setting(47)"" value=""1"" "
				If Setting(47) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(47)"" value=""0"" "
				If Setting(47) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    <span>--游客提的问题只能由管理员设置最佳答案，并且使用系统函数标签查询调用问题时性能有所影响，常规情况下建议不要开启游客提问</span>"
				.Write "</dd>"
				

				
			    .Write "   <dd><div>是否允许游客回答问题：</div>"
				
				.Write " <input type=""radio"" name=""Setting(13)"" value=""1"" "
				If Setting(13) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(13)"" value=""0"" "
				If Setting(13) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </dd>"
				
				
			    .Write "   <dd><div>列表页每页显示条数：</div>"
				
				.Write " <input class='textbox' style=""text-align:center"" type=""text"" name=""Setting(14)"" value=""" & Setting(14) & """ size=""6"">条"
				
				.Write "    <span>--对应前台的showlist.asp</span></dd>"
			    .Write "   <dd><div>问题详情每页显示条数：</div>"
				
				.Write " <input class='textbox' style=""text-align:center"" type=""text"" name=""Setting(15)"" value=""" & Setting(15) & """ size=""6"">条"
				
				.Write "    <span>--对应前台的q.asp</span>"
				.Write "</dd>"
			 .Write " </dl>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=td-page>"
			.Write "  <H2 class=tab>专家团队</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""td-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <dl class=""dtable"">"
            .Write "    <dd><div>专家团队分类</div><textarea name=""Setting(48)"" style='height:250px' cols='40' rows='8' class='textbox' id=""Setting48"">" & Setting(48) & "</textarea><span class=""block"">TIPS:一行一个分类</span>"
			.Write "    </dd>" 
			.Write "  </dl>"
			.Write "</div>"
			 
			.Write " <div class=tab-page id=template-page>"
			.Write "  <H2 class=tab>问答模板</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""template-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <dl class=""dtable"">"
            .Write "    <dd><div>问答首页模板：</div><input name=""Setting(20)"" class='textbox' id=""Setting20"" type=""text"" value=""" & Setting(20) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting20')[0]") & "<a href='../" & KS.ASetting(1) & "index.asp' target='_blank' style='color:blue'>index.asp</a>"
			.Write "    </dd>"            
			.Write "    <dd><div>提问问题模板：</div>"
			.Write "     <input name=""Setting(21)"" class='textbox' id=""Setting21"" type=""text"" value=""" & Setting(21) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting21')[0]") & "<a href='../" & KS.ASetting(1) & "a.asp' target='_blank' style='color:blue'>a.asp</a>"
			.Write "    </dd>"
			.Write "    <dd><div>问题列表页模板：</div> <input name=""Setting(22)"" class='textbox' id=""Setting22"" type=""text"" value=""" & Setting(22) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting22')[0]") & "<a href='../" & KS.ASetting(1) & "showlist.asp' target='_blank' style='color:blue'>showlist.asp</a>"
			.Write "    </dd>"
			.Write "    <dd><div>问题内容页模板：</div>"
			.Write "    <input class='textbox' name=""Setting(23)"" id=""Setting23"" type=""text"" value=""" & Setting(23) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting23')[0]") & "<a href='../" & KS.ASetting(1) & "q.asp' target='_blank' style='color:blue'>q.asp</a>"
			.Write "    </dd>"
			.Write "    <dd><div>问题搜索页模板：</div>"
			.Write "   <input name=""Setting(24)"" class='textbox' id=""Setting24"" type=""text"" value=""" & Setting(24) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting24')[0]") & "<a href='../" & KS.ASetting(1) & "search.asp' target='_blank' style='color:blue'>search.asp</a>"
			.Write "    </dd>"
			.Write "    <dd><div>专列团队模板：</div>"
			.Write "      <td> <input name=""Setting(49)"" class=""textbox"" id=""Setting49"" type=""text"" value=""" & Setting(49) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting49')[0]") & "<a href='../" & KS.ASetting(1) & "zj.asp' target='_blank' style='color:blue'>zj.asp</a>"
			.Write "    </dd>"
			.Write "    <dd><div>所有问题列表模板：</div>"
			.Write "    <input name=""Setting(51)"" id=""Setting51"" class=""textbox"" type=""text"" value=""" & Setting(51) & """ size=""50"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting51')[0]") & "<a href='../" & KS.ASetting(1) & "all.asp' target='_blank' style='color:blue'>all.asp</a>"
			.Write "    </dd>"
			 .Write " </dl>"
			.Write " </div>"

			.Write " <div class=tab-page id=user-page>"
			.Write "  <H2 class=tab>积分设置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""user-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <dl class=""dtable"">"
			.Write "    <dd><div>用户回答一个问题所得的积分：</div>"
			 .Write "    <input class=""textbox"" type=""text"" name=""Setting(30)"" size=""10"" value=""" & Setting(30) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>回答被提问者采纳所额外奖励的积分：</div>"
			 .Write "     <input type=""text"" class=""textbox"" name=""Setting(31)"" size=""10"" value=""" & Setting(31) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>用户处理问题所得的积分：</div>"
			 .Write "   <input type=""text"" class=""textbox"" name=""Setting(32)"" size=""10"" value=""" & Setting(32) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>问题被选为精彩推荐提问者所得的积分：</div>"
			 .Write "    <input class=""textbox"" type=""text"" name=""Setting(33)"" size=""10"" value=""" & Setting(33) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>问题被选为精彩推荐最佳回答者所得的积分：</div>"
			 .Write "    <input type=""text"" class=""textbox"" name=""Setting(34)"" size=""10"" value=""" & Setting(34) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>用户发表一个问题所获的积分：</div>"
			 .Write "    <input type=""text""class=""textbox"" name=""Setting(35)"" size=""10"" value=""" & Setting(35) & """> 分  <span>设置成负数则表示提问题要消费</span>"
			 .Write "   </dd>"
			.Write "    <dd><div>匿名提问减去积分：</div>"
			 .Write "   <input type=""text"" class=""textbox"" name=""Setting(36)"" size=""10"" value=""" & Setting(36) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>问题被删除减去提问者积分：</div>"
			 .Write "    <input type=""text"" class=""textbox"" name=""Setting(50)"" size=""10"" value=""" & Setting(50) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>删除答案减去回答者积分：</div>"
			 .Write "   <input type=""text"" class=""textbox"" name=""Setting(37)"" size=""10"" value=""" & Setting(37) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>删除最佳答案减去回答者积分：</div>"
			 .Write "   <input type=""text"" class=""textbox"" name=""Setting(38)"" size=""10"" value=""" & Setting(38) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>删除未解决问题减去积分：</div>"
			 .Write "   <input type=""text"" class=""textbox"" name=""Setting(39)"" size=""10"" value=""" & Setting(39) & """> 分"
			 .Write "   </dd>"
			.Write "    <dd><div>删除已解决问题减去积分：</div>"
			 .Write "    <input type=""text"" class=""textbox"" name=""Setting(40)"" size=""10"" value=""" & Setting(40) & """> 分"
			 .Write "   </dd>"
			
			.Write " </dl>"
			.Write " </div>"
			

			.Write "<div style=""clear:both;text-align:center;color:#003300;padding-bottom:20px;"">--------------------------------------------------------------------------------<br/>KeSion CMS X" & GetVersion & ", Copyright (c) 2006-" & year(now) &" <a href=""http://www.kesion.com/"" target=""_blank"">KeSion.Com</a>. All Rights Reserved . </div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			
			.Write " function CheckForm()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "if ($('#Setting20').val()=='')" & vbCrLf
			.Write "{ top.$.dialog.alert('请选择问答首页模板!');" & vbCrLf
			.Write "  $('#Setting20').focus();" & vbCrLf
			.Write "  return false;" & vbCrLf
			.Write "}" & vbCrLf

			.Write "     $('#myform').submit();" & vbCrLf
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
	
		

End Class
%> 

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
Set KSCls = New Admin_ShopSetting
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ShopSetting
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
			
					If Not KS.ReturnPowerResult(5, "M510021") Then          '检查是否有基本信息设置的权限
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
				ElseIf I=82 Then
				  SetStr=SetStr & "^%^" & KS.G("Setting(82)")&"|" & KS.G("Setting(82)_1") & "|" & KS.G("Setting(82)_2")& "|" &KS.G("Setting(82)_3") 
			   ElseIf Request("Setting(" & I & ")")<>"" or  I=208 Or I=209 Or I=210 Or I=211 Or I=71 Or I=72 Then
				 SetStr=SetStr & "^%^" & Request("Setting(" & I & ")")
			   Else
				SetStr=SetStr & "^%^" & SetArr(I)
			   End If
			  Next
			
				RS("Setting")=SetStr
				RS.Update
				RS.Close:Set RS=Nothing

			    KS.Die ("<script>top.$.dialog.alert('商城系统配置成功！',function(){ location.href='" & KS.Setting(3) & KS.Setting(89) & "shop/KS.ShopSetting.asp';});</script>")

			End If
			
			.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=SetParam&OpStr=" & Server.URLEncode("系统设置 >> <font color=red>基本信息设置</font>") & "';</script>")

			.Write "<body  bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>商城系统参数配置</div>"
			.Write ""
			.Write "<div class=tab-page id=shopconfigPanel>"
			.Write " <form name='myform' method=post action=""KS.ShopSetting.asp"" id=""myform"">"
			.Write " <input type=""hidden"" value=""Edit"" name=""Flag""/>"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""shopconfigPanel"" ), 1 )"
            .Write " </SCRIPT>"
             
		
					
								 '=====================================================商城系统参数配置开始=========================================
			 .Write "<div class=tab-page id=Shop_Option>"
			 .Write "<H2 class=tab>基本参数</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""Shop_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			.Write "  <dl class=""dtable"">"
			
			 .Write "<dd><div>前缀设置：</div>"
			.Write " 订单编号前缀<input class='textbox' style='text-align:center' name=""Setting(71)"" size=""6"" value=""" & Setting(71) & """>"
			 .Write " 在线支付单编号前缀： <input class='textbox' style='text-align:center' name=""Setting(72)"" size=""6"" value=""" & Setting(72) & """><span>不加前缀请留空</span>"
			 .Write "</dd>"
			 .Write "    <dd><div>商城付款方式：</div>"
			Dim PArr:Parr=Split(Setting(82)&"|0|0|0|0||||","|")
			.Write "①<label><input type='radio' name='Setting(82)'"
			If Parr(0)="1" Then .Write " checked"
			.Write " value='1'>一次性付款</label><br/>"
			.Write "②<label><input type='radio' name='Setting(82)'"
			If Parr(0)="2" Then .Write " checked"
			.Write " value='2'>不允许一次性付款，只能固定付订单总款的<input name=""Setting(82)_1"" style=""text-align:center"" class=""textbox"" size=""3"" value=""" & Parr(1) & """> % 作为定金</label><br/>"
			.Write "③<label><input type='radio' name='Setting(82)'"
			If Parr(0)="3" Then .Write " checked"
			.Write " value='3'>可以付全款，也可以付定金，但定金不能少于订单总款的<input style=""text-align:center"" class=""textbox"" name=""Setting(82)_2"" size=""4"" value=""" & Parr(2) & """> % <Br/>当选择第②种或第③种付款方式时，如果订单总款小于<input style=""text-align:center"" class=""textbox"" name=""Setting(82)_3"" size=""4"" value=""" & Parr(3) & """> 元时,则按订单全额付款</label>"
			
			.Write "</dd>"
			
			.Write "    <dd><div>是否允许游客购买商品: </div>"
			.Write "   <input  name=""Setting(63)"" type=""radio"" value=""1"""
			 If Setting(63)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(63)"" type=""radio"" value=""0"""
			 If Setting(63)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</dd>"

			.Write "    <dd><div>是否启用只有管理员后台确认的订单才能付款: </div><input  name=""Setting(49)"" type=""radio"" value=""1"""
			 If Setting(49)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(49)"" type=""radio"" value=""0"""
			 If Setting(49)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</dd>"
			 
			 .Write "    <dd><div>新订单支付成功通知管理员设置：</div>"
			.Write "<input type='checkbox' name='Setting(208)' value='1'"
			If Setting(208)="1" Then .Write "checked"
			.Write ">后台弹出提示"
			.Write "&nbsp;<input type='checkbox' name='Setting(209)' value='1'"
			If Setting(209)="1" Then .Write "checked"
			.Write ">手机短信提示"
			.Write "&nbsp;<input type='checkbox' name='Setting(210)' value='1'"
			If Setting(210)="1" Then .Write "checked"
			.Write ">电子邮件提示<br/>"
			.Write "接收新订单通知邮箱：<input name=""Setting(211)"" class='textbox' type=""text""  value=""" & Setting(211) & """ size=40><span>多个邮件要接收通知，请用英文逗号隔开。</span>"
			.Write "</dd>"
			 
			 
			 .Write "    <dd><div>会员交易管理费<font>(仅当启用会员可以在本商发布商品销售时有效。相当于交易中介服务费用)</font>：</div>"
			.Write "     总交易金额的<input class='textbox' name=""Setting(79)"" style=""text-align:center"" size=""6"" value=""" & Setting(79) & """>% <span>会员成功在本站销售商品收取的交易管理费。。</span>"
			
			.Write "</dd>"
			 
			.Write "    <dd><div>商品价格是否含税：</div><input onclick=""$('#rate').hide();"" name=""Setting(64)"" type=""radio"" value=""1"""
			 If Setting(64)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input onclick=""$('#rate').show();"" name=""Setting(64)"" type=""radio"" value=""0"""
			 If Setting(64)="0" Then .Write " Checked"
			 .Write ">否"
			 
			 .Write "<div class=""clear""></div><font id='rate'"
			 If Setting(64)="1" Then .Write " style='display:none'"
			 .Write">税率设置： <input class='textbox' name=""Setting(65)"" style=""text-align:center"" size=""6"" value=""" & Setting(65) & """>%</font>"
			 
			 .Write "</dd>"
			.Write " <dd><div>客户需要另外支付运费：</div>"
			.Write " <input name=""Setting(180)"" onclick=""$('#yf').show()"" type=""radio"" value=""1"""
			
			 If Setting(180)="1" Then .Write " Checked"
			 .Write ">需要"
			 .Write "&nbsp;&nbsp;<input name=""Setting(180)"" onclick=""$('#yf').hide()"" type=""radio"" value=""0"""
			 If Setting(180)="0" Then .Write " Checked"
			 .Write ">不需要" 
			 If Setting(180)="1" Then
			  .Write "<div class=""clear""></div><font id='yf'>"
			 Else
			  .Write "<div class=""clear""></div><font id='yf' style='display:none'>"
			 End If 
			  .Write "买满  <input name=""Setting(207)"" class='textbox' style=""text-align:center"" value=""" & Setting(207) & """ size=""10"" type=""text""  onKeyUp=""value=value.replace(/\D/g,'')"" onafterpaste=""value=value.replace(/\D/g,'')"" >元 免快递费"
			 .Write "</font></dd>"
			 .Write "<dd><div>允许积分扣减购物金额：</div>"
			 .Write " <input onclick=""$('#scores').show();"" name=""Setting(181)"" type=""radio"" value=""1"""
			
			 If Setting(181)="1" Then .Write " Checked"
			 .Write ">允许"
			 .Write "&nbsp;&nbsp;<input onclick=""$('#scores').hide();"" name=""Setting(181)"" type=""radio"" value=""0"""
			 If Setting(181)="0" Then .Write " Checked"
			 .Write ">不允许"
			 .write "<div style='"
			 If KS.ChkClng(Setting(181))=0 then .write "display:none;"
			 .Write "padding:5px;margin:3px;font-weight:normal;' id='scores'>"
			 .write "<FIELDSET><LEGEND align=left></LEGEND>抵扣比率： <input type='text' class='textbox' name='Setting(182)' value='" & Setting(182) &"' style='text-align:center;width:30px'/> 积分=1元 <br/> 限制订单总金额大于等于<input type='text' class='textbox' name='Setting(183)' value='" & Setting(183) &"' style='text-align:center;width:30px'/>元时才能使用,抵扣金额不能大于订单总金额的<input type='text' class='textbox' name='Setting(184)' value='" & Setting(184) &"' style='text-align:center;width:30px'/> % <span class='tips'>tips:不限制请输入0</span>"
			 
			 .Write "</FIELDSET></div></dd>"
			 
			.Write " <dd><div>美元汇率：</div><input class='textbox' name=""Setting(81)"" style=""text-align:center"" size=""6"" value=""" & Setting(81) & """>  <span>如:1美元=6.7784人民币元 则这里填6.7784 ，当启用paypal国际版支付平台时，系统将根据此汇率将人民币转换为美元进行支付"
			 .Write "</span></dd>"
			
			 
			
			
			.Write "    <dd><div>团购是否开启伪静态</div>"
			.Write " <label><input type='radio' name='Setting(179)'"
			If Setting(179)="0" Then .Write " checked"
			.Write " value='0'>不开启</label>"
			.Write " <label><input type='radio' name='Setting(179)'"
			If Setting(179)="1" Then .Write " checked"
			.Write " value='1'>开启</label><span>需要服务器支持Rewrite组件</span>"
			.Write "</dd>"	
			.Write "</dl>"	
			.Write "</div>"
			
			.Write "<div class=tab-page id=email_Option>"
			.Write "<H2 class=tab>站内短信/Email模板</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""email_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			.Write "  <dl class=""dtable"">"
			.Write "    <dd><div>订单确认站内短信/Email通知内容：</div>"
			.Write "<textarea name='Setting(73)' cols='60' rows='4'>" & Setting(73) & "</textarea>"
			.Write "<span class=""block"">支持HTML代码，可用标签详见下面的标签说明</span>"	
			.Write "</dd>"
			.Write " <dd><div>收到汇款后站内短信/Email通知内容：</div><textarea name='Setting(74)' cols='60' rows='4'>" & Setting(74) & "</textarea>"
		   .Write "<span class=""block"">支持HTML代码，可用标签详见下面的标签说明</span>"	
			.Write "</dd>"
			.Write " <dd><div>退款后站内短信/Email通知内容：</div>"
			.Write " <textarea name='Setting(75)' cols='60' rows='4'>" & Setting(75) & "</textarea>"
			.Write "<span class=""block"">支持HTML代码，可用标签详见下面的标签说明</span>"	
			.Write "</dd>"
			.Write "    <dd><div>开发票后站内短信/Email通知内容：</div>"
			.Write "  <textarea name='Setting(76)' cols='60' rows='4'>" & Setting(76) & "</textarea>"
			.Write "<span class=""block"">支持HTML代码，可用标签详见下面的标签说明</span></dd>"	
			.Write "    <dd><div>发出货物后站内短信/Email通知内容：</div>"
			.Write "   <textarea name='Setting(77)' cols='60' rows='4'>" & Setting(77) & "</textarea>"
			.Write "<span>支持HTML代码，可用标签详见下面的标签说明</span>"
			.Write "</dd>"
			.Write "<dd><div>支付货款给卖方的站内短信/Email通知内容：</div>"
			.Write "    <textarea name='Setting(80)' cols='60' rows='4'>" & Setting(80) & "</textarea>" 
			.Write "     <span class=""block"">标签说明：{$ContactMan}-卖家名称 {$OrderID}-订单编号 {$TotalMoney}-总货款 {$ServiceCharges}-服务费 {$RealMoney}-实到账</span>"
            .Write "</dd>" &vbcrlf
			.Write "<dd><div>发放优惠券站内短信/Email通知内容：</div>"
			.Write "    <textarea name='Setting(186)' cols='60' rows='4'>" & Setting(186) & "</textarea>" 
			.Write "     <span class=""block"">标签说明：{$UserName}-用户名称 {$CouponNum}-优惠券号码 {$Money}-可抵用金额 {$EndDate}-截止使用时间</span>"
            .Write "</dd>" &vbcrlf
			
			.Write "<dd><div>标签含义：</div"
			.Write " {$OrderID} --定单ID号<br>{$ContactMan} --收货人姓名<br>{$InputTime} --订单提交时间<br>{$OrderInfo} --订单详细信息"
			.Write "</dd>"
			.Write "   </dl>"
			.Write " </div>"							 '========================================================商城系统参数配置结束=========================================
				
			
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

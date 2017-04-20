<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
'-----------------------------------------------------------------------------------------------
'科汛网站管理系统,频道栏目通用类
'开发:厦门科汛软件有限公司 版本 X1.5
'-----------------------------------------------------------------------------------------------
Class ClassCls
        Private KS,KSCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		    '添加频道或目录的过程
			'参数 channelID--频道ID,FolderID 父目录,FormProcesPage--表单处理的页面
			Sub GetAddChannelFolder(Action,FolderID, FormProcesPage)
			 Dim WapSwitch,WapFolderTemplateID,WapTemplateID
			 Dim Folder,CurrPath,TemplateRS, TemplateSql, TypeList, NowDate, YearStr, MonthStr, DayStr,DefaultArrGroupID,ReadPoint,ChargeType,PitchTime,ReadTimes,AllowArrGroupID,DividePercent,K,PubTF,MailTF,FilterTF,MapTF,firstAlphabet,Target
			 Dim ClassBasicInfoArr,FolderName,FolderEname, ClassPic,ClassDescript,MetaKeyWord,MetaDescript,CommentTF,TopFlag,FolderTemplateID,FsoType,FolderFsoIndex,FolderDomain,TemplateID,FnameType,ClassPurview,ClassDefineContentArr,ClassContent
			 Dim TopTitle,SelStr,ClassType,ChannelID,ModelXML,Node,AdminPurview
			 dim ShowADTF,AdParam,AdUrl,AdLinkUrl,AdP,AdType,AdminMobile
			  CurrPath = KS.GetUpFilesDir():If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath, Len(CurrPath) - 1)
			  NowDate = Now():YearStr = Year(NowDate):MonthStr = Right("0"&Month(NowDate),2):DayStr = Right("0" & Day(NowDate),2)
			  If Action="Edit" Then
			    Dim RSE:Set RSE=Server.CreateObject("ADODB.RECORDSET")
				RSE.Open "Select top 1 * From KS_Class Where ID='" & FolderID & "'",Conn,1,1
				If Not RSE.Eof Then
				   FolderID=Rse("TN")
				   ChannelID=RSE("ChannelID")
				  FolderName       = Rse("FolderName")
				  ClassType        = Rse("ClassType")
				  If ClassType=2 Then
				  FolderEname      = Rse("Folder")
				  Else
				  FolderEname      = Split(Rse("Folder"), "/")(Rse("tj") - 1)
				  End If
				  CommentTF        = Rse("CommentTF")
				  PubTF            = KS.ChkClng(Rse("PubTF"))
				  MailTF           = KS.ChkClng(Rse("MailTF"))
				  FilterTF         = KS.ChkClng(Rse("FilterTF"))
				  MapTF            = KS.ChkClng(Rse("MapTF"))
				  TopFlag          = Rse("TopFlag") 
				  WapSwitch        = Rse("WapSwitch")
				  WapFolderTemplateID = Rse("WapFolderTemplateID")
				  WapTemplateID       = Rse("WapTemplateID")
				  FolderTemplateID = Rse("FolderTemplateID") 
				  TemplateID       = Rse("TemplateID")
				  FolderFsoIndex   = Rse("FolderFsoIndex")
				  FnameType        = Rse("FnameType")
				  FsoType          = Rse("FsoType")
				  FolderDomain     = Rse("FolderDomain")
				  ClassPurview     = Rse("ClassPurview")
				  DefaultArrGroupID= Rse("DefaultArrGroupID")
				  AllowArrGroupID  = Rse("AllowArrGroupID")
				  ReadPoint        = Rse("DefaultReadPoint")
				  PitchTime        = Rse("DefaultPitchTime")
				  ReadTimes        = Rse("DefaultReadTimes")
				  ChargeType       = Rse("DefaultChargeType")
				  DividePercent    = Rse("DefaultDividePercent")
				  firstAlphabet    = Rse("firstAlphabet")
				  Target           = Rse("Target")
				  AdminPurview     = Rse("AdminPurview")
				  
				  ClassBasicInfoArr=Split(Rse("ClassBasicInfo")&"||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||","||||")
				  ClassPic=ClassBasicInfoArr(0)
				  ClassDescript=ClassBasicInfoArr(1)
				  MetaKeyWord=ClassBasicInfoArr(2)
				  MetaDescript=ClassBasicInfoArr(3)
				  '画中画
				  AdP=Split(ClassBasicInfoArr(4),"%ks%")
				  ShowADTF=Adp(0)
				  AdParam=Adp(1)
				  AdType=KS.ChkClng(Adp(2))
				  AdUrl=Adp(3)
				  AdLinkUrl=Adp(4)
				  if AdType=0 Then AdType=1
				  if ks.isnul(Rse("ClassDefineContent")) then                  
                   redim ClassDefineContentArr(0)
                  else
                   ClassDefineContentArr=Split(Rse("ClassDefineContent"),"||||")
                  end if
				  AdminMobile=ClassBasicInfoArr(5)
				  ClassContent=ClassDescript
				  TopTitle="编辑"& MyFolderName
				End If
			  Else
			    TopTitle="创建新"& MyFolderName
				ClassType=1 : PubTF=1 : MailTF=0 : FilterTF=1 : MapTF=1
			    CommentTF=1:TopFlag=1:WapSwitch=1:FsoType=11:FolderFsoIndex="index.html":FnameType=".html"
				ClassPurview=0:DefaultArrGroupID=0:AllowArrGroupID=0:ReadPoint=0:DividePercent=0:PitchTime=12:ReadTimes=10:ShowADTF=0:AdParam="250,left,300,300":AdUrl="":AdLinkUrl="http://www.kesion.com":AdType=1:firstAlphabet="0":Target="_blank"
				ChannelID=KS.ChkClng(Request("ChannelID"))
				AdminPurview=KS.C("GroupID")
				If ChannelID=0 Then Channelid=KS.ChkClng(KS.C_C(FolderID,12))
				If ChannelID=0 Then ChannelID=KS.ChkClng(session("class_modelid"))
				If ChannelID=0 Then ChannelID=1

				If FolderID="0" Or FolderID="" or FolderID="1"  Then
				FolderTemplateID="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/频道首页.html"
				WapFolderTemplateID="{@TemplateDir}/3G/" & KS.C_S(ChannelID,10) & "/list.html"
				Else
				FolderTemplateID="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/" & MyFolderName & "页.html"
				WapFolderTemplateID="{@TemplateDir}/3G/" & KS.C_S(ChannelID,10) & "/list.html"
				End If
				TemplateID="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/内容页.html"
				WapTemplateID="{@TemplateDir}/3G/" & KS.C_S(ChannelID,10) & "/show.html"
			  End If
			  TypeList = Replace(KS.LoadClassOption(0,false),"value='" & FolderID & "'","value='" & FolderID &"' selected")
			  
			With Response
                .Write"<!DOCTYPE html>" & vbcrlf
			    .Write "<html>"
			   	.Write "<head>" & vbCrLf
				.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
				.Write "<link href='../Include/admin_style.css' rel='stylesheet'>" & vbCrLf
				.Write "<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>" & vbCrLf
				.Write "<script language='JavaScript' src='../../KS_Inc/Common.js'></script>" & vbCrLf
				.Write "<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
				.Write "<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
				.Write EchoUeditorHead()
				.Write "<script language='Javascript'>" & vbcrlf
				
				If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				.Write "var marr = new Array();" & vbCrlf
				K=0
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks0!=6 and @ks0!=9 and @ks0!=10 || @ks0=5]")
				.Write "marr[" & K & "] = new Array('" & Node.SelectSingleNode("@ks0").text & "','" & Node.SelectSingleNode("@ks1").text & "','" & Node.SelectSingleNode("@ks10").text & "');" & vbCrlf
				K=K+1
				Next
				.Write "</script>" & vbcrlf
				.Write "</head>" & vbCrLf
				.Write "<body>" & vbCrLf
				.Write "<div class=""topdashed sort menu_top_fixed"">" & TopTitle & "</div>" & vbCrLf
				.Write "<div class=""menu_top_fixed_height""></div>"
				.Write " <form  action='" & FormProcesPage & "' method='post' name='CreateFolderForm'>" & vbCrLf
				.Write "<input type='hidden' name='page' value='" & Request("Page") & "'/>"
				.Write "<input type='hidden' name='a' value='" & Request("a") & "'/>"
				.Write "<input type='hidden' name='oids' value='" & Request("oids") & "'/>"
				.Write "<input type='hidden' name='schannelid' value='" & Request("channelid") & "'/>"
				
				.Write "<div class=tab-page id=ClassPane>"
				.Write " <SCRIPT type=text/javascript>"
				.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""ClassPane"" ), 1 )"
				.Write " </SCRIPT>"
				 
				.Write " <div class=tab-page id=site-page>"
				.Write "  <H2 class=tab>基本信息</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
				.Write "	</SCRIPT>"

				'基本信息设置
				.Write "      <dl class=""dtable"">" & vbCrLf
				.Write "            <dd><div>所属" & MyFolderName & "：</div>" & vbCrLf
				If KS.G("Action")="Edit" Then
				.Write "<input type='hidden' name='parentid' value='" & FolderID & "'>"
				.Write "<select name='parentID1' Disabled style='width:300px'>" & vbCrLf
				Else
				.Write "<select onchange='setchannel(this.value)' name='parentID'  style='width:300px'>" & vbCrLf
				End If
				.Write "<option value='0'>无（作为频道)</option>" & vbcrlf
				.Write TypeList & " </select>" & vbcrlf
				.Write "</dd>" & vbCrLf
						 
				.Write "        <dd><div>绑定模型：</div>"
				If Action="Edit" Then
				.Write "  <input type='hidden' id='ChannelID' name='ChannelID' value='" & ChannelID & "'><select style='width:300px' Disabled name='ChannelIDs' class='textbox' onchange='changemodel(this.value)'>" & vbCrLf
				Else
				.Write "   <select name='ChannelID'  style='width:300px' id='ChannelID' class='textbox' onchange='changemodel(this.value)'>" & vbCrLf
				.Write "<option value='0'>---请选择模型---</option>"
				End If
				
				Dim Pstr:Pstr="@ks21=1 and @ks0!=6 "
				If KS.SSetting(0)="0" Then
				Pstr="@ks21=1 and @ks0!=6"
				End If
				If KS.C("SuperTF")<>"1" Then Pstr=Pstr & " and @ks0=" & ChannelID 
				For Each Node In ModelXML.documentElement.SelectNodes("channel[" & Pstr & "]")
				
				  If trim(ChannelID)=trim(Node.SelectSingleNode("@ks0").text) Then
				  .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  Else
				  .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  End If
				Next
				
				             
				.Write "             </select> <span>请选择该" & MyFolderName & "要绑定的模型</span>" & vbCrLf 
				.Write "            </dd>" & vbCrLf	
				
						  
				.Write "          <dd><div>" & MyFolderName & "名称：</div><font id='add1'><INPUT class='textbox' name='FolderName' "
				if KS.S("action")<>"Edit" Then
				 .Write " onkeyup='ctoe()'"
				End If
				.Write " type='text' value='" & FolderName & "' id='FolderName' title='请输入" & MyFolderName & "名称' size=30><font color=red>*</font><span>概括性的说明文字</span> &nbsp;&nbsp;<Select name='firstAlphabet' id='firstAlphabet'><option value='0'>首字母</option>"
				Dim I
				For I=65 To 90
					If firstAlphabet=chr(I) Then
					.Write "<option value='" & I& "' selected>" & chr(I) &"</option>"
					Else
					.Write "<option value='" & I& "'>" & chr(I) &"</option>"
					End If
				Next
			
			.Write "        </select>" & vbCrLf
				
				.Write "</font>"
				.Write "             <div id='add2' style='display:none;color:blue'><strong>录入格式:</strong>" & MyFolderName & "中文名称|英文名称,说明每行一个<br/>"
				.Write "             <textarea id='FolderNames' name='FolderNames' style='width:300px;height:150px'>" & MyFolderName & "名称1|英文名称1</textarea>"
				.Write "             </div>"
				
				If KS.G("Action")<>"Edit" Then
				.Write "<label><input type='checkbox' onclick='ChangeAddMode()' name='AddMore' id='AddMore' value='1'><font color=red><strong>切换到批量添加模式</strong></font></label>"
				End If
				
				.Write "</dd>" & vbCrLf
				.Write "<dd id='typearea'>" & vbCrLf
				.Write "  <div>" & MyFolderName & "类型：" & vbCrLf
				.Write "   <font style=""font-weight:normal"" id=""showtype"">" & vbCrLf
				If Action="Edit" Then
				 .Write "("
				  Select Case ClassType
				   Case "1": .Write "系统"
				   Case "2": .Write "外部链接"
				   Case "3": .Write "单页面"
				  End Select
				  .Write ")"
				Else
				.Write "<label><input type='radio' onclick='changetype(this.value)' name='classtype' value='1'"
				If ClassType="1" Then .Write " checked"
				.Write ">系统</label>"
				.Write "             <label><input type='radio' onclick='changetype(this.value)' name='classtype' value='2'"
				If ClassType="2" Then .Write " checked"
				.Write ">外部链接</label>"
				.Write "             <label><input type='radio' onclick='changetype(this.value)' name='classtype' value='3'"
				If ClassType="3" Then .Write " checked"
				.Write ">单页面</label>"
				End If
				.Write "  </font></div><font id='classarea'>英文名称：</font>" &vbcrlf
				.Write "             <input style='display:none' name='OldFolderEname' class='textbox' type='text' id='OldFolderEname' value='" & FolderEname & "' size=30/>"
				.Write "             <input  class='textbox' name='FolderEname' type='text' id='FolderEname' value='" & FolderEname & "' size=30/>"

				.Write "             <font color=red>*</font><span id='classtips'>不能带\/：*？“ < > | 等特殊符号,且设定后建议不要修改。</span>" & vbCrLf
				.Write "   </dd>" & vbCrLf
				
				.Write "  <span id='templatearea'>"
				.Write "     <dd><div>" & vbCrLf
				
					   If FolderID = "0" Then  .Write ("" & MyFolderName & "首页模板：")  Else  .Write ("" & MyFolderName & "模板：")
					   
				.Write "</div>" & vbCrLf
				.Write " <input class='textbox' type='text' id='FolderTemplateID' name='FolderTemplateID' value='" & FolderTemplateID & "' size=50>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('FolderTemplateID')")
				.Write "  </dd>" & vbCrLf
				
				.Write " <span class='temparea'>" & vbcrlf
					 .Write "     <dd><div>内容页模板：</div>" & vbCrLf
					 .Write "    <input class='textbox' type='text' id='TemplateID' name='TemplateID' value='" & TemplateID & "' size='50'>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('TemplateID')")										              
			   if action<>"Add" then
					.Write "  <label><input checked type='checkbox' value='1' name='autotemplate1'/>自动更换已添加文档的内容页模板</label> "
				End If
				.Write " </dd></span>"
					
				If KS.WSetting(0)="1" Then
				.Write "<span id='3gtemplate'>"
				Else
				.Write "<span id='3gtemplate' style='display:none'>"
				End If
				.Write " <dd><div>手机版" & MyFolderName & "模板：</div>" & vbCrLf
				.Write " <input class='textbox' type='text' id='WAPFolderTemplateID' name='WAPFolderTemplateID' value='" & WAPFolderTemplateID & "' size=50>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('WAPFolderTemplateID')")
				.Write " </dd></span>" & vbCrLf
					
					 .Write " <span class='temparea'><dd><div>手机版内容页模板：</div>" & vbCrLf
					 .Write "<input class='textbox' type='text' id='WAPTemplateID' name='WAPTemplateID' value='" & WAPTemplateID & "' size='50'>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('WAPTemplateID')")
					 
					 if action<>"Add" then									  
					.Write "  <label><input checked type='checkbox' value='1' name='autotemplate2'/>自动更换已添加文档的内容页模板</label>"
					end if
					.Write " </dd></span>"
					
					
					
					If FolderID="0" Then
					.Write "<span id='channel'>"
					Else
					.Write "<span id='channel' style='display:none'>"
					End If
					.Write " <dd><div>绑定域名<font color='#FF0000'>(子域名)</font>：</div>" & vbCrLf
					.Write " <input name='FolderDomain' TYPE='text' value='" & FolderDomain & "' id='FolderDomain' class='textbox' size=50></b>&nbsp;如：http://news.kesion.com/，只对一级" & MyFolderName & "有效 " & vbCrLf
					.Write " </dd>" & vbCrLf
					.Write "</span>"
					
					
                     .Write "<span id=""fsohtmlarea"">"
				     .Write "         <dd><div>生成的" & MyFolderName & "首页文件：</div>" & vbCrLf
					 .Write "       <select name='FolderFsoIndex' class='textbox'>" & vbCrLf
					 .Write "               <option value='index.html'>index.html</option>" & vbCrLf
					 .Write "               <option value='index.htm' selected>index.htm</option>" & vbCrLf
					 .Write "               <option value='index.shtm'>index.shtm</option>" & vbCrLf
					 .Write "               <option value='index.shtml'>index.shtml</option>" & vbCrLf
					 .Write "               <option value='default.html'>default.html</option>" & vbCrLf
					 .Write "               <option value='default.htm'>default.htm</option>" & vbCrLf
					 .Write "               <option value='default.shtm'>default.shtm</option>" & vbCrLf
					 .Write "               <option value='default.shtml'>default.shtml</option>" & vbCrLf
					 .Write "               <option value='index.asp'>index.asp</option>" & vbCrLf
					 .Write "               <option value='default.asp'>default.asp</option>" & vbCrLf
					 .Write "               <option value=""" & FolderFsoIndex & """ selected>" & FolderFsoIndex & "</option>"
					 .Write "             </select>" & vbCrLf
					 .Write "         </dd>" & vbCrLf
					 .Write "        <dd><div>内容页生成的扩展名：</div>"
					 .Write "          <input class='textbox' type='text' ID='FnameType' name='FnameType' value='" & FnameType & "' size='15'> <-<select name='FnameTypes'  class='textbox' onchange=""$('#FnameType').val(this.value);"">" & vbCrLf
					 .Write "               <option value='.html' selected>.html</option>" & vbCrLf
					 .Write "               <option value='.htm'>.htm</option>" & vbCrLf
					 .Write "               <option value='.shtm'>.shtm</option>" & vbCrLf
					 .Write "               <option value='.shtml'>.shtml</option>" & vbCrLf
					 .Write "               <option value='.asp'>.asp</option>" & vbCrLf
					 .Write "             </select>" & vbCrLf
					  .Write "        </dd>" & vbCrLf
					  .Write "        <dd><div>内容页生成HTML格式：</div>" & vbCrLf
					  .Write "   <select style='width:200;' name='FsoType' id='select5' onChange='SelectFsoType(options[selectedIndex].value);'>" & vbCrLf
							   If FsoType = 1 Then SelStr = " Selected"  Else SelStr = ""
							   .Write ("<option value=""1""" & SelStr & ">" & YearStr & "/" & MonthStr & "-" & DayStr & "/RE</option>")
							  If FsoType = 2 Then SelStr = " Selected" Else SelStr = ""
							   .Write ("<option value=""2""" & SelStr & ">" & YearStr & "/" & MonthStr & "/" & DayStr & "/RE</option>")
							  If FsoType = 3 Then SelStr = " Selected" Else SelStr = ""
							   .Write ("<option value=""3""" & SelStr & ">" & YearStr & "-" & MonthStr & "-" & DayStr & "/RE</option>")
							  If FsoType = 4 Then SelStr = " Selected" Else SelStr = ""
							   .Write ("<option value=""4""" & SelStr & ">" & YearStr & "/" & MonthStr & "/RE</option>")
							  If FsoType = 5 Then SelStr = " Selected"  Else	SelStr = ""
							  .Write ("<option value=""5""" & SelStr & ">" & YearStr & "-" & MonthStr & "/RE</option>")
							  If FsoType = 12 Then SelStr = " Selected"  Else	SelStr = ""
							  .Write ("<option value=""12""" & SelStr & ">" & YearStr & MonthStr & "/RE</option>")
							  If FsoType = 6 Then SelStr = " Selected" Else	SelStr = ""
							  .Write ("<option value=""6""" & SelStr & ">" & YearStr & MonthStr & DayStr & "/RE</option>")
							  If FsoType = 7 Then SelStr = " Selected" Else	SelStr = ""
							  .Write ("<option value=""7""" & SelStr & ">" & YearStr & "/RE</option>")
							  If FsoType = 8 Then SelStr = " Selected" Else SelStr = ""
							  .Write ("<option value=""8""" & SelStr & ">" & YearStr & MonthStr & DayStr & "RE</option>")
							  If FsoType = 9 Then SelStr = " Selected" Else SelStr = ""
							  .Write ("<Option value=""9""" & SelStr & ">RE</Option>")
							  If FsoType = 10 Then SelStr = " Selected"  Else SelStr = ""
							  .Write ("<option value=""10""" & SelStr & ">SCE</option>")
							  If FsoType = 11 Then SelStr = " Selected"  Else SelStr = ""
							  .Write ("<option value=""11""" & SelStr & ">文档IDE</option>")

					  .Write "            </select>"
					  .Write "        </dd>" & vbCrLf
					  .Write "        <dd> <div><span id='ShowAS1'></Span></strong> </div></dd>" & vbCrLf
					  .Write "</span>" &vbcrlf
				      .Write "     </span>" &vbcrlf
					  .Write "     </span>" & vbcrlf
					  
				 .Write "         <dd id=""editorarea"">"
				 .Write "           <div>单页内容：<font>(使用标签{$GetClassIntro}在模板里调用)</font></div>" & vbCrLf
				 %>
				<table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td height='30' width=70>&nbsp;<strong>附件上传:</strong></td><td><iframe id='upiframe' name='upiframe' src='../../user/BatchUploadForm.asp?UPFrom=Admin&ChannelID=1' frameborder=0 scrolling=no width='620' height='24'></iframe></td></tr></table>
				 <%
				 Response.Write EchoEditor("ClassContent",ClassContent,"NewsTool","96%","280px")
				 
				 
				 .Write "         </dd>" & vbCrLf
					 .Write "        <dd><div>打开方式：</div>"
					 .Write "          <input class='textbox' type='text' ID='Target' name='Target' value='" & Target & "' size='15'> <-<select name='Targets'  class='textbox' onchange=""$('#Target').val(this.value);"">" & vbCrLf
					 .Write "               <option value='_blank' selected>新窗口(_blank)</option>" & vbCrLf
					 .Write "               <option value='_parent'>父窗口(_parent)</option>" & vbCrLf
					 .Write "               <option value='_self'>本窗口(_self)</option>" & vbCrLf
					 .Write "               <option value='_top'>整页(_top)</option>" & vbCrLf
					 .Write "             </select>" & vbCrLf
					  .Write "        </dd>" & vbCrLf

				 
				 .Write "       </dl>" & vbCrLf
				 .Write "</div>"
				 
				.Write " <div class=tab-page id=classoption-page>"
				.Write "  <H2 class=tab>" & MyFolderName & "选项</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""classoption-page"" ) );"
				.Write "	</SCRIPT>"

				 '频道（栏目）选项
				 .Write " <dl class=""dtable"">" & vbCrLf
				 .Write "          <dd><div>" & MyFolderName & "图片地址：</div>" & vbCrLf
				.Write "         <INPUT NAME='ClassPic' value='" & ClassPic &"' TYPE='text' id='ClassPic' class='textbox' size=61>"
				.Write "                  <input class=""button""  type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,document.CreateFolderForm.ClassPic);"">  <input class=""button"" onClick=""top.openWin('抓取远程图片','include/SaveBeyondfile.asp?fieldid=ClassPic&CurrPath=" & CurrPath & "',false,500,100);"" type='button' name='Submit' value='远程抓取图片...'/>"
				.Write "           <span>用于在" & MyFolderName & "页显示指定的图片</span>" & vbCrLf
				.Write "          </dd>" & vbCrLf
				If ClassType=3 and Action="Edit" then
				 .Write "         <dd style='display:none'>"
                Else
				 .Write "         <dd>"
				End if
				 .Write " <div>" & MyFolderName & "介绍：</div>" & vbCrLf
				 .Write "     <textarea name='ClassDescript' id='ClassDescript' class='textbox' cols='60' rows='3'>" & Server.Htmlencode(ClassDescript) & "</textarea>"
				 .Write "      <span class=""block"">用于在" & MyFolderName & "页详细介绍" & MyFolderName & "信息，支持HTML<br>可在对应的" & MyFolderName & "模板页使用标签<br><font color=red>""{$GetClassIntro}""</font> 进行调用</font></font></span>" & vbCrLf
				 .Write "  </dd>" & vbCrLf		 
							  
				 .Write "  <dd><div>" & MyFolderName & "META关键词：</div>" & vbCrLf
				 .Write "             <textarea name='MetaKeyWord' id='MetaKeyWord' class='textbox' cols='60' rows='4'>" & MetaKeyWord & "</textarea><span class=""block"">用于设置针对搜索引擎的关键词<br>可在对应的" & MyFolderName & "模板页使用标签<br><font color=red>""{$GetClass_Meta_KeyWord}""</font> 进行调用</span>" & vbCrLf
				 .Write "  </dd>" & vbCrLf
				 
				  .Write "         <dd><div>" & MyFolderName & "META网页描述：</div>" & vbCrLf
				 .Write "             &nbsp;<textarea name='MetaDescript' id='MetaDescript' class='textbox' cols='60' rows='4'>" & MetaDescript & "</textarea><span class=""block"">用于设置针对搜索引擎的网页描述<br>可在对应的" & MyFolderName & "模板页使用标签<br><font color=red>""{$GetClass_Meta_Description}""</font> 进行调用</font></span>" & vbCrLf
				 .Write "         </dd>" & vbCrLf
				 
				.Write "          <dd><div>" & MyFolderName & "顶部导航：</div>" & vbCrLf
					       If TopFlag = 1 Then
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""1"">")
						   End If
							.Write ("显示 ")
							If TopFlag = 0 Then
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""0"">")
						   End If
						 .Write "不显示"
					.Write "          </dd>" & vbCrLf
			        .Write "          <dd><div>" & MyFolderName & "3G状态：</div>" & vbCrLf
					       If WapSwitch = 1 Then
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""1"">")
						   End If
							.Write ("显示 ")
							If WapSwitch = 0 Then
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""0"">")
						   End If
						 .Write "不显示"
					.Write "          </dd>" & vbCrLf
			        .Write "          <dd><div>允许邮件订阅：</div>" & vbCrLf
					       If MailTF = 1 Then
						   .Write ("<input name=""MailTF"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""MailTF"" type=""radio"" value=""1"">")
						   End If
							.Write ("允许 ")
							If MailTF = 0 Then
						   .Write ("<input name=""MailTF"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""MailTF"" type=""radio"" value=""0"">")
						   End If
						 .Write "不允许"
					.Write "<span color='blue'>此设置不继承，子" & MyFolderName & "允许订阅需单独设置。</span>"
			        .Write "</dd> <dd><div>允许当筛选项：</div>" & vbCrLf
					       If FilterTF = 1 Then
						   .Write ("<input name=""FilterTF"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""FilterTF"" type=""radio"" value=""1"">")
						   End If
							.Write ("允许 ")
							If FilterTF = 0 Then
						   .Write ("<input name=""FilterTF"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""FilterTF"" type=""radio"" value=""0"">")
						   End If
						 .Write "不允许"
					.Write " <span>指是否在/item/index.asp的" & MyFolderName & "筛选项里显示</span>"
					.Write " </dd>" & vbCrLf
			        .Write "  <dd><div>网站地图里是否显示：</div>" & vbCrLf
					       If MapTF = 1 Then
						   .Write ("<input name=""MapTF"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""MapTF"" type=""radio"" value=""1"">")
						   End If
							.Write ("显示 ")
							If MapTF = 0 Then
						   .Write ("<input name=""MapTF"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""MapTF"" type=""radio"" value=""0"">")
						   End If
						 .Write "不显示"
					.Write " <span>指是否在/plus/map.asp里是否显示该" & MyFolderName & "</span>"
					.Write "</dd>" & vbCrLf
				
				
                 .Write "<font id='ShowAD'>" & vbcrlf
						 .Write "   <dd><div>内容中显示画中画：</div>"
                          if KS.ChkClng(ShowADTF) = "1" Then
						   .Write ("<input onclick=""$('#Ad').show();"" name=""ShowADTF"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input onclick=""$('#Ad').show();"" name=""ShowADTF"" type=""radio"" value=""1"">")
						   End If
							.Write ("显示 ")
						   If KS.ChkClng(ShowADTF) = "0" Then
						   .Write ("<input onclick=""$('#Ad').hide();"" name=""ShowADTF"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input onclick=""$('#Ad').hide();"" name=""ShowADTF"" type=""radio"" value=""0"">")
						   End If
						 .Write "不显示"
						 
					
                         .Write " <table"
						 If KS.ChkClng(ShowADTF)="0" Then .Write "  style=""display:none"""
						 .Write " id=""Ad"" class=""border"" style=""margin:5px"" border=""0"" cellpadding=""5"" cellspacing=""1"">"
                         .Write "<tr class=""tdbg"">"
                         .Write "<td width=""22%""><div align=""right"">画中画参数设置：</div></td>"
                         .Write "<td width=""78%""><input class=""textbox"" name=""AdParam"" type=""text"" id=""AdParam"" size=""20"" maxlength=""20"" value=""" & AdParam & """>(插入位置在内容前多少字,左(left)右(right),宽度,高度：如500,left,300,300)</td>"
                         .Write "</tr>"
						 .Write "<tr class=""tdbg"">"
						 .Write "<td><div align=""right"">广告类型：</div></td>"
						 .Write "<td>"
						 if KS.ChkClng(ADType) = "1" Then
						   .Write ("<input onclick=""$('#adcodearea').hide();$('#adimgarea').show();"" name=""ADType"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input onclick=""$('#adcodearea').hide();$('#adimgarea').show();"" name=""ADType"" type=""radio"" value=""1"">")
						   End If
							.Write ("图片/Flash ")
							If KS.ChkClng(ADType) = "2" Then
						   .Write ("<input onclick=""$('#adimgarea').hide();$('#adcodearea').show();"" name=""ADType"" type=""radio"" value=""2"" checked>")
						   Else
						   .Write ("<input onclick=""$('#adimgarea').hide();$('#adcodearea').show();"" name=""ADType"" type=""radio"" value=""2"">")
						   End If
						 .Write "代码广告（支持Google广告)"
						 .Write "</td>"
						 .Write "</tr>"
						 
						 if KS.ChkClng(ADType)="1" Then
						 .Write "<tbody id='adcodearea' style='display:none'>"
						 Else
						 .Write "<tbody id='adcodearea'>"
						 End IF
                         .Write "<tr class=""tdbg"">"
                         .Write "<td><div align=""right"">广告代码：<br><font color=red>支持HTML语法</font></div></td>"
                         .Write "<td><textarea style='height:60px' name=""AdCode"" class=""textbox"" cols='60' rows=6>" & AdUrl & "</textarea>"
			             .Write "</td></tr>"
						 .Write "</tbody>"
						 
						 if KS.ChkClng(ADType)="2" Then
						 .Write "<tbody id='adimgarea' style='display:none'>"
						 Else
						 .Write "<tbody id='adimgarea'>"
						 End IF
                         .Write "<tr class=""tdbg"">"
                         .Write "<td><div align=""right"">图片地址：</div></td>"
                         .Write "<td><input name=""AdUrl"" class=""textbox"" type=""text"" id=""AdUrl""  size=""36"" maxlength=""250"" value=""" & AdUrl & """>"
                         .Write " <input class=""button""  type='button' name='Submit' value='选择图片或FLASH' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,document.CreateFolderForm.AdUrl);""> "
			             .Write "</td></tr>"
						 .Write "<tr class=""tdbg"">"
						 .Write "<td><div align=""right"">链接地址：</div></td>"
						 .Write "<td><input name=""AdLinkUrl"" type=""text"" class=""textbox"" id=""AdLinkUrl""  size=""36"" maxlength=""250"" value=""" & AdLinkUrl & """>仅对图片有效</td>"
                         .Write "</tr>"
						 .Write "</tbody>"
						 
                         .Write " </table>"
						 
						 .Write "</dd>" & vbCrLf
                  .Write "</span>"

				  .Write "</dl>" & vbCrLf 
				  .Write "</div>"
				
				If ChannelID<>5 Then
				.Write " <div class=tab-page id=poweroption-page>"
				.Write "  <H2 class=tab>权限选项</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""poweroption-page"" ) );"
				.Write "	</SCRIPT>"

				 '权限收费选项设置
				 .Write "<dl class=""dtable"">" & vbCrLf
				.Write "          <dd><div>浏览/查看权限：</div>" & vbCrLf
					 If ClassPurview=0 Then SelStr=" checked" Else SelStr=""
				.Write "<input name='ClassPurview' onclick=""$('#show2').hide()"" type='radio' value='0'" & SelStr &">"
				.Write "              开放" & MyFolderName & "&nbsp;&nbsp;<font color=red>任何人（包括游客）可以浏览和查看此" & MyFolderName & "下的信息。</font><br>"
				If  ChannelID<>8 Then
					 If ClassPurview=1 Then SelStr=" checked" Else SelStr=""
				.Write "<INPUT type='radio'  onclick=""$('#show2').hide()"" name='ClassPurview' value='1'" & SelStr &">" & vbCrLf
				.Write "              半开放" & MyFolderName & "&nbsp;&nbsp;<font color=red>任何人（包括游客）都可以浏览。游客不可查看，其他会员根据会员组的" & MyFolderName & "权限设置决定是否可以查看。</font><br/>"
				End If
					 If ClassPurview=2 Then SelStr=" checked" Else SelStr=""
				.Write " <INPUT type='radio' onclick=""$('#show2').show()"" name='ClassPurview' value='2'" & SelStr &">" & vbCrLf
				If  ChannelID<>8 Then
				.Write "              认证" & MyFolderName & "&nbsp;&nbsp;<font color=red>游客不能浏览和查看，其他会员根据会员组的" & MyFolderName & "权限设置决定是否可以浏览和查看。</font><br>"
				Else
				.Write "              认证" & MyFolderName & "&nbsp;&nbsp;<font color=red>只有指定的会员组才可以查看供求信息的联系方式。</font>"
				End If
				.Write "</dd>" & vbCrLf
				
				If ClassPurview=2 Then
				 .Write " <dd id='show2'>" & vbCrLf
				Else
				.Write " <dd  id='show2' style='display:None'>" & vbCrLf
				End If
				
				If  ChannelID<>8 Then
				.Write "  <div><strong>允许查看此" & MyFolderName & "下信息的会员组：</strong></div><font>如果" & MyFolderName & "是“认证" & MyFolderName & "”，请在此设置允许查看此" & MyFolderName & "下信息的会员组,如果在信息中设置了查看权限，则以信息中的权限设置优先</font>" & vbCrLf
				Else
				.Write " <div>允许查看此" & MyFolderName & "下供求信息联系方式的会员组：</div><font style='text-align:left;'>1、供求内容页模板里一旦放“[KS_Charge][/KS_Chagrge]”标签，则表示联系方式加密；<br/>2、这里勾选的用户组不需要另外付点券，如果联系方式都要扣点才能查看，这里请不要勾选；</font>" & vbCrLf
				End If
				
				.Write KS.GetUserGroup_CheckBox("GroupID",DefaultArrGroupID,5)
				.Write "</dd>" & vbCrLf
				
				.Write "          <dd><div>新用户投稿通知管理员手机号：</div>" & vbCrLf
				.Write "   <input name='AdminMobile' type='text' id='AdminMobile'  value='" & AdminMobile & "' size='50' class='textbox'>不启用此功能，请留空。否则请输入接收短信通知手机号，多个手机号可以用英文逗号隔开。"
				.Write "<span>TIPS:需要在网站的基本信息设置-》短信平台配置好并启用方可以使用本功能。</span>"
				.Write "          </dd>" & vbCrLf
				.Write "          <dd><div>默认与投稿者的分成比率：</div>" & vbCrLf
				.Write "         <input name='DividePercent' type='text' value='" & DividePercent & "' size='6' class='textbox' style='text-align:center'>% 系统将根据这里设置的分成比率将收成分给投稿者。建议设成10的整数倍!"
				.Write "</dd>" & vbCrLf
				
				If  ChannelID<>8 Then
				.Write "          <dd><div>默认阅读信息所需点数：</div>" & vbCrLf
				.Write "   <input name='ReadPoint' type='text' id='ReadPoint'  value='" & ReadPoint & "' size='6' class='textbox' style='text-align:center'>免费阅读请设为""<font color=red>0</font>""，否则有权限的会员阅读该" & MyFolderName & "下的信息时将消耗相应点数，游客将无法阅读。"
				.Write "<span>如果在信息中设置了阅读点数，则以信息中的点数设置优先</span>"
				.Write "          </dd>" & vbCrLf
				.Write "          <dd><div>默认与投稿者的分成比率：</div>" & vbCrLf
				.Write "         <input name='DividePercent' type='text' value='" & DividePercent & "' size='6' class='textbox' style='text-align:center'>% 系统将根据这里设置的分成比率将收成分给投稿者。建议设成10的整数倍!"
				.Write "</dd>" & vbCrLf
				.Write " <dd><div>默认阅读信息重复收费：</div><font>如果在信息中设置了阅读点数，则以信息中的点数设置优先</font><br/>" & vbCrLf
				.Write "<input name='ChargeType' type='radio' value='0' "
					 IF ChargeType=0 Then .Write " checked"
				.Write" >不重复收费(如果信息需扣点数才能查看，建议使用)<br>"
				.Write "&nbsp;<input name='ChargeType' type='radio' value='1'"
					 IF ChargeType=1 Then .Write " checked"
				.write ">距离上次收费时间 <input name='PitchTime' type='text' class='textbox' value='" & PitchTime & "' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='2'"
					 IF ChargeType=2 Then .Write " checked"
				.write ">会员重复阅信息 &nbsp;<input name='ReadTimes' type='text' class='textbox' value='" & ReadTimes & "' size='8' maxlength='8' style='text-align:center'> 页次后重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='3'"
					 IF ChargeType=3 Then .Write " checked"
				.write ">上述两者都满足时重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='4'"
					 IF ChargeType=4 Then .Write " checked"
				.write ">上述两者任一个满足时就重新收费<br>            &nbsp;<input name='ChargeType' type='radio' value='5'"
					 IF ChargeType=5 Then .Write " checked"
				.write ">每阅读一页次就重复收费一次（建议不要使用,多页信息将扣多次点数）"
				.Write ""
				.Write "          </dd>" & vbCrLf
                  End If
				  
				  .Write "<dd><div>允许管理此" & MyFolderName & "的管理员角色：</div>" & vbCrLf
				
				.Write KS.GetRoleList("AdminPurview",AdminPurview)
				.Write "</dd>" & vbCrLf
				  
				  
				 .Write "       </dl>" & vbCrLf 
				 .Write "</div>" & vbcrlf
				 End If
				 
				.Write " <div class=tab-page id=classtg-page>"
				.Write "  <H2 class=tab>投稿选项</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""classtg-page"" ) );"
				.Write "	</SCRIPT>"

				.Write "<dl class=""dtable"">" & vbCrLf
				.Write "    <dd><div>允许在本" & MyFolderName & "下发布文档：</div>" & vbCrLf
				
						if PubTF = 1 Then
						   .Write ("<input name=""PubTF"" type=""radio"" value=""1"" checked>允许")
						Else
						   .Write ("<input name=""PubTF"" type=""radio"" value=""1"">允许")
						End If
				
						if PubTF = 0 Then
						   .Write ("<input name=""PubTF"" type=""radio"" value=""0"" checked>不允许")
						Else
						   .Write ("<input name=""PubTF"" type=""radio"" value=""0"">不允许")
						End If
				
				.Write "  <span>当" & MyFolderName & "不是终级" & MyFolderName & "时,建议选择不允许</span>"
				.Write " </dd>"
				
				.Write "  <dd><div>是否允许在本" & MyFolderName & "下投稿：</div>" & vbCrLf
						If CommentTF = 0 Then
						   .Write ("①<input name=""CommentTF"" onclick=""$('#sag').hide()"" type=""radio"" value=""0"" checked>不允许<br>")
						Else
						   .Write ("①<input name=""CommentTF"" onclick=""$('#sag').hide()"" type=""radio"" value=""0"">不允许<br>")
						End If
						if CommentTF = 1 Then
						   .Write ("②<input name=""CommentTF"" onclick=""$('#sag').hide()""  type=""radio"" value=""1"" checked>允许所有会员投稿<font color=blue>(游客除外)</font><br>")
						Else
						   .Write ("②<input name=""CommentTF"" onclick=""$('#sag').hide()"" type=""radio"" value=""1"">允许所有会员投稿<font color=blue>(游客除外)</font><br>")
						End If
						if CommentTF = 2 Then
						   .Write ("③<input name=""CommentTF"" onclick=""$('#sag').hide()"" type=""radio"" value=""2"" checked>允许所有人投稿<font color=red>(包括游客)</font><br>")
						Else
						   .Write ("③<input name=""CommentTF"" onclick=""$('#sag').hide()"" type=""radio"" value=""2"">允许所有人投稿<font color=red>(包括游客)</font><br>")
						End If
						if CommentTF = 3 Then
						   .Write ("④<input name=""CommentTF"" onclick=""$('#sag').show()"" type=""radio"" value=""3"" checked>只允许指定用户组的会员投稿<br>")
						Else
						   .Write ("④<input name=""CommentTF"" onclick=""$('#sag').show()"" type=""radio"" value=""3"">只允许指定用户组的会员投稿<br>")
						End If

					
				.Write "   </dd>"
				if CommentTF = 3 Then
				.Write "   <dd id=""sag"">"
				Else
				.Write "   <dd id=""sag"" style=""display:none"">"
				End If
				.Write "<div>允许此" & MyFolderName & "下投稿的会员组：</div>" & vbCrLf
				.Write "<font>当上面选择④时，请在此设置允许在此" & MyFolderName & "下投稿的会员组</font>" & KS.GetUserGroup_CheckBox("AllowArrGroupID",AllowArrGroupID,5)
				.Write "   </dd>" & vbCrLf
				.Write "</dl>"
				.Write "</div>" 

				 
				.Write " <div class=tab-page id=defineoption-page>"
				.Write "  <H2 class=tab>自设选项</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""defineoption-page"" ) );"
				.Write "	</SCRIPT>"

				 .Write " <dl class=""dtable"">" & vbCrLf
				  .Write "         <dd><div>自设内容数：</div>" & vbCrLf
				 .Write " <select name=""ClassDefine_Num"" onChange=""setFileFileds(this.value)"">"
				  Dim DefineNum,SelDefineNum
				  If IsArray(ClassDefineContentArr) Then SelDefineNum=Ubound(ClassDefineContentArr)+1 Else SelDefineNum=1
				  For DefineNum=1 To 20
				   If DefineNum=SelDefineNum Then
				    .Write "<option value=""" & DefineNum & """ selected>" & DefineNum & "</option>"
				   Else
				    .Write "<option value=""" & DefineNum & """>" & DefineNum & "</option>"
				   End If
                  Next
			     .Write " </select>"
				 .Write "  </dd>" & vbCrLf
				
				 For DefineNum=1 To 20
				 .Write "        <dd id='objFiles" & DefineNum & "' style=""display:none"">"
				 .Write "         <div>自设内容" & DefineNum & "：</div> " & vbCrLf
				 
				  If Action="Edit" Then
				     IF DefineNum-1<=Ubound(ClassDefineContentArr) Then
				      .Write "   <TEXTAREA class='textbox' Name='ClassDefineContent" & DefineNum &"' ROWS='' COLS=''style='width:500px;height:80px'>" &ClassDefineContentArr(DefineNum-1)& "</TEXTAREA> " & vbCrLf
					 Else
					  .Write "  <TEXTAREA class='textbox' Name='ClassDefineContent" & DefineNum &"' ROWS='' COLS=''style='width:500px;height:80px'></TEXTAREA> " & vbCrLf
					 End If
				  Else
				    .Write "    <TEXTAREA class='textbox' Name='ClassDefineContent" & DefineNum &"' ROWS='' COLS=''style='width:500px;height:80px'></TEXTAREA> " & vbCrLf
				  End If
				 .Write "   <span>在" & MyFolderName & "模板页插入{$GetClassDefineContent" & DefineNum & "} 调用</span>" & vbCrLf
				 .Write "         </dd>" & vbCrLf
				 Next 
				 .Write "       </dl>" & vbCrLf 
				 .Write "</div>"
				 .Write "</div>"
				 .Write " </form>" & vbCrLf
				.Write "</body>" & vbCrLf
				.Write "</html>" & vbCrLf
				.Write "<Script Language='javascript'>" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "$(document).ready(function(){" & vbcrlf
				.Write " SelectFsoType('11')" & vbcrlf
				If Action="Edit" Then .Write "showad('" & ChannelID & "');" & vbcrlf
				.Write "})"& vbcrlf
				
				.Write "changetype('" & ClassType &"');" & vbcrlf
				.Write "function ChangeAddMode(){" & vbcrlf
				.Write " if ($('#AddMore').prop('checked')==true){"
				.Write "  $('#add1').hide(); $('#add2').show(); $('#typearea').hide();"
				.Write " }else{"
				.Write "  $('#add1').show();$('#add2').hide();$('#typearea').show();"
				.Write " }"
				.Write "}" & vbcrlf
				.Write "function SelectFsoType(ObjValue)" & vbCrLf
				.Write "{ var ChannelDomain='" & KS.GetChannelDomain(ChannelID) & KS.C_S(ChannelID,43) &"';" & vbCrLf
				 
					Dim N
					Randomize
					N = Rnd * 3 + 5
				.Write "switch (ObjValue)" & vbCrLf
				.Write "  {" & vbCrLf
				.Write "   case '1' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & MonthStr & "-" & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '2' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & MonthStr & "/" & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '3' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "-" & MonthStr & "-" & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '4' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & MonthStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '5' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "-" & MonthStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '12' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & MonthStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '6' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & MonthStr & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '7' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '8' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & MonthStr & DayStr & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '9' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & KS.MakeRandom(N) & "'+ $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '10' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & KS.MakeRandomChar(N) & "'+ $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '11' :$('#ShowAS1').html(ChannelDomain+'<font color=red>文档ID'+ $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "  }"
				.Write "}" & vbCrLf
				.Write "function changemodel(mid){" &vbcrlf
				.Write " if (mid==9 || mid==10 || mid==11 || mid==12 || mid==13){" & vbcrlf
				.Write "changetype(2);" & vbcrlf
				.Write "$('input[name=""classtype""]').each(function(){ if ($(this).val()==2){$(this)[0].checked=true; }  });" &vbcrlf
				.Write "$('#showtype').hide();" &vbcrlf
				.Write "}else{" &vbcrlf
				.Write "$('input[name=""classtype""]').each(function(){ if ($(this).val()==1){$(this)[0].checked=true; }  });" &vbcrlf
				.Write "changetype(1);" & vbcrlf
				.Write "$('#showtype').show();" &vbcrlf
				.Write "}" &vbcrlf
				.Write "" & vbcrlf
				.Write " for(i=0;i<marr.length;i++){" & vbcrlf
				.Write "  if (mid==marr[i][0]){$('#FolderTemplateID').val('{@TemplateDir}/'+marr[i][1]+'/" & MyFolderName & "页.html');$('#TemplateID').val('{@TemplateDir}/'+marr[i][1]+'/内容页.html');$('#WAPFolderTemplateID').val('{@TemplateDir}/3G/'+marr[i][2]+'/list.html');$('#WAPTemplateID').val('{@TemplateDir}/3G/'+marr[i][2]+'/show.html');}"
				.Write "  } return; showad(mid);" & vbcrlf
				.Write "}" & vbcrlf
			
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.CreateFolderForm;" & vbCrLf
				.Write "   if ($('#ChannelID').val()==0)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     alert('请选择此" & MyFolderName & "要绑定的模型!');" & vbCrLf
				.Write "     $('#ChannelID').focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "    }" & vbCrLf
				.Write "   if ($('input[name=FolderName]').val()=='' && $('#AddMore').prop('checked')==false)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     top.$.dialog.alert('请输入" & MyFolderName & "的中文名称!',function(){" & vbCrLf
				.Write "     $('input[name=FolderName]').focus();});" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "    }" & vbCrLf
				.Write "    if ($('#FolderName').val().length>50)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     top.$.dialog.alert('" & MyFolderName & "中文名称不能超过25个汉字(50个英文字符)!',function(){" & vbCrLf
				.Write "     $('#FolderName').focus();});" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "    if ($('#FolderEname').val()==''&& $('#AddMore').prop('checked')==false)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     top.$.dialog.alert('请输入" & MyFolderName & "的英文名称!',function(){" & vbCrLf
				.Write "     $('#FolderEname').focus();});" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "    }" & vbCrLf
				If Action<>"Edit" Then
				.Write "    if (form.classtype[0].checked && CheckEnglishStr(form.FolderEname,'" & MyFolderName & "的英文名称')==false)" & vbCrLf
				.Write "     return false;" & vbCrLf
				End If
				.Write "    if ($('#FolderTemplateID').val()=='')" & vbcrlf
				.Write "     { top.$.dialog.alert('请绑定" & MyFolderName & "模板!',function(){" & vbcrlf
				.Write "       $('#FolderTemplateID').focus();});"
				.Write "       return false;}" & vbcrlf 
				.Write "    if ($('#TemplateID').val()=='')" & vbcrlf
				.Write "     { top.$.dialog.alert('请绑定内容页页模板!',function(){" & vbcrlf
				.Write "       $('#TemplateID').focus();});"
				.Write "       return false;}" & vbcrlf 
				.Write "    form.submit();" & vbCrLf
				.Write "}"
				.Write "function ctoe()" & vbCrLf
				.Write "{" & vbCrLf
				.Write " var folderName=escape($('input[name=FolderName]').val());" & vbcrlf
				.Write "$.get('../../plus/ajaxs.asp', { foldername: folderName, action: 'Ctoe' }," &vbCrlf
				.Write "	function(data){ " & vbcrlf
				.Write "    if ($(""input[name='classtype']:checked"").val()==3){"& vbcrlf
				.Write "	  $('input[name=FolderEname]').val(unescape(data)+'.html');" & vbcrlf
				.Write "    }else{" &vbcrlf
				.Write "	  $('input[name=FolderEname]').val(unescape(data));" & vbcrlf
				.Write "    }" &vbcrlf
				.Write "  });"
				.Write "}" & vbCrLf
				.Write "setFileFileds($('select[name=ClassDefine_Num]').val());" & vbcrlf
				.Write "function setFileFileds(num){    " &vbcrlf
				.Write "for(var i=1,str="""";i<=20;i++){" & vbcrlf
				.Write "	$(""#objFiles"" + i).hide();" & vbcrlf
				.Write "}" & vbcrlf
				.Write "for(var i=1,str="""";i<=num;i++){"
				.Write "	$(""#objFiles"" + i).show();" & vbcrlf
				.Write "}" & vbcrlf
			    .Write "}" & vbcrlf
				.Write "function setchannel(v)" & vbcrlf
				.Write "{ if (v=='0') {$('#channel').show();} else {$('#channel').hide()}}"
				.Write "function changetype(v)" & vbcrlf
				.Write "{"
				.Write " switch(parseInt(v))"&vbcrlf
				.Write "  {case 1:$('#3gtemplate').show();$('#editorarea').hide();$('#fsohtmlarea').show();$('#classarea').html('英文名称：');$('#classtips').html('不能带\/：*？“ < > | 等特殊符号,且设定后不能改');$('#templatearea').show();$('.temparea').show();if ($('#FolderEname').val().indexOf('.')!=-1){$('#FolderEname').val($('#FolderEname').val().split('.')[0])} break;" & vbcrlf
				.Write "   case 2:$('#3gtemplate').hide();$('#editorarea').hide();$('#fsohtmlarea').hide();$('#classarea').html('链接地址：');$('#classtips').html('如 <font color=blue>http://www.kesion.com</font> 等');$('#channel').hide();$('#templatearea').hide();$('.temparea').hide();break;" & vbcrlf
				.Write "   case 3:$('#3gtemplate').show();$('#editorarea').show();$('#fsohtmlarea').hide();$('#classarea').html('生成文件名：');$('#classtips').html('如 <font color=blue>about.html,intro.html,help.html</font>等');$('#templatearea').show();$('.temparea').hide();$('#channel').hide();"
				.Write "if ($('#FolderEname').val()!='' && $('#FolderEname').val().indexOf('.')==-1){ "
				.Write "  $('#FolderEname').val($('#FolderEname').val()+"".html"");" &vbcrlf
				.Write "  }"
				.Write "break;" & vbcrlf
				
				.Write " }"
				If channelid>0 and (KS.WSetting(0)="0" or KS.ChkClng(KS.C_S(Channelid,53))=0) THEN .Write "$('#3gtemplate').hide();"
				.Write " }"&vbcrlf
				.Write "function showad(v){" & vbcrlf
				.Write " $('#ShowAD').show();" & vbcrlf
				.Write "}" & vbcrlf
				.Write "function insertHTMLToEditor(codeStr) {" & InsertEditor("ClassContent","codeStr") &"} " &vbcrlf
				.Write "//-->"
				.Write "</Script>"
				
				
				
			End With
			End Sub
			
			
			'添加频道目录的保存过程
			'参数:ChannelID--频道ID
			Sub ChannelFolderAddSave(Go)
			Dim ID, TJ, FolderName, Folder,ChannelID, ClassID, TS, FolderTemplateID, FolderFsoIndex
			Dim TemplateID, FnameType, FsoType, FolderDomain, FolderOrder, CurrPath, TopFlag,ClassType,WapSwitch,WapFolderTemplateID,WapTemplateID,PubTF,MailTF,FilterTF,MapTF
			Dim RSC,FolderEName,CommentTF,ClassPurview,GroupID,ReadPoint,ChargeType,DividePercent,PitchTime,ReadTimes,AllowArrGroupID,AddMore,ParentFolder,j,Root,Child,PrevOrderID
			Dim ClassPic,ClassDescript,MetaKeyWord,MetaDescript,ClassDefine_Num,N,ClassDefineContent,Action,AdminMobile
				
				Action=KS.G("Action")
				AddMore=Request.Form("AddMore")
				
				If AddMore="1" Then
				 FolderName=Request.Form("FolderNames")
				 ClassType=1
				 If Trim(FolderName) = "" Then Call KS.AlertHistory("批量添加" & MyFolderName & ",请按格式输入" & MyFolderName & "名称及" & MyFolderName & "英文名称!",-1):.End
				 FolderName=Split(FolderName,vbcrlf)
				Else
				 FolderName = KS.G("FolderName")
				 ClassType  = KS.ChkClng(KS.G("ClassType"))
				 FolderEName = Replace(KS.G("FolderEName")," ","")
				 If Trim(FolderName) = "" Then Call KS.AlertHistory("目录中文名称不能为空!",-1):.End
				 If KS.strLength(Trim(FolderName)) > 50 Then Call KS.AlertHistory("目录中文名称不能超过25个汉字(50个英文字符)!", -1): .End 
				 If Trim(FolderEName) = "" Then Call KS.AlertHistory("目录英文名称不能为空!",-1):.End
				End If
				
				if ClassType=1 Then
				 If Instr(FolderEName,".") <>0 Then Call KS.AlertHistory("目录英文名称不能含有“.”!",-1):.End
				Elseif ClassType=3 Then
				 If right(lcase(FolderEName),4) <>".htm" and right(lcase(FolderEName),5)<>".html" and right(lcase(FolderEName),6)<>".shtml" and right(lcase(FolderEName),5)<>".shtm" Then Call KS.AlertHistory("单页面扩展名不正确，只能是.html,.htm,.shtm,.shtml中的一个!",-1):.End
				End If
				
				ID                 = Trim(Request("parentID")):If ID = "" Then ID = "0"
				FolderTemplateID   = KS.G("FolderTemplateID")
				TemplateID         = KS.G("TemplateID")
				WapFolderTemplateID= KS.G("WapFolderTemplateID")
				WapTemplateID      = KS.G("WapTemplateID")
				ChannelID          = KS.ChkClng(KS.G("ChannelID"))
				MailTF             = KS.ChkClng(KS.G("MailTF"))
				FilterTF           = KS.ChkClng(KS.G("FilterTF"))
				MapTF              = KS.ChkClng(KS.G("MapTF"))
				
			
				If FolderTemplateID = "" Or TemplateID = "" Then Call KS.AlertHistory("对不起,添加新频道应先选择模板绑定!", -1): Exit Sub
				If ClassType=3 Then
				 	If Instr(FolderEName,".")=0 Then
						Call KS.AlertHistory("单页面保存的文件格式不正确!", -1)
						Set KS = Nothing:Response.End
					 Else
					   Dim FileExt:FileExt=lcase(Split(FolderEName,".")(1))
					   If FileExt<>"html" and FileExt<>"htm" and FileExt<>"shtml" and FileExt<>"shtm" Then
						Call KS.AlertHistory("单页面保存的文件格式不正确,只能以html,htm,shtml或shtm为扩展名!", -1)
						Set KS = Nothing:Response.End
					   End If
					 End If
				End If
				
			   If ID <> "0" And ID<>"" Then  
				     Dim FolderRS,MaxOrderID
					 Set FolderRS = Server.CreateObject("ADODB.RECORDSET")
					 FolderRS.Open"Select Folder,FolderName,FolderDomain,TS,Tj,Root,FolderOrder,Child From KS_Class Where ID='" & ID & "'",conn,1,1
					 If FolderRS.EOF Then
					    FolderRS.Close:Set FolderRS=Nothing
						KS.AlertHintScript "父" & MyFolderName & "不存在！"
					 Else
					    Root=FolderRS("Root")
						PrevOrderID=FolderRS("FolderOrder")
						Child=FolderRS("Child")
						TS = Trim(FolderRS("TS"))

						if (Child > 0) Then
							'得到与本栏目同级的最后一个栏目的OrderID
							PrevOrderID = Conn.Execute("select Max(FolderOrder) From KS_Class where tn='" &ID& "'")(0)
	
							'得到同一父栏目但比本栏目级数大的子栏目的最大OrderID，如果比前一个值大，则改用这个值
							MaxOrderID =  KS.ChkClng(Conn.Execute("select Max(FolderOrder) from [KS_Class] where ts like '" & ts & "%'")(0))
							if (MaxOrderID > PrevOrderID) Then	PrevOrderID = MaxOrderID
                        end if
						
					    ParentFolder=Trim(FolderRS("Folder"))
						Folder = ParentFolder & FolderEName
						FolderDomain = Trim(FolderRS("FolderDomain"))
						TJ = FolderRS("TJ")+1
					    
					 End If
					 FolderRS.Close:Set FolderRS = Nothing
			   Else 
					Folder = FolderEName
					TJ=1
					FolderDomain = KS.G("FolderDomain")
					Root=Conn.Execute("Select Max(root) From KS_Class")(0)
					If KS.IsNul(Root) Then 
					 Root=1
					Else
					 Root=Root+1
					End If
					
			   End If
			   
			   If ClassType=1 Then Folder=trim(Folder) & "/"
				
				If Action="Add" Then
					Set RSC=Server.CreateObject("ADODB.Recordset")
					RSC.Open "Select FolderName,Folder From KS_Class Where ChannelID=" & ChannelID & " and TN='" & ID & "'", Conn, 1, 1
					If Not RSC.EOF Then
					  If AddMore="1" Then
					      '检查输入的是否有同名
						  For I=0 To Ubound(FolderName)
						   For J=0 To Ubound(FolderName)
							   If Ubound(split(FolderName(j),"|"))<1 Then
								Call KS.AlertHistory("批量输入的" & MyFolderName & "格式不正确!请按＂" & MyFolderName & "中文名称|" & MyFolderName & "英文名称＂和格式录入!", -1):.End
							   End If
							   If Not IsAlphabet(replace(Split(FolderName(i),"|")(1)," ","")) Then
								Call KS.AlertHistory("批量输入的" & MyFolderName & "英文名称不正确!请输英文名称!", -1):.End
							   End If
							   
						       If Split(FolderName(i),"|")(0)=Split(FolderName(j),"|")(0) and i<>j Then
							     Call KS.AlertHistory("批量输入的" & MyFolderName & "[" & Split(FolderName(i),"|")(0) & "]存在重复!", -1):.End
							   End If
						       If trim(Split(FolderName(i),"|")(1))=trim(Split(FolderName(j),"|")(1)) and i<>j Then
							    Call KS.AlertHistory("批量输入的英文" & MyFolderName & "[" & Split(FolderName(i),"|")(1) & "]存在重复!", -1):.End
							   End If
						   Next
						  Next
						  
						  Do While Not RSC.Eof
						   For I=0 To Ubound(FolderName)
						    If RSC(0) = Split(FolderName(i),"|")(0) Then  Call KS.AlertHistory("批量输入的" & MyFolderName & "[" & Split(FolderName(i),"|")(0) & "]已存在,请用其它名称!", -1):.End
						    If RSC(1) = Split(FolderName(i),"|")(1) Then Call KS.AlertHistory("批量输入的英文名称[" & Split(FolderName(i),"|")(1) & "]已存在,请用其它英文名称!",-1): .End
						   Next
						   RSC.MoveNext
						  Loop
					  Else
						  Do While Not RSC.Eof
						   If RSC(0) = FolderName Then  Call KS.AlertHistory("名称已存在,请用其它名称!", -1):.End
						   If RSC(1) = Folder Then Call KS.AlertHistory("英文名称已存在,请用其它英文名称!",-1): .End
						   RSC.MoveNext
						  Loop
					  End If
					End If
					RSC.Close:Set RSC=Nothing
				End If
				

			   TopFlag = KS.ChkClng(KS.G("TopFlag"))
			   PubTF   = KS.ChkClng(KS.G("PubTF"))
			   WapSwitch  = KS.ChkClng(KS.G("WapSwitch"))
			   FolderFsoIndex = Request("FolderFsoIndex")
			   FnameType = Request("FnameType")
			   FsoType = Request("FsoType")
			   ClassPurview= KS.ChkClng(KS.G("ClassPurview"))
			
				CommentTF=Request.Form("CommentTF")
				GroupID=KS.G("GroupID"):if GroupID="" Then GroupID=0
				AllowArrGroupID=KS.G("AllowArrGroupID"):iF AllowArrGroupID="" Then AllowArrGroupID=0
				ClassPic=Request.Form("ClassPic")
				ClassDescript=Request.Form("ClassDescript")
				If ClassDescript="" and ClassType=3 Then ClassDescript=Request.Form("ClassContent")
				
				
				MetaKeyWord=Request.Form("MetaKeyWord")
				MetaDescript=Request.Form("MetaDescript")
				ClassDefine_Num=KS.ChkClng(KS.G("ClassDefine_Num"))
				For N=1 To ClassDefine_Num
				  If N=1 Then
				   ClassDefineContent=Request.Form("ClassDefineContent"& N)
				  Else
				   ClassDefineContent=ClassDefineContent & "||||" & Request.Form("ClassDefineContent"& N)
				  End If
				Next
				
				ReadPoint=KS.ChkClng(KS.G("ReadPoint"))
				ChargeType=KS.ChkClng(KS.G("ChargeType"))
				PitchTime=KS.ChkClng(KS.G("PitchTime"))
				ReadTimes=KS.ChkClng(KS.G("ReadTimes"))
				DividePercent=KS.G("DividePercent")
				AdminMobile=KS.G("AdminMobile")
				If Not IsNumeric(DividePercent) Then
				 DividePercent=0
				End If
				Dim AdParam,AdPa
				AdPa="0%ks%0,0,0,0%ks%0%ks%%ks%"
				'If KS.C_S(ChannelID,6)=1  Then
					AdParam=KS.G("AdParam")
					if Ubound(Split(AdParam,","))<>3 Then Call KS.AlertHistory("输入的画中画广告参数设置有误!",-1).end
					if KS.ChkClng(KS.G("ShowADTF"))=1 and KS.G("ADtype")="1" and KS.G("AdUrl")="" then Call KS.AlertHistory("输入的画中画广告的图片地址!",-1).end
					if KS.ChkClng(KS.G("ShowADTF"))=1 and KS.G("ADtype")="2" and KS.G("AdCode")="" then Call KS.AlertHistory("输入的画中画广告的代码!",-1).end
					If KS.G("ADtype")="2" then
					AdPa=KS.ChkClng(KS.G("ShowADTF")) & "%ks%" & AdParam &"%ks%" & KS.G("ADType") & "%ks%" & Request.Form("AdCode") & "%ks%"
					else
					AdPa=KS.ChkClng(KS.G("ShowADTF")) & "%ks%" & AdParam &"%ks%" & KS.G("ADType") & "%ks%"& KS.G("AdUrl") & "%ks%" & KS.G("AdLinkUrl")
					end if
				'End If
				
				Dim Node,oldnode,m,OldTemplateID,OldWapTemplateID

				Dim RST:Set RST=Server.CreateObject("ADODB.Recordset")
				If Action="Add" Then
				     If Not IsArray(FolderName) Then FolderName=Split(FolderName,vbcrlf)
				     For I=Ubound(FolderName) To Lbound(FolderName) Step -1
						RST.Open "select top 1 * from KS_Class where 1=0", Conn, 1, 3
						RST.AddNew
						ClassID = KS.GetClassID()   '调用函数取新的目录ID
						RST("ID") = ClassID
						RST("Creater") = KS.C("AdminName")
						RST("AdminPurView")=KS.S("AdminPurview")
						RST("CreateDate") = Now
						If AddMore="1" Then
							if ID<>"" Then
							 RST("folder") = ParentFolder & trim(Split(FolderName(i),"|")(1)) & "/"
							Else
							 RST("Folder")=trim(Split(FolderName(i),"|")(1)) & "/"
							End If
						Else
						    if ClassType=2 Then
							 RST("folder") = FolderEname
							Else
							 RST("Folder")=Folder
							End If
						End If
						RST("FolderName") = Split(FolderName(i),"|")(0)
						RST("ClassType")=ClassType
						If ID <> "" Then  RST("TN") = ID Else  RST("TN") = "0"  
						RST("TJ") = TJ
						RST("TS") = "" & TS & "" & ClassID & ","
						RST("FolderTemplateID") = FolderTemplateID
						RST("TopFlag")   = TopFlag
						RST("PubTF")     = PubTF
						RST("MailTF")    = MailTF
						RST("FilterTF")  = FilterTF
						RST("MapTF")     = MapTF
						RST("firstAlphabet")= Chr(KS.G("firstAlphabet"))
						RST("Target")    = KS.G("Target")
						RST("WapSwitch") = WapSwitch
						
						RST("FolderFsoIndex") = FolderFsoIndex
						RST("TemplateID") = TemplateID
						If KS.WSetting(0)="1" Then
						RST("WapFolderTemplateID")=WapFolderTemplateID
						RST("WapTemplateID")=WapTemplateID
						end if
						RST("FnameType") = FnameType
						RST("FsoType") = FsoType
						RST("FolderDomain") = FolderDomain
						If RST("TN") <> "0" Then
						RST("FolderOrder") = PrevOrderID+I
						Else
						RST("FolderOrder") = 0
						End If
						If ID="" Or ID="0" Then
						 RST("Root")=Root+i
						Else
						 RST("Root")=Root
						End If
						RST("Child")=0
						RST("ChannelID") = ChannelID
						RST("DelTF") = 0
						RST("ClassPurview")=ClassPurview
						RST("CommentTF")=CommentTF
						RST("DefaultArrGroupID")=KS.FilterIDS(GroupID)
						RST("AllowArrGroupID")=KS.FilterIDS(AllowArrGroupID)
						RST("DefaultReadPoint")=ReadPoint
						RST("DefaultChargeType")=ChargeType
						RST("DefaultDividePercent")=DividePercent
						RST("DefaultPitchTime")=PitchTime
						RST("DefaultReadTimes")=ReadTimes
						RST("ClassBasicInfo")=ClassPic & "||||" & ClassDescript & "||||" & MetaKeyWord   &"||||" & MetaDescript & "||||" & AdPa &"||||" & AdminMobile
						RST("ClassDefineContent")=ClassDefineContent
						RST.Update
						
						Call KS.FileAssociation(1000,RST("ClassID"),RST("ClassBasicInfo")&ClassDefineContent,0)
						
						if (ID <>"" and id<>"0") Then
                           Conn.Execute ("update ks_class set Child=Child+1 where id='" & ID & "'")
                           '更新该栏目排序以及大于本需要和同在本分类下的栏目排序序号
						   Conn.Execute ("update ks_class set FolderOrder=FolderOrder+1 where root=" & Root & " and FolderOrder>" & KS.ChkClng(PrevOrderID))
						   Conn.Execute ("update ks_Class set FolderOrder=" & PrevOrderID & "+1 where ID='" & RST("ID") & "'")
                       End If
					   
					   
					   	'采用向内存追加节点
						
						If ClassType=3  And KS.C_S(ChannelID,7)=1 Then  '单页面生成
						  Dim KSR:Set KSR=New Refresh
						  Call KSR.RefreshFolder(ChannelID,RST)  '调用栏目刷新函数
						  Set KSR=Nothing
						End If
							
    					RST.Close
				     Next
					 
                         Application(KS.SiteSN&"_class")=""
                         Application(KS.SiteSN&"_classpath")=""
						 
						 
						
						 
						 KSCls.ClassAction  channelid          '生成搜索JS
						If session("class_from")="all" Then
						Call KS.Confirm("创建成功,继续创建吗?",KS.Setting(3) & KS.Setting(89) & "system/KS.Class.asp?Action=" & Action &"&oids=" & request("oids")&"&Go=" & Go & "&FolderID=" & ID & "",KS.Setting(3) & KS.Setting(89) & "system/KS.Class.asp?oids=" & request("oids"))
						Else
						Call KS.Confirm("创建成功,继续创建吗?",KS.Setting(3) & KS.Setting(89) & "system/KS.Class.asp?ChannelID=" & ChannelID &"&oids=" & request("oids")&"&Action=" & Action &"&Go=" & Go & "&FolderID=" & ID & "",KS.Setting(3) & KS.Setting(89) & "system/KS.Class.asp?ChannelID=" & ChannelID & "&oids=" & request("oids"))
						End If
					Else
					    RST.Open "select top 1 * from KS_Class Where ID='" &KS.G("FolderID") & "'", Conn, 1, 1
					  
						if request("FolderEname")<>request("oldFolderEname")  and RST("classtype")=1 then
							Dim An_folder:An_folder=RST("Folder")
							An_folder=left(An_folder,len(An_folder)-len(request("oldFolderEname"))-1)&request("FolderEname")&"/"
							 IF databaseType=1 Then
							  IF Not(Conn.Execute("Select folder from KS_Class Where TN in(Select TN from KS_Class Where id='"&KS.G("FolderID")&"') and  id<>'"&KS.G("FolderID")&"' and Convert(nvarchar(2000),Folder)='"&An_folder&"'").Eof) Then
								Call KS.AlertHistory("栏目英文名已存在！",-1): .End
							  End IF
							 Else
							  IF Not(Conn.Execute("Select folder from KS_Class Where TN in(Select TN from KS_Class Where id='"&KS.G("FolderID")&"') and  id<>'"&KS.G("FolderID")&"' and Folder='"&An_folder&"'").Eof) Then
								Call KS.AlertHistory("栏目英文名已存在！",-1): .End
							  End IF
							 End If
						End If
					    RST.Close
					
						RST.Open "select top 1 * from KS_Class Where ID='" &KS.G("FolderID") & "'", Conn, 1, 3
						ClassType=RST("ClassType")
						RST("FolderName") = FolderName
						If  ClassType="2" Then
						  RST("Folder")=FolderEname
						End If
						RST("FolderTemplateID") = FolderTemplateID
						RST("TopFlag")          = TopFlag
						RST("PubTF")            = PubTF
						RST("MailTF")           = MailTF
						RST("FilterTF")         = FilterTF
						RST("MapTF")            = MapTF
						RST("WapSwitch")        = WapSwitch
						RST("firstAlphabet")= Chr(KS.G("firstAlphabet"))
						RST("Target")    = KS.G("Target")
						RST("AdminPurView")=KS.S("AdminPurview")
						OldTemplateID=RST("TemplateID")
						OldWapTemplateID=RST("WapTemplateID")
						
						RST("FolderFsoIndex")   = FolderFsoIndex
						RST("TemplateID")       = TemplateID
						If KS.WSetting(0)="1" Then
						RST("WapFolderTemplateID")=WapFolderTemplateID
						RST("WapTemplateID")    = WapTemplateID
						end if
						RST("FnameType")        = FnameType
						RST("FsoType")          = FsoType
						RST("FolderDomain")     = FolderDomain
						RST("ClassPurview")     = ClassPurview
						RST("CommentTF")        = CommentTF
						RST("DefaultArrGroupID")= KS.FilterIDS(GroupID)
						RST("AllowArrGroupID")  = KS.FilterIDS(AllowArrGroupID)
						RST("DefaultReadPoint") = ReadPoint
						RST("DefaultChargeType")= ChargeType
						RST("DefaultDividePercent")=DividePercent
						RST("DefaultPitchTime") = PitchTime
						RST("DefaultReadTimes") = ReadTimes
						If RST("ClassType")=3 Then ClassDescript=Request.Form("ClassContent")
						RST("ClassBasicInfo")   = ClassPic & "||||" & ClassDescript & "||||" & MetaKeyWord   &"||||" & MetaDescript& "||||" & AdPa&"||||" & AdminMobile
						RST("ClassDefineContent")=ClassDefineContent
						RST.Update
					 
					 
					 
					    Call KS.FileAssociation(1000,RST("ClassID"),RST("ClassBasicInfo")&ClassDefineContent ,1)
					  	  
						
						  
						 
						  If RST("TN") = "0" Then
						 
						  dim FieldXML,XMLStr,XML2,NewNode,TN
						  set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
						  FieldXML.async = false
						  FieldXML.setProperty "ServerHTTPRequest", true 
						  FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/Class_sitecache.xml"))	
						  if FieldXML.parseError.errorCode<>0 Then
							 XMLStr=""
							 XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
							 XMLStr=XMLStr&"<website>" &vbcrlf
							 XMLStr=XMLStr&"</website>" &vbcrlf
						 	 Call KS.WriteTOFile(KS.Setting(3)&"Config/Class_sitecache.xml",xmlstr)
						  end if 
						  Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@id=" & RST("ClassID") & "]")
						  If  Node Is Nothing Then
								XMLStr="<item id="""& RST("ClassID") &""">" &vbcrlf
								XMLStr=XMLStr&" <weburl>" & FolderDomain &"</weburl>" &vbcrlf
								XMLStr=XMLStr&"</item>" &vbcrlf
								set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
								XML2.LoadXml(XMLStr)
								set NewNode=XML2.documentElement
								Set TN=FieldXML.DocumentElement
								TN.appendChild(NewNode)
								FieldXML.Save(Server.MapPath(KS.Setting(3)&"Config/Class_sitecache.xml"))
						  else
						   		Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@id=" & RST("ClassID")& "]")
								Node.childnodes(0).text=FolderDomain
								FieldXML.Save(Server.MapPath(KS.Setting(3)&"Config/Class_sitecache.xml"))
						  end if 
						  
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FolderDomain,ClassPurview,ClassID from KS_Class where TS Like '%" & KS.G("FolderID") & "%'", Conn, 1, 3
						   Do While Not RS.EOF
							RS("FolderDomain") = FolderDomain
							RS.Update
							RS.MoveNext
						   Loop
						   RS.Close
						  End If
						  Set RS = Nothing
						  
						
						 Dim ENode:Set ENode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & KS.G("FolderID") & "']")
						
						 
						 
						 If TemplateID<>OldTemplateID AND KS.ChkClng(request("autotemplate1"))=1 Then
						  Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set TemplateID='" & TemplateID &"' Where Tid='" & KS.G("FolderID") & "'")
						 End If
						 If ChannelID<>8 and KS.WSetting(0)="1" and WapTemplateID<>OldWapTemplateID AND KS.ChkClng(request("autotemplate2"))=1 Then
						  Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set WapTemplateID='" & WapTemplateID &"' Where Tid='" & KS.G("FolderID") & "'")
						 End If
						 
						 TJ=RST("tj")
						RST.Close 
						'==========修改栏目英文名称===================================
					    if request("FolderEname")<>request("oldFolderEname") then
						  rst.open "select id,foldername,folder,tj from KS_Class Where classtype=1 and ts like '%" & KS.G("FolderID") &"%'",conn,1,3
						  do while not rst.eof
						   dim ntj:ntj=tj-1
						   dim ii,folderarr:folderarr=split(rst("folder"),"/")
						   dim newFolder:NewFolder=""
						   for ii=0 to ubound(folderarr)
						     if ii=ntj then
							   if newfolder="" then
							   newfolder=KS.G("folderename")
							   else
							    newfolder=newfolder & "/" & KS.G("folderename")
							   end if
							 else
							  if newfolder="" then
							  newfolder=folderarr(ii)
							  else
							  newfolder=newfolder & "/" & folderarr(ii)
							  end if
							 end if
						   next
						   rst("folder")=newfolder
						  rst.movenext
						  loop
						  rst.close
						end if
						'================================================================
                         Application(KS.SiteSN&"_class")=""
					    KSCls.ClassAction  channelid
						If ClassType=3  And KS.C_S(ChannelID,7)=1 Then  '单页面生成
						  RST.Open "select top 1 * From KS_Class Where ID='"& KS.G("FolderID") &"'",CONN,1,1
						  If NOT RST.Eof Then
						  Set KSR=New Refresh
						  Call KSR.RefreshFolder(ChannelID,RST)  '调用栏目刷新函数
						  Set KSR=Nothing
						  End IF
						  RST.Close
						End If
						 KS.AlertDoFun "" & MyFolderName & "信息修改成功!","location.href='" & KS.Setting(3) & KS.Setting(89) & "system/KS.Class.asp?channelid=" &request("schannelid")&"&page=" & Request("page")&"&a=" & request("a")&"&oids=" & request("oids")&"';"
					End If
			        Set RST = Nothing
                 
			End Sub
			
			Function IsAlphabet(ByVal str )
				dim re
				set re = New RegExp 
				re.Global = True 
				re.IgnoreCase = True 
				re.Pattern="^[A-Za-z\d\s\_]+$" 
				IsAlphabet = re.Test(str) 
			End Function
			
End Class
%> 

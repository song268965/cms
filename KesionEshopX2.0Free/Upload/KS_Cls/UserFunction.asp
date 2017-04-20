<!--#include file="Kesion.IfCls.asp"-->
<%
Sub Echo(sStr)
	 Response.Write sStr 
	 'Response.Flush()
End Sub
  
public Sub Scan(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "[#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = Parse(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function Parse(sTemplate, iPosBegin)
	Dim iPosCur, sToken, sTemp,MyNode
	iPosCur      = InStr(iPosBegin, sTemplate, "]")
	sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
	iPosBegin    = iPosCur+1
	select case Lcase(sTemp)
		case "pubtips"  
		  If Action="Edit" Then
		    echo "修改" & KS.C_S(Channelid,3)
		  Else
		    echo "发布" & KS.C_S(Channelid,3)
		  End If
		case "selectclassid"
		   If KS.C("UserName")="" Then  '游客投稿
		    echo "[" & KS.GetClassNP(KS.S("ClassID")) & "] <a href=""Contributor.asp""><<重新选择>></a>"
			echo "<input type=""hidden"" name=""ClassID"" value=""" & KS.S("ClassID") & """>"
		   Elseif action="Edit" Then 
		   Call KSUser.GetClassByGroupID(ChannelID,ClassID,"ClassID",Selbutton,0) 
		   Else
		   Call KSUser.GetClassByGroupID(ChannelID,KS.S("ClassID"),"ClassID",Selbutton,0) 
		   End If
		case "status"
		   if action="Edit" Then
		     If RS("Verific")<>1 Then
			  if rs("verific")=2 Then
		       echo "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' checked>放入草稿</label>"
			  else
		       echo "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' value='2'>放入草稿</label>"
			  end if
			 Else
			  echo "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
		   Else
		    echo "<label style='color:#999999'><input type='checkbox' name='status' value='2'>放入草稿</label>"
		   End If
		case "readpoint" 
		   if action="Edit" Then echo rs("readpoint") else echo "0"
		case "showsetthumb"
		   if action<>"Edit" Then echo "<label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>使用图集的第一幅图</label>"
		case "showphotourl"
			If KS.C("UserName")="" Then%>
				  <td width="240"><input class="textbox" name='PhotoUrl'  type='text' style="width:230px;" id='PhotoUrl' maxlength="100" /></td>
			 <%Else
			   if action="Edit" Then PhotoUrl=rs("PhotoUrl")
			   %>
				<td width="340"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:230px;" id='PhotoUrl' maxlength="100" />
                 <input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('SelectPhoto.asp?pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=4',500,360,window,document.myform.PhotoUrl);" />
			 </td>
		 <%End If
		case "showselecttk"
		 ' If KS.C("UserName")<>"" Then echo "<button type=""button""  class=""pn"" onClick=""AddTJ();"" style=""margin: -6px 0px 0 0;""><strong>图片库...</strong></button>"
		case "showquestionandverify"
			If KS.C("UserName")="" Then
			Call PubQuestion()
			%>
				<dd>
					<div>验证码：</div>
 <input name="Verifycode" id="Verifycode" type="text" class="yzmInput" maxlength="6" size="10" autocomplete="off" />
  <span id="showVerify"><img style='height:28px;cursor:pointer' title='点击刷新' align='absmiddle' src='../../plus/verifycode.asp' onClick='this.src="../../plus/verifycode.asp?n="+ Math.random();'></span>
				</dd>
		<% End If
		case "showstyle"
		   if action="Edit" Then
		    ShowStyle=RS("ShowStyle"): PageNum=RS("PageNum")
		   Else
		    ShowStyle=4 : PageNum=10
		   End If
		   %>
		   <table width='80%'><tr><td>
								  <input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='4'<%If ShowStyle="4" Then response.Write " checked"%>><img src='../images/default/p4.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td>
								  <input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='1'<%If ShowStyle="1" Then response.Write " checked"%>><img src='../images/default/p1.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>
		   <td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='2'<%If ShowStyle="2" Then Response.Write " checked"%>><img src='../images/default/p2.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='3'<%If ShowStyle="3" Then Response.Write " checked"%>><img src='../images/default/p3.gif'>
		   </td><td><input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='5'<%If ShowStyle="5" Then Response.Write " checked"%>><img src='../images/default/p5.gif'>
		   </td></tr></table><div style="margin:5px" id="pagenums"
			<%If ShowStyle="1" or ShowStyle="4" Then Response.Write " style='display:none'"%>
			>每页显示<input type="text" name="pagenum" value="<%=PageNum%>" style="text-align:center;width:30px">张</div>
		<%
		case "downlblist" echo DownLBList
		case "downyylist" echo DownYYList
		case "downsqlist" echo DownSQList
		case "downptlist" echo DownPTList
		case "sizeunit"
		   Dim SizeUnit      
		   If Action="Edit" Then SizeUnit = Right(rs("DownSize"), 2) Else SizeUnit="KB"
			Response.Write "<input name=""SizeUnit"" type=""radio"" value=""KB"" id=""kb"""
			If SizeUnit = "KB" Then response.write "checked"
			Response.Write "><label for=""kb"">KB</label> " & vbCrLf
			Response.Write "<input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb"""
			If SizeUnit = "MB" Then response.write "checked"
			Response.Write "><label for=""mb"">MB</label> " & vbCrLf
		case "downurls" If Action="Edit" Then echo Split(RS("DownUrls")&"|||","|")(2)
		case "content"
		  If Action="Edit" Then
		   select case KS.ChkClng(KS.C_S(ChannelID,6))
		    case 1 if not KS.IsNul(rs("ArticleContent")) then echo (rs("ArticleContent"))
			case 2 if not KS.IsNul(rs("PictureContent")) then echo (rs("PictureContent"))
			case 3 if not KS.IsNul(rs("DownContent")) then echo (rs("DownContent"))
			case 4 if not KS.IsNul(rs("FlashContent")) then echo (rs("FlashContent"))
		   end select
		  End If
		case "changesurl"
		 If Action="Edit" Then
		    Dim ChangesUrl
		    If RS("Changes")="1" Then ChangesUrl=RS("ArticleContent") Else ChangesUrl=""
		    If ChangesUrl = "" Then
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='50' class='textbox'>")
			Else
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl& "' size='50' class='textbox'>")
			End If
			If RS("Changes") = "1" Then
				 Response.Write (" <input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
			Else
				 Response.Write (" <input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'> 使用转向链接</font>")
		    End If
		End If
		case else
		   Dim II,DV,XNode
		   if instr(sTemp,"|select")<>0 then  '下拉及联动
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
				 If Action="Edit" Then DV=RS(split(sTemp,"|")(0)) Else DV=xnode.selectsinglenode("defaultvalue").text
				 KS.Echo KSUser.GetSelectOption(ChannelID,FieldDictionary,FieldXML,xnode.selectsinglenode("fieldtype").text,split(sTemp,"|")(0),xnode.selectsinglenode("width").text,xnode.selectsinglenode("options").text,DV) 
			 end if
		   elseif instr(sTemp,"|radio")<>0 then  '单选
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
			      If Action="Edit" Then DV=RS(split(sTemp,"|")(0)) Else DV=xnode.selectsinglenode("defaultvalue").text
			       KS.Echo  KSUser.GetRadioOption(split(sTemp,"|")(0),xnode.selectsinglenode("options").text,DV)
			 End If
		   elseif instr(sTemp,"|checkbox")<>0 then  '多选
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
			      If Action="Edit" Then DV=RS(split(sTemp,"|")(0)) Else DV=xnode.selectsinglenode("defaultvalue").text
			       KS.Echo  KSUser.GetCheckBoxOption(split(sTemp,"|")(0),xnode.selectsinglenode("options").text,DV)
			 End If
		   elseif instr(sTemp,"|unit")<>0 then  '单位
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
			      If Action="Edit" Then DV=RS(split(sTemp,"|")(0)&"_unit") Else DV=xnode.selectsinglenode("defaultvalue").text
			       KS.Echo  KSUser.GetUnitOption(split(sTemp,"|")(0),xnode.selectsinglenode("unitoptions").text,DV)
			 End If
		   elseif action="Add" then
		     if lcase(trim(stemp))="author" and Not KS.IsNul(KS.C("UserName")) then
			   echo KSUser.GetUserInfo("RealName")
			 end if
			 if lcase(trim(stemp))="origin" and Not KS.IsNul(KS.C("UserName")) then
			   echo LFCls.GetSingleFieldValue("SELECT top 1 CompanyName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'")
			 end if
			 if instr(stemp,"@")<>0 then  '有设置默认值
			   dim tempstr:tempstr=split(stemp,"@")
			   if tempstr(1)="now" then 
			    echo now
			   elseif tempstr(1)="date" then
			    echo date
			   else
			     echo LFCls.GetSingleFieldValue("SELECT top 1 " & split(tempstr(1),"|")(1) &" From " & split(tempstr(1),"|")(0) &" Where UserName='" & KSUser.UserName & "'")
			   end if
			 end if
			 
		   elseif action="Edit" Then
		     echo rs(trim(stemp))
		   Elseif left(lcase(sTemp),3)="ks_" then
		     echo server.htmlencode(GetDiyFieldValue(FieldXML,sTemp))
		   End If
	end select
	Parse    = iPosBegin
 End Function
 
 
'=========================扫描会员中心主体框架 增加于2010年6月========================================

Public Sub Kesion()
         Dim LoginTF:LoginTF=Cbool(KSUser.UserLoginChecked)
		 Dim FileContent,MainUrl,RequestItem,TemplateFile
		 Dim KSR,ParaList
		 FCls.RefreshType = "Member"   '设置当前位置为会员中心
		 Set KSR = New Refresh
		 TemplateFile=KS.Setting(116)
		 If LoginTF=True Then  TemplateFile=KS.U_G(KSUser.GroupID,"templatefile")
		 If trim(TemplateFile)="" Then TemplateFile=KS.Setting(116)
         If trim(TemplateFile)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		 FileContent = KSR.LoadTemplate(TemplateFile)
		 If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		  FileContent = KSR.KSLabelReplaceAll(FileContent)
		 Set KSR = Nothing
		 ScanTemplate RexHtml_IF(FileContent)
End Sub	
 
public Sub ScanTemplate(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "{#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = ParseTemplate(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function ParseTemplate(sTemplate, iPosBegin)
		Dim iPosCur, sToken, sTemp,MyNode,CheckJS
		iPosCur      = InStr(iPosBegin, sTemplate, "}")
		sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
		iPosBegin    = iPosCur+1
		select case Lcase(sTemp)
			case "showusermain"  loadMain
			case "showmymenu"  ShowMyMenu
			case "userid"  echo ks.c("userid")
			case "username" echo ksuser.username
			case "groupname" echo KS.U_G(KSUser.GroupID,"groupname")
			case "showsynchronizedoption"  echo KSUser.ShowSynchronizedOption(CheckJS)
			case "checkjs" echo checkjs
			
			case "userface"
			  Dim UserFaceSrc:UserFaceSrc=KSUser.GetUserInfo("UserFace")
			  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.Setting(3) & userfacesrc
			  response.write userfacesrc
			case else
			  response.write ksuser.getuserinfo(sTemp)
		end select
		 ParseTemplate=iPosBegin
End Function


 
 Sub ShowMyMenu()
   %>
		
	<%if ks.Setting(201)="1" then
       Call KSUser.QianDao()
	 end if%>
		
		<h3>个人中心</h3>
		<div class="left02">
		  <ul>
		     <li><a href="user_editinfo.asp">会员资料</a>
			 <span><a href="user_rz.asp" title="实名认证">实名认证</a></span>
			 </li>
		     <li><a href="user_logmoney.asp">消费明细</a>
			 <span><a href="user_payonline.asp">充值</a></span></li>
		 <%
		
		
		 If KS.C_S(5,21)=1 Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<a href=""user_order.asp"">商城订单</a>"
			Response.Write "<span><a href=""user_order.asp?action=coupon"">优惠券</a></span></li>"
		 End if
		

		 If KSUser.CheckPower("s20")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"<a href=""User_ItemSign.asp"">文档签收</a>"
			Response.Write "<span><a  href=""User_ItemSign.asp"">查看</a></span></li>"
		 End If
         if KSUser.CheckPower("s16")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"<a href=""User_favorite.asp"">我的收藏</a>"
			Response.Write "<span><a  href=""User_MyComment.asp"">评论</a></span></li>"
		 End If
         if KSUser.CheckPower("s17")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write" <a href=""Complaints.asp"">投诉建议</a>"
			Response.Write "<span><a  href=""Complaints.asp?Action=Add"">发布</a></span></li>"
		 End If
		 If KSUser.CheckPower("s19")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<a href=""user_form.asp"">表单管理</a>"
			Response.Write "</li>"
		 End If
		   
		 %>
		  </ul>
		</div>
     <h3>内容发布</h3>
		<div class="left02">
		  <ul>
		    <%
  
  
  	 
'模型的投稿
if KSUser.CheckPower("s18")<>false Then 
			 Dim Node,Ico,ItemUrl,PubUrl,Itemname,ThumbnailsConfig,groupidok
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig()
			 
			 
			 For Each Node In Application(KS.SiteSN&"_ChannelConfig").DocumentElement.SelectNodes("channel[@ks21=1 and @ks36!=0 and @ks0!=6]")
				Ico=Node.SelectSingleNode("@ks51").text
				If KS.IsNul(Ico) Then Ico="images/icon7.png"
				Select Case KS.ChkClng(Node.SelectSingleNode("@ks6").text) 
				  Case 1 ItemUrl="User_ItemInfo.asp":PubUrl="user_post.asp"
				  Case 2 ItemUrl="User_ItemInfo.asp":PubUrl="user_post.asp"
				  Case 3 ItemUrl="User_ItemInfo.asp":PubUrl="user_post.asp"
				  Case 4 ItemUrl="User_ItemInfo.asp":PubUrl="User_Myflash.asp"
				  Case 5 ItemUrl="User_ItemInfo.asp":PubUrl="User_MyShop.asp"
				  Case 7 ItemUrl="User_ItemInfo.asp":PubUrl="User_MyMovie.asp"
				  Case 8 ItemUrl="User_ItemInfo.asp":PubUrl="User_MySupply.asp"
				  Case 9 ItemUrl="User_MyExam.asp":ItemUrl="User_MyExam.asp"
			   End Select
			        ItemName=Node.SelectSingleNode("@ks52").text : groupidok=0
					If KS.IsNul(ItemName) Then ItemName=KS.C_S(Node.SelectSingleNode("@ks0").text,3)
					ThumbnailsConfig=Split(Node.SelectSingleNode("@ks46").text&"||||||||||||||||||||||||||||||||||","|")
					if  Node.SelectSingleNode("@ks36").text =3 then
					   if InStr(ThumbnailsConfig(17),ksuser.groupid) >0 then
					   		groupidok=22
					   else
					   		groupidok=33		
					   end if
					end if
					if groupidok=22 or groupidok=0 then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<a href=""" & ItemUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & """>" & ItemName & "</a>" 
					end if	
			   		
					If KS.ChkClng(Node.SelectSingleNode("@ks6").text) =9 Then
					Response.Write "<span><a href=""User_MyExam.asp?action=record"">记录</a></span></li>"
					Else
						if groupidok=22 or groupidok=0 then
							Response.Write "<span><a href=""" & PubUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & "&Action=Add"">发布</a></span></li>"
						end if	
					End If
			 Next
	   End If
	   
		 '团购
		  If KSUser.CheckPower("s09")=true  and   KS.C_S(5,21)="1" Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"<a href=""User_groupbuy.asp"">团购</a>"
			Response.Write "<span><a  href=""User_groupbuy.asp?Action=Add"" >发布</a></span></li>"
		 End If
		 
		 
		 '求职
		If KS.C_S(10,21)=1 Then
			If KSUser.GetUserInfo("UserType")=0 Then
				If KSUser.CheckPower("s14")=true Then 
					 If KS.C_S(10,21)="1" Then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<a href=""User_JobResume.asp"">求职</a>"
						Response.Write "<span><a href=""User_JobResume.asp"">+简历</a></span></li>"

					 End If
				End If
			Else			 
			if KSUser.CheckPower("s14")=true  then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<a href=""user_Enterprise.asp?action=job"">招聘</a>"
						Response.Write "<span><a href=""User_JobCompanyZW.asp?Action=Add"">发布</a></span></li>"
			 end if
            End If
		End If
		 If KSUser.CheckPower("s09")=true  and  KS.ASetting(0)="1" Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"<a href=""User_Askquestion.asp"">问答</a>"
			Response.Write "<span><a  href=""../ask/a.asp"" target=""_blank"">提问</a></span></li>"
		 End If

		 %>
		 </ul>
		</div>
		
		 <%
End Sub
 
'------扫描会员中心主体框架------

 
 
 
 '取得某个字段的默认值
 Function GetDiyFieldValue(FieldXML,FieldName) 
        Dim V,Xnode:Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & FieldName &"']")
		if Not Xnode is nothing then
		  v=Xnode.selectsinglenode("defaultvalue").text
		End If
		If Instr(V,"|")<>0 Then
			 If Not KS.IsNul(KS.C("UserName")) Then
			 V=LFCls.GetSingleFieldValue("select top 1 " & Split(V,"|")(1) & " from " & Split(V,"|")(0) & " where username='" & KSUser.UserName & "'") 
			 Else
			 V=""
			 End If
		End If
		GetDiyFieldValue=v
 End Function

'参数 isTemplate true 后台生成表单模板调用,channelid 模型id, id 编辑时的文章ID
Function GetInputForm(IsTemplate,ChannelID,FieldXML,FieldNode,FieldDictionary,id,KSUser,RS)
  Dim FNode,ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,BigPhoto,Intro,FullTitle,ReadPoint,Province,City,County,UserDefineFieldArr,I,SelButton,MapMarker,PicUrls,ShowStyle,PageNum,DownSize,SizeUnit,DownPT,YSDZ,ZCDZ,JYMM,DownUrls,FlashUrl,ChargeTips,AddDate,oTid,oID,SeloTid,ChangesUrl,Changes
if IsObject(RS) And IsTemplate=false Then
	If Not RS.Eof Then
		     If KS.C_S(ChannelID,42) =0 And RS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
			   RS.Close():Set RS=Nothing
			   KS.ShowTips "error",server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
			   KS.Die ""
			 End If
		     ClassID  = RS("Tid")
			 oTid     = RS("oTid")
			 oID      = RS("oID")
			 Title    = RS("Title")
			 KeyWords = RS("KeyWords")
			 AddDate  = RS("AddDate")
			 If ChannelID<>5 and KS.C_S(ChannelID,6)<>7 and KS.C_S(ChannelID,6)<>8 Then 
			  Author   = RS("Author")
			  Origin   = RS("Origin")
			 End If
			 If ChannelID<>5 and KS.C_S(ChannelID,6)<>8 Then 
			  ReadPoint= RS("ReadPoint")
			 End If
			 Select Case KS.ChkClng(KS.C_S(ChannelID,6))
			  case 1 
			    ChargeTips="阅读"
			    Content  = RS("ArticleContent"):FullTitle= RS("FullTitle")
				Province = RS("Province"):  City  = RS("City") : County = RS("County")
                Intro    = RS("Intro")
				Changes  = RS("Changes")
				If RS("Changes") = "1" Then ChangesUrl     = Trim(RS("ArticleContent"))
			  case 2 
			    ChargeTips="查看"
				Province = RS("Province"):  City  = RS("City"): County = RS("County")
			    PicUrls  = RS("PicUrls"):Content  = RS("PictureContent")
				ShowStyle= RS("ShowStyle"):PageNum  = RS("PageNum")
			  case 3
			    ChargeTips="下载"
			    DownSize = RS("DownSize") : DownPT = RS("DownPT") :DownUrls=Split(RS("DownUrls")&"|||","|")(2)
				YSDZ = RS("YSDZ") : ZCDZ = RS("ZCDZ") : JYMM = RS("JYMM") : BigPhoto=RS("BigPhoto")
				SizeUnit = Right(DownSize, 2):DownSize = Replace(DownSize, SizeUnit, "") : Content=RS("DownContent")
			  case 4
			    ChargeTips="观看"
			    FlashUrl = RS("FlashUrl")
				Content  = RS("FlashContent")
			  case 5
			    BigPhoto=RS("BigPhoto")
				Content=RS("ProIntro")
			  case 7
			    ChargeTips="观看"
				Content=RS("MovieContent")
			  case 8
			    Province = RS("Province"):  City  = RS("City"): County = RS("County")
			 End Select
			 
			 Verific  = RS("Verific")
			 If Verific=3 Then Verific=0
			 PhotoUrl   = RS("PhotoUrl")
			 
			 
			 if KS.ChkClng(KS.C_S(ChannelID,6))<=2 Then
			  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonuserform").text="1" Then	MapMarker=RS("MapMarker")
			 End If
				'自定义字段
				Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
				If diynode.length>0 Then
					Set FieldDictionary=KS.InitialObject("Scripting.Dictionary")
					For Each FNode In DiyNode
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),RS(FNode.SelectSingleNode("@fieldname").text)
					   If FNode.SelectSingleNode("showunit").text="1" Then
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text) &"_unit",RS(FNode.SelectSingleNode("@fieldname").text&"_Unit")
					   End If
					Next
				End If
		   End If
		   SelButton=KS.C_C(ClassID,1)
		   SeloTid=KS.C_C(oTid,1)
		Else
		 If IsTemplate=false Then
		     If Not KS.IsNul(KS.C("UserName")) Then
		     Call KSUser.CheckMoney(ChannelID)
			 Author=KSUser.GetUserInfo("RealName")
			 Origin=LFCls.GetSingleFieldValue("SELECT top 1 CompanyName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'")
			 End If
			 ClassID=KS.S("ClassID")
			 If ClassID="" Then ClassID="0"
			 If ClassID="0" Then
			  SelButton="选择" & KS.GetClassName(ChannelID) &"..."
			 Else
			 SelButton=KS.C_C(ClassID,1)
			 End If
			 ReadPoint=0 : Verific=0 : ShowStyle=4 : PageNum=10
			 YSDZ="http://" : ZCDZ="http://":AddDate=now:ChangesUrl=""
		 Else
		    ShowStyle="[#ShowStyle]":PageNum="[#PageNum]":PicUrls="[#PicUrls]":Title="[#Title]":FullTitle="[#FullTitle]":KeyWords="[#KeyWords]":Author="[#Author]":Origin="[#Origin]":Province="[#Province]":City="[#City]":Count="[#County]":Author="[#Author]":Intro="[#Intro]":Content="[#Content]":PhotoUrl="[#PhotoUrl]":BigPhoto="[#BigPhoto]":ReadPoint="[#ReadPoint]":Verific="[#Verific]":MapMarker="[#MapMarker]":DownSize="[#DownSize]":DownPT="[#DownPT]":YSDZ="[#YSDZ]":ZCDZ="[#ZCDZ]":JYMM="[#JYMM]":DownUrls="[#DownUrls]":AddDate="[#AddDate]":ChangesUrl="[#ChangesUrl]"
			'自定义字段
			Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
			If diynode.length>0 Then
					Set FieldDictionary=KS.InitialObject("Scripting.Dictionary")
					For Each FNode In DiyNode
					   Dim dv:dv=lcase(FNode.SelectSingleNode("defaultvalue").text&"")
					   if dv="now" or dv="date" or instr(dv,"ks_user")<>0 or instr(dv,"ks_enterprise")<>0 then '有默认值
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),"[#" & FNode.SelectSingleNode("@fieldname").text & "@" & FNode.SelectSingleNode("defaultvalue").text &"]"
					   else
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),"[#" & FNode.SelectSingleNode("@fieldname").text  &"]"
					   end if
					Next
			End If
			
		 End If
		End If
		%><div class="title" style="text-align:center"><%
	      If IsTemplate Then
		  Response.Write "[#PubTips]"
		  ElseIF ID<>0 Then
			  response.write "修改" & KS.C_S(ChannelID,3)
		  Else
		      response.write "发布" & KS.C_S(ChannelID,3)
		 End iF%></div>
 

 <dl class="dtable">
 <%
 
Call KS.LoadFieldGroupXML()
			
Dim TypeNode,TypeNodes
IF IsObject(Application(KS.SiteSN & "_FieldGroupXml")) Then
  Set TypeNodes=Application(KS.SiteSN & "_FieldGroupXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &"]")
  For Each TypeNode In TypeNodes
   Dim nNode:Set nNode=FieldXML.DocumentElement.SelectNodes("fielditem[showonuserform=1&&fieldtype!=13&&@groupid=" & TypeNode.SelectSingleNode("@id").text & "]")
   If nNode.Length>=1 Then
    if TypeNodes.length>1 then response.write "<FIELDSET><LEGEND align=""left"">" & TypeNode.SelectSingleNode("@groupname").text &"</LEGEND>" & vbcrlf
 
For Each FNode In nNode
	    If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
		    Response.Write KSUser.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary)
		Else
		 Dim XWidth:XWidth =KS.ChkClng(FNode.SelectSingleNode("width").text) : If  XWidth=0 Then  XWidth=250
		 Dim XTitle:XTitle=FNode.SelectSingleNode("title").text
	     Select Case lcase(FNode.SelectSingleNode("@fieldname").text)
   
   case "tid"
		   if IsTemplate=False and conn.execute("select count(1) from KS_Class Where ChannelID=" & ChannelID)(0)=1 Then
		   response.write "<input type='hidden' value='" & Conn.Execute("select top 1 ID From KS_Class Where ChannelID=" & ChannelID)(0) &"' name='ClassID' id='ClassID'/>"
		   Else
 %>
 <dd>
  <div><%=Replace(XTitle,"栏目",KS.GetClassName(ChannelID))%>：</div>
  <%
				If IsTemplate Then
				  Response.Write "[#SelectClassID]"
				Else
				 If KS.C("UserName")="" and KS.S("ClassID")<>"" Then  '游客投稿
					response.write "[" & KS.GetClassNP(KS.S("ClassID")) & "] <a href=""Contributor.asp""><<重新选择>></a>"
					response.write "<input type=""hidden"" name=""ClassID"" value=""" & KS.S("ClassID") & """>"
				 Else
				  Call KSUser.GetClassByGroupID(ChannelID,ClassID,"ClassID",Selbutton,0) 
				 End If
				End If
			If ChannelID=5 Then %><span id="brandarea">
					<%If ID<>"0" Then
						     Response.Write GetBrandByClassID(ClassID,BrandID)
				    End If%></span></dd><dd><div>我的分类：</div><select class="textbox" size='1' name='UserClassID' style="width:150">
					<option value="0">-不指定分类-</option>
						<%=KSUser.UserClassOption(3,UserClassID)%>
					 </select>		
				 <a href="User_Class.asp?Action=Add&typeid=3"><font color="red">添加</font></a>	
			<%End If%>
 </dd>
<%    End If
case "otid"
  Dim OtherModel:OtherModel=KS.ChkClng(FNode.SelectSingleNode("defaultvalue").text)
  If OtherModel<>0 Then
    If KS.IsNul(SeloTid) Then SeloTid="选择" & KS.GetClassName(OtherModel) &"..."
    %>
	 <dd>
     <div><%=XTitle%>：</div>
	<%
    If IsTemplate Then
		 Response.Write "[#SelectClassID]"
	Else
      Call KSUser.GetClassByGroupID(OtherModel,oTid,"oTid",SeloTid,oid) 
   End If
   %>
   </dd>
   <%
  End If
case "title"%>
 <dd>
    <div><%=XTitle%>：</div>
    <input class="textbox" name="Title" type="text" id="Title" style="width:<%=XWidth%>px;" value="<%=Title%>" maxlength="100" /><span style="color: #FF0000">*</span>
 </dd>
<%case "fulltitle"%>
 <dd>
    <div><%=XTitle%>：</div>
    <input class="textbox" name="FullTitle" type="text" style="width:<%=XWidth%>px; " value="<%=FullTitle%>" maxlength="100" /><span class="msgtips"> 完整标题，可留空</span>
 </dd>
<%case "turnto"%>
 <dd><div><%=XTitle%>:</div>
	 <%if IsTemplate Then%>
	 [#ChangesUrl]
	 <%Else%>
		<%If ChangesUrl = "" Then
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='50' class='textbox'>")
				Else
				 Response.Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl & "' size='50' class='textbox'>")
				End If
				If Changes = "1" Then
				 Response.Write (" <input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
				Else
				 Response.Write (" <input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'> 使用转向链接</font>")
		End If
	End If%>
  </dd>
<%case "keywords"%>
 <dd>
    <div><%=XTitle%>：</div>
    <input name="KeyWords"  class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:<%=XWidth%>px; " /><a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号&quot;,&quot;隔开</span>
  </dd>
<%case "author"%>
<dd>
    <div><%=XTitle%>：</div>
    <input name="Author" class="textbox" type="text" id="Author" style="width:<%=XWidth%>px; " value="<%=Author%>" maxlength="30" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的作者</span>
  </dd>
<%case "origin"%>
<dd>
   <div><%=XTitle%>：</div>
  <input class="textbox" name="Origin" type="text" id="Origin" style="width:<%=XWidth%>px; " value="<%=Origin%>" maxlength="100" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的来源</span>
</dd>
<%case "adddate"%>
<dd>
   <div><%=XTitle%>：</div>
  <input class="textbox" name="AddDate" type="text" id="AddDate" style="width:<%=XWidth%>px; " onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});"  value="<%=AddDate%>" maxlength="100" /> <span class="msgtips">格式：<%=now%></span>
</dd>
<%case "nature"%>
 <dd>
       <div><%=XTitle%>：</div>
	   <%If Istemplate Then%>
	   类别:<select name='DownLB'>[#DownLBList]</select> 语言:<select name='DownYY' size='1'>[#DownYYList]</select>授权:<select name='DownSQ' size='1'>[#DownSQList]</select>
	   <%Else%>
	   类别:<select name='DownLB'><%=DownLBList%></select> 语言:<select name='DownYY' size='1'><%=DownYYList%></select>授权:<select name='DownSQ' size='1'><%=DownSQList%></select>
	   <%End If%>
	   <%
		Response.Write "大小:<input type='text' class='textbox' style='text-align:center' size=4 id='DownSize' name='DownSize' value='" & DownSize & "'> "
If Istemplate Then
      Response.Write "[#SizeUnit]"
Else
		If SizeUnit = "KB" Then
			Response.Write "<input name=""SizeUnit"" type=""radio"" value=""KB"" checked id=""kb""><label for=""kb"">KB</label> " & vbCrLf
			Response.Write "<input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb""><label for=""mb"">MB</label> " & vbCrLf
		Else
			Response.Write "<input name=""SizeUnit"" type=""radio"" value=""KB""  id=""kb""><label for=""kb"">KB</label> " & vbCrLf
			Response.Write "<input type=""radio"" name=""SizeUnit"" value=""MB"" checked id=""mb""><label for=""mb"">MB</label> " & vbCrLf
		End If
	End If%>                      
</dd>
<%case "platform"%>
<dd>
     <div><%=XTitle%>：</div>
     <input class='textbox' type='text' size=70 style="width:<%=XWidth%>px" name='DownPT' value="<%=DownPT%>"><br>
		 <font class="platselect">平台选择 <%If Istemplate Then%>[#DownPTList]<%Else%><%=DownPTList%><%End If%></font>
</dd>
<%case "ysdz"%>
<dd>
   <div><%=XTitle%>：</div>
   <input class="textbox" name="YSDZ" type="text" value="<%=YSDZ%>" id="YSDZ" style="width:<%=XWidth%>px; " maxlength="100" />
</dd>
<%case "zcdz"%>
<dd>
   <div><%=XTitle%>：</div>
   <input class="textbox" name="ZCDZ" type="text" value="<%=ZCDZ%>" id="ZCDZ" style="width:<%=XWidth%>px; " maxlength="100" />
</dd>
<%case "jymm"%>
<dd>
   <div><%=XTitle%>：</div>
   <input class="textbox" name="JYMM" type="text" value="<%=JYMM%>" id="JYMM" style="width:<%=XWidth%>px; " maxlength="100" />
</dd>
<%case "area"%>
<dd>
    <div><%=XTitle%>：</div>
    <script>try{setCookie("pid",'<%=province%>');setCookie("cid",'<%=city%>');}catch(e){}</script>
							<script src="../plus/area.asp" language="javascript"></script>
							<script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
							  <%end if%>
							  <%if City<>"" Then%>
							  $('#City').val('<%=City%>');
							  <%end if%>
							  <%if County<>"" Then%>
							  $('#County').val('<%=County%>');
							  <%end if%>
	</script>
 </dd>
<%case "map"%>
<dd>
    <div><%=XTitle%>：</div>
    经纬度：<input class="textbox" value="<%=MapMarker%>" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/edit_add.gif' align='absmiddle' border='0'>添加电子地图标志</a>
  </dd>
<%case "intro"%>
<%if KS.ChkClng(KS.C_S(ChannelID,6))=1 then%>
 <dd>
   <div><%=XTitle%>：<font>（<input name='AutoIntro' type='checkbox' checked value='1'>自动截取内容的200个字作为导读）</font></div>
   <textarea class='textbox' name="Intro" style='width:95%;height:65px'><%=intro%></textarea>
  </dd>
<%end if%>
<%case "content"%>
<%  select case KS.ChkClng(KS.C_S(ChannelID,6))
     case 1
%>
<dd ID='ContentArea'>
   <div><%=XTitle%>:<font>（如果<%=KS.C_S(ChannelID,3)%>较长可以使用分页标签：[NextPage]）</font></div>
  <%
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/showonuserform").text="1" Then
	%>
		<strong><%=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/title").text%>:</strong>
		<iframe id='upiframe' name='upiframe' src='BatchUploadForm.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' height='24'></iframe>
	 <%end if%>
		<%
		If GetEditorType()="CKEditor" Then
		    Response.write "<table><tr><td><textarea id=""Content"" name=""Content"">"& Server.HTMLEncode(Content) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('Content', {width:""850px"",height:""200px"",toolbar:""Simple"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script></td></tr></table>"
		Else
			 Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:98%;height:200px;"">" &Content&"</script>"
	         Response.Write "<script>setTimeout(""baiduContent = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(GetEditorToolBar("NewsTool"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:200,elementPathEnabled:false });"",10);</script>"
       End If
		%>
</dd>
<%case 2%>
	<dd>
		<div>显示样式：</div>
		<%if IsTemplate Then%>
		[#ShowStyle]
		<%Else%><table width='80%'><tr><td>
					<input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='4'<%If ShowStyle="4" Then response.Write " checked"%>><img src='../images/default/p4.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td>
					<input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='1'<%If ShowStyle="1" Then response.Write " checked"%>><img src='../images/default/p1.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>
		   <td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='2'<%If ShowStyle="2" Then Response.Write " checked"%>><img src='../images/default/p2.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='3'<%If ShowStyle="3" Then Response.Write " checked"%>><img src='../images/default/p3.gif'>
		   </td><td><input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='5'<%If ShowStyle="5" Then Response.Write " checked"%>><img src='../images/default/p5.gif'>
		   </td></tr></table><div style="margin:5px" id="pagenums"
			<%If ShowStyle="1" or ShowStyle="4" Then Response.Write " style='display:none'"%>
			>每页显示<input type="text" name="pagenum" value="<%=PageNum%>" style="text-align:center;width:30px">张</div>
		<%End If%>
  </dd>
 <dd>
       <div><%=XTitle%>：</div>
	    <table>
		 <tr>
		  <td><div class="pn">
			 <span id="spanButtonPlaceholder"></span>
			</div></td>
		 <td>
		 <button type="button"  class="pn" onClick="OnlineCollect()" style="margin: -6px 0px 0 0;"><strong>网上地址</strong></button><%if IsTemplate Then%>
		 [#ShowSelectTK]
	   <%ElseIf KS.C("UserName")<>"" Then%>
		<!-- <button type="button"  class="pn" onClick="AddTJ();" style="margin: -6px 0px 0 0;"><strong>图片库...</strong></button>-->
	   <%End If%>
		 </td>
		 </tr>
		</table>

		<label><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>图片添加水印</label>
		<div id="divFileProgressContainer"></div>
	    <div id="thumbnails"></div>
		<input type='hidden' name='PicUrls' id='PicUrls' value="<%=PicUrls%>">
</dd>
<%case Else%>
<dd>
     <div><%=XTitle%>：</div>
      <%
	    If GetEditorType()="CKEditor" Then
		    Response.write "<table><tr><td><textarea id=""Content"" name=""Content"">"& Server.HTMLEncode(Content) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('Content', {width:""850px"",height:""200px"",toolbar:""Basic"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script></td></tr></table>"
		Else
			 Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:80%;height:200px;"">" &Content&"</script>"
	         Response.Write "<script>setTimeout(""editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:200,elementPathEnabled:false });"",10);</script>"
		End If
		%>
</dd>
<%end select%>
<%case "address"%>
  <%If channelid=8 Then%>
    <%if KS.IsNul(KS.C("UserName")) or ContactMan="" Then%>
	
	  <dd> <div>您的联系资料:</div></dd>
		 <dd>
			   <table width="98%" border=0 align="center" cellspacing="1" cellpadding=2 style="BORDER-COLLAPSE: collapse">
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>联 系 人：</p></td>
					<td valign=top width="34%"><input type="text" name="ContactMan" id="ContactMan" class="textbox"/></td>
					<td valign=top width="16%"><p align=right>联系电话：</p></td>
					<td valign=top width="35%"><input type="text" name="Mobile" id="Mobile" class="textbox"/></td>
				  </tr>
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>公司名称：</p></td>
					<td valign=top width="34%"><input type="text" name="CompanyName" id="CompanyName" class="textbox"/></td>
					<td valign=top width="16%"><p align=right>联系地址：</p></td>
					<td valign=top width="35%"><input type="text" name="Address" id="Address" class="textbox"/></td>
				  </tr>
		
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>电子邮件：</p></td>
					<td valign=top width="34%"><input type="text" name="Email" id="Email" class="textbox"/></td>
					<td valign=top width="16%"><p align=right>邮政编码：</p></td>
					<td valign=top width="35%"><input type="text" name="Zip" id="Zip" class="textbox"/></td>
				  </tr>
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>传真号码：</p></td>
					<td valign=top width="34%"><input type="text" name="Fax" id="Fax" class="textbox"/></td>
					<td valign=top width="16%"><p align=right>公司网址：</p></td>
					<td valign=top width="35%"><input type="text" name="HomePage" id="HomePage" class="textbox"/></td>
				  </tr>
			  </table>
		  </dd>
	
	
	<%Else%>
		 <dd>
			  <div>您的联系资料:
			  <font>(
			  <%if KSUser.GetUserInfo("UserType")=1 Then
				 response.write "<a href='" & EditInfoUrl & "'><font color=#999999>修改联系资料</font></a>"
				else
				 response.write "<a href='" & EditInfoUrl & "'><font color=#999999>修改联系资料</font></a>"
				end if
			  %>)
			  </font>
			  </div>
		  </dd>
		 <dd>
			   <table width="98%" height=121 border=0 align="center" cellspacing="1" cellpadding=2 style="BORDER-COLLAPSE: collapse">
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>联 系 人：</p></td>
					<td valign=top width="34%"><%=ContactMan%></td>
					<td valign=top width="16%"><p align=right>联系电话：</p></td>
					<td valign=top width="35%"><%=Tel%> / <%=Mobile%></td>
				  </tr>
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>公司名称：</p></td>
					<td valign=top width="34%"><%=CompanyName%></td>
					<td valign=top width="16%"><p align=right>所在地区：</p></td>
					<td valign=top width="35%"><%=Address%></td>
				  </tr>
		
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>电子邮件：</p></td>
					<td valign=top width="34%"><%=email%></td>
					<td valign=top width="16%"><p align=right>邮政编码：</p></td>
					<td valign=top width="35%"><%=zip%></td>
				  </tr>
				  <tr class="tdbg">
					<td valign=top width="15%"><p align=right>传真号码：</p></td>
					<td valign=top width="34%"><%=fax%></td>
					<td valign=top width="16%"><p align=right>公司网址：</p></td>
					<td valign=top width="35%"><%=HomePage%></td>
				  </tr>
			  </table>
		  </dd>
   <%end if%>
   <%else%>
		 <dd>
			 <div><%=KS.C_S(ChannelID,3)%>地址：</div>
				<table class="downtable" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td  width="275"><input type="text" class="textbox" name='DownUrlS' id='DownUrlS' value='<%=DownUrls%>' style="width:<%=XWidth%>px; "> <span style="color: #FF0000">*</span>
						 </td>
							<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadsoft']/showonuserform").text="1" Then%>
							<td><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../user/User_upfile.asp?type=UpByBar&channelid=<%=ChannelID%>' frameborder="0" scrolling="no" width='280' height='25'></iframe></td>
							<%end if%>
					</tr>
				</table>
		</dd>
<%end if%>
<%case "price"%>
<dd>
	<div><%=XTitle%>：</div>
	 <%if channelid=8 then%>
	 <input class="textbox" size=35  value="<%=Price%>" name="Price">
      <font color=#ff6600> *</font> <span class="msgtips">如:3-5千</span>
	 <%else%>
<font color=blue>参考价<input name="Price" type="text" id="Price" value="<%=Price%>" size="6" class="textbox" onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">元</font>&nbsp;&nbsp; 会员价<input name="Price_Member" type="text" id="Price_Member" value="<%=Price_Member%>" size="6" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">元
    <%end if%>
</dd>
<%
case "totalnum"%>
<dd>
		<div>库存设置：</div>
		库存数量&nbsp;<input name="TotalNum" type="text" class="textbox" id="TotalNum" style="width:40px; " value="<%=TotalNum%>" size="40" maxlength="40" />&nbsp;库存报警下限数&nbsp;<input name="AlarmNum" type="text" class="textbox" id="AlarmNum" style="width:40px; " value="<%=AlarmNum%>" size="40" maxlength="40" />
		 <span style="color: #FF0000">*</span>
		 单件重量<input name="Weight" type="text" class="textbox" id="Weight" style="width:40px; " value="<%=Weight%>" size="10" maxlength="10" /> KG 
 </dd>
<%
case "prointro"
%><dd>
     <div><%=XTitle%>：</div>
      <%If GetEditorType()="CKEditor" Then
		    Response.write "<table><tr><td><textarea id=""Content"" name=""Content"">"& Server.HTMLEncode(Content) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('Content', {width:""850px"",height:""200px"",toolbar:""Basic"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script></td></tr></table>"
		Else
			 Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:80%;height:200px;"">" &Content&"</script>"
	         Response.Write "<script>setTimeout(""editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:200 ,elementPathEnabled:false});"",10);</script>"
		End If
		%>
</dd>
<dd><div>空间显示：</div>
    <input name="ShowOnSpace" type="radio" value="1" <%If ShowOnSpace="1" Then Response.Write " checked"%> />是
	<input name="ShowOnSpace" type="radio" value="0" <%If ShowOnSpace="0" Then Response.Write " checked"%>/>否	
</dd>
<%
case "photourl"
%>
<% select case KS.ChkClng(KS.C_S(ChannelID,6))
case 2%>
<dd>
     <div><%=XTitle%>：</div>
     <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
									 <tr>
									 <%if IsTemplate Then%>
									 [#ShowPhotoUrl]
									 <%
									 Else
									   If KS.C("UserName")="" Then%>
									  <td width="240"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:<%=XWidth%>px;" id='PhotoUrl' maxlength="100" />
									  </td>
									 <%Else%>
									  <td width="240"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:<%=XWidth%>px;" id='PhotoUrl' maxlength="100" /></td><%if KS.C_S(ChannelID,16)="1" then%><td width="80" valign="top" style="padding-top:2px;"><input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('SelectPhoto.asp?pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=<%=ChannelID%>',500,360,window,document.myform.PhotoUrl);" /></td><%end if%>
								      
									 <%End If
									 End If%>
									  <td>
									  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../user/User_upfile.asp?channelid=<%=ChannelID%>&Type=Pic' frameborder=0 scrolling=no width='250' height='32'> </iframe>
									  </td>
									 </tr>
									 </table><%if IsTemplate Then%>
										[#ShowSetThumb]
										<%elseif action<>"Edit" Then%>
										 <label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>使用图集的第一幅图</label>
										<%end if%>
	</dd>
<%case 3%>
<dd>
      <div><%=XTitle%>：</div>
      <input class="textbox"  name="PhotoUrl" value="<%=PhotoUrl%>" type="text" id="PhotoUrl" style="width:<%=XWidth%>px; float:left;margin-top:3px;margin-right:3px;" maxlength="100" /><input type="hidden" name="BigPhoto" id="BigPhoto" value="<%=BigPhoto%>"/>
		<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadphoto']/showonuserform").text="1" Then%>
		<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../user/User_upfile.asp?channelid=<%=ChannelID%>&Type=Pic' frameborder=0 scrolling=no width='250' height='30'> </iframe>
		<%end if%>
</dd>
<%case 5%>         <dd>
                                        <div>商品图片：</div>
										 小图：
                                       <input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:<%=XWidth%>px;" id='PhotoUrl' maxlength="100" />
                                          &nbsp;
                                          <input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('SelectPhoto.asp?pagetitle=<%=Server.URLEncode("选择图片")%>&channelid=5',500,360,window,document.myform.PhotoUrl);" />
							            <br/>大图：
                                        <input class="textbox" name='BigPhoto' value="<%=BigPhoto%>" type='text' style="width:<%=XWidth%>px;" id='BigPhoto' maxlength="100" />
                                          &nbsp;
                                          <input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('SelectPhoto.asp?pagetitle=<%=Server.URLEncode("选择图片")%>&channelid=5',500,360,window,document.myform.BigPhoto);" />
							   </dd>
								<dd><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='../user/User_upfile.asp?channelid=5&Type=Pic' frameborder=0 scrolling=no width='95%' height='30'> </iframe></dd>
                    <dd>
                           <div>关 键 字：</div>
                           <input name="KeyWords" class="textbox" type="text" value="<%=KeyWords%>" id="KeyWords" style="width:<%=XWidth%>px; " /> <a href="javascript:void(0)" onclick="GetKeyTags()">【自动获取】</a>
						                <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span> 
                                </dd>
                                <dd>
                                        <div><%=KS.C_S(ChannelID,3)%>型号：</div>
                                       <input name="ProModel" class="textbox" type="text" value="<%=ProModel%>" id="ProModel" style="width:<%=XWidth%>px; "  maxlength="30" />
                                        <span style="color: #FF0000">*</span>
                                </dd>
                                <dd>
                                        <div><%=KS.C_S(ChannelID,3)%>规格：</div>
                                        <input name="ProSpecificat" class="textbox" type="text" id="ProSpecificat" value="<%=ProSpecificat%>" style="width:<%=XWidth%>px; " maxlength="100" />
                                        <span style="color: #FF0000">*</span>
								</dd><dd>
								  <div>品牌/商标：</div>
								 <input name="TrademarkName" class="textbox" type="text" id="TrademarkName" value="<%=TrademarkName%>" style="width:<%=XWidth%>px; " maxlength="100" />
				    </dd>
								<dd>
								  <div>生产商：</div>
								  <input name="ProducerName" class="textbox" type="text" id="ProducerName" value="<%=ProducerName%>" style="width:<%=XWidth%>px; " maxlength="100" />
							      <span style="color: #FF0000">*</span>
				    </dd>
								<dd>
								  <div>商品单位：</div>
								 <input name="Unit" type="text" class="textbox" id="Unit" style="width:40px; " value="<%=Unit%>" size="40" maxlength="40" />&nbsp;(例:本)<span style="color: #FF0000">*</span>
				    </dd>
								
								
								
								
<%case Else%>
<dd>
    <div><%=XTitle%>：</div>
   <input name='PhotoUrl' style="float:left;width:230px;margin-top:3px;margin-right:4px" type='text' id='PhotoUrl' value="<%=PhotoUrl%>" size='40'  class="textbox"/>
    <table border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td width="80" valign="top" style="padding-top:2px;">
   <input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('SelectPhoto.asp?pagetitle=<%=Server.URLEncode("选择图片")%>&channelid=<%=ChannelID%>',500,360,window,$('#PhotoUrl')[0]);" /></td>
	<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadphoto']/showonuserform").text="1" Then%>	
	<td><iframe  id='UpPhotoFrame' name='UpPhotoFrame' src='../user/User_UpFile.asp?Type=Pic&ChannelID=<%=ChannelID%>' frameborder="0" scrolling="No"  width='340' height='32'></iframe></td>
	<%end if%>   
	</tr>
	</table>
</dd>
<%
end select
case "uploadflash"
%><dd>
      <div><%=KS.C_S(ChannelID,3)%>地址：</div>
       <input class="textbox" name='FlashUrl' value="<%=FlashUrl%>" type='text' style="width:<%=XWidth%>px;float:left" id='FlashUrl' maxlength="100" />
		 <iframe id='UpFlashFrame' name='UpFlashFrame' src='../user/User_Upfile.asp?type=UpByBar&channelid=<%=channelid%>' frameborder=0 scrolling=no width='300' height='30'> </iframe>
 </dd>
<%
case "chargeoption"%>
<dd>
        <div><%=ChargeTips%><%=KSUser.GetModelCharge(channelid)%>：</div>
       <input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6">
		 <%if KS.ChkClng(KS.C_S(ChannelID,34))=0 then
		  response.write KS.Setting(46)
		  elseif KS.ChkClng(KS.C_S(ChannelID,34))=1 then
		   response.write "元"
		  else
		   response.write "个"
		  end if
		 %> <span class="msgtips">如果免费<%=ChargeTips%>请输入“<font color=red>0</font>”</span>
   </dd>
<%case "picturecontent"%>
<dd>
          <div><%=XTitle%>：</div>
		  <%
		  If GetEditorType()="CKEditor" Then
		    Response.write "<table><tr><td><textarea id=""Content"" name=""Content"">"& Server.HTMLEncode(Content) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('Content', {width:""850px"",height:""200px"",toolbar:""Basic"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script></td></tr></table>"
		Else
			 Response.Write "<script id=""Content"" name=""Content"" type=""text/plain"" style=""width:90%;height:200px;"">" &Content&"</script>"
	         Response.Write "<script>setTimeout(""editor = " & GetEditorTag() &".getEditor('Content',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:200,elementPathEnabled:false });"",10);</script>"
		End If
		%>
</dd>
<%case "movieact"%>
 <dd>
    <div><%=XTitle%>：</div>
     <input name="MovieAct" class="textbox" type="text" value="<%=MovieAct%>" id="MovieAct" style="width:<%=XWidth%>px; "  maxlength="30" />
 </dd>
<%case "moviedy"%>
 <dd>
    <div><%=XTitle%>：</div>
     <input name="MovieDY" class="textbox" type="text" value="<%=MovieDY%>" id="MovieDY" style="width:<%=XWidth%>px; "  maxlength="30" />
 </dd>
<%case "screentime"%>
 <dd>
    <div><%=XTitle%>：</div>
     <input name="ScreenTime" class="textbox" type="text" value="<%=ScreenTime%>" id="ScreenTime" style="width:<%=XWidth%>px; "  maxlength="30" /> <span>如：<%=year(now)%>年<%=month(now)%>月<%=day(now)%>日</span>
 </dd>
<%case "uploadmovie"%>
<dd>
   <div><%=KS.C_S(ChannelID,3)%>地址：</div>
		 <%If id<>"0" Then
		    Dim mi,MovieArr:MovieArr=split(MovieUrl,"|||")
			For Mi=1 To Ubound(MovieArr)+1
			  dim mmarr:mmarr=split(MovieArr(mi-1)&"§","§")
			  response.write "<input  name='MovieName"&mi&"' value='" & mmarr(0) & "' style='width:50px' type='text' class='textbox'/>" &vbcrlf
			  response.write " <input class='textbox' name='MovieUrl"&mi&"' type='text' value='" & mmarr(1)&"' style='width:250px;' id='MovieUrl"&mi&"' maxlength='200' />" &vbcrlf
			  response.write "<iframe id='UpFlashFrame' name='UpFlashFrame' src='../user/User_Upfile.asp?type=UpByBar&channelid=7&FieldName=MovieUrl"&mi&"' frameborder=0 scrolling=no width='300' height='30'> </iframe><br/>" &vbcrlf
			Next
		 %>
		 <script>
		  var i=<%=Ubound(MovieArr)+1%>;
		 </script>
		 <input type="hidden" name="totalnum" id="totalnum" value="<%=Ubound(MovieArr)+1%>"/>
		<%else%>
		 <script>
		   var i=0;
		  $(document).ready(function(){ addOne(); });
		 </script>
		  <input type="hidden" name="totalnum" id="totalnum" value="0"/>
		<%end if%>	
		 <div id="uparea"></div>
		 <input type="button" value="增加一集" class="button" onclick="addOne()"/>
		 <script> 
		  function addOne(){
		    i++;
		    var str="<input  name='MovieName"+i+"' value='第"+i+"集' style='width:50px' type='text' class='textbox'/>";
			str+=" <input class='textbox' name='MovieUrl"+i+"' type='text'  style='width:250px;' id='MovieUrl"+i+"' maxlength='200' />";
			str+="<iframe id='UpFlashFrame' name='UpFlashFrame' src='../user/User_Upfile.asp?type=UpByBar&channelid=7&FieldName=MovieUrl"+i+"' frameborder=0 scrolling=no width='300' height='30'> </iframe><br/>";
			$("#totalnum").val(i);
			$("#uparea").append(str);
		  }
		 </script>
</dd>
<%case "typeid"%>
 <dd>
   <div>交易类别：</div>
   <%=KS.ReturnGQType(TypeID,0)%>
				
                    <font color=#ff6600> *</font>　 有 效 期：
                    <select class="textbox" size=1 name="ValidDate">
					 <option value="3" <% if ValidDate=3 Then Response.Write " selected"%>>三天</option>
					 <option value="7"<% if ValidDate=7 Then Response.Write " selected"%>>一周</option>
					 <option value="15"<% if ValidDate=15 Then Response.Write " selected"%>>半个月</option>
					 <option value="30"<% if ValidDate=30 Then Response.Write " selected"%>>一个月</option>
					 <option value="90"<% if ValidDate=90 Then Response.Write " selected"%>>三个月</option>
					 <option value="180"<% if ValidDate=180 Then Response.Write " selected"%>>半年</option>
					 <option value="365"<% if ValidDate=365 Then Response.Write " selected"%>>一年</option>
					 <option value="0"<% if ValidDate=0 Then Response.Write " selected"%>>长期</option>
                    </select>
                    <font color=#ff6600> *</font>
  </dd>	
<%case "gqcontent"%>
<dd>
                <div>信息内容：
                  <font color=#800000>（请详细描述您发布的供求信息）</font></div>
                  <%
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/showonuserform").text="1" Then
	%>
		<table border='0' width='100%' cellspacing='0' cellpadding='0'>
		<tr><td height='35' width=70 nowrap="nowrap">&nbsp;<strong><%=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/title").text%>:</strong></td><td><iframe id='upiframe' name='upiframe' src='../user/BatchUploadForm.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' height='24'></iframe></td></tr>
		</table>
	 <%end if%>
				<%
				 Response.Write "<script id=""GQContent"" name=""GQContent"" type=""text/plain"" style=""width:90%;height:220px;"">" &GQContent&"</script>"
	             Response.Write "<script>setTimeout(""editor = " & GetEditorTag() &".getEditor('GQContent',{toolbars:[" & Replace(GetEditorToolBar("Basic"),"'source',","") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:220,elementPathEnabled:false });"",10);</script>"
				%>
 </dd>							  
<%case else
    ' response.write "<dd>" & lcase(FNode.SelectSingleNode("@fieldname").text) &"</dd>"
   End Select
 End IF
Next
   if TypeNodes.length>1 then response.write "</FIELDSET>"
   End If
  Next
END IF



If IsTemplate Then
%>[#ShowQuestionAndVerify]
<%
ElseIf KS.C("UserName")="" Then
Call PubQuestion()
%>
    <dd>
			<div>验证码：</div>
 <input name="Verifycode" id="Verifycode" type="text" class="yzmInput" maxlength="6" size="10" autocomplete="off" />
  <span id="showVerify"><img style='height:28px;cursor:pointer' title='点击刷新' align='absmiddle' src='../../plus/verifycode.asp' onClick='this.src="../../plus/verifycode.asp?n="+ Math.random();'></span>
	</dd>
<%
End If
%>
 <dd>
  <button class="pn" id="submit1" type="submit" onclick="return(CheckForm())"><strong>OK, 保 存</strong></button>&nbsp;<%if IsTemplate Then 
   Response.Write "[#Status]" 
   Elseif id<>0 Then
		     If Verific<>1 Then
			  if Verific=2 Then
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' checked>放入草稿</label>"
			  else
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' value='2'>放入草稿</label>"
			  end if
			 Else
			  response.write "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
	ElseIf KS.C("UserName")<>"" Then
		    response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2'>放入草稿</label>"
   End If%>
 </dd>
</dl>
<br/><%
End Function
%>
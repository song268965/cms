<%@ Language="VBSCRIPT" codepage="65001" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Kesion.KeyCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim SupplyPayPoint:SupplyPayPoint=1    '默认查看供求信息扣一个点券，设置为0不扣点！

Dim KS:Set KS=New PublicCls
Dim Action
Action=KS.S("Action")
'If KS.IsNul(Request.ServerVariables("HTTP_REFERER")) and action<>"GetClubPushModel" and action<>"AjaxSqlLabel" and action<>"AjaxLabel" Then KS.Die "error"

Select Case Action
 Case "getinputclass" getinputclass
 Case "ShowUserRZ" ShowUserRZ
 case "LoadItemInfo" LoadItemInfo
 case "getmobilecode" getmobilecode
 Case "paySupplyShow" paySupplyShow
 Case "DelPhoto"  DelPhoto
 Case "SetAttributeFields" SetAttributeFields
 Case "Ctoe" CtoE
 Case "GetTags" GetTags
 Case "GetRelativeItem" GetRelativeItem
 Case "GetClassOption" GetClassOption
 Case "GetFieldOption" GetFieldOption
 Case "GetOrderOption" GetOrderOption
 Case "GetModelAttr" GetModelAttr
 Case "SpecialSubList" SpecialSubList
 Case "GetArea" GetArea
 Case "GetFunc" GetFunc
 Case "GetSchool" GetSchool
 Case "AddFriend" AddFriend
 Case "MessageSave" MessageSave
 Case "CheckMyFriend" CheckMyFriend
 Case "SendMsg" SendMsg
 Case "SearchUser" SearchUser
 Case "CheckLogin" CheckLogin
 Case "relativeDoc" relativeDoc
 Case "getModelType" getModelType
 Case "addCart" addShoppingCart
 Case "GetPackagePro" GetPackagePro
 Case "GetSupplyContact" GetSupplyContact
 Case "GetClubBoardOption" GetClubBoardOption
 Case "getclubboard" GetClubboard
 Case "GetClubPushModel" GetClubPushModel
 Case "getclubboardcategory" getclubboardcategory
 Case "getonlinelist" getonlinelist
 Case "AjaxLabel" AjaxLabel
 Case "AjaxSqlLabel"  AjaxSqlLabel
End Select
Set KS=Nothing
CloseConn()

'后台添加文档栏目选择
Sub GetInputClass()
  Call KS.LoadClassConfig()
  Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
  Dim ParentID:ParentID=KS.S("ParentID")
  If KS.S("FolderID")<>"0"  and lcase(KS.S("ischange"))="false" Then   '编辑
     Dim I
	 Dim FolderIDArr:FolderIDArr=Split(KS.C_C(KS.S("FolderID"),8),",")
     For i=0 To Ubound(FolderIDArr)
	   If Not KS.IsNul(FolderIDArr(i)) Then
	     Call LoadInputClassByParam(ChannelID,KS.C_C(FolderIDArr(I),13))
	   End If
	 Next
  Else  '添加
     Call LoadInputClassByParam(ChannelID,ParentID)
  End If
  	  KS.Die ""
End Sub

Sub LoadInputClassByParam(ChannelID,ParentID)
      Dim folderidArr:folderidArr=Split(KS.C_C(KS.S("folderid"),8)&",,,,,,,,,,,,,,,,,,,,,,",",")
	  Dim Pstr
	  If ChannelID<>0 Then Pstr=" and @ks12=" & channelid & ""
	  Pstr=Pstr & " and @ks13='" & ParentID &"'"
	  Dim TempStr
	  Dim depth:depth=1
	  Dim PubTF:PubTF=0
	  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
		depth=KS.ChkClng(Node.SelectSingleNode("@ks10").text)
		PubTF=KS.ChkClng(Node.SelectSingleNode("@ks20").text)
		Dim SelectStr:SelectStr=""
		Dim ID:ID=Node.SelectSingleNode("@ks0").text
		if trim(folderidArr(depth-1))=trim(ID) Then SelectStr=" selected=""true"""
		If KS.C("SuperTF")="1" or KS.FoundInArr(Node.SelectSingleNode("@ks16").text,KS.C("GroupID"),",") or checkxjtk(id,depth)  or Instr(KS.C("ModelPower"),KS.C_S(Node.SelectSingleNode("@ks12").text,10)&"1")>0 Then
		TempStr=TempStr &" <option value=""" & ID & """ ispub=""" & PubTF & """" & SelectStr&">" & Node.SelectSingleNode("@ks1").text& "</option>" &vbcrlf
		End If
	  Next
	  
	  if Not KS.IsNul(TempStr) or ParentID="0" Then
		 KS.Echo "<label depth=""" & depth &""">"
		 KS.Echo "<select onchange=""changeclass(this.value,'" & parentid &"'," & depth &")"" name=""m" & parentid &""" id=""m" & parentid &"""><option value='-1'>--请选择" & KS.GetClassName(ChannelID) &"--</option>" & tempstr & "</select>"
		 KS.Echo "</label>"
	  End If
 End Sub
 
 '检查栏目ID检查下级有没有允许投稿的栏目
function checkxjtk(id,tj)
     Dim Xml,Node
	 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1 and @ks10>" & tj & "]")
	 For Each Node In Xml
	   If KS.FoundInArr(Node.SelectSingleNode("@ks8").text,id,",")=true Then  '如果是他的下级
		  If KS.FoundInArr(Node.SelectSingleNode("@ks16").text,KS.C("GroupID"),",") Then
		   checkxjtk=true
		   exit function
		  End If
	   End If
	Next

  checkxjtk=false
end function



Sub ShowUserRZ()
  Dim UserName:UserName=KS.DelSQL(KS.S("UserName"))
  If KS.IsNul(UserName) Then KS.Die ""
  Dim Str,RS:Set RS=Conn.Execute("select top 1 * From KS_User Where UserName='" & UserName &"'")
  If RS.Eof And RS.Bof Then
    str="<li><span>是否实名认证：</span>否</li>"
  Else
    If RS("UserType")="0" Then
     str="<li><span>用户类型：</span>个人会员</li>"
	  If RS("IsSFZRZ")="1" Then
      str=str & "<li><span>是否实名认证：</span><font style=""color:green"">已认证</font></li>"
	  Else
      str=str & "<li><span>是否实名认证：</span><font style=""color:red"">未认证</font></li>"
	  End If
      str=str & "<li><span>真实姓名：</span>" & RS("RealName") & "</li>"
      str=str & "<li><span>所在地区：</span>" & RS("Province") & RS("City") & "</li>"
	Else
      str="<li><span>用户类型：</span>企业会员</li>"
	   Dim RSE:Set RSE=Conn.Execute("select top 1 * From KS_Enterprise Where UserName='" & UserName &"'")
	   If RSE.Eof Then
	      str=str & "<li><span>是否实名认证：</span><font style=""color:red"">未认证</font></li>"
	   Else
		  If RS("IsRz")="1" Then
		  str=str & "<li><span>是否实名认证：</span><font style=""color:green"">已认证</font></li>"
		  Else
		  str=str & "<li><span>是否实名认证：</span><font style=""color:red"">未认证</font></li>"
		  End If
      str=str & "<li><span>公司名称：</span>" & RSE("CompanyName") & "</li>"
      str=str & "<li><span>负 责 人：</span>" & RSE("ContactMan") & "</li>"
      str=str & "<li><span>所在地区：</span>" & RSE("Province") & RSE("City") & "</li>"
	  End If
	  RSE.Close
	  Set RSE=Nothing
	End If
	
	
  End If
  
  RS.Close
  Set RS=Nothing
  KS.Die "document.write('" & str &"');"
End Sub

Sub LoadItemInfo()
   Dim ChannelID:ChannelID=KS.ChkClng(Request("ChannelID"))
   Dim Tid:Tid=KS.S("Tid")
   If Tid="" or tid="0" Then KS.Die ""
   Dim oID:oID=KS.ChkClng(Request("Oid"))
   If ChannelID=0 Then KS.Die ""
   response.write "<option value='0'>--请选择" & KS.C_S(ChannelID,3) & "--</option>"
   Dim RS:SET RS=Conn.Execute("select top 500 id,title From " & KS.C_S(ChannelID,2) & " where deltf=0 and verific=1 and tid in("& KS.GetFolderTid(tid) &") order by id desc")
   Do While Not RS.Eof
    If KS.ChkClng(Oid)=KS.ChkClng(rs(0)) Then
    response.write "<option value='" & RS(0) & "' selected>" & escape(rs(1)) &"</option>"
	Else
    response.write "<option value='" & RS(0) & "'>" & escape(rs(1)) &"</option>"
	End If
   RS.MoveNext
   Loop
   RS.Close
   Set RS=Nothing
   KS.Die ""
End Sub


'通用获取手机验证码
Sub GetMobileCode()
 Dim ModelId:ModelId=KS.ChkClng(KS.S("ModelID"))
 Dim Mobile:Mobile=KS.DelSQL(KS.S("Mobile"))
 Dim UserName:UserName=KS.DelSQL(KS.S("UserName"))
 If KS.IsNul(Mobile) Then KS.Die "没有输入手机号码"
 Dim PerTime:PerTime=KS.ChkClng(split(KS.Setting(156)&"∮∮","∮")(1))
 Dim AllowPostNum:AllowPostNum=KS.ChkClng(split(KS.Setting(156)&"∮∮","∮")(2))
 Dim AllowPostNumByIP:AllowPostNumByIP=KS.ChkClng(split(KS.Setting(156)&"∮∮","∮")(3))
 if PerTime=0 Then PerTime=10
 
 Dim SmsContentArr,Content
 SmsContentArr=Split(KS.Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")
 select case ModelId
   case 101  '注册验证码
      Content=SmsContentArr(0)
	  If KS.Setting(129)="1" Then
		   If Not Conn.Execute("Select top 1 userid from KS_User Where Mobile='" & Mobile & "'").eof Then
			 KS.Die "对不起，您输入的手机号码已注册过!"
		   End If
	  End If
   case 102  '取回密码
      Content=SmsContentArr(2)
	   If  Conn.Execute("Select top 1 userid from KS_User Where UserName='" & UserName & "' and Mobile='" & Mobile & "'").eof Then
			 KS.Die "对不起，您输入的手机号码和您绑定的手机不一致!"
	   End If
   case 103  '取回密码
      Content=SmsContentArr(3)
   case 104  '自定义表单提交
      Content=SmsContentArr(21)
	  if ks.isnul(content) then Content="尊敬的用户，{$sitename}网站的表单提交验证码：{$code}。"
	   Content=replace(Content,"{$formname}",username)
   case else
      KS.Die "非法调用"
 end select
 
 If AllowPostNum>0 Then
	 Dim MobileNum:MobileNum=KS.ChkClng(Conn.Execute("Select count(*) From KS_UserRecord Where flag=" & ModelID & " and UserName='" & Mobile & "' and year(adddate)=" & Year(now) & " and month(adddate)=" & Month(Now) & " and day(adddate)=" & Day(Now))(0))
	 If MobileNum>=AllowPostNum Then KS.Die "对不起，您的操作过于频繁！请明天再重试。"
 End If	
 
 If AllowPostNumByIP>0 Then
	 Dim IpNum:IpNum=KS.ChkClng(Conn.Execute("Select count(*) From KS_UserRecord Where flag=" & ModelID & " and userip='" & KS.GetIp & "' and year(adddate)=" & Year(now) & " and month(adddate)=" & Month(Now) & " and day(adddate)=" & Day(Now))(0))
	 If IpNum>=AllowPostNumByIP Then KS.Die "对不起，您的操作过于频繁！已达到每天的发送限制，请明天再重试。"
 End If 

 Dim AllowPub:AllowPub=true 
 Dim ErrTips
 Dim Code:Code=KS.MakeRandom(6)
 Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
 RS.Open "Select Top 1 * From KS_UserRecord Where UserName='" & Mobile & "' and year(adddate)=" & Year(now) & " and month(adddate)=" & Month(Now) & " and day(adddate)=" & Day(Now) & " order by id desc",conn,1,1
 If RS.Eof And RS.Bof Then
      RS.Close
	  Conn.Execute("Insert Into KS_UserRecord([userid],[username],[flag],[note],[adddate],[userip]) values(0,'" & Mobile & "'," & ModelID & ",'" & Code & "'," & SqlNowString & ",'" & KS.GetIP() & "')")
 Else
     Dim AddDate:AddDate=RS("AddDate")
	 If PerTime>0 Then
	    If DateDiff("s",addDate,Now)<PerTime Then
		  ErrTips=(PerTime-KS.ChkClng(DateDiff("s",addDate,Now))) & "秒后，才能重新发送！"
		  AllowPub=false
		Else
	     Conn.Execute("Insert Into KS_UserRecord([userid],[username],[flag],[note],[adddate],[userip]) values(0,'" & Mobile & "'," & ModelID & ",'" & Code & "'," & SqlNowString & ",'" & KS.GetIP() & "')")
		End If
	 End If
 End If
 
'删除大于5天的无用记录
Conn.Execute("Delete From KS_UserRecord Where flag=" & ModelID & " and datediff(" & DataPart_D & ",adddate," & sqlnowstring &")>5")


 If AllowPub Then
	 Content=Replace(Content,"{$code}",code)
	 Dim Rstr
	 Rstr=KS.SendMobileMsg(Mobile,Content)
	 If Isnumeric(Rstr) and KS.ChkClng(Rstr)>0 Then
	   KS.Die "true"
	 Else
	   KS.Die Rstr
	 End If
 Else
     KS.Die ErrTips
 End If
End Sub



'标签Ajax输出
Sub AjaxLabel()
	Dim KSCls:Set KSCls=New RefreshFunction
	Dim LabelID:LabelID=KS.R(KS.S("LabelID"))   '标签ID
	Dim InfoID:InfoID=KS.R(KS.S("InfoID"))      '信息ID
	FCls.RefreshInfoID=InfoID      '设置信息ID，以取得相关链接
	IF KS.S("labtype")="-1" Then
	FCls.RefreshFolderID=KS.S("ClassID")
	End IF
	FCls.ChannelID=KS.ChkCLng(KS.S("Channelid"))
	If LabelID="" Then Response.Write "非法调用！":Response.End
	If KS.S("labeltype")="SQL" Then
		Dim KSRCls:Set KSRCls=New DIYCls
		Dim LabelName:LabelName=replace(replace("{"&split(Request.QueryString("LabelID"),"ksr")(0)&")}","ksl","("),"ksu","_")
		KS.Echo KSRCls.ReplaceDIYFunctionLabel(LabelName,"ajax")
		Set KSRCls=Nothing
	Else
		 Dim L_P
		 Dim RCls:Set RCls=New Refresh
		 Call RCls.LoadLabelToCache()    '加载标签
		 L_P=Replace(RCls.LabelXML.documentElement.selectSingleNode("labellist[@labelid='" & LabelID & "']").text,LabelID,"ajax")
		 Set RCls=Nothing
		 If L_P="" Then Response.End
		 KS.Echo KSCls.GetLabel(l_p)
	End If
	Set KSCls=Nothing
	KS.Die ""
End Sub

'SQL分页标签
Sub AjaxSqlLabel()
            Dim KSCls:Set KSCls=New DIYCls
			Dim I,KS_RS_Obj,LabelName,UserParamArr,FunctionLabelParamArr,CirLabelContent,FunctionSQL,LabelContent,TempCirContent
			Dim FunctionLabelType,ItemName,PageStyle,PerPageNumber,TotalPut,PageNum,J,TempStr,Ajax,DataSourceType,DataSourceStr
			Dim Str,CurrPage,Tconn
			Dim SqlLabel:SqlLabel=KS.S("LabelID")
			CurrPage=KS.ChkClng(KS.S("curpage"))
			If CurrPage<=0 Then CurrPage=1
			
          if SqlLabel="" Then KS.Die "error"
		  LabelName    = Replace(Replace(Split(SqlLabel,"(")(0),"'",""),"""","")
		  '用户函数参数
		  UserParamArr = Split(Replace(Replace(Replace(Replace(SqlLabel,LabelName&"(",""),")}",""),"""",""),"'",""),",")   
		  
		   Dim L_Description:L_Description=KSCls.G_S_P(LabelName &"}",1)
		   If L_Description="" Then
		    ks.die "对不起，标签不存在!"
		   Else
		    FunctionLabelParamArr = Split(L_Description,"@@@")
		    LabelContent          = Replace(KSCls.G_S_P(LabelName &"}",2),Chr(10) ,"$KS:Page$")
		   End If
		  
		   FunctionSQL=FunctionLabelParamArr(0)           '查询语句
		   FunctionSQL=Replace(FunctionSQL,"{$CurrClassID}",KS.S("classID"))
		   FunctionSQL=Replace(FunctionSQL,"{$CurrInfoID}",KS.ChkClng(KS.S("infoID")))
		   FunctionSQL=Replace(FunctionSQL,"{$CurrClassChildID}",KS.GetFolderTid(KS.S("classID")))
		   FunctionSQL=Replace(FunctionSQL,"{$CurrUserName}",KS.C("UserName"),1,-1,1)
		   If Instr(FunctionSQL,"{$GetUserName}")<>0 Then
		    If Not KS.IsNul(KS.S("UserName")) Then
		     FunctionSQL=Replace(FunctionSQL,"{$GetUserName}",KS.DelSql(KS.UrlDecode(KS.S("UserName"))),1,-1,1)
			ElseIf Not KS.IsNul(Session("SpaceUserName")) Then
			 FunctionSQL=Replace(FunctionSQL,"{$GetUserName}",Session("SpaceUserName"))
            Else
		     FunctionSQL=Replace(FunctionSQL,"{$GetUserName}",Split(KS.DelSql(Replace(KS.UrlDecode(Request.ServerVariables("QUERY_STRING")),"'","")),"/")(0),1,-1,1)
			End If
		   End If
		   LabelContent = KS.ReplaceRequest(LabelContent)    '替换request的值
		   FunctionSQL = KS.ReplaceRequest(FunctionSQL)    '替换request的值

		   For I=0 To Ubound(UserParamArr)
		    FunctionSQL  = Replace(FunctionSQL,"{$Param("&I&")}",KS.DelSQL(UserParamArr(I)))
			LabelContent = Replace(LabelContent,"{$Param("&I&")}",KS.DelSQL(UserParamArr(I)))
		   Next
		   FunctionLabelType=FunctionLabelParamArr(2)
		   If Not Isnumeric(FunctionLabelType) Then FunctionLabelType=0
		   Ajax=FunctionLabelParamArr(5)
           		   
		   ItemName=FunctionLabelParamArr(3)
		   PageStyle=FunctionLabelParamArr(4)
		   DataSourceType=FunctionLabelParamArr(6)
		   DataSourceStr=FunctionLabelParamArr(7)
		   if DataSourceType=1 Or DataSourceType=5 Or DataSourceType=6 then	DataSourceStr=LFCls.GetAbsolutePath(DataSourceStr)
		   
		   If DataSourceType=0 Then
		   Else
				on error resume next
				Set tconn = Server.CreateObject("ADODB.Connection")
				tconn.open datasourcestr
				If Err Then 
				  Err.Clear
				  Set tconn = Nothing
				  KS.Die "外部数据库连接出错!"
				End If
			 End If
		   
		   on error resume next
		   Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
           If DataSourceType=0 Then
		    KS_RS_Obj.Open FunctionSQL,Conn,1,1
		   Else
		    KS_RS_Obj.Open FunctionSQL,TConn,1,1
		   End IF
		   if err then 
		    err.clear
			KS_RS_Obj.close: set KS_RS_Obj=nothing
			KS.Die "非法调用!"
		   end if
		   
		   If Not KS_RS_Obj.Eof Then
			    Dim regEx, Matches, Match,LoopTimes
				Set regEx = New RegExp
				regEx.Pattern = "\[loop=\d*].+?\[/loop]"
				regEx.IgnoreCase = True
				regEx.Global = True
				Set Matches = regEx.Execute(LabelContent)
				If FunctionLabelType=1 Then                  '分页标签
				         PerPageNumber=0
				         For Each Match In Matches
							PerPageNumber=PerPageNumber+KSCls.GetLoopNum(Match.Value)   '每页记录数
						 Next
                         If PerPageNumber=0 Then ks.die "自定义函数标签的循环次数必须大于0"
						 
				  		TotalPut = KS_RS_Obj.recordcount
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
							 TempCirContent    = LabelContent
							 KS_RS_Obj.Move (CurrPage - 1) * PerPageNumber
						     For Each Match In Matches
								  LoopTimes=KSCls.GetLoopNum(Match.Value)   '循环次数
								  CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
								  TempCirContent    = Replace(TempCirContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",KSCls.GetCirLabelContent(CirLabelContent,KS_RS_Obj,LoopTimes,CurrPage),1,1)

								  If KS_RS_Obj.Eof Then Exit For
							 Next
							  TempStr = TempCirContent
						      str=Replace(KSCls.CleanLabel(TempStr),"$KS:Page$",vbcrlf)

				End If		 
		   Else
		     ks.die "对不起，没有内容!"
		   End if
		   KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   
		If DataSourceType=0 Then
		   Conn.Close:Set Conn=Nothing
		Else
		   TConn.Close:Set TConn=Nothing
	    End If

		Dim Tp,homeUrl,endUrl,prevUrl,nextUrl,startpage
		Dim XML:Set XML=LFCls.GetXMLFromFile("pagestyle")
		Dim Node:Set Node= XML.documentElement.selectSingleNode("/pagestyle/item[@name='" &PageStyle & "']/content")
		If Not Node Is Nothing Then
				 Tp=Node.text
		End If
		 tp=replace(tp,"{$maxperpage}",PerPageNumber)
		 tp=replace(tp,"{$totalput}",TotalPut)
		 tp=replace(tp,"{$totalpage}",PageNum)
		 tp=replace(tp,"{$currentpage}",CurrPage)
		 tp=replace(tp,"{$itemunit}",ItemName)
		 
		 
		 dim lid:lid=replace(replace(SqlLabel,"{",""),"}","")

		if (currPage=1) then
		 homeUrl = "javascript:turn(-1,'" & lid &"');"
		 prevUrl = "javascript:turn(-1,'" & lid &"')"
		else
		 homeUrl = "javascript:turn(1,'" & lid &"');"
		 prevUrl = "javascript:turn(" & currpage-1 & ",'" & lid &"');"
		end if
		if (currpage=pagenum) then
		 NextUrl = "javascript:turn(-2,'" & lid &"');"
		 endUrl = "javascript:turn(-2,'" & lid &"');"
		else
		 NextUrl = "javascript:turn(" & currpage+1 & ",'" & lid &"');"
		 endUrl="javascript:turn(" & pagenum & ",'" &lid &"');"
		end if
						
						
		 Tp=Replace(Tp,"{$homeurl}",homeurl)
		 Tp=Replace(Tp,"{$prevurl}",prevurl)
		 Tp=Replace(Tp,"{$nexturl}",nexturl)
		 Tp=Replace(Tp,"{$endurl}",endurl) 
		
		 if (instr(Tp,"{$pagenumlist}")<>0) then
						Dim p,pageStr:pageStr=""
						startpage = 1
						if (CurrPage >= 7)  then startpage = CurrPage - 5
						if (PageNum - CurrPage < 5) then startpage = pageNum - 10
						if (startpage <= 0) then startpage = 1
						Dim nn:nn = 1
						for p = startpage to pageNum
							if (p = CurrPage) then
								pageStr=pageStr & " <a href=""#"" class=""curr"">" & p & "</a>"
							else
								 pageStr=pageStr & " <a class=""num"" href=""javascript:turn(" & p& ",'" &lid &"');"">" & p & "</a>"
							end if
							nn=nn+1
							if (nn >= 10) then exit for
						Next
						Tp = replace(Tp, "{$pagenumlist}", pagestr)
		End If	  
		
		if (instr(Tp,"{$turnpage}")<>0) then
						pageStr="<select name=""page"" id=""turnpage"" onchange=""javascript:turn(this.options[this.selectedIndex].value,'" &lid &"');"">"
						for j = 1 to pageNum
						  pageStr=pageStr &"<option value=""" & j & """"
						  if j=currPage then pageStr=pageStr &" selected"
						  pageStr=pageStr &">第" & j & "页</option>"
						next
						pageStr=pageStr &"</select>"
						Tp = replace(Tp, "{$turnpage}", pageStr)
		 end if
		
		 if instr(str,"[KS:PageStyle]")=0 then  str=str  &"[KS:PageStyle]"
		 KS.Die replace(str,"[KS:PageStyle]",tp)
End Sub


Sub SetAttributeFields()
  Dim ChannelID:ChannelID=KS.ChkClng(KS.S("Channelid"))
  If ChannelID=0 Then KS.Die ""
  If KS.C_S(ChannelID,6)="1" Then
  %>
<tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eIsVideo' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("是否带视频")%>:</strong></td>
				<td><input name='IsVideo' type='radio' value='1'> 是  <input name='IsVideo' type='radio' value='0' checked> 否</td>
		  </tr>	  <%
  ElseIf KS.C_S(ChannelID,6)="3" Then
  '取得下载参数
		 Dim DownLBList, DownYYList, DownSQList, DownPTList, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
		  Set RSP = Server.CreateObject("Adodb.RecordSet")
		  RSP.Open "Select top 1 * From KS_DownParam Where ChannelID=" & ChannelID, conn, 1, 1
		  If Not RSP.Eof Then
		   DownLBStr = RSP("DownLB"):DownYYStr = RSP("DownYY"): DownSQStr = RSP("DownSQ"): DownPTStr = RSP("DownPT")
		  End If
		  RSP.Close:Set RSP = Nothing
		  '下载类别
		  LBArr = Split(DownLBStr, vbCrLf)
		  For I = 0 To UBound(LBArr)
			DownLBList = DownLBList & "<option value='" & escape(LBArr(I)) & "'>" & escape(LBArr(I)) & "</option>"
		  Next
		  '下载语言
		  YYArr = Split(DownYYStr, vbCrLf)
		  For I = 0 To UBound(YYArr)
			DownYYList = DownYYList & "<option value='" & escape(YYArr(I)) & "'>" & escape(YYArr(I)) & "</option>"
		  Next
		'下载授权
		  SQArr = Split(DownSQStr, vbCrLf)
		  For I = 0 To UBound(SQArr)
			DownSQList = DownSQList & "<option value='" & escape(SQArr(I)) & "'>" & escape(SQArr(I)) & "</option>"
		  Next
		'下载平台
		  PTArr = Split(DownPTStr, vbCrLf)
		  For I = 0 To UBound(PTArr)
			DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
		  Next
		  %>
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownServer' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("设置服务器")%>:</strong></td>
				<td><select name="DownServer"><option value='0'>↓不使用下载服务器↓</option><%
				Dim rsobj
			Set rsobj = conn.Execute("SELECT downid,DownloadName,depth,rootid FROM KS_DownSer WHERE depth=0 And ChannelID="& ChannelID)
			Do While Not rsobj.EOF
				 response.write escape("<option value=""" & rsobj("downid") & """>" & rsobj(1) & "</option>")
				rsobj.movenext
			Loop
			rsobj.Close:Set rsobj = Nothing
				
				%></select></td>
		  </tr>
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownLB' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("下载类别")%>:</strong></td>
				<td><select name="DownLB"><%=DownLBList%></select></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownYY' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("语言")%>:</strong></td>
				<td><select name="DownYY"><%=DownYYList%></select></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownSQ' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("授权方式")%>:</strong></td>
				<td><select name="DownSQ"><%=DownSQList%></select></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownPT' value='1'></td>
				<td class='clefttitle' align='right' nowrap><strong><%=escape("运行平台")%>:</strong></td>
				<td><input type="text" size='40' name='DownPT' id='DownPT' class='textbox'><br/><%=DownPTList%></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eYSDZ' value='1'></td>
				<td class='clefttitle' align='right' nowrap><strong><%=escape("演示地址")%>:</strong></td>
				<td><input type="text" size='40' name='YSDZ' id='YSDZ' class='textbox'></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eZCDZ' value='1'></td>
				<td class='clefttitle' align='right' nowrap><strong><%=escape("注册地址")%>:</strong></td>
				<td><input type="text" size='40' name='ZCDZ' id='ZCDZ' class='textbox'></td>
		  </tr>	
		  <%
  End If
  Dim RS:Set RS=server.CreateObject("ADODB.RECORDSET")
  RS.Open "Select FieldName,Title,FieldType,Options,Width From KS_Field Where FieldType<>0 and ChannelID="& ChannelID &" Order By OrderID ,FieldID",conn,1,1
  If Not RS.EOf Then
      %>
		<tr><td colspan=3 align='center' class='clefttitle' style="font-weight:bold;color:blue;height:20px;text-align:center">========<%=escape("以下列出自定义字段")%>===========</td></tr>
	<%
	  Do While Not RS.Eof 
		%>
		<tr class='tdbg'> 
			<td class='clefttitle' height='25' align='center'><input type='checkbox' name='e<%=rs(0)%>' value='1'></td>
			<td class='clefttitle' nowrap align='right'><strong><%=escape(rs(1))%>:</strong></td>
			<td><%
			Dim O_Arr,O_Len,O_Value,O_Text,F_V,K,BrStr
			select case rs(2) 
			  case 3,11
			   KS.Echo "<select class=""upfile"" style=""width:" & rs(4) & "px"" name=""" & rs(0) & """>"
			   O_Arr=Split(RS(3),vbcrlf): O_Len=Ubound(O_Arr)
				 For K=0 To O_Len
					If O_Arr(K)<>"" Then
							 F_V=Split(O_Arr(K),"|")
							 If Ubound(F_V)=1 Then  O_Value=F_V(0):O_Text=F_V(1) Else  O_Value=F_V(0):O_Text=F_V(0)
							KS.Echo Escape("<option value=""" & O_Value& """>" & O_Text & "</option>")
					End If
				 Next
			   KS.Echo "</select>"
			 case 6,7
			   O_Arr=Split(RS(3),vbcrlf): O_Len=Ubound(O_Arr)
						   For K=0 To O_Len
							   F_V=Split(O_Arr(K),"|")
							   If O_Arr(K)<>"" Then
							    If Ubound(F_V)=1 Then O_Value=F_V(0):O_Text=F_V(1) Else	O_Value=F_V(0):O_Text=F_V(0)
								If rs(2) = 6 Then
							     KS.Echo escape("<input type=""radio"" name=""" & RS(0) & """ value=""" & O_Value& """>" & O_Text)&BrStr
								Else
							    KS.Echo escape("<input type=""checkbox"" name=""" & RS(0) & """ value=""" & O_Value& """>" & O_Text)&BrStr
								End If
							 End If
						   Next
			  case else
			%><input type="text" size='40' name='<%=rs(0)%>' id='<%=rs(0)%>' class='textbox'/>
			<%
			end select 
			%>
			</td>
		 </tr>
		<%
	  RS.MoveNext
	  Loop
  End If
  RS.Close :Set RS=Nothing
End Sub

Sub getModelType()
 Dim ChannelID:ChannelID=KS.ChkClng(Request("channelid"))
 If ChannelID<>0 Then KS.Echo KS.C_S(Channelid,6)
End Sub




'取中文首字母
Sub Ctoe()
 Dim FolderName:FolderName=KS.DelSQL(UnEscape(Request("FolderName")))
 Dim CE:Set CE=New CtoECls
 Response.Write Escape(CE.CTOE(FolderName))
 Set CE=Nothing
End Sub

'取关键词tags
Sub GetTags()
 Dim Text:Text=UnEscape(Request("Text"))
 If Text<>"" Then
     Dim MaxLen:MaxLen=KS.ChkClng(KS.S("MaxLen"))
	 Dim WS:Set WS=New Wordsegment_Cls
	 Response.Write Escape(WS.SplitKey(text,4,MaxLen))
	 Set WS=Nothing
 End If
End Sub


'相关信息
Sub GetRelativeItem()
 Dim Key:Key=KS.DelSql(UnEscape(request("Key")))
 Dim Rtitle:rtitle=lcase(KS.G("rtitle"))
 Dim RKey:Rkey=lcase(KS.G("Rkey"))
 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("Channelid"))
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim Param,RS,SQL,k,SqlStr
 If Key<>"" Then
   If (Rtitle="true" Or RKey="true") Then
	 If Rtitle="true" Then
	   param=Param & " title like '%" & key & "%'"
	 end if
	 If Rkey="true" Then
	   If Param="" Then
	     Param=Param & " keywords like '%" & key & "%'"
	   Else
	     Param=Param & " or keywords like '%" & key & "%'"
	   End If
	 End If
 Else
    Param=Param & " keywords like '%" & key & "%'"
 End If
End If

 
 If Param<>"" Then 
  	Param=" where verific=1 and InfoID<>" & id & " and (" & param & ")"
 else
    Param=" where verific=1 and  InfoID<>" & id
 end if
 
  If ChannelID<>0 Then Param=Param & " and ChannelID=" & ChannelID


 SqlStr="Select top 30 ChannelID,InfoID,Title From KS_ItemInfo " & Param & " order by id desc"
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open SqlStr,conn,1,1
 If Not RS.Eof Then
  SQL=RS.GetRows(-1)
 End If
 RS.Close
 Set RS=Nothing
 If IsArray(SQL) Then
	 For k=0 To Ubound(SQL,2)
	   Response.Write "<option value='" & SQL(0,K) & "|" & SQL(1,K) & "'>" & SQL(2,K) & "</option>" 
	 Next
 End If
End Sub


'检查是否登录
Sub CheckLogin()
  If KS.C("UserName")="" Then KS.Echo "false" Else  KS.Echo "true"
End Sub

'取栏目选项
Sub GetClassOption()
 Dim From:From=KS.S("From")
 Dim ChannelID:ChannelID=KS.ChkCLng(Request.Querystring("ChannelID"))
   Dim KSUser:Set KSUser=New UserCls
   KSUser.UserLoginChecked
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr,nbsp
		KS.LoadClassConfig()
		If ChannelID<>0 Then Pstr=" and @ks12=" & channelid & ""
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
		  SpaceStr="" 
		 If ((Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or Node.SelectSingleNode("@ks20").text="0") and from<>"label" Then
		 Else
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "──" 
				 Next
			  End If
			  TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
		 End If 
		Next
		Set KSUser=Nothing
 KS.Die Escape(TreeStr)
End Sub


Sub SpecialSubList()
	  Dim ClassID, RS,SpecialXML,Node
	  ClassID=KS.ChkClng(Request.QueryString("ClassID"))
	  If ClassID=0 Then Exit Sub
	  Set RS=Conn.Execute("Select SpecialID,SpecialName from KS_Special Where ClassID=" & ClassID & " Order BY SpecialAddDate Desc")
	  If Not RS.Eof Then Set SpecialXML=KS.RsToXml(RS,"row","xmlroot")
	  RS.Close:Set RS=Nothing
	  If IsObject(SpecialXml) Then
	  	For Each node in SpecialXml.DocumentElement.SelectNodes("row")
		  KS.Echo Escape("<div><img src=""../../images/folder/Special.gif"" align='absmiddle'>")
          KS.Echo Escape("<a href=""#"">"  & Trim(Node.SelectSingleNode("@specialname").text) & "</a><input type='checkbox' onclick=""set(" & Node.SelectSingleNode("@specialid").text & ",'" & Node.SelectSingleNode("@specialname").text & "');"" value='" & Node.SelectSingleNode("@specialid").text & "'></div>")
	    Next
		 Set SpecialXml=Nothing
      End If
End Sub


Sub GetOrderOption()
    Dim Node,ChannelID
	ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
	If ChannelID=0 Then Exit Sub
	
	KS.Echo "<option value='avgscore Desc' style='color:blue'>点评平均分(降序)</option>"
	KS.Echo "<option value='avgscore Asc' style='color:blue'>点评平均分(升序)</option>"

	
	Dim FieldXML,FieldNode,KSUser
	Set KSUser=New UserCls
	Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	If Not IsObject(FieldXML) Then Exit Sub
	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype=4 || fieldtype=12 || fieldtype=5]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
							KS.Echo "<option value='" & Node.SelectSingleNode("@fieldname").text&" Desc' style='color:blue'>"  & Node.SelectSingleNode("title").text & "(降序)</option>"
							KS.Echo "<option value='" & Node.SelectSingleNode("@fieldname").text&" Asc' style='color:blue'>"  & Node.SelectSingleNode("title").text & "(升序)</option>"

						Next
				  End If
	End If
	Set KSUser=Nothing
End Sub

Sub GetFieldOption()
    Dim Node,ChannelID
	ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
	If ChannelID=0 Then Exit Sub
	Dim FieldXML,FieldNode,KSUser
	Set KSUser=New UserCls
	Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	If Not IsObject(FieldXML) Then Exit Sub
	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0&&fieldtype!=13]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
							KS.Echo "<li class='diyfield' title=""" & Node.SelectSingleNode("title").text &""" onclick=""FieldInsertCode('" & Node.SelectSingleNode("@fieldname").text & "'," & Node.SelectSingleNode("fieldtype").text & ")"">" & Node.SelectSingleNode("title").text & "</li>"
							if Node.SelectSingleNode("showunit").text="1" then
							KS.Echo "<li class='diyfield' style='color:#ff3300' title=""" & Node.SelectSingleNode("title").text &""" onclick=""FieldInsertCode('" & Node.SelectSingleNode("@fieldname").text & "_unit'," & Node.SelectSingleNode("fieldtype").text & ")"">“" & Node.SelectSingleNode("title").text & "”单位</li>"
							end if
							if Node.SelectSingleNode("fieldtype").text="14" then
							KS.Echo "<li class='diyfield' style='color:#ff3300' title=""" & Node.SelectSingleNode("title").text &""" onclick=""FieldInsertCode('" & Node.SelectSingleNode("@fieldname").text & "->title'," & Node.SelectSingleNode("fieldtype").text & ")"">“" & Node.SelectSingleNode("title").text & "”标题</li>"
							KS.Echo "<li class='diyfield' style='color:#ff3300' title=""" & Node.SelectSingleNode("title").text &""" onclick=""FieldInsertCode('" & Node.SelectSingleNode("@fieldname").text & "->url'," & Node.SelectSingleNode("fieldtype").text & ")"">“" & Node.SelectSingleNode("title").text & "”链接(URL)</li>"
							end if
						Next
				  End If
	End If
	Set KSUser=Nothing
End Sub

Sub GetModelAttr()
    Dim Node,ChannelID
	ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
	If ChannelID=0 Then Exit Sub
	Dim FieldXML,FieldNode,KSUser,Attr
	Attr=Request("Attr")
	Set KSUser=New UserCls
	Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	If Not IsObject(FieldXML) Then Exit Sub
	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype=13]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
						  If KS.FoundInArr(lcase(Attr&""),lcase(Node.SelectSingleNode("@fieldname").text),"|") Then
							KS.Echo "<label style='color:brown'><input type='checkbox' checked name='attr' value='" & Node.SelectSingleNode("@fieldname").text & "'>" & Node.SelectSingleNode("title").text & "</label>"
						  Else
							KS.Echo "<label style='color:brown'><input type='checkbox' name='attr' value='" & Node.SelectSingleNode("@fieldname").text & "'>" & Node.SelectSingleNode("title").text & "</label>"
						  End If
						Next
				  End If
	End If
	Set KSUser=Nothing
End Sub

'取得ajax选项
sub GetArea()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
if parentid<>0 then param=" where parentid=" & parentid
If Parentid<>0 Then
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
End If
Dim ors : set ors=Conn.Execute("select ID,City FROM KS_Province " & Param & " order by orderid")
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(1) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'取得院校
sub GetSchool()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
Dim sqlstr
if parentid<>0 then
  sqlstr="select id,schoolname from ks_job_school where provinceid=" & parentid & " order by orderid,id"
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
else
  sqlstr="select id,city from ks_province where parentid=0 order by orderid,id"
End If
Dim ors : set ors=Conn.Execute(sqlstr)
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(0) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'取得职能
sub GetFunc()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
if parentid<>0 then param=" where parentid=" & parentid
If Parentid<>0 Then
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
End If
Dim ors : set ors=Conn.Execute("select ID,hymc FROM KS_Job_hyzw " & Param & " order by orderid")
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(1) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'请求加为好友
Sub AddFriend()
 If KS.C("UserName")="" Then KS.Echo "nologin" : Response.End
 Dim UserName:UserName=KS.DelSQL(UnEscape(Request("UserName")))
 Dim Message:Message=KS.CheckXSS(KS.DelSQL(UnEscape(Request("Message"))))
 If Len(Message)>255 Then 
   KS.Echo escape("附言字数太多,最多只能输入255个字符!")
   exit sub
 End If
 If UserName="" Then KS.Echo escape("没有输入好友名称!") : Exit Sub
 call saveFriend(username,message,0)
 KS.Echo "success"
End Sub
'检查是否好友
Sub CheckMyFriend()
 If KS.C("UserName")="" Then KS.Echo "nologin" : Response.End
 Dim UserName:UserName=KS.DelSQL(UnEscape(Request("UserName")))
 Dim RS:Set RS=Conn.Execute("Select Top 1 accepted from KS_Friend Where UserName='" & KS.C("UserName") & "' and friend='" & username & "'")
 If rs.eof then
  KS.Echo "false"
 Else
  If rs(0)="1" then
   KS.Echo "true"
  Else
   KS.Echo "verify"
  End If
 End If
 RS.Close:Set RS=Nothing
End Sub

sub saveFriend(username,message,accepted)
		dim incept,i,sql,rs
		incept=KS.R(username)
		incept=split(incept,",")
		set rs=server.createobject("adodb.recordset")
		for i=0 to ubound(incept)
			sql="select top 1 UserName from KS_User where UserName='"&incept(i)&"'"
			set rs=Conn.Execute(sql)
			if rs.eof and rs.bof then
				rs.close:set rs=nothing
				KS.Echo escape("系统没有（"&incept(i)&"）这个用户，操作未成功。")
				Set KS=Nothing
				Response.End
			end if
			set rs=Nothing
			
			if KS.C("UserName")=Trim(incept(i)) then
			   KS.Echo escape("不能把自已添加为好友。")
			   Set KS=Nothing
			   Response.End
			end if
			
			sql="select top 1 id,friend,accepted from KS_Friend where username='"&KS.C("UserName")&"' and  friend='"&incept(i)&"'"
			set rs=Conn.Execute(sql)
			if rs.eof and rs.bof then
				sql="insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&KS.C("UserName")&"','"&Trim(incept(i))&"',"&SqlNowString&",1,'" & replace(message,"'","") & "'," & accepted & ")"
				set rs=Conn.Execute(sql)
			else
			    if rs("accepted")=0 then
				  conn.execute("update ks_friend set message='" & replace(message,"'","") & "' where id=" & rs("id"))
				end if
			end if
		
		next
		set rs=nothing
end sub
'发送短消息
Sub SendMsg()
     If Request.ServerVariables("HTTP_REFERER")="" Then KS.Die "error!"
     If KS.C("UserName")="" Then Response.End
	 Dim UserName:UserName=KS.DelSQL(UnEscape(Request("UserName")))
	 Dim Message:Message=KS.DelSQL(UnEscape(Request("Message")))
	 If Len(Message)>255 Then 
	   KS.Echo escape("附言字数太多,最多只能输入255个字符!")
	   exit sub
	 End If
     Call KS.SendInfo(UserName,KS.C("UserName"),KS.Gottopic(Message,100),Replace(Message,chr(10),"<br/>"))
	 KS.Echo "success"
End Sub

'搜索好友
Sub SearchUser()
 Dim Page:Page=KS.ChkClng(Request("Page")) : If Page= 0 Then Page=1
 Dim Province:Province=KS.DelSQL(UnEscape(Request("Province")))
 Dim City:City=KS.DelSQL(UnEscape(Request("City")))
 Dim County:County=KS.DelSQL(UnEscape(Request("County")))
 Dim Birth_Y:Birth_Y=KS.ChkClng(Request("Birth_Y"))
 Dim Birth_M:Birth_M=KS.ChkClng(Request("Birth_M"))
 Dim Birth_D:Birth_D=KS.ChkClng(Request("Birth_D"))
 Dim RealName:RealName=KS.DelSQL(UnEscape(Request("RealName")))
 Dim Sex:Sex=KS.DelSQL(UnEscape(Request("Sex")))
 Dim RS:Set RS=Server.CreateObject("Adodb.recordset")
 Dim Param,SQLStr,XML,Node,totalPut,MaxPerPage,TotalPage,N
 MaxPerPage=10
 Param="Where locked=0"
 If Province<>"" Then Param=Param &" and Province='"& Province & "'"
 If City<>"" Then Param=Param & " and city='" & city & "'"
 If County<>"" Then Param=Param & " and County='" & County & "'"
 If Sex<>"" Then Param=Param & " and sex='" & Sex & "'"
 If RealName<>"" Then Param=Param & " and realname like '%" & RealName & "%'"
 If Birth_Y<>0 Then Param=Param & " and year(birthday)=" & Birth_Y & ""
 If Birth_M<>0 Then Param=Param & " and month(birthday)=" & Birth_m & ""
 If Birth_D<>0 Then Param=Param & " and day(birthday)=" & Birth_d & ""

 
 SQLStr="Select userid,username,realname,sex,birthday,province,city,userface,isonline from ks_user " & param & " order by userid desc"
 'response.write sqlstr
 RS.Open SQLStr,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close: Set RS=Nothing
    KS.Echo Escape("<div style='text-align:center'>对不起,找不到您要查找的用户!请更换查询条件,重新检索!</div>")
 Else
    totalPut = Conn.Execute("Select Count(*) From KS_User " & Param)(0)
	If Page < 1 Then	Page = 1
	If (totalPut Mod MaxPerPage) = 0 Then
		TotalPage = totalPut \ MaxPerPage
	Else
		TotalPage = totalPut \ MaxPerPage + 1
	End If
	
	If Page > 1  and (Page - 1) * MaxPerPage < totalPut Then
		RS.Move (Page - 1) * MaxPerPage
	Else
		Page = 1
	End If
	Set XML=KS.ArrayToXML(RS.GetRows(MaxPerPage),RS,"row","")
	RS.Close : Set RS=Nothing
	If IsObject(XML) Then
	  Dim user_face,UserName
	 For Each Node In XML.DocumentElement.SelectNodes("row")
	  user_face=node.selectsinglenode("@userface").text
	  If user_face="" then 
	    if node.selectSingleNode("@sex").text="男" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
	  End If
	  If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
      username=Node.selectsinglenode("@username").text
	  KS.Echo "<li>"
	  KS.Echo "<table border='0' width='100%'>"
	  KS.Echo "<tr><td class='face'> <a href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'><img src='" & user_face & "' alt='" & username & "' /></a></td>"
	  KS.Echo " <td align='left' class='realname'>"
	  KS.Echo   Escape(Username & "(" & Node.SelectSingleNode("@realname").text & ")")
	  if isdate(Node.SelectSingleNode("@birthday").text) then
	  KS.Echo Escape(" <br />性别：" & Node.SelectSingleNode("@sex").text & "　出生：" & formatdatetime(Node.SelectSingleNode("@birthday").text,2))
	  else
	  KS.Echo Escape(" <br />性别：" & Node.SelectSingleNode("@sex").text & "　出生：" & Node.SelectSingleNode("@birthday").text)
	  end if
	  KS.Echo Escape(" <br />来自：" & Node.SelectSingleNode("@province").text & Node.SelectSingleNode("@city").text)
	  KS.Echo Escape(" <br />状态：")
	  If Node.SelectSingleNode("@isonline").text="1" Then KS.Echo escape("<font color='red'>在线</font>") else KS.Echo Escape("离线")
	  KS.Echo Escape(" <br /><img src='" & KS.Setting(3) & "images/user/log/106.gif' border='0' align='absmiddle'> <a href='javascript:void(0)' onclick=""addF(event,'" & username & "')"">加为好友</a> <img src='" & KS.Setting(3) & "images/user/mail.gif' align='absmiddle'> <a href=""javascript:void(0)"" onClick=""sendMsg(event,'" & username & "')"">发送消息</a>")
	  KS.Echo " </td>"
	  KS.Echo "</tr>"
	  KS.Echo "</table>"
	  KS.Echo "</li>"
	 Next
	End If
 End If
 If TotalPut<>0 Then
	 KS.Echo "<div id=""pageNext"" style=""text-align:center;clear:both;"">"
	 KS.Echo "<table align=""center""><tr><td>"
	 If Page>=2 Then
	  KS.Echo Escape("<a class='prev' href='javascript:void(0)' onclick=""query.page(" & Page-1 & ")"">上一页</a>")
	 End If
	 
	 If Page>=10 Then
	  KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(1)"">1</a> <a class=""num"" href=""javascript:void(0)"" onclick=""query.page(2)"">2</a> <a href='#' class='dh'>...</a>"
	 End If
	 
	 Dim StartPage,EndPage
	 If TotalPage<10 Or Page<10 Then
	  StartPage=1
	  If Page<10 Then EndPage=10 Else  EndPage=TotalPage
	 ElseIf Page>=10 Then
	  StartPage=Page-4
	  EndPage=Page+4
	 ElseIf Page<TotalPage Then
	  StartPage=TotalPage-10
	  EndPage=TotalPage
	 End If
	 If EndPage>TotalPage Then EndPage=TotalPage : StartPage=TotalPage-10
	 If StartPage<0 Then StartPage=1
	 For N=StartPage To EndPage
	  If N=Page Then
	   KS.Echo "<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</a> "
	  Else
	   KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(" & n &")"">" & N & "</a> "
	  End If
	 Next
	 
	 If TotalPage>10 And Page<TotalPage-4 Then
	  KS.Echo "<a href='#' class='dh'>...</a>"
	  KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(" & TotalPage-1 & ")"">" & TotalPage-1 & "</a> <a href=""javascript:void(0)"" class=""num"" onclick=""query.page(" & TotalPage & ")"">" & TotalPage & "</a>"
	 End If
	 If Page<>TotalPage Then
	  KS.Echo Escape("<a class='next' href='javascript:void(0)' onclick=""query.page(" & Page+1 & ")"">下一页</a>")
	 End If
	 KS.Echo "</td></tr></table>"
	 
	 KS.Echo "</div>"
	End If
End Sub

'保存空间留言
Sub MessageSave()

		If Request.servervariables("REQUEST_METHOD") <> "POST" Then
			KS.Die "<script>alert('请不要非法提交！');</script>"
		End If

		if KS.IsNul(Request.ServerVariables("HTTP_REFERER")) or instr(lcase(Request.ServerVariables("HTTP_REFERER")&""),"space")=0 then
			KS.Die "<script>alert('请不要非法提交留言!');</script>"
		end if
		
		If IsDate(Session(KS.SiteSN & "spacemsgposttime"))  Then
				If DateDiff("s",Session(KS.SiteSN & "spacemsgposttime"),Now())<30 Then '限制30秒内不允许重重提交
					KS.Die "<script>alert('请不要非法重复提交!');</script>"
				End If
		 End If

         If KS.SSetting(25)="0" Then
		   Dim KSUser:Set KSUser=New UserCls
           If KSUser.UserLoginChecked=false Then
		    Set KSUser=Nothing
		    KS.Die "<script>alert('登录后才可以留言!');</script>"
		   End IF
		   Set KSUser=Nothing
         End If
		 
		 Dim Content:Content=KS.FilterIllegalChar(KS.LoseHtml(Request("Content")))
		 Dim AnounName:AnounName=KS.LoseHtml(KS.S("AnounName"))
         Dim HomePage:HomePage=KS.LoseHtml(KS.S("HomePage"))
         Dim Title:Title=KS.FilterIllegalChar(KS.S("Title"))
		if AnounName="" Then  KS.Die "<script>alert('请填写你的昵称!');</script>"
		if Title="" Then 
		 'Response.Write("请填写留言主题!")
		 'Response.End
		End if
		if Content="" Then KS.Die "<script>alert('请填写留言内容!');</script>"
		If Len(KS.LoseHtml(Content))>=500 Then KS.Die "<script>alert('留言内容不能超过500个字!');</script>"
		IF lcase(Trim(KS.S("Verifycode")))<>lcase(Trim(Session("Verifycode"))) Then KS.Die "<script>alert('你输入的认证码不正确!');</script>"
		
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_BlogMessage where 1=0",Conn,1,3
		RS.AddNew
		 RS("AnounName")=AnounName
		 RS("Title")=Title
		 RS("UserName")=KS.S("UserName")
		 RS("HomePage")=HomePage
		 RS("Content")=Content
		 RS("UserIP")=KS.GetIP
		 If KS.SSetting(24)="1" Then
		 RS("Status")=0
		 Else
		 RS("Status")=1
		 End If
		 RS("AddDate")=Now
		RS.UpDate
		 RS.Close:Set RS=Nothing
		 Session(KS.SiteSN & "spacemsgposttime")=Now()

		 KS.Die "<script>alert('恭喜，您的留言已提交!');top.location.reload();</script>"
End Sub 

'用户名
Function GetUserID()
		  If KS.IsNul(KS.C("UserName")) Then
			GetUserID=KS.C("CartID")
		  Else
		    GetUserID=KS.C("UserName")
		  End If
End Function
'加到购物车
Sub addShoppingCart()
   Dim RS,RealPrice,n,arrGroupID,KSUser,LoginTF,str
   Dim Prodid:Prodid=KS.ChkClng(request("id"))
   Dim AttrID:AttrID=KS.ChkClng(request("attrid"))
   Dim KBID:KBID=KS.FilterIds(KS.S("KBID"))
   Dim istype:istype=KS.ChkClng(KS.S("istype"))
   if Prodid=0 then KS.Die ""
   Dim ProductList:ProductList=Session("ProductList")
   Dim Num:Num=KS.ChkClng(Request("Num"))
   Dim Attr:Attr=KS.DelSQL(UnEscape(Request("AttributeCart")))
   Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select top 1 arrGroupID From KS_Product Where id=" & Prodid,conn,1,1
   If RS.Eof And RS.Bof Then
      RS.Close : Set RS=Nothing
	  ks.die "var data={'flag':'error','str':''}"
   End If
   arrGroupID=RS(0)
     Set KSUser=New UserCls
     LoginTF=KSUser.UserLoginChecked
   If Not KS.IsNul(arrGroupID) Then
     If KS.FoundInArr(arrGroupID,KSUser.GetUserInfo("GroupID"),",")=false Then
	  RS.Close:Set RS=Nothing
	  ks.die "var data={'flag':'error1','str':''}"
	 End If
   End If
   RS.Close
   
   
   Dim RSA:Set RSA=Server.CreateObject("ADODB.RECORDSET")
   rsA.open "select top 1 * from KS_ShoppingCart where flag=0 and attrid=" & attrid & " and username='" & GetUserID & "' And proid=" & Prodid,conn,1,3
   if rsa.eof and rsa.bof then
			   rsa.addnew
    end if
	  rsa("flag")=0
	  rsa("proid")=Prodid
	  rsa("username")=GetUserID
	  rsa("attr")=attr
	  rsa("adddate")=now
	  rsa("amount")=Num
	  rsa("attrid")=attrid
	  if istype=1 then rsa("istype")=1 else rsa("istype")=0
	  rsa.update
	rsa.close:set rsa=nothing

	
	if KBID<>"" then  '加捆绑商品
	      Dim K,Price,KBIDArr
		  KBIDArr=Split(KBID,",")
		  For K=0 To Ubound(KBIDArr)
		   If KS.ChkClng(KBIDArr(K))<>0 Then
			  RS.Open "Select top 1 KBPrice From KS_ShopBundleSale Where ProID=" & Prodid & " And KBProID=" &KBIDArr(K),conn,1,1
				 If Not RS.Eof Then
			       Set RSA=Server.CreateObject("ADODB.RECORDSET")
				   RSA.Open "Select top 1 * From KS_ShopBundleSelect where username='" & GetUserID & "' and pid=" & KBIDArr(K) & " and proid=" & Prodid,conn,1,3
				  If RSA.Eof Then
					RSA.AddNew
					RSA("UserName")=GetUserID
					RSA("Pid")=KBIDArr(K)
					RSA("ProID")=Prodid
					RSA("Amount")=1
					RSA("AddDate")=Now
					RSA("Price")=RS(0)
					RSA.Update
				  End If
				  RSA.Close:Set RSA=Nothing
				 End If
				 RS.Close
		 End If
		Next
	end if
	
   str=("<div style=""FONT-SIZE: 10pt;OVERFLOW-y: auto;overflow-x:hidden; WIDTH: 100%; LINE-HEIGHT: 20px; HEIGHT: 130px"">")
   RS.Open "Select c.cartid,i.id,i.title,i.Price_Member,i.Price,i.vipprice,i.isdiscount,c.attr,c.amount,c.attrid,c.istype from KS_Product i Inner Join KS_ShoppingCart c on i.id=c.proid where c.flag=0 and c.username='" & GetUserID & "' and i.verific=1 order by i.id desc",conn,1,1
   if not rs.eof then
      str=str & ("购物车里已有<font style=""color:red"">" & rs.recordcount & "</font>样商品。")
	  str=str & "<table border=""0"" width=""380"" align=""center"" cellspacing=""0"" cellpadding=""0"" style=""margin-top:10px;"">"
	  n=1
	   Do While Not RS.Eof
	   
      If RS("AttrID")<>0 Then 
	  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RS("AttrID"))
	  If Not RSAttr.Eof Then
		RealPrice=RSAttr("Price")
	  Else
		RealPrice=RS("Price_Member")
	  End If
	  RSAttr.CLose:Set RSAttr=Nothing
	 Else
	    RealPrice=RS("Price_Member")
	 End If	  
	 
	IF Cbool(LoginTF)=true Then
	   Dim Discount:Discount=KS.U_S(KSUser.GroupID,17)
	   If Not IsNumeric(Discount) Then Discount=0
	     If KS.U_S(KSUser.GroupID,21)="1" and rs("vipprice")<>"0" then
				RealPrice=RS("VipPrice")
		 ElseIf KS.ChkClng(RS("isdiscount"))<>0 and Discount<>0 Then
	   RealPrice=FormatNumber(RealPrice*discount/10,2,-1)
	  End If
    End If 
	   
		
	    Num=rs("amount")
		If Num=0 Then Num=1
	    str=str & ("<tr><td style=""line-height:24px; border-bottom:#f1f1f1 1px solid; font-size:12px;""><input type=""checkbox"" name=""id"" value=""" & rs("cartid") & """ checked>" & n & "、<font color=#555>" & ks.gottopic(rs("title"),36) & "</font>&nbsp;&nbsp;<span style=""font-size:12px;font-weight:normal;color:#999999"">" & rs("attr") & "</span>")
		set rsa=conn.execute("Select I.ID,I.Title,i.weight,b.Price,b.amount,b.id as selid From KS_Product I inner Join KS_ShopBundleSelect b on i.id=b.pid Where B.ProID=" & RS("ID") & " and b.username='" & GetUserID & "' order by I.id")
		if not rsa.eof then
		    str=str & "<div style=""color:green;font-size:12px"">捆绑购买:</div>"
			do while not rsa.eof
			  str=str & "<div style=""line-height:20px; font-weight:normal;font-size:12px""><span style=""color:#999;float:right"">￥" &rsa("price") & "×1</span>" & rsa("title") & "</div>"
			rsa.movenext
			loop
		end if
		rsa.close

		str=str & ("</td><td width=""80"" style=""line-height:24px;  font-weight:bold; border-bottom:#f1f1f1 1px solid; color:#ff6600;"">￥" & RealPrice & "×" & Num & "</td></tr>")
		n=n+1
	   RS.MoveNext
	   Loop
	  str=str & "</table><br/>"
   end if
   str=str & "</div>"
   RS.Close: Set RS=Nothing
   ks.die "var data={'flag':'ok','str':'" & str & "'}"
End Sub

Sub GetClubBoardOption()
 Call KS.LoadClubBoard()
   Dim node,Xml,n
   Set Xml=Application(KS.SiteSN&"_ClubBoard")
        KS.Echo Escape("<select name=""boardid"">")
   for each node in xml.documentelement.selectnodes("row[@parentid=0]")
		KS.Echo Escape("<optgroup label='" & node.selectsinglenode("@boardname").text &"'>")
		for each n in xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
		   KS.Echo Escape("<option value='" & N.SelectSingleNode("@id").text & "'>---" & n.selectsinglenode("@boardname").text &"</option>")
		next
	next
	KS.Echo Escape("</select>")
    Set Xml=Nothing
End Sub

Sub GetPackagePro()
    Dim RS,Key,pricetype,tid,minPrice,maxPrice,param,sqlstr,xml,node
	dim id:id=ks.chkclng(request("id"))
	dim proid:proid=ks.s("proid")
	Key=KS.DelSQL(unescape(Request("Key")))
	pricetype=KS.ChkClng(KS.S("pricetype"))
	tid=KS.S("tid")
	minPrice=KS.S("minPrice"):If Not Isnumeric(minPrice) Then minPrice=0
	maxPrice=KS.S("maxPrice"):If Not Isnumeric(maxPrice) Then maxPrice=0
	param=" where deltf=0 and verific=1"
	if tid<>"" and tid<>"0" then param=param & " and tid in(" & KS.GetFolderTid(TID) &")"
	if proid<>"" then param=param & " and proid='"& proid & "'"
    if id<>0 then param=param & " and id<>" & id 

	If PriceType<>0 Then
	  Select Case PriceType
	   case 1 : param=param & " and price>=" & minPrice & " and price<=" & maxPrice
	   case 2 : param=param & " and VipPrice>=" & minPrice & " and VipPrice<=" & maxPrice
	   case 3 : param=param & " and Price_Member>=" & minPrice & " and Price_Member<=" & maxPrice
	  End Select
	End If
	if key<>"" Then
	  Param=Param & " and title like '%" & key & "%'"
	End If
	sqlstr="select top 500 id,title from ks_product" & param & " order by id desc"
	
	set rs=conn.execute(sqlstr)
	if not rs.eof then
	 set xml=KS.RstoXml(rs,"row","")
	end if
	rs.close:set rs=nothing
	if isobject(xml) then
	  for each node in xml.documentelement.selectnodes("row")
       ks.echo "<option value='" & node.selectsinglenode("@id").text & "'>" & node.selectsinglenode("@title").text & "</option>"
	  next
    end if
End Sub

'查看联系信息
Sub GetSupplyContact()
 Dim ID:ID=KS.ChkClng(Request("id"))
 Set RS=Server.CreateObject("Adodb.Recordset")
 RS.Open "Select top 1 b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.verific=1 and a.ID=" & ID,Conn,1,1
 if rs.eof and rs.bof then
   rs.close:set rs=nothing
   ks.echo escape("加载出错!")
 else
     dim inputer:inputer=rs("inputer")
    if not conn.execute("select top 1 adminid from ks_admin where username='" & inputer & "'").eof then '判断是管理员发布的信息，则直接显示网站的联系方式
	 rs.close:set rs=nothing
	 KS.Die LFCls.GetConfigFromXML("supply","/labeltemplate/label","noencrypted")
	end if
   Dim KSUser:Set KSUser=New UserCls
   Dim UserLoginTF:UserLoginTF=KSUser.UserLoginChecked
    Dim ClassPurView:ClassPurview=rs("classpurview")
	Dim DefaultArrGroupID:DefaultArrGroupID=rs("defaultarrgroupid")
	' If ClassPurView="2" Then
	     If SupplyPayPoint=0 Then Call ShowSupplyContactInfo(rs):rs.close:set rs=nothing:exit sub
		 IF UserLoginTF=false Then
		        response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您还没有登录，请<a href='" & KS.Setting(2) & "/user/login/' target='_blank'>登录</a>后再查看联系信息。</div>")
				rs.close:set rs=nothing
				response.end
		 ElseIf KS.FoundInArr(DefaultArrGroupID,KSUser.GroupID,",")=false Then
		        If SupplyPayPoint>0 Then
				   Dim ModelChargeType,ChargeTableName,DateField,ChargeStr,ChargeStrUnit,CurrPoint,IncomeOrPayOut  
				   ModelChargeType=KS.ChkClng(KS.C_S(8,34))
				   Select Case ModelChargeType
					case 1 ChargeStrUnit="元人民币": ChargeTableName="KS_LogMoney" : DateField="PayTime": IncomeOrPayOut="IncomeOrPayOut" : CurrPoint=KSUser.GetUserInfo("Money")
					case 2  ChargeStrUnit="分积分": ChargeTableName="KS_LogScore": DateField="AddDate":IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetUserInfo("Score")
					case else   '按点券
					  ChargeStrUnit=KS.Setting(46)&KS.Setting(45) : ChargeTableName="KS_LogPoint" : DateField="AddDate" :IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetUserInfo("Point")
					End Select
				  If Conn.Execute("Select top 1 Times From " & ChargeTableName & " Where ChannelID=8 And InfoID=" & ID & " And " & IncomeOrPayOut & "=2 and UserName='" & KSUser.UserName & "'").eof and ksuser.username<>inputer Then
		          response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>需要支付 <span style=""color:red"">" & SupplyPayPoint & " </span>" & ChargeStrUnit  & "才可以查看联系方式,您当前余额 <span style='color:green'>" & CurrPoint & " </span>" & ChargeStrUnit & ",确认支付吗？<br/><input type='button' class='btn' value='确认支付' onclick=""payShow(" & ID & ")""/></div>")
				  else
				   Call ShowSupplyContactInfo(rs)
				  end if
				Else
		          response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您的级别不够,无法查看联系信息!得到更好服务,请联系本站管理员。</div>")
				End If
				rs.close:set rs=nothing
				response.end
		 End If
	 'End If
     Call ShowSupplyContactInfo(rs)
 end if
 rs.close:set rs=nothing
End Sub
Sub ShowSupplyContactInfo(rs)
   Dim template:template=LFCls.GetConfigFromXML("supply","/labeltemplate/label","contactinfo")
   template=replace(template,"{$GetContactMan}",LFCls.ReplaceDBNull(rs("contactman"),"---"))
   template=replace(template,"{$GetContactTel}",LFCls.ReplaceDBNull(rs("tel"),"---"))
   template=replace(template,"{$GetContactMobile}",LFCls.ReplaceDBNull(rs("mobile"),"---"))
   template=replace(template,"{$GetFax}",LFCls.ReplaceDBNull(rs("fax"),"---"))
   template=replace(template,"{$GetEmail}",LFCls.ReplaceDBNull(rs("email"),"---"))
   template=replace(template,"{$GetHomePage}",LFCls.ReplaceDBNull(rs("homepage"),"---"))
   template=replace(template,"{$GetAddress}",LFCls.ReplaceDBNull(rs("address"),"---"))
   ks.echo (template)   
End Sub
Sub paySupplyShow()
 Dim ID:ID=KS.ChkClng(KS.S("ID"))
 If ID=0 Then KS.Die escape("error:参数出错!")
 Dim KSUser:Set KSUser=New UserCls
 If Cbool(KSUser.UserLoginChecked)=false Then KS.Die escape("error:请先登录!")
 If SupplyPayPoint<=0 Then Exit Sub
 Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
 RS.Open "Select top 1 * From KS_GQ where verific=1 and ID=" & ID,Conn,1,1
 If RS.Eof And RS.Bof Then 
  RS.Close :Set RS=Nothing
  KS.Die escape("error:找不到记录了!")
 End If
 Descript="查看供求信息[" & RS("Title") & "]的联系方式"
  Select Case KS.ChkClng(KS.C_S(8,34))
		 case 1 
		   If round(KSUser.GetUserInfo("money"),2)<round(SupplyPayPoint,2) Then rs.close:set rs=nothing :KS.Die escape("error:对不起，您的可用余额不足，您当前余额为 " & KSUser.GetUserInfo("money") & " 元!")
		  Call KS.MoneyInOrOut(KSUser.UserName,KSUser.UserName,SupplyPayPoint,4,2,now,0,"系统",Descript,8,ID,1)
		 case 2 
		   If round(KSUser.GetUserInfo("score"),2)<round(SupplyPayPoint,2) Then rs.close:set rs=nothing :KS.Die escape("error:对不起，您的可用余额不足，您当前积分为 " & KSUser.GetUserInfo("score") & " 分!")
		   Session("ScoreHasUse")="+" '设置只累计消费积分
		  Call KS.ScoreInOrOut(KSUser.UserName,2,KS.ChkClng(SupplyPayPoint),"系统",Descript,8,ID)
		 case else
		 	If round(KSUser.GetUserInfo("point"),2)<round(SupplyPayPoint,2) Then rs.close:set rs=nothing :KS.Die escape("error:对不起，您的可用余额不足，您当前" & KS.Setting(45) & "为 " & KSUser.GetUserInfo("point") & " " & KS.Setting(46) & "!")
		   Call KS.PointInOrOut(8,ID,KSUser.UserName,2,SupplyPayPoint,"系统",Descript,0)
  End Select
  ShowSupplyContactInfo(rs)
  RS.Close:Set RS=Nothing
End Sub

'删除图片
Sub DelPhoto()
 Dim UserName,Pass,UserID,i,p,picarr,pic:pic=KS.S("Pic")
 Dim Flag:Flag=KS.ChkClng(Request("flag"))
 Dim PicID:PicID=KS.ChkClng(Request("picid"))
 If Not KS.IsNul(Pic) Then
    PicArr=Split(pic,"|")
	If flag=1 then
	  UserName=KS.C("AdminName")
	  Pass=KS.C("AdminPass")
	  if KS.IsNul(UserName) Or KS.IsNul(Pass) Then
	   ks.die "error"
	  End If
	  If Conn.Execute("Select top 1 * From KS_Admin Where UserName='" & UserName & "' and PassWord='" & Pass & "'").eof Then
	    KS.Die "error"
	  End If
	Else
	  Set KSUser=New UserCls
      LoginTF=KSUser.UserLoginChecked
	  If LoginTF=false Then KS.Die "error!"
	  UserID=KSUser.GetUserInfo("userid")
	end if
	for i=0 to ubound(PicArr)-1
	  p=PicArr(i)
	  If Not KS.IsNul(p) Then 
	     p=replace(p,KS.Setting(2),"")
		 if flag=1 then
		  Call KS.DeleteFile(p)
		 else
		   if instr(lcase(p),lcase("/" & KS.Setting(91) & "user/" & userid & "/"))<>0 then
		    Call KS.DeleteFile(p)
		   end if
		 end if
	  End If
	next
	if picid<>0 then conn.execute("delete from ks_proimages where id=" & picid)
 End If
End Sub



Sub GetClubboard()
   Dim Xml,Node,Pid
   Pid=KS.ChkClng(KS.G("pid"))
   KS.LoadClubBoard()
	
%>
<form name="postform" method="get" action="<%=KS.Setting(3)&KS.Setting(66)%>/post.asp">
<table border="0">
 <tr>
 <td><select name="pid" id="pid" size="10" style="width:220px;height:270px" onChange="loadBoard(this.value)">
 <%
 if isobject(Application(KS.SiteSN&"_ClubBoard")) then
	 Set Xml=Application(KS.SiteSN&"_ClubBoard")
	for each node in xml.documentelement.selectnodes("row[@parentid=0]")
	  If trim(Pid)=trim(Node.SelectSingleNode("@id").text) Then
		KS.Echo "<option value='" & Node.SelectSingleNode("@id").text & "' selected>" & node.selectsinglenode("@boardname").text &"</option>"
	  Else
		KS.Echo "<option value='" & Node.SelectSingleNode("@id").text & "'>" & node.selectsinglenode("@boardname").text &"</option>"
	  End If
	next
 end if
 %>
 </select></td>
 <td><select name="bid" id="bid" size="10" style="width:220px;height:270px" onChange=" $('#navlist2').html('->'+$('#bid>option:selected').text());">
 <%
 if isobject(Application(KS.SiteSN&"_ClubBoard")) and pid<>0 then
	for each node in xml.documentelement.selectnodes("row[@parentid=" & pid &"]")
		KS.Echo "<option value='" & Node.SelectSingleNode("@id").text & "'>" & node.selectsinglenode("@boardname").text &"</option>"
	next
 end if
 %>
 </select></td>
 <td id="btns">
 <input type="button" value=" 进 入 " style="margin-bottom:6px" class="btn" onClick="toBoard()"><br/>
 <input type="submit" value=" 发 帖 " style="margin-bottom:6px" class="btn" onClick="return(toPost())"><br/>
 <input type="button" value=" 关 闭 " style="margin-bottom:6px" class="btn" onClick="parent.box.close()">
 </td>
 </tr>
</table>
 </form>
<%		  
 xml=empty
 set node=nothing
End Sub

Sub GetClubPushModel()
  KS.Echo "<select name=""ModelID"" style=""width:150px;height:220px"" onchange=""getpushclass(this.value)"" Id=""ModelID"" size=""5"">"
  Dim ModelXML,Node
  Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
  For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks6=1]")
	if Node.SelectSingleNode("@ks21").text="1" Then
	  KS.echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
	End If
  next
  KS.Die "<select>"
End Sub

Sub getclubboardcategory()
Dim BoardID:BoardID=KS.ChkClng(Request("boardid"))
If BoardID<>0 Then
     Dim CategoryStr
	 KS.LoadClubBoardCategory
	 For Each CategoryNode In Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" &BoardID &"]")
     	CategoryStr=CategoryStr & "<option value='" &CategoryNode.SelectSingleNode("@categoryid").text  & "'>" & CategoryNode.SelectSingleNode("@categoryname").text &"</option>"
	Next
	If Not KS.IsNul(CategoryStr) Then
		CategoryStr="<strong>主题分类:</strong><select name=""CategoryId"" id=""CategoryId""><option value='0'>==选择分类==</option>"  & CategoryStr &"</select>"
	End If
	KS.Die Escape(CategoryStr)
End If

End Sub

Sub getonlinelist()
 KS.Echo "<hr size='1' color='#cccccc'/>"
 Dim RS,UserName,Page,PageNum,MaxPerPage,TotalPut,n
 MaxPerPage=24
 Page=KS.ChkClng(Request("page")) : If Page=0 Then Page=1
 Set RS=Conn.Execute("select * from [KS_Online] order by startTime desc")
 If Not RS.Eof Then
            TotalPut=Conn.Execute("Select Count(1) From [KS_Online]")(0)
            If Page < 1 Then Page = 1
			If (totalPut Mod MaxPerPage) = 0 Then
				PageNum = totalPut \ MaxPerPage
			Else
				PageNum = totalPut \ MaxPerPage + 1
			End If

			If (Page - 1) * MaxPerPage < totalPut Then
				RS.Move (Page - 1) * MaxPerPage
			Else
				Page = 1
			End If
	 n=0
	 Do While NOt RS.Eof
	   n=n+1
	   userName=RS("UserName")
	   If UserName="匿名用户" Then
	   KS.Echo "<li><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/guest.png' align='absmiddle'> <a title=""用 户 名:游客&#13;当前位置:" & RS("station") & "&#13;来访时间:" & rs("starttime") & """ href='#'>游客</a></li>"
	   Else
	   KS.Echo "<li><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/" & GetOnlinePic(UserName) & "' align='absmiddle'> <a title=""用 户 名:" & username & "&#13;当前位置:" & RS("station") & "&#13;来访时间:" & rs("starttime") & """ href='" & KS.GetDomain & "space/?" & UserName & "' target='_blank'>" & KS.Gottopic(UserName,15) &"</a></li>"
	   End If
	   If N>=MaxPerPage Then Exit Do
	   RS.MoveNEXT
	 Loop
 End If
 RS.Close
 Set RS=Nothing
 KS.Echo "<div style='clear:both'></div>"
  KS.Echo "<hr size='1' color='#f1f1f1'/>"
 KS.Echo "<div style=""text-align:left"">总在线:<span color='green'>" & TotalPut & "</span> 人 共分为<font color=red>" & PageNum & "</font>页,当前第<font color=red>" & Page & "</font>页"
			  if page>1 then
			  KS.Echo " <a href=""javascript:onlineList(1);"">首页</a>"
			  KS.Echo " <a href=""javascript:onlineList(" & page-1 & ");"">上一页</a>"
			  end if
			  
			  If page<>PageNum Then
			  KS.Echo " <a href=""javascript:onlineList(" & page+1 & ");"">下一页</a>"
			  KS.Echo " <a href=""javascript:onlineList(" & pagenum & ");"">末页</a>"
			  End If
			  KS.Echo "</div>"
 
End Sub
Function GetOnlinePic(username)
 if not conn.execute("select top 1 username from ks_admin where username='" & username & "'").eof then
   GetOnlinePic="admin.png" 
 Elseif not conn.execute("select top 1 master from ks_guestboard where master+',' like'%" & username & "%,'").eof then
   GetOnlinePic="mod.png"
 Else
   GetOnlinePic="member.png"
 end if
End Function
%>
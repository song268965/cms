<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Class LocationCLs
        Private KS,DomainStr,WebNameStr,NaviStr
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.Setting(3) &KS.WSetting(4)&"/"
		  WebNameStr="首页"
		  NaviStr=" > "
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub

		Function Kesion()
		 Dim Str
		 Select Case UCase(FCls.RefreshType)
			 Case "FOLDER"  str = GetFolderLocation
			 Case "CONTENT" str = GetContentLocation
		 End Select
         Kesion=str
       End Function

	
	Function GetFolderNaviStr()
				  Dim TSArr, I
				  TSArr = Split(KS.C_C(FCls.RefreshFolderID,8), ",")
				  For I = LBound(TSArr) To UBound(TSArr) - 1
						GetFolderNaviStr = GetFolderNaviStr & NaviStr & "<a href=""" & DomainStr &"list.asp?id=" & KS.C_C(TSArr(I),9) & """>" & KS.C_C(TSArr(I),1) & "</a>"
				  Next
	End Function
	
	'取得栏目的位置导航
	Function GetFolderLocation()
				Dim FolderNaviStr:FolderNaviStr = GetFolderNaviStr()
				If FolderNaviStr="" Then
				  If Request("modelid")<>"" then
						  GetFolderLocation = "<a  href=""" & DomainStr & """>" & WebNameStr & "</a>" & FolderNaviStr & NaviStr &  "<a href=""" & DomainStr &"channel.asp?id=" & Request("modelid") & """>" & KS.C_S(KS.ChkClng(Request("modelid")),1) & "</a>" 
						  if request("tj")="1" then
						  FolderNaviStr = NaviStr & "推荐" & KS.C_S(KS.ChkClng(Request("modelid")),3)
						  elseif request("rm")="1" then
						  FolderNaviStr = NaviStr & "热门" & KS.C_S(KS.ChkClng(Request("modelid")),3)
						  else
						  FolderNaviStr = NaviStr & "最新" & KS.C_S(KS.ChkClng(Request("modelid")),3)
						  end if
						  GetFolderLocation=GetFolderLocation & FolderNaviStr
				  ElseIf KS.ChkClng(Request("tid"))<>0 Then
					dim rst:set rst=conn.execute("select top 1 tname from ks_sjclass where id=" & KS.ChkClng(request("tid")))
					if not rst.eof then
					 FolderNaviStr=NaviStr & split(rst("tname"),"|")(ubound(split(rst("tname"),"|"))-1)
					end if
					rst.close:set rst=nothing
				    GetFolderLocation = "<a  href=""" & DomainStr & """>" & WebNameStr & "</a>" & NaviStr & "<a href=""channel.asp?id=9"">考试系统</a>"  & FolderNaviStr
				  end if
				Else
					If FCls.BrandName<>"" Then
						  GetFolderLocation = "<a  href=""" & DomainStr & """>" & WebNameStr & "</a>" & FolderNaviStr & NaviStr & FCls.BrandName
					Else
						  GetFolderLocation = "<a href=""" & DomainStr & """>" & WebNameStr & "</a>" & FolderNaviStr
				   End If
				End If
	End Function
	'取得信息页导航位置的函数
	Function GetContentLocation()
				Dim Str,FolderNaviStr
				'FolderNaviStr = GetFolderNaviStr()
				FolderNaviStr = NaviStr  &  "<a href=""" & DomainStr &"list.asp?id=" & KS.C_C(Fcls.RefreshFolderID,9) & """>" & KS.C_C(Fcls.RefreshFolderID,1) &"</a>"
				Str =  "<a href=""" & DomainStr & """>" & WebNameStr & "</a>" & FolderNaviStr & NaviStr  & "浏览"& KS.C_S(FCls.Channelid,3)
				GetContentLocation = Str
	End Function
End Class



Class RefreshLocationCls
		Private KS  
		Private KMRFObj,DomainStr,WebNameStr        
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		  WebNameStr=KS.Setting(0)
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		'***********************************************************************************************************
		'取得位置导航
		'***********************************************************************************************************
		Function GetLocation(ParamNode)
		    Dim Bold, StartTag, NavType, Nav, OpenType, TitleCss,ShowTitle
			Bold       = ParamNode.GetAttribute("bold")
			StartTag   = ParamNode.GetAttribute("starttag")
			NavType    = ParamNode.getAttribute("navtype")
			Nav        = ParamNode.getAttribute("nav")
			OpenType   = ParamNode.getAttribute("opentype")
			TitleCss   = ParamNode.getAttribute("titlecss")
			ShowTitle  = ParamNode.getAttribute("showtitle")
			Dim NaviStr
			If CBool(Bold) = True Then StartTag = "<strong>" & StartTag & "</strong>"
			NaviStr = GetLocationNav(NavType, Nav)
			TitleCss=KS.GetCss(TitleCss)
			Dim ShowHomePage:ShowHomePage=ParamNode.getAttribute("showhomepage")
			if KS.isNul(ShowHomePage) Then ShowHomePage=true
			If Cbool(ShowHomePage)=TRUE Then
			 StartTag=StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr
			End If
			Select Case UCase(FCls.RefreshType)
			   Case "MORESPACE","MORELOG","MOREFRESH","MOREGROUP","MOREXC" GetLocation = GetMoreSpaceLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "SPECIALINDEX" GetLocation = GetSpecialIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "FOLDER" GetLocation = GetFolderLocation(StartTag, NaviStr, OpenType, TitleCss, FCls.RefreshFolderID)
			   Case "CONTENT" GetLocation = GetContentLocation(StartTag, NaviStr, OpenType, TitleCss,FCls.RefreshFolderID,ShowTitle)
			   Case "CHANNELSPECIAL" GetLocation = GetSpecialClassLocation(StartTag, NaviStr, OpenType, TitleCss, FCls.RefreshFolderID)
			   Case "SPECIAL"  GetLocation = GetSpecialLocation(StartTag, NaviStr, OpenType, TitleCss, FCls.RefreshFolderID)
					 
	    '--------------------------------------------会员中心导航-------------------------------------------		   
			   Case "USERREGSTEP1" GetLocation = GetUserRegLocation(1,StartTag, NaviStr, OpenType, TitleCss)
			   Case "USERREGSTEP2" GetLocation = GetUserRegLocation(2,StartTag, NaviStr, OpenType, TitleCss)
			   Case "USERREGSTEP3" GetLocation = GetUserRegLocation(3,StartTag, NaviStr, OpenType, TitleCss)
			   Case "USERLIST"  GetLocation = GetUserListLocation(StartTag, NaviStr, OpenType, TitleCss)	
			   Case "SHOWUSER"  GetLocation = GetUserInfoLocation(StartTag, NaviStr, OpenType, TitleCss)	
			   Case "MEMBER"  GetLocation = GetMemberLocation(StartTag, NaviStr, OpenType, TitleCss)	
		'-------------------------------------------会员中心导航结束----------------------------------------
		
		
		'-------------------------------------------购物流程------------------------------------------------
		      Case "SHOPPINGCART" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,1)
			  Case "SHOPPINGPAYMENT" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,2)
			  Case "SHOPPINGPREVIEW" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,3)
			  Case "SHOPPINGSUCCESS" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,4)
			  Case Else GetLocation = GetIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
		   End Select
		 
		End Function
		
		'取得网站首页导航位置的函数
		Function GetIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
		   Dim str,Node
		   Select Case UCase(FCls.RefreshType)
		     case "INDEX" str="网站首页"
			 case "COMMENT"str="查看["&fcls.ItemTitle&"]的所有评论"
			 case "SEARCH" str="搜索结果"
			 case "SPACEINDEX" str="空间首页"
			 case "LINKINDEX" str="友情链接"
			 case "MAP" str="网站地图"
			 case "RSS" str="RSS订阅服务"
			 case "GUESTINDEX"
			  str="<a href='" & KS.GetClubListUrl(0) & "'>" & KS.Setting(61) & "</a>"  
			  If FCls.RefreshFolderID="0" Then
			  str=str & NaviStr & "首页"
			  Else
			   If KS.S("KeyWord")<>"" Then
			    str=str & NaviStr & "搜索结果 关键字:<span style=""color:red"">" & KS.CheckXSS(KS.S("KeyWord")) & "</span>"
			   Else
			    str=str & NaviStr & GetBoardNavigator(KS.ChkClng(FCls.RefreshFolderID),NaviStr)
			   End If
			  End If
			 case "GUESTWRITE" 
			  str="<a href='" & KS.GetClubListUrl(0) & "'>" & KS.Setting(61) & "</a>"
			  If KS.ChkClng(Request("bid"))<>0 Then
			    str=str & NaviStr & GetBoardNavigator(KS.ChkClng(Request("bid")),NaviStr)
			  End If
			  str=str &  Navistr &"发表" & KS.Setting(62)
			 case "GUESTDISPLAY" 
			  str="<a href='" & KS.GetClubListUrl(0) & "'>" & KS.Setting(61) & "</a>"
			  if FCls.RefreshFolderID<>0 then  str=str & NaviStr & GetBoardNavigator(KS.ChkClng(FCls.RefreshFolderID),NaviStr)
			  str=str & Navistr & "查看" & KS.Setting(62) & ""
			 case "CLUBSEARCH"
			  str="<a href='" & KS.GetClubListUrl(0) & "'>" & KS.Setting(61) & "</a>"& Navistr & "论坛搜索"
			 case "JOBINDEX" str="求职招聘"
			 case "RESUMESEARCH" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "查找人才" 
			 case "RESUMESCHOOL" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "按院校查看简历" 
			 case "SEARCHZW" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "查找职位" 
			 case "RESUMESC" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "简历收藏夹" 
			 case "COMPANYSHOW" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "查看公司详情" 
			 case "JOBREAD" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "查看职位详情" 
			 case "JOBAPPLY" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "申请职位" 
			 case "LETTER" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "求职信" 
			 case "ZHAOPIN" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "找人才" 
			 case "QIUZHI" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "找工作" 
			 case "JOBLTINDEX" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "猎头服务首页" 
			 case "JOBLTINTRO" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "猎头介绍" 
			 case "JOBLTNEWS" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "最新猎头职位"
			 case "JOBJZJOB" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "最新兼职职位"
			 case "JOBJZRESUME" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "最新兼职人才"
			 case "JOBJZINDEX" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "兼职天地首页" 
			 case "RESUMESEARCH" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "简历搜索列表" 
			 case "JOBSEARCH" str= "<a " & TitleCss & " href=""" & JLCls.GetJobHomeUrl & """" & KS.G_O_T_S(OpenType) & ">求职招聘</a>" & NaviStr & "职位搜索列表" 
			 case "ENTERPRISE" str="企业大全"
			 case "ENTERPRISELIST" str="<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">企业大全</a>" & NaviStr & FCls.LocationStr
			 case "ENTERPRISEPRO" str="产品库"
			 case "ENTERPRISEPROLIST" str="<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">产品库</a>" & NaviStr & FCls.LocationStr 
			 case "ENTERPRISEZS" str="装饰企业大全"
			 case "ENTERPRISELISTZS" str="<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">装饰企业大全</a>" & NaviStr & FCls.LocationStr
			 case else str=""
		   End Select
			  GetIndexLocation = StartTag &  str
		End Function
		
      '根据论坛版面ID取导航
		Function GetBoardNavigator(boardid,NaviStr)
			Dim Node,ParentID
			KS.LoadClubBoard
			Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			If Not Node Is Nothing Then
			      ParentID=KS.ChkClng(Node.SelectSingleNode("@parentid").Text)
				  if ParentID<>0 Then
					  If KS.ChkClng(session("clubnowboardpage"))>1 and UCase(FCls.RefreshType)="GUESTDISPLAY" Then
					  GetBoardNavigator=GetBoardNavigator(ParentID,NaviStr) & NaviStr & "<a href='" & KS.GetClubListUrlByPage(boardid,KS.ChkClng(session("clubnowboardpage"))) &"'>" & Node.SelectSingleNode("@boardname").Text &"</a>"
					  Else
					  GetBoardNavigator=GetBoardNavigator(ParentID,NaviStr) & NaviStr & "<a href='" & KS.GetClubListUrl(boardid) &"'>" & Node.SelectSingleNode("@boardname").Text &"</a>"
					  End If
				  else
				  GetBoardNavigator="<a href='" & KS.GetClubListUrl(boardid) &"'>" & Node.SelectSingleNode("@boardname").Text & "</a>"
				  End If
			End If
			 
		End Function		

		'取得更多空间导航位置的函数
		Function GetMoreSpaceLocation(StartTag, NaviStr, OpenType, TitleCss)
		   Dim MoreStr
		   Select Case UCase(FCls.RefreshType)
		    Case "MORESPACE":MoreStr="个人空间列表"
			Case "MOREFRESH":MoreStr="新鲜事"
			Case "MORELOG":MoreStr="日志列表"
			Case "MOREGROUP":MoreStr="圈子列表"
			Case "MOREXC":MoreStr="相册列表"
		   End Select 
			  GetMoreSpaceLocation = StartTag &  "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">空间首页</a>" & NaviStr  &MoreStr
		End Function

		'所有会员列表页
		Function GetUserListLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetUserListLocation = StartTag &  "所有注册会员列表"
		End Function
		'所有会员信息页
		Function GetUserInfoLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetUserInfoLocation = StartTag &  "<a href=""" & DomainStr & "User/UserList.asp"" " & KS.G_O_T_S(OpenType) & ">所有会员列表</a>"& NaviStr & "会员信息"
		End Function
		'会员中心
		Function GetMemberLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetMemberLocation = StartTag &  "会员中心"
		End Function
		'取得会员注册导航
		Function GetUserRegLocation(Step,StartTag, NaviStr, OpenType, TitleCss)
		  Select Case Step
		    Case 1 GetUserRegLocation = StartTag &  "服务条款和声明"
			Case 2 GetUserRegLocation = StartTag &  "填写注册表单"
			Case 3 GetUserRegLocation = StartTag &  "完成注册"
		  End Select
		End Function
		'取得专题首页导航位置的函数
		Function GetSpecialIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetSpecialIndexLocation = StartTag &  "专题首页"
		End Function
		'取得专题分类导航
		Function GetSpecialClassLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			 Dim SpecialIndexUrl,SpecialDir:SpecialDir = KS.Setting(95)
			 If KS.Setting(78)<>"0" Then SpecialIndexUrl=DomainStr & SpecialDir Else SpecialIndexUrl=DomainStr & "item/SpecialIndex.asp"
			 If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			 GetSpecialClassLocation = StartTag &  "<a href=""" & SpecialIndexUrl & """" & KS.G_O_T_S(OpenType) & ">专题首页</a>" & NaviStr & KS.C_C(RefreshFolderIDValue,1)  & KS.GetSpecialClass(RefreshFolderIDValue,"classname")
		
		End Function
		'取得专题页的位置导航" 
		Function GetSpecialLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			 Dim SpecialIndexUrl,SpecialDir:SpecialDir = KS.Setting(95)
			 If KS.Setting(78)<>"0" Then SpecialIndexUrl=DomainStr & SpecialDir Else SpecialIndexUrl=DomainStr & "item/SpecialIndex.asp"
			 If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			 Dim TempFolderStr
				  TempFolderStr = "<a " & TitleCss & " href=""" & KS.GetFolderSpecialPath(RefreshFolderIDValue, True) & """" & KS.G_O_T_S(OpenType) & ">" & KS.GetSpecialClass(RefreshFolderIDValue,"classname") & "</a>" & NaviStr
			 GetSpecialLocation = StartTag &  "<a href=""" & SpecialIndexUrl & """" & KS.G_O_T_S(OpenType) & ">专题首页</a>" & NaviStr & TempFolderStr & "浏览专题"
		End Function
		'取得栏目的位置导航
		Function GetFolderLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			Dim FolderNaviStr:FolderNaviStr = GetFolderNaviStr(NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			If FCls.BrandName<>"" Then
				  GetFolderLocation = StartTag & FolderNaviStr & NaviStr & FCls.BrandName
			Else
				If FCls.RefreshChannelHomeFlag = True Then
				  GetFolderLocation = StartTag & FolderNaviStr & NaviStr & "频道首页"
				Else
				  GetFolderLocation = StartTag & FolderNaviStr
				End If
		   End If
		End Function
		'取得信息页导航位置的函数
		Function GetContentLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue,ShowTitle)
		    Dim Str,FolderNaviStr:FolderNaviStr = GetFolderNaviStr(NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			Str = StartTag & FolderNaviStr & NaviStr
			If Cbool(ShowTitle)=true Then Str=Str & Fcls.ItemTitle Else Str=Str & "浏览"& KS.C_S(FCls.Channelid,3)
			GetContentLocation = Str
		End Function
		
		
        '购物流程
		Function GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,TypeID)
		   GetShoppingLocation = StartTag &  "商城中心" & NaviStr
		   Select Case TypeID
		    Case 1: GetShoppingLocation=GetShoppingLocation & "我的购物车"
			Case 2: GetShoppingLocation=GetShoppingLocation & "收银台"
			Case 3: GetShoppingLocation=GetShoppingLocation & "预览订单并确认"
			Case 4: GetShoppingLocation=GetShoppingLocation & "订单提交成功"
		   End Select
		End Function
         
		'******************************************************************************************************
		'函数名：GetFolderNameStr
		'作  用：返回栏目顺序列表
		'参  数：NaviStr--链接字符串,RefreshFolderIDValue--栏目ID, OpenType---新窗口打开, TitleCss---名称样式
		'返回值：形如: 科汛网络 >> 产品列表 >> 科汛网站管理系统
		'******************************************************************************************************
		Function GetFolderNaviStr(NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			  Dim TSArr, I
			  TSArr = Split(KS.C_C(RefreshFolderIDValue,8), ",")
			  For I = 0 To UBound(TSArr) - 1
			      If I=0 Then
					GetFolderNaviStr =  "<a " & TitleCss & " href=""" & KS.GetFolderPath(TSArr(I)) & """" & KS.G_O_T_S(OpenType) & ">" & KS.C_C(TSArr(I),1) & "</a>"
				  Else
					GetFolderNaviStr = GetFolderNaviStr & NaviStr & "<a " & TitleCss & " href=""" & KS.GetFolderPath(TSArr(I)) & """" & KS.G_O_T_S(OpenType) & ">" & KS.C_C(TSArr(I),1) & "</a>"
				 End If
			  Next
		End Function

		
		Function GetLocationNav(NavType, Nav)
			If CStr(NavType) = "0" Then
			  If Nav = "" Then
			   GetLocationNav = " >> "
			  Else
			   GetLocationNav = Nav
			  End If
			Else
			  If Nav = "" Then
				GetLocationNav = " >> "
			  Else
				If Left(Nav, 1) = "/" Or Left(Nav, 1) = "\" Then Nav = Right(Nav, Len(Nav) - 1)
				GetLocationNav = "<img src=""" & DomainStr & Nav & """ border=""0"" align=""absmiddle"" />"
			  End If
			End If
		End Function

End Class
%> 
<!--#include file="Kesion.Thumbs.asp"-->
<!--#include file="Kesion.TranPinYinCls.asp"-->
<!--#include file="Kesion.VersionCls.asp"-->
<!--#include file="EditorAPI.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Const ClassField="ID,FolderName,Folder,ClassPurview,FolderDomain,TemplateID,ClassBasicInfo,ClassDefineContent,TS,ClassID,Tj,DefaultDividePercent,ChannelID,TN,ClassType,FolderOrder,AdminPurview,AllowArrGroupID,CommentTF,Child,PubTf,MailTF,WapTemplateID,FnameType,FsoType,FilterTF,MapTF,TopFlag,firstAlphabet"           rem 定义载入缓存的栏目字段
Const ModelField="ChannelID,ChannelName,ChannelTable,ItemName,ItemUnit,FieldBit,BasicType,FsoHtmlTF,FsoFolder,RefreshFlag,ModelEname,MaxPerPage,VerificCommentTF,CommentVF,CommentLen,CommentTemplate,UserSelectFilesTF,InfoVerificTF,UserAddMoney,UserAddPoint,UserAddScore,ChannelStatus,CollectTF,UpFilesTF,UpFilesDir,UpFilesSize,UserUpFilesTF,UserUpFilesDir,AllowUpPhotoType,AllowUpFlashType,AllowUpMediaType,AllowUpRealType,AllowUpOtherType,SearchTemplate,ChargeType,FsoListNum,UserTF,DiggByVisitor,DiggByIP,DiggRepeat,DiggPerTimes,UserClassStyle,UserEditTF,FsoContentRule,FsoClassListRule,FsoClassPreTag,ThumbnailsConfig,LatestNewDay,StaticTF,PubTimeLimit,AnnexPoint,ModelIco,ModelShortName,WapSwitch"  rem 定义载入缓存的模型字段

Class PublicCls
		Public SiteSN,Version
		Public Setting,TbSetting,SSetting,JSetting,ASetting,WSetting
	    Private Sub Class_Initialize()
		if Not Response.IsClientConnected then die ""
		  Call InitializeKesionCMS
        End Sub
	    Private Sub Class_Terminate()

	    End Sub
	 
	 Function InitialObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set InitialObject=CreateObject(str)
	 End Function
	 '*******************************************************************************************************************
	 '函数名：InitializeKesionCMS
	 '作  用: 加载KesionCMS的必要参数
	 '备  注：以下参数请不要更改。否则系统可能无法正常运行
	 '*******************************************************************************************************************
	 Public Function InitializeKesionCMS()
		Dim KCls:Set KCls=New KesionCls
		SiteSN =KCls.SiteSN 
		Version = KCls.KSVer
        Set KCls=Nothing
		Call InitialConfig()
	 End Function
	 
	 '*********************************************************************************************************
		'函数名：FSOSaveFile
		'作  用：生成文件
		'参  数： Content内容,路径 注意虚拟目录
	'*********************************************************************************************************
	Public Sub WriteLog(Content)
		    Dim FileName:FileName="/log.txt"
			dim stm:set stm=server.CreateObject("adodb.stream")
			dim oldText
			dim fs:Set fs=Server.CreateObject("Scripting.FileSystemObject")
			If fs.FileExists(server.MapPath(FileName)) Then   '判断文件是否存在      
			 oldText=ReadFromFile(FileName)
			end if
			set fs=nothing
			dim writeText:writetext=content
			if Not IsNul(oldText) Then
			 writeText=oldText & vbcrlf & content
			end If
			stm.Type=2 '以文本模式读取
			stm.mode=3
			stm.charset="utf-8"
			stm.open
			stm.WriteText writeText
			stm.SaveToFile server.MapPath(FileName),2 
			stm.flush
			stm.Close
			set stm=nothing
		End Sub
		
    '当前管理的后台风格皮肤		
	Public Function GetAdminSkinID()
	    dim adminName:adminName=C("AdminName")
		 Dim Doc :set Doc = InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(Setting(3)&"Config/adminconfig.xml"))
			Dim Node:Set Node=Doc.documentElement.selectSingleNode("/admin/item[@name='" &adminName & "']")
			 if not node is nothing then  
			  GetAdminSkinID=Node.text
			 Else
			   GetAdminSkinID=1
			 End If
	End Function

	'获得当前版本号
	Function GetVersion()
		Dim Doc:set Doc = CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Doc.async = false
		Doc.setProperty "ServerHTTPRequest", true 
		Doc.load(Server.MapPath(Setting(3) &"config/version.xml"))
		if Doc.readystate=4 and Doc.parseError.errorCode=0 Then 
		Dim Node:Set Node= Doc.documentElement.selectSingleNode("//kesioncms/version")
		If Not Node Is Nothing Then GetVersion=Node.text Else GetVersion="1.0"
		end if
	End Function

	'不提示,批量清除缓存,参数 PreCacheName-前段匹配
	Public Sub DelCaches(PreCacheName)
	    Dim i
		Dim CacheList:CacheList=split(GetCacheList(PreCacheName),",")
		If UBound(CacheList)>1 Then
			For i=0 to UBound(CacheList)-1
				DelCahe CacheList(i)
			Next
		End IF
	End Sub
	'取得缓存列表 参数 PreCacheName-前段匹配
	Public Function GetCacheList(PreCacheName)
		Dim Cacheobj
		For Each Cacheobj in Application.Contents
		If CStr(Left(Cacheobj,Len(PreCacheName)))=CStr(PreCacheName) Then GetCacheList=GetCacheList&Cacheobj&","
		Next
	End Function
	'清除缓存,参数 MyCaheName-缓存名称
	Public Sub DelCahe(MyCaheName)
	    Application(MyCaheName)=""
	End Sub
	
	'发送手机验证码,返回成功发送的条件
	Function SendMobileMsg(Mobile,ByVal Content)
	 ' on error resume next
	
	  If Setting(157)="0" Then SendMobileMsg=0 : Exit Function
	  if ChkClng(Setting(133))=0 then   'GBK编码
	    Session.CodePage=936
        Response.Charset="gb2312"
	  else
	    Session.CodePage=65001
        Response.Charset="utf-8"
	  end if
	  Dim PostUrl:PostUrl=Setting(150)
	  Dim Sign:Sign=split(Setting(156)&"∮","∮")(0)
	  If Not IsNul(Sign) Then Content=Content & Sign
	  If IsNUL(PostUrl) Or IsNul(Setting(152)) Or IsNul(split(Setting(153)&"∮","∮")(0)) Then SendMobileMsg=0 : Exit Function
	  PostUrl=Replace(PostUrl,"{$user}",Setting(152))
	  dim passarr:passarr=split(Setting(153)&"∮","∮")
	  dim pass:pass=passarr(0)
	  dim passType:passType=passarr(1)
	  if passType="1" Then
		  pass=md5(pass,16)
	  ElseIF PassType="2" Then
		  pass=md5(pass,32)
	  END iF
	  Content=Replace(Content,"{$sitename}",Setting(0))
	  Content=Replace(Content,"{$domain}",Setting(2))

	  PostUrl=replace(PostUrl,"{$pass}",pass)
	  PostUrl=replace(PostUrl,"{$content}",server.URLEncode(Content))
	  PostUrl=replace(PostUrl,"{$mobile}",mobile)
	  Dim Rstr:Rstr=do_post(posturl,"gbk")
	  If Rstr="error" Then
	    SendMobileMsg=0 
	  Else
	    If Instr(lcase(Rstr),lcase(Setting(151)))<>0 Then
		  SendMobileMsg=1
		Else
	     SendMobileMsg=Rstr
		End if
	  End If
	End Function
	
	'发送手机短信验证码等接口调用
	 Function do_post(HttpUrl,CharsetCode)
		   on error resume next
		   Dim Http:Set Http = Server.CreateObject("MSXML2.ServerXMLHTTP") 
		   Http.Open "GET", HttpUrl, False
		   Http.Send
		   If Http.Readystate <> 4 Then
			  Set Http = Nothing
			  do_post="error"
			  Exit Function
		   End If
		   do_post = BytesToBstr(Http.ResponseBody, CharsetCode)
		   Set Http = Nothing
		   If Err.Number <> 0 Then
		      do_post="error"
			  Err.Clear
		   End If
	End Function
		'==================================================
		'函数名：BytesToBstr
		'作  用：将获取的源码转换为中文
		'参  数：Body ------要转换的变量
		'参  数：Cset ------要转换的类型
		'==================================================
	Function BytesToBstr(Body, Cset)
		   Dim Objstream
		   Set Objstream = Server.CreateObject("adodb.stream")
		   Objstream.Type = 1
		   Objstream.Mode = 3
		   Objstream.Open
		   Objstream.Write Body
		   Objstream.Position = 0
		   Objstream.Type = 2
		   Objstream.Charset = Cset
		   BytesToBstr = Objstream.ReadText
		   Objstream.Close
		   Set Objstream = Nothing
	End Function
	
	 '判断是否启用日文
     Function WithKorean()
	     Dim Str:Str=""
	    If DataBaseType=1 AND IsWithKorean=TRUE Then str= "N"
		WithKorean=str
	 End Function

	 Public Sub GetSetting()
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		    RSObj.Open "SELECT top 1 Setting,TbSetting,SpaceSetting,JobSetting,AskSetting,WapSetting from [KS_Config]",conn,1,1
		    Dim i,node,xml,j,DataArray,rs
			Set xml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			xml.appendChild(xml.createElement("xml"))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(1)
						For i=0 To UBound(DataArray,2)
							Set Node=xml.createNode(1,"config","")
							j=0
							For Each rs in RSObj.Fields
								node.attributes.setNamedItem(xml.createNode(2,LCase(rs.name),"")).text= Replace(DataArray(j,i),vbcrlf,"$br$")& ""
								j=j+1
							Next
							xml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
		   Set Application(SiteSN&"_Config")=Xml
		   RSObj.Close:Set RSObj=Nothing
	 End Sub

	 Public Sub InitialConfig()
		If not IsObject(Application(SiteSN&"_Config")) then  GetSetting
		Setting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@setting").text&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","$br$",vbcrlf),"^%^")
		TbSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@tbsetting").text&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","$br$",vbcrlf),"^%^")
        SSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@spacesetting").text&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","$br$",vbcrlf),"^%^")
		JSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@jobsetting").text&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","$br$",vbcrlf),"^%^")
		ASetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@asksetting").text&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","$br$",vbcrlf),"^%^")
		WSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@wapsetting").text&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","$br$",vbcrlf),"^%^")
	 End Sub
	
	 'xmlroot跟节点名称 row记录行节点名称
	 Public Function RecordsetToxml(RSObj,row,xmlroot)
	  Dim i,node,rs,j,DataArray
	  If xmlroot="" Then xmlroot="xml"
	  If row="" Then row="row"
	  Set RecordsetToxml=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	  RecordsetToxml.appendChild(RecordsetToxml.createElement(xmlroot))
	  If Not RSObj.EOF Then
	   DataArray=RSObj.GetRows(-1)
	   For i=0 To UBound(DataArray,2)
		Set Node=RecordsetToxml.createNode(1,row,"")
		j=0
		For Each rs in RSObj.Fields		   
		   node.attributes.setNamedItem(RecordsetToxml.createNode(2,"ks"&j,"")).text= DataArray(j,i)& ""
		   j=j+1
		Next
		RecordsetToxml.documentElement.appendChild(Node)
	   Next
	  End If
	  DataArray=Null
	 End Function
	 
	 'xmlroot跟节点名称 row记录行节点名称
	Public Function RsToxml(RSObj,row,xmlroot)
			Dim i,node,rs,j,DataArray,fieldName
			If xmlroot="" Then xmlroot="xml"
			If row="" Then row="row"
			Set RsToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			RsToxml.appendChild(RsToxml.createElement(xmlroot))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(-1)
						For i=0 To UBound(DataArray,2)
							Set Node=RsToxml.createNode(1,row,"")
							j=0
							For Each rs in RSObj.Fields
							    fieldName=LCase(rs.name)
								if fieldName="username" then
								node.attributes.setNamedItem(RsToxml.createNode(2,fieldName,"")).text= lcase(DataArray(j,i))& ""
								else
								node.attributes.setNamedItem(RsToxml.createNode(2,fieldName,"")).text= DataArray(j,i)& ""
								end if
								j=j+1
							Next
							RsToxml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
	End Function
	Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
				Dim i,node,rs,j
				If xmlroot="" Then xmlroot="xml"
				Set ArrayToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
				If row="" Then row="row"
				For i=0 To UBound(DataArray,2)
					Set Node=ArrayToxml.createNode(1,row,"")
					j=0
					For Each rs in Recordset.Fields
							 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
							 j=j+1
					Next
					ArrayToxml.documentElement.appendChild(Node)
				Next
		End Function
	 
	 Public Function LoadChannelConfig()
	 Application.Lock
	 Dim RS:Set Rs=conn.execute("select " & ModelField &" From KS_Channel Order by ChannelID")
	 Set Application(SiteSN&"_ChannelConfig")=RecordsetToxml(rs,"channel","ChannelConfig")
	 Set Rs=Nothing
	 Application.unLock
	 End Function
	 
	 Function C_S(sChannelID,FieldID)
	  If IsNul(sChannelID) Then Exit Function
	  If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig()
	  Dim Node:Set Node=Application(SiteSN&"_ChannelConfig").documentElement.selectSingleNode("channel[@ks0=" & sChannelID & "]/@ks" & FieldID & "")
	  If Not Node Is Nothing  Then C_S = Node.Text Else C_S=0
	  Set Node = Nothing
	 End Function
	 
	 Function M_C(ChannelId,n)
	   M_C=Split(C_S(ChannelID,46)&"||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||","|")(n)
	 End Function
	 
	 Function GetClassName(ChannelID)
	   Dim ClassName:ClassName=M_C(ChannelID,26)
	   If IsNul(ClassName) Then ClassName="栏目"
	   GetClassName=ClassName
	 End Function
	 
	 Public Function LoadClassConfig()
		If not IsObject(Application(SiteSN&"_class")) Then
		 Application.Lock
		 Dim RS:Set Rs=conn.execute("select " & ClassField & " From KS_Class Order by root,folderorder")
		 Set Application(SiteSN&"_class")=RecordsetToxml(rs,"class","classConfig")
		 Set Rs=Nothing
		 Application.unLock
	   End If
	 End Function

	 '栏目ID得到栏目配置信息
	 Function C_C(ClassID,FieldID)
	   If ClassID="" Or IsNull(ClassID) Then Exit Function
	   LoadClassConfig()
	   Dim Node:Set Node=Application(SiteSN&"_class").documentElement.selectSingleNode("class[@ks0=" & classID & "]/@ks" & FieldID & "")
	   If Not Node Is Nothing Then C_C=Node.text
	   Set Node=Nothing
	 End Function
	 '栏目classID得到栏目配置信息
	 Function C_CID(ClassID,FieldID)
	   If ClassID="" Or IsNull(ClassID) Then Exit Function
	   LoadClassConfig()
	   Dim Node:Set Node=Application(SiteSN&"_class").documentElement.selectSingleNode("class[@ks9=" & classID & "]/@ks" & FieldID & "")
	   If Not Node Is Nothing Then C_CID=Node.text
	   Set Node=Nothing
	 End Function
	 
	
	 '加载品牌缓存
	 Sub LoadBrandCache()
	   If Not IsObject(Application(SiteSN&"_ClassBrand")) Then
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select id,BrandName,BrandEname,PhotoUrl From KS_ClassBrand Order by ID")
		 Set Application(SiteSN&"_ClassBrand")=RsToxml(rs,"row","classbrand")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 Function C_B(BrandID,FieldName)
	   If BrandID="" Or IsNull(BrandID) Then Exit Function
	   LoadBrandCache()
	   Dim Node:Set Node=Application(SiteSN&"_classbrand").documentElement.selectSingleNode("row[@id=" & BrandID & "]/@" & LCase(FieldName) & "")
	   If Not Node Is Nothing Then C_B=Node.text
	   Set Node=Nothing
	 End Function
	 
	 '加载用户组缓存
	 Sub LoadUserGroup()
	   If Not IsObject(Application(SiteSN&"_UserGroup")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select id,groupname,powerlist,descript,usertype,formid,templatefile,showonreg,ChargeType,GroupPoint,GroupSetting,tj,tn,ts From KS_UserGroup Where [Type]<2 Order by root,orderID")
		 Set Application(SiteSN&"_UserGroup")=RsToxml(rs,"row","groupConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 '获取用户组特殊权限
	 Function U_S(GroupID,i)
	   If IsNul(GroupID) Then U_S=0 : Exit Function
	   Dim GroupSetting:GroupSetting=U_G(GroupID,"GroupSetting") &",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
	   Dim GroupSetArr:GroupSetArr=Split(GroupSetting,",")
	   U_S=GroupSetArr(i)
	 End Function
	 Function U_G(GroupID,FieldName)
	   If IsNul(GroupID) Then Exit Function
	   LoadUserGroup
	   Dim Node:Set Node=Application(SiteSN&"_UserGroup").DocumentElement.selectSingleNode("row[@id=" & GroupID & "]/@" & Lcase(FieldName))
	   If Not Node Is Nothing Then U_G=Node.text
	   Set Node=Nothing
	 End Function
	 
	 
	 
	'**************************************************
	'函数名：LoadUserGroupOption
	'作  用：加载用户组选项
	'参  数：ShowPub 不允许发布的栏目显示灰色
	'返回值：整棵树
	'**************************************************
	Public Function LoadUserGroupOption(ShowPub)
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr,nbsp
		LoadUserGroup()
		'If ShowPub=false Then Pstr=Pstr & " and @ks27=1"
		For Each Node In Application(SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
		  SpaceStr="" 
			  TJ=Node.SelectSingleNode("@tj").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "──" 
				 Next
			  End If
			  If ShowPub=true Then
			    If Node.SelectSingleNode("@showonreg").text="1" Then
			      TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@id").text & "'>" & SpaceStr & Node.SelectSingleNode("@groupname").text & " </option>"
				Else
			      TreeStr = TreeStr & "<OPTGROUP label=&nbsp;"& SpaceStr & Node.SelectSingleNode("@groupname").text & " </OPTGROUP>"
				End If
			  Else
			  TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@id").text & "'>" & SpaceStr & Node.SelectSingleNode("@groupname").text & " </option>"
			  End If
		Next
		LoadUserGroupOption=TreeStr
	End Function
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 '加载论坛/问答等级
	 Sub LoadAskGrade()
	   If Not IsObject(Application(SiteSN&"_AskGrade")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select gradeid,UserTitle,score,Ico,ClubPostNum,Color,TypeFlag,Special From KS_AskGrade Order by GradeID")
		 Set Application(SiteSN&"_AskGrade")=RsToxml(rs,"row","AskGrade")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 '取KS_AskGrade表的配置
	 Function A_G(GradeID,FieldName)
	   If IsNul(GradeID) Then Exit Function
	   LoadAskGrade
	   Dim Node:Set Node=Application(SiteSN&"_AskGrade").DocumentElement.selectSingleNode("row[@gradeid=" & GradeID & "]/@" & Lcase(FieldName))
	   If Not Node Is Nothing Then A_G=Node.text
	   Set Node=Nothing
	 End Function
	 
	 '加载Tags缓存
	 Sub LoadTagsCache()
	   If Not IsObject(Application(SiteSN&"_ClassTags")) Then
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select id,KeyText From KS_KeyWords Order by ID")
		 Set Application(SiteSN&"_ClassTags")=RsToxml(rs,"row","classtags")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 Function Tags(TagsName)
	   If TagsName="" Or IsNull(TagsName) Then Tags=TagsName:Exit Function
	   LoadTagsCache()
	   Dim Node:Set Node=Application(SiteSN&"_classtags").documentElement.selectSingleNode("row[@keytext='" & TagsName & "']/@id")
	   If Not Node Is Nothing Then
	    Tags=Node.text
	   Else
	    Tags=TagsName
	   End If
	   Set Node=Nothing
	 End Function
	 '关键字URL
	 Function TagsUrl(ID,ChannelID,ClassID,Page)
	   If Setting(185)="1" Then
	   TagsUrl=GetDomain & "tags/list-" & server.URLEncode(Tags(ID)) & "-" & ChannelID & "-" & CLassID &"-" & Page &".html"
	   Else
	   TagsUrl=GetDomain & "plus/tags/index.asp?n=" & Tags(ID) & "&ChannelID=" & ChannelID & "&ClassID=" & CLassID &"&page=" & Page
	   End If
	 End Function
	 
	 
	 '加载留言版面缓存
	 Sub LoadClubBoard()
	   If Not IsObject(Application(SiteSN&"_ClubBoard")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select [id],[boardname],[note],[master],[todaynum],[postnum],[topicnum],[parentid],[LastPost],[BoardRules],[Settings] From KS_GuestBoard Where Locked<>1 Order by orderid,ID")
		 Set Application(SiteSN&"_ClubBoard")=RsToxml(rs,"row","clubConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 Sub LoadClubBoardCategory()
	   If Not IsObject(Application(SiteSN&"_ClubBoardCategory")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select [categoryid],[categoryname],[BoardID],[ico] From KS_GuestCategory Where Status=1 Order by orderid,CategoryID")
		 Set Application(SiteSN&"_ClubBoardCategory")=RsToxml(rs,"row","boardcategory")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub

	Public Sub LoadCategoryList()
	  If Not IsObject(Application(SiteSN&"_askclasslist")) Then
		Dim Rs,SQL,TempXmlDoc
		Set Rs = Conn.Execute("SELECT classid,ClassName,Readme,rootid,depth,parentid,Parentstr,child,c_setting FROM KS_AskClass ORDER BY orders,classid")
		If Not (Rs.BOF And Rs.EOF) Then
			SQL=Rs.GetRows(-1)
			Set TempXmlDoc = ArrayToxml(SQL,Rs,"row","classlist")
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(TempXmlDoc) Then
			Application.Lock
				Set Application(SiteSN&"_askclasslist") = TempXmlDoc
			Application.unLock
		End If
	 End If
	End Sub
	
	'**************************************************
	'函数名：LoadClassOption
	'作  用：加载栏目选项
	'参  数：ChannelID-----当前模型ID,ShowPub 不允许发布的栏目显示灰色
	'返回值：整棵树
	'**************************************************
	Public Function LoadClassOption(ChannelID,ShowPub)
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr,nbsp
		LoadClassConfig()
		If ChannelID<>0 Then Pstr=" and @ks12=" & channelid & ""
		'If ShowPub=false Then Pstr=Pstr & " and @ks27=1"
		For Each Node In Application(SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
		  SpaceStr="" 
		  If (C("SuperTF")=1 or FoundInArr(Node.SelectSingleNode("@ks16").text,C("GroupID"),",") or Instr(C("ModelPower"),C_S(Node.SelectSingleNode("@ks12").text,10)&"1")>0) and (C_S(Node.SelectSingleNode("@ks12").text,21)=1 or Node.SelectSingleNode("@ks12").text=5) Then 
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "──" 
				 Next
			  End If
			  If ShowPub=true Then
			    If Node.SelectSingleNode("@ks20").text="1" Then
			      TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
				Else
			      TreeStr = TreeStr & "<OPTGROUP label='&nbsp;"& SpaceStr & Node.SelectSingleNode("@ks1").text & "'> </OPTGROUP>"
				End If
			  Else
			  TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
			  End If
		 End If 
		Next
		LoadClassOption=TreeStr
	End Function
	'取防注册机组的问题及答案组 Flag 162取问题组 163取答案组
	Function GetCurrQuestion(Flag)
		  Dim QuestionArr:QuestionArr=Split(Setting(Flag),"#####"&vbcrlf)
		  Dim CurrHour:CurrHour=Hour(now)
		  Dim CurrQuestion,N: N=Ubound(QuestionArr)
		  IF N=0 Then 
		    CurrQuestion=Setting(Flag)
		  ElseIf CurrHour<=N Then 
		    CurrQuestion=QuestionArr(CurrHour)
		  Else
		    CurrQuestion=QuestionArr(currhour mod n)
		  End If
		  GetCurrQuestion=CurrQuestion
    End Function
		
	Sub Echo(Str)
	  Response.Write Str
	End Sub
	
	Sub Die(Str)
	  Response.Write Str : Response.End
	End Sub
	
	Function IsNul(Str)
	  If replace(Str&""," ","")="" Or IsNull(Str) Then IsNul=True Else IsNul=false
	End Function
	 
	
	'===================获取插件状态=======================================================
	 '获得应用插件的状态,打开返回true 否则返回false ,参数：AppName-插件英文名称
	Function GetAppStatus(AppEname)
		Dim IsOpen:IsOpen=false
		Dim APPXML:set APPXML = InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		APPXML.async = false
		APPXML.setProperty "ServerHTTPRequest", true 
		APPXML.load(Server.MapPath(Setting(3) & "Config/AppSetting.xml"))
		If APPXML.readystate=4 and APPXML.parseError.errorCode=0 Then
             Dim AppNode:Set AppNode=APPXML.DocumentElement.SelectSingleNode("app[@name='" & Lcase(AppEname) &"']")	
		     If Not AppNode Is Nothing Then
			   If AppNode.SelectSingleNode("@status").text="1" Then IsOpen=true
			 End If
		End If
		GetAppStatus=IsOpen
	End Function
	
	'检查插件是否打开，未打开则返回网站首页,参数：AppName-插件英文名称
	Sub CheckAppStatusAndDie(AppEname)
	  If GetAppStatus(AppEname) = False Then Die "<script>alert('对不起，本频道未开启！');location.href='" & Setting(3) & "';</script>"
	End Sub
	'==========================================================================================

	 
	 '保存配置信息到config/config.txt
     Sub settingsave(no,val)
		  dim i,str,config:config=ReadFromFile(Setting(3) & "config/config.txt")
		  if config="" or instr(config,"♂")=0 then config=config &"♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂"
		  config=split(config,"♂")
		  for i=0 to ubound(config)
		    if i=0 then
			  if i=no then str=val else str=config(i)
			else
			  if i=no then str=str & "♂" & val else  str=str & "♂" & config(i)
			end if
		  next
		  Call WriteTOFile(Setting(3) & "config/config.txt",str)
	 End Sub
	 Function ReadSetting(no)
		  dim config:config=split(ReadFromFile(Setting(3) & "config/config.txt") &"♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂♂","♂")
		  ReadSetting=config(no)
	 End Function
	 
	 '生成地区缓存
	 Sub CreateAreaCache()
		 Dim XMLStr,RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Province Order By ParentID,Depth,OrderID,ID",conn,1,1
		 If Not RS.EOf Then
			XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
			XMLStr=XMLStr&"<area>" &vbcrlf
			Do While Not RS.Eof 
			 XMLStr=XMLStr & "<item id=""" & RS("ID") & """ city=""" & RS("City") & """ parentid=""" & RS("ParentID") &""" depth=""" & rs("depth") & """ filtertf=""" & rs("filtertf") &"""></item>" &vbcrlf
			 RS.MoveNext
			Loop
			XMLStr=XMLStr&"</area>" &vbcrlf
			Call WriteTOFile(Setting(3) & "config/area.xml",xmlstr)
		 End If
		 RS.CLose
		 Set RS=Nothing
	End Sub
	
	'生成商品品牌缓存
	Sub CreateBrandCache()
		 Dim XMLStr,RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select B.ID,R.ClassID,B.BrandName From KS_ClassBrand B inner join KS_ClassBrandR R On B.id=R.BrandID  order by B.orderid",conn,1,1
		 If Not RS.EOf Then
			XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
			XMLStr=XMLStr&"<brand>" &vbcrlf
			Do While Not RS.Eof 
			 XMLStr=XMLStr & "<item id=""" & RS("ID") & """ classid=""" & RS("classid") & """><brandname><![CDATA[ " & RS("BrandName") &"]]></brandname></item>" &vbcrlf
			 RS.MoveNext
			Loop
			XMLStr=XMLStr&"</brand>" &vbcrlf
			Call WriteTOFile(Setting(3) & "config/shopbrand.xml",xmlstr)
		 End If
		 RS.CLose
		 Set RS=Nothing
	End Sub
	 
	 
    Public Function Conn()
	  On Error Resume Next
	  Dim ConnObj:Set ConnObj=Server.CreateObject("ADODB.Connection")
	  ConnObj.Open ConnStr
	  Set Conn = ConnObj
	End Function
	'采集数据库连接
	Public Function ConnItem()
	  On Error Resume Next
	  Dim ConnObj:Set ConnObj=Server.CreateObject("ADODB.Connection")
	  ConnObj.Open CollcetConnStr
	  Set ConnItem = ConnObj
	End Function

	
	'***************************************************************************************************************
	'函数名：GetDomain
	'作  用：获取URL,包括虚拟目录 如http://www.kesion.com/ 或 http://www.kesion.com/Sys/  其中 Sys/为虚拟目录
	'参  数：  无
	'返回值：完整域名
	'***************************************************************************************************************
	Public Function GetDomain()
	    GetDomain = Trim(Setting(2) & Setting(3))
	End Function
	'**************************************************
	'函数名：GetChannelDomain
	'作  用：获取包含频道的完整Url
	'参  数：ChannelID频道ID
	'返回值：完整域名
	'**************************************************
	Public Function GetChannelDomain(ChannelID)
		GetChannelDomain=C_S(ChannelID,8)
		If Left(GetChannelDomain, 1) = "/" Then GetChannelDomain = Right(GetChannelDomain, Len(GetChannelDomain) - 1)
		GetChannelDomain = GetDomain() & GetChannelDomain
	End Function
	'**************************************************
	'函数名：GetAutoDoMain()
	'作  用：取得当前服务器IP 如：http://127.0.0.1
	'参  数：无
	'**************************************************
	Public Function GetAutoDomain()
		Dim TempPath
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetAutoDomain = Request.ServerVariables("SERVER_NAME")
		Else
			GetAutoDomain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		 If Instr(UCASE(GetAutoDomain),"/W3SVC")<>0 Then
			   GetAutoDomain=Left(GetAutoDomain,Instr(GetAutoDomain,"/W3SVC"))
		 End If
		 GetAutoDomain = "http://" & GetAutoDomain
	End Function
	
	'取得当前完整URL
	Public Function GetCurrentURL() 
			Dim Url 
			Dim ServerPort,ServerName,ScriptName,QueryString 
			ServerName = Request.ServerVariables("SERVER_NAME") 
			ServerPort = Request.ServerVariables("SERVER_PORT") 
			ScriptName = Request.ServerVariables("SCRIPT_NAME") 
			QueryString = Request.ServerVariables("QUERY_STRING") 
			Url="http://"&ServerName 
			If ServerPort <> "80" Then Url = Url & ":" & ServerPort 
			Url=Url&ScriptName 
			If QueryString <>"" Then Url=Url&"?"& QueryString 
			GetCurrentURL=Url 
	End Function

	
	Function CutFixContent(ByVal str, ByVal start, ByVal last, ByVal n)
		Dim strTemp
		On Error Resume Next
		If InStr(str, start) > 0 Then
			Select Case n
			Case 0  '左右都截取（都取前面）（去处关键字）
				strTemp = Right(str, Len(str) - InStr(str, start) - Len(start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) - 1)
			Case Else  '左右都截取（都取前面）（保留关键字）
				strTemp = Right(str, Len(str) - InStr(str, start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) + Len(last) - 1)
			End Select
		Else
			strTemp = ""
		End If
		CutFixContent = strTemp
	End Function
	
	'取得Tag之间的循环体
	Function GetTagLoop(ByVal Content)
			Dim regEx, Matches, Match, LoopStr
			Set regEx = New RegExp
			regEx.Pattern = "{Tag([\s\S]*?):(.+?)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				Content=Replace(Content,Match.Value,"")
				Content=Replace(Content,"{/Tag}","")
			Next
			GetTagLoop=Content
    End Function
	
	
	'==================================================
	'函数名：ScriptHtml
	'作  用：过滤html标记
	'参  数：ConStr ------ 要过滤的字符串
	'==================================================
	Function ScriptHtml(ByVal Constr, TagName, FType)
			Dim re
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			Select Case FType
			Case 1
			   re.Pattern = "<" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			Case 2
			   re.Pattern = "<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			Case 3
			   re.Pattern = "<" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			   re.Pattern = "</" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			End Select
			ScriptHtml = Constr
			Set re = Nothing
	End Function
	

	'*************************************************************************
	'函数名：gotTopic
	'作  用：截字符串，汉字一个算两个字符，英文算一个字符
	'参  数：str   ----原字符串
	'       strlen ----截取长度
	'返回值：截取后的字符串
	'*************************************************************************
	Public Function GotTopic(ByVal Str, ByVal strlen)
		If Str = "" OR IsNull(Str) Then GotTopic = "":Exit Function
		If strlen=0 Then GotTopic=Str:Exit Function
		Dim l, T, c, I, strTemp
		Str = Replace(Replace(Replace(Replace(Str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
		l = Len(Str)
		T = 0
		strTemp = Str
		strlen = CLng(strlen)
		For I = 1 To l
			c = Abs(Ascw(Mid(Str, I, 1)))
			If c > 255 Then
				T = T + 2
			Else
				T = T + 1
			End If
			If T >= strlen Then
				strTemp = Left(Str, I)
				Exit For
			End If
		Next
		If strTemp <> Str Then	strTemp = strTemp
		GotTopic = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
	End Function
	
	'**************************************************
	'函数名：ListTitle
	'作  用：取标题
	'参  数：TitleStr 标题, TitleNum 取字符数
	'返回值：将标题分解成两行
	'**************************************************
	Public Function ListTitle(TitleStr, TitleNum)
		  Dim LeftStr, RightStr
			ListTitle = Trim(GotTopic(Trim(TitleStr), TitleNum))
			If Len(ListTitle) > CInt(TitleNum / 2) Then
			  LeftStr = GotTopic(ListTitle, CInt(TitleNum / 2))
			  RightStr = Mid(ListTitle, Len(LeftStr) + 1)
			  ListTitle = LeftStr & "<br>" & RightStr
			End If
	 End Function
	Function ListTitle1(TitleStr, TitleNum)
		   Dim ClsTitleStr, ClsTitleNum, I, J, ClsTempNum, k, ClsTitleStrResult, LeftStr, RightStr
			   ClsTitleNum = CInt(TitleNum)
			   ClsTempNum = Len(CStr(TitleStr))
			   If ClsTitleNum > ClsTempNum Then
				   ClsTitleNum = ClsTempNum
			   End If
			   ClsTitleStr = Left(CStr(TitleStr), ClsTitleNum)
			   Dim TempStr
			   For I = 1 To ClsTitleNum - 1
				   TempStr = TempStr & Mid(ClsTitleStr, I, 1) & "<br />"
			   Next
			   TempStr = TempStr & Right(ClsTitleStr, 1)
			   ListTitle1 = TempStr
	End Function

	'**************************************************
	'函数名：GetIP
	'作  用：取得正确的IP
	'返回值：IP字符串
	'**************************************************
	Public Function GetIP() 
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		End If 
		getIP = Checkstr(Trim(Mid(strIPAddr, 1, 30)))
	End Function
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function
	'================================================
	'函数名：URLDecode
	'作  用：URL解码
	'================================================
	Function URLDecode(ByVal urlcode)
		Dim start,final,length,char,i,butf8,pass
		Dim leftstr,rightstr,finalstr
		Dim b0,b1,bx,blength,position,u,utf8
		On Error Resume Next
	
		b0 = Array(192,224,240,248,252,254)
		urlcode = Replace(urlcode,"+"," ")
		pass = 0
		utf8 = -1
	
		length = Len(urlcode) : start = InStr(urlcode,"%") : final = InStrRev(urlcode,"%")
		If start = 0 Or length < 3 Then URLDecode = urlcode : Exit Function
		leftstr = Left(urlcode,start - 1) : rightstr = Right(urlcode,length - 2 - final)
	
		For i = start To final
			char = Mid(urlcode,i,1)
			If char = "%" Then
				bx = URLDecode_Hex(Mid(urlcode,i + 1,2))
				If bx > 31 And bx < 128 Then
					i = i + 2
					finalstr = finalstr & ChrW(bx)
				ElseIf bx > 127 Then
					i = i + 2
					If utf8 < 0 Then
						butf8 = 1 : blength = -1 : b1 = bx
						For position = 4 To 0 Step -1
							If b1 >= b0(position) And b1 < b0(position + 1) Then
								blength = position
								Exit For
							End If
						Next
						If blength > -1 Then
							For position = 0 To blength
								b1 = URLDecode_Hex(Mid(urlcode,i + position * 3 + 2,2))
								If b1 < 128 Or b1 > 191 Then butf8 = 0 : Exit For
							Next
						Else
							butf8 = 0
						End If
						If butf8 = 1 And blength = 0 Then butf8 = -2
						If butf8 > -1 And utf8 = -2 Then i = start - 1 : finalstr = "" : pass = 1
						utf8 = butf8
					End If
					If pass = 0 Then
						If utf8 = 1 Then
							b1 = bx : u = 0 : blength = -1
							For position = 4 To 0 Step -1
								If b1 >= b0(position) And b1 < b0(position + 1) Then
									blength = position
									b1 = (b1 xOr b0(position)) * 64 ^ (position + 1)
									Exit For
								End If
							Next
							If blength > -1 Then
								For position = 0 To blength
									bx = URLDecode_Hex(Mid(urlcode,i + 2,2)) : i = i + 3
									If bx < 128 Or bx > 191 Then u = 0 : Exit For
									u = u + (bx And 63) * 64 ^ (blength - position)
								Next
								If u > 0 Then finalstr = finalstr & ChrW(b1 + u)
							End If
						Else
							b1 = bx * &h100 : u = 0
							bx = URLDecode_Hex(Mid(urlcode,i + 2,2))
							If bx > 0 Then
								u = b1 + bx
								i = i + 3
							Else
								If Left(urlcode,1) = "%" Then
									u = b1 + Asc(Mid(urlcode,i + 3,1))
									i = i + 2
								Else
									u = b1 + Asc(Mid(urlcode,i + 1,1))
									i = i + 1
								End If
							End If
							finalstr = finalstr & Chr(u)
						End If
					Else
						pass = 0
					End If
				End If
			Else
				finalstr = finalstr & char
			End If
		Next
		URLDecode = leftstr & finalstr & rightstr
	End Function
	
Function URLDecode_Hex(ByVal h)
	On Error Resume Next
	h = "&h" & Trim(h) : URLDecode_Hex = -1
	If Len(h) <> 4 Then Exit Function
	If isNumeric(h) Then URLDecode_Hex = cInt(h)
End Function
	'**************************************************
	'函数名：R
	'作  用：过滤非法的SQL字符
	'参  数：strChar-----要过滤的字符
	'返回值：过滤后的字符
	'**************************************************
	Public Function R(strChar)
		If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
		Dim strBadChar, arrBadChar, tempChar, I
		'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		arrBadChar = Split(strBadChar, ",")
		tempChar = strChar
		For I = 0 To UBound(arrBadChar)
			tempChar = Replace(tempChar, arrBadChar(I), "")
		Next
		tempChar = Replace(tempChar, "@@", "@")
		R = tempChar
	End Function
	'过滤xss
	Function CheckXSS(ByVal Str)
		If IsNull(Str) Then
			CheckXSS = ""
			Exit Function
		End If
		Str = Replace(Str, "&", "&amp;")
		Str = Replace(Str, "'", "&acute;")
		Str = Replace(Str, """", "&quot;")
		Str = Replace(Str, "<", "&lt;")
		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "/", "&#47;")
		Str = Replace(Str, "*", "&#42;")
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(w)(here)"
		Str = re.Replace(Str, "$1h&#101;re")
		re.Pattern = "(s)(elect)"
		Str = re.Replace(Str, "$1el&#101;ct")
		re.Pattern = "(i)(nsert)"
		Str = re.Replace(Str, "$1ns&#101;rt")
		re.Pattern = "(c)(reate)"
		Str = re.Replace(Str, "$1r&#101;ate")
		re.Pattern = "(d)(rop)"
		Str = re.Replace(Str, "$1ro&#112;")
		re.Pattern = "(a)(lter)"
		Str = re.Replace(Str, "$1lt&#101;r")
		re.Pattern = "(d)(elete)"
		Str = re.Replace(Str, "$1el&#101;te")
		re.Pattern = "(u)(pdate)"
		Str = re.Replace(Str, "$1p&#100;ate")
		re.Pattern = "(\s)(or)"
		Str = re.Replace(Str, "$1o&#114;")
		re.Pattern = "(\n)"
		Str = re.Replace(Str, "$1")
		'----------------------------------
			re.Pattern = "(java)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
			re.Pattern = "(j)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
			re.Pattern = "(vb)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
        '----------------------------------
        If Instr(Str, "expression") > 0 Then
                Str = Replace(Str, "expression", "e&#173;xpression", 1, -1, 0) '防止xss注入
        End If
		Set re = Nothing
		Checkxss = Str
   End Function
	
	Function FilterIDs(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs&"",",")
	For i=0 To Ubound(arrIds)
		If ChkClng(Trim(arrIDs(i)))<>0 Then
			strReturn=strReturn & "," & Int(arrIDs(i))
		End If
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterIDs=strReturn
	End Function
	'********************************************
	'函数名：IsValidEmail
	'作  用：检查Email地址合法性
	'参  数：email ----要检查的Email地址
	'返回值：True  ----Email地址合法
	'       False ----Email地址不合法
	'********************************************
	Public Function IsValidEmail(Email)
		Dim names, name, I, c
		IsValidEmail = True
		names = Split(Email, "@")
		If UBound(names) <> 1 Then IsValidEmail = False: Exit Function
		For Each name In names
			If Len(name) <= 0 Then IsValidEmail = False:Exit Function
			For I = 1 To Len(name)
				c = LCase(Mid(name, I, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then IsValidEmail = False:Exit Function
		   Next
		   If Left(name, 1) = "." Or Right(name, 1) = "." Then IsValidEmail = False:Exit Function
		Next
		If InStr(names(1), ".") <= 0 Then IsValidEmail = False:Exit Function
		I = Len(names(1)) - InStrRev(names(1), ".")
		If I <> 2 And I <> 3 Then IsValidEmail = False:Exit Function
		If InStr(Email, "..") > 0 Then IsValidEmail = False
	End Function
	'**************************************************
	'函数名：strLength
	'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
	'参  数：str  ----要求长度的字符串
	'返回值：字符串长度
	'**************************************************
	Public Function strLength(Str)
		On Error Resume Next
		Dim WINNT_CHINESE:WINNT_CHINESE = (Len("中国") = 2)
		If WINNT_CHINESE Then
			Dim l, T, c,I
			l = Len(Str)
			T = l
			For I = 1 To l
				c = Ascw(Mid(Str, I, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then
					T = T + 1
				End If
			Next
			strLength = T
		Else
			strLength = Len(Str)
		End If
		If Err.Number <> 0 Then Err.Clear
	End Function

	'**************************************************
	'函数名: GetFolderPath
	'功 能:取得目录Url
	'参 数: FolderID目录的ID
	'**************************************************
	Public Function GetFolderPath(FolderID)
			If Not IsObject(Application(SiteSN&"_classpath")) Then
		     Dim Folder,ClassPurview,ChannelFsoHtmlTF,Node,K,SQL,RS
			 Set  Application(SiteSN&"_classpath")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		     Application(SiteSN&"_classpath").appendChild( Application(SiteSN&"_classpath").createElement("xml"))
              Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select C.ClassID,C.ChannelID,TN,Folder,FolderDomain,ClassPurview,FsoHtmlTF,StaticTF,C.ID,ClassType,M.FsoClassListRule,M.FsoClassPreTag,FolderFsoIndex,StaticTF From KS_Class C inner join KS_Channel M On C.ChannelID=M.ChannelID Order BY FolderOrder", Conn, 1, 1
			  If RS.Eof And RS.Bof Then RS.Close:Set RS=Nothing:Exit Function
			  SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
			  For K=0 To Ubound(SQL,2)
					       ClassPurview=SQL(5,K)
						   ChannelFsoHtmlTF=SQL(6,K)
						   If SQL(9,K)="2" Then    '外部链接
						    GetFolderPath=SQL(3,K)
						   Else
							   If Not IsNul(SQL(4,K)) And SQL(2,K) = "0" Then
								   IF ClassPurview=2 Then
									 GetFolderPath= GetChannelNoHtmlUrl(GetDomain,SQL(7,K),SQL(0,K))
								   Else
									 GetFolderPath=Trim(SQL(4,K))
								   End If
							   ElseIf Not IsNul(SQL(4,K)) Then
								  Folder = Trim(SQL(3,K))
								  Folder = Right(Mid(Folder, InStr(Folder, "/")), Len(Mid(Folder, InStr(Folder, "/"))) - 1)
								   IF ClassPurview=2 Or ChannelFsoHtmlTF=0 Or ChannelFsoHtmlTF=3 Then
								      'GetFolderPath=Trim(SQL(4,K)) & "item/list.asp?id=" & sql(0,k)
									  GetFolderPath= GetChannelNoHtmlUrl(SQL(4,K),SQL(7,K),SQL(0,K))
								   Else
									 GetFolderPath= Trim(SQL(4,K)) & Folder
								   End If
							  Else
							       IF ClassPurview=2 Or ChannelFsoHtmlTF=0 Or ChannelFsoHtmlTF=2 Then
									 GetFolderPath= GetChannelNoHtmlUrl(GetDomain,SQL(7,K),SQL(0,K))
								   Else
									 	 GetFolderPath= GetChannelDomain(SQL(1,K)) 
										 If SQL(9,K)="3" Then
										  GetFolderPath= GetChannelDomain(SQL(1,K)) & SQL(3,K)
										 Else
										     Dim FsoClassPreTag:FsoClassPreTag=SQL(11,K)
											 Dim FolderEname:FolderEname=Split(Trim(SQL(3,K)),"/")(Ubound(Split(Trim(SQL(3,K)),"/"))-1)
											 FsoClassPreTag=Replace(FsoClassPreTag,"{$TopClassEname}",Split(Trim(SQL(3,K)),"/")(0)&"")
											 FsoClassPreTag=Replace(FsoClassPreTag,"{$ClassEname}",FolderEName)
											 FsoClassPreTag=Replace(FsoClassPreTag,"{$ClassID}",SQL(0,K))
											 FsoClassPreTag=Replace(FsoClassPreTag,"{$BigClassID}",SQL(8,K))
											 Select Case SQL(10,K)
											   Case "1":GetFolderPath= GetChannelDomain(SQL(1,K)) & SQL(3,K)
											   Case "2":GetFolderPath= GetChannelDomain(SQL(1,K)) & FsoClassPreTag &"_" & SQL(0,K) &Mid(Trim(SQL(12,K)), InStrRev(Trim(SQL(12,K)), ".")) '分离出扩展名
											   Case "3":
												 GetFolderPath= GetChannelDomain(SQL(1,K)) & Split(SQL(3,K),"/")(0) & "/"
												 If SQL(2,K) <> "0" Then GetFolderPath= GetFolderPath & FsoClassPreTag &"_" & SQL(0,K) &Mid(Trim(SQL(12,K)), InStrRev(Trim(SQL(12,K)), ".")) '分离出扩展名
											   Case "4":GetFolderPath=GetChannelDomain(SQL(1,K)) & FsoClassPreTag &Mid(Trim(SQL(12,K)), InStrRev(Trim(SQL(12,K)), ".")) '分离出扩展名
											   
											 End Select
                                         End If
								   End If
							  End If
						 End If
		            Set Node=Application(SiteSN&"_classpath").documentElement.appendChild(Application(SiteSN&"_classpath").createNode(1,"classpath",""))
			        Node.attributes.setNamedItem(Application(SiteSN&"_classpath").createNode(2,"classid","")).text=SQL(8,K)
			        Node.text=GetFolderPath
               Next			
     End If
	 Dim NodeText:Set NodeText=Application(SiteSN&"_classpath").documentElement.selectSingleNode("classpath[@classid='" & FolderID & "']")
	 If Not NodeText Is Nothing Then GetFolderPath=NodeText.text
	End Function
	'************************************************************************
	'函数名: GetClassNP
	'功 能: 取得目录名称并加上链接
	'参 数: ClassID目录的ID	          
	'*************************************************************************
	Function GetClassNP(ClassID)
		If Not IsObject(Application(SiteSN&"_classnamepath")) Then
		    Dim Folder,ClassPurview,ChannelFsoHtmlTF,Node,K,SQL,RS
			Dim OpenTypeStr:OpenTypeStr=" target=""_blank"""
			Set  Application(SiteSN&"_classnamepath")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		    Application(SiteSN&"_classnamepath").appendChild( Application(SiteSN&"_classnamepath").createElement("xml"))
              Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select ID,FolderName From KS_Class Order BY FolderOrder", Conn, 1, 1
			  If RS.Eof And RS.Bof Then RS.Close:Set RS=Nothing:Exit Function
			  SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
			  For K=0 To Ubound(SQL,2)
		            Set Node=Application(SiteSN&"_classnamepath").documentElement.appendChild(Application(SiteSN&"_classnamepath").createNode(1,"classnamepath",""))
			        Node.attributes.setNamedItem(Application(SiteSN&"_classnamepath").createNode(2,"classid","")).text=SQL(0,K)
			        Node.text="<a href=""" & GetFolderPath(SQL(0,K)) & """" & OpenTypeStr & ">" & Trim(SQL(1,K)) & "</a>"
              Next			
     End If
	 Dim NodeText:Set NodeText=Application(SiteSN&"_classnamepath").documentElement.selectSingleNode("classnamepath[@classid=" &ClassID & "]")
	 If Not NodeText Is Nothing Then GetClassNP=NodeText.text
	End Function
	
	'替换内容页生成规则
	 Function LoadFsoContentRule(ChannelID,ClassID,InfoID,AddDate)
	    on error resume next
		Dim FsoContentRule:FsoContentRule=C_S(ChannelID,43)
		if not isDate(addDate) then AddDate=Now
        FsoContentRule=Replace(FsoContentRule,"{$ChannelEname}",Split(C_C(ClassID,2),"/")(0))
        FsoContentRule=Replace(FsoContentRule,"{$ClassDir}",C_C(ClassID,2))
        FsoContentRule=Replace(FsoContentRule,"{$ClassID}",C_C(ClassID,9))
        FsoContentRule=Replace(FsoContentRule,"{$InfoID}",InfoID)
        FsoContentRule=Replace(FsoContentRule,"{$Year}",year(AddDate))
        FsoContentRule=Replace(FsoContentRule,"{$Month}",month(AddDate))
        FsoContentRule=Replace(FsoContentRule,"{$Day}",day(AddDate))
        FsoContentRule=Replace(FsoContentRule,"{$ClassEname}",Split(C_C(ClassID,2), "/")(C_C(ClassID,10)- 1))
		FsoContentRule=Replace(Setting(3) & C_S(ChannelID,8),"//","/") & FsoContentRule
		LoadFsoContentRule=FsoContentRule
	 End Function
     Function LoadInfoUrl(ChannelID,ClassID,Fname,InfoID,AddDate)
	   If C_C(ClassID,4)<>"" Then
	    LoadInfoUrl=GetFolderPath(ClassID) & Fname
	   Else
	    LoadInfoUrl=Setting(2) & LoadFsoContentRule(ChannelID,ClassID,InfoID,AddDate) & Fname
	   End If
	 End Function
		'----------------------------------------------------------------------------------------------------------------------
		'函数名: GetSpecialPath
		'功 能: 取得专题目录Url
		'参 数: SpecialrRS
		'-----------------------------------------------------------------------------------------------------------------------
		Public Function GetSpecialPath(SpecialID,SpecialEname,FsoSpecialIndex)
		      Dim SpecialDir:SpecialDir = Setting(95)
			  If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			  If Setting(78)="0" Then
				  GetSpecialPath=GetDomain & "item/Special.asp?ID=" & SpecialID
			  Else
			     if fsoSpecialIndex="{$id}" Then
				  GetSpecialPath = GetDomain & SpecialDir & replace(FsoSpecialIndex,"{$id}",specialID&".html")
				 Else
				  GetSpecialPath = GetDomain & SpecialDir & SpecialEname & "/" & FsoSpecialIndex
				 End If
              End iF
		End Function
		'----------------------------------------------------------------------------------------------------------------------
		'函数名: GetFolderSpecialPath
		'功 能: 取得栏目专题汇总Url
		'参 数: ClassID目录的ID,FullPathFlag是否完整路径(取栏目首页与否),包括专题首页
		'-----------------------------------------------------------------------------------------------------------------------
		Function GetFolderSpecialPath(ClassID, FullPathFlag)
		   Dim SpecialDir:SpecialDir =Setting(95)
		    If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
		     IF Setting(78)="0" Then
			     GetFolderSpecialPath = GetDomain &"item/SpecialList.asp?ID="&ClassID
			 Else
			  Dim RS:Set RS=Conn.Execute("Select ClassEname,FsoIndex From KS_SpecialClass Where ClassID=" & ChkClng(ClassID))
			  If RS.Eof Then
			   GetFolderSpecialPath = GetDomain &"item/SpecialList.asp?ID="&ClassID
			  Else
			    GetFolderSpecialPath = GetDomain & SpecialDir & RS(0) & "/"
			    If FullPathFlag = True Then
			     GetFolderSpecialPath=GetFolderSpecialPath & RS(1)
			    End If
              	RS.Close:Set RS = Nothing
			 End IF
			End If
		End Function
		'取得栏目的链接URL
		Public Function GetChannelNoHtmlUrl(Url,StaticTF,ClassID)
		     If StaticTF=0 Then
		      GetChannelNoHtmlUrl=Url &"Item/list.asp?id=" & ClassID
			 ElseIf StaticTF=2 Then
		      GetChannelNoHtmlUrl=Url & GCls.StaticPreList & "-" & ClassID & GCls.StaticExtension
			 Else
		      GetChannelNoHtmlUrl=Url & "?" & GCls.StaticPreList & "-" & ClassID & GCls.StaticExtension
			 End If
		End Function
		
		'模型内容页URL(PC)
		Public Function GetItemURL(ByVal ChannelID,ByVal Tid,ByVal InfoID,ByVal Fname,AddDate)
		  IF Not Isnumeric(ChannelID) Then GetItemURL="#":Exit Function
		  If  C_S(ChannelID,7)=0 Then 
		        dim preDomain:PreDomain=GetDomain
				if Not IsNul(C_C(Tid,4)) then PreDomain=C_C(Tid,4)
		        if C_S(ChannelID,48)=0 Then
				 GetItemURL=PreDomain & "Item/Show.asp?m=" & ChannelID & "&d=" &InfoID
				ElseIf C_S(ChannelID,48)=2 Then
				 GetItemURL=PreDomain & GCls.StaticPreContent & "-" & InfoID & "-"& ChannelID & GCls.StaticExtension
				Else
				 GetItemURL=PreDomain & "?" & GCls.StaticPreContent & "-" & InfoID & "-"& ChannelID & GCls.StaticExtension
				End If
		  Else
				GetItemURL=LoadInfoUrl(ChannelID,TID,Fname,InfoID,AddDate)
		  End If
		End Function
		
		'模型内容页URL(手机版)
		Public Function Get3GItemURL(ByVal ChannelID,ByVal Tid,ByVal InfoID,ByVal Fname)
		  Dim PreDomain
		  If IsNul(WSetting(1)) Then
		     PreDomain=GetDomain
		  Else
		     PreDomain="http://" & WSetting(1) &"/"
		  End If
		  If ChkClng(M_C(ChannelID,28))=1  Or ChkClng(M_C(ChannelID,28))=2 Then
		   Dim FsoDir:FsoDir=WSetting(10)
		   If left(FsoDir,1)="/" Then FsoDir=Right(FsoDir,len(FsoDir)-1)
		   Get3GItemURL=PreDomain & FsoDir & C_S(ChannelID,10) & "/" & Fname
		  Else
		   Get3GItemURL=PreDomain & Wsetting(4) &"/show.asp?m=" & channelid &"&d=" & InfoID
		  End If
		End Function
		
		
		'根据ID取空间Url
		Public Function GetSpaceUrl(ByVal UserID)
		  If SSetting(21)="1" Then
		  GetSpaceUrl=GetDomain & "space/" & UserID
		  Else
		  GetSpaceUrl=GetDomain & "space/?" & UserID
		  End If
		End Function
		'帖子版面URL
		Public Function GetClubListUrl(ByVal Id)
			If ID=0 Then   '首页
			  If Not IsNul(Setting(69)) Then 
				  GetClubListUrl="http://" & Setting(69) & "/"
			  Else
				  GetClubListUrl=GetDomain & Setting(66) & "/index.asp"
			  End If
			Else
			  If Not IsNul(Setting(69)) Then
				 If Setting(70)="1" Then
					GetClubListUrl="http://" & Setting(69) &"/" & GCls.ClubPreList & "-" & id&GCls.StaticExtension
				 Else
					GetClubListUrl="http://" & Setting(69) &"/?" & GCls.ClubPreList & "-" & id&GCls.StaticExtension
				 End If
			  Else
				 If Setting(70)="1" Then
					GetClubListUrl=GetDomain & "" & GCls.ClubPreList & "-" & id&GCls.StaticExtension
				 Else
					GetClubListUrl=GetDomain & Setting(66) & "/index.asp?boardid=" & id
				 End If
			  End If
			 End If
		End Function
		Public Function GetClubListUrlByPage(ByVal Id ,ByVal Page)
			  If Not IsNul(Setting(69)) Then
				 If Setting(70)="1" Then
					GetClubListUrlByPage="http://" & Setting(69) &"/" & GCls.ClubPreList & "-" & id & "-" & Page &GCls.StaticExtension
				 Else
					GetClubListUrlByPage="http://" & Setting(69) &"/?" & GCls.ClubPreList & "-" & id & "-" & Page &GCls.StaticExtension
				 End If
			  Else
				 If Setting(70)="1" Then
					GetClubListUrlByPage=GetDomain & "" & GCls.ClubPreList & "-" & id& "-" & Page & GCls.StaticExtension
				 Else
					GetClubListUrlByPage=GetDomain & Setting(66) & "/index.asp?boardid=" & id & "&page=" & page
				 End If
			  End If
		End Function
		
		'帖子Url
		Public Function GetClubShowUrl(ByVal Id)
		     If Not IsNul(Setting(69)) Then
			    If Setting(70)="1" Then
			    GetClubShowUrl="http://" & Setting(69) &"/" & GCls.ClubPreContent & "-" & id&GCls.StaticExtension
				Else
			    GetClubShowUrl="http://" & Setting(69) &"/?" & GCls.ClubPreContent & "-" & id&GCls.StaticExtension
				End If
			 Else
			    If Setting(70)="1" Then
			    GetClubShowUrl=GetDomain & "" & GCls.ClubPreContent & "-" & id&GCls.StaticExtension
				Else
			    GetClubShowUrl=GetDomain & Setting(66) & "/display.asp?id=" & id
				End If
			 End If
		End Function
		'帖子Url带分页
		Public Function GetClubShowUrlPage(ByVal Id,ByVal Page)
		     If Not IsNul(Setting(69)) Then
			   If Setting(70)="1" Then
			    GetClubShowUrlPage="http://" & Setting(69) &"/" & GCls.ClubPreContent & "-" & id & "-" & page &GCls.StaticExtension
			   Else
			    GetClubShowUrlPage="http://" & Setting(69) &"/?" & GCls.ClubPreContent & "-" & id & "-" & page &GCls.StaticExtension
			   End If
			 Else
			   If Setting(70)="1" Then
			    GetClubShowUrlPage=GetDomain & "" & GCls.ClubPreContent & "-" & id & "-" & page &GCls.StaticExtension
			   Else
			    GetClubShowUrlPage=GetDomain & Setting(66) & "/display.asp?id=" & id &"&page=" & page
			   End If
			 End If
	End Function
	'论坛分页	
    Function GetClubPageList(BoardID,MaxPerPage,CurrentPage,TotalPut,ID,PreStatic)
	    dim totalPage,BStr
		If BoardID=0 Then Bstr="返回首页" Else Bstr="返回列表"
		GetClubPageList= "<table align=""right"" cellspacing=""0"" border=""0"" cellpadding=""0""><tr><td ><div class='fenye'><a href='" & GetClubListUrl(BoardID) &"' class='num'> << "&Bstr &"</a></div></td><td align=""right"">"
		If Not IsNul(Setting(69)) Then
			If totalput Mod MaxPerPage = 0 Then
				TotalPage=totalput\MaxPerPage
			Else
				TotalPage=totalput\MaxPerPage + 1
			End If
			If Setting(70)="1" Then
			GetClubPageList=GetClubPageList & "<div id='fenye' class='fenye'>" & GetStaticPageList ("http://" & Setting(69) & "/" & PreStatic & "-" & ID & "-",4,CurrentPage,TotalPage,true,GCls.StaticExtension) &"</div>"
			Else
			GetClubPageList=GetClubPageList & "<div id='fenye' class='fenye'>" & GetStaticPageList ("http://" & Setting(69) & "/?" & PreStatic & "-" & ID & "-",4,CurrentPage,TotalPage,true,GCls.StaticExtension) &"</div>"
			End if
		Else
		    If Setting(70)="1" Then
			If totalput Mod MaxPerPage = 0 Then
				TotalPage=totalput\MaxPerPage
			Else
				TotalPage=totalput\MaxPerPage + 1
			End If
			GetClubPageList=GetClubPageList & "<div id='fenye' class='fenye'>" & GetStaticPageList (PreStatic & "-" & ID & "-",4,CurrentPage,TotalPage,true,GCls.StaticExtension) &"</div>"
			Else
		    GetClubPageList=GetClubPageList &ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)
			End If
		End If
		GetClubPageList=GetClubPageList & "</td></tr></table>"
	 End Function	
	 
	 '显示伪静态分页
	 Function GetStaticPageList(FileName,PageStyle,CurrPage,TotalPage, ShowTurnToFlag,Extension)
			Dim PageStr, I, J, SelectStr
			 If ChkClng(PageStyle)=0 Then PageStyle=1
			 Select Case PageStyle
			  Case 1
			   If CurrPage = 1 And CurrPage <> TotalPage Then
				PageStr = "首页  上一页 <a href=""" & FileName & CurrPage + 1 & Extension & """>下一页</a>  <a href= """ & FileName & TotalPage & Extension & """>尾页</a>"
			   ElseIf CurrPage = 1 And CurrPage = TotalPage Then
				PageStr = "首页  上一页 下一页 尾页"
			   ElseIf CurrPage = TotalPage And CurrPage <> 2 Then  '对于最后一页刚好是第二页的要做特殊处理
				 PageStr = "<a href=""" & FileName & "1" & Extension & """>首页</a>  <a href=""" & FileName & CurrPage - 1 & Extension & """>上一页</a> 下一页  尾页"
			   ElseIf CurrPage = TotalPage And CurrPage = 2 Then
				 PageStr = "<a href=""" & FileName & "1" & Extension & """>首页</a>  <a href=""" & FileName & "1" & Extension & """>上一页</a> 下一页  尾页"
			   ElseIf CurrPage = 2 Then
				PageStr = "<a href=""" & FileName & "1" & Extension & """>首页</a>  <a href=""" & FileName & "1" & Extension & """>上一页</a> <a href=""" & FileName & CurrPage + 1 & Extension & """>下一页</a>  <a href= """ & FileName & TotalPage & Extension & """>尾页</a>"
			   Else
				PageStr = "<a href=""" & FileName & "1" & Extension & """>首页</a>  <a href=""" & FileName & CurrPage - 1 & Extension & """>上一页</a> <a href=""" & FileName & CurrPage + 1 & Extension & """>下一页</a>  <a href= """ & FileName & TotalPage & Extension & """>尾页</a>"
			   End If
			 Case 2
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				'ElseIf CurrPage=2 Then
			   '  PageStr="<a href=""" & FileName & "1" & Extension & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & FileName & "-" & Extension &""" title=""上一页""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName &"1"& Extension&""" title=""首页""><font face=webdings>9</font></a> <a href=""" & FileName & CurrPage - 1 & Extension&""" title=""上一页""><font face=webdings>7</font></a> "
				End If
				 dim startpage,n
				 startpage=1
				 if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
				
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"">" & J &"</a>"
				    Else
				     PageStr=PageStr & " <a href=""" & FileName & J& Extension&""">" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & CurrPage + 1& Extension & """ title=""下一页""><font face=webdings>8</font></a> <a href=""" & FileName & TotalPage & Extension& """><font face=webdings>:</font></a> "
				 End If
			 Case 3
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName &"1" & Extension & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & FileName & "1"  & Extension & """ title=""上一页""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName & "1" & Extension & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & FileName & CurrPage - 1 & Extension &""" title=""上一页""><font face=webdings>7</font></a> "
				End If
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & CurrPage + 1 & Extension & """ title=""下一页""><font face=webdings>8</font></a> <a href=""" & FileName & TotalPage & Extension & """><font face=webdings>:</font></a> "
				 End If
			 Case 4
			     n=0:startpage=1
				 pageStr=pageStr & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""right""><tr><td>" & vbcrlf
				 if (CurrPage>1) then pageStr=PageStr & "<a href=""" & FileName & CurrPage - 1 & Extension & """ class=""prev"">上一页</a>"
				 pageStr=pageStr & "<a href=""" & FileName &"1" & Extension & """ class=""prev"">首 页</a>"
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-10
				 If startpage<=0 Then startpage=1 
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"" class=""curr"">" & J &"</a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & FileName & J& Extension&""">" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 If TotalPage>10 Then
					 If CurrPage=TotalPage Then
					 pageStr=pageStr & "<a href=""" & FileName & TotalPage & Extension &""" class=""curr"">..." & TotalPage & "</a>"
					 Else
					 pageStr=pageStr & "<a href=""" & FileName & TotalPage & Extension &""" class=""num"">..." & TotalPage & "</a>"
					 End If
				 End If
				 if (CurrPage<>TotalPage) then pageStr=PageStr & "<a href=""" & FileName & CurrPage + 1 & Extension &""" class=""next"">下一页</a>"
				 pageStr=PageStr & " </td></tr></table>"
			 End Select
			   
			   If CBool(ShowTurnToFlag) = True and pageStyle<>4 Then
				  PageStr = PageStr & " 转到：<select name=""page"" size=""1"" onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"
				  For J = 1 To TotalPage
				   If J = CurrPage Then
					 SelectStr = " selected"
				   Else
					 SelectStr = ""
				   End If
				   If J = 1 Then
					 PageStr = PageStr & "<option value=""" & FileName & "1" & Extension & """" & SelectStr & ">第" & J & "页</option>"
				   Else
					 PageStr = PageStr & "<option value=""" & FileName & J & Extension & """" & SelectStr & ">第" & J & "页</option>"
				   End If
			   Next
				  PageStr = PageStr & "</select>"
			   End If
			   	GetStaticPageList=PageStr	   
      End Function
	 
	 	
	 '论坛顶部广告
	 Function GetClubTopAdList()
	     IF Not IsNUL(Setting(159)) Then
			Dim ADArr,I,J,Str,N,JJ
			ADArr=Split(Setting(159),"@") : N=0
			str="<table cellpadding=""1"" align=""center"" cellspacing=""1"" class=""tableborder1"">" &vbcrlf
			For I=0 To Ubound(AdArr)
			 str=str & "<tr>"& VBCRLF 
			 For J=1 To 4
			   str=str & "<td width=""25%"" class=""tablebody1"" height=""20"" style=""text-align : center; "">" & AdArr(N) & "</td>"  & VBCRLF        
			   n=n+1
			   If N>Ubound(AdArr) Then Exit For
			 Next
			 If N>Ubound(AdArr) Then 
				 do while j<4
				  str=str &"<td width=""25%"" class=""tablebody1"">&nbsp;</td>"
				  j=j+1
				 loop
				 str=str &"</tr>"& VBCRLF 
			     Exit For
			 End If
			 str=str &"</tr>"& VBCRLF 
			Next
			str=str &"</table>"& VBCRLF 
		 End If
		 GetClubTopAdList=str
	 End Function
		
		'取消HTML,UBB代码
		Public Function LoseHtml(ByVal ContentStr)
		    If IsNul(ContentStr) Then LoseHtml="": Exit Function
			Dim objRegExp, Match, Matches
			Set objRegExp = New Regexp
			objRegExp.IgnoreCase = True
			objRegExp.Global = True
			
			objRegExp.Pattern = "<.+?>"
			Set Matches = objRegExp.Execute(ContentStr)
			For Each Match in Matches
			ContentStr=Replace(ContentStr,Match.Value,"")
			Next
			objRegExp.Pattern = "\[.+?\]"
			Set Matches = objRegExp.Execute(ContentStr)
			For Each Match in Matches
			ContentStr=Replace(ContentStr,Match.Value,"")
			Next
			LoseHtml=ContentStr
			Set objRegExp = Nothing
		End Function
		
		
		'---------------------------------------------------------------------------------------------------
		'函数名: G_O_T_S
		'功 能:取得打开类型
		'参 数: OpenType 取true时,新窗口打开
		'--------------------------------------------------------------------------------------------
		Function G_O_T_S(OpenType)
			  If OpenType = "" Or OpenType = False Then
				G_O_T_S = ""
			  ElseIf OpenType = True Then
				G_O_T_S = " target=""_blank"""
			  Else
				G_O_T_S = " target=""" & OpenType & """"
			  End If
		End Function
		'--------------------------------------------------------------------------------------------------
		'函数名: GetCss
		'功 能:取得样式
		'参 数: CssName样式名称
		'--------------------------------------------------------------------------------------------
		Function GetCss(CssName)
			 If CssName = "" Or IsNull(CssName) Then  GetCss = "" Else GetCss = " class=""" & CssName & """"
		End Function
				
		'取得CSS的ID
		Function GetCssID(ID)
		  If ID="" Then GetCssID="" Else GetCssID=" id=""" & ID & """"
		End Function  
		'-------------------------------------------------------------------------------------------------------------
		'函数名: G_R_H
		'功 能:取得单元格行距
		'参 数: RowHeight 默认行距
		'-----------------------------------------------------------------------------------------------------------
		Function G_R_H(RowHeight)
			If IsNumeric(RowHeight) Then G_R_H = RowHeight Else G_R_H = 20
		End Function
	'----------------------------------------------------------------------------------------------------------------------------
		'函数名:GetMenuBg
		'功 能:取得表头背景
		'参 数: MenuBGType 类型 1 取背景图片 0 取背景颜色, MenuBg 背景颜色的值 如#CCCCCC 或 /Upfies/TITLE_BG.GIF ,ColNumber列数
   '---------------------------------------------------------------------------------------------------------------------------
		Function GetMenuBg(MenuBgType, MenuBg, ColNumber)
		  If MenuBgType = 0 Then
			 If MenuBg = "" Then GetMenuBg = "" Else GetMenuBg = MenuBg
		  Else
			 If MenuBg = "" Then
			   GetMenuBg = "url(" & GetDomain & "Images/Default/MenuBg" & ColNumber & ".Gif)"
			 Else
			   If Left(MenuBg, 1) = "/" Or Left(MenuBg, 1) = "\" Then MenuBg = Right(MenuBg, Len(MenuBg) - 1)
			   If LCase(Left(MenuBg, 4)) = "http" Then MenuBg = MenuBg Else MenuBg = GetDomain & MenuBg
			   GetMenuBg = "url(" & MenuBg & ")"
			 End If
		  End If
		End Function

		'--------------------------------------------------------------------------------------------------------------------
		'函数名: GetNavi
		'功 能: 取得导航值
		'参 数: NaviType 导航类型, NaviStr导航值
		'---------------------------------------------------------------------------------------------------------------
		Function GetNavi(NaviType, NaviStr)
		 If NaviType = "0" Then
			 If NaviStr = "" Then GetNavi = "" Else GetNavi = NaviStr
		 ElseIf NaviType = "1" Then
		   If NaviStr <> "" Then  GetNavi = "<img src=""" & NaviStr & """ alt="""" border=""0""/>"
		 Else
			 GetNavi = ""
		 End If
		End Function
		
	  '返回格式化后的时间,showTime显示时间部分
	  Function GetTimeFormat1(DateTime,showTime)
		      If Not IsDate(DateTime) Then GetTimeFormat1=DateTime : Exit Function
		      if DateDiff("n",DateTime,now)<5 then
			   GetTimeFormat1="刚刚"
			  elseif DateDiff("n",DateTime,now)<60 then
			   GetTimeFormat1=DateDiff("n",DateTime,now) & " 分钟前"
			  elseif DateDiff("h",DateTime,now)<5 Then
			   GetTimeFormat1=DateDiff("h",DateTime,now) & " 小时前"
			  elseif DateDiff("d",DateTime,now)=0 Then
			   GetTimeFormat1=" 今天 " & right("0"&Hour(DateTime),2) & ":" & right("0"&Minute(DateTime),2)
			  elseif DateDiff("d",DateTime,now)=1 Then
			   GetTimeFormat1=" 昨天 " & right("0"&Hour(DateTime),2) & ":" & right("0"&Minute(DateTime),2)
			  elseif DateDiff("d",DateTime,now)=2 Then
			   GetTimeFormat1=" 前天 " & right("0"&Hour(DateTime),2) & ":" & right("0"&Minute(DateTime),2)
			  else
			   if showTime=true then
			   GetTimeFormat1=DateTime 
			   else
			   GetTimeFormat1=formatdatetime(DateTime,2)
			   end if
			  end if
	  End Function
	  '返回格式化后的时间
	  Function GetTimeFormat(DateTime)
		      GetTimeFormat=GetTimeFormat1(DateTime,false)
	  End Function

		 '----------------------------------------------------------------------------------------------------------------------------
		'函数名:GetOrigin
		'功 能:取得文章来源并附加上链接
		'参 数: OriginName名称
		'返回值: 形如 <a href="http://www.xinhua.com" target="_blank">新华网</a>
'----------------------------------------------------------------------------------------------------------------------------
		Function GetOrigin(OriginName)
		  Dim RS: Set RS=Server.CreateObject("ADODB.Recordset")
		  RS.Open "select top 1 OriginName,HomePage From KS_Origin Where OriginName='" & Trim(OriginName) & "'", Conn, 1, 1
		  If RS.EOF Then
		   GetOrigin = OriginName
		  Else
		   If RS("HomePage") <> "" And UCase(Trim(RS("HomePage"))) <> "HTTP://" Then
		   GetOrigin = "<a href=""" & Trim(RS("HomePage")) & """ target=""_blank"">" & OriginName & "</a>"
		   Else
			GetOrigin = OriginName
		   End If
		 End If
		 RS.Close:Set RS = Nothing
		End Function
	'----------------------------------------------------------------------------------------------------------------------------
		'函数名:GetMoreLink
		'功 能:取得更多链接
		'参 数: ColNum列数, RowHeight行距, MoreLinkType链接类型, LinkUrl链接地址, OpenTypeStr是否新窗口打开
	'----------------------------------------------------------------------------------------------------------------------------
		Function GetMoreLink(PrintType,ColNum, RowHeight, MoreLinkType, LinkNameStr, LinkUrl, OpenTypeStr)
		   If LinkNameStr = "" Then GetMoreLink = "":Exit Function
	      If PrintType=2 Then
		   If MoreLinkType = "0" Then
			  GetMoreLink = "<li><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></li>"
		   ElseIf MoreLinkType = "1" Then
			  GetMoreLink = "<li><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmiddle""/></a></li>"
		   Else
			 GetMoreLink = ""
		   End If
		  Else
			   LinkNameStr = Trim(LinkNameStr):LinkUrl = Trim(LinkUrl)
			   If MoreLinkType = "0" Then
				  GetMoreLink = "<tr><td colspan= """ & ColNum+1 & """ height=""" & RowHeight & """ align=""right""><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></td></tr>"
			   ElseIf MoreLinkType = "1" Then
						GetMoreLink = "<tr><td colspan= """ & ColNum+1 & """ height=""" & RowHeight & """ align=""right""><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmiddle""/></a></td></tr>"
				 
			   Else
				 GetMoreLink = ""
			   End If
		  End If
		End Function			
 '----------------------------------------------------------------------------------------------------------------------------
		'函数名: GetSplitPic
		'功 能:取得分隔图片
		'参 数: ColSpanNum 列数, SplitPic 图片SRC		'-------------------------------------------------------------------------------------------------------------------------------
		Function GetSplitPic(SplitPic, ColSpanNum)
		     Dim ColStr
			 If SplitPic = "" or IsNull(SplitPic) Then
			   GetSplitPic = ""
			 Else
			   If ColSpanNum>=2 Then ColStr=" colspan=""" & ColSpanNum & """"
			   GetSplitPic = "<tr><td height=""1"""  & ColStr & " background=""" & SplitPic & """ ></td></tr>" & vbcrlf
			 End If
		End Function
	'-------------------------------------------------------------------------------------------------------------------
		'函数名:GetFolderTid
		'功 能:取得子目录的ID集合
		'参 数:  FolderID父目录ID
		'返回值: 形如 1255555,111111,4444的ID集合
   '---------------------------------------------------------------------------------------------------------
		Function GetFolderTid(FolderID)
			GetFolderTid="Select ID From KS_Class Where DelTF=0 AND TS LIKE '%" & FolderID & "%'":Exit Function
		End Function
		'取得专题查询参数,应用于Sql条件
		Function GetSpecialPara(ChannelID,SpecialID)
			   If SpecialID = "-1" Then
					 If FCls.RefreshType = "Special" Then
					   If ChannelID<>0 Then
						GetSpecialPara=" And ID in(select infoid from ks_specialr where ChannelID=" & ChannelID & " and  SpecialID=" & ChkClng(FCls.CurrSpecialID) & ") "
					   Else
						GetSpecialPara=" And InfoID in(select infoid from ks_specialr r where SpecialID=" & ChkClng(FCls.CurrSpecialID) & " and i.channelid=r.channelid) "
					   End If
					 Else
						 GetSpecialPara = ""
					 End If
			  ElseIf (SpecialID = "" Or SpecialID = "0" Or IsNull(SpecialID))  Then
					 GetSpecialPara = ""
			  Else
			      If ChannelID<>0 Then
			      GetSpecialPara=" And ID in(select infoid from ks_specialr where ChannelID=" & ChannelID & " and SpecialID=" & ChkClng(SpecialID) & ") "
				  Else
			      GetSpecialPara=" And InfoID in(select infoid from ks_specialr r where SpecialID=" & ChkClng(SpecialID) & " and i.channelid=r.channelid) "
				  End If
			  End If
		End Function
		
	'载入文件类自定义字段
	Sub LoadFieldToXml()
	  If Not IsObject(Application(SiteSN & "_FeildXml")) then
			Dim Rs:Set Rs = Conn.Execute("Select ChannelID,FieldName,fieldtype From KS_Field Where FieldType=9 or FieldType=10 Order By FieldID")
			Set Application(SiteSN & "_FeildXml")=RsToxml(Rs,"row","FeildXml")
			Set Rs = Nothing
	  End If
	End Sub
	
	
	 '生成模型字段xml
	   Sub CreateFieldXML(ChannelID,Param)
		  CloseConn: OpenConn
	      Dim RS:Set RS=Conn.Execute("select * from KS_Field Where ChannelID=" & ChannelID & " " & Param & " Order By OrderID,FieldID")
	      Dim XMLStr:XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
		  XMLStr=XMLStr&" <field>" &vbcrlf
		  If Not RS.Eof Then
					Do While Not RS.Eof
					    XMLStr=XMLStr & "  <fielditem id=""" & RS("FieldID") &""" fieldname=""" & replace(rs("fieldname"),"&","") & """ groupid=""" & RS("GroupID") &""">"&vbcrlf
						XMLStr=XMLStr & "    <title>" & rs("title") & "</title>" &vbcrlf
						XMLStr=XMLStr & "    <tips><![CDATA[" & rs("tips") &"]]></tips>" &vbcrlf
						XMLStr=XMLStr & "    <fieldtype>" & rs("fieldtype") & "</fieldtype>" &vbcrlf
						XMLStr=XMLStr & "    <defaultvalue><![CDATA[" & rs("defaultvalue") &"]]></defaultvalue>" &vbcrlf
						If Not IsNul(rs("options")) Then
						XMLStr=XMLStr & "    <options><![CDATA[" & replace(rs("options"),vbcrlf,"\n") &"]]></options>" &vbcrlf
						Else
						XMLStr=XMLStr & "    <options><![CDATA[" & rs("options") &"]]></options>" &vbcrlf
						End If
						XMLStr=XMLStr & "    <mustfilltf>" & rs("mustfilltf") & "</mustfilltf>" &vbcrlf
						XMLStr=XMLStr & "    <showonform>" & rs("showonform") & "</showonform>" &vbcrlf
						XMLStr=XMLStr & "    <showonuserform>" & rs("showonuserform") & "</showonuserform>" &vbcrlf
						XMLStr=XMLStr & "    <showonclubform>" & chkclng(rs("showonclubform")) & "</showonclubform>" &vbcrlf
						XMLStr=XMLStr & "    <allowfileext>" & rs("AllowFileExt") & "</allowfileext>" &vbcrlf
						XMLStr=XMLStr & "    <width>" & rs("width") & "</width>" &vbcrlf
						XMLStr=XMLStr & "    <height>" & rs("height") & "</height>" &vbcrlf
						XMLStr=XMLStr & "    <maxfilesize>" & rs("maxfilesize") & "</maxfilesize>" &vbcrlf
						XMLStr=XMLStr & "    <editortype>" & rs("editortype") & "</editortype>" &vbcrlf
						XMLStr=XMLStr & "    <showunit>" & rs("showunit") & "</showunit>" &vbcrlf
						if not IsNul(rs("unitoptions")) Then
						XMLStr=XMLStr & "    <unitoptions><![CDATA[" & replace(rs("unitoptions"),vbcrlf,"\n") &"]]></unitoptions>" &vbcrlf
						Else
						XMLStr=XMLStr & "    <unitoptions><![CDATA[" & rs("unitoptions") &"]]></unitoptions>" &vbcrlf
						End If
						XMLStr=XMLStr & "    <parentfieldname>" & rs("ParentFieldName") & "</parentfieldname>" &vbcrlf
						XMLStr=XMLStr & "    <maxlength>" & rs("maxlength") & "</maxlength>" &vbcrlf
					    XMLStr=XMLStr & "  </fielditem>"&vbcrlf
					 RS.MoveNext
					Loop
		  End If
		   XMLStr=XMLStr &" </field>" &vbcrlf
		   Call WriteTOFile(Setting(3) & "config/fielditem/field_" & ChannelID & ".xml",xmlstr)
		  RS.Close :Set RS=Nothing
	   End Sub
	
	
	'自定义字段用户组
	Sub LoadFieldGroupXML()
	  If Not IsObject(Application(SiteSN & "_FieldGroupXml")) then
			Dim Rs:Set Rs = Conn.Execute("Select ID,ChannelID,GroupName,IsSys,OrderID From KS_FieldGroup Order By OrderID,ID")
			Set Application(SiteSN & "_FieldGroupXml")=RsToxml(Rs,"row","FieldGroupXml")
			Set Rs = Nothing
	  End If
	End Sub
		
	'添加自关联数据库	
	Sub FileAssociation(ByVal ChannelID,ByVal InfoID,ByVal Content,ByVal Flag)
	  on error resume next
	  If Flag<>0 Then
	  Conn.Execute("Delete From KS_UploadFiles Where IsAnnex<>1 and ChannelID=" & ChannelID & " and InfoID=" & InfoID)
	  End If
	  If ChannelID<>0 And ChannelID<1000 and channelid<>7 Then
	     Dim Node,tempStr
	     LoadFieldToXml()
		 For Each Node In Application(SiteSN & "_FeildXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &" and @fieldtype=9 or @fieldtype=10]")
		    tempStr=tempStr & Request(Node.SelectSingleNode("@fieldname").text)
		 Next
		 Content=Content & tempStr
	  End If
	  Dim FileLists,I,FileArr
	  FileLists=GetFilesList(ChannelID,Content)
	  If Not IsNul(FileLists) Then
	    FileArr=Split(FileLists,"|")
		For I=0 To Ubound(FileArr)
		    Dim strFile:strFile=trim(FileArr(i))
            If instr("00000000"&lcase(strfile),".asp")<>0 or instr("00000000"&lcase(strfile),".aspx")<>0 or instr("00000000"&lcase(strfile),".asa")<>0 or instr("00000000"&lcase(strfile),".php")<>0 or instr("00000000"&lcase(strfile),".jsp")<>0 then
			ElseIf instr("00000000"&lcase(strfile),".mp3")<>0 or  instr("00000000"&lcase(strfile),".mp4")<>0 or instr("00000000"&lcase(strfile),".flv")<>0 or instr("00000000"&lcase(strfile),".f4v")<>0 or instr("00000000"&lcase(strfile),".swf")<>0  or instr("00000000"&lcase(strfile),".wmv")<>0 or instr("00000000"&lcase(strfile),".rmvb")<>0 or instr("00000000"&lcase(strfile),".rm")<>0 Then
		     Conn.Execute("Update [KS_UploadFiles] Set InfoID=" & InfoID &" Where filename='"& strFile&"' and title<>''")
			ElseIf Conn.Execute("select top 1 ID From [KS_UploadFiles] Where ChannelID=" & ChannelID& " and infoid=" & InfoId &" and filename='"& strFile&"'").eof Then
		     Conn.Execute("Insert Into [KS_UploadFiles](ChannelID,InfoID,ClassID,FileName,IsAnnex,UserName,AddDate) values(" &ChannelID &"," & InfoID &",0,'" & replace(strFile,"'","''") & "',0,'" & C("UserName") & "'," & SQLNowString&")")
		   End If
		Next
	  End If
	End Sub
	
	'根据内容获取上传文件名
	Public Function GetFilesList(ChannelID,Content)
	    If IsNul(Content) Then Exit Function
		Dim re, UpFile, BFU, FileName,SaveFileList,FileExt
		If ChannelID<1000 Then FileExt=ReturnChannelAllowUpFilesType(ChannelID,0) Else FileExt=Setting(7)
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(\/" & replace(Setting(91),"/","\/") &")[^(\/" & replace(Setting(91),"/","\/") &")]?(.*?)[.]{1}(" & FileExt & "|wma|mp3)"
		're.Pattern = "(\/uploadfiles\/)[^(\/uploadfiles\/)](.*?)[.]{1}(" & FileExt & "|wma|mp3)"
		Set UpFile = re.Execute(Content)
		Set re = Nothing
		For Each BFU In UpFile
		  If Instr(SaveFileList,trim(BFU))=0 and len(trim(BFU))>len(setting(91))+1 Then
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
	
	'**************************************************
	'函数名：ReturnChannelAllowUpFilesTF
	'作  用：返回频道的是否允许上传文件
	'参  数：ChannelID--频道ID
	'**************************************************
	Public Function ReturnChannelAllowUpFilesTF(ChannelID)
	  If ChannelID = "" Or Not IsNumeric(ChannelID) Then  ChannelID = 0
	   Dim CRS:Set CRS=Server.CreateObject("ADODB.RECORDSET")
	   CRS.Open "Select top 1 UpFilesTF From KS_Channel Where ChannelID=" & ChannelID, Conn, 1, 1
	  If CInt(ChannelID) = 0 Or (CRS.EOF And CRS.BOF) Then  '默认允许上传文件
		ReturnChannelAllowUpFilesTF = True
	  Else
		If CRS(0) = 1 Then ReturnChannelAllowUpFilesTF = True	Else ReturnChannelAllowUpFilesTF = False
	  End If
	CRS.Close:Set CRS = Nothing
	End Function
	
	'取上传目录6.0改为按日期存放
	Function GetUpFilesDir()
	   Dim DateFolder
	   Dim UserFolder:UserFolder=ChkClng(C("UserID"))
	   Select Case ChkClng(Setting(96))
	    case 1 DateFolder=Setting(3) & Setting(91) & Year(Now) & "-" & Right("0"&Month(Now),2) & "/" & UserFolder   
	    case 2 DateFolder=Setting(3) & Setting(91) & Year(Now) & "-" & Right("0"&Month(Now),2) & "-" & Right("0"&Day(Now),2) & "/" & UserFolder   
	    case 3 DateFolder=Setting(3) & Setting(91) & Year(Now) & "/" & UserFolder   
	    case 5 DateFolder=Setting(3) & Setting(91) & Year(Now) 
	    case 6 DateFolder=Setting(3) & Setting(91) & Year(Now) & "-" & Right("0"&Month(Now),2)
	    case 7 DateFolder=Setting(3) & Setting(91) & Year(Now) & "-" & Right("0"&Month(Now),2) & "-" & Right("0"&Day(Now),2) 
	    case Else DateFolder=Setting(3) & Setting(91) & UserFolder   
	   End Select
	   CreateListFolder(DateFolder)
	   GetUpFilesDir=DateFolder
	End Function
	
	'取得后台公共管理部分的上传目录,一般用于广告,公告设置等
	Function GetCommonUpFilesDir()
	  Dim Str
	  If C("SuperTF")="1" Then 
	    Str=Setting(3) & Setting(91)
	  Else
	    Str=GetUpFilesDir()
	  End If
	  If Right(Str,1)="/" Then Str=Left(Str,Len(Str)-1)
	  GetCommonUpFilesDir=Str
	End Function

	'**************************************************
	'函数名：ReturnChannelAllowUserUpFilesTF
	'作  用：返回频道是否允许会员上传文件
	'参  数：ChannelID--频道ID
	'**************************************************
	Public Function ReturnChannelAllowUserUpFilesTF(ChannelID)
	  If ChannelID = "" Or Not IsNumeric(ChannelID) Then '默认允许上传文件
	  ReturnChannelAllowUserUpFilesTF = True:Exit Function
	  End If
		If C_S(ChannelID,26) = 1 Then
		 ReturnChannelAllowUserUpFilesTF = True
		Else
		 ReturnChannelAllowUserUpFilesTF = False
		End If
	End Function

	'**************************************************
	'函数名：ReturnChannelUserUpFilesDir
	'作  用：返回频道前台会员的上传目录
	'参  数：ChannelID--频道ID,UserFolder-按用户名生成的目录
	'返回值：目录字符串
	'**************************************************
	Public Function ReturnChannelUserUpFilesDir(ChannelID,UserFolder)
	   If HasChinese(UserFolder) Then
	     Dim Ce:Set Ce=new CtoeCls
	     UserFolder="[" & Ce.CTOE(R(UserFolder)) & "]"
	     Set Ce=Nothing
	   End If
	   
	   ChannelID = ChkCLng(ChannelID)
	   If UserFolder="" Then UserFolder="Temp"
	   Select Case ChannelID
	    Case 9999,9990 '用户头像,固定放到uploadfiles/user/avatar目录下
		   ReturnChannelUserUpFilesDir=Setting(3)&"UploadFiles/User/avatar/"
		Case 9998,9997 '相册
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/xc/"
		Case 9996 '圈子图片
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/team/"
		Case 9995 '音乐
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/music/"
		Case 9994 '论坛
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)& UserFolder &"/" & Year(Now) & "-" & Right("0"&Month(Now),2) &"/"
		Case 9993 '日志
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/blog/"
		Case 9001
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/ad/"
		Case 9002
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/zs/"
		Case 8666
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/message/"
		Case Else
		  ReturnChannelUserUpFilesDir = Setting(3) & Setting(91)&"User/" & UserFolder &"/"
	   End Select
	End Function
	
	'判断有没有中文
	Function HasChinese(Content)
				 Dim regEx, Matches, Match
				 Set regEx = New RegExp
				 regEx.Pattern="[\u4e00-\u9fa5]+"
				 regEx.IgnoreCase = True
				 regEx.Global = True
				 Set Matches = regEx.Execute(Content)
				 If Matches.count > 0 Then
				   HasChinese=true
				 Else
				   HasChinese=false
				 end if
	End Function
	
	
	'**************************************************
	'函数名：ReturnChannelAllowUpFilesSize
	'作  用：返回频道的最大允许上传文件大小
	'参  数：ChannelID--频道ID
	'**************************************************
	Public Function ReturnChannelAllowUpFilesSize(ChannelID)
	   ChannelID = ChkClng(ChannelID)
	   Dim CRS:Set CRS=conn.execute("Select top 1 UpFilesSize From KS_Channel Where ChannelID=" & ChannelID)
	  If CInt(ChannelID) = 0 Or (CRS.EOF And CRS.BOF) Then
		ReturnChannelAllowUpFilesSize = Setting(6)
	  Else
		ReturnChannelAllowUpFilesSize = CRS(0)
	  End If
	CRS.Close:Set CRS = Nothing
	End Function
	'**************************************************
	'函数名：ReturnChannelAllowUpFilesType
	'作  用：返回频道的允许上传的文件类型
	'参  数：ChannelID--频道ID,TypeFlag 0-取全部 1-图片类型 2-flash 类型 3-Windows 媒体类型 4-Real 类型 5-其它类型
	'**************************************************
	Public Function ReturnChannelAllowUpFilesType(ChannelID, TypeFlag)
	  If ChkClng(ChannelID) = 0 Then  ReturnChannelAllowUpFilesType = Setting(7):Exit Function
	  If Not IsNumeric(TypeFlag) Then TypeFlag = 0
		If TypeFlag = 0 Then   '所有允许的类型
		 ReturnChannelAllowUpFilesType = Replace(Replace(C_S(ChannelID,28) & "|" & C_S(ChannelID,29) & "|" & C_S(ChannelID,30) & "|" & C_S(ChannelID,31) & "|" & C_S(ChannelID,32),"||","|"),"||","|")
		Else
		 ReturnChannelAllowUpFilesType = Replace(C_S(ChannelID,27+TypeFlag),"||","|")
		End If
		If Right(ReturnChannelAllowUpFilesType,1)="|" Then ReturnChannelAllowUpFilesType=left(ReturnChannelAllowUpFilesType,len(ReturnChannelAllowUpFilesType)-1)
	End Function
	'返回付款方式名称,参数TypeID,0名称 1折扣率
	Function ReturnPayment(ID,TypeID)
	  If Application(SiteSn &"Payment_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 TypeName,Discount From KS_PaymentType Where TypeID=" & ChkClng(ID),conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnPayment=rs(0)
			  If RS(1)<100 Then ReturnPayment=ReturnPayment & "&nbsp;&nbsp;<span style=""color:red"">折扣率:" & RS(1) & "%</span>"
			 Else
			  ReturnPayment=rs(1)
			 End if
		End iF 
		Application(SiteSn &"Payment_" & ID&TypeID)=ReturnPayment
	  Else
	    ReturnPayment=Application(SiteSn &"Payment_" & ID&TypeID)
	  End If
	End Function
		
			
	'===========================计算订单运费===============================
	Function GetFreight(ExpressID,ToCity,ByVal totalweight,ByRef ExpressCompany)
	     if TotalWeight=-1 then GetFreight=0 : Exit Function
	     IF ChkClng(Setting(180))<>1 Then GetFreight=0 : Exit Function
	     Dim RST,fweight,carriage,C_fee,W_fee,foundexpress
		 Set RST=Server.CreateObject("ADODB.RECORDSET")
		 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& expressid &" and a.tocity like '%"&ToCity&"%'",conn,1,1
		 If RST.Eof Then
			foundexpress=false:fweight=0 :carriage=0 : C_fee=0 : W_fee=0
		 Else
			foundexpress=true:fweight=rst("fweight"):carriage=rst("carriage"):C_fee=rst("C_fee"):W_fee=rst("W_fee"):ExpressCompany=rst("typename")
		 End If
		 RST.Close
		 If foundexpress=false Then
		  If DataBaseType=1 Then
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& expressid &" and (convert(varchar(200),tocity)='全国统一运费' or convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
		  Else
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& expressid &" and (a.tocity='全国统一运费' or a.tocity='' or a.tocity is null)",conn,1,1
		  End If
		  if rst.eof then
			rst.close : set rst=nothing
			GetFreight="0"
			Exit Function
		  else
			fweight=rst("fweight"):carriage=rst("carriage"):C_fee=rst("C_fee"):W_fee=rst("W_fee"):ExpressCompany=rst("typename")
		  end if
		  rst.close
		 End If
		 set rst=nothing
		 if (totalweight<=fweight) or (C_fee=0 and W_fee=0) then
		   GetFreight=carriage
		 else
		   totalweight=int(totalweight+0.99) 
		   if totalweight<=fweight then
		    GetFreight=carriage
		   else
		   GetFreight=carriage+int((totalweight-fweight) / W_fee+0.99)*C_fee
		   end if
		 end if
		End Function
	'**********************************************************************
	'函数名：ReturnSpecial
	'作  用：返回专题名称
	'参  数：Selected-预选中项
	'返回值：专题名称
	'**********************************************************************
	Public Function ReturnSpecial(SelectID)
	 Dim RS,ParaStr,SpecialChannelStr,SQL,K
	 Set RS=Conn.Execute("Select ClassID,ClassName From KS_SpecialClass Order By OrderID")
	 If Not RS.Eof Then SQL=RS.GetRows(-1)
     RS.Close
	 If IsArray(SQL) Then
	  For K=0 To Ubound(SQL,2)
	  ReturnSpecial = ReturnSpecial & "<optgroup label='---" & SQL(1,K) & "---'>"
	  Set RS=Conn.Execute("Select SpecialName,SpecialID From KS_Special Where ClassID=" & SQL(0,K) & " Order By SpecialID Desc")
		 If Not RS.EOF Then
		  Do While Not RS.EOF
			 If FoundInarr(Trim(SelectID),Trim(RS(1)),",")<>0 Then
				  ReturnSpecial = ReturnSpecial & "<Option value=" & RS(1) & " Selected>" & Trim(RS("SpecialName")) & SpecialChannelStr & "</Option>"
			 Else
				  ReturnSpecial = ReturnSpecial & "<Option value=" & RS(1) & ">" & Trim(RS("SpecialName")) & SpecialChannelStr & "</Option>"
			 End If
			 RS.MoveNext
		  Loop
		End If
	 Next
	  RS.Close:Set RS = Nothing
	 Else
	  Set RS = Nothing
	 End If
	End Function
	
	'**************************************************
	'函数：FoundInArr
	'作  用：检查一个数组中所有元素是否包含指定字符串
	'参  数：strArr     ----字符串
	'        strToFind    ----要查找的字符串
	'       strSplit    ----数组的分隔符
	'返回值：True,False
	'**************************************************
	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function
	
	'检查是否是数字 ，并转换为长整型
	Public Function ChkClng(ByVal str)
	    On error resume next
		If IsNumeric(str) Then
			ChkClng = CLng(str)
		Else
			ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function
	'**************************************************
	'函数名：ShowPage
	'作  用：显示“上一页 下一页”等信息
	'参  数：filename文件名 TotalNumber总数量 MaxPerPage每页数量 ShowTurn显示转到 PrintOut立即输出
	'**************************************************
	Function ShowPage(totalnumber, MaxPerPage, FileName, CurrPage,ShowTurn,PrintOut)
	             Dim n,j,startpage,pageStr,TotalPage,ParamStr
				 If totalnumber Mod MaxPerPage = 0 Then
						TotalPage = totalnumber \ MaxPerPage
				 Else
						TotalPage = totalnumber \ MaxPerPage + 1
				 End If
				 
				 ParamStr=QueryParam("page") : If ParamStr<>"" Then ParamStr="&" & ParamStr	
				 n=0:startpage=1:CurrPage=ChkClng(CurrPage)
				 pageStr=pageStr & "<div id='fenye' class='fenye' align=""right""><table  border=""0""><form action=""" & FileName & "?1=1" & ParamStr & """ name=""pageform"" method=""post""><tr><td nowrap>" & vbcrlf
				 if (CurrPage>1) then pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage-1 & ParamStr & """ class=""prev"">上一页</a>"
				 pageStr=pageStr & "<a href=""" & FileName & "?page=1" & ParamStr & """ class=""prev"">首 页</a>"
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-9
				 If startpage<0 Then startpage=1
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"" class=""curr"">" & J &"</a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & FileName & "?page=" &J& ParamStr & """>" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 if CurrPage<>TotalPage Then pageStr=pageStr & "<a href=""" & FileName & "?page=" & TotalPage & ParamStr & """ class=""next"">末 页</a>"
				 if CurrPage<>TotalPage and totalnumber>MaxPerPage then pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage+1 & ParamStr & """ class=""next"">下一页</a>"
				 pageStr=PageStr & " </td><td>"
				 If ShowTurn=true Then
				 If CurrPage=TotalPage Then CurrPage=0
				 pageStr=PageStr & " 转到:<input class='textbox' type='text' value='" & (CurrPage + 1) &"' name='page' style='width:30px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>"
				 End If
				 PageStr=PageStr & "</td></tr></form></table></div>"
				If PrintOut=true Then echo PageStr Else ShowPage=PageStr
	End Function

	
	'模型选项
	Sub LoadChannelOption(ChannelID)
		If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig
		Dim ModelXML,Node
		Set ModelXML=Application(SiteSN&"_ChannelConfig")
		For Each Node In ModelXML.documentElement.SelectNodes("channel")
		 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and ChkClng(Node.SelectSingleNode("@ks6").text)<9 Then
		  If Trim(ChannelID)=Trim(Node.SelectSingleNode("@ks0").text) Then
		  echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"
		  Else
		  echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
		  End If
		 End If
		next
	End Sub
	
	'从字符串里返回图片地址，当n=1时，只返回一张图片地址	
	Function GetPictureFromStr(byval str,n)  
	 Dim objRegExp,strs,RetStr,Matches,Match
	 Set objRegExp = New Regexp'设置配置对象   
	objRegExp.IgnoreCase = True'忽略大小写   
	objRegExp.Global = True'设置为全文搜索   
	objRegExp.Pattern ="src\=.+?\.(png|gif|jpg)"
	strs=""
	Set Matches =objRegExp.Execute(trim(str))'开始执行配置   
	For Each Match in Matches   
		Retstr=replace(Match.Value,"src=","")
		Retstr=Replace(Replace(Retstr,"""",""),"'","")
		if strs="" then
		   strs=Retstr 
		   if n=1 then GetPictureFromStr=Retstr : Exit Function
		else
		   strs=strs & "|" & Retstr
		end if
	Next   
	GetPictureFromStr = strs 
	End Function  

	
	'**************************************************
	'函数名：ReplaceBeyondUrl
	'作  用：保存远程文件到本地
	'参  数：LocalFile 本地文件,BFU远程文件
	'返回值：无
	'**************************************************
	Public Function ReplaceBeyondUrl(ReplaceContent, SaveFilePath)
		Dim re, BeyondFile, BFU, SaveFileName,SaveFileList
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp|jpeg)))"
		
		Set BeyondFile = re.Execute(ReplaceContent)
		Set re = Nothing
		For Each BFU In BeyondFile
		  If Instr(SaveFileList,BFU)=0 Then
			SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & MakeRandom(10) & Mid(BFU, InStrRev(BFU, "."))
			If Instr(BFU,Setting(2))<=0  AND Instr(lcase(BFU),"kesion.com")=0  Then
			Call SaveBeyondFile(SaveFilePath&SaveFileName,BFU)
			ReplaceContent = Replace(ReplaceContent, BFU, Setting(2) & SaveFilePath & SaveFileName)
			End If
		  End If
		   SaveFileList=SaveFileList & "," & BFU
		Next
		ReplaceBeyondUrl = ReplaceContent
	End Function

	'==================================================
	'过程名：SaveBeyondFile
	'作  用：保存远程的文件到本地
	'参  数：LocalFileName ------ 本地文件名
	'参  数：RemoteFileUrl ------ 远程文件URL
	'==================================================
	Function SaveBeyondFile(ByVal LocalFileName,RemoteFileUrl)
	   If Instr(lcase(RemoteFileUrl),"kesion.com")<>0 then SaveBeyondFile=RemoteFileUrl : Exit Function
	    on error resume next
		Dim SaveRemoteFile:SaveRemoteFile=True
		dim Ads,Retrieval,GetRemoteData
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", RemoteFileUrl, False, "", ""
			.Send
			If .Readystate<>4 then
				SaveRemoteFile=False
				Exit Function
			End If
			GetRemoteData = .ResponseBody
		End With
		Set Retrieval = Nothing
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile server.MapPath(LocalFileName),2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
		SaveBeyondFile=SaveRemoteFile
		IF Setting(174)="1" Then
		'加水印
		Dim T:Set T=New Thumb
		call T.AddWaterMark(LocalFileName)
		Set T=Nothing
		End If
	end Function
	'****************************************************
	'参数说明
	  'Subject     : 邮件标题
	  'MailAddress : 发件服务器的地址,如smtp.163.com
	  'LoginName     ----登录用户名(不需要请填写"")
	  'LoginPass     ----用户密码(不需要请填写"")
	  'Email       : 收件人邮件地址
	  'Sender      : 发件人姓名
	  'Content     : 邮件内容
	  'Fromer      : 发件人的邮件地址
	'****************************************************
	  Public Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	   On Error Resume Next
		Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
			jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j
			jmail.Charset = "gb2312" '邮件的文字编码为国标
			'jmail.ContentType = "text/html" '邮件的格式为HTML格式
			jmail.AddRecipient Email '邮件收件人的地址
			jmail.From = Fromer '发件人的E-MAIL地址
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '您的邮件服务器登录名
				JMail.MailServerPassword = Decrypt(LoginPass) '登录密码
			  End If

			jmail.Subject = Subject '邮件的标题 
			JMail.Body = Content
			JMail.HTMLBody = Content
			JMail.Priority = 1'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
			jmail.Send(MailAddress) '执行邮件发送（通过邮件服务器地址）
			jmail.Close() '关闭对象
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function
	
	'**************************************************
	'函数名： ReplaceUserFile
	'作  用：将会员上传的文件移到指定的上传目录下
	'**************************************************
	Public Function ReplaceUserFile(ReplaceContent,ChannelID)
		Dim re, BeyondFile, BFU, SaveFileName
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		're.Pattern = "(" &Setting(3)&Setting(91) & "user(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp|rar|doc|xsl|zip|exe)))"
		re.Pattern = "(" &Setting(3)&Setting(91) &"user[^(""|'|\s)]*[.]{1}(dwg|gif|jpg|png|bmp|flv|swf|mp3|mp4|rar|doc|xsl|zip|exe|docx))"
		Set BeyondFile = re.Execute(ReplaceContent)
		Set re = Nothing
		Dim Path,DateDir
		Path = GetUpFilesDir()
		DateDir = Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		Path = Path & "/" & DateDir
		For Each BFU In BeyondFile
		    Dim NewPath:NewPath=Path & Split(BFU,"/")(Ubound(Split(bfu,"/")))
			Call CopyFile(BFU,NewPath)
			ReplaceContent = Replace(ReplaceContent, BFU, NewPath)
		Next
		ReplaceUserFile = ReplaceContent
	End Function
	
	'模拟剪切文件操作
	Public Function CopyFile(OldPath,NewPath)
		CopyFile=false
		Call CreateListFolder(Replace(NewPath,Split(NewPath,"/")(Ubound(Split(NewPath,"/"))),""))
		on error resume next
		dim fso:set fso = Server.CreateObject(Setting(99))
	    fso.CopyFile Server.MapPath(OldPath), server.mappath(NewPath), True
		DeleteFile(OldPath)
		if err then
			CopyFile=false
		else
			CopyFile=true
		end if
		IF err Then	 CopyFile=false
	End Function
	
	'**************************************************
	'函数名：CreateListFolder
	'作  用：不限分级创建目录 形如 1\2\3\ 则在网站根目录下创建分级目录
	'参  数：Folder要创建的目录
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function CreateListFolder(ByVal Folder)
		Dim FSO, WaitCreateFolder, SplitFolder, CF, k
		 On Error Resume Next
		If Folder = "" Then
		 CreateListFolder = False:Exit Function
		End If
	   Folder = Replace(Folder, "\", "/")
	   If Right(Folder, 1) <> "/" Then Folder = Folder & "/"
	   If Left(Folder, 1) <> "/" Then Folder = "/" & Folder

		 Set FSO = CreateObject(Setting(99))
		 If Not FSO.FolderExists(Server.MapPath(Folder)) Then
		   SplitFolder = Split(Folder, "/")
		 For k = 0 To UBound(SplitFolder) - 1
		  If k = 0 Then
		   CF = SplitFolder(k) & "/"
		  Else
		   CF = CF & SplitFolder(k) & "/"
		  End If
		  If (Not FSO.FolderExists(Server.MapPath(CF))) Then
			 FSO.CreateFolder (Server.MapPath(CF))
			 CreateListFolder = True
		  End If
		 Next
	   End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear
	   CreateListFolder = False
	   Else
	   CreateListFolder = True
	   End If
	 End Function
	
	 '**************************************************
	'函数名：DeleteFolder
	'作  用：删除指定目录
	'参  数：FolderStr要删除的目录
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function DeleteFolder(FolderStr)
	   Dim FSO
	   On Error Resume Next
	   FolderStr = Replace(FolderStr, "\", "/")
	   Set FSO = CreateObject(Setting(99))
		If FSO.FolderExists(Server.MapPath(FolderStr)) Then
			  FSO.DeleteFolder (Server.MapPath(FolderStr))
		Else
		DeleteFolder = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFolder = False
	   Else
	   DeleteFolder = True
	   End If
	End Function
	 '**************************************************
	'函数名：DeleteFile
	'作  用：删除指定文件
	'参  数：FileStr要删除的文件
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function DeleteFile(FileStr)
	   Dim FSO
	   On Error Resume Next
	   Set FSO = CreateObject(Setting(99))
		If FSO.FileExists(Server.MapPath(FileStr)) Then
			FSO.DeleteFile Server.MapPath(FileStr), True
		Else
		DeleteFile = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFile = False
	   Else
	   DeleteFile = True
	   End If
	End Function
	'**********************************************************************
	'函数名：CheckFileShowOrNot
	'参数：AllowShowExtNameStr允许的文件扩展名，ExtName实际文件扩展名
	'**********************************************************************
	Public Function CheckFileShowOrNot(AllowShowExtNameStr, ExtName)
		If ExtName = "" Then
			CheckFileShowOrNot = False
		Else
			If InStr(1, AllowShowExtNameStr, ExtName) = 0 Then
				CheckFileShowOrNot = False
			Else
				CheckFileShowOrNot = True
			End If
		End If
	End Function
	'**********************************************************************
	'函数名：GetFieSize
	'作用：取得指定文件的大小
	'参数：FilePath--文件位置
	'**********************************************************************
	Public Function GetFieSize(FilePath)
			GetFieSize = 0
			Dim FSO, F
			On Error Resume Next
			Set FSO = Server.CreateObject(Setting(99))
			Set F = FSO.GetFile(FilePath)
			GetFieSize = F.size
			Set F = Nothing:Set FSO = Nothing
	End Function
    '取得目录大小
	Public Function GetFolderSize(FolderPath)
		dim fso:Set FSO = Server.CreateObject(Setting(99))
		if fso.FolderExists(Server.MapPath(FolderPath)) then
		dim userfilespace:set UserFileSpace=FSO.GetFolder(Server.MapPath(FolderPath))
        GetFolderSize=UserFileSpace.size
		else
		 GetFolderSize=0:exit function
		end if
		set userfilespace=nothing:set fso=nothing
	End Function
	'*************************************************************************************
	'文件备份过程
	'过程名：backupdata
	'参数：CurrPath原文件完整物理地址，BackPath目标备份文件完整物理地址
	'*************************************************************************************
	
	Public Function BackUpData(CurrPath, BackPath)
	  On Error Resume Next
	  Dim FSO:Set FSO = Server.CreateObject(Setting(99))
	 FSO.copyfile CurrPath, BackPath
	 If Err Then
	   BackUpData = False
	 Else
	   BackUpData = True
	 End If
	  FSO.Close:Set FSO = Nothing
	End Function
	'------------------检查某一目录是否存在-------------------
	Public Function CheckDir(FolderPath)
	        Dim fso:Set fso = InitialObject(Setting(99))
			CheckDir=fso.FolderExists(Server.MapPath(FolderPath))
			Set fso = Nothing
	End Function
	'------------------检查某一文件是否存在-------------------
	Public Function CheckFile(FileName)
		 On Error Resume Next
		 Dim FsoObj
		 Set FsoObj = Server.CreateObject(Setting(99))
		  If Not FsoObj.FileExists(Server.MapPath(FileName)) Then
			  CheckFile = False
			  Exit Function
		  End If
		 CheckFile = True:Set FsoObj = Nothing
	End Function
	
	
	'**************************************************
	'函数名：WriteTOFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'        Content   ------要写入目标文件的内容
	'返回值：成功返回true ,失败返回false
	'**************************************************
	Public Function WriteTOFile(FileName, Content)
	    On Error Resume Next
		dim stm:set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以文本模式读取
		stm.mode=3
		stm.charset="utf-8"
		stm.open
		stm.WriteText content
		stm.SaveToFile server.MapPath(FileName),2 
		stm.flush
		stm.Close
		set stm=nothing
	   If Err.Number <> 0 Then
		 WriteTOFile = False
	   Else
		 WriteTOFile = True
	   End If
	End Function
	'**************************************************
	'函数名：ReadFromFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'返回值：成功返回文件内容 ,失败返回""
	'**************************************************
	Public Function ReadFromFile(FileName)
	    On Error Resume Next
		dim str,stm
		set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以本模式读取
		stm.mode=3 
		stm.charset="utf-8"
		stm.open
		stm.loadfromfile server.MapPath(FileName)
		str=stm.readtext
		stm.Close
		set stm=nothing
		if err.number<>0 then Die ("Error:<br/>[" & Server.MapPath(FileName) & "] File does not exist!"):Exit Function
		ReadFromFile=str
	End Function
	'**************************************************
	'函数名：MakeRandom
	'作  用：生成指定位数的随机数
	'参  数： maxLen  ----生成位数
	'返回值：成功:返回随机数
	'**************************************************
	Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function
	'生成随机密码
	Function GetRndPassword(PasswordLen)
		Dim Ran, i, strPassword
		strPassword = ""
		For i = 1 To PasswordLen
			Randomize
			Ran = CInt(Rnd * 2)
			Randomize
			If Ran = 0 Then
				Ran = CInt(Rnd * 25) + 97
				strPassword = strPassword & UCase(Chr(Ran))
			ElseIf Ran = 1 Then
				Ran = CInt(Rnd * 9)
				strPassword = strPassword & Ran
			ElseIf Ran = 2 Then
				Ran = CInt(Rnd * 25) + 97
				strPassword = strPassword & Chr(Ran)
			End If
		Next
		GetRndPassword = strPassword
	End Function
	'**************************************************
	'函数名：MakeRandomChar
	'作  用：生成指定位数的随机数字符串 如 "sJKD_!@KK"
	'参  数： Length  ----生成位数
	'返回值：成功返回随机字符串
	'**************************************************
	Public Function MakeRandomChar(Length)
	  Dim I, tempS, v
	  Dim c(65)
	   tempS = ""
	   c(1) = "a": c(2) = "b": c(3) = "c": c(4) = "d": c(5) = "e": c(6) = "f": c(7) = "g"
	   c(8) = "h": c(9) = "i": c(10) = "j": c(11) = "k": c(12) = "l": c(13) = "m": c(14) = "n"
	  c(15) = "o": c(16) = "p": c(17) = "q": c(18) = "r": c(19) = "s": c(20) = "t": c(21) = "u"
	  c(22) = "v": c(23) = "w": c(24) = "x": c(25) = "y": c(26) = "z": c(27) = "1": c(28) = "2"
	   c(29) = "3": c(30) = "4": c(31) = "5": c(32) = "6": c(33) = "7": c(34) = "8": c(35) = "9"
	  c(36) = "-": c(37) = "_": c(38) = "@": c(39) = "!": c(40) = "A": c(41) = "B": c(42) = "C"
	  c(43) = "D": c(44) = "E": c(45) = "F": c(46) = "G": c(47) = "H": c(48) = "I": c(49) = "J": c(50) = "K"
	  c(51) = "L": c(52) = "M": c(53) = "N": c(54) = "O": c(55) = "P": c(56) = "Q": c(57) = "R": c(58) = "S"
	  c(59) = "J": c(60) = "U": c(61) = "V": c(62) = "W": c(63) = "X": c(64) = "Y": c(65) = "Z"
	
	  If IsNumeric(Length) = False Then
		 MakeRandomChar = "":Exit Function
	  End If
	  For I = 1 To Length
		 Randomize
		 v = Int((65 * Rnd) + 1):tempS = tempS & c(v)
		 Next
		MakeRandomChar = tempS
	End Function
	'**************************************************
	'函数名：GetFileName
	'作  用：构造文件名。
	'参  数：FsoType  ----生成类型,addDate   -----添加时间,Ext--扩展名
	'**************************************************
	Public Function GetFileName(FsoType, AddDate, Ext)
		Dim N,Fname
		Do While True
			Randomize
			N = chkclng(Rnd * 12 + 10)
			Dim Y,M,D
			Y=Year(AddDate):M=Right("0"&Month(AddDate),2):D=Right("0"&Day(AddDate),2)
			 Select Case FsoType
			  Case 1:Fname = Y & "/" & M & "-" & D & "/" & MakeRandom(N) & Ext  '年/月-日/随机数+扩展名
			  Case 2:Fname = Y & "/" & M & "/" & D & "/" & MakeRandom(N) & Ext '年/月/日/随机数+扩展名
			  Case 3:Fname = Y & "-" & M & "-" & D & "/" & MakeRandom(N) & Ext '年-月-日/随机数+扩展名
			  Case 4:Fname = Y & "/" & M & "/" & MakeRandom(N) & Ext '年/月/随机数+扩展名
			  Case 5:Fname = Y & "-" & M & "/" & MakeRandom(N) & Ext '年-月/随机数+扩展名
			  Case 12:Fname = Y & M & "/" & MakeRandom(N) & Ext '年-月/随机数+扩展名
			  Case 6:Fname = Y & M & D & "/" & MakeRandom(N) & Ext '年月日/随机数+扩展名
			  Case 7:Fname = Y & "/" & MakeRandom(N) & Ext '年/随机数+扩展名
			  Case 8:Fname = Y & M & D & MakeRandom(N) & Ext '年+月+日+随机数+扩展名
			  Case 9:Fname = MakeRandom(N) & Ext
			  Case 10:Fname = MakeRandomChar(N) & Ext '随机字符
			  Case 11:Fname ="ID"
			  Case Else
			   Fname = Y & M & D & Ext '年月日+扩展名
			End Select
			If Fname="ID" Then Exit Do
			If  Conn.Execute("Select top 1 ID from KS_ItemInfo Where fname='" & Fname & "'").Eof Then Exit Do
		Loop
		
		GetFileName=Fname
	End Function
	
	Function EchoHtmlBody()
	   Dim Str:Str ="<!DOCTYPE html><html><head><script src=""" & Setting(3) &"KS_Inc/jquery.js""></script>"
	   Str ="<script src=""" & Setting(3) &"KS_Inc/jquery.js""></script><script src=""" & Setting(3) &"ks_inc/common.js""></script></head><body>"
	   EchoHtmlBody=Str
	End Function
	
	'**************************************************
	'函数名：Alert
	'作  用：弹出成功提示。
	'参  数：SuccessStr  ----成功提示信息
	'        Url   ------成功提示按下"确定"转向链接
	'返回值：无
	'**************************************************
	Public Function Alert(SuccessStr, Url)
	  Echo (EchoHtmlBody)
	 If Url <> "" Then
	  die ("<script> $.dialog.alert('" & SuccessStr & "', function (){ location.href='" & Url & "'; });</script>")
	 Else
	  echo ("<script> $.dialog.alert('" & SuccessStr & "');</script>")
	 End If
	End Function
	Public Function AlertDoFun(SuccessStr, fun)
	  Echo (EchoHtmlBody)
	  die ("<script> $.dialog.alert('" & SuccessStr & "', function (){ " & fun & "; });</script>")
	End Function
	'**************************************************
	'函数名：AlertHistory
	'作  用：弹出警告消息后,停止所在页面的执行,返回n级。
	'参  数：SuccessStr  ----成功提示信息
	'        n   ------返回级数
	'返回值：无
	'**************************************************
	Public Function AlertHistory(SuccessStr, N)
	    echo (EchoHtmlBody)
		echo ("<script> $.dialog.alert('" & SuccessStr & "',function(){ history.back(" & N & ");});</script>")
		die ""
	End Function
	'提示成功。并返回
	Sub AlertHintScript(SuccessStr)
	  echo (EchoHtmlBody)
	  echo "<script language=JavaScript>" & vbCrLf
	  echo "$.dialog.alert('" & SuccessStr & "',function(){"
	  echo "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "');" & vbCrLf
	  echo "});</script>" & vbCrLf
	  die ""
	End Sub
	'**************************************************
	'函数名：Confirm
	'作  用：弹出成功提示。
	'参  数：SuccessStr  ----成功提示信息
	'        Url   ------成功提示按下"确定"转向链接
	'        Url1   ------confirm按下"取消"转向链接
	'返回值：无
	'**************************************************
	Public Function Confirm(SuccessStr, Url, Url1)
	 echo (EchoHtmlBody)
	  echo ("<script> $.dialog.confirm('" & SuccessStr & "',function(){location.href='" & Url & "';},function(){location.href='" & Url1 & "';});</script>")
	End Function
	Public Function ConfirmDoFun(SuccessStr, fun1, fun2)
	 echo (EchoHtmlBody)
	  echo ("<script> $.dialog.confirm('" & SuccessStr & "',function(){ " & fun1 & "},function(){" & fun2 & "});</script>")
	End Function
    
	Public Sub ShowTips(Action,Message)
		 Response.Redirect(Setting(3) & "plus/error.asp?action="&action &"&message="&Server.URLEncode(message))
	End Sub
	'**************************************************
	'函数名：ShowError
	'作  用：显示错误信息。
	'参  数：Errmsg  ----出错信息
	'返回值：无
	'**************************************************
	Public Sub ShowError(Errmsg)
		echo ("<br><br><div align=""center"">")
		echo ("  <center>")
		echo ("  <table border=""0"" cellpadding='2' cellspacing='1' class='border' width=""75%"" style=""MARGIN-TOP: 3px"" class='border'>")
		echo ("	 <tr class=tdbg>")
		echo ("			  <td width=""100%"" height=""30"" align=""center"">")
		echo ("				<b> " & Errmsg & "&nbsp; </b>")
		echo ("				</b>")
		echo ("			  </td>")
		echo ("	 </tr>")
		echo ("	 <tr  class=tdbg>")
		echo ("			  <td width=""100%"" height=""30"" align=""center"">")
		echo ("				<p><b><a href=""javascript:history.go(-1)"">...::: 点 此 返 回 ")
		echo ("				:::...</a></b>")
		echo ("			  </td>")
		echo ("			</tr>")		
		echo ("  </table>")
		echo ("  </center>")
		echo ("</div>")
		die  ("")
    end sub
	'*****************************************************************************************
	'函数名：ReturnPowerResult
	'作  用：检查操作权限。
	'参  数：ChannelID---所在系统(频道) 1文章系统2图片系统等 PowerOpName ---当前操作的权限名称
	'返回值：允许返回true,否则返回false
	'******************************************************************************************
	Public Function ReturnPowerResult(ChannelID, PowerOpName)
		If C("AdminName") = "" Then
			 ReturnPowerResult = False
			 Exit Function
		ElseIf C("SuperTF") = "1" Then    '超级管理组拥有所有权限
			ReturnPowerResult = True
			Exit Function
		Else
		   If Instr(C("ModelPower"),C_S(ChannelID,10)&"0")>0  and ChannelID<>0 then          '没有任何管理权
			ReturnPowerResult = False
		   ElseIf Instr(C("ModelPower"),C_S(ChannelID,10)&"1")>0 and ChannelID<>0 then      '拥有所有权限
			ReturnPowerResult = True
		   ElseIf Instr(C("ModelPower"),C_S(ChannelID,10)&"2")>0 and ChannelID<>0 then      '限制栏目,拥有部分权限
			ReturnPowerResult = CheckPower(PowerOpName)
		   Else
			ReturnPowerResult = CheckPower(PowerOpName)
		   End If
	   End If
	End Function
	'结合上面ReturnPowerResult过程序使用
	Public Function CheckPower(PowerOpName)
	        Dim PowerList, ModelPower
		    PowerList = Trim(C("PowerList"))
			If (PowerList <> "") And (PowerOpName <> "") Then
				Select Case Left(PowerOpName, 4)     '检查是否有模块的总权限
				  Case "KMST" '系统
				   If Instr(C("ModelPower"),"sysset0") >0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case "KMUA" '用户
				   If Instr(C("ModelPower"),"user0") >0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case "KMTL"
				  If Instr(C("ModelPower"),"lab0")>0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case "KSMM"
				  If Instr(C("ModelPower"),"model0")>0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				 ' Case "KSMS"
				 ' If Instr(C("ModelPower"),"subsys0")>0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case Else
				   ModelPower = true
				End Select
				   If InStr(PowerList, PowerOpName) <> 0 And ModelPower Then
					 CheckPower = True:Exit Function
				   Else
					 CheckPower = False:Exit Function
				   End If
			Else
			   CheckPower = False:Exit Function
			End If
			
	End Function
	'结合上面ReturnPowerResult过程使用,     ReturnFlag  ----类型 0关闭,1返回前一页2,转向URL, Url    -错误后转向的Url
	Sub ReturnErr(ReturnFlag, Url)
	   If ReturnFlag = 0 Then
		 die ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');window.close();</script>")
	   ElseIf ReturnFlag = 1 Then
		 die ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');history.back();</script>")
	  ElseIf ReturnFlag = 2 Then
	     die ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');location.href='" & Url & "';</script>")
	  End If
	End Sub
	'插入网站后台日志 , UserName --- 管理员账号 , ResultTF ---0登录失败 1---登录成功 ,ScriptName---登录路径 ,Descript---描述信息
	Sub InsertLog(UserName, ResultTF, ScriptName, Descript)
		Dim SystemStr:SystemStr = Request.ServerVariables("HTTP_USER_AGENT")
		If InStr(SystemStr, "Windows NT 5.2") Then
		  SystemStr = "Win2003"
		ElseIf InStr(SystemStr, "Windows NT 5.0") Then
		  SystemStr = "Win2000"
		ElseIf InStr(SystemStr, "Windows NT 5.1") Then
		  SystemStr = "WinXP"
		ElseIf InStr(SystemStr, "Windows NT") Then
		  SystemStr = "WinNT"
		ElseIf InStr(SystemStr, "Windows 9") Then
		  SystemStr = "Win9x"
		ElseIf InStr(SystemStr, "unix") Or InStr(SystemStr, "linux") Or InStr(SystemStr, "SunOS") Or InStr(SystemStr, "BSD") Then
		  SystemStr = "类似Unix"
		ElseIf InStr(SystemStr, "Mac") Then
		  SystemStr = "Mac"
		Else
		  SystemStr = "Other"
		End If
		Conn.Execute("Insert into KS_Log(UserName,ResultTF,LoginTime,LoginOS,LoginIP,ScriptName,Description) values('" & UserName & "'," & ResultTF & "," & SqlNowString & ",'" & replace(SystemStr,"'","''") & "','" & getip & "','" & replace(scriptname,"'","''") & "','" & replace(descript,"'","""") & "')")
	End Sub
	
	'获取分页链接
		Function GetPageUrl(ChannelID,ID,CurrPage)
		      Dim Url
			  If ID=0 Then  '手机版本
			    Url="?modelid=" & channelid &"&page=" & CurrPage
				if not isnul(request("tj")) then url=url&"&tj=" &s("tj")
				if not isnul(request("rm")) then url=url&"&rm=" &s("rm")
			  ElseIf C_S(ChannelID,48)=0 or channelid=-1 Then
			   Url="?id=" & ID &"&page=" & CurrPage
			  ElseIf C_S(ChannelID,48)=2 Then
			   Url=GCls.StaticPreList & "-" & ID & "-"& CurrPage & GCls.StaticExtension
			  Else
			   Url="?" & GCls.StaticPreList & "-" & ID & "-" & CurrPage& GCls.StaticExtension
			  End If
			GetPageUrl=Url
		End Function
		
		Function ReplacePage(PageStyle,ChannelID,id,CurrPage,TotalPut,PerPageNumber)
		 Dim n:n=PageStyle
         Dim Tp,TotalPage,ItemUnit
		 Dim XML:Set XML=LFCls.GetXMLFromFile("pagestyle")
		 Dim Node:Set Node= XML.documentElement.selectSingleNode("/pagestyle/item[@name='" & n & "']/content")
		  If Not Node Is Nothing Then
		   Tp=Node.text
		  End If
		  
		  if (TotalPut mod PerPageNumber)=0 then
				TotalPage= TotalPut \ PerPageNumber
		  else
				TotalPage = TotalPut \ PerPageNumber + 1
		 end if
		 ItemUnit=C_S(ChannelID,4): if IsNul(ItemUnit) Then ItemUnit="条"
		  
		  
		  Dim homeUrl,endUrl,prevUrl,nextUrl 

            if (CurrPage = 1 and CurrPage <>TotalPage) then
                homeUrl = "javascript:;"
                prevUrl = "javascript:;"
                nextUrl = GetPageUrl(ChannelID,ID,CurrPage + 1)
                endUrl = GetPageUrl(ChannelID,ID,TotalPage)
            elseif (CurrPage = 1 and CurrPage = TotalPage) then
                homeUrl = "javascript:;"
                prevUrl = "javascript:;"
                nextUrl = "javascript:;"
                endUrl = "javascript:;"
            elseif (CurrPage = TotalPage and  CurrPage <> 2)  then '对于最后一页刚好是第二页的要做特殊处理
                homeUrl = GetPageUrl(ChannelID,ID,1)
                prevUrl = GetPageUrl(ChannelID,ID,CurrPage - 1)
                nextUrl = "javascript:;"
                endUrl = "javascript:;"
            elseif (CurrPage = TotalPage and CurrPage = 2) then
                homeUrl = GetPageUrl(ChannelID,ID,1)
                prevUrl = GetPageUrl(ChannelID,ID,1)
                nextUrl = "javascript:;"
                endUrl = "javascript:;"
            elseif (CurrPage = 2) then
                homeUrl = GetPageUrl(ChannelID,ID,1)
                prevUrl = GetPageUrl(ChannelID,ID,1)
                nextUrl = GetPageUrl(ChannelID,ID,CurrPage + 1)
                endUrl = GetPageUrl(ChannelID,ID,TotalPage)
            else
                homeUrl = GetPageUrl(ChannelID,ID,1)
                prevUrl = GetPageUrl(ChannelID,ID,CurrPage - 1)
                nextUrl = GetPageUrl(ChannelID,ID,CurrPage + 1)
                endUrl = GetPageUrl(ChannelID,ID,TotalPage)
            end if
		 
		  Tp=Replace(Tp,"{$homeurl}",homeurl)
		  Tp=Replace(Tp,"{$prevurl}",prevurl)
		  Tp=Replace(Tp,"{$nexturl}",nexturl)
		  Tp=Replace(Tp,"{$endurl}",endurl)

          if (instr(Tp,"{$pagenumlist}")<>0) then
                Dim j,p,pageStr:pageStr=""
                Dim StartPage:startpage = 1
                if (CurrPage >= 7)  then startpage = CurrPage - 5
                if (TotalPage - CurrPage < 5) then startpage = TotalPage - 9
                if (startpage <= 0) then startpage = 1
                Dim nn:nn = 1
                for p = startpage to TotalPage
                    if (p = CurrPage) then
                        pageStr=pageStr & " <a href=""#"" class=""curr"">" & p & "</a>"
                    else
                         pageStr=pageStr & " <a class=""num"" href=""" & GetPageUrl(ChannelID,ID,p)& """>" & p & "</a>"
                    end if
					if (nn >= 10) then exit for
					nn=nn+1
                Next
                Tp = replace(Tp, "{$pagenumlist}", pagestr)
         End If
		 
		 if (instr(Tp,"{$turnpage}")<>0) then
                pageStr="<select name=""page"" id=""turnpage"" onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"
                for j = 1 to totalPage
                  pageStr=pageStr &"<option value=""" & GetPageUrl(ChannelID,ID,j) & """"
				  if j=currPage then pageStr=pageStr &" selected"
				  pageStr=pageStr &">第" & j & "页</option>"
                next
                pageStr=pageStr &"</select>"
                Tp = replace(Tp, "{$turnpage}", pageStr)
        end if

		  
		 Tp=Replace(Tp,"{$currentpage}",CurrPage)
		 Tp=Replace(Tp,"{$maxperpage}",PerPageNumber)
		 Tp=Replace(Tp,"{$totalpage}",TotalPage)
		 Tp=Replace(Tp,"{$totalput}",TotalPut)
		 Tp=Replace(Tp,"{$itemunit}",ItemUnit)
		 ReplacePage=Tp
		End Function
	
	
		
	'*************************************************************************************
	'函数名:GetClassID
	'作  用:生成新目录或频道的ID号,生成目录ID 年+10位随机
	'参  数:无
	'*************************************************************************************
	Public Function GetClassID()
		Do While True
		 GetClassID = Year(Now()) & MakeRandom(10)
		 If Conn.Execute("Select ID from KS_Class Where ID='" & GetClassID & "'").Eof Then Exit Do
		Loop
	End Function
	
	'取专题分类参数
	Function GetSpecialClass(ClassID,FieldName)
	  If Not IsObject(Application(SiteSN & "_SpecialClass")) then
			Dim Rs:Set Rs = Conn.Execute("Select ClassID,ClassName,ClassEname,Descript,FsoIndex From KS_SpecialClass Order By ClassID")
			Set Application(SiteSN & "_SpecialClass")=RsToxml(Rs,"row","root")
			Set Rs = Nothing
	  End If
	  Dim Node:Set Node=Application(SiteSN&"_SpecialClass").documentElement.selectSingleNode("row[@classid=" & ClassID & "]/@" & Lcase(FieldName) & "")
	  If Not Node Is Nothing  Then GetSpecialClass=Node.text
	  Set Node = Nothing
	End Function
	
	'载入供求类型
	Sub LoadGQTypeToXml()
	  If Not IsObject(Application(SiteSN & "_SupplyType")) then
			Dim Rs:Set Rs = Conn.Execute("Select TypeID,TypeName,TypeColor From KS_GQType Order By TypeID")
			Set Application(SiteSN & "_SupplyType")=RsToxml(Rs,"row","SupplyType")
			Set Rs = Nothing
	  End If
	End Sub
	
	'*************************************************************************************
	'函数名:GetGQTypeName
	'作  用:获得供求的交易类别名称
	'参  数:TypeID
	'*************************************************************************************
	Public Function GetGQTypeName(TypeID)
	   If Not IsNumeric(TypeID) Then GetGQTypeName="":Exit Function
	   LoadGQTypeToXml()
	   Dim NodeName,NodeColor
	   Set NodeName=Application(SiteSN & "_SupplyType").documentElement.SelectSingleNode("row[@typeid=" & ChkClng(TypeID) & "]/@typename")
	   If Not NodeName  Is Nothing Then
		 Set NodeColor=Application(SiteSN & "_SupplyType").documentElement.SelectSingleNode("row[@typeid=" & ChkClng(TypeID) & "]/@typecolor")
		 GetGQTypeName="<span style=""color:" & NodeColor.Text & """>" & NodeName.Text & "</span>"
	   End If 
	End Function
	'返回供求交易类型列表
	'参数：Flag:1-标签调用 0-添加信息时调用
	Public Function ReturnGQType(SelID,Flag)
	   If S("TypeID")<>"" Then
	     ReturnGQType=LoseHtml(GetGQTypeName(S("TypeID")))
		 ReturnGQType=ReturnGQType & "<input type=""hidden"" name=""TypeID"" id=""TypeID"" value=""" & S("TypeID") &"""/>"
	   Else
		   Dim Node
		   LoadGQTypeToXml()
			If Flag=1 Then 
			   ReturnGQType="<select class=""textbox"" name=""TypeID"" id=""TypeID"" style=""width:70%"">"
				If SelID = "0" Then ReturnGQType=ReturnGQType & "<option  value=""0"" selected>- 交易类型不限 -</option>"	Else ReturnGQType=ReturnGQType & "<option  value=""0"">- 交易类型不限 -</option>"
		   Else
			   ReturnGQType="<select class=""textbox"" name=""TypeID"" id=""TypeID"">"
		   End If
		   For Each Node In Application(SiteSN & "_SupplyType").DocumentElement.SelectNodes("row")
			 If trim(SelID)=trim(node.SelectSingleNode("@typeid").text) Then
				 ReturnGQType=ReturnGQType & "<option value=""" & node.SelectSingleNode("@typeid").text & """ style=""color:" & node.SelectSingleNode("@typecolor").text & """ selected>" & node.SelectSingleNode("@typename").text & "</option>"
			 else
				  ReturnGQType=ReturnGQType & "<option value=""" & node.SelectSingleNode("@typeid").text & """ style=""color:" & node.SelectSingleNode("@typecolor").text & """>" & node.SelectSingleNode("@typename").text & "</option>"
			 end if
		   Next
		   ReturnGQType=ReturnGQType & "</select>"
	  End If
	End Function
	
	'*************************************************************************************
	'函数名:GetInfoID
	'作  用:生成文章,图片或下载等的唯一ID
	'参  数:ChannelID--频道ID
	'*************************************************************************************
	Public Function GetInfoID(ChannelID)
	   Dim TableNameStr,ID
	   TableNameStr = "Select top 1 ProID From " & C_S(ChannelID,2) & " Where ProID='"
	   Do While True
		   ID = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "上午", ""), "下午", ""),"/","") & MakeRandom(3)
		   If Conn.Execute(TableNameStr & ID & "'").Eof Then Exit Do
	   Loop
		GetInfoID=ID
	End Function
	'*************************************************************************************
	'函数名:ReplaceInnerLink
	'作  用:替换站内链接
	'参  数:Content-待替换内容
	'*************************************************************************************
	Public Function ReplaceInnerLink(ByVal Content)
	  on error resume next
	  'Content=HTMLCode(Content)
	  If Not IsObject(Application(SiteSN & "_InnerLink")) then
			Dim Rs:Set Rs = Conn.Execute("Select Title,Url,OpenType,CaseTF,Times,Start From KS_InnerLink Where OpenTF=1 Order By ID")
			Set Application(SiteSN & "_InnerLink")=RecordsetToxml(Rs,"InnerLink","InnerLinkList")
			Set Rs = Nothing
	  end if
		Dim Node,CaseTF,Times,Inti,DLocation,XLocation,StrReplace,CurrentTimes,SourceStr
		For Each Node In Application(SiteSN & "_InnerLink").DocumentElement.SelectNodes("InnerLink")
			 CurrentTimes=0
			 Dim OpenTypeStr:OpenTypeStr = G_O_T_S(Node.selectSingleNode("@ks2").text)
			 CaseTF=Cint(Node.selectSingleNode("@ks3").text)
			 Times=Cint(Node.selectSingleNode("@ks4").text)
			 Inti=ChkClng(Node.selectSingleNode("@ks5").text)
			 StrReplace=Node.selectSingleNode("@ks0").text
			 If Inti=0 Then Inti=1
			If InStr(Inti,Content,StrReplace,CaseTF)>0 Then
			  Do While instr(Inti,Content,StrReplace,CaseTF)>0
			    Inti=instr(Inti,Content,StrReplace,CaseTF)
				If Inti<>0 then
				  DLocation=instr(Inti,Content,">") '仅替换在><之间的关键字
				  XLocation=instr(Inti,Content,"<")
				  If DLocation >= XLocation Then
				    dim sstr:sstr="<a href="""&Node.selectSingleNode("@ks1").text&"""" & OpenTypeStr & " class=""innerlink"">"&Node.selectSingleNode("@ks0").text&"</a>"
					Content=left(Content,Inti-1) & sstr &mid(Content,Inti+len(StrReplace))
					Inti=Inti+len(sstr)
					CurrentTimes=CurrentTimes+1
					If Times<>-1 And CurrentTimes>= Times Then Exit Do
				 Else
				    Inti=Inti+len(StrReplace)
				 End If
			   End If
			  Loop	
			End if
		Next
	 ReplaceInnerLink = Content
	End Function
	
	'=============================================================
	'函数作用：判断来源URL是否来自外部
	'=============================================================
	Public Function CheckOuterUrl()
		On Error Resume Next
		Dim server_v1, server_v2
		server_v1 = LCase(Trim(Request.ServerVariables("HTTP_REFERER")))
		server_v2 = LCase(Trim(Request.ServerVariables("SERVER_NAME")))
		CheckOuterUrl = True
		If Mid(server_v1,8,len(server_v2))=server_v2 Then CheckOuterUrl=False 
	End Function 
	
	'加密
	Function Encrypt(ecode)
	dim texts,i
	for i=1 to len(ecode)
	texts=texts & chr(asc(mid(ecode,i,1))+3)
	next
	Encrypt = texts
	End Function
	'解密
	Function Decrypt(dcode)
	 If IsNul(dcode) then exit function
	dim texts,i
	for i=1 to len(dcode)
	texts=texts & chr(asc(mid(dcode,i,1))-3)
	next
	Decrypt=texts
	End Function
	'匹配 img src,结果以|隔开 
	Function GetImgSrcArr(strng) 
	If strng="" Or IsNull(strng) Then GetImgSrcArr="":Exit Function
	Dim regEx,Match,Matches,values
	Set regEx = New RegExp
	regEx.Pattern = "src\=.+?\.(gif|jpg)"
	regEx.IgnoreCase = true 
	regEx.Global = True 
	Set Matches = regEx.Execute(strng)
	For Each Match in Matches
		if instr(lcase(Match.Value),"fileicon")=0 then
		 values=values&Match.Value&"|" 
		end if
	Next 
	GetImgSrcArr = Replace(Replace(Replace(Replace(values,"'",""),"""",""),"src=",""),Setting(2),"")
	If GetImgSrcArr<>"" Then GetImgSrcArr = left(GetImgSrcArr,len(GetImgSrcArr)-1)
	End Function
	

	'取得Request.Querystring 或 Request.Form 的值
	Public Function G(Str)
	 G = Replace(Replace(Replace(Replace(Request(Str)&"", "'", ""), """", ""),"%","％"),"*","＊")
	End Function
	Function DelSql(Str)
		Dim SplitSqlStr,SplitSqlArr,I
		SplitSqlStr="'|%|dbcc |alter|drop|*|and |exec|or |insert|select|delete|update|count |master|truncate|declare|char|mid|chr|set |where|xp_cmdshell"
		SplitSqlArr = Split(SplitSqlStr,"|")
		For I=LBound(SplitSqlArr) To Ubound(SplitSqlArr)
			If Instr(LCase(Str),SplitSqlArr(I))>0 Then
				Die "<script>alert('系统警告！\n\n1、您提交的数据有恶意字符" & SplitSqlArr(I) &";\n2、您的数据已经被记录;\n3、您的IP："&GetIP&";\n4、操作日期："&Now&";\n		Powered By Kesion.Com!');window.close();</script>"
			End if
		Next
		DelSql = Str
    End Function
	'取得Request.Querystring 或 Request.Form 的值
	Public Function S(Str)
	 S = DelSql(Replace(Replace(Replace(Request(Str)&"", "'", ""), """", ""),"%","％"))
	End Function
	'读Cookies值
	Public Function C(Str)
	 C=DelSql(Request.Cookies(SiteSN)(Str))
	End Function
	
	'取得QueryString,或Form参数集合,参数NoCollect表示不收集的字段,多个用英文逗号隔开
	Function QueryParam(NoCollect)
		 Dim Param,R
		 For Each r In Request.QueryString
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.QueryString(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				End If
			  End If
		 End If
		 Next
	
		' If Param<>"" Then QueryParam=Param:Exit Function
		 For Each r In Request.Form
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.Form(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				End If
			  End If
		 End If
		 Next
		 QueryParam=CheckXSS(Param) '防跨站
	End Function


	
   	'进行脚本过滤
	Function CheckScript(byVal Content)
	    If IsNul(Content) Then  Exit Function
		Dim oRegExp,oMatch,spamCount
		Set oRegExp = New Regexp
		oRegExp.IgnoreCase = True
		oRegExp.Global = True
		oRegExp.pattern ="<script(.|\n)+?/script>"
		Content=oRegExp.replace(Content,"")
		Set oRegExp=Nothing
		CheckScript=Content
	End Function


	'关闭采集数据库对象
	Public Sub CloseConnItem()
	   On Error Resume Next
	   If IsObject(ConnItem) Then
		 ConnItem.Close:Set ConnItem = Nothing
	   End If
	End Sub
	'文章自动分页
	'参数：Content-文章内容 SplitPageStr-文章分隔线 PerPageLen-每页大约字符数
	Function AutoSplitPage(Content,SplitPageStr,maxPagesize)
	    Dim sContent,ss,i,IsCount,c,iCount,strTemp,Temp_String,Temp_Array
		sContent=Content
		
		If maxPagesize<100 Or Len(sContent)<maxPagesize+100 Then
			AutoSplitPage=sContent
		End If
		
		sContent=Replace(sContent, SplitPageStr, "")
		sContent=Replace(sContent, "&nbsp;", "<&nbsp;>")
		sContent=Replace(sContent, "&gt;", "<&gt;>")
		sContent=Replace(sContent, "&lt;", "<&lt;>")
		sContent=Replace(sContent, "&quot;", "<&quot;>")
		sContent=Replace(sContent, "&#39;", "<&#39;>")
		
		If sContent<>"" and maxPagesize<>0 and InStr(1,sContent,SplitPageStr)=0 then
			IsCount=True:Temp_String=""
			For i= 1 To Len(sContent)
				c=Mid(sContent,i,1)
				If c="<" Then
					IsCount=False
				ElseIf c=">" Then
					IsCount=True
				Else
					If IsCount=True Then
						'If Abs(Asc(c))>255 Then
						'	iCount=iCount+2
						'Else
							iCount=iCount+1
						'End If
						If iCount>=maxPagesize And i<Len(sContent) Then
							strTemp=Left(sContent,i)
							'If CheckPagination(strTemp,"table|a|b>|i>|strong|div|span") then
								Temp_String=Temp_String & Trim(CStr(i)) & "," 
								iCount=0
							'End If
						End If
					End If
				End If	
			Next
			If Len(Temp_String)>1 Then Temp_String=Left(Temp_String,Len(Temp_String)-1)
			Temp_Array=Split(Temp_String,",")
			For i = UBound(Temp_Array) To LBound(Temp_Array) Step -1
				ss = Mid(sContent,Temp_Array(i)+1)
				If Len(ss) > 100 Then
					sContent=Left(sContent,Temp_Array(i)) & SplitPageStr & ss
				Else
					sContent=Left(sContent,Temp_Array(i)) & ss
				End If
			Next
		End If
		sContent=Replace(sContent, "<&nbsp;>", "&nbsp;")
		sContent=Replace(sContent, "<&gt;>", "&gt;")
		sContent=Replace(sContent, "<&lt;>", "&lt;")
		sContent=Replace(sContent, "<&quot;>", "&quot;")
		sContent=Replace(sContent, "<&#39;>", "&#39;")
		AutoSplitPage=sContent
	End Function
    '结合以上函数使用
	Private Function CheckPagination(strTemp,strFind)
		Dim i,n,m_ingBeginNum,m_intEndNum
		Dim m_strBegin,m_strEnd,FindArray
		strTemp=LCase(strTemp)
		strFind=LCase(strFind)
		If strTemp<>"" and strFind<>"" then
			FindArray=split(strFind,"|")
			For i = 0 to Ubound(FindArray)
				m_strBegin="<"&FindArray(i)
				m_strEnd  ="</"&FindArray(i)
				n=0
				do while instr(n+1,strTemp,m_strBegin)<>0
					n=instr(n+1,strTemp,m_strBegin)
					m_ingBeginNum=m_ingBeginNum+1
				Loop
				n=0
				do while instr(n+1,strTemp,m_strEnd)<>0
					n=instr(n+1,strTemp,m_strEnd)
					m_intEndNum=m_intEndNum+1
				Loop
				If m_intEndNum=m_ingBeginNum then
					CheckPagination=True
				Else
					CheckPagination=False
					Exit Function
				End If
			Next
		Else
			CheckPagination=False
		End If
	End Function
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) then
		    fString = ClearBadChr(fString)
			fString = Replace(fString, "&", "&amp;")
			fString = Replace(fString, "'", "&#39;")
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
			fString = Replace(fString, Chr(32), " ")
			fString = Replace(fString, Chr(9), " ")
			fString = Replace(fString, Chr(34), "&quot;")
			fString = Replace(fString, Chr(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			'fString = Replace(fString, " ", "&nbsp;")
			'fString = Replace(fString, Chr(10), "<br />")
		HTMLEncode = fString
		End If
	End Function
	
	Function ClearBadChr(str)
	  If Str<>"" Then
	     Dim re:Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True

		re.Pattern="(on(error|load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)='[^']+')"
		str = re.Replace(str, "")
		re.Pattern="(on(error|load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)=""[^""]+"")"
		str = re.Replace(str, "")
		re.Pattern="(on(error|load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)=[\s]*)"	
		str = re.Replace(str, "")

		
		're.Pattern="((name|id|class)=""[^""]+"")"
		'str = re.Replace(str, "")
		re.Pattern = "(<s+cript[^>]*?>([\w\W]*?)<\/s+cript>)"
		str = re.Replace(str, "")
		re.Pattern = "(<iframe[^>]*?>([\w\W]*?)<\/iframe>)"
		str = re.Replace(str, "")
		re.Pattern = "(<p>&nbsp;<\/p>)"
		str = re.Replace(str, "")
		Set re=Nothing
		ClearBadChr = str
	 End If	
	End Function

	
	Public Function HTMLCode(HtmlStr)
		If Not IsNul(HtmlStr) then
		'HtmlStr = Replace(HtmlStr, "&nbsp;", " ")
		HtmlStr = Replace(HtmlStr, "&quot;", Chr(34))
		HtmlStr = Replace(HtmlStr, "&#39;", Chr(39))
		HtmlStr = Replace(HtmlStr, "&#123;", Chr(123))
		HtmlStr = Replace(HtmlStr, "&#125;", Chr(125))
		HtmlStr = Replace(HtmlStr, "&#36;", Chr(36))
		HtmlStr = Replace(HtmlStr, "&amp;", "&")
		'HtmlStr = Replace(HtmlStr, vbCrLf, "")

		HtmlStr = Replace(HtmlStr, "&gt;", ">")
		HtmlStr = Replace(HtmlStr, "&lt;", "<")
		'HtmlStr = Replace(HtmlStr, chr(13)&chr(10), "<br/>")
		
		HTMLCode = HtmlStr
		End If
	End Function

	

	Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
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
	

	'====================================商城相关_begin=============================================
	    '订单支付成功，发送Email,手机消息通知
	    Sub OrderPaySuccessTips(RS)
		   '发送Email/手机通知
			Dim EmailMsg:EmailMsg=Setting(74)
			Dim SmsMsg:SmsMsg=Split(Setting(155)&"∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮∮","∮")(5)
			If Not IsNul(EmailMsg) Then
				 If RS("UserName")<>"游客" Then
				    Call SendInfo(rs("username"),"system","订单[" & RS("OrderID") &"]支付成功通知",ReplaceOrderLabel(EmailMsg,rs))
						 If Not IsNul(RS("Email")) Then
						  CALL SendMail(Setting(12), Setting(13), Setting(14), "订单[" & RS("OrderID") &"]支付成功通知", RS("Email"),RS("ContactMan"), ReplaceOrderLabel(EmailMsg,rs),Setting(11))
						 End If
				 End If
			End If
			 Dim Mobile:Mobile=RS("Mobile")		
			 If Not IsNUL(SmsMsg) And Setting(157)="1" Then
						      If IsNul(Mobile) and Trim(RS("UserName"))<>"游客" Then
								  Dim RSU:Set RSU=Conn.Execute("select top 1 Mobile From KS_User Where UserName='" & RS("UserName") &"'")
								  If Not RSU.Eof Then
									Mobile=RSU(0)
								  End If
								  RSU.Close
								  Set RSU=Nothing
							   End If 
						      SmsMsg=Replace(SmsMsg,"{$contactman}",RS("ContactMan"))
							  SmsMsg=Replace(SmsMsg,"{$orderid}",rs("orderid"))
							  SmsMsg=Replace(SmsMsg,"{$time}",now)
							  SmsMsg=Replace(SmsMsg,"{$money}",rs("MoneyReceipt")&"")
							  Call SendMobileMsg(Mobile,SmsMsg)
			 End If
			
			 '邮件通知管理员
			 Dim MailBodyStr
			If Setting(210)="1" And Not IsNul(Setting(211)) Then
						MailBodyStr=MailBodyStr&"订单号:"&rs("orderid")&"<br/>"
						MailBodyStr=MailBodyStr&"用户:"&RS("UserName")&"<br/>"
						MailBodyStr=MailBodyStr&"已支付金额:"&RS("MoneyReceipt")&"元<br/>"
						MailBodyStr=MailBodyStr&"收货地址:"&RS("Address")&"<br/>"
						MailBodyStr=MailBodyStr&"收货人姓名:"&RS("ContactMan")&"<br/>"
						MailBodyStr=MailBodyStr&"收货人手机:"&RS("Mobile")&"<br/>"
						MailBodyStr=MailBodyStr&"收货人电话:"&RS("Phone")&"<br/>"
						MailBodyStr=MailBodyStr&"收货人QQ:"&RS("QQ")&"<br/>"
						MailBodyStr=MailBodyStr&"收货人Email:"&RS("Email")&"<br/>"
						Dim Ei,Earr:Earr=split(Setting(211),",")
						For Ei=0 To Ubound(Earr)
					       Call SendMail(Setting(12), Setting(13), Setting(14), Setting(0) & "-订单" & rs("orderid") & "付款成功",Earr(Ei),RS("UserName"), MailBodyStr,Setting(11))
                        Next
			 End If
				   '手机短信通知管理员
			 If Setting(209)="1" And Setting(157)="1" And Not IsNul(Setting(154)) Then
				SmsMsg="新订单,订单号:"&rs("orderid")&",已付金额:"&RS("MoneyReceipt")&"元，请尽快安排发货！"
				Call SendMobileMsg(Setting(154),SmsMsg)
			 End If
		End Sub
		
		
	   '替换可用订单标签
		Function ReplaceOrderLabel(MailContent,RS)
				 MailContent=Replace(MailContent,"{$ContactMan}",RS("ContactMan"))
				 MailContent=Replace(MailContent,"{$InputTime}",RS("InputTime"))
				 MailContent=Replace(MailContent,"{$OrderID}",RS("OrderID"))
				 MailContent=Replace(MailContent,"{$OrderInfo}",OrderDetailStr(RS,0))
				ReplaceOrderLabel=MailContent
		End Function
	   '返回订单详细信息
		Function  OrderDetailStr(RS,flag)
		 Dim DomainStr:DomainStr=GetDomain
		 OrderDetailStr="<table width='100%' border='0' style='text-align:center' cellpadding='2' cellspacing='1' class='border'> "&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr style='text-align:center' class='title'>    <td height='22'><b>订 单 信 息</b>（订单编号：" & RS("ORDERID") & "）</td>  </tr>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr>" & vbcrlf
		 OrderDetailStr=OrderDetailStr & " <td height='25'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "  <table width='100%'  border='0' cellpadding='2' cellspacing='0'> "   & vbcrlf
		 OrderDetailStr=OrderDetailStr & "    <tr class='tdbg'>"
		 OrderDetailStr=OrderDetailStr & "	         <td width='18%'>客户姓名：<font color='red'>" & RS("Contactman") & "</td>      "
		 OrderDetailStr=OrderDetailStr & "			 <td width='20%'>用 户 名：<font color='red'>" & rs("username") & "</td> " &vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='20%'>获赠积分："
		if rs("totalscore")=0 or rs("DeliverStatus")=3 then
			OrderDetailStr=OrderDetailStr & "无"
		else
			if rs("scoretf")=1 then
			OrderDetailStr=OrderDetailStr & "<font color=green>" & rs("totalscore") & "分,已送出</font>"
			else
			OrderDetailStr=OrderDetailStr & "<font color=red>" & rs("totalscore") & "分,未送出</font>"
			end if
		end if
		OrderDetailStr=OrderDetailStr & "			</td> <td width='18%'>购买日期：<font color='red'>" & formatdatetime(rs("inputtime"),2) & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='24%'>下单时间：<font color='red'>" & rs("inputtime") & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "	</tr>"
		OrderDetailStr=OrderDetailStr & "	<tr class='tdbg'> "      
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>需要发票："
			    If RS("NeedInvoice")=1 Then
				  OrderDetailStr=OrderDetailStr & "<Font color=red>√</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "		 </td> "       
		OrderDetailStr=OrderDetailStr & "	 <td width='20%'>已开发票："	
				  If RS("Invoiced")=1 Then
				   OrderDetailStr=OrderDetailStr & "<font color=green>√</font>"
				  Else
				   OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td> "
		OrderDetailStr=OrderDetailStr & "	<td width='20%'>订单状态："	
			if RS("Status")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>等待确认</font>"
				  ElseIf RS("Status")=1 Then
				 OrderDetailStr=OrderDetailStr & "<font color=green>已经确认</font>"
				  ElseIf RS("Status")=2 Then
				 OrderDetailStr=OrderDetailStr & "<font color=#a7a7a7>已结清</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td>"
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>付款情况："
		if rs("paystatus")="100" then
				OrderDetailStr=OrderDetailStr & "<font color=""green"">凭单消费</font>"
		elseif rs("paystatus")="3" then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>退款</font>"
		elseif rs("paystatus")="1" then
				  OrderDetailStr=OrderDetailStr & "<font color=green>已经付清</font>"
		   else	
			     If RS("MoneyReceipt")<=0 Then
				   OrderDetailStr=OrderDetailStr & "<font color=red>等待汇款</font>"
				  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>已收定金</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=green>已经付清</font>"
				  End If
           end if
       OrderDetailStr=OrderDetailStr & "</td>"
	   OrderDetailStr=OrderDetailStr & "        <td width='24%'>物流状态："
				if RS("DeliverStatus")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>未发货</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已发货</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已签收</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  OrderDetailStr=OrderDetailStr & "<font color=#ff6600>退货</font>"
				 End If
	OrderDetailStr=OrderDetailStr & "		</td></tr>    </table> "
    OrderDetailStr=OrderDetailStr & " </td>  </tr> " 
	OrderDetailStr=OrderDetailStr & "   <tr style='text-align:center'>"
	OrderDetailStr=OrderDetailStr & "       <td height='25'>"
	OrderDetailStr=OrderDetailStr & "	   <table width='100%' border='0' style='text-align:center' cellpadding='2' cellspacing='1'>"
	OrderDetailStr=OrderDetailStr & "	           <tr class='tdbg'>"
	OrderDetailStr=OrderDetailStr & "			             <td width='12%' style='text-align:right'>收货人姓名：</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("contactman") & "</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='12%' style='text-align:right'>联系电话：</td> "      
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("phone") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>"
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>"
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' style='text-align:right'>收货人地址：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("address") & "</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' style='text-align:right'>邮政编码：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" &rs("zipcode") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>  "      
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' style='text-align:right'>收货人邮箱：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("email") & " 联系QQ: " & rs("qq") & "</td> "         
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' style='text-align:right'>收货人手机：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("mobile") & "</td>       "
	OrderDetailStr=OrderDetailStr & "			   </tr>"        
	OrderDetailStr=OrderDetailStr & "			   <tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "			              <td width='12%' style='text-align:right'>付款方式：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & ReturnPayMent(rs("PaymentType"),0) & "</td>       "   
	If rs("tocity")="" Then
		OrderDetailStr=OrderDetailStr & "						  <td width='12%' style='text-align:right'>送货方式：</td>"          
		OrderDetailStr=OrderDetailStr & "						  <td width='38%'>免运费订单，由商家指定" 
	Else
		OrderDetailStr=OrderDetailStr & "						  <td width='12%' style='text-align:right'>快递公司：</td>"          
		OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" 
		
		  dim rst,foundexpress
		  Set RST=Server.CreateObject("ADODB.RECORDSET")
		 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
		 If RST.Eof Then
			foundexpress=false
		 Else
			foundexpress=true
		OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
		 End If
		 RST.Close
		 If foundexpress=false Then
		  If DataBaseType=1 Then
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
		  Else
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
		  End If
		  if rst.eof then
		  else
		OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
		  end if
		  rst.close
		 End If
		 set rst=nothing
		
		
		OrderDetailStr=OrderDetailStr & " 发往<span style='color:red'>" & rs("tocity") & "</span>"
    End If
	OrderDetailStr=OrderDetailStr & "				</td></tr> "       
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>  "        
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' style='text-align:right'>发票信息：</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>"
	 'If RS("NeedInvoice")=1 Then OrderDetailStr=OrderDetailStr & replace(rs("InvoiceContent"),chr(10),"<br/>") &"</td>"
	 If RS("NeedInvoice")=0 Then
	  OrderDetailStr=OrderDetailStr & "不需要"
	ElseIf RS("NeedInvoice")=1 Then
	  OrderDetailStr=OrderDetailStr & "发票类型：普通发票<br/>"
	  OrderDetailStr=OrderDetailStr & "单位名称：" &rs("InvoiceContent")
	Else
	  OrderDetailStr=OrderDetailStr & "发票类型：增值税发票<br/>"
	  OrderDetailStr=OrderDetailStr & "单位名称：" &rs("InvoiceContent") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "纳税人识别码："&rs("InvoiceCode") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "注册地址："&rs("InvoiceAddress") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "注册电话："&rs("InvoiceTel") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "开户银行："&rs("Invoicebank") &"<br/>"
	  OrderDetailStr=OrderDetailStr & "银行账号："&rs("Invoicebankcard")
	End If
    OrderDetailStr=OrderDetailStr & "						 <td width='12%' style='text-align:right'>备注/留言：</td>"          
	OrderDetailStr=OrderDetailStr & "							<td width='38%'>" & rs("Remark") & "</td>       "
	OrderDetailStr=OrderDetailStr & "				 </tr>  "  
	OrderDetailStr=OrderDetailStr & "				 </table>"
	if flag=1  then
	 OrderDetailStr=OrderDetailStr & "<div style='text-align:left'><input type='button' onclick=""modifyInfo(" & rs("id") & ")"" class='button' value='修改/打印快递单'/>"
	 IF Rs("OrderType")<>"1" Then
	 OrderDetailStr=OrderDetailStr & " <input type='button' onclick=""modifyproduct('" & rs("orderid") & "')"" class='button' value='修改/添加商品'/>"
	 End If
	 OrderDetailStr=OrderDetailStr & " <input type='button' onclick=""modifytotalprice(" & rs("id") & "," & rs("moneytotal") &")"" class='button' value='修改订单总价'/>"
	 OrderDetailStr=OrderDetailStr & " <input type='button' onclick=""addservice(" & rs("id") & ")"" class='button' value='添加服务跟踪'/> </div>"
	End If
	OrderDetailStr=OrderDetailStr & "			</td>  "
	OrderDetailStr=OrderDetailStr & "		</tr>  "
	
	OrderDetailStr=OrderDetailStr & "		<tr><td>"
	OrderDetailStr=OrderDetailStr & "		<table width='100%' border='0' style='text-align:center' cellpadding='2' cellspacing='1' class='border'> "
	OrderDetailStr=OrderDetailStr & "		  <tr style='text-align:center' class='title' height='25'>  "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>商 品 名 称</b></td> "   
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>单位</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='55'><b>数量</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>参考价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>商城价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>您的价格</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='85'><b>金 额</b></td>   " 
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>赠送积分</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>备注</b></td>  "
	OrderDetailStr=OrderDetailStr & "		  </tr> "
			 Dim attributecart,TotalPrice:totalprice=0
			 Dim RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			 Dim SQLStr:SQLStr="Select * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id"
			   RSI.Open SQLStr,conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
			  Else
			   Dim OrderType:OrderType=ChkClng(RS("OrderType"))
			   Do While Not RSI.Eof
			    If OrderType=1 Then
				 SQLStr="select top 1 subject from ks_groupbuy where id=" & rsi("proid")
				ELSE
				 SqlStr="select top 1 title from ks_product where id=" & rsi("proid")
				End If
			   If Conn.execute(SqlStr).eof Then
			   		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"    
					OrderDetailStr=OrderDetailStr & "	  <td colspan='9'>该商品已被删除！</td>"   
					OrderDetailStr=OrderDetailStr & "	  </tr>"   
			   Else
			  attributecart=rsi("attributecart")
			  if not isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
			  Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  If OrderType=1 Then
			    SqlStr="Select top 1 Subject as title,'件' as unit,0 as IsLimitBuy,0 as LimitBuyPrice,0 as LimitBuyPayTime From KS_GroupBuy Where ID=" & RSI("ProID")
			  Else
			    SqlStr="Select top 1 I.Title,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID")
			  End If
			  RSP.Open SqlStr,conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(换购)"
				  Else
				     If RSP("LimitBuyPayTime") Then
				  	   If LimitBuyPayTime="" Then
					   LimitBuyPayTime=RSP("LimitBuyPayTime")
					   ElseIf LimitBuyPayTime>RSP("LimitBuyPayTime") Then
						LimitBuyPayTime=RSP("LimitBuyPayTime")
					   End If
					 End If
				  End If
				  If RSI("IsLimitBuy")="1" Then  title=title & "<span style='color:green'>(限时抢购)</span>"
				  If RSI("IsLimitBuy")="2" Then title=title & "<span style='color:blue'>(限量抢购)</span>"
			  End If
			  RSP.Close:Set RSP=Nothing
			  
		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>" 
		If OrderType=0 Then   
		OrderDetailStr=OrderDetailStr & "	   <td width='*'><a href='" & DomainStr & "item/show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & title & "</a>" & attributecart & "</td>    "
		Else
		OrderDetailStr=OrderDetailStr & "	   <td width='*'><a href='" & DomainStr & "shop/groupbuyshow.asp?id=" & RSi("proid") & "' target='_blank'>" & title & "</a>" & attributecart & "</td>    "
		End If
		OrderDetailStr=OrderDetailStr & "	   <td width='45' style='text-align:center'>"& Unit & "</td>               <td width='55' style='text-align:center'>" & rsi("amount") &"</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' style='text-align:center'>" & formatprice(rsi("price_original")) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' style='text-align:center'>" & formatprice(rsi("realprice")) & "</td>    "
		
		if flag=1 then
			OrderDetailStr=OrderDetailStr & "	   <td width='65' style='text-align:center'>" &  formatprice(rsi("realprice")) 
			If RSI("IsChangedBuy")<>1 And RSI("IsLimitBuy")<>"1" And RSI("IsLimitBuy")<>"2" Then
			OrderDetailStr=OrderDetailStr& " <a href=""javascript://"" onclick=""modifyPrice(event,'" & title & "','" & rs("orderid") & "'," & rsi("id")&"," & rsi("realprice") & ")""><font color=blue>改</font></a>"
			End If
			OrderDetailStr=OrderDetailStr & "</td>    "
		else
		OrderDetailStr=OrderDetailStr & "	   <td width='65' style='text-align:center'>" &  formatprice(rsi("realprice")) & "</td>    "
		end if
		OrderDetailStr=OrderDetailStr & "	   <td width='120' style='text-align:right'>"
		if ChkClng(rs("UseScoreisshop"))>0 then
			OrderDetailStr=OrderDetailStr & ""& ChkClng(rs("UseScoreisshop")) &" 积分 + "	&  formatprice(rsi("realprice")) 
		else
			OrderDetailStr=OrderDetailStr  &  formatprice(rsi("realprice")*rsi("amount")) 		
		end if
		OrderDetailStr=OrderDetailStr  & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' style='text-align:center'>" & chkclng(rsi("score")) & " 分</td>    "'bus改动
		OrderDetailStr=OrderDetailStr & "	   <td style='text-align:center' width='45'>" & rsi("Remark") & "</td>  "
		OrderDetailStr=OrderDetailStr & "	   </tr> " 
		
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  '取得捆绑销售商品
		     end if
			    TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
			End If
			
			
			OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '超值礼包
			
			
		OrderDetailStr=OrderDetailStr & "	   <tr class='tdbg' height='30' > "   
		OrderDetailStr=OrderDetailStr & "	    <td colspan='6' style='text-align:right'><b>合计：</b></td> "   
		OrderDetailStr=OrderDetailStr & "		<td style='text-align:right'>"
		if ChkClng(rs("UseScoreisshop"))>0 then
			OrderDetailStr=OrderDetailStr & "  "& ChkClng(rs("UseScoreisshop")) & " 积分 + "  
		end if
		OrderDetailStr=OrderDetailStr & "<b>" &  formatprice(totalprice) & "</b></td>    "
		OrderDetailStr=OrderDetailStr & "		<td colspan='3'> </td>  "
		OrderDetailStr=OrderDetailStr & "	  </tr>    "
		OrderDetailStr=OrderDetailStr & "	  <tr class='tdbg'>"
       OrderDetailStr=OrderDetailStr & "         <td colspan='4'>付款方式折扣率：" & rs("Discount_Payment") & "%&nbsp;&nbsp;" 
	   If RS("Weight")>0 Then
	   OrderDetailStr=OrderDetailStr & "重量：" & rs("weight") & " KG"
	   End If
	   OrderDetailStr=OrderDetailStr & "&nbsp;&nbsp;运费：" & rs("Charge_Deliver")&" 元&nbsp;&nbsp;&nbsp;&nbsp;税率：" & Setting(65) &"%&nbsp;&nbsp;&nbsp;&nbsp;价格含税："
				IF Setting(64)=1 Then 
				   OrderDetailStr=OrderDetailStr & "是"
				  Else
				   OrderDetailStr=OrderDetailStr & "不含税"
				  End If
				  Dim TaxMoney
				  Dim TaxRate:TaxRate=Setting(65)
				 If Setting(64)=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+TaxRate/100

				OrderDetailStr=OrderDetailStr & "<br>订单金额：(" & rs("MoneyGoods") & "×" & rs("Discount_Payment") & "%＋"&rs("Charge_Deliver") & ")×"
				if Setting(64)=1 Or rs("NeedInvoice")=0 Then OrderDetailStr=OrderDetailStr & "100%" Else OrderDetailStr=OrderDetailStr & "(1＋" & TaxRate & "%)" 
				OrderDetailStr=OrderDetailStr & "＝" &  formatprice(rs("NoUseCouponMoney")) & "  "
				if ChkClng(rs("UseScoreisshop"))>0 then
					OrderDetailStr=OrderDetailStr & "  <font color=""#006600"">本单为积分兑换订单，已支付积分<font color=""#FF0000"">"& ChkClng(rs("UseScoreisshop")) & "</font> 积分</font>"  
				end if
				OrderDetailStr=OrderDetailStr & "</td>"
    OrderDetailStr=OrderDetailStr & "<td  colspan='3' style='text-align:right'><b>订单金额：</b> " &  formatprice(rs("NoUseCouponMoney")) & " <br>"

	If ChkClng(RS("CouponUserID"))<>0 And RS("UseCouponMoney")>0 Then
	OrderDetailStr=OrderDetailStr & "<b>使用优惠券：</b> <font color=#ff6600>" &  formatprice(RS("UseCouponMoney")) & " </font><br>"
    ElseIf RS("UseScoreMoney")<>"0" Then
	OrderDetailStr=OrderDetailStr & "<b>花费<font color=green>" &RS("UseScore") & "</font>积分抵扣了<font color=#ff6600>" &  formatprice(RS("UseScoreMoney")) & "</font><br>"
	End If
	
	
	
	OrderDetailStr=OrderDetailStr & "<b>应付金额：</b> <span style=""font-weight:bold;font-size:15px;color:#ff6600"">" &  formatprice(rs("MoneyTotal")) & "</span>  "
	OrderDetailStr=OrderDetailStr &"</td>"
    OrderDetailStr=OrderDetailStr & "<td colspan='3' align='left'><b>已付款：</b> "
	if isnumeric(rs("MoneyReceipt")) then
	 OrderDetailStr=OrderDetailStr & "<font color=red>" &  formatprice(rs("MoneyReceipt")) & "</font></b>"
	else
	 OrderDetailStr=OrderDetailStr & "<font color=red>0</font></b>"
	end if
	
	If RS("MoneyReceipt")<RS("MoneyTotal") Then
	OrderDetailStr=OrderDetailStr & "<br><B>尚欠款：<font color=blue>" &  formatprice(RS("MoneyTotal")-RS("MoneyReceipt")+rs("UseCouponMoney")) &"</B>"
	End If
	OrderDetailStr=OrderDetailStr & "</td></tr></table></td>  "
	OrderDetailStr=OrderDetailStr & "</tr>"  
	OrderDetailStr=OrderDetailStr & "     <tr><td style='text-align:left'><br><b>注：</b><br/>1、“<font color='blue'>参考价</font>”指商品的市场参考价，“<font color='green'>商城价</font>”指本商城的销售价格，“<font color='red'>您的价格</font>”指根据会员灯级折扣系统自动算出的优惠价。商品的最终销售价格以“您的价格”为准。“订单金额”指系统自动算出来的价格，本订单的最终价格以“<font color=#ff6600>应付金额</font>”为准。<br>2、积分的赠送为结清订单后送出；<br/>3、限时抢购商品、特惠礼包、使用优惠券及使用积分抵扣费用的订单将不再赠送积分。"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"

	If not conn.execute("select top 1 * from ks_orderitem where orderid='" & RS("OrderID") &"' and islimitbuy<>0").eof Then
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:red;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单是限时/限量抢购订单,限制下单后" & LimitBuyPayTime & "小时之内必须付款,即在[" & DateAdd("h",LimitBuyPayTime,RS("InputTime")) & "]之前用户没有付款,本订单自动作废。</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	If RS("DeliverStatus")=1 Then
	 Dim RSD,DeliverStr
	 Set RSD=Conn.Execute("Select Top 1 * From KS_LogDeliver Where DeliverType=1 And OrderID='" & RS("OrderID") & "'")
	 If Not RSD.Eof Then
	  DeliverStr="快递公司:" & RSD("ExpressCompany") & " 物流单号:" & RSD("ExpressNumber") & " 发货日期:" & RSD("DeliverDate") & " 发货经手人:" & RSD("HandlerName")
	 End If
	 RSD.Close : Set RSD=Nothing
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='text-align:left;margin:10px;color:blue;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单已发货。" & DeliverStr & "</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	
	OrderDetailStr=OrderDetailStr & "	</table>"
 End Function
 
 '返回格式化后的价格,带单位
 Function FormatPrice(Price)
  if Not IsNumeric(Price) Then FormatPrice="¥0.00元":Exit Function
  If Price<1 and Price>0 Then FormatPrice="¥0"& Price &"元":Exit Function
  if Price<0 and Price>-1 Then FormatPrice="¥-0"& abs(Price) &"元":Exit Function
  FormatPrice="¥" & formatnumber(Price,2) &"元"
 End Function
 
 '取得格式化后的价格，不带单位
 Function GetPrice(price)
     If Not IsNumeric(price) Then GetPrice="0.0": Exit Function
	 If Price=0 Then
	   GetPrice="0.0"
	 ElseIf Price<1 Then
	   GetPrice="0"&Price
	 Else
	   price=round(price,2)
	   if instr(price,".")=0 then price=price&".00"
	   GetPrice=price
	 End If
   End Function

'取得捆绑销售商品
Dim ProIds
Function GetBundleSalePro(ByRef TotalPrice,ProID,OrderID)
  If FoundInArr(ProIDS,ProID,",")=true Then Exit Function
  ProIds=ProIDs & "," & ProID
  Dim Str,RS,XML,Node
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "Select I.Title,I.Unit,O.* From KS_OrderItem O inner join KS_Product I On O.ProID=I.ID Where O.SaleType=6 and BundleSaleProID=" & ProID & " and o.OrderID='" & OrderID & "' order by O.id",conn,1,1
  If Not RS.Eof Then
    Set XML=RsToXml(rs,"row","")
  End If
  RS.Close:Set RS=Nothing
  If IsObject(XML) Then
	     str=str & "<tr height=""25"" align=""left""><td colspan=9 style=""color:green"">&nbsp;&nbsp;选购捆绑促销:</td></tr>"
       For Each Node In Xml.DocumentElement.SelectNodes("row")
         str=str & "<tr>"
		 str=str &" <td style='color:#999999'>&nbsp;" & Node.SelectSingleNode("@title").text &"</td>"
		 str=str &" <td style='text-align:center'>" & Node.SelectSingleNode("@unit").text &"</td>"
		 str=str &" <td style='text-align:center'>" & Node.SelectSingleNode("@amount").text &"</td>"
		 str=str &" <td style='text-align:center'>" & formatnumber(Node.SelectSingleNode("@price_original").text,2,-1) &"</td>"
		 str=str &" <td style='text-align:center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td style='text-align:center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td style='text-align:right'>" & formatnumber(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2,-1) &"</td>"
		 str=str &" <td style='text-align:center'>" & Node.SelectSingleNode("@serviceterm").text &"</td>"
		 str=str &" <td style='text-align:center'>" & Node.SelectSingleNode("@remark").text &"</td>"
		 str=str & "</tr>"
		 TotalPrice=TotalPrice +round(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2) 
       Next
  End If
  GetBundleSalePro=str
End Function

 '得到超值礼包
 Function GetPackage(ByRef TotalPrice,OrderID)
	    If IsNul(OrderID) Then Exit Function
		Dim RS,RSB,GXML,GNode,str,n,Price
		Set RS=Conn.Execute("select packid,OrderID from KS_OrderItem Where SaleType=5 and OrderID='" & OrderID & "' group by packid,OrderID")
		If Not RS.Eof Then
		 Set GXML=RsToXml(Rs,"row","")
		End If
		RS.Close : Set RS=Nothing
		If IsOBJECT(GXml) Then
		   FOR 	Each GNode In GXML.DocumentElement.SelectNodes("row")
		     Set RSB=Conn.Execute("Select top 1 * From KS_ShopPackAge Where ID=" & GNode.SelectSingleNode("@packid").text)
			 If Not RSB.Eof Then
					  
						Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
						RSS.Open "Select a.title,a.Price_Member,a.Price,b.* From KS_Product A inner join KS_OrderItem b on a.id=b.proid Where b.SaleType=5 and b.packid=" & GNode.SelectSingleNode("@packid").text & " and  b.orderid='" & OrderID & "'",Conn,1,1
						  str=str & "<tr class='tdbg' height=""25"" align=""center""><td colspan=2><strong><a href='" & GetDomain & "shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a></strong></td>"
						  n=1
						  Dim TotalPackPrice,tempstr,i
						  TotalPackPrice=0 : tempstr=""
						Do While Not RSS.Eof
						 
						  For I=1 To RSS("Amount") 
							  '得到单件品价格 
							  If RSS("AttrID")<>0 Then 
							  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RSS("AttrID"))
							  If Not RSAttr.Eof Then
								Price=RSAttr("Price")
							  Else
								Price=RSS("Price_member")
							  End If
							  RSAttr.CLose:Set RSAttr=Nothing
							 Else
								Price=RSS("Price_member")
							 End If
							
							   TotalPackPrice=TotalPackPrice+Price
							  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
							  n=n+1
						  Next
						  RSS.MoveNext
						Loop
						
						str=str &"<td>1</td><td>￥" & TotalPackPrice & "</td><td>" & rsb("discount") & "折</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>---</td><td>---</td>"
					   
						str=str & "</tr><tr><td align='left' colspan=9>选择的套装详细如下:<br/>" & tempstr & "</td></tr>" 
						
						TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1))   '将礼包金额加入总价
						
						RSS.Close
						Set RSS=Nothing
					
			End If
			RSB.Close
		   Next
			
	    End If
		GetPackage=str
		
End Function
	'====================================商城相关_end=============================================
	
	
	
	
	
	
	

  
  '======================================会员相关函数====================================
    '取得会员组选项--下拉列表  参数：Selected--默认选项
	Public Function GetUserGroup_Option(Selected)
	    Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr,nbsp
		LoadUserGroup()
		For Each Node In Application(SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
		  SpaceStr="" 
			  TJ=Node.SelectSingleNode("@tj").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "──" 
				 Next
			  End If
              If ChkClng(Selected)=ChkClng(Node.SelectSingleNode("@id").text) Then
			  TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@id").text & "' selected>" & SpaceStr & Node.SelectSingleNode("@groupname").text & " </option>"
			  Else
			  TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@id").text & "'>" & SpaceStr & Node.SelectSingleNode("@groupname").text & " </option>"
			  End If
		Next
	   GetUserGroup_Option=TreeStr
	End Function
	
	'取得会员组选项--下载列表 参数：SelectArr--默认选择项以","隔开,RowNum--每行显示选项数
	Public Function GetUserGroup_CheckBox(OptionName,SelectArr,RowNum)
	   Dim n:n=0
	   IF RowNum<=0 Then RowNum=3
	   
	   LoadUserGroup()
	    Dim Node,str,i,DocNode
		Set DocNode=Application(SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
		str="<div style=""clear:both""></div><select name=""" & OptionName &""" id=""" & replace(replace(OptionName,"(",""),")","") &""" multiple style=""width:380px;height:200px"">"
		Dim TJ,SpaceStr,k,Nstr
		For I=0 To DocNode.Length-1
		    Set Node=DocNode.Item(i)
			SpaceStr="&nbsp;&nbsp;"
			TJ=Node.SelectSingleNode("@tj").text
			For k = 1 To TJ - 1
			 SpaceStr = SpaceStr & "──"
			Next
		   
			 If TJ=1 Then
			  Nstr="+ " & Node.SelectSingleNode("@groupname").text &""
			 Else
			  Nstr=Node.SelectSingleNode("@groupname").text 
			 End If
			 NStr=SpaceStr & Nstr
		    If FoundInArr(SelectArr,Node.SelectSingleNode("@id").text,",")<>0 Then
			 str=str & "<option value=""" & Node.SelectSingleNode("@id").text & """ selected>" & Nstr & "</option>" &vbcrlf
			Else
			 str=str & "<option value=""" & Node.SelectSingleNode("@id").text & """>" & Nstr & "</option>" &vbcrlf
			End If
		Next
		str=str &"</select><br/><span class=""tips"">TIPS:按Ctrl或Shift可以多选,<a href=""javascript:;"" onclick=""$('#" & replace(replace(OptionName,"(",""),")","") &"').val('');"">[点此]</a>取消选中项。</span>"
	   GetUserGroup_CheckBox=str
	  
	End Function
	
	
	'取得角色选项--下载列表 参数：SelectArr--默认选择项以","隔开
	Public Function GetRoleList(OptionName,SelectArr)
	   Dim n:n=0
	    Dim str,i
		str="<div style=""clear:both""></div><select name=""" & OptionName &""" id=""" & OptionName &""" multiple style=""width:380px;height:200px"">"
		Dim RS:Set RS=Conn.Execute("Select ID,GroupName From KS_UserGroup Where [Type]=2 order by id")
		Dim SpaceStr,k,Nstr
		Do While Not RS.Eof
			SpaceStr="&nbsp;&nbsp;"
			 NStr=SpaceStr & "+ " & rs("groupname")
		    If FoundInArr(SelectArr,rs("id"),",")<>0 Then
			 str=str & "<option value=""" & rs("id") & """ selected>" & Nstr & "</option>" &vbcrlf
			Else
			 str=str & "<option value=""" & rs("id") & """>" & Nstr & "</option>" &vbcrlf
			End If
		RS.MoveNext
		Loop
		RS.Close
		Set RS=Nothing
		str=str &"</select><br/><span class=""tips"">TIPS:按Ctrl或Shift可以多选,<a href=""javascript:;"" onclick=""$('#" & OptionName &"').val('');"">[点此]</a>取消选中项。</span>"
	   GetRoleList=str
	  
	End Function
	
	 
  	'取得用户组名称
	Public Function GetUserGroupName(GroupID)
	 On Error Resume Next
	 GetUserGroupName=Conn.Execute("Select top 1 GroupName From KS_UserGroup Where ID=" & GroupID)(0)
	 if err then GetUserGroupName=""
	End Function
    
	'会员投稿文章，图片，下载等增加积分,发送站内短信操作
	'参数ChannelID-频道ID,UserName---用户名称,InfoTitle---投稿的主题
	Public Sub SignUserInfoOK(ChannelID,UserName,InfoTitle,InfoID)
	    IF Not IsNumeric(ChannelID) Then Exit Sub
	    Dim ClientName,GroupID,RSObj:Set RSObj=Conn.Execute("Select top 1 RealName,GroupID,Mobile From KS_User Where UserName='" & UserName & "'")
		IF Not RSObj.Eof Then
					ClientName=RSObj(0):If ClientName="" Then ClientName=UserName
					GroupID=RSObj(1)
					Dim ScoreRate:ScoreRate=ChkClng(U_S(GroupID,3))
					Dim PointRate:PointRate=ChkClng(U_S(GroupID,4))
					Dim MoneyRate:MoneyRate=ChkClng(U_S(GroupID,5))
					Dim Mobile:Mobile=RSObj(2)
					
					
					
					'成功则发送站内通知信件
					Dim Sender:Sender=Setting(0)
					Dim Title:Title="恭喜，您发表的" & C_S(ChannelID,3) & "[" & InfoTitle & "]已通过审核！！！"
					Dim Message:Message="" & C_S(ChannelID,3) & "标题：" & InfoTitle &" 已通过审核!<br>"
					
					If Conn.Execute("Select top 1 * From KS_LogMoney Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID).Eof And C_S(ChannelID,18)*MoneyRate<>0 Then    '没有记录才给增加金钱
					 If C_S(ChannelID,18)>0 Then
					  Message = Message & "获得金钱：<font color=red>" & C_S(ChannelID,18)*MoneyRate & "</font> 元人民币<br>"
					 ElseIf C_S(ChannelID,18)<0 Then
					  Message = Message & "消耗金钱：<font color=red>" & Abs(C_S(ChannelID,18))*MoneyRate & "</font> 元人民币<br>"
					 End IF
					End If
					 
					If Conn.Execute("Select top 1 * From KS_LogPoint Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " and ContributeFlag=1").Eof And C_S(ChannelID,19)*PointRate<>0 Then
					 If C_S(ChannelID,19)>0 Then
					  Message = Message & "获得" & Setting(45) & "：<font color=red>" & C_S(ChannelID,19)*PointRate & "</font> " & Setting(46) & Setting(45) & "<br>"
					 ElseIf C_S(ChannelID,19)<0 Then
					  Message = Message & "消耗" & Setting(45) & "：<font color=red>" & Abs(C_S(ChannelID,19))*PointRate & "</font> " & Setting(46) & Setting(45) & "<br>"
					 End If
					End If
					 
					If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID).Eof And C_S(ChannelID,20)*ScoreRate<>0 Then    '没有记录才给增加积分
						 If C_S(ChannelID,20)>0 Then
						  Message = Message & "获得积分：<font color=red>" & C_S(ChannelID,20)*ScoreRate & "</font> 分积分<br>"
						 ElseIf C_S(ChannelID,20)<0 Then
						  Message = Message & "消耗积分：<font color=red>" & Abs(C_S(ChannelID,20))*ScoreRate & "</font> 分积分<br>"
						 End If
					End If
					
					Message = Message & "<br />备注：此信息由系统自动发布，请不要回复！！！"
					If C_S(ChannelID,19)<0 Then  
					Call PointInOrOut(ChannelID,InfoID,UserName,2,-C_S(ChannelID,19)*PointRate,"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生",1)            
					Else
					Call PointInOrOut(ChannelID,InfoID,UserName,1,C_S(ChannelID,19)*PointRate,"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生",1)            
					End If
					
					If C_S(ChannelID,20)<0 Then
					 Call ScoreInOrOut(UserName,2,-C_S(ChannelID,20)*ScoreRate,"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生",ChannelID,InfoID)            
					Else
					 Call ScoreInOrOut(UserName,1,C_S(ChannelID,20)*ScoreRate,"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生",ChannelID,InfoID)            
					End If
					If C_S(ChannelID,18)<0 Then
					Call MoneyInOrOut(UserName,ClientName,-C_S(ChannelID,18)*MoneyRate,4,2,SqlNowString,"0","系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生",ChannelID,InfoID,0)
					Else
					Call MoneyInOrOut(UserName,ClientName,C_S(ChannelID,18)*MoneyRate,4,1,SqlNowString,"0","系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生",ChannelID,InfoID,0)
					End If
					If ChkClng(U_S(GroupID,10))=1 Then Call SendInfo(UserName,Sender,Title,Message)
					
					'发送手机短信
					If Not IsNul(Mobile) Then 
					   Dim SmsContent:SmsContent=Split(Setting(155)&"∮∮∮∮∮∮∮∮","∮")(4)
					   If Not IsNul(SmsContent) And Setting(157)="1" Then
					      SmsContent=Replace(SmsContent,"{$inputer}",UserName)
					      SmsContent=Replace(SmsContent,"{$title}",InfoTitle)
						  Call SendMobileMsg(Mobile,SmsContent)
					   End If
					End If
		End IF
		RSObj.Close:Set RSObj=Nothing
	End Sub
	'功能:会员点券明细出入函数	                                                       '参数:Channelid-模块ID,InfoID-信息ID，UserName-用户名,InOrOutFlag-操作类型1收入2支出,Point-交易点数,User-操作员,Descript-操作备注
	Public Function PointInOrOut(ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript,ContributeFlag)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Point) Or Point=0 Then PointInOrOut=false:Exit Function
	  Dim PointParam,CurrPoint
	  If InOrOutFlag=1 Then 
	     PointParam="Set Point=Point+" & Point
	  ElseIF InOrOutFlag=2 Then
	     PointParam="Set Point=Point-" & Point
	  Else
	    PointInOrOut=false:Exit Function
	  End If
	  If (Conn.Execute("Select top 1 * From KS_LogPoint Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " And InOrOutFlag=" & InOrOutFlag).Eof) Or (ChannelID=0 And InfoID=0) or ContributeFlag=0 Then
		  on error resume next
		  CurrPoint=Conn.Execute("Select top 1 Point From KS_User Where UserName='" & UserName & "'")(0)
		  If Not IsNumeric(CurrPoint) Then PointInOrOut=false:Exit Function
		  If InOrOutFlag=2 and CurrPoint-Point<0 Then
		   PointInOrOut=false:Exit Function
		  End If
		  Conn.Execute("Update KS_User " & PointParam & " Where UserName='" & UserName & "'")
		  CurrPoint=Conn.Execute("Select top 1 Point From KS_User Where UserName='" & UserName & "'")(0)
		  Conn.Execute("Insert into KS_LogPoint(ChannelID,InfoID,UserName,InOrOutFlag,Point,Times,[User],Descript,Adddate,IP,CurrPoint,ContributeFlag) values(" & ChannelID & "," & InfoID & ",'" & UserName & "',"& InOrOutFlag & "," & Point & ",1,'" & replace(User,"'","""") & "','" & replace(Descript,"'","""") & "'," & SqlNowString & ",'" & replace(getip,"'","""") & "'," & CurrPoint & "," & ContributeFlag & ")")
	  End If
	  IF Err Then PointInOrOut=false Else PointInOrOut=true
	  Session(SiteSN&"UserInfo")=""
	End Function
	
	'功能:会员积分明细出入函数	
	'参数:UserName-用户名,InOrOutFlag-操作类型1收入2支出,Score-交易点数,User-操作员,Descript-操作备注
	Public Function ScoreInOrOut(UserName,InOrOutFlag,Score,User,Descript,ChannelID,InfoID)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Score) Or Score=0 Then ScoreInOrOut=false:Exit Function
	  Dim ScoreParam,CurrScore
	  If InOrOutFlag=1 Then 
	     ScoreParam="Set Score=Score+" & Score
		 '判断有没有到达每天增加的总限
		 If ChkClng(Setting(165))<>0 Then
		  Dim TodayScore:TodayScore=ChkClng(Conn.Execute("select sum(Score) from ks_logscore where InOrOutFlag=1 and year(adddate)=year(" & SQLNowString & ") and month(adddate)=month(" & SQLNowString & ") and day(adddate)=day(" & SQLNowString & ") and username='" & UserName & "'")(0))
		  If TodayScore>=ChkClng(Setting(165)) Then Exit Function
		 End If
	  ElseIF InOrOutFlag=2 Then
	     ScoreParam="Set Score=Score-" & Score
	  Else
	    ScoreInOrOut=false:Exit Function
	  End If
	  If (Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " And InOrOutFlag=" & InOrOutFlag).Eof) Or (ChannelID=0 And InfoID=0)  or session("scoremustin")="true" Then
		  session("scoremustin")="false"
		  on error resume next
		  CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & UserName & "'")(0)
		  If Not IsNumeric(CurrScore) Then ScoreInOrOut=false:Exit Function
		  If InOrOutFlag=2 and CurrScore-Score<0 and isnul(Session("ScoreHasUse")) Then
		   ScoreInOrOut=false:Exit Function
		  End If
		  if Session("ScoreHasUse")<>"" Then '累计总消费积分
		   if Session("ScoreHasUse")="+" then
		   Conn.Execute("Update KS_User Set ScoreHasUse=ScoreHasUse+" & Score &" Where UserName='" & UserName & "'")
		   else
		   Conn.Execute("Update KS_User Set ScoreHasUse=ScoreHasUse-" & Score &" Where UserName='" & UserName & "'")
		   end if
		   Session("ScoreHasUse")=""
		  Else
		   Conn.Execute("Update KS_User " & ScoreParam & " Where UserName='" & UserName & "'")
		  End If
		  CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & UserName & "'")(0)
		  Dim AvailableScore:AvailableScore=CurrScore-ChkClng(Conn.Execute("Select top 1 ScoreHasUse From KS_User Where UserName='" & UserName & "'")(0))
		  If AvailableScore<0 Then AvailableScore=0

		  
		  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,AvailableScore,[User],Descript,Adddate,IP,ChannelID,InfoID,[Times]) values('" & UserName & "',"& InOrOutFlag & "," & Score & ","&CurrScore & "," & AvailableScore & ",'" & replace(User,"'","""") & "','" & replace(Descript,"'","""") & "'," & SqlNowString & ",'" & replace(getip,"'","""") & "'," & ChannelID &"," & InfoID &",1)")
	  End If
	  IF Err Then ScoreInOrOut=false Else ScoreInOrOut=true
	  Session(SiteSN&"UserInfo")=""
	End Function
	
	'功能:资金明细出入函数	                 
	'参数:UserName-用户名,ClientName-客户姓名,Money-金钱,MoneyType-类型,InOrOutFlag-操作类型1收入2支出,PayTime-汇款日期,OrderID-订单号,Inputer-操作员,Remark-操作备注,MustIn -强行写入 1是 0否
	Public Function MoneyInOrOut(UserName,ClientName,Money,MoneyType,InorOutFlag,PayTime,OrderID,Inputer,Remark,ChannelID,InfoID,MustIn)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Money) Or Money="0" Then MoneyInOrOut=false:Exit Function
	  Dim MoneyParam,CurrMoney,PaymentPlat
	  If InOrOutFlag=1 Then 
	     MoneyParam="Set [Money]=[Money]+" & Money
	  ElseIF InOrOutFlag=2 Then
	     MoneyParam="Set [Money]=[Money]-" & Money
	  Else
	    MoneyInOrOut=false:Exit Function
	  End If
	  If (Conn.Execute("Select top 1 * From KS_LogMoney Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " And IncomeOrPayOut=" & InOrOutFlag).Eof) Or (ChannelID=0 And InfoID=0) Or MustIn=1 Then
		  on error resume next
		  CurrMoney=Conn.Execute("Select top 1 Money From KS_User Where UserName='" & UserName & "'")(0)
		  If Not IsNumeric(CurrMoney) Then MoneyInOrOut=false:Exit Function
		  If InOrOutFlag=2 and CurrMoney-Money<0 Then
		   MoneyInOrOut=false:Exit Function
		  End If
		  dim RowsAffected
		  Conn.Execute "Update KS_User " & MoneyParam & " Where UserName='" & UserName & "'",RowsAffected,&H0001
          if RowsAffected>=1 Then
			  CurrMoney=Conn.Execute("Select top 1 Money From KS_User Where UserName='" & UserName & "'")(0)
			  PaymentPlat=ChkClng(S("PaymentPlat")): If PayMentPlat=0 And Request("c_memo1")<>"" Then PayMentPlat=ChkClng(Split(Request("c_memo1"),"|")(0))  '判断是不是云网支付的
              if isnul(PaymentPlat) then PaymentPlat=16  '微信支付    
			  Conn.Execute("Insert into KS_LogMoney([UserName],[ClientName],[Money],[MoneyType],[IncomeOrPayOut],[OrderID],[Remark],[PayTime],[LogTime],[Inputer],[IP],[CurrMoney],[ChannelID],[InfoID],[Times],[PaymentID] ) values('" & UserName & "','" & ClientName & "'," & Money & "," & MoneyType & ","& InOrOutFlag & ",'" & OrderID & "','" & replace(Remark,"'","""") & "'," & SqlNowString & "," &SqlNowString & ",'" & replace(inputer,"'","""") & "','" & replace(getip,"'","""") & "'," & CurrMoney & "," & ChannelID & "," & InfoID & ",1," & PaymentPlat &")")
		 end if
	  End If
	  IF Err Then MoneyInOrOut=false Else MoneyInOrOut=true
	  Session(SiteSN&"UserInfo")=""
	End Function
	'会员有效期明细出入函数
	'参数:UserName,InOrOutFlag,Edays,User,Descript
	Function EdaysInOrOut(UserName,InOrOutFlag,Edays,User,Descript)
		 If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Edays) Or Edays=0 Then EdaysInOrOut=false:Exit Function
		  Conn.Execute("insert into KS_LogEdays(UserName,InOrOutFlag,Edays,[user],descript,adddate,ip) values('" & UserName & "'," & InOrOutFlag & "," & Edays & ",'" & user & "','" & replace(descript,"'","""") & "'," & SqlNowString & ",'" & getip & "')")
		  IF Err Then EdaysInOrOut=false Else EdaysInOrOut=true
	 End Function
	'发送站内信息
	'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
	Public Function SendInfo(Incept,Sender,title,Content)
	  Dim GroupID,Max_Num,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "select top 1 groupid from ks_user where username='" & Incept & "'",conn,1,1
	  if RS.Eof And RS.Bof Then
	    RS.Close:Set RS=Nothing
		SendInfo=false:exit Function
	  End If
	  GroupID=RS(0)
	  RS.Close:Set RS=Nothing
	  Max_Num=ChkClng(U_S(GroupID,13))
	  If Max_Num<>0 And ChkClng(conn.execute("select count(*) from KS_Message where IsSend=1 and delR=0 and Incept='"&Incept&"'")(0))>=ChkClng(Max_Num) Then
		SendInfo=false:exit Function
	  End If
	  
	  Content=ScriptHtml(Content,"script",2)  '过滤脚本
	  Conn.Execute("insert Into KS_Message(Incept,Sender,Title,Content,SendTime,Flag,IsSend,DelR,DelS,AutoDelDays) values('" & Incept & "','" & Sender & "','" & CheckXSS(replace(Title,"'","""")) & "','" & replace(Content,"'","''") & "'," & SqlNowString & ",0,1,0,0," & ChkClng(S("AutoDelDays"))&")")
	  SendInfo=true
	End Function
	
	Function WordFilter()
	   WordFilter=ReadFromFile(Setting(3) & "config/filter.txt")
	End Function
	
	'过滤非法字符
	Public Function FilterIllegalChar(ByVal Content)
	   If IsNul(Content) Then Exit Function
	   Dim K
		dim sarr:sarr=split(WordFilter,",")
        for k=0 to ubound(sarr)
		   Content=Replace(Content,sarr(k),"***")
		next
        FilterIllegalChar=Content 
	End Function
	
	 '从字符字组中过滤得重复项
	 public Function FilterRepeatInArray(stringArray, splitStr)
            If stringArray="" Then FilterRepeatInArray="" : Exit Function
            Dim Str,I
            Dim arr:arr = split(stringArray, splitStr)
			For I=0 To Ubound(Arr)
			 if arr(i)<>"" then
			  If Not FoundInArr(str, arr(i), splitStr) Then
			     if str="" then
				   str=arr(i)
				 else
				   str=str & splitStr & arr(i)
				 end if
			  End If
			 End If
			Next
			FilterRepeatInArray=str
   End Function
   

	
	'删除微博相关
	 sub delweibo(copyfrom,topicids)
		Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open "select b.id as msgid,a.id,a.username From KS_UserLogR a left join ks_userlog b on a.msgid=b.id Where b.CopyFrom='" & copyfrom & "' and b.infoid in(" & topicids &")",conn,1,1
			  do while not rs.eof
				  Dim UserName:UserName=RS("UserName")
				  Dim MsgId:MsgId=RS("MsgId")
				  Dim Id:Id=RS("Id")
					Conn.Execute("Delete From KS_UserLog Where ID=" & MsgId)
					Conn.Execute("Delete From KS_UserLogCMT Where MsgID=" & MsgId)
					Conn.Execute("Delete From KS_UserLogR Where ID=" & id)
					Conn.Execute("Update KS_User set MsgNum=MsgNum-1  Where UserName='" & UserName &"' and MsgNum>=1")
			   rs.movenext
			  loop
			  rs.close
			  set rs=nothing
	End Sub	
	
	'替换request的值,支持ReqNum和ReqStr两个标签
	Function ReplaceRequest(Content)
		     Dim regEx, Matches, Match,TempStr,QStr,ReqType
			 Set regEx = New RegExp
			 regEx.Pattern= "{(ReqNum|ReqStr)[^{}]*}"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 For Each Match In Matches
				On Error Resume Next
				TempStr = Match.Value
				ReqType=Split(TempStr,"(")(0)
				QStr=Replace(Split(TempStr,"(")(1),")}","")
				If ReqType="{ReqNum" Then
				 Content=Replace(Content,TempStr,ChkClng(S(QStr)))
				Else
				 Content=Replace(Content,TempStr,S(QStr))
				End If
			Next
			ReplaceRequest=Content
	End Function
	
	'分页SQL语句生成代码
	Function GetPageSQL(tblName,fldName,PageSize,PageIndex,OrderType,strWhere,fieldIds)
			Dim strTemp,strSQL,strOrder
			
			'根据排序方式生成相关代码
			if OrderType=0 then
				strTemp=">(select max([" & fldName & "])"
				strOrder=" order by [" & fldName & "] asc"
			else
				strTemp="<(select min([" & fldName & "])"
				strOrder=" order by [" & fldName & "] desc"
			end if
			
			'若是第1页则无须复杂的语句
			if PageIndex=1 then
				strTemp=""
				if strWhere<>"" then strTemp = " where " + strWhere
				strSQL = "select top " & PageSize & " " & fieldIds & " from [" & tblName & "]" & strTemp & strOrder
			else '若不是第1页，构造SQL语句
				strSQL="select top " & PageSize & " " & fieldIds & " from [" & tblName & "] where [" & fldName & "]" & strTemp & _
				" from (select top " & (PageIndex-1)*PageSize & " [" & fldName & "] from [" & tblName & "]" 
				if strWhere<>"" then
				strSQL=strSQL & " where " & strWhere
				end if
				strSQL=strSQL & strOrder & ") as tblTemp)"
				if strWhere<>"" then
				strSQL=strSQL & " And " & strWhere
				end if
				strSQL=strSQL & strOrder
			end if
			GetPageSQL=strSQL 
	End Function
	
  '======================================================================================
End Class
%> 
<!--#include file="../../Plus/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Const TemplatePath="3g"
Class MyCls
        Private KS, KSR,F_C,ModelID,ItemId,ID,ClassID,RS,DocXML,UserLoginTF,KSUser,XML,Node,MaxPerPage,Key,TopNum,TopStr
		Private InfoPurview,ReadPoint,ChargeType,PitchTime,ReadTimes,ClassPurview,UserName,ModelChargeType,ChargeStr,ChargeStrUnit,ChargeTableName,DateField,IncomeOrPayOut,CurrPoint,Content,UrlsTF, CurrPage,SqlStr,ModelTable,FieldStr,Param,TotalPut,PayTF,domainstr
		Private Sub Class_Initialize()
		  MaxPerPage=10  '每页显示条数
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		  If KS.WSetting(0)<>"1" Then KS.die "<div style=""text-align:center;margin:20px"">对不起，本站没有开通3G手机访问频道！</div>"
		  Fcls.CallFrom3g="true"
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#Include file="function.asp"-->
		<%
		Public Sub Kesion()
		   F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/index.html")
		   InitialCommon
		   F_C = KSR.KSLabelReplaceAll(F_C) 
		   KS.Die F_C
		End Sub
		
		
		'频道首页
		Sub Channel()
		 ModelID=KS.ChkClng(Request("ID"))
		 If ModelID=0 Then Call Kesion() : Exit Sub
		 FCls.ChannelID=ModelID
		 F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/index.html")
		 InitialCommon
		 F_C = Replace(F_C,"{$GetModelID}",ModelID)
		 F_C = KSR.KSLabelReplaceAll(F_C) 
		 KS.Die F_C
		End Sub
		
		
		
		
		'静态化列表
		Sub List()
		 ID=KS.ChkClng(Request("ID"))
		 CurrPage=KS.ChkClng(KS.S("page"))
		 If Currpage<1 Then Currpage=1
		' If ID=0 THEN Call List1():Exit Sub
		 
		 UserLoginTF=Cbool(KSUser.UserLoginChecked)
		 Dim RSObj,ChannelID,TemplateID
		 If ID<=0 Then
		    ModelID=KS.ChkClng(Request("ModelID"))
			Call FCls.SetClassInfo(ModelID,0,0)
			TopNum=500  '没有传栏目ID，限制只查询500条记录
			If TopNum<>0 Then TopStr=" Top " & TopNum
		 Else
			 Set RSObj=Conn.Execute("Select top 1 ID,ClassPurview,TN,FolderTemplateID,FolderDomain,DefaultArrGroupID,ChannelID,WapFolderTemplateID From KS_Class Where ClassID=" & ID)
			 IF Not RSObj.Eof Then 
			  If RSObj("ClassPurview")=2 and  RSObj("channelid")<>8 Then
				If Cbool(KSUser.UserLoginChecked)=false Then 
				 Call KS.Alert("本栏目为认证栏目，至少要求本站的注册会员才能浏览!",KS.GetDomain & KS.Wsetting(4) &"/login.asp"):Response.End
				elseIF KS.FoundInArr(RSObj("DefaultArrGroupID"),KSUser.GroupID,",")=false Then
				 Call KS.Alert("对不起，你所在的用户级没有权限浏览!",Request.ServerVariables("http_referer")):Response.End
				End If
			   End If
				 ModelID=RSObj("ChannelID")
				 Dim BigID:BigID=RSObj("ID")
				 Dim ClassID:ClassID=RSObj("ID")
				 if KS.ChkClng(KS.C_C(ClassID,14))=2 then
					   response.Redirect(KS.C_C(ClassID,2))
				 end if
				 TemplateID=RSObj("WapFolderTemplateID")
	
				 Call FCls.SetClassInfo(ModelID,RSObj("ID"),RSObj("TN"))
			 End If
			 RSObj.Close:Set RSObj=Nothing   
		End If	 
			 
		   if KS.ChkClng(request("tid"))<>0 then ModelID=9
		   If ModelID=9 Then 
				 Call SJList()
				 Exit Sub
		   End If
			If KS.ChkClng(KS.C_S(ModelID,21))=0 Then   Call KS.ShowTips("error","对不起，本频道已关闭!"):Response.End
		   If TemplateID="" Then TemplateID=KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/list.html"
		   
		   F_C = KSR.LoadTemplate(TemplateID)
           InitialCommon

		    F_C=Replace(F_C,"{$GetModelID}",ModelID)
			 
			 F_C = KSR.KSLabelReplaceAll(F_C)
			Dim LabelParamStr:LabelParamStr=Application(KS.SiteSN&"PageParam")

			If Not KS.IsNul(LabelParamStr) And Instr(F_C,"{KS:PageList}")=0 Then
				 Dim XMLDoc,XMLSql,LabelStyle,KMRFOBJ
				 Dim ParamNode,IncludeSubClass,ModelID,OrderStr,PrintType,PageStyle,PicStyle,ShowPicFlag,FieldStr,Param
				 Dim PerPageNumber,TotalPut,PageNum,TempStr,TableName
				 Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 If XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
					 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
					 
					 if modelid=0 Then
					 ModelID         = ParamNode.getAttribute("modelid") : If Not IsNumeric(ModelID) Then ModelID=1
					 End If
					 
					 IncludeSubClass = ParamNode.getAttribute("includesubclass"):If KS.IsNul(IncludeSubClass) Then IncludeSubClass=true 
					 PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					 PageStyle       = ParamNode.getAttribute("pagestyle") : If PageStyle="" Or IsNull(PageStyle) Then PageStyle=1
					 PicStyle        = ParamNode.getAttribute("picstyle")
					 OrderStr        = ParamNode.getAttribute("orderstr") : If OrderStr="" Or IsNull(OrderStr) Then OrderStr="ID Desc"
					 ShowPicFlag     = ParamNode.getAttribute("showpicflag") : If ShowPicFlag="" Or IsNull(ShowPicFlag) Then ShowPicFlag=false
					 PerPageNumber   = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNumber) Then PerPageNumber=10
						FCls.PerPageNum=PerPageNumber
					 Param = " Where I.Verific=1 And I.DelTF=0"
					 If CBool(IncludeSubClass) = True Then 
					 Param= Param & " And I.Tid In (" & KS.GetFolderTid(BigID) & ")" 
					 Else 
					 Param= Param & " And I.Tid='" & BigID & "'"
					 End If
					 
					 Dim SQLParam:SQLParam=ParamNode.getAttribute("sqlparam")
						If Not KS.IsNul(SQLParam) Then
						 Param=Param & " " & SQLParam
					  End If
					  
					   if request("tj")="1" then param=param &"  and recommend=1"
					   if request("rm")="1" then param=param &"  and Popular=1"
					   
					   if ks.chkclng(request("typeid"))<>0 and modelid=8 then  Param=Param & " and typeid=" & ks.chkclng(request("typeid"))
					   
					 
					 
					 Set KMRFObj= New RefreshFunction
					 Set KMRFObj.ParamNode=ParamNode
				     Call KMRFObj.LoadField(ModelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param)
				
					If Lcase(Left(Trim(OrderStr),2))<>"id" Then  OrderStr=OrderStr & ",I.ID Desc"	
					SqlStr = "SELECT " & FieldStr & " FROM " & KS.C_S(ModelID,2) & " I " & Param & " ORDER BY I.IsTop Desc," & OrderStr
					Set RS=Server.CreateObject("ADODB.RECORDSET")
					'KS.DIE SQLSTR
					RS.Open SqlStr, Conn, 1, 1
					If RS.EOF And RS.BOF Then
						TempStr = "<p>此栏目下没有" & KS.C_S(ModelID,3) & "</p>"
					Else
						PerPageNumber=cint(PerPageNumber)
						TotalPut = Conn.Execute("select Count(id) from " & KS.C_S(ModelID,2) & " I " & Param)(0)
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
						If CurrPage >1 and (CurrPage - 1) * PerPageNumber < totalPut Then
							RS.Move (CurrPage - 1) * PerPageNumber
						Else
							CurrPage = 1
						End If
						Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNumber),RS,"row","root")
						Call KMRFObj.LoadPageParam(XMLSQL,ParamNode,ModelID)
						LabelStyle=Application("LabelStyle")
						TempStr = KMRFObj.ExplainGerericListLabelBody(LabelStyle)
						XMLSql=Empty
						FCls.TotalPut=TotalPut
						FCls.PageStyle=PageStyle       '分页样式
						FCls.TotalPage=PageNum         '总页数
					End If
						if Instr(tempstr,"[KS:PageStyle]")=0 Then
							tempstr=tempstr & "[KS:PageStyle]"
						End If
                        F_C=Replace(F_C,"{Tag:Page}",tempstr)

					RS.Close:Set RS=Nothing					
					XMLDoc= Empty : Set ParamNode=Nothing
				End If	
				
			End If
			if Instr(F_C,"[KS:PageStyle]")<>0 Then
		  	F_C=Replace(F_C,"[KS:PageStyle]",KS.ReplacePage(FCls.PageStyle,ModelID,id,CurrPage,FCls.TotalPut,FCls.PerPageNum))
			End If
			F_C=Replace(F_C,"[#CurrPage]",CurrPage)

		 
		 Set KMRFObj=Nothing
		 KS.Echo F_C
		End Sub
		
		
		
		'列表页
		Sub SJList()
		 ID=KS.ChkClng(Request("ID"))
		 CurrPage=KS.ChkClng(KS.S("page"))
		 If Currpage<1 Then Currpage=1
		 If ID<>0 Then
			 UserLoginTF=Cbool(KSUser.UserLoginChecked)
			 Dim RSObj
				 Set RSObj=Conn.Execute("Select top 1 ID,ClassPurview,TN,FolderTemplateID,WapFolderTemplateID,FolderDomain,DefaultArrGroupID,ChannelID From KS_Class Where ClassID=" & ID)
			 IF RSObj.Eof And RSObj.Bof Then  RSObj.Close:Set RSObj=Nothing:Call KS.Alert("非法参数!",""):Exit Sub
				 ModelID=RSObj("ChannelID")
				 ClassID=RSObj("ID")
				 if KS.ChkClng(KS.C_C(ClassID,14))=2 then
				   response.Redirect(KS.C_C(ClassID,2))
				 end if
				 Dim Templateid:TemplateID=RSObj("WapFolderTemplateID")
				 Call FCls.SetClassInfo(ModelID,ClassID,RSObj("TN"))
				 RSObj.Close:Set RSObj=Nothing
		   Else
		        ModelID=KS.ChkClng(Request("ModelID"))
				Call FCls.SetClassInfo(ModelID,0,0)
				TopNum=500  '没有传栏目ID，限制只查询500条记录
				If TopNum<>0 Then TopStr=" Top " & TopNum
		   End If
		   
		  if KS.ChkClng(request("tid"))<>0 then ModelID=9
		   If ModelID=9 Then 
		   TemplateID=KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/list.html"
		   End If
		   
		  If TemplateID="" Then TemplateID=KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/list.html"
		   
		  
			 F_C = KSR.LoadTemplate(TemplateID)
			 InitialCommon
			 F_C=Replace(F_C,"{$GetModelID}",ModelID)
			 F_C = KSR.KSLabelReplaceAll(F_C)
		         

		     ModelTable=KS.C_S(ModelID,2)
		    FieldStr=FieldStr & GetDiyFieldStr(ModelID)
		     Param=" Where dtfs<>0 and verific=1"
			 if request("tj")="1" then param=param &"  and recommend=1"
			  if ks.chkclng(request("tid"))<>0 then param=param &"  and tid in(select id from ks_sjclass where ts like '%"&ks.chkclng(request("tid")) & "%')"
			 
			 if request("rm")="1" then param=param & " and popular=1"
		     SqlStr="Select " & TopStr & " * From " & ModelTable & Param & " Order by ID Desc"
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  'ks.die sqlstr
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     TotalPut = Conn.Execute("select Count(1) from " & ModelTable & " " & Param)(0)
			 If TotalPut>TopNum And TopNum<>0 Then TotalPut=TopNum
			 If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrPage - 1) * MaxPerPage
			 End If
			 Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
		  End If
		 RS.Close : Set RS=Nothing
         Scan F_C
		End Sub
		
		
		Sub Echo(sStr)
				Response.Write    sStr
		End Sub 
		
		Sub Scan(sTemplate)
			Dim iPosLast, iPosCur
			iPosLast    = 1
			While True 
				iPosCur    = InStr(iPosLast, sTemplate, "{@")
				If iPosCur>0 Then
					Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
					iPosLast    = Parse(sTemplate, iPosCur+2)
				Else 
					Echo    Mid(sTemplate, iPosLast)
					Exit Sub  
				End If 
		   Wend 
		End Sub 
		
		Function Parse(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sValue, sTemp
			iPosCur        = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			iPosCur       = InStr(sTemp, ".")
			if iPosCur>1 Then
			sToken        = Left(sTemp, iPosCur-1)
			End If
			sValue        = Mid(sTemp, iPosCur+1) 
		
			Select Case sValue
				Case "begin"
					sTemp            = "{@" & ( sToken & ".end}" )
					iPosCur            = InStr(iPosBegin, sTemplate, sTemp)
					ParseArea      sToken, Mid(sTemplate, iPosBegin, iPosCur-iPosBegin)
					iPosBegin        = iPosCur+Len(sTemp)
				Case Else
					ParseNode sToken, sValue 
			End Select 
			Parse    = iPosBegin
		End Function 

		
		Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "loop"
				      If IsObject(XML) Then
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,找不到任何相关记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName
				case "list" 
				          select case sTokenName
						    case "showpage" echo ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							case "leavetime" echo FormatNumber((timer-starttime),5)
							case "keyword" echo KS.R(key)
							case "channelid" echo modelid
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "linkurl" 
			   If ModelID=0 Then
			       echo "show.asp?m=" & GetNodeText("channelid") &"&d=" & GetNodeText("infoid")
			   ElseIf ModelID=9 Then
			       echo "exam/index.asp?id=" & GetNodeText("id")
			   Else
				   echo "show.asp?m=" & modelid &"&d=" & GetNodeText("id")
			   End If 
			case "getdate" 
			 if modelid=9 then
			 echo KS.GetTimeFormat(GetNodeText("date"))
			 else
			 echo KS.GetTimeFormat(GetNodeText("adddate"))
			 end if
			case "photourl" dim photourl:photourl=GetNodeText("photourl")
			if ks.isnul(photourl) then photourl="../images/nopic.gif"
			echo photourl
			case "classname" 
			  if modelid<>9 then
			   echo KS.C_C(GetNodeText("tid"),1)
			  else
			    dim rst:set rst=conn.execute("select top 1 tname from ks_sjclass where id=" & KS.ChkClng(GetNodeText("tid")))
				if not rst.eof then
				 echo split(rst("tname"),"|")(ubound(split(rst("tname"),"|"))-1)
				end if
				rst.close
				set rst=nothing
			  end if
			case "classurl" 
			 if modelid<>9 then
			  echo "list.asp?id=" & KS.C_C(GetNodeText("tid"),9)
			 else
			  echo "list.asp?tid=" & GetNodeText("tid")
			 end if
			case "typename" echo KS.GetGQTypeName(GetNodeText("typeid"))
			case "intro" 
			 Dim Intro:intro=KS.Gottopic(KS.LoseHtml(GetNodeText("intro")),160)
			 Intro=Replace(Intro,"&nbsp;","")
			 If Not KS.IsNul(Key) Then
			  echo Replace(Intro,key,"<span style='color:red'>" & key & "</span>")
			 Else
			 echo intro
			 End If
			case else
			  echo GetNodeText(sTokenName)
		  End Select
		End Sub
		Function GetNodeText(NodeName)
		 Dim N,Str
		 NodeName=Lcase(NodeName)
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then Str=N.text
		  If Not KS.IsNul(Key)  And NodeName="title" Then
		   Str=Replace(Str,key,"<span style='color:red'>" & key & "</span>")
		  End If
		  GetNodeText=Str
		 End If
		End Function
		
		
		
		'内容页
		Sub Show()
		 ModelID=KS.ChkClng(Request("m"))
		 ItemID=KS.ChkClng(Request("d"))
		 PayTF=KS.ChkClng(KS.S("Pt"))
		 ID=ItemID
		 If ModelID=0 or ItemID=0 Then Call Kesion() : Exit Sub
		 CurrPage=KS.ChkClng(KS.S("P"))
		 If CurrPage<=0 Then CurrPage=1
		 UserLoginTF=Cbool(KSUser.UserLoginChecked)
		 Select Case (KS.C_S(modelid,6))
		   Case 1 Call StaticArticleContent()
		   Case 2 Call StaticPhotoContent()
		   Case 3,5,7 Call StaticContent()
		   Case 8 Call StaticSupplyContent()
		  End Select
		 
		End Sub
		Sub GetRecords()
		  If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ShowContent"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@ID",3)
				Cmd.Parameters.Append cmd.CreateParameter("@TableName",200,1,220)
				Cmd("@ID")=id
				Cmd("@TableName")=KS.C_S(ModelID,2)
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
		  Else
			    Set RS=Conn.Execute("Select top 1 a.*,ClassPurview,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ModelID,2) & " a inner join KS_Class b on a.tid=b.id Where a.ID=" & ItemID)
		  End If
		End Sub
		
		Function GetPageStr(Page)
		  GetPageStr="?m=" & ModelID & "&d="& ItemID & "&p="&Page
		End Function
		
		'检查收费及权限
		Sub CheckCharge(KSR)

		  With KSR
			 InfoPurview = Cint(.Node.SelectSingleNode("@infopurview").text)
			 ReadPoint   = Cint(.Node.SelectSingleNode("@readpoint").text)
			 ChargeType  = Cint(.Node.SelectSingleNode("@chargetype").text)
			 PitchTime   = Cint(.Node.SelectSingleNode("@pitchtime").text)
			 ReadTimes   = Cint(.Node.SelectSingleNode("@readtimes").text)
			 ClassPurview= Cint(.Node.SelectSingleNode("@classpurview").text)
			If InfoPurview=2 or ReadPoint>0 Then
			  
				   IF UserLoginTF=false Then
					 Call GetNoLoginInfo(Content)
				   Else
				   
						 IF KS.FoundInArr(.Node.SelectSingleNode("@arrgroupid").text,KSUser.GroupID,",")=false and readpoint=0 Then
						   Content="<div style=""text-align:center"">对不起，你所在的用户组没有查看本" & KS.C_S(ModelID,3) & "的权限!</div>"
						 Else 
						   
							  Call PayPointProcess()
						 End If
				   End If
			  ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2) Then 
				  If UserLoginTF=false Then
					Call GetNoLoginInfo(Content)
				  Else     
				  
					 '============继承栏目收费设置时,读取栏目收费配置===========
					 ReadPoint  = Cint(.Node.SelectSingleNode("@defaultreadpoint").text)   
					 ChargeType = Cint(.Node.SelectSingleNode("@defaultchargetype").text)
					 PitchTime  = Cint(.Node.SelectSingleNode("@defaultpitchtime").text)
					 ReadTimes  = Cint(.Node.SelectSingleNode("@defaultreadtimes").text)
					 '============================================================
					 If ClassPurview=2 Then
						 IF KS.FoundInArr(.Node.SelectSingleNode("@defaultarrgroupid").text,KSUser.GroupID,",")=false Then
							Content="<div style=""text-align:center"">对不起，你所在的用户组没有查看的权限!</div>"
						 Else
							Call PayPointProcess()
						 End If
					Else    
					 Call PayPointProcess()
					End If
				  End If
			 Else
			   Call PayPointProcess()
			 End If 
		  End With
		End Sub
		
		Sub StaticArticleContent()
		  Call GetRecords()
		 IF RS.Eof And RS.Bof Then
		  RS.Close:Set RS=Nothing
		  KS.ShowTips "error","您要查看的" & KS.C_S(ModelID,3) & "已删除。或是您非法传递注入参数!"
		 ElseIF Cint(RS("Changes"))=1 Then 
		   Dim ClassID:ClassID=RS("Tid")
		   Dim Fname:Fname=RS("articlecontent")
		   RS.Close:Set RS=Nothing
		   Response.Redirect Fname
		 End IF
		  Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		  With KSR 
			 Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
		      .oTid=.Node.SelectSingleNode("@otid").text
			 If .Node.SelectSingleNode("@verific").text<>1 And UserLoginTF=False And KSUser.UserName<>.Node.SelectSingleNode("@inputer").text Then
			   KS.ShowTips "error","对不起，该" & KS.C_S(ModelID,3) & "还没有通过审核!"
			 End If
			 Call FCls.SetContentInfo(ModelID,.Tid,.oTid,ItemID,.Node.SelectSingleNode("@title").text)
		 
		     Call CheckCharge(KSR) '检查权限
		 'F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/show.html")
		 F_C = KSR.LoadTemplate(.Node.SelectSingleNode("@waptemplateid").text)
		 InitialCommon
		 If InStr(F_C,"[KS_Charge]")=0 Then
		   F_C = Replace(F_C,"{$GetArticleContent}","[KS_Charge]{$GetArticleContent}[/KS_Charge]")
		 End If
		 Dim ContentArr

		 If .Node.SelectSingleNode("@postid").text<>"0" Then
		  ContentArr=Split(.UbbCode(.Node.SelectSingleNode("@articlecontent").text,1),"[NextPage]")
		 Else
		  ContentArr=Split(.Node.SelectSingleNode("@articlecontent").text,"[NextPage]")
		 End If
		 Dim TotalPage,N,K,PageStr,NextUrl,PrevUrl
			TotalPage = Cint(UBound(ContentArr) + 1)
			   If TotalPage > 1 Then
					   If CurrPage = 1 Then
					     PrevUrl="" : NextUrl=GetPageStr(CurrPage + 1)
					   ElseIf CurrPage = TotalPage Then
					     NextUrl = KS.GetFolderPath(.Tid) : PrevUrl = GetPageStr(CurrPage - 1)
					   Else
					     NextUrl = GetPageStr(CurrPage + 1) :PrevUrl = GetPageStr(CurrPage - 1)
					   End If
					   PageStr =  "<div id=""pageNext"" style=""text-align:center""><table align=""center""><tr><td>"
					   If CurrPage > 1 And PrevUrl<>"" Then PageStr = PageStr & "<a class=""prev"" href=""" & PrevUrl & """>上一页</a> "
					 Dim StartPage:StartPage=1
					 if (CurrPage>=10) then StartPage=(CurrPage\10-1)*10+CurrPage mod 10+2
				     For N = StartPage To TotalPage
						 If CurrPage = N Then
						  PageStr = PageStr & ("<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</span></a> ")
						 Else
						  PageStr = PageStr & ("<a class=""num"" href=""" & GetPageStr(N) & """>" & N & "</a> ")
						 End If
						 K=K+1
						 If K>=10 Then Exit For
					 Next
					 PageStr = "<div id=""MyContent"">" & ContentArr(CurrPage-1) & "</div>" & PageStr 
					 If CurrPage<>TotalPage Then PageStr = PageStr & " <a class=""next"" href=""" & NextUrl & """>下一页</a>"
					 PageStr = PageStr & "</td></tr></table></div>"

					 Dim PageTitleArr,PageTitle
					 PageTitle=	.Node.SelectSingleNode("@pagetitle").text
					 
					 If Not KS.IsNul(PageTitle) Then
					  PageTitleArr=Split(PageTitle,"§")
					  If CurrPage-1<=Ubound(PageTitleArr) Then F_C=Replace(F_C,"{$GetArticleTitle}",PageTitleArr(CurrPage-1))
					 ElseIF Currpage>0 Then
					   F_C=Replace(F_C,"{$GetArticleTitle}",.Node.SelectSingleNode("@title").text & "(" & currpage & ")")
					 End IF
				 Else
				  NextUrl=KS.GetFolderPath(.Tid)
				  PageStr = "<div id=""MyContent"">" & ContentArr(0) & "</div>"
				 End If

'KS.DIE F_C
				 .ModelID = ModelID
				 .ItemID  = ItemID
				 .PageContent=PageStr
				 .NextUrl=NextUrl
				 .TotalPage=TotalPage
				 .Templates=""
				 .Scan F_C
		 		 F_C = .Templates
				 				 					 
		  If Content<>"True" Then
		   Dim ChargeContent:ChargeContent=KS.CutFixContent(F_C, "[KS_Charge]", "[/KS_Charge]", 0)
		   F_C=Replace(F_C,"[KS_Charge]" & ChargeContent &"[/KS_Charge]",Content)
		  Else
		   F_C=Replace(Replace(F_C,"[KS_Charge]",""),"[/KS_Charge]","")
		  End If
		  If Instr(F_C,"[KS_ShowIntro]")<>0 Then
			  If CurrPage=1 Then
		        F_C=Replace(Replace(F_C,"[KS_ShowIntro]",""),"[/KS_ShowIntro]","")
			  Else
		        F_C=Replace(F_C,KS.CutFixContent(F_C, "[KS_ShowIntro]", "[/KS_ShowIntro]", 1),"")
			  End If
		  End If

		  F_C = .KSLabelReplaceAll(F_C)
		End With
          F_C=Replace(Replace(Replace(Replace(F_C,"{§","{$"),"{#LB","{LB"),"{#SQL","{SQL"),"{#=","{=")

		  KS.Echo F_C

	   End Sub
	   
	   
	   Sub StaticPhotoContent()
	      Call GetRecords()
		 IF RS.Eof And RS.Bof Then
		  RS.Close : Set RS=Nothing
		  KS.ShowTips "error","对不起,您要查看的" & KS.C_S(ModelID,3) & "已删除。或是您非法传递注入参数!"
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		    .Tid=.Node.SelectSingleNode("@tid").text
            .oTid=.Node.SelectSingleNode("@otid").text
		 If .Node.SelectSingleNode("@verific").text<>1 And UserLoginTF=False And KSUser.UserName<>.Node.SelectSingleNode("@inputer").text Then
		   KS.ShowTips "error","对不起，该" & KS.C_S(ModelID,3) & "还没有通过审核!"
		   Response.End
		 End If
		 Call FCls.SetContentInfo(ModelID,.Tid,.Otid,ID,.Node.SelectSingleNode("@title").text)
         Dim ShowStyle,PageNum
		 PageNum     = KS.ChkClng(.Node.SelectSingleNode("@pagenum").text) : If PageNum=0 Then PageNum=10
		 ShowStyle   = KS.ChkClng(.Node.SelectSingleNode("@showstyle").text) : If ShowStyle=0 Then ShowStyle=1
		   
         Call CheckCharge(KSR)  
		 
		 	Dim KSLabel:Set KSLabel =New RefreshFunction
			F_C = KSR.LoadTemplate(.Node.SelectSingleNode("@waptemplateid").text)
			'F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/show.html")
			InitialCommon
			 Dim PicUrlsArr,N,PageStr,TotalPage,NextUrl,Tp,Tpage,r,thumblist
			 If Cbool(UrlsTF)=true Then
			            PicUrlsArr = Split(Content, "|||")
						
			           Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style3")
                        n=0
						dim bigpic:BigPic=Split(PicUrlsArr(n), "|")(1) : If lcase(Left(BigPic,4))<>"http" Then BigPic=KS.Setting(2) & BigPic
						ThumbList="<div><img title=""" &Split(PicUrlsArr(n), "|")(0) & """ href=""" & BigPic &""" class=""scrollLoading swipebox"" alt='" & Split(PicUrlsArr(n), "|")(0) & "' style=""cursor:pointer;background:url(" & KS.GetDomain &"images/default/loading.gif) no-repeat center;""  data-url=""" & BigPic &""" src=""" & KS.GetDomain &"images/default/pixel.gif"" border='0'><div class=""imgtitle"">共有 <span>" & (UBound(PicUrlsArr)+1) &"</span> 张图片，点击上图浏览。</div></div>"

					   For n=1 to UBound(PicUrlsArr)
						 BigPic=Split(PicUrlsArr(n), "|")(1) : If lcase(Left(BigPic,4))<>"http" Then BigPic=KS.Setting(2) & BigPic
						 ThumbList=ThumbList & "<div style=""display:none""><img title=""" &Split(PicUrlsArr(n), "|")(0) & """ href=""" & BigPic &""" class=""scrollLoading swipebox"" alt='" & Split(PicUrlsArr(n), "|")(0) & "' style=""cursor:pointer;background:url(" & KS.GetDomain &"images/default/loading.gif) no-repeat center;""  data-url=""" & BigPic &""" src=""" & KS.GetDomain &"images/default/pixel.gif"" border='0'><div class=""imgtitle"">" & Split(PicUrlsArr(n), "|")(0) & "</div></div>"
					   Next
					   Tp=Replace(Tp,"{$ShowPage}","")
					   Tp=Replace(Tp,"{$ShowImgList}",ThumbList)
			 
			 
			 
			 
			 
			 
			 
			 
					
				F_C=Replace(F_C,"{$ShowPictures}",Tp)
                If Tpage>1 Then F_C=Replace(F_C,"{$GetPictureName}",.Node.SelectSingleNode("@title").text & "(" & currpage & ")")
			Else
			    PageStr = Content
			End If
			     
				 .ModelID = ModelID
				 .ItemID  = ItemID
				 .PageContent=PageStr
				 .NextUrl=NextUrl
				 .TotalPage=TotalPage
				 .Templates=""
				 .Scan F_C
		 		  F_C = .Templates
				
			F_C = KSR.KSLabelReplaceAll(F_C)
		  End With
		  KS.Echo F_C
		  Set KSLabel=Nothing
	   End Sub
	   
	   Sub StaticContent()
	       SQLStr="Select top 1 * From " & KS.C_S(ModelID,2) & "  Where verific=1 And ID=" & ItemID
	       If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_TSql"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@SQL",202,1,8000,SQLStr)
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
			Else
			    Set RS=Conn.Execute(SqlStr)
			End If
		 IF RS.Eof And RS.Bof Then
		    RS.Close:Set RS=Nothing
		    KS.ShowTips "error","您要查看的" & KS.C_S(ModelID,3) & "不存在!"
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
		    
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  .oTid=.Node.SelectSingleNode("@otid").text
			  Call FCls.SetContentInfo(ModelID,.Tid,.oTid,ID,.Node.SelectSingleNode("@title").text)
			  If KS.C_S(ModelID,6)=3 or KS.C_S(ModelID,6)=5 Then
			   F_C = KSR.LoadTemplate(.Node.SelectSingleNode("@waptemplateid").text)
			  Else  '影视
			   Session("movie_id")=ItemID
			  ' F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/show.html")
			   F_C = KSR.LoadTemplate(.Node.SelectSingleNode("@waptemplateid").text)
			   F_C = Replace(F_C,"{$PrePlayTime}",KS.ChkClng(.Node.SelectSingleNode("@preplaytime").text))
			  End If
		      InitialCommon
			  .ModelID = ModelID
			  .ItemID  = ItemID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan F_C
			  F_C = .Templates 
			  F_C = .KSLabelReplaceAll(F_C)
		 End With
		 KS.Echo F_C
	   End Sub
	   
       Sub StaticSupplyContent()
	    If Not KS.IsNul(KS.C("AdminName")) Then
		 SQLStr="Select top 1 b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.ID=" & ItemID
		 Else
		 SQLStr="Select top 1 b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.verific=1 and a.ID=" & ItemID
		 End If
		 
		 If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_TSql"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@SQL",202,1,8000,SQLStr)
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
		Else
			    Set RS=Conn.Execute(SqlStr)
		End If
		 
		 IF RS.Eof And RS.Bof Then
		  RS.Close :Set RS=Nothing
		  KS.ShowTips "error","您要查看的信息已删除或未通过审核!"
		 End IF
		 With KSR 
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			   .oTid=.Node.SelectSingleNode("@otid").text
		  F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/" & KS.C_S(ModelID,10) &"/show.html")
		  Call FCls.SetContentInfo(8,.Tid,.oTid,ItemID,.Node.SelectSingleNode("@title").text)
		  InitialCommon
			  
			  .ModelID = 8
			  .ItemID  = ItemID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan F_C
			  F_C = .Templates 
			  
			  If Instr(F_C,"[KS_Charge]")<>0 Then
				 Dim ChargeContent:ChargeContent=KS.CutFixContent(F_C, "[KS_Charge]", "[/KS_Charge]", 1)
				 F_C=Replace(F_C,ChargeContent,LFCls.GetConfigFromXML("supply","/labeltemplate/label","divajax"))
				End If
				
			  F_C = .KSLabelReplaceAll(F_C)
		 End With
		 KS.Echo F_C
	   
	   End Sub
	   
	   
		
					  
End Class
%>

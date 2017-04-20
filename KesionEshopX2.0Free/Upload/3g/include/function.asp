<%
'3g版通用标签替换
Public Sub InitialCommon()
    Dim RCls:Set RCls=new Refresh
    Call RCls.Replace3GCommonLabel(F_C)
	Set RCls=nothing
	If InStr(F_C,"{$GetModelList}")<>0 Then GetModelList
	If Instr(F_C,"{$GetUserLogin}")<>0 Then GetUserLogin
	If Instr(F_C,"{$GetSJCategory}")<>0 Then GetSJCategory
End Sub


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
				 ParamStr=KS.QueryParam("page") : If ParamStr<>"" Then ParamStr="&" & ParamStr	
				 n=0:startpage=1:CurrPage=KS.ChkClng(CurrPage)
				 pageStr=pageStr & "<form action=""" & FileName & "?1=1" & ParamStr & """ name=""pageform"" method=""post""><div id='fenye' class='fenye'><table border=""0"" align=""center"" cellspacing=""0"" cellpadding=""0""><tr><td nowrap>" & vbcrlf
				 pageStr=pageStr & "<a href=""" & FileName & "?page=1" & ParamStr & """ class=""prev"">首 页</a>"
				 if (CurrPage>1) then 
				  pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage-1 & ParamStr & """ class=""prev"">上一页</a>"
				 Else
				  pageStr=PageStr & "<a href=""#"" onclick=""return false;"" class=""prev"">上一页</a>"
				 End If
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-9
				 If startpage<0 Then startpage=1
				 'For J=startpage To TotalPage
				 '   If J= CurrPage Then
				 '    PageStr=PageStr & " <a href=""#"" class=""curr"">" & J &"</a>"
				 '   Else
				 '    PageStr=PageStr & " <a class=""num"" href=""" & FileName & "?page=" &J& ParamStr & """>" & J &"</a>"
				'	End If
				'	n=n+1
				'	if n>=10 then exit for
				' Next
				If TotalPage<=0 Then TotalPage=1
				pageStr=PageStr & "</td><td nowrap style=""text-align:center"">第" & CurrPage & "页 共" & TotalPage &"页</td><td nowrap>"

                 if CurrPage<>TotalPage and totalnumber>MaxPerPage then 
				  pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage+1 & ParamStr & """ class=""next"">下一页</a>"
				 Else
				  pageStr=PageStr & "<a href=""#"" onclick=""return false"" class=""next"">下一页</a>"
				 End IF

				 if CurrPage<>TotalPage Then 
				  pageStr=pageStr & "<a href=""" & FileName & "?page=" & TotalPage & ParamStr & """ class=""next"">末 页</a>"
				 Else
				  pageStr=pageStr & "<a href=""#"" onclick=""return false"" class=""next"">末 页</a>"
				 End If	
				
				 pageStr=PageStr & " </td><td>"
				 If ShowTurn=true Then
				 If CurrPage=TotalPage Then CurrPage=0
				 pageStr=PageStr & " 转到:<input class='textbox' type='text' value='" & (CurrPage + 1) &"' name='page' style='width:30px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>"
				 End If
				 PageStr=PageStr & "</td></tr></table></div></form>"
				If PrintOut=true Then echo PageStr Else ShowPage=PageStr
End Function

'用户登录
Sub GetUserLogin()
  Dim Str
  If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then
   Str="<a href=""" & KS.GetDomain & "3g/login.asp"" class=""login"">登录</a>" &vbcrlf
   str=str &" | <a href=""" & KS.GetDomain & "3g/reg.asp"" class=""reg"">注册</a>"&vbcrlf
  Else
   Str="您好,<a href=""" & KS.GetDomain & "3g/user.asp"" style=""color:brown"">" & KS.C("UserName") & "</a> <a onclick=""return(confirm('确定退出登录吗？'));"" href=""" & KS.GetDomain & "3g/Logout.asp"">退出</a>"
  End If
  F_C=Replace(F_C,"{$GetUserLogin}",STR)
End Sub

'频道列表
Sub GetModelList()
	If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
	Dim ModelXML,Node,Str
	Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
	For Each Node In ModelXML.documentElement.SelectNodes("channel")
			if Node.SelectSingleNode("@ks21").text="1" And Node.SelectSingleNode("@ks53").text="1" AND (Node.SelectSingleNode("@ks6").text<4 OR Node.SelectSingleNode("@ks6").text="5"  OR Node.SelectSingleNode("@ks6").text="8" OR Node.SelectSingleNode("@ks6").text="9" OR Node.SelectSingleNode("@ks6").text="7") Then
					  
			Str=Str & "<li><a href=""channel.asp?id=" &Node.SelectSingleNode("@ks0").text &"""><img onerror=""this.src='" & KS.GetDomain &KS.WSetting(4) & "/images/ico/1.gif';"" src=""" & KS.GetDomain &KS.WSetting(4) & "/images/ico/" &Node.SelectSingleNode("@ks0").text &".gif""></a><a href=""channel.asp?id=" &Node.SelectSingleNode("@ks0").text &""">" & Node.SelectSingleNode("@ks1").text &"</a></li>"
			End If
	next 
	F_C=Replace(F_C,"{$GetModelList}",STR)
End Sub 

'试卷分类
Sub GetSJCategory()
   dim str,rs,param
   if ks.chkclng(request("tid"))=0 then
     param=" where tj=1"
   else
     param=" where tn=" & ks.chkclng(request("tid"))
   end if
   set rs=conn.execute("select id,tname from ks_sjclass " & param & " order by orderid,id")
   if rs.eof and rs.bof then
     set rs=conn.execute("select id,tname from ks_sjclass where tn=(select tn from ks_sjclass where id=" & ks.chkclng(request("tid")) &") order by orderid,id")
   end if
   do while not rs.eof
     str=str & "<li><a href=""list.asp?tid=" & rs("id") & """>" & split(rs("tname"),"|")(ubound(split(rs("tname"),"|"))-1) & "</a></li>"
    rs.movenext
   loop
   rs.close
   set rs=nothing
   
   F_C=Replace(F_C,"{$GetSJCategory}",str)
End Sub

'返回自定义字段
Public Function GetDiyFieldStr(ByVal ChannelID)
		  If ChannelID=0 Then Exit Function
		    Dim FieldXML,FieldNode,KSCls,TStr,N
			set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
			If Not IsObject(FieldXML) Then Exit Function
			if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
						  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
						  If DiyNode.Length>0 Then
								For Each N In DiyNode
									Tstr=Tstr & "," & N.SelectSingleNode("@fieldname").text
			                        If N.SelectSingleNode("showunit").text="1" then Tstr=Tstr &"," & N.SelectSingleNode("@fieldname").text&"_unit"
								Next
								Set N=nothing
						  End If
			End If
		  GetDiyFieldStr=Tstr
End Function

'=============================================查看文档扣费处理开始=========================================================
'收费扣点处理过程
Sub PayPointProcess()
	      ModelChargeType=KS.ChkClng(KS.C_S(ModelID,34))
	       Select Case ModelChargeType
			case 1 ChargeStr="资金" : ChargeStrUnit="元人民币": ChargeTableName="KS_LogMoney" : DateField="PayTime": IncomeOrPayOut="IncomeOrPayOut" : CurrPoint=KSUser.GetUserInfo("Money")
			case 2 ChargeStr="积分" : ChargeStrUnit="分积分": ChargeTableName="KS_LogScore": DateField="AddDate":IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetScore
			case else   '按点券
			 ChargeStr=KS.Setting(45) : ChargeStrUnit=KS.Setting(46)&KS.Setting(45) : ChargeTableName="KS_LogPoint" : DateField="AddDate" :IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetUserInfo("Point")
			End Select
	   
	       Dim UserChargeType:UserChargeType=KSUser.ChargeType
	        If (Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2))) and KSUser.UserName<>UserName Then
					 
					     If UserChargeType=1 Then
							 Select Case ChargeType
							  Case 0:Call CheckPayTF("1=1")
							  Case 1:Call CheckPayTF("datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<" & PitchTime)
							  Case 2:Call CheckPayTF("Times<" & ReadTimes)
							  Case 3:Call CheckPayTF("datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
							  Case 4:Call CheckPayTF("datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
							  Case 5:Call PayConfirm()
							  End Select
						Elseif UserChargeType=2 Then
				          If KSUser.GetEdays <=0 Then
						     Content="<div align=center>对不起，你的账户已过期 <font color=red>" & KSUser.GetEdays & "</font> 天,此文需要在有效期内才可以查看，您可以<a href='../user/user_payonline.asp' target='_blank'>点此在线充值</a>或与我们联系充值！</div>"
						  Else
						   Call KSUser.UseLogConsum(KS.C_S(ModelID,6),ModelID,ID,KSR.Node.SelectSingleNode("@title").text)
						   Call GetContent()
						  End If
						Else
						 Call KSUser.UseLogConsum(KS.C_S(ModelID,6),ModelID,ID,KSR.Node.SelectSingleNode("@title").text)
						 Call GetContent()
						end if
					   Else
						  Call GetContent()
					   End IF
End Sub

'检查是否过期，如果过期要重复扣点券
'返回值 过期返回 true,未过期返回false
Sub CheckPayTF(Param)
	   
	    Dim SqlStr:SqlStr="Select top 1 Times From " & ChargeTableName & " Where ChannelID=" & ModelID & " And InfoID=" & ID & " And " & IncomeOrPayOut & "=2 and UserName='" & KSUser.UserName & "' And (" & Param & ")"
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,conn,1,3

		IF RS.Eof And RS.Bof Then
			Call PayConfirm()	
		Else
		       RS.Movelast
			   RS(0)=RS(0)+1
			   RS.Update
			   Call KSUser.UseLogConsum(KS.C_S(ModelID,6),ModelID,ID,KSR.Node.SelectSingleNode("@title").text)
			   Call GetContent()
		End IF
		 RS.Close:Set RS=nothing
 End Sub
	   
Sub PayConfirm()
	     If UserLoginTF=false Then Call GetNoLoginInfo(Content):Exit Sub
		 If ReadPoint<=0 Then Call GetContent():Exit Sub

			 If KS.ChkClng(CurrPoint)<ReadPoint Then
					 Content="<div style=""text-align:center"">对不起，你的可用" & ChargeStr & "不足!阅读本文需要 <span style=""color:red"">" & ReadPoint & "</span> " & ChargeStrUnit &",你还有 <span style=""color:green"">" & CurrPoint & "</span> " & ChargeStrUnit & "</div>,请及时与我们联系！" 
			 Else
					If PayTF="1" Then
					 Dim PayPoint : PayPoint=(ReadPoint*KS.C_C(KSR.Tid,11))/100
					 Dim Descript:Descript="阅读收费" & KS.C_S(ModelID,3) & "“" & KSR.Node.SelectSingleNode("@title").text & "”"
					 Dim TcMsg:TcMsg=KS.C_S(ModelID,3) & "“" & KSR.Node.SelectSingleNode("@title").text & "”的提成"
					 Dim ClientName:ClientName=KSUser.GetUserInfo("realname")
					 If KS.IsNul(ClientName) Then ClientName=KSUser.UserName
					 Select Case ModelChargeType
					   case 1
					     If PayPoint>0 Then Call KS.MoneyInOrOut(KSR.Node.SelectSingleNode("@inputer").text,KSR.Node.SelectSingleNode("@inputer").text,PayPoint,4,1,now,0,"系统",KS.C_S(ModelID,3) & TcMsg,ModelID,ID,1)
					     Call KS.MoneyInOrOut(KSUser.UserName,ClientName,ReadPoint,4,2,now,0,"系统",Descript,ModelID,ID,1)
						 Call GetContent()
					   case 2
					     If KS.ChkClng(PayPoint)>0 Then Call KS.ScoreInOrOut(KSR.Node.SelectSingleNode("@inputer").text,1,KS.ChkClng(PayPoint),"系统",TcMsg,0,0)
						 Session("ScoreHasUse")="+" '设置只累计消费积分
					     Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(ReadPoint),"系统",Descript,ModelID,ID)
						 Call GetContent()
					   case else
					        If PayPoint>0 Then Call KS.PointInOrOut(ModelID,ID,KSR.Node.SelectSingleNode("@inputer").text,1,PayPoint,"系统",TcMsg,0)
							 Call KS.PointInOrOut(ModelID,ID,KSUser.UserName,2,ReadPoint,"系统",Descript,0)
							 Call GetContent()
					 End Select
					 Call KSUser.UseLogConsum(KS.C_S(ModelID,6),ModelID,ID,KSR.Node.SelectSingleNode("@title").text)
					Else
					    Dim PayUrl
						PayUrl=DomainStr & "Item/Show.asp?m=" & ModelID & "&d=" &ID&"&pt=1"
						PayUrl="Show.asp?m=" & ModelID & "&d=" &ID&"&pt=1"

						Content="<div style=""text-align:center"">阅读本文需要消耗 <span style=""color:red"">" & ReadPoint & "</span> " & ChargeStrUnit &",你目前尚有 <span style=""color:green"">" & CurrPoint & "</span> " & ChargeStrUnit &"可用,阅读本文后，您将剩下 <span style=""color:blue"">" & CurrPoint-ReadPoint & "</span> " & ChargeStrUnit &"</div><div style=""text-align:center"">你确实愿意花 <span style=""color:red"">" & ReadPoint & "</span> " & ChargeStrUnit & "来阅读此文吗?</div><div>&nbsp;</div><div align=center><a href=""" & PayUrl & """>我愿意</a>    <a href=""index.asp"">我不愿意</a></div>"
					End If
			 End If
End Sub
Sub GetNoLoginInfo(ByRef content)
	       GCls.ComeUrl=GCls.GetUrl()
		   Content="<div style='text-align:center'>对不起，你还没有登录，本文至少要求本站的注册会员才可查看!</div><div style='text-align:center'>如果你还没有注册，请<a href=""../user/reg/""><span style='color:red'>点此注册</span></a>吧!</div><div style='text-align:center'>如果您已是本站注册会员，赶紧<a href=""javascript:ShowLogin();""><span style='color:red'>点此登录</span></a>吧！</div>"
End Sub

Sub GetContent()
	     Select Case KS.ChkClng((KS.C_S(Modelid,6)))
		  Case 1 Content="True"
		  Case 2 Content=KSR.Node.SelectSingleNode("@picurls").text
		  Case 4 Content="True"
		 End Select
		 UrlsTF=true
 End Sub
'=============================================查看文档扣费处理结束=========================================================
	 
%>
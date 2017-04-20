<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 9.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Link
KSCls.Kesion()
Set KSCls = Nothing

Class Link
        Private KS,ChannelID,ModelTable,Param,XML,Node,StartTime,FormID,TableName,Surveylx,KSUser
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,stype,OrderStr,SurveyName,SurveyContent,TimeLimit,StartDate,ExpiredDate,OnlyUser,AllowGroupID,UserOnce,num
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  MaxPerPage=1000
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		'-------------------------主体------------------------------------
		Public Sub Kesion()

		   Dim Template,KSR
		   FCls.RefreshType = "searchIndex"   
		 Set KSR = New Refresh
		   dim rs,Templ_url
		   FormID=KS.ChkClng(KS.G("id"))
		   if  FormID=0 then Call KS.AlertHistory("ID错误!",-1):response.end
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   RS.Open "Select top 1 ProjectName,ProjectContent,OnlyUser,Template_a,Template_b,TimeLimit,StartDate,ExpiredDate,UserCk,AllowGroupID,UserOnce From KS_Survey Where ID=" & FormID,conn,1,1
		   If RS.EOF And RS.Bof Then
			 Call KS.AlertHistory("ID错误!",-1):response.end
		   else
			 Templ_url=RS("Template_a")	
			 SurveyName= rs("ProjectName")
			 SurveyContent= rs("ProjectContent")
			 TimeLimit=rs("TimeLimit")
			 StartDate=rs("StartDate")
			 ExpiredDate=rs("ExpiredDate")
			 OnlyUser=KS.ChkClng(rs("OnlyUser"))
			 AllowGroupID=rs("AllowGroupID")
			 UserOnce=KS.ChkClng(rs("UserOnce"))
		   End If
		   RS.Close
		   Dim LoginTF:LoginTF=KSUser.UserLoginChecked()
			If OnlyUser=1 and LoginTF=false Then
			    KS.Die("<script>alert('对不起，只会登录会员才能进入!');location.href='index.asp';</script>"):response.end 
			End If
			if OnlyUser=1 then
				If Not KS.IsNul(AllowGroupID) And KS.FoundInArr(AllowGroupID, KSUser.GroupID, ",")=False Then
				  KS.Die("<script>alert('对不起，您所在的会员组不允许投票!');location.href='index.asp';</script>"):response.end 
				End If
			end if
			If TimeLimit="1" Then
				if now<StartDate then KS.Die("<script>alert('对不起，该投票于" & StartDate & "开放！');location.href='index.asp';</script>"):response.end 
		        if now>ExpiredDate then call KS.Die("<script>alert('对不起，该投票已在" & ExpiredDate  & "停止！');location.href='index.asp';</script>"):response.end
		  End If
		   If UserOnce<>0 Then
		   		If Conn.Execute("Select Count(ID) From KS_SurveyResult Where userip='" & KS.GetIP & "' and  SurveyID=" & FormID & "")(0)>=UserOnce  Then
				    KS.Die("<script>alert('对不起，最多只能投" & UserOnce & "次!');location.href='index.asp';</script>"):response.end 
	            End If
		   		If Conn.Execute("Select Count(ID) From KS_SurveyResult Where UserName='" & KS.GetIP & "' and  SurveyID=" & FormID & "")(0)>=UserOnce  Then
				    KS.Die("<script>alert('对不起，最多只能投" & UserOnce & "次!');location.href='index.asp';</script>"):response.end 
	            End If
		  End If
		   Template = KSR.LoadTemplate(Templ_url)'读取模板
		   Template = KSR.KSLabelReplaceAll(Template)
		   Set KSR = Nothing
		   StartTime = Timer()
		   InitialSearch
		   Scan Template
	   End Sub
	 
	    '-------------------------------------------------------------
	   Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "loop"
				      echo "<input name='radioidstr' id='radioidstr'  type='hidden'  value="""" />" &vbcrlf
					   echo "<input name='itemidstr' id='itemidstr'  type='hidden'  value="""" />" &vbcrlf
					   echo "<input name='textastr' id='textastr'  type='hidden'  value="""" />" &vbcrlf
					  If IsObject(XML) Then
					   num=1
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
						num=num+1
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,根据您的查找条件,找不到任何相关记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item"  EchoItem sTokenName
				case "survey" 
				          select case sTokenName
						 	case "TimeLimit" echo TimeLimit
						    case "StartDate" echo StartDate
						    case "ExpiredDate" echo ExpiredDate
						  	case "surveyid" echo FormID
							case "Surveyname" echo SurveyName
						    case "showpage" echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							case "ProjectContent" echo SurveyContent
							case "ProjectDateLimit"
							  If TimeLimit="1" Then
							   echo "(从" & StartDate & "开始，截止到" & ExpiredDate & ")"
							  End If
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "num" echo num
		    case "id" echo GetNodeText("id")
			case "SurveyBox" 
					dim rs
					Select Case GetNodeText("lx")
					case 0
						Surveylx="radio"
					case 1
						Surveylx="checkbox"
					case 2
						Surveylx="textc"
					case else
						Surveylx="类型出错"
					End Select
					if Surveylx<>"类型出错" then
						Set Rs = Server.CreateObject("adodb.recordset")
						Rs.Open "select * from KS_SurveyItem where SurveySTID="&KS.ChkClng(GetNodeText("id")) & " Order By SurveyItemOrder" , Conn, 1, 1	
						dim n,nstr:n=0
						Do While Not rs.Eof
						nstr=chr(65+n)
						n=n+1
							if Surveylx="radio" then
								if rs("SurveyItemType")=0 then 
									echo "<input name='radioItem"&GetNodeText("id") & "' onclick=""textsh(0,"& rs("id") &");"" id='"& rs("id") &"' type='"& Surveylx &"' value='1' />"& nstr & "、" & rs("SurveyItemName") &"<br/>"&vbcrlf
								else
									echo "<input name='radioItem"&GetNodeText("id") & "' onclick=""textsh(1,"& rs("id")&");"" id='"& rs("id") &"' type='"& Surveylx &"' value='1' />"& nstr& "、" & rs("SurveyItemName") 	&vbcrlf	
									echo  " <span style=""display:none"" id=""ot" & rs("id") &"""><textarea name='texta_"& rs("id")  &"' class=""input"" id='textarea_"& rs("id")  &"'></textarea></span><br/>" &vbcrlf 
								end if
							elseif Surveylx="checkbox" then
								if rs("SurveyItemType")=0 then 
									echo "<input name='"& rs("id") &"'  type='"& Surveylx &"' value='0' />"& nstr& "、" & rs("SurveyItemName") &"<br/>"&vbcrlf
								else
									echo "<input name='"& rs("id") &"' onclick=""$('#ot" & rs("id")&"').toggle();"" id='no0'  type='"& Surveylx &"' value='0' />"& nstr& "、" & rs("SurveyItemName") &vbcrlf
									echo   "<span style=""display:none"" id=""ot" & rs("id") &"""><textarea name='texta_"& rs("id")  &"' class=""input"" id='textarea_"& rs("id")  &"' cols='36' rows='6' ></textarea></span><br/>" &vbcrlf
								end if	
							else
								echo   rs("SurveyItemName") &" <textarea class=""input"" name='texta_"& rs("id")  &"' id=""tk_text"" cols='36' rows='6' ></textarea><br/>" &vbcrlf
							end if
						rs.MoveNext 
						loop
						Rs.Close
						Set Rs = Nothing
					end if
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
		  	GetNodeText=Str
		 End If
		End Function
		
		Sub InitialSearch()
		  Dim FieldStr,SqlStr,TopStr,TopNum
		  If CurrPage<=0 Then CurrPage=1
		  If TopNum<>0 Then TopStr=" Top " & TopNum
		  Param=" Where 1=1 " 
		  Param=Param & " and SurveyID=" & FormID
		  ModelTable="KS_SurveyST"
		  FieldStr="*"
		  OrderStr=" Order By ID"
		  SqlStr="Select " & TopStr & " " & FieldStr & " From " & ModelTable & Param & OrderStr
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     TotalPut = Conn.Execute("select Count(1) from " & ModelTable & " " & Param)(0)
			 If TotalPut>TopNum And TopNum<>0 Then TotalPut=TopNum
			 If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrPage - 1) * MaxPerPage
			 Else
					CurrPage = 1
			 End If
			 Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
		  End If
		 RS.Close
		 Set RS=Nothing
		End Sub	
End Class
%>

 

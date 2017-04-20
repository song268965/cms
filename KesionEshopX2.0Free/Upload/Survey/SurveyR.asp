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
        Private KS,KSUser,ChannelID,ModelTable,Param,XML,Node,StartTime,FormID,TableName,Surveylx
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,stype,OrderStr,Userck,TimeLimit,StartDate,ExpiredDate,OnlyUser,AllowGroupID,UserOnce,ProjectName,ProjectContent,num,Score
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		'-------------------------主体------------------------------------
		Public Sub Kesion()
		
		   Dim Template,KSR,ItemNum,radioidstr,itemidstr,textastr
		   Set KSR = New Refresh
		   dim rs,Templ_url
		   FormID=KS.ChkClng(KS.G("SurveyID"))
		   if KS.G("action")="Surveyshow" then
				Call Surveyshow()
				Response.end()
		   end if
		   if  FormID=0 then Call KS.AlertHistory("ID错误!",-1):response.end
		   FCls.RefreshType = "searchIndex"   
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   RS.Open "Select top 1 Score,ProjectName,ProjectContent,OnlyUser,Template_a,Template_b,UserCk,TimeLimit,StartDate,ExpiredDate,UserCk,AllowGroupID,UserOnce From KS_Survey Where ID=" & FormID,conn,1,1
		   If RS.EOF And RS.Bof Then
			 Call KS.AlertHistory("ID错误!",-1):response.end
		   else
			 Templ_url=RS("Template_b")	
			 UserCk=KS.ChkClng(RS("UserCk"))
			 TimeLimit=rs("TimeLimit")
			 StartDate=rs("StartDate")
			 ExpiredDate=rs("ExpiredDate")
			 OnlyUser=KS.ChkClng(rs("OnlyUser"))
			 AllowGroupID=rs("AllowGroupID")
			  UserOnce=KS.ChkClng(rs("UserOnce"))
			  ProjectName=rs("ProjectName")
			  ProjectContent=rs("ProjectContent")
			  Score=RS("Score")
		   End If
		   RS.Close
		    Dim LoginTF:LoginTF=KSUser.UserLoginChecked()
			
		   if KS.G("action")="getItemNum" then
		     If OnlyUser=1 and LoginTF=false Then
			    call KS.AlertHistory("对不起，只会登录会员才能进入!",-1):response.end	
			End If
			If OnlyUser=1 then
				If Not KS.IsNul(AllowGroupID) And KS.FoundInArr(AllowGroupID, KSUser.GroupID, ",")=False Then
					call KS.AlertHistory("对不起，您所在的会员组不允许投票!",-1):response.end
				End If
			end if	
				If TimeLimit="1" Then
					if now<StartDate then call KS.AlertHistory("对不起，该投票于" & StartDate & "开放！",-1):response.end 
					if datediff("s",ExpiredDate,now)>0 then call KS.AlertHistory("对不起，该投票已在" & ExpiredDate  & "停止！",-1):response.end
			   End If
				If UserOnce<>0 Then
					If Conn.Execute("Select Count(ID) From KS_SurveyResult Where userip='" & KS.GetIP & "' and  SurveyID=" & FormID & "")(0)>=UserOnce  Then
						call KS.AlertHistory("对不起，最多只能投" & UserOnce & "次!",-1):response.end
					End If
					If Conn.Execute("Select Count(ID) From KS_SurveyResult Where UserName='" & KS.GetIP & "' and  SurveyID=" & FormID & "")(0)>=UserOnce  Then
				    KS.Die("<script>alert('对不起，最多只能投" & UserOnce & "次!');location.href='index.asp';</script>"):response.end 
	                End If
			    End If
				radioidstr=KS.G("radioidstr")
				Call ItemNumSql(radioidstr,0)
				itemidstr=KS.G("itemidstr")
				Call ItemNumSql(itemidstr,1)
				textastr=KS.G("textastr")
				Call ItemNumSql(textastr,2)
				
				if LoginTF=true and Score>0 Then  '增加积分
				   Call KS.ScoreInOrOut(KSUser.UserName,1,Score,"系统","参与问卷调查[" & ProjectName &"]所得!",-10,FormID)
				End If
				
				if UserCk=1 then
				response.write "<script>alert('恭喜，提交成功！');location.href='Survey.asp?id="&FormID&"';</script>"
				else
				response.write "<script>location.href='SurveyR.asp?Surveyid="&FormID&"';</script>"
				end if
		   End if
		   if userck=1 then Call KS.AlertHistory("对不起，没开放前台查看权限!",-1):response.end
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
				      If IsObject(XML) Then
					   num=1
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
						num=Num+1
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,根据您的查找条件,找不到任何相关记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName
				case "survey" 
				          select case sTokenName
							case "Surveyname" echo ProjectName
							case  "UserNum" 
							   dim rsc:set rsc=server.CreateObject("adodb.recordset")
							   rsc.open "select distinct username from KS_SurveyResult where SurveyID=" & FormID,conn,1,1
							   echo rsc.recordcount
							   rsc.close
							   set rsc=nothing
							case  "VoteNum" 
							   echo conn.execute("select Count(username) from KS_SurveyResult where SurveyID=" & FormID)(0)
						    case "showpage" echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							case "ProjectContent" echo ProjectContent
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
					dim rs,nstr
					Set Rs = Server.CreateObject("adodb.recordset")
					Rs.Open "select * from KS_SurveyItem where SurveySTID="& KS.ChkClng(GetNodeText("id")) & " Order By SurveyItemOrder" , Conn, 1, 1	             
					dim n:n=0
					dim tnum:tnum=conn.execute("select count(1) from KS_SurveyResult where SurveySTID=" & KS.ChkClng(GetNodeText("id")))(0)          
					echo "<table border='0'>"
					Do While Not rs.Eof
					    nstr=chr(65+n)
						n=N+1
						dim per
						'per=tnum
						if tnum>0 then
						per=round(KS.ChkClng(rs("ItemNum"))/tnum*100,2)
						else
						per=0
						end if
						
						echo "<tr><td width=350><b>"&  nstr  & "、"& rs("SurveyItemName") & "</b> &nbsp;(投票数:" & KS.ChkClng(rs("ItemNum")) &")</td><td><img src='../images/Default/bar.gif' width='"& per &"' height='15' align='absmiddle' /> " & per & "%"
						if rs("SurveyItemType")=1 then
							echo "<input name=""详细查看"" onClick=""Surveyshow("& rs("id") &");"" value=""详细查看"" type=""button"">"
						end if
						echo "</td></tr>"
					rs.MoveNext 
					loop
					echo "</table>"
					Rs.Close
					Set Rs = Nothing
			case "SurveyNum"
					Set Rs = Server.CreateObject("adodb.recordset")
					Rs.Open "select * from KS_SurveyItem where SurveySTID="&KS.ChkClng(GetNodeText("id")) & " Order By SurveyItemOrder" , Conn, 1, 1	
					dim ii
					Do While Not rs.Eof
					    nstr=chr(65+ii)
						ii=ii+1
						echo "{ data : [[0, " &  KS.ChkClng(rs("ItemNum")) &"]], label : '"& nstr &"、"&rs("SurveyItemName") &"' },"
					rs.MoveNext 
					loop
					Rs.Close
					Set Rs = Nothing
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
		
		
		Function ItemNumSql(txtstr,lx)
			dim i,ID,SurveySTID
			txtstr=ks.gottopic(txtstr,Len(txtstr)-1)
			txtstr=Split(txtstr,"|") 
			if UBound(txtstr)>=0 then
				for i=0 to UBound(txtstr) 
					if lx=2 then ID=KS.ChkClng(Replace(txtstr(i),"texta_","")) else ID=KS.ChkClng(txtstr(i))
					SurveySTID=Conn.Execute("Select top 1 SurveySTID from KS_SurveyItem  WHERE ID = "&ID&"")(0)
					Conn.Execute("UPDATE KS_SurveyItem SET ItemNum=ItemNum + 1 WHERE ID = "&ID&"")
					Conn.Execute("INSERT INTO KS_SurveyResult (SurveyID,SurveyItemID,Content,lx,userName,AddDate,userip,SurveySTID) VALUES ( "&FormID&","&ID&",'"& KS.G("texta_"&ID) &"',"& lx &",'"& KSUser.UserName &"',"&SQLNowString&",'"&KS.GetIP&"',"& SurveySTID &")")

				next	
			end if
		End Function
		Sub Surveyshow()
			if KS.ChkClng(KS.G("surveystid"))<>0 then
					dim rs
					Set Rs = Server.CreateObject("adodb.recordset")
					dim Param:Param=" where SurveyItemID="& KS.ChkClng(KS.G("surveystid"))
					dim CurrPage:CurrPage=KS.ChkClng(KS.G("page"))
					if CurrPage<=0 then CurrPage=1
					Rs.Open "select * from KS_SurveyResult "& Param &" ORDER BY ID DESC" , Conn, 1, 1
					echo "<!DOCTYPE HTML>"
					echo "<html>"
					echo "<head>"
					echo "<title>详细查看</title>"
					echo "<meta http-equiv=Content-Type content=""text/html; charset=utf-8"">"
					echo "<meta http-equiv=""X-UA-Compatible"" content=""IE=EmulateIE7"" /> "
					echo "<link href=""/images/style.css"" type=text/css rel=stylesheet>"
					echo "<style>"
					echo ".SurveyR_co{ border:1px solid #CCCCCC; background:#F5F5F5;font-size:14px ;margin:0px 50px 0px 10px ; padding:0px 5px 0px 5px ;}"
					echo ".SurveyR_list{ padding:10px;}"
					echo "</style>"
					echo "<div class='SurveyR_list'>" &vbcrlf
					If RS.Eof And RS.Bof Then
					   echo "该选项，没有用户反馈！"
				    Else
						 TotalPut = Conn.Execute("select Count(1) from KS_SurveyResult "& Param )(0)
						 If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
								RS.Move (CurrPage - 1) * MaxPerPage
						 End If
					    dim i:i=0
					Do While Not rs.Eof
						echo "<li style=""margin-top:10px;""><b>用户:</b> "& rs("username") & " &nbsp;<b>时间: "& rs("AddDate") & "</b> &nbsp;<b>IP: "& rs("userip")& "</b></li>"
						echo "<li style=""border-bottom:1px dashed #CCCCCC;""><div class=""SurveyR_co"">"&  Replace(rs("Content"),Chr(13),"<br>") & "</div> &nbsp;</li>"
						I=i+1
						if i>=MaxPerPage then exit do
					rs.MoveNext 
					loop
					End If
					echo "</div>"					
					echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
					Rs.Close
					Set Rs = Nothing
			end if
		end sub

		Sub InitialSearch()
		  Dim FieldStr,SqlStr,TopStr
		  If CurrPage<=0 Then CurrPage=1
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

 

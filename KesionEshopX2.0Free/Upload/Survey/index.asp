<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
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
        Private KS,ChannelID,ModelTable,Param,XML,Node,StartTime,FormID,TableName,Surveylx
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,stype,OrderStr
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		'-------------------------主体------------------------------------
		Public Sub Kesion()

		   Dim Template,KSR
		 Set KSR = New Refresh
		   dim rs,Templ_url
		   Templ_url = KSR.LoadTemplate("Template/index.html" ) 
		   Template = KSR.KSLabelReplaceAll(Templ_url)
		   Set KSR = Nothing
		   InitialSearch
		   Scan Template
	   End Sub
	 
	    '-------------------------------------------------------------
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
				case "search" 
				          select case sTokenName
						    case "showpage" echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							case "leavetime" 
							   dim leavetime:leavetime=FormatNumber((timer-starttime),5)
							   if leavetime<1 then leavetime="0"&leavetime
							   echo leavetime
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "status" 
			  If GetNodeText("timelimit")="1" Then
				 if datediff("s",GetNodeText("expireddate"),now)>0 then
				   echo "<span style='color:red'>[已截止]</span>"
				 else
				   echo "<span style='color:blue'>[进行中]</span>"
				 end if
			  End If
			case "showvotelink"
			  If (GetNodeText("timelimit")="1" and datediff("s",GetNodeText("expireddate"),now)<0) or GetNodeText("timelimit")="0" Then
				   echo "<a href='Survey.asp?id=" & GetNodeText("id") & "' target='_blank'>[进入问卷]</a>"
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
		  if NodeName="title" Then Str= GetNodeText("KS_title")
		  	GetNodeText=Str
		  End If
		End Function
		
		
		Sub InitialSearch()
		  Dim FieldStr,SqlStr,TopStr,TopNum
		  ChannelID=KS.ChkClng(Request("M"))
		  CurrPage=KS.ChkClng(Request("Page"))
		  If CurrPage<=0 Then CurrPage=1
		  Key=KS.CheckXSS(KS.R(KS.S("Key")))
		  stype=KS.ChkClng(Request("stype"))
		  
		 if request("classid")<>"" and request("classid")<>"0" then
		   If ChannelID<>102 Then
		     Param=Param & " And Tid In(" & KS.GetFolderTid(KS.S("ClassID")) & ")"
		   end if
		 end if
		 
		  If TopNum<>0 Then TopStr=" Top " & TopNum
		  Param=" Where 1=1 "
		  ModelTable="KS_Survey"
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

	
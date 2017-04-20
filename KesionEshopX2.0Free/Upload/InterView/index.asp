<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 9.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New CartCls
KSCls.Kesion()
Set KSCls = Nothing

Class CartCls
        Private KS, KSR,KSUser,DomainStr,Template,TotalPut,MaxPerPage,CurrPage
		Private XML,Node,LoginTF
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		  Set KSUser = New UserCls
		  Set KSR = New Refresh
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		Public Sub Kesion()
		  Template = KSR.LoadTemplate("{@TemplateDir}/访谈系统/访谈频道首页.html")
		  Fcls.RefreshFolderID = "0"        '设置当前刷新目录ID 为"0" 以取得通用标签
          Template = KSR.KSLabelReplaceAll(Template)
          CurrPage=KS.ChkClng(KS.S("Page"))
		  If CurrPage<0 Then CurrPage=1
		   Initial()
		   Scan Template
	   End Sub
	   
	   Sub Initial()
	      Dim SqlStr:SqlStr="Select * From KS_InterView order by id desc"
		  'ks.echo sqlstr
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     TotalPut = RS.RecordCount
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
	  
	   Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "loop"
				      If IsObject(XML) Then
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,没有访谈记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName
				case "search" 
				          select case sTokenName
						    case "menu"  SearchMenu
						    case "showpage" echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							
							
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "intro" 
			 Dim Intro:intro=KS.Gottopic(KS.LoseHtml(GetNodeText("content")),160)
			 Intro=Replace(Intro,"&nbsp;","")
			 echo intro
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
	  
End Class
%>

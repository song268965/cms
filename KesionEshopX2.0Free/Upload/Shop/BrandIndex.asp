<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New ClassCls
KSCls.Kesion()
Set KSCls = Nothing

Class ClassCls
        Private KS,KSRFObj,BrandID,TempStr,SQL,Intro
		Private FileContent,SqlStr,ID,CurrentPage,RSObj,MaxPerPage,TotalPut
		Private BrandName,BrandArr,ClassArr,Param,FolderID,K,PageNum,TS,TN
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		  MaxPerPage =20    '定义每页数量
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  
		  ID=KS.ChkClng(KS.S("ID"))
		  CurrentPage=KS.ChkClng(KS.S("Page"))
		  If CurrentPage<=0 Then CurrentPage=CurrentPage+1

         If ID<>0 Then
			 SqlStr= "Select Top 1 ID,TN,TS From KS_Class Where ClassID="& ID
			 Set RSObj=Server.CreateObject("Adodb.Recordset")
			 RSObj.Open SqlStr,Conn,1,1
			 IF RSObj.Eof Then
			  Call KS.Alert("非法参数!","")
			  Exit Sub
			 End IF
			 FolderID=RSObj("ID")
			 TS=RSObj("TS")
			 TN=RSObj("tn")
			 RSObj.Close:Set RSObj=Nothing
		 End If
		  
			 FileContent = KSRFObj.LoadTemplate(KS.Setting(135))
			 FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			 LoadClassList()
			 LoadBrandList()		 
			 Scan FileContent
		End Sub
		
		Sub ParseArea(sTokenName, sTemplate)
			Select Case lcase(sTokenName)
			 case "brandlist"
				  If IsArray(BrandArr) Then
					For K=0 To Ubound(BrandArr,2)
					  Scan sTemplate
					Next
				  Else 
				   	  if Not KS.IsNul(KS.S("Letter")) Then
					    echo "找不到首字母为[" & KS.CheckXSS(KS.S("Letter")) & "]的品牌!"
					  Elseif KS.ChkClng(KS.S("ID"))=0 Then
					    echo "找不到品牌!"
					  Else
					   echo "分类[" & KS.C_C(FolderID,1) & "]下找不到品牌!"
					  End If

				  End If
				
			End Select 
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
		     on error resume next
		    Dim I
			Select Case lcase(sTokenType)
			 case "brand" 
			   select case lcase(sTokenName)
			     case "id" echo BrandArr(0,k)
				 case "name"
				   'if Not KS.IsNul(KS.S("Letter")) Then
				   'echo KS.CheckXSS(KS.S("Letter"))
				   'else 
				   echo BrandArr(1,k)
				   'end if
				 case "url" echo "brand.asp?id=" & KS.C_C(FolderID,9) & "&brandid=" & BrandArr(0,k)
				 case "photourl" If BrandArr(2,k)="" Or IsNull(BrandArr(2,k)) Then echo "../Images/nopic.gif" else echo BrandArr(2,k)
				 case "showclass"
				   	  if Not KS.IsNul(KS.S("Letter")) Then
					    echo "<strong>首字母为[" & KS.CheckXSS(KS.S("Letter")) & "]的品牌</strong>"
					  Elseif KS.ChkClng(KS.S("ID"))=0 Then
					    echo "<strong>品牌中心</strong>"
					  Else
						  If IsArray(ClassArr) Then
						   For I=0 To Ubound(ClassArr,2)
							echo "<li><a href=""?id=" & ClassArr(2,i) & """>" & ClassArr(1,i) & "</a></li>"
						   Next
						  End If
					  End If

				 case "shownav"
				  If TS<>"" Then
					  Dim TSArr:TSArr=Split(TS,",")
					  For I=0 To Ubound(TSArr)-1
					   echo " &gt; "
					   echo "<a href=""?id=" & KS.C_C(TSArr(i),9) & """>" & KS.C_C(TSArr(i),1) & "</a>"
					  Next
				  End If
				 case "amount"
				   echo conn.execute("select count(id) from ks_product where tid in (" & KS.GetFolderTid(FolderID) & ") and brandid=" & BrandArr(0,k))(0)
			   end select
			 case "foot"
			    if lcase(sTokenName)="showpage" then echo KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)
		    End Select
		End Sub
		
		Sub LoadClassList()
		 Dim SqlParam
		 If FolderID="" Then SqlParam=" and tj=1" Else SqlParam=" and Tn='" & FolderID & "'"
		 Dim RS:Set RS=Conn.Execute("SELECT ID,FolderName,ClassID From KS_Class Where Channelid=5" & SqlParam &"  Order By FolderOrder")
		 If Not RS.Eof Then
		  ClassArr=RS.GetRows(-1)
		 End If
		 If Not IsArray(ClassArr) Then
		  RS.Close
		  Set RS=Conn.Execute("SELECT ID,FolderName,ClassID From KS_Class Where Channelid=5 And Tn='" & TN & "' Order By FolderOrder")
		  If Not RS.Eof Then
		  ClassArr=RS.GetRows(-1)
		  End If
		  
		 End If
		 RS.Close
		 Set RS=Nothing
		End Sub
		
		Sub LoadBrandList()	
		  if KS.ChkClng(KS.S("ID"))<>0 Then
		    If FolderID="" Then
			 If IsArray(ClassArr) Then FolderID=ClassArr(0,0) 
			End If
			Param = " WHERE B.ClassID in (" & KS.GetFolderTid(FolderID) & ")"
		  end if
		  If Not KS.IsNul(KS.S("Letter")) Then Param=" Where A.firstAlphabet='" & ucase(KS.S("Letter")) & "'"

			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select distinct a.id,a.BrandName,a.photourl from KS_ClassBrand A Inner Join KS_ClassBrandR B ON A.ID=B.BrandID" & Param & "", Conn, 1, 1
			If Not RS.Eof Then
						  TotalPut=RS.Recordcount
						  If CurrentPage < 1 Then CurrentPage = 1
						  if (TotalPut mod MaxPerPage)=0 then
								PageNum = TotalPut \ MaxPerPage
						  else
								PageNum = TotalPut \ MaxPerPage + 1
						  end if
				
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								 RS.Move (CurrentPage - 1) * MaxPerPage
							Else
								 CurrentPage = 1
							End If
							BrandArr=RS.GetRows(MaxPerPage)
			 End If
			 RS.Close:Set RS=Nothing		
	   End Sub
	   
End Class
%>

 

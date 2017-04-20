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
        Private KS,KSRFObj,BrandID,TempStr,SQL,Intro,ClassName,ChannelID
		Private FileContent,SqlStr,ID,CurrentPage,RSObj,MaxPerPage,TotalPut
		Private BrandName,BrandPhotoUrl,Product,Param,CurrentPageStr,FolderID,K,PageNum,TS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		  MaxPerPage = 20   '定义每页数量
		  ChannelID=5
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  
		  ID=KS.ChkClng(KS.S("ID"))
		  BrandID=KS.ChkCLng(KS.S("BrandID"))
		  CurrentPage=KS.ChkClng(KS.S("Page"))
		  If CurrentPage<=0 Then CurrentPage=CurrentPage+1

		 IF BrandID=0 Then Exit Sub
		 
		 If ID<>0 Then
			 SqlStr= "Select Top 1 ID,TN,FolderName,TS From KS_Class Where ClassID="& ID
			 Set RSObj=Server.CreateObject("Adodb.Recordset")
			 RSObj.Open SqlStr,Conn,1,1
			 IF Not RSObj.Eof Then
				 Call FCls.SetClassInfo(ChannelID,RSObj("ID"),RSObj("TN"))
				 ClassName=RSObj("FolderName")
				 TS=RSObj("TS")
				 FolderID=RSObj("ID")
			 End If
			 RSObj.Close:Set RSOBj=Nothing
		 End If
			 FileContent = KSRFObj.LoadTemplate(KS.Setting(136))
			 Set RSObj=Conn.Execute("select top 1 * From KS_ClassBrand Where ID=" & BrandID)
			 If Not RSObj.Eof Then
			   BrandName    = RSObj("BrandName")
			   Intro        = RSObj("Intro")
			   BrandPhotoUrl= RSObj("PhotoUrl")
			   If KS.IsNul(BrandPhotoUrl) Then BrandPhotoUrl=KS.Setting(3) & "images/nopic.gif"
			 Else
			   Set RSObj=Nothing
			   KS.Die "error!"
			 End If
			 Set RSOBj=Nothing
			 FileContent = Replace(FileContent,"{$GetBrandName}",BrandName)			
			 FCls.BrandName="[品牌] " & BrandName
			 FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			 LoadProductList()		 
			 Scan FileContent
		End Sub
		
		Sub ParseArea(sTokenName, sTemplate)
			Select Case lcase(sTokenName)
			 case "productlist"
			  If IsArray(Product) Then
			    For K=0 To Ubound(Product,2)
				  Scan sTemplate
				Next
			  Else 
			     echo "找不到属于品牌[" & BrandName & "]的商品!" 
			  End If
			End Select 
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			 case "brand" 
			   select case lcase(sTokenName)
			    case "classid" echo id
			    case "classname" echo ClassName
			    case "id" echo BrandID
				case "name" echo BrandName
				case "intro" echo intro
				case "photourl" echo BrandPhotoUrl
				 case "shownav"
				  If TS<>"" Then
					  Dim I,TSArr:TSArr=Split(TS,",")
					  For I=0 To Ubound(TSArr)-1
					   echo " &gt; "
					   echo "<a href=""showbrand.asp?id=" & KS.C_C(TSArr(i),9) & """>" & KS.C_C(TSArr(i),1) & "</a>"
					  Next
				  End If

			   end select
			 case "product"
			   select case lcase(sTokenName)
			     case "id" echo Product(0,k)
				 case "name" echo Product(1,k)
				 case "url" echo KS.GetItemUrl(ChannelID,Product(2,K),Product(0,K),Product(4,K),Product(7,K))
				 case "photourl" If Product(3,k)="" Or IsNull(Product(3,k)) Then echo "../Images/nopic.gif" else echo Product(3,k)
				 case "pricemarket" echo product(5,k)
				 case "price" echo product(6,k)
			   end select
			 case "foot"
			    if lcase(sTokenName)="showpage" then echo KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)
		    End Select
		End Sub
		
		Sub LoadProductList()	
			Param = " WHERE BrandID=" & BrandID & " AND Verific=1 AND DelTF=0"
		    If FolderID<>"" And FolderID<>"0" Then  Param=Param & " and Tid in (" & KS.GetFolderTid(FolderID) & ")"
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select id,title,tid,photourl,Fname,price,price_member,adddate from ks_product " & Param & " order by IsTop Desc,ID Desc", Conn, 1, 1

				If Not RS.Eof Then
						  TotalPut= Conn.Execute("select Count(id) from ks_product" & Param)(0)
						  If CurrentPage < 1 Then CurrentPage = 1
						  if (TotalPut mod MaxPerPage)=0 then
								PageNum = TotalPut \ MaxPerPage
						  else
								PageNum = TotalPut \ MaxPerPage + 1
						  end if
				
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								 RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Product=RS.GetRows(MaxPerPage)
				 End If
				 RS.Close:Set RS=Nothing		
	   End Sub
	   
End Class
%>

 

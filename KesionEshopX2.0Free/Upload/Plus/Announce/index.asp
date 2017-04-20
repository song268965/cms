﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Announce
KSCls.Kesion()
Set KSCls = Nothing

Class Announce
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		
		    Call KS.CheckAppStatusAndDie("announce")
			
			Dim AnnounceID, FileContent,RefreshRS, KMRFObj
			Set KMRFObj = New Refresh
		   AnnounceID = KS.ChkClng(request.QueryString)
		   FileContent = KMRFObj.LoadTemplate("{@TemplateDir}common/announce.html")
		   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		   Set RefreshRS = Server.CreateObject("Adodb.Recordset")
		   RefreshRS.Open "Select top 1 Title,Author,AddDate,Content From KS_Announce Where ID=" & AnnounceID, Conn, 1, 1
		   If Not RefreshRS.EOF Then
			FileContent = ReplaceAnnounceContent(RefreshRS, FileContent)     '替换公告内容标签为内容
		   Else
			FileContent = "参数传递错误!"
		   End If
		   RefreshRS.Close:Set RefreshRS = Nothing
		   FileContent=KMRFObj.KSLabelReplaceAll(FileContent)
		   Set KMRFObj = Nothing
		   Response.Write FileContent   '输出公告内容页
		End Sub
		'*********************************************************************************************************
		'函数名：ReplaceAnnounceContent
		'作  用：替换公告内容页标签为内容
		'参  数：FileContent待替换的内容
		'*********************************************************************************************************
		Function ReplaceAnnounceContent(RefreshRS, FileContent)
			   If InStr(FileContent, "{$GetAnnounceTitle}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceTitle}", RefreshRS(0))
			   End If
			   If InStr(FileContent, "{$GetAnnounceAuthor}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceAuthor}", RefreshRS(1))
			   End If
			   If InStr(FileContent, "{$GetAnnounceDate}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceDate}", RefreshRS(2))
			   End If
			   If InStr(FileContent, "{$GetAnnounceContent}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceContent}", RefreshRS(3))
			   End If
			   ReplaceAnnounceContent = FileContent
		End Function

End Class
%>

 

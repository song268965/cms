﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.WebFilesCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Frame
KSCls.Kesion()
Set KSCls = Nothing

Class Frame
        Private KS,KSUser
		Private TopDir
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
		 Response.Buffer = True
		Response.Expires = -1
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
		 if request.ServerVariables("http_referer")="" then
		  ks.die "error!"
		 end if
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>window.close();</script>"
		  Exit Sub
		End If
     
	  If ChannelID<100 then
	      
		If KS.C_S(ChannelID,16)=0 Then
		  Response.Write "<script>alert('对不起，您没有选择已上传图片的权限，请与网站管理员联系!');window.returnValue='';window.close();</script>"
		  Exit Sub
		End IF
	 end if
		'If ChannelID<5000 Then
		'   TopDir=KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID"))
		'Else
		   TopDir=KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID"))
		'End If
		
		IF TopDir="" Then 
			TopDir=KS.S("topdir")
		Else
		 TopDir=TopDir &"/"
		End IF
        Call KS.CreateListFolder(TopDir)
		Dim WFCls:Set WFCls = New WebFilesCls
		Call WFCls.Kesion(ChannelID,TopDir,"select",20,"选择图片","Images/Css.css")
		Set WFCls = Nothing
			Response.Write "<script>"
		    Response.Write "window.onunload=SetReturnValue;"
            Response.Write "function SetReturnValue()"
            Response.Write "{"
            Response.Write "	if (typeof(window.returnValue)!='string') window.returnValue='';"
            Response.Write "}</script>"

      End Sub
End Class
%> 

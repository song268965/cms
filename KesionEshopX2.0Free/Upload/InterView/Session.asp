<%
'****************************************************
' Software name:Kesion CMS 9.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCheck
Set KSCheck = New LoginCheckCls1
KSCheck.Run()
Set KSCheck = Nothing

Class LoginCheckCls1
		Private ComeUrl
		Private TrueSiteUrl
		Private AdminDirStr
		Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		
		
		Sub Run()
		 
		  
		  If KS.C("AdminName")<>"" and KS.C("AdminPass")<>"" Then '管理员
		  
		     Dim ChkRS:Set ChkRS = Server.CreateObject("ADODB.RecordSet")
			 ChkRS.Open "Select top 1 * From KS_Admin Where UserName='" & KS.R(KS.C("AdminName")) & "' and password='" & KS.C("AdminPass") & "'",Conn, 1, 1
			 If ChkRS.EOF And ChkRS.BOF Then
			     ChkRS.Close:Set ChkRS=Nothing
				 Response.Write ("<script>top.location.href='../';</script>")
				 Response.End
			 End If
		    ChkRS.Close:Set ChkRS = Nothing
		  ElseIf KS.IsNul(KS.C("InterViewUserName")) Or KS.IsNul(KS.C("InterViewPass")) Or KS.IsNul(KS.C("InterRole"))="" Or KS.IsNUL(KS.C("InterViewID")) Then
			Response.Write ("<script>top.location.href='../';</script>")
			Response.End()
		  Else
			 Set ChkRS = Server.CreateObject("ADODB.RecordSet")
			 ChkRS.Open "Select top 1 * From KS_InterView Where ID=" & KS.ChkClng(KS.C("InterViewID")) & " and HostUserID='" & KS.C("InterViewUserName") &"' and HostUserPass='" & KS.C("InterViewPass") &"'",Conn, 1, 1
			 If ChkRS.EOF And ChkRS.BOF Then
			     ChkRS.Close:Set ChkRS=Nothing
				 Response.Write ("<script>top.location.href='login.asp?id=" &KS.C("InterViewID") &"' ;</script>")
				 Response.End
			 End If

		   ChkRS.Close:Set ChkRS = Nothing
		 End If
		End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/ClubCls.asp"-->
<!--#include file="Include/3GCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Display
KSCls.Kesion()
Set KSCls = Nothing

Class Display
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Fcls.CallFrom3g="true"
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    FCls.ChannelID=11  '论坛系统
			dim Club:set Club=new ClubDisplayCls
			 Club.kesion
			 Set Club=Nothing
		End Sub
End Class
%>

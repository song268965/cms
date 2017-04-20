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
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 	Dim Letter:Letter=KS.S("Letter")
			%>
			<html>
			<head>
			 <title>按字母查看品牌</title>
			 <style>
			 body{font-size:12px;margin:0px;padding:0px}
			 ul,li{margin:0px;padding:0px}
			 li{list-style-type:none;height:23px;line-height:23px;padding-left:5px}
			 .title{font-weight:bold;background:#f1f1f1;height:20px;line-height:20px;margin-bottom:3px}
			 a{color:#000000;text-decoration:none}
			 </style>
			</head>
			<body>
			<%Dim I,sql
			For I=65 To 90
				 Response.Write "<a name='" & chr(I) & "'></a><div class='title'>" & chr(I) &"</div>"
			     sql="select * from ks_classbrand where firstAlphabet='" & UCase(chr(i)) & "' order by id"
				 dim rs:set rs=server.createobject("adodb.recordset")
				 rs.open sql,conn,1,1
				 response.write "<div><ul>"
				 do while not rs.eof 
				  response.write "<li><a href='brand.asp?brandid=" & rs("id") & "' target='_blank'>" & rs("brandname") & " "&rs("BrandEname")& "</a></li>"
				  rs.movenext
				 loop
				 rs.close
				 set rs=nothing
				 response.write "</ul></div>"
			Next
			%>
			</body>
			</html>
			<%
	   End Sub
	   
End Class
%>

 

﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.5
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   If KS.C_S(13,21) = "0" Then KS.Die "<script>alert('对不起，本频道未开启！');location.href='" & KS.Setting(3) & "';</script>"
		   
		    Dim TpDir:TpDir=KS.SSetting(64)
			 If KS.IsNUL(TpDir) Then TpDir=KS.Setting(3) & KS.Setting(90) & "空间模板/企业通用/product_index.html"
			 Dim Template:Template = KSR.LoadTemplate(TpDir)

		   
		   FCls.RefreshType = "enterprisepro" '设置刷新类型，以便取得当前位置导航等
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   call getclasslist()
		   Template=Replace(Template,"{$ShowClass}",str)
		   Template=KSR.KSLabelReplaceAll(Template)
		   Response.Write Template  
		End Sub
		Sub GetClassList()
		 Dim RS,I,RSS,N
		 Set RS=Conn.Execute("select id,foldername,classid from ks_class where classtype=1 and channelid=5 and tj=2 order by root,folderorder")
		 str="<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""productclass"">" & vbcrlf
		 n=1
		 Do While Not RS.Eof
		 str=str & "<tr>" & vbcrlf
		 for i=1 to 2
		   if n mod 2=0 then
		   str=str & "<td width=""50%"" style=""background:#F7F7F7;padding:5px"">" & vbcrlf
		   else
		   str=str & "<td width=""50%"" style=""padding:5px"">" & vbcrlf
		   end if
		   str=str & "<div class=""classname""> <a href=""list.asp?id=" & rs(2) & """><u>" & rs(1) &"</u> <span class='num'>(" & conn.execute("select count(id) from KS_Product where tid in(select id from ks_class where ts like '%" & rs(0) &"%')")(0)&")</span></a></div>" & vbcrlf
		   str=str & "<div class=""seconditem"">"
		   KS.LoadClassConfig()
		   Dim Node,k
		   k=0
		   For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1 && @ks13=" & rs(0) & "]")
		     if k<>0 then str=str & " | "
		 	 str=str & "<a href='list.asp?id=" & Node.SelectSingleNode("@ks9").text & "'>" & Node.SelectSingleNode("@ks1").text & "</a>"
			 k=K+1
		   Next
		   str=str & "</div>"
		   str=str & "</td>" & vbcrlf
		   rs.movenext
		   if rs.eof then exit for
		 next
		 str=str & "</tr>"
		  n=n+1
		 Loop
		 str=str & "</table>" & vbcrlf
		 
		End Sub
		
End Class
%>

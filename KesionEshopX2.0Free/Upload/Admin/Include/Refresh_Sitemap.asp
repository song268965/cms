<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************%>
<!DOCTYPE html>
<html>
<title>Google Sitemap</title>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="Admin_Style.CSS" rel="stylesheet" type="text/css">
<script src="../../ks_inc/jquery.js"></script>

<%


dim xmlstr,lastmod
dim sql_KS_Class,SqlStr,rs,rsclass,i,Classpath
dim sitemappath
Dim KS:Set KS=New PublicCls


'=========================主程序==========================
If KS.G("Action")<>"" Then
    Dim changefreq:changefreq=KS.G("changefreq")
	Dim prioritynum:prioritynum=KS.ChkCLng(KS.G("prioritynum"))
	dim tmFile,objFso,smw
	sitemappath=KS.Setting(3)&"sitemap.xml"
	Set objFso = KS.InitialObject(KS.Setting(99))

	if KS.G("Action")="creategoogle" then
		If prioritynum=0 then prioritynum=15
		Dim big:big=KS.G("Big")
		Dim SQL,K
		Set RS=KS.InitialObject("ADODB.RECORDSET")
		RS.Open "Select BasicType,ChannelTable,ChannelID From KS_Channel Where ChannelStatus=1 And ChannelID<>6 And BasicType<=8 Order By ChannelID",Conn,1,1
		SQL=RS.GetRows(-1)
		RS.Close

		xmlstr="<?xml version=""1.0"" encoding=""UTF-8""?>"&vbcrlf
		xmlstr=xmlstr&"<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">"&vbcrlf
	
		For K=0 To Ubound(SQL,2)
		 Select Case  SQL(0,K)
		  Case 1 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate"
		  Case 2 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 3 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 4 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 5 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		  Case 7 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		  Case 8 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		 End Select
		
		SqlStr=SqlStr & " from "& SQL(1,K) & " where verific=1 and deltf=0 order by id desc"
		rs.Open SqlStr,conn,1,1
		for i=1 to rs.RecordCount
			xmlstr=xmlstr&"    <url>"&vbcrlf
			xmlstr=xmlstr&"        <loc><![CDATA["&KS.GetItemUrl(SQL(2,K),RS(2),RS(0),RS(5),RS(7))&"]]></loc>"&vbcrlf
			xmlstr=xmlstr&"        <lastmod>" & GetDate(rs(7)) & "</lastmod>"&vbcrlf
			xmlstr=xmlstr&"        <changefreq>"&changefreq&"</changefreq>"&vbcrlf
			xmlstr=xmlstr&"        <priority>"&big&"</priority>"&vbcrlf
			xmlstr=xmlstr&"    </url>"&vbcrlf
			rs.movenext 
		next
		rs.close
	  Next
	'=sitemap===============================================================================================================
		xmlstr=xmlstr&"</urlset>"
	
	
		'==============写入sitemap======================
		Call KS.WriteTOFile(sitemappath,xmlstr)
	   '===========sitemap================================
	
	response.write("<script language='JavaScript' type='text/JavaScript'>")
	response.write("function yy() {")
	response.write("overstr.innerHTML='<div align=center>恭喜,sitemap.xml生成完毕！<br><br><a href=" & KS.Setting(3) & "sitemap.xml target=_blank>点击查看生成好的sitemap.xml文件</a></div>'; }")
	response.write("</script>")
	
	elseif  KS.G("Action")="createbaidu" then
	
		xmlstr=""
		Dim Num:Num=0
		Set RS=KS.InitialObject("ADODB.RECORDSET")
		
		SqlStr="Select top " & prioritynum & " InfoID,Title,Tid,0,0,Fname,ChannelID,AddDate From KS_ItemInfo Where deltf=0 And verific=1 Order By ID desc"
		rs.Open SqlStr,conn,1,1
		for i=1 to rs.RecordCount
			xmlstr=xmlstr&KS.GetItemUrl(RS("ChannelID"),RS(2),RS(0),RS(5),RS(7)) &vbcrlf
			Num=Num+1
			If Num>=50000 Then Exit For
			rs.movenext 
		next
		rs.close
	'=sitemap===============================================================================================================
		
		
	
		'==============写入news.txt======================
		Dim NewsPath:NewsPath=KS.Setting(3) &"sitemap.txt"
		Call KS.WriteTOFile(NewsPath,xmlstr)
	   '===========sitemap================================

	
	response.write("<script>")
	response.write("function yy() {")
	response.write("overstr.innerHTML='<div align=center>恭喜,360/百度sitemap.txt生成完毕！<br><br><a href=" & KS.Setting(3) & "sitemap.txt target=_blank>点击查看生成好的sitemap.txt文件</a></div>'; }")
	response.write("</script>")
	end if
	
	'===================================================
		set rs=nothing
End If


response.write("<script>")
response.write("function ll() { ")
response.write("overstr.innerHTML='<div align=center>正在生成，请耐心等待。。。<br></div>'; } ")
response.write("</script>")

set rs=nothing
conn.Close:set conn=nothing
'===================================================结束
Function GetDate(DateStr)
	if KS.G("Action")="creategoogle" then
	GetDate=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
	else
	GetDate=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)& " " & Right("0" &hour(DateStr),2) &":" & Right("0" &minute(DateStr),2)& ":" & Right("0" & Second(DateStr),2)
	end if
End Function
%>


</head>

<body <%if request("action")<>"" then response.write "onLoad='yy()'" end if%>>
<div class="pageCont2">

 <div class='tabTitle'>针对百度或360搜索引擎的SiteMap 生成操作</div>

<table width="100%" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td><div id="overstr" class="attention"></div></td>
  </tr>
</table>




<form id="form1" name="bqsitemapform" method="post" action="?action=createbaidu">

<table width="100%" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr class="Title" align="center">
    <td>★360/百度Sitemap生成操作(<a href='http://zhanzhang.baidu.com/wiki/93' target='_blank'>查看百度Sitemap工具详情</a>)</td>
  </tr>


  <tr class="tdbg">
    <td height="35" style="text-align:center">生成URL数：
      <input name="prioritynum" type="text" class="textbox" id="prioritynum" value="10000" size="6" />
      条信息内容为最高注意度(最多50000条)	 </td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td height="45" align="center"><input name="Submit1"  class="button" onClick="ll();" type="submit" id="Submit1" value="开始生成sitemap" /></td>
  </tr>
</table>
</form>

<br/>

<div style="display:none">
<form id="form1" name="bqsitemapform" method="post" action="?action=creategoogle">

<table width="600" border="0" align="center" cellpadding="6" cellspacing="0" class="border">
  <tr class="Title">
    <td>★生成符合GOOGLE规范的XML格式地图操作<a href='http://www.google.com/webmasters/sitemaps/login' target='_blank'>(查看介绍)</a></td>
  </tr>

  <tr class="tdbg">
    <td height="18">更新频率：
      <select name="changefreq" id="changefreq">
        <option value="always ">频繁的更新</option>
        <option value="hourly">每小时更新</option>
        <option value="daily" selected="selected">每日更新</option>
        <option value="weekly">每周更新</option>
        <option value="monthly">每月更新</option>
        <option value="yearly">每年更新</option>
        <option value="never">从不更新</option>
      </select></td>
  </tr>
  <tr class="tdbg">
    <td height="35">每个系统调用：
      <input name="prioritynum" type="text" class="textbox" id="prioritynum" value="15" size="6" />条信息内容为最高注意度
	 </td>
  </tr>
  <tr class="tdbg">
    <td height="35">注 意 度：
      <input name="big" type="text" class="textbox" id="big" value="0.5" size="6" />0-1.0之间,推荐使用默认值

	  <br>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td height="45" align="center"><input name="Submit1"  class="button" onClick="ll();" type="submit" id="Submit1" value="开始生成sitemap" /></td>
  </tr>
</table>
</form>

</div>
</div>

</body>
</html>

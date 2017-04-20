<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New User_ScoreDetail
KSCls.Kesion()
Set KSCls = Nothing

Class User_ScoreDetail
        Private KS,KSCls
		Private MaxPerPage,RS,TotalPut,TotalPages,I,Page,SQL,ComeUrl
		Private Sub Class_Initialize()
		  MaxPerPage=15
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub

       Sub Kesion()
	     select case KS.G("action")
		 case "delall"
		 	qiandao_delall
		 case "del"
		 	qiandao_del
		 end select
		  Response.Write "<!DOCTYPE html><html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script src=""../../ks_inc/jquery.js""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
	     	 If Not KS.ReturnPowerResult(0, "KMUA10005") Then
			  response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
			Response.Write"<div class='topdashed quickLink' style='text-align:left'>签到明细: <a href='KS.qiandao.asp'>所有签到明细</a> |"
			%>
			<a href="KS.qiandao.asp?Action=day" >今日签到明细</a> | 
			<a href="KS.qiandao.asp?Action=month">本月签到明细</a> | 
			<a href="javascript:void(0);" onclick="if(confirm('是否删除所有签到!')){location.href='KS.qiandao.asp?Action=delall'};" >点击删除所有签到</a></div>
			<%
		dim Param
		%>
<div class="tableTop">
<form action="?" name="myform" method="post" >
   <div>
      &nbsp;<strong>按用户搜索=></strong>
     &nbsp;用户名:<input type="text" class='textbox' name="keyword">
      &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
      </div>
</form>
</div>
<div class="pageCont2 mt20">            
<div class="tabTitle">会员签到管理</div>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr class="sort">
    <td width="10%" align="center"><strong> 用户名</strong></td>
    <td width="15%" align="center"><strong>签到时间</strong></td>
    <td width="10%" align="center"><strong>IP地址</strong></td>
    <td width="10%"  align="center"><strong>签到心情</strong></td>
    <td width="40%" align="center"><strong>签到内容</strong></td>
    <td width="10%" align="center"><strong>状态</strong></td>
    <td align="10%"><strong>操作</strong></td>
  </tr>
  <%
  Page	= KS.ChkClng(request("page"))
  If Page<=0 Then Page=1
  Set RS=Server.CreateObject("ADODB.RecordSet")
  If KS.G("action")="Status" Then
    Param=" Status=1"
  else
  	Param=" Status=0"	
  End If
  If KS.G("action")="day" Then
  	Param=Param & "  and year(AddDate)=" & Year(now()) & " and month(AddDate)=" & month(now()) &" and day(AddDate)=" & day(now())  & " "
  End If
  
  If KS.G("action")="month" Then
  	Param=Param & "  and year(AddDate)=" & Year(now()) & " and month(AddDate)=" & month(now()) &" "
  End If
  
  if request("keyword")<>"" then
    Param= Param & " and username='" & request("keyword") & "'"
  end if
  If Param="" Then Param=" 1=1"
  Dim FieldStr,SQLStr
  FieldStr="id,content,qdxq,adddate,username,userip,status"
 
	SQLStr=KS.GetPageSQL("KS_qiandao","ID",MaxPerPage,Page,1,Param,FieldStr)
	Set RS = Server.CreateObject("AdoDb.RecordSet")
	RS.Open SQLStr, conn, 1, 1

		   
	If RS.Eof And RS.Bof Then
	 Response.Write "<tr><td colspan=9 align=center height=25 class='splittd'>还没有会员签到记录！</td></tr>"
	Else
                    TotalPut=conn.execute("select count(1) from KS_qiandao Where " & Param)(0)
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
   End If
%>		
</table>


</div>

<%End Sub

Sub ShowContent
 Dim InScore,OutScore
For i=0 To Ubound(SQL,2)
	%>
  <tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
    <td class="splittd" width="80" align="center"><%=SQL(4,i)%></td>
    <td class="splittd" align="center"><%=SQL(3,i)%></td>
    <td class="splittd" align="center"><%=SQL(5,i)%></td>
    <td class="splittd" align="center"> <img src="../../images/emot/<%=SQL(2,i)%>.gif"  style="height:20px;"></td>
    <td class="splittd" align="center"><%=SQL(1,i)%></td>
    <td class="splittd" align="center"><% if KS.ChkClng(SQL(6,i))=0 then Response.Write("<font color=""green"">已签到</font>") else Response.Write("<font color=""#FF0000"">未签到</font>")%></td>
	<td class="splittd" align="center"><a href="javascript:void(0);" onclick="if(confirm('是否删除!')){location.href='KS.qiandao.asp?Action=del&id=<%=SQL(0,i)%>'};">删除</a></td>
  </tr>
  <%Next%>
  <%  
  Response.Write "<tr><td colspan=9 align=right class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
  Response.Write "</td></tr>"
End Sub

Sub qiandao_del()
	if KS.ChkClng(KS.G("id"))<>0 then
		conn.execute("update KS_qiandao set Content='未签到',Status=1 where id="& KS.ChkClng(KS.G("id")) )	
		KS.echo "<script>alert('恭喜,删除成功!');location.href='?page=" & KS.G("Page") & "';</script>"
	end if
End Sub

Sub qiandao_delall()
	conn.execute("delete from KS_qiandao")
	conn.execute("update KS_user set qiandao=0,qiandao_m=0,qiandao_xqco=''")
	KS.echo "<script>alert('恭喜,删除成功!');location.href='KS.qiandao.asp';</script>"
End Sub

				
End Class
%> 

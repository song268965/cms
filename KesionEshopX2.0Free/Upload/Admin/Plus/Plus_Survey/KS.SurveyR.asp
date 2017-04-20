<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"--> 
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS X1.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_SurverCls
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_SurverCls
        Private KS,KSCls,I,TypeFlag,ItemStr
		Private MaxPerPage,CurrentPage,TotalPut,ID,RS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  MaxPerPage=20
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub

		
		Public Sub Kesion()
		  With KS
		   If Not KS.ReturnPowerResult(0, "Survey0001") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 Exit Sub
		   End If
		   CurrentPage=KS.ChkClng(request("Page"))
		   if CurrentPage<=0 then CurrentPage=1
		   TypeFlag=KS.ChkClng(KS.S("TypeFlag"))
		    ItemStr="项目"
		   
		    .echo "<!DOCTYPE html><html>"
			.echo"<title>项目设置</title>"
			.echo"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo"<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
			.echo"<script src=""../../../ks_inc/jQuery.js"" language=""JavaScript""></script>"
			.echo"<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo"</head>"
			.echo"<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.echo"<ul id='menu_top'>"
			if KS.G("Action")="EditST" then
			.echo"<li class='parent' onclick='history.go(-1)'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>返回</span></li>"
			end if
             if KS.G("Action")<>"EditST" then
				 If KS.G("Action")="" Then
					.echo"<li class='parent' disabled"
				 Else
					.echo"<li class='parent'"
				 End If
				 if KS.G("Action")<>"Surveyshow" then
					.echo" onclick='location.href=""KS.SurveyR.asp?typeflag=" & typeflag & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon mainer'></i>管理首页</span></li>"
				 else
				 	.echo" ></li>"
				 end if
			end if	
				.echo"</ul>"
				.echo "<div class='pageCont2'>"
		  Select Case KS.G("Action")
		   Case "SurveyR" Call SurveyRMain() 
		   Case "Surveyshow" Call Surveyshow()
		   Case Else Call Main()
		  End Select
		  End With
		End Sub
 
		Sub Main()
		   With KS
			.echo"<script>"
			.echo"$(document).ready(function(){"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button1').attr('disabled',true);"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button2').attr('disabled',true);"
			.echo"});</script>"
			.echo "<div class='tabTitle'>问卷结果查看</div>"
			.echo("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_Survey Order By ID",conn,1,1
		    .echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo"  <td width='50' align=center>ID</td><td align=center width='150'>项目名称</td><td width='150' align=center>总投票数</td><td align=center>↓操作</td>"
			.echo"</tr>"
			If RS.Eof And RS.Bof Then
			 .echo "<tr><td class='splittd' align='center' height='40' colspan=10>还没有添加项目！</td></tr>"
			Else
			            totalPut = RS.RecordCount
						If CurrentPage < 1 Then	CurrentPage = 1
			            If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
						End If
			
			            dim i:i=0
					  Do While Not RS.Eof 
						.echo"<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
						.echo"<td align=center class='splittd' style='height:35px;'>" & RS("ID")&"</td>"
						.echo"<td class='splittd' style='height:35px;' align=center><a href='?action=SurveyR&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=GoSave&OpStr=问卷结果查看 >> <font color=red>结果查看</font>';"">" & RS("ProjectName") 
						.echo"</a></td>"
						.echo"<td align=center class='splittd' style='height:35px;'>" 
						.echo " &nbsp;<font color='#339900'> "&Conn.Execute("SELECT COUNT(ID) FROM KS_SurveyResult where SurveyID = "&RS("ID"))(0) &"</font>"
						.echo"</td>"
						.echo"<td align=center class='splittd' style='height:35px;'>"
						.echo"<a href='?action=SurveyR&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='post.asp?ButtonSymbol=GoSave&OpStr=问卷结果查看 >> <font color=red>结果查看</font>';"">结果查看</a>"
						.echo"</td></tr>"
						i=i+1
						if i>=maxperpage then exit do
						RS.MoveNext 
					  Loop
			end if
		    .echo"</table>"
			.echo  KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo"</div>"
			.echo"</div>"
		   RS.Close:Set RS=Nothing
		    .echo"</body>"
			.echo"</html>"
		  End With
		End Sub
		
		Sub SurveyRMain()
			With KS
					dim SurveyName:SurveyName=Conn.Execute("select top 1 ProjectName from KS_Survey where ID="&KS.ChkClng(KS.G("ID")))(0)
					.echo"<script>"
					.echo"$(document).ready(function(){"
					.echo"$(parent.frames['BottomFrame'].document).find('#Button1').attr('disabled',true);"
					.echo"$(parent.frames['BottomFrame'].document).find('#Button2').attr('disabled',true);"
					.echo"});</script>"
					.echo"<script src=""../../../ks_inc/flotr2.min.js"" type=""text/javascript""></script>" & vbcrlf
					.echo"<script>"&vbcrlf
					.echo"function Surveyshow(id){"&vbcrlf
					.echo"	top.openWin('查看投票详情','plus/plus_survey/KS.SurveyR.asp?action=Surveyshow&surveystid='+id,false,850,500);"&vbcrlf
					.echo"}"&vbcrlf
					.echo"</script>"&vbcrlf
					.echo("<div style="" overflow: auto; width:100%"" align=""center"">")
					.echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
					.echo"<tr height='25' class='sort'>"
					.echo"  <td width='98%' align=center>"& SurveyName &"</td><td align=center width='1'></td><td width='1' align=center></td><td align=center width='1'></td>"
					.echo"</tr>"
					.echo"</table>"
					.echo "</div>"
					Set Rs = Server.CreateObject("adodb.recordset")
					Rs.Open "select * from KS_SurveyST where SurveyID="& KS.ChkClng(KS.G("ID")) &" ORDER BY SurveyOrder" , Conn, 1, 1
					dim n,nstr
					Do While Not rs.Eof
			        n=N+1
						.echo"<div style='border-bottom:1px dashed #CCCCCC; margin-top:10px;'>" &vbcrlf
						.echo"<div style='float:left;width:400px;'>"
						.echo "<li  style='height:35px;font-size:14px;font-weight:bold' align=left>&nbsp;" & n & "、["& rs("SurveySTName") & "] &nbsp;</li>"
						Call SurveyBox(rs("ID"),"lx1")
						.echo"</div>"
						.echo"<div style='float:right;margin-right:30px;' >"
						%>
						
						 <div id="container<%=n%>" style="margin:0 auto;width:350px;height:200px"></div>
	
							<script type="text/javascript">
						   (function basic_pie(container) {
						  var graph;
						  graph = Flotr.draw(container, [
							<%
							Call SurveyBox(rs("ID"),"lx2")
							%>
						  ],
						   {
							HtmlText : true,
							grid : {
							  verticalLines : false,
							  horizontalLines : false
							},
							xaxis : { showLabels : false },
							yaxis : { showLabels : false },
							pie : {
							  show : true, 
							  explode : 6
							},
							mouse : { track : true },
							legend : {
							  position : 'se',
							  backgroundColor : '#D2E8FF'
							}
						  });
						})(document.getElementById("container<%=n%>"));
						</script>
						
						<%
						
						.echo"</div><div style='clear:both;height:0px; overflow:hidden;'></div>"
						.echo "</div>"&vbcrlf
					rs.MoveNext 
					loop
					Rs.Close
					Set Rs = Nothing
					
					
					.echo"</body>"
					.echo"</html>"
			End With
		end sub
		
		Sub SurveyBox(ID,s_lx)
		   on error resume next
			With KS
			dim I_Rs
			Set I_Rs = Server.CreateObject("adodb.recordset")
				I_Rs.Open "select * from KS_SurveyItem where SurveySTID="& KS.ChkClng(ID) &" ORDER BY SurveyItemOrder" , Conn, 1, 1
				dim n,nstr:n=0
				dim tnum:tnum=conn.execute("select count(1) from KS_SurveyResult where SurveySTID=" & KS.ChkClng(ID))(0)          
				if s_lx="lx1" then
				 .echo "<table>"
				end if
				Do While Not I_Rs.Eof
					nstr=chr(65+n):n=n+1
					if s_lx="lx1" then
					    dim per:per=round(KS.ChkClng(I_Rs("ItemNum"))/tnum*100,2)
						.echo "<tr><td style='width:250px;height:35px;'>&nbsp;&nbsp;"& nstr &"、"& I_Rs("SurveyItemName") & "(投票数:" & KS.ChkClng(I_Rs("ItemNum")) & ")</td><td><img src='../../../images/Default/bar.gif' width='"& per &"' height='15' align='absmiddle' /> " & Per & "%"
						if I_Rs("SurveyItemType")=1 then
							.echo "<input name=""详细查看"" class='button' onClick=""Surveyshow("&I_Rs("id") &");"" value=""详细查看"" type=""button"">"
						end if
						.echo "</td></tr>"
					else
					
					.echo "{ data : [[0, " &  KS.ChkClng(I_Rs("ItemNum")) &"]], label : '"& nstr &"、"&I_Rs("SurveyItemName") &"' },"
					end if		
				I_Rs.MoveNext 
				loop
				if s_lx="lx1" then
				 .echo "</table>"
				end if
			I_Rs.Close
			Set I_Rs = Nothing
			End With
		end sub
		
		Sub Surveyshow()
			dim Param:Param=" where SurveyItemID="& KS.ChkClng(KS.G("surveystid"))
			dim CurrPage:CurrPage=KS.ChkClng(KS.G("page"))
			if CurrPage<=0 then CurrPage=1	
			if KS.ChkClng(KS.G("surveystid"))<>0 then
					dim rs
					Set Rs = Server.CreateObject("adodb.recordset")
					Rs.Open "select * from KS_SurveyResult where SurveyItemID="& KS.ChkClng(KS.G("surveystid")) &" ORDER BY ID DESC" , Conn, 1, 1
					ks.echo "<style>"
					ks.echo "*{margin:0;padding:0;word-wrap:break-word;}"
					ks.echo "body{font:12px/1.75 ""宋体"", arial, sans-serif,'DejaVu Sans','Lucida Grande',Tahoma,'Hiragino Sans GB',STHeiti,SimSun,sans-serif;color:#444;}"
					ks.echo "a{color:#333;text-decoration:none;}"
					ks.echo "a:hover{text-decoration:underline;}"
					ks.echo "a img{border:none;} "
					ks.echo "div,ul,li,p,form{padding: 0px; margin: 0px;list-style-type: none;}"
					ks.echo ".SurveyR_co{background:#F5F5F5;font-size:14px; border:1px solid #CCCCCC;width:330px;text-indent:10px;overflow:hidden;margin-top:10px;text-align:left;padding:5px;}"
					ks.echo "</style>"
					ks.echo"<table border='0' cellpadding='0' cellspacing='0' width='100%'  align='center'>"
					ks.echo"<tr class='sort'>"
					ks.echo "<td  width='50%'>内容</td><td width='18%'>用户</td><td width='30%'>时间</td>"
					ks.echo"</tr>"
					If RS.Eof And RS.Bof Then
				    Else
						TotalPut = Conn.Execute("select Count(1) from KS_SurveyResult "& Param )(0)
						If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrPage - 1) * MaxPerPage
						End If
						dim i:i=0
						Do While Not rs.Eof
							ks.echo"<tr class='sort'>"
							ks.echo "<td ><div class=""SurveyR_co"">"&  Replace(rs("Content"),Chr(13),"") & "</div> &nbsp;</td>"
							ks.echo "<td style=""margin-top:10px;font-size:12px;""> "& rs("username") & " </td>"
							ks.echo "<td style=""margin-top:10px;font-size:12px;""> "& rs("AddDate") & "</td>"
							ks.echo"</tr>"
							I=i+1
							if i>=MaxPerPage then exit do
						rs.MoveNext 
						loop
						ks.echo "</table>"
					end if
					ks.echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
					ks.echo "</body></html>"
					Rs.Close
					Set Rs = Nothing
			end if
			Response.end()
		End sub
		
		
	
		
		
End Class
%> 

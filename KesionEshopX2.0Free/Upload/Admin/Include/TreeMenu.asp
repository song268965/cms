<%@language=vbscript codepage="65001" %>
<%
Option Explicit
Response.buffer = True
Server.ScriptTimeout=9999999
%>
<!--#include file="../../conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KS:Set KS=New PublicCls
Dim strInstallDir,ComeUrl
If Not KS.ReturnPowerResult(0, "KMSL10009") Then          
	'Call KS.ReturnErr(1, "")
	Response.End
End If

ComeUrl=Request.ServerVariables("http_referer")
strInstallDir=KS.Setting(3)

Dim ChannelUrl, UseCreateHTML,  ListFileType, FileExt_List

Dim hf, strTopMenu, pNum, pNum2, OpenTyKS_Class, strMenuJS
Dim FSO
Set FSO = KS.InitialObject(KS.Setting(99))
Response.Write "<!DOCTYPE html><html><head><title>顶部栏目菜单管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<link href='../include/Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "<script src='../../ks_inc/jquery.js'></script>" &vbcrlf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<div class='tabTitle mt20'>生成树型菜单</div>"

Response.Write "<div class='pageCont2'>"

Dim Action:Action=KS.G("Action")
If Action = "Create" Then
    Call Create_RootClass_Menu
Else
    Call Create_Tree()
End If
Response.Write "</body></html>" & vbCrLf
Response.Write "</div>"
Sub Create_Tree()
    Response.Write "<form method='POST' action='?Action=Create' id='myform' name='myform'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
	Response.Write "  <tr class='title' style='display:none;'>"
    Response.Write "    <td height='22' colspan='6'><strong>树型菜单参数设置</strong> </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>选择频道：</strong></td>"
    Response.Write "    <td>"
    Response.Write ReturnAllChannel()
    Response.Write "    </td>"
	Response.Write " </tr>"
	Response.Write " <tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成样式：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <select name='fsostyle' class='textbox' onchange=""if (this.value==2) {$('#s2').show();$('#s1').hide();}else {$('#s1').show();$('#s2').hide();}"">"
	Response.Write "        <option value=1>样式一(ztree插件）</option>"
	Response.Write "        <option value=2>样式二</option>"
	Response.Write "      </select>"
    Response.Write "    </td>"
	Response.Write "</tr>"
	Response.Write "<tbody id='s1'>"
	Response.Write " <tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>默认是否展开：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <input type='radio' name='isopen' value='1' checked>第<input type='text' name='opencol' class='textbox' style='width:40px;text-align:center' value='2'/>个大类展开<br/>"
    Response.Write "      <input type='radio' name='isopen' value='2'>全部展开<br/>"
    Response.Write "      <input type='radio' name='isopen' value='3'>全部关闭<br/>"
    Response.Write "    </td>"
	Response.Write "</tr>"
	Response.Write "</tbody>"
	Response.Write "<tbody id='s2' style='display:none'>"
	Response.Write " <tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成列数：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <input type='text' class='textbox' name='col' value='2' size=""6"">列"
    Response.Write "    </td>"
	Response.Write "</tr>"
	Response.Write "</tbody>"
	Response.Write "<tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成文件名：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <input name='JsFileName' class='textbox' type='text' id='JsFileName' value='Tree.js' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
	Response.Write "</table>"
    Response.Write "<br><div style='text-align:center'><input type='submit' name='Submit' value=' 生成树型导航 ' class='button'></div>"
	Response.Write "</form>"
End Sub
Sub Create_RootClass_Menu()
    If KS.ChkCLng(KS.S("fsostyle"))=1 Then
    strTopMenu = TreeList 
	Else
	strTopMenu = HtreeList
	End If
    Call KS.WriteTOFile(KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName"), strTopMenu)
	Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='ctable'>"
    Response.Write "  <tr class='sort'>"
    Response.Write "    <td height='22' align='center'><strong> 生 成 树 形 导 航 菜 单 </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td>"
	Response.Write "<p align='center'><font color=red><b>恭喜您！树形导航菜单成功生成,请按以下提示完成最好操作。</b></font></p>"
	If KS.ChkCLng(KS.S("fsostyle"))=1 Then
	%>
	
	<br/><strong>第一步：在需要显示的页面模板的&lt;HEAD&gt;与&lt;/HEAD&gt;之间放入以下代码:</strong><br/>
	&lt;script type="text/javascript" src="<%=KS.Setting(3)%>ks_inc/ztree/jquery-1.4.4.min.js">&lt;/script><br/>
		&lt;link rel="stylesheet" href="<%=KS.Setting(3)%>ks_inc/ztree/css/zTreeStyle/zTreeStyle.css" type="text/css"><br/>
	&lt;script type="text/javascript" src="<%=KS.Setting(3)%>ks_inc/ztree/jquery.ztree.core-3.5.js">&lt;/script><br/>
	&lt;script type="text/javascript" src="<%=KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName")%>">&lt;/script><br/>
<br/>
<strong>第二步：在需要显示的页面模板的&lt;body&gt;与&lt;/body&gt;之间放入以下代码:</strong><br/>
&lt;ul id="ztree" class="ztree">&lt;/ul>

<br/><br/><strong>以下是本次生成的效果：</strong><br/>
	<link rel="stylesheet" href="<%=KS.Setting(3)%>ks_inc/ztree/css/zTreeStyle/zTreeStyle.css" type="text/css">
	<script type="text/javascript" src="<%=KS.Setting(3)%>ks_inc/ztree/jquery-1.4.4.min.js"></script>
	<script type="text/javascript" src="<%=KS.Setting(3)%>ks_inc/ztree/jquery.ztree.core-3.5.js"></script>
	<script type="text/javascript" src="<%=KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName")%>"></script>

<ul id="ztree" class="ztree"></ul>


	<%
	Else
    Response.Write "<p><b>将以下代码复制到在模板里要显示的地方。</b></p>"
	Response.Write "<input class='textbox' name='s2' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName") & "&quot;&gt;&lt;/script&gt;' size='80'>&nbsp;<input class=""button"" onClick=""jm_cc('s2')"" type=""button"" value=""复制到剪贴板"" name=""button1"">"
	End IF
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
 %>
 <script>
function jm_cc(ob)
{
	var obj=MM_findObj(ob); 
	if (obj) 
	{
		obj.select();js=obj.createTextRange();js.execCommand("Copy");}
		alert('复制成功，粘贴到你要调用的模板里即可!');
	}
	function MM_findObj(n, d) { //v4.0
  var p,i,x;
  if(!d) d=document;
  if((p=n.indexOf("?"))>0&&parent.frames.length)
   {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
   }
  if(!(x=d[n])&&d.all) x=d.all[n];
  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
  </script>
 <%
End Sub


    
Function EncodeJS(str)
    EncodeJS = Replace(Replace(Replace(Replace(Replace(str, Chr(10), ""), "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "")
End Function


'取得网站的所有频道及其子栏目
Function ReturnAllChannel()
	  Dim RS:Set RS=KS.InitialObject("ADODB.Recordset")
	  Dim SQL,K,ChannelStr:ChannelStr = ""
	   ChannelStr = "<select class='textbox' name=""ChannelID"" style=""width:200;border-style: solid; border-width: 1""><option value='0'>---不指定栏目---</option>"
	   RS.Open "Select channelid,channelname From [KS_Channel] Where ChannelStatus=1", Conn, 1, 1
	   If RS.EOF And RS.BOF Then
		  RS.Close:Set RS = Nothing:Exit Function
	   Else
	     SQL=RS.GetRows(-1):rs.close:set rs=nothing
	   End iF
		
	    For K=0 To ubound(sql,2)
		   ChannelStr = ChannelStr & "<option value=" & sql(0,k) & ">" & sql(1,k) & "</option>"
		Next 
		ChannelStr = ChannelStr & "<optgroup  label=""-----指定到具体的栏目(以下列出了整站的导航树)----"">"  
	   For K=0 To Ubound(sql,2)
	        ChannelStr=ChannelStr & KS.LoadClassOption(sql(0,k),false)
	    Next
	   ReturnAllChannel = ChannelStr &"</select>"
End Function
	
	Function TreeList()
				Dim RS,TreeStr,ID,i,ii,Param,ChannelID
				ChannelID=KS.S("ChannelID")
				 If Len(Channelid)>4 Then
					 Param=" and a.tn='" & ChannelID & "'"
				 Else
				   If ChannelID<>"0" Then  Param=" and B.ChannelID=" & KS.ChkCLng(KS.S("ChannelID"))
				 End If
				TreeStr="var setting = {" &vbcrlf
				TreeStr=TreeStr &"data: {" &vbcrlf
				TreeStr=TreeStr &"simpleData: {" &vbcrlf
				TreeStr=TreeStr &"	enable: true" &vbcrlf
				TreeStr=TreeStr &"}" &vbcrlf
			    TreeStr=TreeStr &"}" &vbcrlf
		        TreeStr=TreeStr &"};" & vbcrlf
				TreeStr=TreeStr&"var zNodes =[" &vbcrlf
				Set RS=KS.InitialObject("ADODB.Recordset")
				RS.Open ("select ID,FolderName,tn,child from KS_Class A,KS_Channel B Where A.ChannelID=B.ChannelID And B.ChannelStatus=1 "  & Param & " Order BY root,folderorder"), Conn, 1, 1
				i=0
				ii=0
				Do While Not RS.EOF
				 if rs("tn")="0" then ii=ii+1
				 TreeStr=TreeStr&"{ id:'" & rs("id") &"', pId:'" & rs("tn")&"', name:'" & rs("foldername")&"', url:'" & KS.GetFolderPath(rs("id")) &"'"
				 if ks.s("isopen")="2" then
				   TreeStr=TreeStr&",open:true"
				 elseif ks.s("isopen")="1" then
				  if ii=ks.chkclng(request("opencol")) then TreeStr=TreeStr&",open:true"
				 end if
				  if rs("child")=0 and rs("tn")="0" then
				 TreeStr=TreeStr&",isParent:true"
				  end if
				  TreeStr=TreeStr&"},"&vbcrlf
				  i=I+1
				  rs.movenext
				 loop
				 rs.close
				 set rs=nothing
				
		      TreeStr=TreeStr&"];"&vbcrlf

		TreeStr=TreeStr&"$(document).ready(function(){"&vbcrlf
		TreeStr=TreeStr&"	$.fn.zTree.init($(""#ztree""), setting, zNodes);"&vbcrlf
		TreeStr=TreeStr&"});"&vbcrlf
				
			 TreeList=TreeStr
	End Function
	
	
	
	
	'横向
	Function HtreeList()
	   Dim RS,TreeStr,ID,i,Param,ChannelID
	   ChannelID=KS.S("ChannelID")
	   If Len(Channelid)>4 Then
	     Param=" and a.tn='" & ChannelID & "'"
	   Else
				If KS.S("ChannelID")<>"0" Then  Param="  and B.ChannelID=" & KS.ChkCLng(ChannelID)
				IF KS.S("ChannelID")<>"8" Then Param=Param &"  And tj=1" Else Param=Param & " and tj=2"
	   End If
				Set  RS=KS.InitialObject("ADODB.Recordset")
				RS.Open ("select ID from KS_Class A,KS_Channel B Where A.ChannelID=B.ChannelID And B.ChannelStatus=1 "  & Param & " Order BY root,folderorder"), Conn, 1, 1
				TreeStr=TreeStr & "document.writeln('<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"">');" & vbcrlf
				Do While Not RS.EOF
				  TreeStr=TreeStr & "document.writeln('<tr>');" & vbcrlf
				  For I=1 To KS.ChkClng(KS.G("Col"))
				   TreeStr = TreeStr & "document.writeln('<td valign=""top"" width=""" & 100 / KS.ChkCLng(KS.G("Col")) & "%"">');" & vbcrlf
				   TreeStr = TreeStr & "document.writeln('<div class=""classtitle"" style=""font-weight:bold""><img src=""" & KS.Setting(3) & "images/default/arrow_r.gif"" align=""absmiddle"">&nbsp;" & KS.GetClassNP(rs(0))& "</div>');" & vbnewline 
				   TreeStr = TreeStr & SubList(RS(0))
				   TreeStr = TreeStr & "document.writeln('</td>');" & vbcrlf
				   RS.MoveNext
				   If RS.EOF Then Exit For
				  Next
				   TreeStr = TreeStr & "document.writeln('</tr>');" & vbcrlf
				  if rs.eof then exit do
				Loop
				TreeStr =TreeStr & "document.writeln('</table>');" & vbcrlf
				RS.Close:Set RS=Nothing
		HtreeList=TreeStr
	End Function	
	
	Function SubList(ParentID)
	  Dim RS:Set RS=Conn.Execute("select id from ks_class where tn='" & ParentID & "' order by root,folderorder")
	  Dim SQL,I
	  If Not RS.Eof Then
	     SQL=RS.GetRows(-1)
		 SubList="document.writeln('<div class=""list"">"
		 For I=0 To Ubound(SQL,2)
		   SubList=SubList & KS.GetClassNP(SQL(0,I)) & "&nbsp;"
		   If I <> Ubound(SQL,2) Then SubList=SubList & "<img src=""" & KS.Setting(3) & "images/nl.gif"" align=""absmiddle"">&nbsp;"
		   
		 Next
		 SubList=SubList & "</div>');"& vbcrlf
	  End IF
	  RS.Close:Set RS=Nothing
	End Function 
%>

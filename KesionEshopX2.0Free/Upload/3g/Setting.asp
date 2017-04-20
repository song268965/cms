<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../plus/Session.asp"-->
<!--#include file="../Plus/md5.asp"-->
<%
Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing

Dim KS:Set KS=New PublicCls

If Not KS.ReturnPowerResult(0, "KSO10003") Then
   Response.Write ("<script>parent.frames['BottomFrame'].location.href='javascript:history.back();';</script>")
   Call KS.ReturnErr(1, "")
   Response.End()
End If

if lcase(request("action")&"")="template" then
 Call template()
else
 Call SetSystem()
end if

Call CloseConn()
Set KS=Nothing


Sub SetSystem()
    on error resume next
	Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
	If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
	Dim SqlStr, RS
	SqlStr = "select WapSetting from KS_Config"
	Set RS = Server.CreateObject("ADODB.recordset")
	RS.Open SqlStr, Conn, 1, 3
	Dim WapSetting:WapSetting=Split(RS(0)&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
	If KS.G("Flag") = "Edit" Then
	   Dim N					
	   Dim WebSetting
	   For N=0 To 50
	       WebSetting=WebSetting & Replace(KS.G("WapSetting(" & N &")"),"^%^","") & "^%^"
	   Next
	   RS("WapSetting")=WebSetting
	   RS.Update
	   Call KS.DelCahe(KS.SiteSn & "_Config")
	   Call KS.DelCahe(KS.SiteSn & "_Date")
	   Response.Write "<script>top.$.dialog.alert('手机版基本参数修改成功！',function(){location.href='../" & WapSetting(4) &"/setting.asp';});</script>"				
	End If
	%>
    <!DOCTYPE html><html>
    <title>手机版基本参数设置</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<script src="<%=KS.Setting(3)%>ks_inc/jquery.js" language="JavaScript"></script>
    <script src="<%=KS.Setting(3)%>ks_inc/Common.js" language="JavaScript"></script>
	<script src="<%=KS.Setting(3)%><%=KS.Setting(89)%>Images/pannel/tabpane.js" language="JavaScript"></script>
    <link href="<%=KS.Setting(3)%><%=KS.Setting(89)%>Images/pannel/tabpane.CSS" rel="stylesheet" type="text/css">
    <link href="<%=KS.Setting(3)%><%=KS.Setting(89)%>Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	<style type="text/css">
	<!--
	.STYLE1 {color: #FF0000}
	.STYLE2 {color: #FF6600}
	-->
    </style>
    </head>

    <body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
    
    <div class="tabTitle mt20">手机版基本参数设置</div>
    <form name="myform" id="myform" method="post" Action="setting.asp" >
    <div class="tab-page" id="3gconfig">
	<script type=text/javascript>var tabPane1 = new WebFXTabPane( document.getElementById( "3gconfig" ), 1 )</script>
    
    <div class=tab-page id=site-page>
    <H2 class=tab>基本参数</H2>
	<script type=text/javascript>
	tabPane1.addTabPage( document.getElementById( "site-page" ) );
    </script>
    <dl class="dtable">
    <input type="hidden" value="Edit" name="Flag">
    
    <dd ><div>是否开启手机版：</div>
    <input  name="WapSetting(0)" type="radio" onClick="$('#sss').show()" value="1" <%If WapSetting(0)="1" Then Response.Write" Checked"%>>开启
    <input name="WapSetting(0)" type="radio" onClick="$('#sss').hide()" value="0" <%If WapSetting(0)="0" Then Response.Write" Checked"%>>关闭
	<span>Tips:开启手机版本，如果使用手机浏览本网站主域名，将自动进入手机专用版本。</span>
    </dd>

    <span id="sss"<%If WapSetting(0)="0" Then Response.Write" style=""display:none"""%>>
    <dd><div>手机版网站名称：</div>
	<input name="WapSetting(3)" class="textbox" type="text" value="<%=WapSetting(3)%>" size="50"> <span>可以在手机版模板里用{$Get3GSiteName}调用</span></td>
    </dd>
    
    <dd><div>安装目录：</div>
	<input name="WapSetting(4)" class="textbox"  type="text" value="<%=WapSetting(4)%>" size="50"><span>* 手机版插件安装的目录，如“3G”,后面不能带“/”,3G模板里可以用标签{$Get3GInstallDir}调用此名称</span>
    </dd>    
    <dd><div>绑定二级域名：</div><input name="WapSetting(1)" class="textbox"  type="text" value="<%=WapSetting(1)%>" size="50"><span>如:3g.kesion.com等,不要带“http://”,如果不绑定请留空,否则可能导致页面路径出错,支持独立域名或二级域名的绑定</span>
    </dd>    

    <dd><div>网站Logo地址：</div><input name="WapSetting(2)" class="textbox"  type="text" value="<%=WapSetting(2)%>" size="50"> <span>手机版模板里可以用标签{$Get3GLogo}调用路径</span></dd>

    <dd><div>底部版权信息：</div>
	<textarea name="WapSetting(5)" cols="60" rows="4" class="textbox"><%=WapSetting(5)%></textarea><span> 可以在手机版模板里用{$Get3GCopyRight}调用</span>
	</dd>
	</span>
    
    </dl>
	</div>
	<div class=tab-page id=seo-page>
		<H2 class=tab>SEO选项</H2>
		<script type=text/javascript>
		tabPane1.addTabPage( document.getElementById( "seo-page" ) );
		</script>
		<dl class="dtable">
		<dd><div>手机版网站标题：</div>
		<textarea name="WapSetting(6)" cols="60" rows="2"><%=WapSetting(6)%></textarea><span> 可以在手机版模板里用{$Get3GSiteTitle}调用</span>
		</dd>
		<dd><div>手机版网站META关键词：</div>
		<textarea name="WapSetting(7)" cols="60" rows="4"><%=WapSetting(7)%></textarea><span> 可以在手机版模板里用{$Get3GMetaKeyWord}调用</span>
		</dd>
		<dd><div>手机版网站META网页描述：</div>
		<textarea name="WapSetting(8)" cols="60" rows="4"><%=WapSetting(8)%></textarea><span> 可以在手机版模板里用{$Get3GMetaDescript}调用</span>
		</dd>
		</dl>
	</div>
	
		<div class=tab-page id=seo-page>
		<H2 class=tab>生成选项</H2>
		<script type=text/javascript>
		tabPane1.addTabPage( document.getElementById( "seo-page" ) );
		</script>
		<dl class="dtable">
		<dd><div>手机版生成的HTML总目录：</div>
		<input type="text" name="WapSetting(10)" value="<%=WapSetting(10)%>" class="textbox"> <span class="tips">如:/HTML/3G/</span>
		</dd>
		<dd><div>手机版生成的扩展名：</div>
		 <select name="WapSetting(9)" class="textbox">
		   <option value=".html"<%If WapSetting(9)=".html" then response.write " selected"%>>.html</option>
		   <option value=".htm"<%If WapSetting(9)=".htm" then response.write " selected"%>>.htm</option>
		   <option value=".shtml"<%If WapSetting(9)=".shtml" then response.write " selected"%>>.shtml</option>
		   <option value=".shtm"<%If WapSetting(9)=".shtm" then response.write " selected"%>>.shtm</option>
		   <option value=".asp"<%If WapSetting(9)=".asp" then response.write " selected"%>>.asp</option>
		 </select>
		</dd>
		
		</dl>
	</div>

	
	</div>
    
    </body>
    </html>
	<script Language="javascript">
	<!--
	function CheckForm(){$('#myform').submit(); }
	//-->
    </script>
    <%
	RS.Close:Set RS = Nothing
End Sub


Sub Template()
If Not KS.ReturnPowerResult(0, "KSO10003") Then
			   Response.Write ("<script>parent.frames['BottomFrame'].location.href='javascript:history.back();';</script>")
			   Call KS.ReturnErr(1, "")
			   Response.End()
			End If
			if request("flag")="" then%>
			<!DOCTYPE html><html>
			<%end if
			Response.Write "<head>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write "<title>模板管理</title>"
			Response.Write "<script src='../ks_inc/jquery.js'></script>"
			Response.Write "<link href=""../" & KS.Setting(89) & "Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "</head>"
			Select Case KS.G("flag")
			    Case "Del" Call LabelDel()
				Case "AddNew","EditLabel"	Call AddLabel()
				Case "AddSave"	Call AddSave()
				Case "EditSave"	Call EditSave()
				Case Else
				Call LabelList()
			End Select
	End Sub
	Sub LabelList()
		    Dim MaxPerPage
			MaxPerPage =22
			
			Response.Write "<script>"
			Response.Write "$(document).ready(function(){"
			Response.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
			Response.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
			Response.Write "})</script>"
			%>

            </head>
            <body topmargin="0" leftmargin="0">
			
<ul id='menu_top'><li class='parent' onClick="location.href='?Action=Template&flag=AddNew';"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='../<%=KS.Setting(89)%>images/ico/iconfont-tianjia.png' border='0' align='absmiddle'>添加页面</span></li></ul>			
			<div class="pageCont2">
            <div class="tabTitle">手机版自定义页面</div>
            <div style="height:94%; overflow: auto; width:100%" align="center">
			<%
			Dim RS,SQL
			Set RS = Server.CreateObject("ADODB.RecordSet")
			SQL = "SELECT * FROM KS_WapTemplate ORDER BY AddDate Desc" 
			RS.Open SQL, Conn, 1, 1
			%>
            <table width='100%' border='0' cellspacing='0' cellpadding='0'>
            <tr>
            <td width='110' class="sort" align='center'><font color="#990000">I D</font></td>
            <td class='sort' align='center'>页面名称</td>
            <td class='sort' align='center'>超链接地址</td>
            <td class='sort' align='center'>修改时间</td>
            <td class='sort' align='center'>操作管理</td>
            </tr>
			<%
			If RS.Eof And RS.Bof Then
			%><tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'">
            <td colspan=5 class='splittd' style="height:50px"><div align="center">您没有添加自定义页面!</div></td></tr>
			<%
			Else
			Dim TotalPut,I
			TotalPut= RS.RecordCount
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			   RS.Move (CurrentPage - 1) * MaxPerPage
			End If	
					 
			Do While Not RS.EOF
			%>
            <tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'">
            <td height="19" align=center class='splittd'><div align="center"><span style="cursor:default"><%=RS("ID")%></span></div></td>
            <td class='splittd' style="text-align:left">&nbsp;<%=Trim(RS("TemplateName"))%></td>
            <td class='splittd' align='center'><a href="../<%=KS.Wsetting(4)%>/diy.asp?id=<%=RS("ID")%>" target="_blank">diy.asp?ID=<%=RS("ID")%></a></td>
            <td class='splittd' align='center'><%=RS("AddDate")%></td>
            <td class='splittd' align='center'><a href='setting.asp?Action=template&flag=EditLabel&TemplateID=<%=RS("ID")%>' onClick="parent.frames['BottomFrame'].location.href='../../<%=KS.Setting(89)%>Post.Asp?OpStr=手机版自定义页面管理中心 >> 3G自定义页面编辑&ButtonSymbol=GoSave'" class='setA'>编辑</a>|<a href="setting.asp?Action=template&flag=Del&TemplateID=<%=RS("ID")%>" onClick="return(confirm('此操作不可逆，确定删除吗？'))" class="setA">删除</a>
            </td>
            </tr>
			<%
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
			Loop	 
			End If
			%>
            </table>
            <table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>
            <tr>
            <td align='right'>
			<%
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			%></td></tr></table>
            <div style="text-align:center;color:#003300">-----------------------------------------------------------------------------------------------------------</div>
            <div style="height:30px;text-align:center">Copyright (c) 2006-<%=year(now)%> <a href="http://www.kesion.com/" target="_blank"><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>
            			<br/><br/>
</div>
</div>
            </body>
            </html>
			<%
        End Sub

		
		'删除模板
		Sub LabelDel()
		    Dim TemplateID:TemplateID=KS.G("TemplateID")
			Conn.Execute("Delete from KS_WapTemplate where ID='" & TemplateID & "'")
			Response.write "<script>window.alert('删除成功');window.location='" & request.ServerVariables("HTTP_REFERER") &"';</script>"
		End Sub
		
		sub checktemplateid()
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select id from KS_WapTemplate",conn,1,3
		  do while not rs.eof
		    if len(rs(0))>6 then
			  rs(0)=gettemplateid(right(rs(0),2))
			  rs.update
			end if
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		end sub
		function gettemplateid(id)
		  dim tempid:tempid=ks.chkclng(id)
		  do while true
		     if conn.execute("select top 1 id from ks_waptemplate where id='" & tempid & "'").eof then
			   exit do
			 else
			   tempid=tempid+1
			 end if
		  loop
		  gettemplateid=tempid
		end function
		
		
		Sub AddSave()
		    checktemplateid
		    Dim RS,RSCheck,TemplateID,TemplateName,TemplateContent
			TemplateName = Replace(Replace(Trim(Request.Form("TemplateName")), """", ""), "'", "")
			TemplateContent = Trim(Request.Form("TemplateContent"))
			If TemplateName="" Then
			   Response.write "<script>window.alert('标题没有填写..');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			If TemplateContent="" Then
			   Response.write "<script>window.alert('操作失败...');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			Set RS = Server.CreateObject("ADODB.RecordSet")
			RS.Open "Select * From [KS_WapTemplate] Where (ID is Null)",Conn,1,3     
			RS.Addnew
			   Set RSCheck = Conn.Execute("Select max(ID) from [KS_WapTemplate] Where ID='" & TemplateID & "'")
			   If RSCheck.EOF And RSCheck.BOF Then
			      TemplateID=KS.ChkClng(RSCheck(0))+1
			   Else
			      TemplateID=1
			   End If
			   TemplateID=gettemplateid(TemplateID)
			RS("ID")=TemplateID
			RS("TemplateName")=TemplateName
			RS("TemplateContent")=TemplateContent
			RS("AddDate")=Now
			RS.Update
			RS.Close:set RS=Nothing
			Response.Write ("<script>if (confirm('成功提示:\n\n添加自定义页面成功,继续添加标签吗?')){location.href='setting.asp?Action=template&flag=AddNew';}else{parent.frames['BottomFrame'].location.href='../../"&KS.Setting(89)&"Post.Asp?OpStr=手机版自定义页面管理中心 >> 手机版页面管理&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='setting.asp?action=template';}</script>")
		End Sub
		
		Sub EditSave()
		    Dim RS
			Dim TemplateID,TemplateName,TemplateContent
			TemplateID=KS.G("TemplateID")
		
			TemplateName = Replace(Replace(Trim(Request.Form("TemplateName")), """", ""), "'", "")
			TemplateContent = Trim(Request.Form("TemplateContent"))
			If TemplateName="" then
			   Response.write "<script>window.alert('标题没有填写..');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			If TemplateContent="" then
			   Response.write "<script>window.alert('操作失败...');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			Set RS = Server.CreateObject("ADODB.RecordSet")
			RS.Open "select * from KS_WapTemplate where ID='" & TemplateID & "'",Conn,1,3
			RS("TemplateName")=TemplateName
			RS("TemplateContent")=TemplateContent
			RS("AddDate")=Now
			RS.Update
			RS.Close:set RS=Nothing
			Response.Write "<script>alert('成功提示:\n\n自定义页面修改成功!');parent.frames['BottomFrame'].location.href='../../"&KS.Setting(89)&"Post.Asp?OpStr=手机版自定义页面管理中心  >> 手机版页面管理&ButtonSymbol=FreeLabel';location.href='setting.asp?action=template';</script>"
		End Sub
		
		'添加页面
		Sub AddLabel()
		    Dim TemplateName,TemplateContent
			If KS.G("flag") = "EditLabel" Then
			   Dim RS,TemplateID
			   TemplateID=KS.G("TemplateID")
			   Set RS = Conn.Execute("select * from KS_WapTemplate where ID='" & TemplateID & "'")
			   TemplateName=RS("TemplateName")
			   TemplateContent=RS("TemplateContent")
			   RS.Close:set RS=Nothing
			Else
			   TemplateName=""
			    TemplateContent=TemplateContent&"<!DOCTYPE html>" & vbcrlf
				TemplateContent=TemplateContent&"<html>" & vbcrlf
				TemplateContent=TemplateContent&"<head> " & vbcrlf
				TemplateContent=TemplateContent&"<title>{$TemplateName}</title>" & vbcrlf
				TemplateContent=TemplateContent&"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbcrlf
				TemplateContent=TemplateContent&"<meta http-equiv=""Cache-control"" content=""max-age=1700"">" & vbcrlf
				TemplateContent=TemplateContent&"<meta name=""viewport"" content=""user-scalable=no, width=device-width"">" & vbcrlf
				TemplateContent=TemplateContent&"<meta name=""MobileOptimized"" content=""320"">" & vbcrlf
				TemplateContent=TemplateContent&"<meta name=""author"" content=""kesion.com"">" & vbcrlf
				TemplateContent=TemplateContent&"<meta name=""format-detection"" content=""telephone=no"">" & vbcrlf
				TemplateContent=TemplateContent&"</head>" & vbcrlf
				TemplateContent=TemplateContent&"<body>" & vbcrlf& vbcrlf
				TemplateContent=TemplateContent&"页面内容" & vbcrlf& vbcrlf
				TemplateContent=TemplateContent&"</body>" & vbcrlf
				TemplateContent=TemplateContent&"</html>" & vbcrlf

	
			End if
			%>
<script language="JavaScript" src="../ks_inc/Common.js"></script>
<script>
parent.frames['BottomFrame'].location.href='../../<%=KS.Setting(89)%>Post.Asp?OpStr=手机版自定义页面管理中心 >> 添加自定义页面&ButtonSymbol=Go'

function InsertFunctionLabel(Url,Width,Height)
{
var Val = OpenWindow(Url,Width,Height,window);if (Val!=''&&Val!=null){ document.LabelForm.TemplateContent.focus();
  var str = document.selection.createRange();
  str.text = Val; }
}

		function show_ln(txt_ln,txt_main){
			var txt_ln  = document.getElementById(txt_ln);
			var txt_main  = document.getElementById(txt_main);
			txt_ln.scrollTop = txt_main.scrollTop;
			while(txt_ln.scrollTop != txt_main.scrollTop)
			{
				txt_ln.value += (i++) + '\n';
				txt_ln.scrollTop = txt_main.scrollTop;
			}
			return;
		}
		function editTab(){
			
			}
			
		</script>
</head>       
<body scroll=no leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<div class='topdashed sort'><font>
      <%
	  If request("flag") = "EditLabel" Then
	     Response.Write "修改页面"
	  Else
	     Response.Write "新建页面"
	  End If
	  %></font></div>
<div class="pageCont2">
<form name="LabelForm" method="post" Action="setting.asp" onSubmit="return(CheckForm())">
<input type="hidden" name="TemplateID" value="<%=TemplateID%>">
<input type="hidden" name="Page" value="<%=KS.S("Page")%>">
<input type="hidden" name="action" value="template">
<table width='100%' height="350" border='0' align='center' cellpadding='0' cellspacing='0' class='otable'>
<%
If KS.G("flag") = "AddNew" Or KS.G("flag") = "" Then Response.Write "<input type='hidden' name='flag' value='AddSave'>"
If KS.G("flag") = "EditLabel" Then Response.Write "<input type='hidden' name='flag' value='EditSave'>"
%>
<tr class="clefttitle">
<td height="30" style="text-align:left" class="splittd noborder"><b>页面名称：</b><input class="textbox" name="TemplateName" type="text" id="TemplateName" size="50" Value="<%=TemplateName%>">  </td>
</tr>

<tr id="toplabelarea" class="clefttitle">
<td valign="top" style="text-align:left" class="splittd noborder">
<strong>插入标签：</strong>
<select name="mylabel" style="width:160px">
<%
		 Response.Write " <option value="""">==选择系统函数标签==</option>"
		  Set RS=KS.InitialObject("adodb.recordset")
		   rs.open "select top 500 LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
%>
</select>&nbsp;
<input class='button' type='button' onclick='LabelInsertCode(document.all.mylabel.value);' value='插入标签'>&nbsp;

&nbsp;页面设计代码建议用：HTML5+CSS3
</td>
</tr>

<tr id="codearea">
<td>
         <table border='0' width="100%" cellspacing='0' cellspadding='0'>
         <tr>
         <td valign="top" width='20'>
         <textarea name="txt_ln" id="txt_ln" cols="6" style="overflow:hidden;height:410;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;" readonly><%
		 Dim N
		 For N=1 To 500
		 Response.Write N & vbcrlf
		 Next
		 %>
         </textarea>
         </td>
         <td valign="top">
         <textarea name="TemplateContent" onclick='setPos()' onkeyup='setPos()' rows="2" cols="30" id="TemplateContent" onscroll="show_ln('txt_ln','TemplateContent')" onKeyDown="editTab()" style="height:410px;width:90%;"><%=TemplateContent%></textarea>
         <script>for(var i=500; i<=500; i++) document.getElementById('txt_ln').value += i + '\n';</script>
         </td>
         </tr>
         </table>

</table>
</form>
</div>
</body>
</html>

<script language="JavaScript">
<!--
var pos=null;
function setPos()
{ if (document.all){
		document.LabelForm.TemplateContent.focus();
		pos = document.selection.createRange();
  }else{
		pos = document.getElementById("TemplateContent").selectionStart;
	 }
}
function LabelInsertCode(Val)
{
 if(pos==null) {alert('请先定位插入位置!');return false;} 
  LabelInsert(Val); 
}
function LabelInsert(Val){
if (Val!=''){ if (document.all){ pos.text=Val; }else{
  var obj=$("#TemplateContent");
  var lstr=obj.val().substring(0,pos);
	   var rstr=obj.val().substring(pos);
	   obj.val(lstr+Val+rstr);			 }
 }
}

					
function CheckForm()
   {
   var form=document.LabelForm;
   if (form.TemplateName.value=='')
      {
	  alert('请输入自定义页面名称!');form.TemplateName.focus();
	  return false;
	  }
   if (form.TemplateContent.value==''||form.TemplateContent.value=='请输入您自定义的代码')
      {
	  alert('请输入自定义页面内容!');
	  form.TemplateContent.focus();
	  return false;
	  }
	  form.submit();
	  return true;
   }
//-->
</script>
<%
		End Sub
		
	
%> 

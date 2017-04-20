<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<!DOCTYPE html><html>
<head>
<link href="Admin_style.css" rel="stylesheet" type="text/css">
<script src="../../ks_inc/jquery.js"></script>
</head>
<body style="background: #EAF0F5;">
<%
Dim KS:Set KS=New PublicCls
Dim Wjj,BH,ext,fname,ItemName
ItemName=KS.G("ItemName")
if KS.IsNul(ItemName) then ItemName="图片"
 if KS.G("wjj")<>"" Then
  Wjj=KS.G("WJJ")
 ELSE
  wjj=request("CurrPath") & "/"
End If
if left(lcase(wjj),len(KS.Setting(3) & KS.Setting(91)))<>lcase(KS.Setting(3) & KS.Setting(91)) then ks.die "error!"
if request("action")="save" then
  call KS.CreateListFolder(wjj)
  http=trim(request.Form("http"))
  if http="" then
   Response.Write"<script>alert('请输入远程" & ItemName &"地址!');</script>"
   Response.End()
  end if
  ext=right(http,4)
  fname=wjj&year(now)&month(now)&day(now)&hour(now)&second(now)&KS.MakeRandom(5)&ext
  dim fname1:fname1=fname
  if instr(fname1,".")=0 then
   KS.AlertHintScript "对不起，远程文件不合法!"
  end if
  ext=lcase(split(fname1,".")(1))
  if (ext<>"jpg" and ext<>"jpeg" and ext<>"gif" and ext<>"bmp" and ext<>"png") or instr(fname1,";")>0 then
  %>
 <script type="text/javascript">
   alert('对不起,只能保存图片jpg|jpeg|gif|png的文件!');
   top.box.close();
 </script>
  <%
   response.end
  end if

  
  Call KS.SaveBeyondFile(fname1,http)
  If KS.Setting(97)="1" Then
    If Left(lcase(fname),4)<>"http" then fname=KS.Setting(2) & fname
  End If
%>
 <script>
    alert('成功保存了远程<%=ItemName%>!');
	 top.frames['MainFrame'].document.getElementById('<%=request("fieldID")%>').value='<%=fname%>';
	<% if request("pic")<>"" and request("pic")<>"undefined" then
		  response.write "top.frames['MainFrame'].document.getElementById('"& request("pic") &"').src='" & fname & "';" &vbcrlf
	  end if
	  if lcase(request("ieditor")&"")="true" then
		  response.write "top.frames['MainFrame'].insertHTMLToEditor(""<img src='" & fname & "' />"");" &vbcrlf
	  end if
	 %>
top.box.close();
 </script>

<%
  Response.Write("远程" & ItemName &"保存成功!")
end if
%>
<script>
  $(document).ready(function(){
    $("#http").focus();
 });

</script>
<div align="center">
<br>
<form name="myform" action="?action=save" method="post">
<input type="hidden" name="ItemName" value="<%=ItemName%>" />
<input type="hidden" name="ieditor" value="<%=request("ieditor")%>" />
<input type="hidden" name="Pic" value="<%=request("pic")%>" />
<input type="hidden" name="FieldID" value="<%=request("fieldid")%>" />
<input type="hidden" value="<%=wjj%>" name="wjj" />
远程<%=ItemName%>地址：<input type="text" class="textbox" id="http" name="http">
<input type="submit" name="Submit" class="button" value="开始抓取" onClick="if ($('#http').val()==''){alert('请输入远程<%=ItemName%>地址！');$('#http').focus(); return false;}"><br><br>
形如:<font color=red>http://www.kesion.com/images/logo.gif</font>
</form>
</div>
</body>
</html>
 

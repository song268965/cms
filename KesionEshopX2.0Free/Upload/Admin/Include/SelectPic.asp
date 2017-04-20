<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Dim KSCls
Set KSCls = New SelectPic
KSCls.Kesion()
Set KSCls = Nothing

Class SelectPic
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		With KS
		If KS.C("AdminName") = "" Then
		 .echo ("<script>alert('对不起，权限不足!');window.close();</script>")
		 Exit Sub
		End If
		Dim ChannelID, CurrPath, ShowVirtualPath
		Dim InstallDir
		Dim LimitUpFileFlag  '上传权限 值yes无上传权限
			InstallDir = KS.Setting(3)
			ChannelID = KS.G("ChannelID")
			CurrPath = KS.G("CurrPath")
			ShowVirtualPath = KS.G("ShowVirtualPath")
			If ChannelID = "" Or Not IsNumeric(ChannelID) Then ChannelID = 0
				 If KS.ReturnChannelAllowUpFilesTF(ChannelID) = False Then
				  LimitUpFileFlag = "yes"
				 End If
				 If KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009") = False Then
				  LimitUpFileFlag = "yes"
				 End If
				 
			if instr(request("currpath"),".")<>0 then
			  ks.die "非法参数!"
			end if
           If InstallDir<>"/" then 
			if instr(CurrPath,InstallDir)=0 Then
			CurrPath = Replace(InstallDir & CurrPath,"//","/")
			End If
		  End iF
		  if left(lcase(currpath),len(KS.Setting(3) & KS.Setting(91)))&"/"<>lcase(KS.Setting(3) & KS.Setting(91)) then currpath=KS.GetUpFilesDir
		  If KS.C("SuperTF")="1" Then CurrPath=KS.Setting(3) & left(ks.setting(91),len(ks.setting(91))-1)
		  
		  if currpath="/" then currpath =ks.setting(3) & left(ks.setting(91),len(ks.setting(91))-1)
		
		.echo "<!DOCTYPE html>"
		.echo "<html>"
		.echo "<head>"
		.echo "<META HTTP-EQUIV=""pragma"" CONTENT=""no-cache"">" 
        .echo "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache, must-revalidate"">"
        .echo "<META HTTP-EQUIV=""expires"" CONTENT=""Wed, 26 Feb 1997 08:21:57 GMT"">"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		.echo "<title>选择文件</title>"
		.echo "<style type=""text/css"">" & vbCrLf
		.echo "<!--" & vbCrLf
		.echo ".PreviewStyle {" & vbCrLf
		.echo "    border: 2px outset #CCCCCC;"
		.echo "}"
		.echo ".ImgOver {"
		.echo "    cursor: default;"
		.echo "    border-top-width: 1px;"
		.echo "    border-right-width: 1px;"
		.echo "    border-bottom-width: 1px;"
		.echo "    border-left-width: 1px;"
		.echo "    border-top-style: solid;"
		.echo "    border-right-style: solid;"
		.echo "    border-bottom-style: solid;"
		.echo "    border-left-style: solid;"
		.echo "    border-top-color: #FFFFFF;"
		.echo "    border-right-color: #999999;"
		.echo "    border-bottom-color: #999999;"
		.echo "    border-left-color: #FFFFFF;"
		.echo "}"
		.echo " BODY   {border: 0; margin: 0; cursor: default; font-family:宋体; font-size:9pt;}"
		.echo " BUTTON {width:5em}" & vbCrLf
		.echo " TABLE  {font-family:宋体; font-size:9pt}"
		.echo " P      {text-align:center}" & vbCrLf
		.echo "-->" & vbCrLf
		.echo "</style>"
		.echo "</head>"
		.echo "<script src=""../../KS_inc/jquery.js""></script>"
		.echo "<script src=""../../KS_inc/Common.js""></script>"
		%>
		<script>
		$(document).ready(function(){
		  rz();
		  uploadBtn("<%=currpath %>");
		  $(window).resize(function () { 
		   rz();
		  });
		});

		function rz(){
		   $("#FolderList").height($(window).height()-115);
		   $("#PreviewArea").height($(window).height()-115);
       }

		function uploadBtn(path){
		  jQuery("#upload").attr("src","../System/KS.UpFileForm.asp?currpath=" +path+"&UPType=UpByBar&ChannelID=<%=channelid%>&from=getfile");
		}
		</script>
		<%
		.echo "<body leftmargin=""0"">"
		.echo "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <tr>"
		 .echo "   <td colspan=""2""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		 .echo "       <tr>"
		 .echo "         <td width=""80"" align=""center"" nowrap>选择目录： </td>"
		 .echo "         <td width=""649""><select onChange=""ChangeFolder(this.value);"" id=""FolderSelectList"" name=""FolderSelectList"" style=""width:100%;"" name=""select"">"
		 .echo "             <option selected value=""" & CurrPath & """>"
		 .echo CurrPath
		 .echo "             </option>"
		 .echo "           </select> </td>"
		.echo "          <td width=""279"" height=""26"" valign=""middle"">"
		 .echo "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "             <tr align=""center"">"
		 .echo "               <td width=""25"">&nbsp;</td>"
		 .echo "                <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><i class=""icon delete"" align=""absmiddle"" onClick=""ChangeViewArea(this);"" id=""Img1"" title=""关闭预览区""></i></td>"
		 .echo "               <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><i class=""icon back"" width=""21"" height=""22"" align=""absmiddle"" onClick=""frames['FolderList'].OpenParentFolder();"" title=""返回上一级目录""></i></td>"
		 .echo "               <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><i class=""icon add3"" width=""19"" height=""17"" align=""absmiddle"" onClick=""frames['FolderList'].AddFolderOperation();"" title=""添加新目录""></i></td>"

		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "              </tr>"
		.echo "            </table>"
		.echo "          </td>"
		.echo "        </tr>"
		.echo "      </table>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td width=""70%"" align=""center""> <iframe name=""FolderList"" id=""FolderList"" width=""100%"" height=""340"" frameborder=""1"" src=""FolderFileList.asp?ChannelID=" & ChannelID & "&CurrPath=" & CurrPath & "&ShowVirtualPath=" & ShowVirtualPath & """ scrolling=""yes""></iframe>"
		.echo "    </td>"
		.echo "    <td width=""30%""  align=""center"" valign=""middle"" id=""ViewArea""> <iframe name=""PreviewArea"" id=""PreviewArea"" scrolling=""yes"" width=""100%"" height=""340"" frameborder=""1"" src=""Preview.asp""></iframe>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td height=""35"" colspan=""2""> <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "        <tr>"
		.echo "          <td width=""80"" height=""40""> <div align=""center"">URL地址：</div></td>"
		.echo "          <td><input style=""width:65%"" class=""textbox"" type=""text"" name=""FileUrl"" id=""FileUrl""> <input type=""button"" onClick=""SetFileUrl();"" name=""Submit"" value="" 确 定 "" class=""button""/>"
		.echo "            <input onClick=""closeWin();"" class=""button"" type=""button"" name=""Submit3"" value="" 取 消 ""/>"
		.echo "          </td>"
		.echo "        </tr>"
		.echo "      </table></td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td height=""35"" colspan=""2""> <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "        <tr>"
		.echo "          <td width=""80"" height=""40""> <div align=""center"">上传文件：</div></td>"
		.echo "          <td><iframe id='upload' name='upload' src='about:blank;' frameborder=0 scrolling=no width='100%' height='30'></iframe>"
		.echo "          </td>"
		.echo "        </tr>"
		.echo "      </table></td>"
		.echo "  </tr>"

		
		.echo "</table>"
		.echo "</body>"
		.echo "</html>"
		.echo "<script language=""JavaScript"">"
		.echo "var ChannelID=" & ChannelID & ";"
		.echo "function closeWin(){" &vbcrlf
		.echo " try{ top.frames['MainFrame'].frames[0].box.close();}catch(e){ try { top.frames['MainFrame'].box.close();}catch(e){top.close();}}" &vbcrlf
		.echo "}" &vbcrlf
		.echo "function ChangeFolder(FolderName)"
		.echo "{"
		.echo "    frames[""FolderList""].location='FolderFileList.asp?CurrPath='+FolderName;"
		.echo "}"
		
		.echo "function SetFileUrl()"
		.echo "{"
		.echo "    if (document.getElementById('FileUrl').value=='') alert('请填写Url地址');"
		.echo "    else"
		 if request("fieldid")<>"" then
		 .echo "{ try{ top.frames['MainFrame'].frames[0].document.getElementById('" &KS.S("fieldID")&"').value=document.getElementById('FileUrl').value;}catch(e){" &vbcrlf
		 .echo "top.frames['MainFrame'].document.getElementById('" &KS.S("fieldID")&"').value=document.getElementById('FileUrl').value;}"
		 if request("pic")<>"" and request("pic")<>"undefined" then
		  .echo "top.frames['MainFrame'].document.getElementById('"& KS.S("pic") &"').src=document.getElementById('FileUrl').value;" &vbcrlf
		 end if
		 .echo " closeWin();" &vbcrlf
		 .echo "}"&vbcrlf
		else
		.echo "    {"
		.echo "       if (document.all){ window.returnValue=document.getElementById('FileUrl').value;}else{window.opener.setVal(document.getElementById('FileUrl').value)}"
		.echo "        closeWin();"
		.echo "    }"
		end if
		.echo "}"
		.echo "window.onunload=CheckReturnValue;"
		.echo "function CheckReturnValue()"
		.echo "{"
		.echo "    if (typeof(window.returnValue)!='string') window.returnValue='';"
		.echo "}"
		.echo "var displayBar=true;"
		.echo "function ChangeViewArea(obj) {"
		.echo "$('#ViewArea').toggle()" &vbcrlf
		.echo "  if (displayBar) {"
		.echo " $('#FolderList').width($(window).width());"
		.echo "    displayBar=false;"
		.echo "    obj.src='../Images/Folder/L.gif';"
		.echo "    obj.title='打开预览区';"
		.echo "  } else {"
		.echo " $('#FolderList').width($(window).width()-$(window).width()*30/100);"
		.echo "    displayBar=true;"
		.echo "    obj.src='../Images/Folder/R.gif';"
		.echo "    obj.title='关闭预览区';"
		.echo "  }"
		.echo "}"
		
		.echo "</script>"
		End With
		End Sub
End Class
%> 

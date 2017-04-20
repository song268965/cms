<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New GetSJList
KSCls.Kesion()
Set KSCls = Nothing

Class GetSJList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
'主体部分
Public Sub Kesion()
Dim InstallDir, CurrPath, FolderID, LabelContent, Action, LabelID, Str, Descript,dtfs,Popular,Recommend
Dim TypeFlag, Num, TitleLen,ChannelID,PrintType,AjaxOut,LabelStyle,ClassID,OrderStr,BigClassID,SmallClassID,DateRule,Page
FolderID = Request("FolderID")
Page=Request("Page")
CurrPath = KS.GetCommonUpFilesDir()
With KS
'判断是否编辑
LabelID = Trim(Request.QueryString("LabelID"))
If LabelID = "" Then
  Action = "Add":DateRule="YYYY-MM-DD"
Else
    Action = "Edit"
  Dim LabelRS, LabelName
  Set LabelRS = Server.CreateObject("Adodb.Recordset")
  LabelRS.Open "Select top 1 * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
  If LabelRS.EOF And LabelRS.BOF Then
     LabelRS.Close
     Set LabelRS = Nothing
     .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
     .End
  End If
    LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
    FolderID = LabelRS("FolderID")
    Descript = LabelRS("Description")
    LabelContent = LabelRS("LabelContent")
    LabelRS.Close
    Set LabelRS = Nothing
            LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetSjxdList", ""),"}" & LabelStyle&"{/Tag}", "")
			' response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			    ClassID          = Node.getAttribute("bigclassid")
				SmallClassID     = Node.getAttribute("smallclassid")
				DateRule         = Node.getAttribute("daterule")
				Num              = Node.getAttribute("num")
				TitleLen         = Node.getAttribute("titlelen")
				AjaxOut          = Node.getAttribute("ajaxout")
				OrderStr         = Node.getAttribute("orderstr")
				dtfs             = Node.getAttribute("dtfs")
				Recommend        = Node.getAttribute("recommend")
				Popular          = Node.getAttribute("popular")
			End If
			XMLDoc=Empty
			Set Node=Nothing
    
End If
		If TitleLen="" Then TitleLen=0
		If Num = "" Then Num = 10
		If dtfs="" Then dtfs=0
		If recommend="" Then recommend=0
		If popular="" Then popular=0
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@sjxdurl}"" target=""_blank"">{@content}</a></li>" & vbcrlf & "[/loop]"
		
		.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
%>
        <style type="text/css">
		 .field{width:720px;}
		 .field li{cursor:pointer;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:18px;line-height:18px;margin:3px 1px 0px;padding:2px}
		 .field li.diyfield{border:1px solid #f9c943;background:#FFFFF6}
		</style>
        <script type="text/javascript">
		$(document).ready(function(){
	   })
		
       function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		 function setPos()
		 { if (document.all){
				$("#LabelStyle").focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("LabelStyle").selectionStart;
			  }
		 }
		 //插入
		function InsertValue(Val)
		{  if (pos==null) {top.$.dialog.alert('请先定位要插入的位置!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#LabelStyle");
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }
	

	function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  top.$.dialog.alert('请输入标签名称!',function(){
			  $("input[name=LabelName]").focus();}); 
			  return false
			  }
	//var ClassID=document.myform.ClassID.value;
	
	//var SmallClassID=document.myform.SmallClassID.value;
	var Num=document.myform.Num.value;
	var TitleLen=document.myform.TitleLen.value;
	var DateRule=document.myform.DateRule.value;
	var OrderStr=$("#OrderStr").val();
	var AjaxOut=false;
	if ($("#AjaxOut").prop("checked")==true){AjaxOut=true}
	var recommend=0;
	if ($("#recommend").prop("checked")==true){recommend=1;}	
	if (Num=='') Num=10
	var tagVal='{Tag:GetSjxdList labelid="0" ajaxout="'+AjaxOut+'" recommend="'+recommend+'" num="'+Num+'" orderstr="'+OrderStr+'" daterule="'+DateRule+'" titlelen="'+TitleLen+'"}'+$("#LabelStyle").val()+'{/Tag}';
	$("input[name=LabelContent]").val(tagVal);
	$("#myform").submit();
}
</script>
<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"">"
		.echo "<div class='pageCont2'>"
		.echo "<iframe src='about:blank' name='_hiddenframe' style='display:none' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Page"" id=""Page"" value=""" & Page & """>"
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSjxdList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"

		
		
		.echo "</td><td>日期格式：" & ReturnDateFormat(DateRule) & "</td>"
		.echo "            </tr>"
		
		
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示条数"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:50px"" value=""" & Num & """>个 名称字数<input name=""TitleLen"" class=""textbox"" type=""text"" id=""TitleLen"" style=""width:50px"" value=""" & TitleLen & """><font color=red>如果不想控制，请设置为“0”</font></td>"

.echo "              <td height=""30"">排序方式"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "a.ID Desc" Then
					.echo ("<option value='a.ID Desc' selected>心得ID(降序)</option>")
					Else
					.echo ("<option value='a.ID Desc'>心得ID(降序)</option>")
					End If
					If OrderStr = "a.ID Asc" Then
					.echo ("<option value='a.ID Asc' selected>心得ID(升序)</option>")
					Else
					.echo ("<option value='a.ID Asc'>心得ID(升序)</option>")
					End If

		.echo "         </select></td>"
.echo "            </tr>"	


		.echo "            <tbody>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@id}')"">考后心得ID</li><li onclick=""InsertLabel('{@sjxdurl}')"">考后心得Url</li><li onclick=""InsertLabel('{@sjid}')"">试卷ID</li> <li onclick=""InsertLabel('{@title}')"">试卷名称</li><li onclick=""InsertLabel('{@sjurl}')"">试卷Url</li> <li onclick=""InsertLabel('{@sjxdusername}')"">用户名称</li><li onclick=""InsertLabel('{@kssj}')"">发布时间</li><li onclick=""InsertLabel('{@content}')"">考后心得</li><li onclick=""InsertLabel('{@userip}')"">用户IP</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
		



.echo "                  </table>"	
.echo "  </form>"
  
.echo "</div>"
.echo "</body>"
.echo "</html>"
End With

End Sub
End Class
%> 

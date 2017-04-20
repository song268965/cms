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
Dim InstallDir, CurrPath, FolderID, LabelContent, Action, LabelID, Str, Descript,dtfs,Popular,Recommend,Page
Dim TypeFlag, Num, TitleLen,ChannelID,PrintType,AjaxOut,LabelStyle,ClassID,OrderStr,BigClassID,SmallClassID,DateRule
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetSjList", ""),"}" & LabelStyle&"{/Tag}", "")
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
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@sjurl}"" target=""_blank"">{@title}</a></li>" & vbcrlf & "[/loop]"
		
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
	var ClassID=document.myform.ClassID.value;
	var SmallClassID=document.myform.SmallClassID.value;
	var Num=document.myform.Num.value;
	var TitleLen=document.myform.TitleLen.value;
	var DateRule=document.myform.DateRule.value;
	var OrderStr=$("#OrderStr").val();
	var AjaxOut=false;
	if ($("#AjaxOut").prop("checked")==true){AjaxOut=true}
	var recommend=0;
	if ($("#recommend").prop("checked")==true){recommend=1;}
	var popular=0;
	if ($("#popular").prop("checked")==true){popular=1;}
			
	if (Num=='') Num=10
	var dtfs=$("input[name='dtfs']:checked").val();
	var tagVal='{Tag:GetSjList labelid="0" ajaxout="'+AjaxOut+'" recommend="'+recommend+'" popular="'+popular+'" dtfs="'+dtfs+'" bigclassid="'+ClassID+'" smallclassid="'+SmallClassID+'" num="'+Num+'" orderstr="'+OrderStr+'" daterule="'+DateRule+'" titlelen="'+TitleLen+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSjList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"

		.echo " &nbsp;&nbsp;&nbsp;属性控制 <label><input id=""recommend"" name=""recommend"" type=""checkbox"" value=""1"""
		If recommend = "1" Then .echo (" Checked")
		.echo ">推荐</label>"
		
		.echo "&nbsp;&nbsp;<label><input name=""popular"" id=""popular"" type=""checkbox"" value=""1"""
		If Popular = "1" Then .echo (" Checked")
		  .echo ">热门</label>"
		
		.echo "</td><td>日期格式：" & ReturnDateFormat(DateRule) & "</td>"
		.echo "            </tr>"
		
		
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示条数"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:50px"" value=""" & Num & """>个 名称字数<input name=""TitleLen"" class=""textbox"" type=""text"" id=""TitleLen"" style=""width:50px"" value=""" & TitleLen & """><font color=red>如果不想控制，请设置为“0”</font></td>"

.echo "              <td height=""30"">排序方式"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>试卷ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>试卷ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>试卷ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>试卷ID(升序)</option>")
					End If
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>点击数(升序)</option>")
					Else
					 .echo ("<option value='Hits Asc'>点击数(升序)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>点击数(降序)</option>")
					Else
					  .echo ("<option value='Hits Desc'>点击数(降序)</option>")
					End If
	

		.echo "         </select></td>"
.echo "            </tr>"	



		.echo "            <tr class='tdbg' id='spaceclass'>"
		.echo "              <td height=""30"" colspan='2'>&nbsp;&nbsp;&nbsp;所属分类"
		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		sqls = "select * from KS_SJClass Where tj=2 order by id"
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("tn"))%>","<%=split(rss("tname"),"|")(ubound(split(rss("tname"),"|"))-1)%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0;
			document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option("--二级分类不限--","0");
 
			var locationid=locationid;
			var i;
			for (i=0;i < onecount; i++)
				{
					if (subcat[i][1] == locationid)
					{ 
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		 <select class="textbox" name="ClassID" id="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
		<option value='0'>--一级分类不限--</option>
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
		sqlb = "select * from ks_sjclass where tj=1 order by id"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    do while not rsb.eof
					  If trim(ClassID)=trim(rsb("id")) then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=split(rsb("tname"),"|")(ubound(split(rsb("tname"),"|"))-1)%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=split(rsb("tname"),"|")(ubound(split(rsb("tname"),"|"))-1)%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" id="SmallClassID" name="SmallClassID">
				   <option value='0'>--二级分类不限--</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_sjclass where tn="& ks.chkclng(ClassID)&" order by orderid"
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if trim(SmallClassID)=trim(rsss("id")) then%>
							<option value="<%=rsss("id")%>" selected><%=split(rsss("tname"),"|")(ubound(split(rsss("tname"),"|"))-1)%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=split(rsss("tname"),"|")(ubound(split(rsss("tname"),"|"))-1)%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select>
				
				试卷限制：<input type='radio' name='dtfs' value='0'<%if dtfs="0" then response.write " checked"%>>不限 <input type='radio' name='dtfs' value='1'<%if dtfs="1" then response.write " checked"%>>整份试卷 <input type='radio' name='dtfs' value='2'<%if dtfs="2" then response.write " checked"%>>随机组合的试卷
		<%
				  
.echo "                </td>"

.echo "            </tr>"
		
		.echo "            <tbody>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@sjurl}')"">试卷Url</li><li onclick=""InsertLabel('{@id}')"">试卷ID</li> <li onclick=""InsertLabel('{@title}')"">试卷名称</li><li onclick=""InsertLabel('{@sjtypeurl}')"">分类Url</li> <li onclick=""InsertLabel('{@sjtypename}')"">分类名称</li><li onclick=""InsertLabel('{@kssj}')"">考试时间</li><li onclick=""InsertLabel('{@form_user}')"">作者</li><li onclick=""InsertLabel('{@user}')"">录入员</li><li onclick=""InsertLabel('{@form_url}')"">来源</li><li onclick=""InsertLabel('{@hits}')"">点击数</li><li onclick=""InsertLabel('{@adddate}')"">上传时间</li><li onclick=""InsertLabel('{@sq}')"">所需点数</li><li onclick=""InsertLabel('{@sjsq}')"" style=""color:red"">点数带图标</li><li onclick=""InsertLabel('{@sjzf}')"">试卷总分</li></td>"
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

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
Set KSCls = New GetJobZWList
KSCls.Kesion()
Set KSCls = Nothing

Class GetJobZWList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim JobType, OpenType, RecommendTF,ShowJiPinFlag, Num,  InfoSort, ColNumber, Province, NavType, Navi, DateRule, DateAlign, JiPin, City,County,ShowStyle, PrintType,AjaxOut,LabelStyle,Page
		FolderID = Request("FolderID")
		Page=Request("Page")
		CurrPath = KS.GetCommonUpFilesDir()

		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  JobType = "0"
		  Action = "Add"
		Else
		  Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Conn.Close:Set Conn = Nothing
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close:Set LabelRS = Nothing

            LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetJobZWList", ""),"}" & LabelStyle & "{/Tag}", "")
			'Response.Write labelcontent
			
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
				JobType    = Node.getAttribute("jobtype")
				ShowStyle  = Node.getAttribute("showstyle")
				Province   = Node.getAttribute("province")
				City       = Node.getAttribute("city")
				County     = Node.getAttribute("county")
				RecommendTF= Cbool(Node.getAttribute("recommendtf"))
				JiPin      = Cbool(Node.getAttribute("jipin"))
				OpenType   = Node.getAttribute("opentype")
				Num        = Node.getAttribute("num")
				InfoSort   = Node.getAttribute("infosort")
				ColNumber  = Node.getAttribute("col")
				NavType    = Node.getAttribute("navtype")
				Navi       = Node.getAttribute("nav")
				PrintType  = Node.getAttribute("printtype")
				AjaxOut    = Node.getAttribute("ajaxout")
				ShowJiPinFlag=Node.GetAttribute("showjipinflag")
			End If
			Set Node=Nothing
			XMLDoc=Empty
		End If
		If PrintType="" Then PrintType=1
		If Num = "" Then Num = 20
		If ColNumber = "" Then ColNumber = 1
		If RecommendTF="" Then RecommendTF=False
		If ShowStyle="" Then ShowStyle=2
		If JobType="" Then JobType=0
		If AjaxOut="" Then AjaxOut=false
		If KS.IsNul(ShowJiPinFlag) Then ShowJiPinFlag=false
		If LabelStyle="" Then LabelStyle="<li><a href=""{@jobzwurl}"">{@jobtitle}</a></li>"
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
		 ChangeOutArea($("#PrintType").val());
		});
		function ChangeOutArea(Val)
		{ 
		 if (Val==2){
		  $("#DiyArea").show();
		 }
		 else{
		  $("#DiyArea").hide();
		 }
		}
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   { $("#NavWord").show();
			 $("#NavPic").hide();
		  }else{
		     $("#NavWord").hide();
		     $("#NavPic").show();
		 }
		}
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
			var JobType;
			var ShowStyle=$('#ShowStyle').val();
			var NavType=1;
			var Province=$('#Province').val();
			var City=$('#City').val();
			var County=$('#County').val();
			var OpenType=$('#OpenType').val();
			var Num=$('input[name=Num]').val();
			var InfoSort=$('select[name=InfoSort]').val();
			var ColNumber=$('input[name=ColNumber]').val();
			var Nav,NavType=$('select[name=NavType]').val();
			var PrintType=$('select[name=PrintType]').val();
            var JobType=$("input[name=JobType]:checked").val();
			var RecommendTF=false;
			if ($("#RecommendTF").prop("checked")==true){RecommendTF=true}
			var JiPin=false;
			if ($("#JiPin").prop("checked")==true){JiPin=true}
			var ShowJiPinFlag=false;
			if ($("#ShowJiPinFlag").prop("checked")==true){ShowJiPinFlag=true}
			if  (Num=='')  Num=10;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav=$('#TxtNavi').val()
			 else  Nav=$('#NaviPic').val();
            var AjaxOut=false;
			if ($("#AjaxOut").prop("checked")==true){AjaxOut=true}

            var tagVal='{Tag:GetJobZWList labelid="0" jobtype="'+JobType+'" showstyle="'+ShowStyle+'" province="'+Province+'" city="'+City+'" county="'+County+'" recommendtf="'+RecommendTF+'" showjipinflag="'+ShowJiPinFlag+'" opentype="'+OpenType+'" num="'+Num+'" infosort="'+InfoSort+'" col="'+ColNumber+'" jipin="'+JiPin+'" navtype="'+NavType+'" nav="'+Nav+'" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Page"" id=""Page"" value=""" & Page & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetJobZWList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通输出</option>"
        .echo " <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">自定义输出样式</option>"
        
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"
		.echo"</td>"
		.echo "            </tr>"
		
        .echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@jobzwurl}')"">职位URL</li> <li onclick=""InsertLabel('{@jobtitle}')"">职位名称</li><li onclick=""InsertLabel('{@companyshortname}')"">单位简称</li><li onclick=""InsertLabel('{@companyname}')"">单位全称</li><li onclick=""InsertLabel('{@jobcompanyurl}')"">单位URL</li><li onclick=""InsertLabel('{@province}')"">工作省份</li><li onclick=""InsertLabel('{@city}')"">城市</li><li onclick=""InsertLabel('{@county}')"">城镇</li><li onclick=""InsertLabel('{@sex}')"">性别</li><li onclick=""InsertLabel('{@qualifications}')"">学历</li><li onclick=""InsertLabel('{@workexperience}')"">工作经验</li> <li onclick=""InsertLabel('{@zpnum}')"">人数</li><li onclick=""InsertLabel('{@salary}')"">待遇</li><li onclick=""InsertLabel('{@refreshtime}')"">时间</li><li onclick=""InsertLabel('{@hits}')"">浏览数</li><li onclick=""InsertLabel('{@jipin}')"" style='color:red'>急聘</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"			
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">工作地区&nbsp;"
		 Response.Write "<script type='text/javascript'>"
							Response.write "try{setCookie(""pid"",'" & Province & "');setCookie(""cid"",'" &  City & "');}catch(e){}" & vbcrlf
							Response.write "</script>"
							%>
							 <script src="../../../plus/area.asp" language="javascript"></script>
							<script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
							 <%end if%>
							  <%if City<>"" Then%>
							 $('#City').val('<%=City%>');
							  <%end if%>
							  <%if County<>"" Then%>
							 $('#County').val('<%=County%>');
							  <%end if%>
							</script>
							<%
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">显示特效"
		.echo "                <select class='textbox' name=""ShowStyle"" id='ShowStyle' style=""width:70%;"">"
		Dim ShowTX
		           If ShowStyle = "0" Then ShowTX = ("<option value=""0"" selected>无</option>") Else	ShowTX = ShowTX & ("<option value=""1"">无</option>")
				   If ShowStyle = "1" Then ShowTX = ShowTX & ("<option value=""1"" selected>首行加红加粗</option>") Else ShowTX = ShowTX & ("<option value=""2"">首行加红加粗</option>")
				   If ShowStyle = "2" Then ShowTX = ShowTX & ("<option value=""2"" selected>隔行加红显示</option>") Else	ShowTX = ShowTX & ("<option value=""2"">隔行加红显示</option>")
		
		
		.echo  ShowTX
		.echo "                  </select></td>"
		.echo "            </tr>"
		
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"">排序方法"
		.echo "                <select style=""width:70%;"" class='textbox' name=""InfoSort"">"
					If InfoSort = "ID Desc" Then
					 .echo ("<option value='ID Desc' selected>按id号降序</option>")
					Else
					 .echo ("<option value='ID Desc'>按id号降序</option>")
					End If
					If InfoSort = "Hits Desc" Then
					 .echo ("<option value='Hits Desc' selected>按点击数降序(热门)</option>")
					Else
					 .echo ("<option value='Hits Desc'>按点击数降序(热门)</option>")
					End If
					If InfoSort = "RefreshTime Desc" Then
					  .echo ("<option value='RefreshTime Desc' selected>按刷新时间降序(最新)</option>")
					Else
					  .echo ("<option value='RefreshTime Desc'>按刷新时间降序(最新)</option>")
					End If

		.echo "         </select></td>"
		.echo "              <td height=""24"">" & ReturnOpenTypeStr(OpenType) & "</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">信息数量"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num2""    style=""width:70%;"" onBlur=""CheckNumber(this,'信息数量');"" value=""" & Num & """></td>"
		.echo "              <td width=""50%"" height=""24"">排列列数"
		 .echo "               <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');""  style=""width:70%;"" value=""" & ColNumber & """ name=""ColNumber"">"
		.echo "            </tr>"

		
		
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">导航类型"
		.echo "                <select name=""NavType"" style=""width:70%;"" class='textbox' onchange=""SetNavStatus()"">"
				   If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>文字导航</option>")
					.echo ("<option value=""1"">图片导航</option>")
				   Else
					.echo ("<option value=""0"">文字导航</option>")
					.echo ("<option value=""1"" selected>图片导航</option>")
				   End If
		 .echo "               </select></td>"
		 .echo "             <td width=""50%"" height=""24"">"
			   If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('include/SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,$('#NaviPic')[0]);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:$('#NaviPic').val('');"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('include/SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,$('#NaviPic')[0]);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:$('#NaviPic').val('');"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
		 .echo "             </td>"
		 .echo "           </tr>"
		 
		 
		 		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"" colspan=""2"">显示条件："
		If JobType="0" Then
		.echo "                &nbsp;&nbsp;<input type='radio' value='0' name='JobType' checked>不限"
		Else
		.echo "                &nbsp;&nbsp;<input type='radio' value='0' name='JobType'>不限"
		End iF
		If JobType="1" Then
		.echo "                &nbsp;&nbsp;<input type='radio' value='1' name='JobType' checked>全职"
		Else
		.echo "                &nbsp;&nbsp;<input type='radio' value='1' name='JobType'>全职"
		End iF
		If JobType="2" Then
		.echo "                &nbsp;&nbsp;<input type='radio' value='2' name='JobType' checked>兼职"
		Else
		.echo "                &nbsp;&nbsp;<input type='radio' value='2' name='JobType'>兼职"
		End iF
		If JobType="3" Then
		.echo "                &nbsp;&nbsp;<input type='radio' value='3' name='JobType' checked>猎头职位"
		Else
		.echo "                &nbsp;&nbsp;<input type='radio' value='3' name='JobType'>猎头职位"
		End iF
						
        .echo "                <br><br>特殊属性："
		.echo "                &nbsp;&nbsp;<input name=""RecommendTF"" id=""RecommendTF"" type=""checkbox"" value=""true"""
		If RecommendTF = true Then .echo (" Checked")
		.echo ">仅显示推荐职位"
			
		.echo "                &nbsp;&nbsp;<input name=""JiPin"" id=""JiPin"" type=""checkbox"" value=""true"""
		If JiPin = true Then .echo (" Checked")
		.echo ">仅显示急聘职位"	
                      .echo "&nbsp;&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowJiPinFlag) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowJiPinFlag"" name=""ShowNewFlag"" checked>显示急聘标志")
					 Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowJiPinFlag"" name=""ShowNewFlag"">显示急聘标志")
					 End If
					  
		.echo "                </td>"
		.echo "            </tr>"

		.echo "                  </table>"			 
		.echo "    </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 

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
Set KSCls = New GetJobList
KSCls.Kesion()
Set KSCls = Nothing

Class GetJobList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag,ShowJiPinFlag
		Dim ClassID, OpenType, RecommendTF, ShowPin,ShowNewFlag,Num, ZWLen, TitleLen, InfoSort, ColNumber, Province, NavType, Navi, DateRule, DateAlign, TitleCss, City,County,ShowStyle, PrintType,AjaxOut,LabelStyle,Page
		FolderID = Request("FolderID")
		Page=Request("Page")
		CurrPath = KS.GetCommonUpFilesDir()

		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  ClassID = "0":DateRule="YYYY-MM-DD"
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetJobList", ""),"}" & LabelStyle & "{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			ClassID = Node.getAttribute("classid")
			ShowStyle=Node.getAttribute("showstyle")
			Province=Node.getAttribute("province")
			City=Node.getAttribute("city")
			County=Node.getAttribute("county")
			RecommendTF=Cbool(Node.getAttribute("recommendtf"))
			OpenType = Node.getAttribute("opentype")
			Num = Node.getAttribute("num")
			ZWLen = Node.getAttribute("zwlen")
			TitleLen = Node.getAttribute("titlelen")
			InfoSort = Node.getAttribute("infosort")
			ColNumber = Node.getAttribute("col")
			ShowPin= Node.getAttribute("showpin")
			ShowNewFlag= Node.getAttribute("shownewflag")
			ShowJiPinFlag=Node.getAttribute("showjipinflag")
			NavType = Node.getAttribute("navtype")
			Navi = Node.getAttribute("nav")
			DateRule = Node.getAttribute("daterule")
			TitleCss = Node.getAttribute("titlecss")
			PrintType= Node.getAttribute("printtype")
			AjaxOut  = Node.getAttribute("ajaxout")
		   End If
		   Set Node=Nothing
		   XMLDoc=Empty
		End If
		If PrintType="" Then PrintType=1
		If Num = "" Then Num = 20
		If ZWLen = "" Then ZWLen = 30
		If TitleLen = "" Then TitleLen = 30
		If ColNumber = "" Then ColNumber = 1
		If RecommendTF="" Then RecommendTF=False
		If ShowStyle="" Then ShowStyle=2
		If AjaxOut="" Then AjaxOUT=false
		If KS.IsNUL(ShowJiPinFlag) Then ShowJiPinFlag=false
		If LabelStyle="" Then LabelStyle="<li><a href=""{@jobcompanyurl}"">{@companyname}</a></li>"
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
		<script>
		$(document).ready(function(){
		 ChangeOutArea($("#PrintType").val());
		});
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
		

		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  top.$.dialog.alert('请输入标签名称!',function(){
			  $("input[name=LabelName]").focus();}); 
			  return false
			  }
			var ClassID=$('#ClassID').val();
			var ShowStyle=$('#ShowStyle').val();
			var NavType=1;
			var ShowNewFlag;
			var Province=$('#Province').val();
			var City=$('#City').val();
			var County=$('#County').val();
			var OpenType=$('#OpenType').val();
			var Num=$('#Num').val();
			var ZWLen=$('input[name=ZWLen]').val();
			var TitleLen=$('input[name=TitleLen]').val();
			var InfoSort=$('select[name=InfoSort]').val();
			var ColNumber=$('input[name=ColNumber]').val();
			var Nav,NavType=$('select[name=NavType]').val();
			var DateRule=$('#DateRule').val();
			var TitleCss=$('input[name=TitleCss]').val();
			var PrintType=$('#PrintType').val();
            var AjaxOut=false;
			if ($("#AjaxOut").prop("checked")==true){AjaxOut=true}
			var RecommendTF=false;
			if ($("#RecommendTF").prop("checked")==true){RecommendTF=true}
            var ShowPin=false;
			if ($("#ShowPin").prop("checked")==true){ShowPin=true}
            var ShowNewFlag=false;
			if ($("#ShowNewFlag").prop("checked")==true){ShowNewFlag=true}
	        var ShowJiPinFlag=false;
			if ($("#ShowJiPinFlag").prop("checked")==true){ShowJiPinFlag=true}
			if  (Num=='')  Num=10;
			if (ZWLen=='') ZWLen=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav=$('#TxtNavi').val()
			 else  Nav=$('#NaviPic').val();
			 
            var tagVal='{Tag:GetJobList labelid="0" classid="'+ClassID+'" showstyle="'+ShowStyle+'" province="'+Province+'" city="'+City+'" county="'+County+'" recommendtf="'+RecommendTF+'" opentype="'+OpenType+'" num="'+Num+'" zwlen="'+ZWLen+'" titlelen="'+TitleLen+'" infosort="'+InfoSort+'" col="'+ColNumber+'" showpin="'+ShowPin+'" shownewflag="'+ShowNewFlag+'" showjipinflag="'+ShowJiPinFlag+'" navtype="'+NavType+'" nav="'+Nav+'" titlecss="'+TitleCss+'" daterule="'+DateRule+'" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetJobList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通输出(Table)</option>"
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
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@id}')"">ID</li><li onclick=""InsertLabel('{@jobcompanyurl}')"">公司URL</li> <li onclick=""InsertLabel('{@companyname}')"">公司名称</li><li onclick=""InsertLabel('{@companyshortname}')"">公司短名称</li><li onclick=""InsertLabel('{@logo}')"">公司Logo</li> <li onclick=""InsertLabel('{@jobzwlist}')""><font color=red>职位列表</font></li><li onclick=""InsertLabel('{@province}')"">公司省份</li><li onclick=""InsertLabel('{@city}')"">城市</li><li onclick=""InsertLabel('{@county}')"">城镇</li><li onclick=""InsertLabel('{@contactman}')"">联系人</li><li onclick=""InsertLabel('{@tel}')"">联系电话</li> <li onclick=""InsertLabel('{@email}')"">公司Email</li><li onclick=""InsertLabel('{@fax}')"">公司传真</li><li onclick=""InsertLabel('{@joindate}')"">加入时间</li><li onclick=""InsertLabel('{@newimg}')"" title='显示新信息图片标志' style='color:red;'>最新图标志</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"		
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">公司行业"
		.echo "                <select class=""textbox"" style=""width:70%;"" name=""ClassID"" id=""ClassID"">"
						
						If ClassID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定行业 -</option>")
						Else
						  .echo ("<option  value=""0"">- 不指定行业 -</option>")
					   End If
					   
					   Dim SQL,I
					   Dim IRS:Set IRS=Conn.Execute("Select id,classname from ks_enterpriseclass where parentid=0")
					  If Not IRS.Eof Then SQL=IRS.GetRows(-1)
					  IRS.Close:Set IRS=Nothing
					  If IsArray(SQL) Then
					  For I=0 To Ubound(SQL,2)
					   if trim(ClassID)=trim(sql(0,i)) then
					   .echo "<option value='" & sql(0,i) & "' selected>" & sql(1,i) &"</option>"
					   else
					   .echo "<option value='" & sql(0,i) & "'>" & sql(1,i) &"</option>"
					   end if
					  Next
					  End If
					   
					   
						  .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "                <input name=""RecommendTF"" id=""RecommendTF"" type=""checkbox"" value=""true"""
		If RecommendTF = true Then .echo (" Checked")
		.echo ">仅显示推荐公司"				  
		.echo "                </td>"
		.echo "            </tr>"
		
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">公司地区&nbsp;"
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
		.echo "              <td width=""50%"" height=""24"">显示样式"
		.echo "                <select class='textbox' name=""ShowStyle"" id=""ShowStyle"" style=""width:70%;"">"
		Dim StyleStr
		           If ShowStyle = "1" Then StyleStr = ("<option value=""1"" selected>①:仅显示公司名称</option>") Else	StyleStr = StyleStr & ("<option value=""1"">①:仅显示公司名称</option>")
				   If ShowStyle = "2" Then StyleStr = StyleStr & ("<option value=""2"" selected>②:上名称+下职位</option>") Else StyleStr = StyleStr & ("<option value=""2"">②:上名称+下职位</option>")
				   If ShowStyle = "3" Then StyleStr = StyleStr & ("<option value=""3"" selected>③:左名称+右职位</option>") Else	StyleStr = StyleStr & ("<option value=""3"">③:左名称+右职位</option>")
				   If ShowStyle = "4" Then StyleStr = StyleStr & ("<option value=""4"" selected>④:上职位+下名称</option>") Else	StyleStr = StyleStr & ("<option value=""4"">④:上职位+下名称</option>")
		
		
		.echo  StyleStr
		.echo "                  </select></td>"
		.echo "            </tr>"
		
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"">排序方法"
		.echo "                <select style=""width:70%;"" class='textbox' name=""InfoSort"">"
					If InfoSort = "A.ID Desc" Then
					 .echo ("<option value='A.ID Desc' selected>最新加盟企业</option>")
					Else
					 .echo ("<option value='A.ID Desc'>最新加盟企业</option>")
					End If
					If InfoSort = "RefreshTime Desc" Then
					 .echo ("<option value='RefreshTime Desc' selected>最新更新时间</option>")
					Else
					 .echo ("<option value='RefreshTime Desc'>最新更新时间</option>")
					End If
					If InfoSort = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>职位点击数(热门职位)</option>")
					Else
					  .echo ("<option value='Hits Desc'>职位点击数(热门职位)</option>")
					End If

		.echo "         </select></td>"
		.echo "              <td height=""24"">" & ReturnOpenTypeStr(OpenType) & "</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">信息数量"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:70%;"" onBlur=""CheckNumber(this,'信息数量');"" value=""" & Num & """></td>"
		.echo "              <td width=""50%"" height=""24"">排列列数"
		 .echo "               <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');""  style=""width:70%;"" value=""" & ColNumber & """ name=""ColNumber"">"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">公司长度"
		.echo "                <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """><br><font color=blue>显示公司名称的字符数，一个汉字算两个字符</font>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">职位长度"
		.echo "                <input name=""ZWLen"" class=""textbox"" type=""text"" id=""ZWLen2""    style=""width:70%;"" onBlur=""CheckNumber(this,'职位字数');"" value=""" & ZWLen & """><br><font color=blue>当公司发布多个职位时，显示的总字数将不超过这里设置的长度</font></td>"
		 .echo "              </td>"
		 .echo "           </tr>"
		
		.echo "           <tr class=tdbg>"
		 .echo "             <td colspan=2 height=""30"">附加显示 "
				   If cbool(ShowPin) = True Then
					  .echo ("&nbsp;&nbsp;&nbsp;<input type=""checkbox"" value=""true"" id=""ShowPin"" name=""ShowPin"" checked>显示“聘”字")
				   Else
					  .echo ("&nbsp;&nbsp;&nbsp;<input type=""checkbox"" value=""true"" id=""ShowPin"" name=""ShowPin"">显示“聘”字")
				   End If
                      .echo "&nbsp;&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowNewFlag) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"" checked>最新标志")
					 Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"">最新标志")
					 End If
                      .echo "&nbsp;&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowJiPinFlag) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowJiPinFlag"" name=""ShowNewFlag"" checked>急聘标志")
					 Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowJiPinFlag"" name=""ShowNewFlag"">急聘标志")
					 End If
				 
		.echo "       　</td>"
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
		

		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">日期格式"
		.echo ReturnDateFormat(DateRule)
		.echo "              </td>"
		.echo "              <td height=""24"">"
		.echo "                <div align=""left"">标题样式<input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """>"
		.echo "                </div></td>"
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

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
Dim KSCls
Set KSCls = New JS_Main
KSCls.Kesion()
Set KSCls = Nothing

Class JS_Main
        Private KS
		'========================================================================
		Private JSSql, JSRS, FolderID, JSID, ChannelID, Channel, Action
		Private i, totalPut, CurrentPage, JSType
		Private KeyWord, SearchType, StartDate, EndDate
		'搜索参数集合
		Private SearchParam
		Private MaxPerPage
		Private Row 
		'========================================================================
		Private Sub Class_Initialize()
		  MaxPerPage = 96
		  Row = 8
		  Set KS=New PublicCls
		   Call KS.DelCahe(KS.SiteSn & "_ReplaceFreeLabel")
		   Call KS.DelCahe(KS.SiteSn & "_jslist")
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		'采集搜索信息
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate = KS.G("EndDate")
		SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate
		JSType = KS.G("JSType"):If JSType = "" Then JSType = 0
		
		Select Case KS.G("Action")
		 Case "JSDel"
		   Call JSDel()
		 Case "JSFolderDel"
		   Call JSFolderDel()
		 Case "JSView"
		   Call JSView()
		 Case Else
		   Call JSMainList()
		End Select
		End Sub
		
		Sub JSMainList()
		   With Response
			If JSType = 0 Then
				If Not KS.ReturnPowerResult(0, "KMTL10004") Then                '系统JS管理的权限检查
				  Call KS.ReturnErr(1, "")
				  .End
				End If
			ElseIf JSType = 1 Then
				If Not KS.ReturnPowerResult(0, "KMTL10005") Then                '自由JS管理的权限检查
				  Call KS.ReturnErr(1, "")
				  .End
				End If
			End If
			
			If Not IsEmpty(KS.G("page")) And KS.G("page") <> "" Then
				  CurrentPage = CInt(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			Action = KS.G("Action")
			FolderID = Trim(KS.G("FolderID"))
			If FolderID = "" Then FolderID = "0"
			Dim UPFolderRS, ParentID
			Set UPFolderRS = Conn.Execute("select * from [KS_LabelFolder] where  ID ='" & FolderID & "'")
			If Not UPFolderRS.EOF Then
			 ParentID = UPFolderRS("ParentID")
			End If
			UPFolderRS.Close:Set UPFolderRS = Nothing
		    .Write "<!DOCTYPE html><html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<title>JS列表</title>"
			.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
			.Write "<script language=""JavaScript"">"
			.Write "var FolderID='" & FolderID & "';         //目录ID" & vbCrLf
			.Write "var ParentID='" & ParentID & "'; //父栏目ID" & vbCrLf
			.Write "var Page='" & CurrentPage & "';   //当前页码" & vbCrLf
			.Write "var KeyWord='" & KeyWord & "';    //关键字" & vbCrLf
			.Write "var SearchParam='" & SearchParam & "';  //搜索参数集合" & vbCrLf
			.Write "var Action='" & Action & "';" & vbCrLf
			.Write "var JSID='" & JSID & "';" & vbCrLf
			.Write "var JSType=" & JSType & ";" & vbCrLf
			.Write "</script>" & vbCrLf
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/jQuery.js""></script>"
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
			%>
			<script language="javascript">
			function ChangeUp()
			{
			 if (FolderID=='0') return;
			 location.href='JS_Main.asp?JSType='+JSType+'&FolderID='+ParentID;
			   if (JSType==0)
				  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS 管理 >> 系统 JS&ButtonSymbol=SysJSList&LabelFolderID='+ParentID;
			   else
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS 管理 >> 自由 JS&ButtonSymbol=FreeJSList&LabelFolderID='+ParentID;
			 }
			function OpenFolder(FolderID)
			{
			 location.href='JS_Main.asp?JSType='+JSType+'&FolderID='+FolderID;
			   if (JSType==0)
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS 管理 >> 系统 JS&ButtonSymbol=SysJSList&LabelFolderID='+FolderID;
			   else
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS 管理 >> 自由 JS&ButtonSymbol=FreeJSList&LabelFolderID='+FolderID;
			}
			function CreateFolder()
			{ 
			  if (JSType==0)
			    top.openWin("新建系统JS目录","include/LabelFolder.asp?LabelType=2&FolderID="+FolderID,true,650,360);
			  else
			    top.openWin("新建自由JS目录","include/LabelFolder.asp?LabelType=3&FolderID="+FolderID,true,650,360);
			}
			function AddJS(TempUrl)
			{
			  if (JSType==0)
				{
				 location.href=TempUrl+'JS/AddSysJS.asp?FolderID='+FolderID+'&JSType="'+JSType+'&Action=AddNew';
				  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS 管理 >> <font color=red>添加系统 JS</font>&ButtonSymbol=JSAdd';
				 }
			  else
				{location.href=TempUrl+'JS/AddFreeJS.asp?FolderID='+FolderID+'&Action='+Action+'&JSID='+JSID+'&JSType='+JSType
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS 管理 >> <font color=red>添加自由 JS</font>&ButtonSymbol=JSAdd';
				 }
			}
			function EditJS(TempUrl,ID)
			{  if (KeyWord=='')
				{   if (JSType==0)
					  {
					   location.href=TempUrl+'EditJS.asp?Page='+Page+'&JSType='+JSType+'&Action=Edit&JSID='+ID;
					   $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS管理 >> <font color=red>修改系统JS</font>&ButtonSymbol=JSEdit';
					  }
				   else
				   {
					 location.href=TempUrl+'EditJS.asp?Page='+Page+'&JSType=1&Action=Edit&JSID='+ID;
					 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS管理 >> <font color=red>修改自由JS</font>&ButtonSymbol=JSEdit';
					}
				}
			   else
				 {  if (JSType==0)
					 {
					  location.href=TempUrl+'EditJS.asp?'+SearchParam+'&Page='+Page+'&JSType='+JSType+'&Action=Edit&JSID='+ID;
					  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS管理 >> 搜索系统JS结果 >><font color=red>修改系统JS</font>&ButtonSymbol=JSEdit';
					  }
				   else
					{
					 location.href=TempUrl+'EditJS.asp?'+SearchParam+'&Page='+Page+'&JSType=1&Action=Edit&JSID='+ID;
					 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>Post.Asp?OpStr=JS管理 >> 搜索自由JS结果 >> <font color=red>修改自由JS</font>&ButtonSymbol=JSEdit';
					 }
			  }
			}
			function EditFolder(ID)
			{
			 top.openWin("编辑标签目录","include/LabelFolder.asp?Action=EditFolder&FolderID="+ID,true,650,360);
			}
			function Edit(TempUrl)
			{   
			var ids=get_Ids(document.myform);
			if (ids!=''){
			      if (ids.indexOf(',')==-1) {
				       var ltype=$("#c"+ids).attr("ltype");
					   if (ltype==1)
					   EditJS(TempUrl,ids);
					    else
						 EditFolder(ids);
					   }
					else alert('一次只能够编辑一个JS标签或目录');
			}
			else 
			{
			alert('请选择要编辑的JS标签');
			}
			
			}
			
			//批量删除标签
		function DeleteLabel(){
		 if (chk_idBatch(myform,'此操作不可逆,确定删除选中的JS标签吗')==true)
		   {
		    $('#Action').val('JSDel'); 
			$('#myform').submit();
		   }
		}//批量删除标签目录
		function DeleteLabelFolder(){
		 if (chk_idBatch(myform,'此操作不可逆,确定删除选中的JS标签目录吗')==true)
		   {
		    $('#Action').val('JSFolderDel'); 
			$('#myform').submit();
		   }
		}
			function Delete(TempUrl)
			{ 
			var ids=get_Ids(document.myform);
			if (ids!=''){
					if (confirm('删除确认:\n\n真的要执行删除操作吗?')){ 
						$('#Action').val('JSDel'); 
			            $('#myform').submit();
					}	
				}
			else alert('请选择要删除的JS标签');
			}
			function DelFolder(ID){
		    if (confirm('删除确认:\n\n真的要执行删除JS目录操作吗?')){
			location='JS_Main.asp?JSType='+JSType+'&Action=JSFolderDel&ID='+ID+'&FolderID='+FolderID;
			}
		}
		function DelJS(ID){
		    if (confirm('删除确认:\n\n真的要执行删除JS操作吗?')){
			location='JS_Main.asp?JSType='+JSType+'&FolderID='+FolderID+'&Action=JSDel&Page='+Page+'&ID='+ID;
			}
		}
			
			function GetKeyDown()
			{
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  { 
				 case 78 : event.keyCode=0;event.returnValue=false; CreateFolder();break;
				 case 77 : event.keyCode=0;event.returnValue=false; AddJS('');break;
				 case 65 : SelectAllElement();break;
				 case 66 : event.keyCode=0;event.returnValue=false;ChangeUp();break;
				 case 69 : event.keyCode=0;event.returnValue=false;Edit('');break;
				 case 68 : Delete('');break;
				 case 86 : JSView();break;
				 case 70 : event.keyCode=0;event.returnValue=false;
				   if (JSType==0)
					parent.initializeSearch('SysJS')
				   else
					parent.initializeSearch('FreeJS')
			 }	
			else if (event.keyCode==46)
			Delete('');
			}

			function JSView(id)
			{  
			    if (id!=''){
				 top.openWin('预览JS显示效果','include/JS_Main.asp?Action=JSView&JSID='+id,false);
				 }
				else
				 alert('请选择您要预览的JS!')
			}

			</script>
			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0""  onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		    .Write "<ul id='menu_top'>"
				 If KeyWord = "" Then
			.Write "<li class='parent' onclick=""AddJS('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add1'></i>添加JS</span></li>"
			.Write "<li class='parent' onclick=""CreateFolder();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加目录</span></li>"
			 .Write "<li class='parent' onclick=""ChangeUp();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>回上一级</span></li>"
				 
				 Else
					  If JSType = 0 Then
					   .Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('JS_Main.asp?JSType=0','Template_Left.asp','../Post.Asp?ButtonSymbol=SysJSList&OpStr=JS管理 >> <font color=red>系统JS管理</font>')"">系统JS首页</span>")
					Else
					   .Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('JS_Main.asp?JSType=1','Template_Left.asp','../Post.Asp?ButtonSymbol=FreeJSList&OpStr=JS管理 >> <font color=red>自由JS管理</font>')"">自由JS首页</span>")
					End If
				   .Write (">>> 搜索结果: ")
					 If StartDate <> "" And EndDate <> "" Then
						.Write ("JS更新日期在 <font color=red>" & StartDate & "</font> 至 <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
					 End If
					Select Case KS.ChkClng(SearchType)
					 Case 0
					  .Write ("名称含有 <font color=red>" & KeyWord & "</font> 的JS")
					 Case 1
					  .Write ("描述中含有 <font color=red>" & KeyWord & "</font> 的JS")
					 Case 2
					  .Write ("文件名中含有 <font color=red>" & KeyWord & "</font> 的JS")
					 End Select
			End If
			
			.Write "    </ul>"

		
		
		Response.Write "<div class='tableTop'><table><tr><td height=""30"" colspan=""3"" style=""text-align:left;"">&nbsp;<b>选择：</b><a href='javascript:void(0)' onclick='javascript:Select(0)'>全选</a>  <a href='javascript:void(0)' onclick='javascript:Select(1)'>反选</a>  <a href='javascript:void(0)' onclick='javascript:Select(2)'>不选</a> <input type='button' value='批量删除选中的JS标签' class='button button2' onclick=""DeleteLabel()""/> <input type='button' value='批量删除选中的JS标签目录' class='button button2' onclick=""DeleteLabelFolder()""/> </td><td colspan=""10""> <form name='searchform' action='JS_Main.asp' method='post'><input type='hidden' name='JSType' value='" & JSType & "'/><strong>搜索标签=》</strong>JS标签名称：<input type='text' name='keyword' class='textbox'/> <input type='submit' value=' 搜索 ' class='button'/></form> </td></tr></table></div>"
		Response.Write "<div class='pageCont2 mt20'>"
		Response.Write "  <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "      <tr class='sort'>"
		Response.Write "       <td width=""40"" align=""center"">选中</td>"
		Response.Write "       <td align=""center"">JS标签名称</td>"
		Response.Write "       <td align=""center"">类型</td>"
		Response.Write "       <td align=""center"">JS文件名</td>"
		Response.Write "       <td align=""center"">更新时间</td>"
		Response.Write "       <td align=""center"">操作</td>"
		Response.Write "      </tr>"
		Response.Write "<form action='JS_Main.asp' id='myform' name='myform' method='post'>"
		Response.Write "<input type='hidden' name='JSType' value='" & JSType & "'/>"
		Response.Write "<input type='hidden' name='FolderID' value='" & FolderID & "'/>"
		Response.Write "<input type='hidden' name='Action' id='Action' value='del'/>"
					Dim FolderSql, Param
					 Param = " Where JsType=" & JSType
					If KeyWord <> "" Then
					   FolderSql = "SELECT ID,FolderName,Description,OrderID,AddDate,'---' as JSFileName FROM [KS_LabelFolder] Where 1=0"
					  Select Case KS.ChkClng(SearchType)
						Case 0
						  Param = Param & " AND JSName like '%" & KeyWord & "%'"
						Case 1
						 Param = Param & " AND Description like '%" & KeyWord & "%'"
						Case 2
						 Param = Param & " AND JSFileName like '%" & KeyWord & "%'"
					  End Select
					  If StartDate <> "" And EndDate <> "" Then
						If CInt(DataBaseType) = 1 Then         'Sql
						   Param = Param & " And (AddDate>= '" & StartDate & "' And AddDate<= '" & DateAdd("d", 1, EndDate) & "')"
						Else                                                 'Access
						   Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
						End If
					  End If
					Else
					   Param = Param & " AND FolderID='" & FolderID & "'"
					   FolderSql = "SELECT ID,FolderName,Description,OrderID,AddDate,'---' AS JSFileName FROM [KS_LabelFolder] Where FolderType=" & JSType + 2 & " And ParentID='" & FolderID & "'"
					End If
					Param = Param & " ORDER BY OrderID"
			Set JSRS = Server.CreateObject("ADODB.recordset")
			JSRS.Open FolderSql & " UNION  Select JSID,JSName,Description,OrderID,AddDate,JSFileName From KS_JSFile " & Param, Conn, 1, 1
			If JSRS.EOF And JSRS.BOF Then
			 Response.Write "<tr class='tdbg'><td class='splittd' colspan='10' style='text-align:center'>没有找到记录！</td></tr>" 
			 Else
						       totalPut = JSRS.RecordCount
								If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										JSRS.Move (CurrentPage - 1) * MaxPerPage
								End If
								Call showContent
							   
			End If
		 Response.Write "</form><tr><td height=""40"" colspan=""3"" style=""text-align:left;"" class='operatingBox'>&nbsp;<b>选择：</b><a href='javascript:void(0)' onclick='javascript:Select(0)'>全选</a>  <a href='javascript:void(0)' onclick='javascript:Select(1)'>反选</a>  <a href='javascript:void(0)' onclick='javascript:Select(2)'>不选</a> <input type='button' value='批量删除选中的JS标签' class='button' onclick=""DeleteLabel()""/> <input type='button' value='批量删除选中的JS标签目录' class='button' onclick=""DeleteLabelFolder()""/> </td><td colspan=""10""> <form name='searchform' action='JS_Main.asp' method='post'><input type='hidden' name='JSType' value='" & JSType & "'/><strong>搜索标签=》</strong>JS标签名称：<input type='text' name='keyword' class='textbox'/> <input type='submit' value=' 搜索 ' class='button'/></form> </td></tr>"
		  Response.Write " <tr><td  align=""right"" colspan=""10"" class='operatingBox'>"
			  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		  Response.Write " </td>"
		  Response.Write "    </tr>"

			.Write "  </table>"
			.Write "  </div>"
			.Write "  </div>"
			.Write "  </body>"
			.Write "  </html>"
			
			Set JSRS = Nothing
			Set Conn = Nothing
			End With
			End Sub
			
			   Sub showContent() 
			    Dim i:i=0
				Do While Not JSRS.EOF
				   Response.Write "<tr class='list' id='u" & JSRS("id") &"' onclick=""chk_iddiv('" & JSRS("id") &"')"" onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
				   Response.Write " <td class='splittd' style='text-align:center'><input ltype='" & JSRS(3) &"' type='checkbox' onclick=""chk_iddiv('" & JSRS("id") &"')"" id='c" & JSRS("id") &"'   name='id' value='" & JSRS("id") &"'/></td>"
				   Response.Write " <td class='splittd' style='text-align:left'>"
							  If JSRS(3) = 0 Then
								 Response.Write ("<i class=""icon folder""></i> ")
							 Else
								 Response.Write ("<img src=""../Images/Label/JS" & JSType & ".gif"" align=""absmiddle"">")
							 End If
				If JSRS(3) = 0 Then
				   Response.Write  "<a href=""javascript:;"" title=""进入该分类"" onclick=""OpenFolder('" & JSRS("id") &"')"">" &  JSRS(1) &"</a>"
				Else
				   Response.Write  "<a href=""javascript:;"" title=""调用：" & JSRS(1)& """ onclick=""EditJS('','" & JSRS("id") &"')"">" & Replace(Replace(JSRS(1), "{JS_", ""),"}","") &"</a>"
				End If
				   
				   Response.Write "</td>"
				   If JSRS(3) = 0 Then
					Response.Write " <td class='splittd tips' style='text-align:center'>JS标签目录</td>"
				   Else
					Response.Write " <td class='splittd tips' style='text-align:center'>JS标签</td>"
				   End If
				   Response.Write " <td class='splittd tips' style='text-align:center'>" & JSRS("JSFileName") & "</td>"
				   Response.Write " <td class='splittd' style='text-align:center'>" & JSRS("AddDate") & "</td>"
				   Response.Write " <td class='splittd' style='text-align:center'>"
				   
					If JSRS(3) = 0 Then
						response.write " <a href=""javascript:;"" onclick=""EditFolder('" & JSRS(0) & "')""  title=""修改JS标签目录"">修改</a> <a href=""javascript:;"" onclick=""DelFolder('" & JSRS(0) & "');"" title=""删除JS标签目录"">删除</a>"
					Else
						response.write " <a href=""javascript:EditJS('','" & JSRS(0) & "');"" title=""修改JS标签"">修改</a>"
						response.write " <a href=""javascript:DelJS('" & JSRS(0) & "');"" title=""删除JS标签"">删除</a>"
						response.write " <a href=""javascript:JSView('" & JSRS(0) & "');"" title=""预览JS标签"">预览</a>"
					 End If
				   
				   Response.Write " </td>"
				   Response.Write "</tr>"
				
						i = i + 1
						JSRS.MoveNext
						If i >= MaxPerPage Then Exit do
		          loop
			 JSRS.Close
					  
		End Sub
		
		'删除JS
		Sub JSDel()
		 Dim K, JSID, Page,RS,ArticleRS,JSType, CurrPath, JSFileName, JSDir, FolderID
		Set RS=Server.CreateObject("ADODB.Recordset")
		Set ArticleRS=Server.CreateObject("ADODB.Recordset")
		Page = Trim(KS.G("Page"))
		JSID = Split(KS.G("ID"), ",") '获得要删除标签的ID集合
		For K = LBound(JSID) To UBound(JSID)
		  RS.Open "SELECT * FROM [KS_JSFile] WHERE JSID='" & JSID(K) & "'", Conn, 1, 3
		  If Not RS.EOF Then
			JSType = RS("JSType")
			FolderID = RS("FolderID")
			  '删除物理JS文件
			  JSFileName = Trim(RS("JSFileName"))
			  JSDir = Trim(KS.Setting(93))
			  If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			  CurrPath = KS.Setting(3) & JSDir
			  Call KS.DeleteFile(CurrPath & JSFileName)
			  '从文章中删除此JSID
			  ArticleRS.Open "Select  JSID From KS_Article Where JSID like '%" & JSID(K) & "%'", Conn, 1, 3
			  If Not ArticleRS.EOF Then
				 While Not ArticleRS.EOF
					ArticleRS(0) = Replace(ArticleRS(0), JSID(K) & ",", "")
					ArticleRS.Update
					ArticleRS.MoveNext
				 Wend
			  End If
			  ArticleRS.Close
		  End If
		 RS.Close
		 Conn.Execute("Delete From KS_JSFile Where JSID='" & JSID(k) &"'")
		Next
		Set RS = Nothing:Set ArticleRS = Nothing
		 Call KS.Alert("恭喜，删除JS成功！","include/JS_Main.asp?FolderID=" & KS.S("FolderID")&"&JSType=" & KS.S("JsType")&"&page=" & KS.S("Page"))

		End Sub
		
		'删除JS目录
		Sub JSFolderDel()
		   Dim RS,K, ID, ParentID, FolderSql,LabelFolderID,LabelType
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   ID = Split(Request("ID"), ",")     '获得要删除目录的ID集合
			For K = LBound(ID) To UBound(ID)
			  FolderSql = "select ID,ParentID,FolderType from [KS_LabelFolder] where ID='" & ID(K) & "'"
			  RS.Open FolderSql, Conn, 1, 1
			  If Not RS.EOF Then
				LabelFolderID = Trim(RS(0))
				ParentID = Trim(RS(1))
				LabelType = RS(2)
						  Dim RSJS,JSDir
						  Set RSJS=Server.CreateObject("ADODB.Recordset")
						  '删除JS物理文件
						  RSJS.Open "Select JSFileName From KS_JSFile Where FolderID='" & LabelFolderID & "'", Conn, 1, 1
								 JSDir = Trim(KS.Setting(93))
								If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
						  Do While Not RSJS.EOF
								Call KS.DeleteFile(KS.Setting(3) & JSDir & RSJS(0))
								RSJS.MoveNext
						  Loop
						  RSJS.Close
						  Set RSJS = Nothing
						  Conn.Execute ("DELETE  FROM KS_JSFILE WHERE FolderID='" & LabelFolderID & "'")
						  Conn.Execute ("DELETE  FROM KS_LabelFolder WHERE ID='" & LabelFolderID & "' OR TS like '%" & LabelFolderID & "%'")
			   End If
			  RS.Close
			Next
		 Set RS = Nothing
		 Call KS.Alert("恭喜，删除JS标签目录成功！","include/JS_Main.asp?JSType=" & KS.S("JsType")&"&page=" & KS.S("Page"))
		End Sub
		
		'预览JS
		Sub JSView()
			Dim JSObj,JSID, JSdir,JSUrlStr
			JSID=Trim(Request.QueryString("JSID"))
			JSDir = KS.Setting(93)
			If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			Set JSObj=Server.CreateObject("Adodb.Recordset")
			JSObj.OPEN "Select JSConfig,JSType,JSFileName From KS_JSFile Where JSID='" & JSID & "'",Conn,1,1
			IF JSObj.EOf AND JSObj.BOF THEN
			  Response.Write("参数传递出错!")
			  JSObj.Close
			  Set JSObj=Nothing
			  Response.End
			ELSE
			  IF (trim(Split(JSObj("JSConfig"),",")(0))="GetExtJS" Or JSObj("JSType")=0) or (Request.QueryString("CanView")="1") Then
			  JSUrlStr="<script language=""javascript"" src=""" & KS.GetDomain & JSDir & Trim(JSObj("JSFileName")) & """></script>"
			  Else
				JSObj.Close:Set JSObj=Nothing
				Response.Redirect "JSFreeView.asp?JSID=" &JSID
			  End IF
			END IF
			JSObj.close:Set JSObj=Nothing
			%>
			<!DOCTYPE html><html>
			<head>
			<meta http-equiv="Expires" CONTENT="0">        
			<meta http-equiv="Cache-Control" CONTENT="no-cache">        
			<meta http-equiv="Pragma" CONTENT="no-cache">      
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<link href="Admin_Style.CSS" rel="stylesheet">
			<title>JS预览</title>
			<script language="JavaScript" src="Common.js"></script>
			</head>
			<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  bgcolor="#F1FAFA">
			<br>
			<table width="90%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
					<tr>
					  <td align="center" valign="top"><%=JSUrlStr%></td>
					</tr>
				  </table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
				<td height="25"><strong> 说明:</strong></td>
			  </tr>
			  <tr>
				<td height="25">１．如果JS中有设置样式,那么这里的预览效果可能会与实际有点差距</td>
			  </tr>
			  <tr> 
				<td height="25">２．如果看不到效果，请单击刷新按钮 <input class="button" type="button" value="刷新" onClick="window.location.reload()"> <input class="button"  type="button" value="关闭" onClick="top.box.close()">
				<%if Request.QueryString("CanView")="1" then
				  Response.Write(" <INPUT TYPE=BUTTON value=""返回"" class=""button""  onclick=""history.back();"">")
				  End IF
				  %></td>
			  </tr>
			</table>
			</body>
			</html>
      <%
		End Sub
End Class
%> 

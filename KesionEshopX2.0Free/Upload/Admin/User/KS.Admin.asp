<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Server.ScriptTimeOut=9999999
Dim KSCls
Set KSCls = New User_AdminMain
KSCls.Kesion()
Set KSCls = Nothing

Class User_AdminMain
        Private KS,UserName
		Private GroupID, I, SqlStr, RSObj,Title, CreateDate, TempStr, GRS,KeyWord, SearchType
		Private PowerRS,RS,AdminID,PowerList,SpecialPower,CollectPower,SystemPower,RefreshPower,UserAdminPower,KMTemplatePower,ModelPower

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 Select Case KS.G("Action")
		   Case "Add","Edit"
			 If Not KS.ReturnPowerResult(0, "KMUA10001") Then                '检查管理员操作(增和改)的权限检查
			  Call KS.ReturnErr(0, "")
			  Exit Sub
			 Else 
			  Call AdminAdd()
		     End If
		   Case "Del"
		       If Not KS.ReturnPowerResult(0, "KMUA10001") Then                '检查管理员操作(增和改)的权限检查
				  Call KS.ReturnErr(0, "")
				  Exit Sub
			   Else
			   Call AdminDel()
		       End If
		   Case "SetPass"
		   	 If Not KS.ReturnPowerResult(0, "KMUA10010") Then           
			  Call KS.ReturnErr(0, "")
			  Exit Sub
		     Else
		      Call SetAdminPass()
			 End If
		   Case "AddRole","EditRole" Call AddRole()
		   Case "AddRoleSave" Call AddRoleSave()
		   Case "Role" Call Role()
		   Case "DelRole" Call DeleteRole()
		   Case Else
		     Call AdminList()
		 End Select
		End Sub
		
		Sub Head()
		'收集搜索参数
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		'搜索参数集合
		Dim SearchParam:SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType
		Const Row = 8 '每行显示数
		Response.Write "<!DOCTYPE html><html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>管理员管理</title>"
		Response.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var GroupID='0';        //管理员组ID" & vbCrLf
		Response.Write "var KeyWord='" & KeyWord & "';         //搜索关键字" & vbCrLf
		Response.Write "var SearchParam='" & SearchParam & "'; //搜索参数集合" & vbCrLf
		Response.Write "</script>" & vbCrLf
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/jquery.js""></script>"
		%>
		<script language="javascript">
		function CreateAdmin()
		{
		 location.href='KS.Admin.asp?Action=Add';
		 $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Go&OpStr=用户管理 >> <font color=red>添加管理员</font>';
		}
		function EditAdmin(AdminID)
		{
		 location.href='KS.Admin.asp?Action=Edit&AdminID='+AdminID;
		 $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=用户管理 >> <font color=red>修改管理员</font>';
		}
		function EditRole(RoleID)
		{
		 location.href='KS.Admin.asp?Action=EditRole&ID='+RoleID;
		 $(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=用户管理 >> <font color=red>修改管理员角色</font>';
		}
		function Create()
		{
		 CreateAdmin();  
		}

		function Delete(id)
		{  if (confirm('确定删除选中管理员吗?')) location='KS.Admin.asp?'+SearchParam+'&Action=Del&AdminID='+id;
		}
		function DeleteRole(id){
		 if (confirm('删除角色，归属该角色的管理员将同时被删除，确定执行删除角色吗?')){ location.href='KS.Admin.asp?Action=DelRole&ID='+id;}
		}
		function SetAdminPassWord(id)
		{
		 	 top.openWin('设置管理员密码','user/KS.Admin.asp?Action=SetPass&AdminID='+id,false,520,265);
		}
		function CreateRole(){
			 location.href='KS.Admin.asp?Action=AddRole';
			 $(parent.document).find('#BottomFrame')[0].src='Post.Asp?Opstr='+escape("管理员管理 >> <font color=red>添加角色</font>")+'&ButtonSymbol=Go';
		}
		function GetKeyDown(){
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 :  Reload(); break;
			 case 78 : event.keyCode=0;event.returnValue=false;
				 CreateAdmin('');
			 case 80 :SetAdminPassWord();break;
			 case 70 : event.keyCode=0;event.returnValue=false;
				parent.initializeSearch('Manager')
		   }	
		else	
		 if (event.keyCode==46)Delete('');
		}
		function Reload()
		{
		location.reload();
		}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body topmargin=""0"" leftmargin=""0"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		Response.Write "<ul id='menu_top'>"
			  If KeyWord = "" Then
			   If KS.G("Action")="Role" Or KS.G("GroupID")<>"" Then
			   Response.Write "<li class='parent' onclick=""location.href='KS.Admin.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>管理员管理</span></li>"	
			   End If
			   Response.Write "<li class='parent' onclick=""CreateAdmin();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add3'></i>添加管理员</span></li>"	
			   
			   Response.Write "<li class='parent' onclick=""location.href='?action=Role'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon add2'></i>角色管理</span></li>"
			   Response.Write "<li class='parent' onclick='javascript:CreateRole();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon audit'></i>添加角色</span></li>"

			  Else
				   Response.Write ("<i class='icon mainer'></i><span style='cursor:pointer'>管理员首页</span>")
			   Response.Write (">>> 搜索结果: ")
				Select Case SearchType
				 Case 0
				  Response.Write ("用户名含有 <font color=red>" & KeyWord & "</font> 的管理员")
				 Case 1
				  Response.Write ("管理员简介含有 <font color=red>" & KeyWord & "</font> 的管理员")
				 End Select
			   End If

		Response.Write "</ul>"
		End Sub
		
		Sub Role()
		 Head
		 Response.Write"<div class=""pageCont2"">"
		 Response.Write("<div class='tabTitle'>管理员角色</div>")
		 Response.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		 Response.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""1"">"
		 				 Response.Write ("<tr> ")
				 Response.Write ("  <td>")
				 Response.Write ("    <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
				 Response.Write ("<tr align=""center""><td height=23 class=""sort"">角色名称</td><td height=23 class=""sort"">类型</td><td height=23 class=""sort"">审核</td><td class=""sort"">简介</td><td class=""sort"">管理员数</td><td class='sort'>管理操作</td></tr>")
		Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		 RSObj.Open "select * from ks_usergroup where [type]>1 order by id", Conn, 1, 1
		 Dim T, TitleStr, LockStr, ShortName
		Do While Not RSObj.EOF
			
			  Response.Write ("<tr><td class='splittd' height=25>&nbsp;<i class='icon manage'></i><span style=""cursor:default"">" & RSObj("GroupName") & "<span></td>")
			  Response.Write ("<td class='splittd' align=""center"">")
			  if rsobj("type")="3" then
			   response.write "<span style='color:red'>超级管理员</span>"
			  else
			   response.write "<span style='color:green'>普通管理员</span>"
			  end if
			  Response.Write ("</td>")
			  Response.Write ("<td class='splittd' align=""center"">")
			  if rsobj("type")="3" then
			   response.write "---"
			  Elseif rsobj("role")="3" then
			   response.write "<span style='color:#999'>终审员</span>"
			  elseif rsobj("role")="2" then
			   response.write "<span style='color:green'>初审员</span>"
			  else
			   response.write "<span style='color:blue'>发稿员</span>"
			  end if
			  Response.Write ("</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("descript") & "</td>")
			  Response.Write ("<td class='splittd' align=""center""><a href='KS.Admin.asp?groupid=" & RSObj("id") & "'>" & conn.execute("select count(1) from ks_admin where groupid=" & rsobj("id"))(0) & " 位</a></td>")
			  Response.Write ("<td class='splittd' align=""center""><a href='javascript:EditRole(" & rsobj("ID") &")' class='setA'>修改</a>|")
			  if rsobj("type")="3" then
			  response.write "<font color='#999999' class='setA'>删除</font>"
			  else
			  response.write "<a href='javascript:DeleteRole("&rsobj("ID")&")' class='setA'>删除</a>"
			  end if
			  Response.Write ("</td></tr>")
			  RSObj.MoveNext
			 If RSObj.EOF Then Exit Do
			Loop
			RSObj.Close:Conn.Close:Set RSObj = Nothing:Set GRS = Nothing
		  
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub AddRole()
		Dim SQL,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
		RSC.Open "Select ChannelID,ChannelName,BasicType,ItemName,ModelEname,ChannelStatus From KS_Channel where channelstatus=1 And BasicType<11 Order By ChannelID",Conn,1,1
		If Not RSC.Eof Then
		  SQL=RSC.GetRows(-1)
		End If
		RSC.Close:Set RSC=Nothing
		 Dim GroupName,Descript,Id,RS,STitle,RoleType,PowerList,ModelPower,Role,ManageOtherDoc
		 ID=KS.ChkCLng(Request("ID"))
		 If Id<>0 Then
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 * From KS_UserGroup Where ID=" & ID,conn,1,1
		  If Not RS.Eof Then
		    GroupName = RS("GroupName")
			Descript  = RS("Descript")
			RoleType  = RS("Type")
			PowerList = RS("powerlist")
	        ModelPower= RS("modelpower")
			Role      = RS("Role")
			ManageOtherDoc = RS("ManageOtherDoc")
			STitle="修改"
		  End If
		  RS.Close
		  Set RS=Nothing
		Else
		     STitle="添加" : RoleType=2 : Role=0 : ManageOtherDoc=1
			 ModelPower="sysset0,user0,lab0,model0,subsys0,other0,ask0,space0,bbs0"
			 For i=0 to ubound(sql,2)
			  ModelPower=Modelpower &"," & sql(4,i)&"0"
			 Next
		End If
		 %>
		<!DOCTYPE html><html>
		 <head>
		 <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		 <title>角色管理</title>
		 <link href="../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		 <script language="JavaScript" src="../../KS_Inc/common.js"></script>
		 <script language="JavaScript" src="../../KS_Inc/Jquery.js"></script>
		 <script type="text/javascript">
		 function CheckForm()
			{
			  if(document.myform.GroupName.value=="")
				{
				  alert("管理员角色名称不能为空！");
				  document.myform.GroupName.focus();
				  return false;
				}
			 $("#myform").submit();
			}
		 </script>
		 <body>
		 
         <div class="pageCont2 mt20">
         <div class='tabTitle'><%=STitle%>角色</div>
		 <form method="post" id="myform" action="KS.Admin.asp" name="myform" >
		 <input type="hidden" name="action" value="AddRoleSave"/>
		 <input type="hidden" name="id" value="<%=ID%>"/>
		 <input type="hidden" name="RoleType" value="<%=RoleType%>"/>
		 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
		 <tr class="tdbg"> 
			  <td height="30" width="150" class="clefttitle" align="right"><strong>角色名称：</strong></td>
			  <td> <input name="GroupName" class="textbox" type="text" size=30 value="<%=GroupName%>">		      </td>
			</tr>
		    <tr class="tdbg"> 
			  <td height="30" class="clefttitle" align="right"><strong>角色介绍：</strong></td>
			  <td> <textarea name="Descript" class="textbox" cols="50" rows="4"><%=Descript%></textarea></td>
			</tr>
			<%if RoleType<>3 then%>
			<tr class="tdbg"> 
			  <td height="30" class="clefttitle" align="right"><strong>审核权限：</strong></td>
			  <td> 
			   <select name="Role" class="textbox">
			    <option value="1"<%If Role="1" then Response.Write " selected"%>>发稿员</option>
				<option value="2"<%If Role="2" then Response.Write " selected"%>>初审员</option>
				<option value="3"<%If Role="3" then Response.Write " selected"%>>终审员</option>
			   </select>
			   <span class="tips">TIPS:只有终审过的文档才会在前台显示,当模型没有启用二级审核时，初审员审核的文档将直接通过显示。</span>
			  </td>
			</tr>
			<tr class="tdbg"> 
			  <td height="30" class="clefttitle" align="right"><strong>只能管理自己添加的文档：</strong></td>
			  <td> 
			   <input type="radio" name="ManageOtherDoc" value="1"<%If ManageOtherDoc="1" then Response.Write " checked"%>>是
			   <input type="radio" name="ManageOtherDoc" value="0"<%If ManageOtherDoc="0" then Response.Write " checked"%>>不是
			  <span class="tips">审核员角色，请选择“不是”。</span> </td>
			</tr>
			
			 <tr>
			 <td height="30" class="clefttitle" align="right"><strong>发布管理：</strong></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="ref20000"<%if InStr(1, PowerList,"ref20000" ,1)<>0 then Response.Write( " checked") %>>
发布站点首页
             &nbsp;&nbsp;&nbsp;<input name="PowerList" type="checkbox" id="PowerList" value="ref20001"<%if InStr(1, PowerList,"ref20001" ,1)<>0 then Response.Write( " checked") %>>
专题发布管理&nbsp;&nbsp;&nbsp;<input name="PowerList" type="checkbox" id="PowerList" value="ref20002"<%if InStr(1, PowerList,"ref20002" ,1)<>0 then Response.Write( " checked") %>>
系统JS发布管理&nbsp;&nbsp;&nbsp;<input name="PowerList" type="checkbox" id="PowerList" value="ref20003"<%if InStr(1, PowerList,"ref20003" ,1)<>0 then Response.Write( " checked") %>>
发布通用页面</td>
                  </tr>
			
			<%end if%>
			<%
			if RoleType=3 then
			Response.write "          <tr class='sort'><td colspan='2' align='center'>====此角色是超级管理员，拥有最高权限====</td></tr>"
			Response.Write "          <tr class='tdbg' style='display:none'><td colspan='2'>"
			else
			Response.write "          <tr class='sort'><td colspan='2' align='center'>====此角色的详细权限设置====</td></tr>"
			Response.Write "          <tr class='tdbg'><td colspan='2'>"
			end if
			%>
			
			
	<table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle' style="text-align:left"><strong> 一、此角色在【<font color="#FF0000">内容</font>】选项的权限</strong></td>
	 </tr>
	 </table>
          <table width="96%" border="0" align="center" cellspacing="0" cellpadding="0">   
		    <tr> 
			  <td style="padding-left:20px">
			      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			  	 <%
				  dim m:m=1
				 for i=0 to ubound(sql,2)%>
				   <tr> 
				   <td width="20%" height="35"> <strong><%=sql(1,i)%></strong></td>
				   </tr>
				   <tr>
				    <td>
					
					<%IF instr(ModelPower,sql(4,i) & "0")>0 Then
			      Response.Write("<input name=""ModelPower" & sql(0,i) & """ type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" value=""" & sql(4,I) & "0"" checked>")
			   Else
			      Response.Write("<input name=""ModelPower" & sql(0,i) & """ type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" value=""" & sql(4,I) & "0"">")
			   End IF
			   %>
                在<%=SQL(1,I)%>无任何管理权限(屏蔽)
					<br/>
					<%IF instr(ModelPower,sql(4,i) & "1")>0 Then
			      Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "1"" Checked>")
				Else
			      Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "1"">")
				End IF
				%>
                模型管理员：拥有该模型的所有管理权限(相当于对<%=sql(1,i)%>没有任何限制)
				 <br>
				 <%IF instr(ModelPower,sql(4,i) & "2")>0 Then
			     Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "2"" Checked>")
			   Else
			     Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "2"">")
			   End IF
			   %>
                栏目管理员：只拥有部分栏目(频道)管理权限
					
					</td>
				   </tr>
				   <tr ID="M_<%=sql(1,i)%>" <%IF instr(ModelPower,sql(4,i) & "2")=0 Then Response.Write("style=""display:none""") End IF%>>	 
			       <td height="22">
					<%
	  Select Case SQL(2,I)
	   Case 1,2,3,4,7,8
	   %>  
              <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td height="25" colspan="7"><strong><font color="#993300">权限设置</font></strong></td>
                  </tr>
                    <%
					Call BasePurview(PowerList,SQL,I)
					%>
                  <tr> 
                    <td height="25" colspan="7"><font color="#993300"><strong>详细指定栏目（频道）权限</strong></font></td>
                  </tr>
                  <tr> 
                    <td colspan="7"> 
					   <%
                       Call ClassList(SQL(0,I))
					   %>
					</td>
                  </tr>
                </table>
	<%case 5%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">权限设置</font></strong></td>
            </tr>
            <%Call BasePurview(PowerList,SQL,I)%>
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">常规管理</font></strong></td>
            </tr>
            <tr>
			  <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10020"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10020" ,1)<>0 then Response.Write(" checked") %>>
订单处理</td>
			  <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10021"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10021" ,1)<>0 then Response.Write(" checked") %>>
删除订单</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10014"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10014" ,1)<>0 then Response.Write(" checked") %>>
                资金明细</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10015"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10015" ,1)<>0 then Response.Write(" checked") %>>
发退货查询</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10016"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10016" ,1)<>0 then Response.Write(" checked") %>>
开发票查询</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10017"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10017" ,1)<>0 then Response.Write(" checked") %>>
销售统计</td>
            </tr>
            <tr>
              <td nowrap="nowrap" title="浏览、编辑、删除厂商等操作的权限"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20003"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20003" ,1)<>0 then Response.Write( " checked") %> />
                厂商管理</td>
				<td nowrap="nowrap" title="浏览、编辑、删除送货方式等操作的权限"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20004"<%if InStr(1, PowerList,"M" & sql(0,i) & "20004" ,1)<>0 then Response.Write( " checked") %> />
                送货/付款方式管理</td>
				<td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20006"<%if InStr(1, PowerList,"M" & sql(0,i) & "20006" ,1)<>0 then Response.Write( " checked") %> />
                商品规格管理</td>
                    
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20007"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20007" ,1)<>0 then Response.Write( " checked") %>> 优惠券管理</td>
                    <td nowrap="nowrap"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20008"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20008" ,1)<>0 then Response.Write( " checked") %>> 限时/限量管理</td>
            </tr>
			<tr>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20009"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20009" ,1)<>0 then Response.Write( " checked") %>> 捆绑销售管理</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20010"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20010" ,1)<>0 then Response.Write( " checked") %>> 换购商品管理</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20012"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20012" ,1)<>0 then Response.Write( " checked") %>> 库存报警管理</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20011"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20011" ,1)<>0 then Response.Write( " checked") %>> 超值礼包管理</td>
              <td nowrap="nowrap"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>30001"<%if InStr(1, PowerList,"M" & SQL(0,I) & "30001" ,1)<>0 then Response.Write( " checked") %> />
                团购管理</td>
				              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10018"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10018" ,1)<>0 then Response.Write(" checked") %>>
品牌管理</td>

			</tr>
			<tr>
               <td><label><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20013"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20013" ,1)<>0 then Response.Write( " checked") %>> 修改订单资料</label></td>
               <td><label><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20014"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20014" ,1)<>0 then Response.Write( " checked") %>> 管理快递单模板</label></td>
               <td><label><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20015"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20015" ,1)<>0 then Response.Write( " checked") %>> 收货人信息管理</label></td>
               <td><label><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20016"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20016" ,1)<>0 then Response.Write( " checked") %>> 商品批量改价</label></td>
			</tr>
            <tr>
              <td height="25" colspan="7"><font color="#993300"><strong>详细指定栏目（频道）权限</strong></font></td>
            </tr>
            <tr>
              <td colspan="7">                      
			        <%
                       Call ClassList(SQL(0,I))
					   %>
               </td>
            </tr>
          </table>

  <%case 9%>
           <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">栏目与信息</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910001"<%if InStr(1, PowerList,"M910001" ,1)<>0 then Response.Write(" checked") %> />
               分类管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910002"<%if InStr(1, PowerList,"M910002" ,1)<>0 then Response.Write( " checked") %> />
添加试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910003"<%if InStr(1, PowerList,"M910003" ,1)<>0 then Response.Write( " checked") %> />
编辑试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910004"<%if InStr(1, PowerList,"M910004" ,1)<>0 then Response.Write( " checked") %> />
删除试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910005"<%if InStr(1, PowerList,"M910005" ,1)<>0 then Response.Write(" checked") %>>
移动试卷</td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910007"<%if InStr(1, PowerList,"M910007" ,1)<>0 then Response.Write(" checked") %>>
发布管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910009"<%if InStr(1, PowerList,"M910009" ,1)<>0 then Response.Write(" checked") %>>
上传文件</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910010"<%if InStr(1, PowerList,"M910010" ,1)<>0 then Response.Write(" checked") %> />
               知识点管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910011"<%if InStr(1, PowerList,"M910011" ,1)<>0 then Response.Write(" checked") %> />
              试题纠错管理</td>
			</tr>
			</table>
		   <%case 10%>
           <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">栏目与信息</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010001"<%if InStr(1, PowerList,"M1010001" ,1)<>0 then Response.Write(" checked") %> />
                招聘系统设置</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010002"<%if InStr(1, PowerList,"M1010002" ,1)<>0 then Response.Write( " checked") %> />
招聘单位管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010003"<%if InStr(1, PowerList,"M1010003" ,1)<>0 then Response.Write( " checked") %> />
个人简历管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010004"<%if InStr(1, PowerList,"M1010004" ,1)<>0 then Response.Write( " checked") %> />
招聘职位管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010005"<%if InStr(1, PowerList,"M1010005" ,1)<>0 then Response.Write(" checked") %>>
行业职位设置</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010007"<%if InStr(1, PowerList,"M1010007" ,1)<>0 then Response.Write(" checked") %>>
简历模板管理</td>
            </tr>
			</table>
			<%
	End Select
	%>  
				   </td>
				   </tr>
				   
                  <%
				  Next%>
			</table>
			</td>
		</tr>
		</table>
		
		<table style="padding-left:40px" border="0" cellpadding="0" cellspacing="0">
			 <tr <%IF KS.C_S(12,21)="0" THEN Response.Write " style='display:none'"%>> 
				   <td width="20%" height="35"> <strong>问答系统权限</strong></td>
				   </tr>
				   <tr<%IF KS.C_S(12,21)="0" THEN Response.Write " style='display:none'"%>>
				    <td>
						<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
						<tr> 
						  <td height="25" colspan="5"> 
							<%
								IF instr(ModelPower,"ask0")>0 Then
								  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Ask','none')"" name=""ask"" value=""ask0"" checked>")
								ELSE
								  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Ask','none')"" name=""ask"" value=""ask0"">")
								END IF
								%>
							在问答系统无任何管理权限(屏蔽)
							  <br/>
							<%
								IF instr(ModelPower,"ask1")>0 Then
								  Response.Write("<input type=""radio"" name=""ask"" onclick=""SetPowerListValue('Ask','')"" value=""ask1"" checked>")
								 ELSE
								  Response.Write("<input type=""radio"" name=""ask"" onclick=""SetPowerListValue('Ask','')"" value=""ask1"">")
								 END IF%>
							拥有指定的部分管理权限 
						 </td>
						</tr>
						<tr ID="Ask" <% IF instr(ModelPower,"ask1")="0" or KS.C_S(12,21)="0" then Response.Write("style=""display:none""") End IF%>> 
										<td><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10000"<%if InStr(1, PowerList,"WDXT10000" ,1)<>0 then Response.Write( " checked") %>> 
										问答参数设置
					</td>
										<td><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10001"<%if InStr(1, PowerList,"WDXT10001" ,1)<>0 then Response.Write( " checked") %>> 
										编辑删除问题</td>
										<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10002"<%if InStr(1, PowerList,"WDXT10002" ,1)<>0 then Response.Write( " checked") %>> 
										问答分类管理</td>
										<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10003"<%if InStr(1, PowerList,"WDXT10003" ,1)<>0 then Response.Write( " checked") %>>等级头衔管理</td>
										<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10004"<%if InStr(1, PowerList,"WDXT10004" ,1)<>0 then Response.Write( " checked") %>>审核问题答案</td>
										<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10005"<%if InStr(1, PowerList,"WDXT10005" ,1)<>0 then Response.Write( " checked") %>>专家审核管理</td>
					
					</tr>
			</table>
		</td>
		</tr>
		</table>
			
			<table  style="padding-left:40px" border="0" cellpadding="0" cellspacing="0">	
			<tr<%IF KS.C_S(11,21)="0" THEN Response.Write " style='display:none'"%>> 
			   <td width="20%" height="35"> <strong>论坛系统权限</strong></td>
			</tr>
			<tr<%IF KS.C_S(11,21)="0" THEN Response.Write " style='display:none'"%>> 
              <td height="25" colspan="5"> 
                <%
					IF instr(ModelPower,"bbs0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('bbs','none')"" name=""bbs"" value=""bbs0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('bbs','none')"" name=""bbs"" value=""bbs0"">")
					END IF
					%>
                在论坛系统无任何管理权限(屏蔽)
				  <br/>
                <%
					IF instr(ModelPower,"bbs1")>0 Then
					  Response.Write("<input type=""radio"" name=""bbs"" onclick=""SetPowerListValue('bbs','')"" value=""bbs1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""bbs"" onclick=""SetPowerListValue('bbs','')"" value=""bbs1"">")
					 END IF%>
                拥有指定的部分管理权限 
			 </td>
            </tr>
			 <tbody ID="bbs" <% IF instr(ModelPower,"bbs1")="0" or KS.C_S(11,21)="0" then Response.Write("style=""display:none""") End IF%>> 
                   <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMB10000"<%if InStr(1, PowerList,"KSMB10000" ,1)<>0 then Response.Write( " checked") %>>
论坛帖子管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMB10001"<%if InStr(1, PowerList,"KSMB10001" ,1)<>0 then Response.Write( " checked") %>>
论坛版面分类管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMB10002"<%if InStr(1, PowerList,"KSMB10002" ,1)<>0 then Response.Write( " checked") %>>
当前数据表管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMB10003"<%if InStr(1, PowerList,"KSMB10003" ,1)<>0 then Response.Write( " checked") %>>
等级头衔勋章管理 </td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMB10004"<%if InStr(1, PowerList,"KSMB10004" ,1)<>0 then Response.Write( " checked") %>>
论坛系统设置 </td>
                  </tr>
		    </tbody>
			
			
			
					
			<tr<%IF KS.C_S(13,21)="0" THEN Response.Write " style='display:none'"%>> 
			   <td width="20%" height="35"> <strong>空间门户权限</strong></td>
			</tr>
			<tr<%IF KS.C_S(13,21)="0" THEN Response.Write " style='display:none'"%>> 
              <td height="25" colspan="5"> 
                <%
					IF instr(ModelPower,"space0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Space','none')"" name=""space"" value=""space0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Space','none')"" name=""space"" value=""space0"">")
					END IF
					%>
                在空间门户无任何管理权限(屏蔽)
				  <br/>
                <%
					IF instr(ModelPower,"space1")>0 Then
					  Response.Write("<input type=""radio"" name=""space"" onclick=""SetPowerListValue('Space','')"" value=""space1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""space"" onclick=""SetPowerListValue('Space','')"" value=""space1"">")
					 END IF%>
                拥有指定的部分管理权限 
			 </td>
            </tr>
            <tbody ID="Space" <% IF instr(ModelPower,"space1")="0" or KS.C_S(13,21)="0" then Response.Write("style=""display:none""") End IF%>> 
                   <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10000"<%if InStr(1, PowerList,"KSMS10000" ,1)<>0 then Response.Write( " checked") %>>
空间参数设置</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10001"<%if InStr(1, PowerList,"KSMS10001" ,1)<>0 then Response.Write( " checked") %>>
个人空间管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10002"<%if InStr(1, PowerList,"KSMS10002" ,1)<>0 then Response.Write( " checked") %>>
空间日志管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10003"<%if InStr(1, PowerList,"KSMS10003" ,1)<>0 then Response.Write( " checked") %>>
用户相册管理 </td>
                  </tr>
				  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10004"<%if InStr(1, PowerList,"KSMS10004" ,1)<>0 then Response.Write( " checked") %>>
用户圈子管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10005"<%if InStr(1, PowerList,"KSMS10005" ,1)<>0 then Response.Write( " checked") %>>
用户留言管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20013"<%if InStr(1, PowerList,"KSMS20013" ,1)<>0 then Response.Write( " checked") %>>
行业广告管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10007"<%if InStr(1, PowerList,"KSMS10007" ,1)<>0 then Response.Write( " checked") %>>
用户歌曲管理</td>
                  </tr>
				  <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10008"<%if InStr(1, PowerList,"KSMS10008" ,1)<>0 then Response.Write( " checked") %>>
企业信息管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10009"<%if InStr(1, PowerList,"KSMS10009" ,1)<>0 then Response.Write( " checked") %>>
企业新闻管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10010"<%if InStr(1, PowerList,"KSMS10010" ,1)<>0 then Response.Write( " checked") %>>
企业产品管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20011"<%if InStr(1, PowerList,"KSMS20011" ,1)<>0 then Response.Write( " checked") %>>
荣誉证书管理</td>
                  </tr>
				  <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20012"<%if InStr(1, PowerList,"KSMS20012" ,1)<>0 then Response.Write( " checked") %>>
行业分类管理</td>
				    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20016"<%if InStr(1, PowerList,"KSMS20016" ,1)<>0 then Response.Write( " checked") %>>
                   微博数据管理</td>

				  </tr>					
				 </tbody>	
					
					
				  </table>
					
		            </td>
				  </tr>
				  
				 
				 </table>
				  
			  </td>
			</tr>
			
		
	
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle' style="text-align:left"><strong> 二、此角色在【<font color="#FF0000">系统设置</font>】选项的权限</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"sysset0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('System','none')"" name=""sysset"" value=""sysset0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('System','none')"" name=""sysset"" value=""sysset0"">")
					END IF
					%>
                在系统管理中心无任何管理权限(屏蔽)
				  <br/>
                <%
					IF instr(ModelPower,"sysset1,")>0 Then
					  Response.Write("<input type=""radio"" name=""sysset"" onclick=""SetPowerListValue('System','')"" value=""sysset1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""sysset"" onclick=""SetPowerListValue('System','')"" value=""sysset1"">")
					 END IF%>
                拥有指定的部分管理权限 
			 </td>
            </tr>
            <tr ID="System" <% IF instr(ModelPower,"sysset1,")="0" then Response.Write("style=""display:none""") End IF%>> 
              <td height="25" colspan="2" style="padding-left:20px;padding-top:5px;">
 
    			<table  width="100%" border="0" cellpadding="0" cellspacing="0">
				 <tr> 
                    <td height="25" width="20%"> <input name="PowerList" type="checkbox" id="PowerList" value="KMST10001"<%if InStr(1, PowerList,"KMST10001" ,1)<>0 then Response.Write( " checked") %>>系统参数配置</td>
                  <td width="20%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010001"<%if InStr(1, PowerList,"M010001" ,1)<>0 then Response.Write(" checked") %>>栏目管理</td>
                    <td width="20%"><input name="PowerList" type="checkbox" id="PowerList" value="M010005"<%if InStr(1, PowerList,"M010005" ,1)<>0 then Response.Write(" checked") %>>文档批量设置</td>
                    <td width="20%"><input name="PowerList" type="checkbox" id="PowerList" value="M010006"<%if InStr(1, PowerList,"M010006" ,1)<>0 then Response.Write(" checked") %>> 回收站管理</td>
				  <td width="20%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010007"<%if InStr(1, PowerList,"M010007" ,1)<>0 then Response.Write(" checked") %>>一键快速生成HTML</td>
				  </tr>
				<tr>
					 <td height="25" width="20%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST10002"<%if InStr(1, PowerList,"KMST10002" ,1)<>0 then Response.Write( " checked") %>>API整合设置</td>
					 <td width="20%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST10003"<%if InStr(1, PowerList,"KMST10003" ,1)<>0 then Response.Write( " checked") %>>在线支付平台设置</td>
                    <td width="20%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST20000"<%if InStr(1, PowerList,"KMST20000" ,1)<>0 then Response.Write( " checked") %>>更新缓存</td>
                       <td width="20%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST10017"<%if InStr(1, PowerList,"KMST10017" ,1)<>0 then Response.Write(" checked") %> />
省市地区管理</td>
                    <td></td>
                  </tr>
				  </table>	
				  
			
			  </td>
            </tr>
			
          </table>
	  </td></tr>
	  </table>
	  
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle' style="text-align:left"><strong> 三、此角色在【<font color="#FF0000">用户管理</font>】的权限</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"user0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('UserAdmin','none')"" name=""userPower"" value=""user0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('UserAdmin','none')"" name=""UserPower"" value=""user0"">")
					END IF
					%>
                在用户管理中心无任何管理权限(屏蔽)
				<br/>
                <%
					IF instr(ModelPower,"user1")>0 Then
					  Response.Write("<input type=""radio"" name=""UserPower"" onclick=""SetPowerListValue('UserAdmin','')"" value=""user1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""UserPower"" onclick=""SetPowerListValue('UserAdmin','')"" value=""user1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="UserAdmin" <% IF instr(ModelPower,"user1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                
				<tr>
				  <td colspan="10"><input onClick="if(this.checked){$('#subuser').slideDown('fast');}else{$('#subuser').slideUp('fast');}" name="PowerList" type="checkbox" id="PowerList" value="KMUA10002"<%if InStr(1, PowerList,"KMUA10002" ,1)<>0 then Response.Write( " checked") %>>
                     <strong>注册用户管理</strong>
					 
					<div class="attention" style="<%if InStr(1, PowerList,"KMUA10002" ,1)=0 then Response.Write( "display:none;") %>margin:10px;padding:5px;" id="subuser"> <strong>子权限：</strong><br/>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100021"<%if InStr(1, PowerList,"KMUA100021" ,1)<>0 then Response.Write( " checked") %>>注册用户查看管理</label>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100022"<%if InStr(1, PowerList,"KMUA100022" ,1)<>0 then Response.Write( " checked") %>>注册用户修改管理</label>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100023"<%if InStr(1, PowerList,"KMUA100023" ,1)<>0 then Response.Write( " checked") %>>注册用户锁定管理</label>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100024"<%if InStr(1, PowerList,"KMUA100024" ,1)<>0 then Response.Write( " checked") %>>注册用户删除管理</label>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100025"<%if InStr(1, PowerList,"KMUA100025" ,1)<>0 then Response.Write( " checked") %>>注册用户续费管理</label>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100026"<%if InStr(1, PowerList,"KMUA100026" ,1)<>0 then Response.Write( " checked") %>>注册用户积分管理</label>
					<label><input name="PowerList" type="checkbox" id="PowerList" value="KMUA100027"<%if InStr(1, PowerList,"KMUA100027" ,1)<>0 then Response.Write( " checked") %>>添加注册用户</label>
					</div>
					 </td>
				 </tr>
				  <tr>
                    <td><input disabled="disabled" name="PowerList" type="checkbox" id="PowerList" value="KMUA10001"<%if InStr(1, PowerList,"KMUA10001" ,1)<>0 then Response.Write( " checked") %>>
                      管理员管理</td>
                                   
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10006"<%if InStr(1, PowerList,"KMST10006" ,1)<>0 then Response.Write( " checked") %>>
                      后台日志管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10003"<%if InStr(1, PowerList,"KMUA10003" ,1)<>0 then Response.Write( " checked") %>>
                     用户短信管理 </td>
                    <td> <input name="PowerList" type="checkbox" id="PowerList" value="KMUA10004"<%if InStr(1, PowerList,"KMUA10004" ,1)<>0 then Response.Write( " checked") %>>
                      用户组管理</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10008"<%if InStr(1, PowerList,"KMUA10008" ,1)<>0 then Response.Write( " checked") %>>
                    充值卡管理</td>                 
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10011"<%if InStr(1, PowerList,"KMUA10011" ,1)<>0 then Response.Write( " checked") %>>
                    查看工作进度</td>					 </tr>
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10005"<%if InStr(1, PowerList,"KMUA10005" ,1)<>0 then Response.Write( " checked") %>>
                    会员点券明细</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10006"<%if InStr(1, PowerList,"KMUA10006" ,1)<>0 then Response.Write( " checked") %>>
                    会员有效期明细</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10007"<%if InStr(1, PowerList,"KMUA10007" ,1)<>0 then Response.Write( " checked") %>>
                    会员资金明细</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10017"<%if InStr(1, PowerList,"KMUA10017" ,1)<>0 then Response.Write( " checked") %>>
                    会员积分明细</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10009"<%if InStr(1, PowerList,"KMUA10009" ,1)<>0 then Response.Write( " checked") %>>
                    发送邮件管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10010"<%if InStr(1, PowerList,"KMUA10010" ,1)<>0 then Response.Write( " checked") %>>
                    修改自己的密码</td>
                  </tr>
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10012"<%if InStr(1, PowerList,"KMUA10012" ,1)<>0 then Response.Write( " checked") %>>
                    会员字段管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10013"<%if InStr(1, PowerList,"KMUA10013" ,1)<>0 then Response.Write( " checked") %>>
                    会员表单管理</td>
                   
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10015"<%if InStr(1, PowerList,"KMUA10015" ,1)<>0 then Response.Write( " checked") %>>
                    会员使用记录</td>
                     <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10016"<%if InStr(1, PowerList,"KMUA10016" ,1)<>0 then Response.Write( " checked") %>>
                    实名认证管理</td>
                    <td height="25">&nbsp;</td>
                    <td height="25">&nbsp;</td>
                  </tr>
			
		
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>
	  
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle' style="text-align:left"><strong> 四、此角色在【<font color="#FF0000">标签模板管理</font>】的权限</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"lab0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('KMTemplatePower','none')"" name=""labPower"" value=""lab0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('KMTemplatePower','none')"" name=""labPower"" value=""lab0"">")
					END IF
					%>
                在模板标签管理管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"lab1")>0 Then
					  Response.Write("<input type=""radio"" name=""labPower"" onclick=""SetPowerListValue('KMTemplatePower','')"" value=""lab1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""labPower"" onclick=""SetPowerListValue('KMTemplatePower','')"" value=""lab1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="KMTemplatePower" <% IF instr(ModelPower,"lab1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                 

                  <tr>
                    <td width="10%"><strong>模板标签</strong></td> 
                    <td width="16%"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10001"<%if InStr(1, PowerList,"KMTL10001" ,1)<>0 then Response.Write( " checked") %>>
                      系统函数标签 </td>
                    <td height="25"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10002"<%if InStr(1, PowerList,"KMTL10002" ,1)<>0 then Response.Write( " checked") %>>
                      自定义SQL标签</td>
                    <td> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10003"<%if InStr(1, PowerList,"KMTL10003" ,1)<>0 then Response.Write( " checked") %>>
                    自定义静态标签 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10004"<%if InStr(1, PowerList,"KMTL10004" ,1)<>0 then Response.Write( " checked") %>>
                    系统JS管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10005"<%if InStr(1, PowerList,"KMTL10005" ,1)<>0 then Response.Write( " checked") %>>
自由JS管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10011"<%if InStr(1, PowerList,"KMTL10011" ,1)<>0 then Response.Write( " checked") %>>
自定义生成XML</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10006"<%if InStr(1, PowerList,"KMTL10006" ,1)<>0 then Response.Write( " checked") %>>自定义页面管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10007"<%if InStr(1, PowerList,"KMTL10007" ,1)<>0 then Response.Write( " checked") %>>
模板管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10008"<%if InStr(1, PowerList,"KMSL10008" ,1)<>0 then Response.Write( " checked") %> />
生成顶部菜单</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10009"<%if InStr(1, PowerList,"KMSL10009" ,1)<>0 then Response.Write( " checked") %> />
生成树型菜单</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10010"<%if InStr(1, PowerList,"KMSL10010" ,1)<>0 then Response.Write( " checked") %> />
通用循环标签</td><td>&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>

	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle' style="text-align:left"><strong> 五、此角色在【<font color="#FF0000">模型字段管理</font>】的权限</strong></td>
	 </tr>
	 </table>
	  
	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"model0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('ModelPowers','none')"" name=""modelPower"" value=""model0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('ModelPowers','none')"" name=""modelPower"" value=""model0"">")
					END IF
					%>
                在模型字段管理管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"model1")>0 Then
					  Response.Write("<input type=""radio"" name=""modelPower"" onclick=""SetPowerListValue('ModelPowers','')"" value=""model1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""modelPower"" onclick=""SetPowerListValue('ModelPowers','')"" value=""model1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="ModelPowers" <% IF instr(ModelPower,"model1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10000"<%if InStr(1, PowerList,"KSMM10000" ,1)<>0 then Response.Write( " checked") %>>
添加模型</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10001"<%if InStr(1, PowerList,"KSMM10001" ,1)<>0 then Response.Write( " checked") %>>
修改模型</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10002"<%if InStr(1, PowerList,"KSMM10002" ,1)<>0 then Response.Write( " checked") %>>
删除模型</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10003"<%if InStr(1, PowerList,"KSMM10003" ,1)<>0 then Response.Write( " checked") %>>
模型字段管理 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10004"<%if InStr(1, PowerList,"KSMM10004" ,1)<>0 then Response.Write( " checked") %>>
信息统计 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10005"<%if InStr(1, PowerList,"KSMM10005" ,1)<>0 then Response.Write( " checked") %>>
开启关闭</td>
                  </tr>
                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>



  <%if KS.CheckDir("../PLUS") Then  '检查应用目录是否存在%>
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle' style="text-align:left"><strong> 六、此角色在【<font color="#FF0000">应用插件</font>】的权限</strong></td>
	 </tr>
	 </table>
<table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"subsys0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('SubSysPowers','none')"" name=""subsysPower"" value=""subsys0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('SubSysPowers','none')"" name=""subsysPower"" value=""subsys0"">")
					END IF
					%>
                在应用插件无任何管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"subsys1")>0 Then
					  Response.Write("<input type=""radio"" name=""subsysPower"" onclick=""SetPowerListValue('SubSysPowers','')"" value=""subsys1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""subsysPower"" onclick=""SetPowerListValue('SubSysPowers','')"" value=""subsys1"">")
					 END IF%>
                拥有指定应用的部分管理权限 </td>
            </tr>
            <tr ID="SubSysPowers" <% IF instr(ModelPower,"subsys1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td style="padding-left:20px"> 
			  
			  <%
			  
			  Dim FsoItem ,Nodek2,MyItem,MyRole,FieldXML
			  Dim FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
			  Dim FolderObj:Set FolderObj = FsoObj.GetFolder(Server.MapPath("../plus"))
			  Dim SubFolderObj:Set SubFolderObj = FolderObj.SubFolders
			  For Each FsoItem In SubFolderObj
						 if KS.CheckFile("../plus/"&FsoItem.name&"/Config.xml") then
						 SET FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
										   FieldXML.async = false
										   FieldXML.setProperty "ServerHTTPRequest", true 
										   FieldXML.load(Server.MapPath("../plus/"&FsoItem.name&"/Config.xml"))
										   if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
										   Dim NodeXML:Set NodeXML=FieldXML.DocumentElement.SelectSingleNode("App")	
											If Not NodeXML Is Nothing Then
												Dim AppName:AppName=NodeXML.SelectSingleNode("AppName").Text
												Dim AppStatus:AppStatus=NodeXML.SelectSingleNode("AppStatus").Text
												 Role=NodeXML.SelectSingleNode("Role").Text
												  %>
												  <li class="App"<%if AppStatus<>"1" then response.write " style='display:none'"%>>
												     <label><input onClick="if (this.checked){$('#<%=Role%>').show();}else{$('#<%=Role%>').hide();}" name="PowerList" type="checkbox" value="<%=Role%>" <%if InStr(1, PowerList,Role ,1)<>0 then Response.Write( " checked") %>><%=AppName%></label>
													 
													 <ul id="<%=Role%>" class="menu"<%if InStr(1, PowerList,Role ,1)=0 then Response.Write( " style='display:none'") %>>
														
														   <%
															Dim NodeXML2:Set NodeXML2=NodeXML.SelectNodes("AppItem")
															 For Each Nodek2 In NodeXML2
																	 MyItem=Nodek2.SelectSingleNode("ItemName").text
																	 MyRole=Nodek2.SelectSingleNode("Role").text
																	%>
																	<li><input name="PowerList" type="checkbox" value="<%=MyRole%>" <%if InStr(1, PowerList,MyRole ,1)<>0 then Response.Write( " checked") %>><%=MyItem %></li>
																	<%
															 Next
														  %>
														
															
														</ul>
												</li>
												  <%
										   End If
							  End If
				   End If
			 Next
			  
			  
			  %>
			  
			  
			  
			  </td>
            </tr>
          </table>
	  </td></tr>
	  </table>
<%END IF%>

<script>
function SetPowerListValue(Module,Value)
{ $('#'+Module)[0].style.display=Value;
}
</script>
		</form>
        </div>
		 
		 <%
		End Sub
		
		Sub AddRoleSave() 
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim GroupName:GroupName=KS.G("GroupName")
		 Dim Descript:Descript=KS.G("Descript")
		 Dim RoleType:RoleType=KS.ChkClng(KS.G("RoleType"))
		 If KS.IsNul(GroupName) Then 
		   KS.AlertHintScript "角色名称必须输入!"
		 End If
		 
		 
		 Dim SQL,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
			RSC.Open "Select ChannelID,ChannelName,BasicType,ItemName,ModelEname,ChannelStatus From KS_Channel where channelstatus=1 Order By ChannelID",Conn,1,1
			If Not RSC.Eof Then
			  SQL=RSC.GetRows(-1)
			End If
			RSC.Close

			 For I=0 To Ubound(sql,2)
			  If I=0 Then
				 ModelPower=Replace(Request("ModelPower" & sql(0,i) &"")," ","")
			  Else
				 ModelPower=ModelPower & "," & Replace(Request("ModelPower" & sql(0,i) &"")," ","")
			  End IF
			 Next
			 ModelPower=request("sysset1") & "," & request("sysset2") & "," & request("otherpower") &"," & request("sysset") &"," & request("userpower") & "," & request("labpower") &"," &request("modelpower") & "," &request("subsyspower")&","&request("ask")&"," & request("bbs") &"," & request("space") & ","& modelpower
			 PowerList=Replace(Trim(KS.G("PowerList"))," ","")
			 IF PowerList="" THEN PowerList=0
			
		' ks.die powerlist
		 
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_UserGroup Where ID=" & id,CONN,1,3
		 If RS.Eof Then
		  RS.AddNew
		 End If
		   RS("GroupName")=GroupName
		   RS("Descript")=Descript
		   If RoleType<>3 Then
		   RS("Type")=2
		   End If
		   RS("ModelPower")= ModelPower
		   RS("PowerList")=PowerList
		   RS("ShowOnReg")=0
		   RS("ChargeType")=0
		   RS("GroupPoint")=0
		   RS("ValidDays")=0
		   RS("PowerType")=0
		   RS("FormID")=0
		   RS("UserType")=0
		   RS("SpaceSize")=0
		   RS("validtype")=0
		   RS("Role")=KS.ChkClng(KS.S("Role"))
		   RS("ManageOtherDoc")=KS.ChkClng(KS.S("ManageOtherDoc"))
		  RS.Update
		  RS.MoveLast
		  Dim RoleId:RoleId=RS("ID")
		 RS.Close
		 Set RS=Nothing
		 
		 
		 RSC.Open "Select AdminPurview,ID From KS_Class Order By ClassID",conn,1,3
			Do While Not RSC.Eof
			    
			  If KS.FoundInArr(Replace(Request("AdminPurview")," ",""),RSC(1),",") Then
			     If KS.IsNul(RSC(0)) Then 
				  RSC(0)=RoleId
				 Else
				  RSC(0)=KS.FilterRepeatInArray(RSC(0) & "," & RoleId,",")
				 End If
				 RSC.Update
			  Else
			     If KS.IsNul(RSC(0)) Then
				  RSC(0)=""
				 Else
					 RSC(0)=DelItemInArr(RSC(0),RoleId,",")
				 End If
			   	 RSC.Update
			  End If
			     on error resume next
				 Dim ENode:Set ENode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & RSC(1) & "']")
				 ENode.SelectSingleNode("@ks16").text=RSC(0)
				 If err Then err.clear
			  
			  RSC.MoveNext
			loop
			RSC.Close
			Set RSC=nothing
		 
		 Application(KS.SiteSN&"_class")=empty
		 If ID=0 Then
		   Response.Write ("<script>if (confirm('添加管理员角色成功,继续添加吗?')) {location.href='?Action=AddRole';} else { location.href='KS.Admin.asp?Action=Role';}</script>")
         Else
		   Response.Write ("<script>alert('管理员角色修改成功!');location.href='KS.Admin.asp?Action=Role';</script>")
		 End if
		End Sub
		
		Sub AdminList()
		 Call Head()
		 Response.Write "<div class='pageCont2'>"
		 Response.Write"<div class='tabTitle'>管理员管理</div>"

		Response.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""1"">"
		
			  Dim Param:Param = " Where 1=1"
			  If KeyWord <> "" Then
				Select Case SearchType
				  Case 0
				   Param = Param & " And UserName like '%" & KeyWord & "%'"
				  Case 1
				   Param = Param & " And Description like '%" & KeyWord & "%'"
				End Select
			  End If
			  If KS.ChkClng(KS.G("GroupID"))<>0 Then Param=Param & " And GroupID=" & KS.ChkClng(KS.G("GroupID"))
			  Param = Param & " Order BY SuperTF Desc,AddDate desc"
			  SqlStr = "Select b.groupname,b.[type],a.* From KS_Admin a inner join KS_UserGroup b on a.groupid=B.id " & Param
				 Response.Write ("<tr> ")
				 Response.Write ("  <td>")
				 Response.Write ("    <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
				 Response.Write ("<tr align=""center""><!--<td  class=""sort"">选择</td>--><td height=23 width=100  class=""sort"">管 理 员</td><td width=120  class=""sort"">类 型</td><td class=""sort"">最后登录IP</td><td class=""sort"">最后登录时间</td><td  class=""sort"">最后登出时间</td><td  class=""sort"">登录次数</td><td class=""sort"">锁 定</td><td class='sort'>管理操作</td></tr>")
		Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		 RSObj.Open SqlStr, Conn, 1, 1
		 Dim T, TitleStr, LockStr, ShortName
		 Dim MaxPerPage:Maxperpage=20
		 CurrentPage=KS.ChkCLng(request("page"))
		 If CurrentPage<=0 Then CurrentPage=1
		 Dim TotalPut:totalPut = RSObj.RecordCount
		 If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			RSObj.Move (CurrentPage - 1) * MaxPerPage
		 End If
		 
		 dim i:i=0
		 
		Do While Not RSObj.EOF
				Response.Write "<tr class='list' id='u" & RSObj("AdminID") &"'"
				if rsobj("supertf")<>"1" Then Response.Write " onclick=""chk_iddiv('" & RSObj("AdminID") &"')"""
				Response.Write " onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
				Response.Write "<!--<td class='splittd' align=""center""><input "
				if rsobj("supertf")="1" Then Response.Write " Disabled" 
				Response.Write " name=""AdminID"" type=""checkbox""  onclick=""chk_iddiv('" & RSObj("AdminID") &"')"" id='c" & RSObj("AdminID") &"'  value=""" & RSObj("AdminID") &"""></td>-->"

				If RSObj("Locked") = 1 Then
					LockStr = "<font color=red>已锁定</font>"
					Else
					LockStr = "<font color=green>正常</font>"
				End If
				TitleStr = " TITLE='姓 名:" & RSObj("RealName") & "&#13;&#10;性 别:" & RSObj("Sex") & "&#13;&#10;添加时间:" & RSObj("AddDate") & "&#13;&#10;简要描述:" & RSObj("Description") & "'"
			  Response.Write ("<td class='splittd' height=25" & TitleStr & ">&nbsp;<span ondblclick=""EditAdmin(" & RSObj("AdminID") & ");"" AdminID=""" & RSObj("AdminID") & """><i class=' icon manage" & Trim(CStr(RSObj("SuperTF"))) & "'></i><span style=""cursor:default"">" & RSObj("UserName") & "</span><span></td>")
			  Response.Write ("<td  class='splittd' align=""center"">")
			  If RSObj("Type")="3" Then
			  response.write "<span style='color:#ff3300'>" & RSObj("GroupName") & "</span>"
			  Else
			  response.write RSObj("GroupName")
			  End If
			  Response.Write ("</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLoginIP") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLoginTime") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLogoutTime") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LoginTimes") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & LockStr & "</td>")
			  Response.Write ("<td class='splittd' align=""center""><a href='javascript:EditAdmin(" & rsobj("AdminID") &")' class='setA'>修改</a>| <a")
			  if rsobj("supertf")="1" then response.write " disabled" else response.write " href='javascript:Delete("&rsobj("AdminID")&")'"
			  Response.Write " class='setA'>删除</a>"
			  Response.write "|<a href='javascript:SetAdminPassWord(" & rsobj("adminid") &")' class='setA'>设置密码</a>"
			  Response.Write ("</td></tr>")
			  RSObj.MoveNext
			 i=i+1
			 if i>=maxperpage then exit do
			Loop
			RSObj.Close:Conn.Close:Set RSObj = Nothing:Set GRS = Nothing
		Response.Write "<tr> <td height='35' colspan='10' align='right'>"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		Response.Write "    </td>"
		Response.Write " </tr>"
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub AdminAdd()
		 IF KS.G("Method")="save" Then
		    Call AdminSave()
		   Else
		    Call AdminAddOrEdit()
		  End IF
		End SUB
		Sub AdminAddOrEdit()
		Response.Write "<!DOCTYPE HTML>"
		Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link href=""../Include/admin_style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/jquery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../ks_Inc/CheckPassWord.js""></script>"
		Response.Write "<style>"
		Response.Write ".rank { border:none; background:url(../../images/rank.gif) no-repeat; width:136px; height:22px; vertical-align:middle; cursor:default; }"
		Response.Write ".r0 { background-position:0 0; }"
		Response.Write ".r1 { background-position:0 -22px; }"
		Response.Write ".r2 { background-position:0 -44px; }"
		Response.Write ".r3 { background-position:0 -66px; }"
		Response.Write ".r4 { background-position:0 -88px; }"
		Response.Write ".r5 { background-position:0 -110px; }"
		Response.Write ".r6 { background-position:0 -132px; }"
		Response.Write ".r7 { background-position:0 -154px; }"
		Response.Write "</style>"
		Response.Write "<title>管理员添加</title>"
		Response.Write "</head>"
		
		Dim AdminID, PrUserName, PassWord, Locked, RealName, Sex, TelPhone, Email, Descript, Action, GroupID, SuperTF
		
		Action = KS.G("Action")
		AdminID = KS.G("AdminID")
		GroupID = KS.G("GroupID")
		If Action = "" Then Action = "AddAdmin"
		If AdminID <> "" Then
		   Dim RSObj:Set RSObj = Conn.Execute("Select top 1 * From KS_Admin Where AdminID=" & AdminID)
		  If Not RSObj.EOF Then
			 UserName = Trim(RSObj("UserName"))
			 PrUserName=Trim(RSObj("PrUserName"))
			 Locked = Trim(CStr(RSObj("Locked")))
			 RealName = Trim(RSObj("RealName"))
			 Sex = Trim(RSObj("Sex"))
			 TelPhone = Trim(RSObj("TelPhone"))
			 Email = Trim(RSObj("Email"))
			 Descript = Trim(RSObj("Description"))
			 SuperTF = Trim(CStr(RSObj("SuperTF")))
			 GroupID=RSObj("GroupID")
		  End If
		   RSObj.Close:Set RSObj = Nothing
		End If
		
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		
		 If AdminID = "" Then
			Response.Write ("<div class='topdashed sort'>添加管理员</div>")
		  Else
			Response.Write ("<div class='topdashed sort'>修改管理员</div>")
		  End If
		Response.Write "<div class='pageCont2'>"
		Response.Write "  <form action=""?Method=save"" name=""AdminForm"" method=""post"">"
		Response.Write "  <table width=""100%"" border=""0"" class=""ctable"" align=""center"" cellpadding=""3"" cellspacing=""1"">"
		Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""" & Action & """>"
		Response.Write "   <input name=""AdminID"" type=""hidden"" value=""" & AdminID & """>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td class='clefttitle' align=""right"">管理员名：</td>"
		Response.Write "      <td height=""25"" colspan='3'>"
					
					If Action = "Edit" Then
						 Response.Write ("<input name=""UserName"" Readonly value=""" & UserName & """ type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
					Else
						 Response.Write ("<input name=""UserName""  type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
					End If
					 
		Response.Write "              用于登录后台的名称</td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "            <td height=""25"" class='clefttitle' align=""right"">前台用户名：</td>"
		Response.Write "            <td colspan='3'>"
					If Action = "Edit" Then
						 Response.Write ("<input name=""PrUserName"" readonly value=""" & PrUserName & """ type=""text"" id=""PrUserName"" size=""30"" class=""textbox"">")
					Else
					     Response.write ("&nbsp;&nbsp;&nbsp;<input type='radio' name='PreAddType' onclick=""$('#pr').hide();"" value='1' checked>自动创建")
					     Response.write ("&nbsp;&nbsp;&nbsp;<input type='radio' name='PreAddType' onclick=""$('#pr').show();"" value='0'>已有会员账号")
						 Response.Write ("<br><br><div id='pr' style='display:none'><input name=""PrUserName""  type=""text"" id=""PrUserName"" size=""30"" class=""textbox""></div>")
					End if
					 
		Response.Write "    </td>"
		Response.Write "          </tr>"		
				  
				  If Action <> "Edit" Then
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" class=""clefttitle"" align=""right"">初始密码：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		Response.Write "             <table border='0' cellspacing='0' cellpadding='0'><tr><td style='padding:0'><input name=""PassWord"" type=""password"" size=""30"" class=""textbox"" onkeyup=""javascript:if(this.value!=''){$('#ps').show();setPasswordLevel(this, document.getElementById('passwordLevel'));}"" onmouseout=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));"" onblur=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));""> </td><td align='left' id='ps' style='display:none'>密码强度："
		Response.Write "         <input name=""Input"" disabled=""disabled"" class=""rank r0"" id=""passwordLevel"" /></td>"

		Response.Write "          </tr></table></td></tr>"
				 
				 End If
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>所属角色：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		 If SuperTf=1 Then
		    Response.Write "<input type=""hidden"" name=""GroupID"" value=""" & GroupID & """/>"
			Response.Write "<select name=""sGroupID"" disabled class='textbox'><option value='0'>---选择管理员角色---</option>"
		 Else
			Response.Write "<select name=""GroupID"" class='textbox'><option value='0'>---选择管理员角色---</option>"
		  End If
			 Dim RSR:Set RSR=Conn.Execute("Select ID,GroupName From KS_UserGroup Where [Type]>1 order by id desc")
			 Do While Not RSR.Eof
			  If KS.ChkClng(GroupID)=KS.ChkClng(RSR(0)) Then
			   Response.Write "<option value='" & RSR(0) & "' selected>" & RSR(1) & "</option>"
			  Else
			   Response.Write "<option value='" & RSR(0) & "'>" & RSR(1) & "</option>"
			  End If
			  RSR.MoveNext
			 Loop
			 RSR.CLose
			 Set RSR=Nothing
			Response.Write "</select>              　　<font color=""green"">请选择该项管理员的角色</font>"
		
		Response.Write "          </td></tr>"

				 
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>是否锁定：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
					
					If SuperTF = "1" Then
					   Response.Write ("<input type=""hidden"" value=""0"" name=""locked""> (否)正常")
					 ElseIf Locked = "1" Then
					 Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Locked"" value=""0""> 正常 ")
					 Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Locked"" value=""1"" checked> 锁定 ")
					 Else
					  Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Locked"" value=""0"" checked> 正常 ")
					  Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Locked"" value=""1""> 锁定 ")
					 End If
					  
		Response.Write "              　　<font color=""#FF0000"">锁定的用户不能登录后台管理</font></td>"
		Response.Write "          </tr>"
		
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>真实姓名：</td>"
		Response.Write "            <td><input name=""RealName"" type=""text"" class=""textbox"" value=""" & RealName & """ id=""RealName"" size=""30""></td>"
		Response.Write "            <td align=""right"" class='clefttitle'>性 别：</td>"
		Response.Write "            <td>"
					 
					 If Trim(Sex) = "女" Then
						  Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Sex"" value=""男""> 男 ")
						  Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Sex"" value=""女"" checked>  女 ")
					  Else
						  Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Sex"" value=""男"" checked> 男 ")
						  Response.Write ("&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""Sex"" value=""女"">  女 ")
					  End If
				   
		Response.Write "             </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>联系电话：</td>"
		Response.Write "            <td><input name=""TelPhone"" type=""text"" class=""textbox"" value=""" & TelPhone & """ id=""TelPhone"" size=""30""></td>"
		Response.Write "            <td align=""right"" class='clefttitle'>电子信箱：</td>"
		Response.Write "            <td><input name=""Email"" type=""text"" class=""textbox"" id=""Email"" value=""" & Email & """ size=""30""></td>"
		Response.Write "          </tr>"
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>简要说明：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		Response.Write "              <textarea class='textbox' name=""Description"" rows=""6"" id=""Description"" style=""width:80%;height:80px;border-style: solid; border-width: 1"">" & Descript & "</textarea></td>"
		Response.Write "          </tr>"
		Response.Write "</table>"
		Response.Write "  </form>"
		Response.Write "  </div>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "<!--" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{  var form=document.AdminForm;" & vbCrLf
		Response.Write "   if (form.UserName.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     top.$.dialog.alert(""请输入管理员名称!"",function(){"
		Response.Write "     form.UserName.focus();});"
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
			
			If Action <> "Edit" Then
		Response.Write "   if ($('input:radio[name=PreAddType]:checked').val()==0 && form.PrUserName.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     top.$.dialog.alert(""请输入前台注册用户名称!"",function(){"
		Response.Write "     form.PrUserName.focus();});"
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf

		Response.Write "    if (form.PassWord.value=='')"
		Response.Write "    {"
		Response.Write "     top.$.dialog.alert(""请输入初始密码!"",function(){"
		Response.Write "     form.PassWord.focus();});"
		Response.Write "     return false;"
		Response.Write "    }"
		Response.Write "   else if (form.PassWord.value.length<6)"
		Response.Write "    {"
		Response.Write "     top.$.dialog.alert(""初始密码不能少于6位!"",function(){"
		Response.Write "     form.PassWord.focus();});"
		Response.Write "     return false;"
		Response.Write "    }"

		
			End If
			
		Response.Write "   if (form.RealName.value=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "     top.$.dialog.alert(""请输入真实姓名"",function(){" & vbCrLf
		Response.Write "     form.RealName.focus();});" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   if (form.Email.value!='')" & vbCrLf
		Response.Write "   if(is_email(form.Email.value)==false)" & vbCrLf
		Response.Write "      { top.$.dialog.alert('非法电子邮箱!',function(){" & vbCrLf
		Response.Write "        form.Email.focus();});" & vbCrLf
		Response.Write "        return false;" & vbCrLf
		Response.Write "     }"
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "//-->" & vbCrLf
		Response.Write "</Script>"
		End Sub
		
		Sub AdminSave()
			Dim AdminID, GroupID, UserName,PrUserName, PassWord, ConPassWord, Locked, RealName, Sex, TelPhone, Email, Descript, TrueIP
			Dim TempObj, AdminRS, AdminSql,ComeUrl,PreAddType
			ComeUrl=Request.ServerVariables("HTTP_REFERER")
			AdminID = KS.G("AdminID")
			GroupID = KS.ChkClng(KS.G("GroupID"))
			UserName = KS.R(KS.G("UserName"))
			PrUserName=KS.R(KS.G("PrUserName"))
			PassWord = KS.G("PassWord")
			PreAddType=KS.ChkClng(KS.G("PreAddType"))

			IF PrUserName="" And PreAddType=0 Then Call KS.Alert("前台注册用户名必须填写!",ComeUrl)
			
			PassWord = MD5(KS.R(PassWord),16)
			Locked = KS.G("Locked")
			RealName = KS.R(KS.G("RealName"))
			Sex = KS.G("Sex")
			TelPhone = KS.R(KS.G("TelPhone"))
			Email = KS.R(KS.G("Email"))
			Descript = KS.R(KS.G("Description"))
			TrueIP = KS.GetIP
			If UserName <> "" Then
					If Len(UserName) >= 100 Then
						Call KS.AlertHistory("管理员名称不能超过50个字符!", -1)
						Set KS = Nothing
						Response.End
					End If
			 Else
					Call KS.AlertHistory("请输入管理员名称!", -1)
					Set KS = Nothing
					Response.End
			 End If
			 If GroupID=0 Then
					Call KS.AlertHistory("请选择管理员角色!", -1)
					Response.End
			 End If
			
			   
			If Request("Action") = "Add" Then
					Set TempObj = Conn.Execute("Select UserName from [KS_Admin] where UserName='" & UserName & "'")
					If Not TempObj.EOF Then
					    KS.Die "<script>top.$.dialog.alert('数据库中已存在该管理员名称!',function(){history.back(-1);});</script>"
						Set KS = Nothing
						Response.End
					End If
					If PreAddType=0 Then
						Set TempObj = Conn.Execute("Select top 1 UserName from [KS_User] where UserName='" & PrUserName & "'")
						If TempObj.BOf And TempObj.EOF Then
							 KS.Die "<script>top.$.dialog.alert('找不到此前台注册用户!',function(){history.back(-1);});</script>"
							 Set KS = Nothing:Response.End
						End If
						IF not Conn.Execute("Select adminid From KS_Admin Where PrUserName='" & PrUserName & "'").eof Then
							KS.Die "<script>top.$.dialog.alert('您填写的前台注册用户已经是管理员了，不能再添加!',function(){history.back(-1);});</script>"
							 Set KS = Nothing:Response.End
						End IF
					Else
					
					IF not Conn.Execute("Select userid From KS_User Where UserName='" & UserName & "'").eof Then
							KS.Die "<script>top.$.dialog.alert('您填写的用户名已经存在，不能再添加!',function(){ history.back(-1);});</script>"
							 Set KS = Nothing:Response.End
					End IF
					
					 Set AdminRS = Server.CreateObject("adodb.recordset")
					  AdminSql = "select top 1 * from [KS_User] Where 1=0"
					  AdminRS.Open AdminSql, Conn, 1, 3
					  AdminRS.AddNew
					  AdminRS("RegDate") = Now
					  AdminRS("JoinDate") = Now
					  AdminRS("UserName") = UserName
					  AdminRS("PassWord") = PassWord
					  AdminRS("Locked") = 0
					  AdminRS("RealName") = RealName
					  AdminRS("Sex") = Sex
					  AdminRS("HomeTel") = TelPhone
					  AdminRS("Mobile") = TelPhone
					  AdminRS("Email") = Email
					  AdminRS("GroupID") = 1
					  AdminRS("LoginTimes") = 0
					  AdminRS("LastLoginIP") = KS.GetIP
					  AdminRS("BeginDate") = Now
					  AdminRS("LastLoginTime")=Now
					  AdminRS("Money")=0
					  AdminRS("Point")=0
					  AdminRS("Score")=0
					  AdminRS.Update
					  AdminRS.Close:Set AdminRS = Nothing
					  PrUserName=UserName
				End If
				
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select top 1 * from [KS_Admin] Where (AdminID IS NULL)"
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS.AddNew
				  AdminRS("AddDate") = Now
				  AdminRS("UserName") = UserName
				  AdminRS("PrUserName")=PrUserName
				  AdminRS("PassWord") = PassWord
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("TelPhone") = TelPhone
				  AdminRS("Email") = Email
				  AdminRS("Description") = Descript
				  AdminRS("SuperTF") = 0
				  AdminRS("LastLoginIP") = TrueIP
				  AdminRS("LastLoginTime") = Now
				  AdminRS("LastLogOutTime") = Now
				  AdminRS("LoginTimes") = 0
				  AdminRS("GroupID")=GroupID
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing
				  
				  '更新前台用户，使之加入管理员组
				  Conn.Execute("Update KS_User Set GroupID=1 Where UserName='" & PrUserName & "'")
				  
				  Response.Write ("<script>top.$.dialog.confirm('添加管理员成功,继续添加吗?',function(){location.href='user/KS.Admin.asp?Action=Add';} ,function(){ location.href='user/KS.Admin.asp';});</script>")
			ElseIf Request("Action") = "Edit" Then
					Set TempObj = Conn.Execute("Select UserName from [KS_Admin] where AdminID<>" & AdminID & " And UserName='" & UserName & "'")
					If Not TempObj.EOF Then
						Call KS.AlertHintScript("数据库中已存在该管理员名称!")
						 Set KS = Nothing
						Response.End
					End If
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select top 1 * from [KS_Admin] Where AdminID=" & AdminID
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS("UserName") = UserName
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("TelPhone") = TelPhone
				  AdminRS("Email") = Email
				  AdminRS("GroupID")=GroupID
				  AdminRS("Description") = Descript
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing
				  Response.Write ("<script>top.$.dialog.alert('修改管理员成功!',function(){ location.href='user/KS.Admin.asp';});</script>")
			End If
			
			
		End Sub
        
		'设置管理员密码
		Sub SetAdminPass()
		Response.Write "<!DOCTYPE html>"
        Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link href=""../Include/admin_style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/CheckPassWord.js""></script>"
		Response.Write "<style>"
		Response.Write ".rank { border:none; background:url(../../images/rank.gif) no-repeat; width:136px; height:22px; vertical-align:middle; cursor:default; }"
		Response.Write ".r0 { background-position:0 0; }"
		Response.Write ".r1 { background-position:0 -22px; }"
		Response.Write ".r2 { background-position:0 -44px; }"
		Response.Write ".r3 { background-position:0 -66px; }"
		Response.Write ".r4 { background-position:0 -88px; }"
		Response.Write ".r5 { background-position:0 -110px; }"
		Response.Write ".r6 { background-position:0 -132px; }"
		Response.Write ".r7 { background-position:0 -154px; }"
		Response.Write "</style>"
		Response.Write "<title>设置管理员密码</title>"
		Response.Write "</head>"
		
		Dim AdminID, UserName, PassWord,RSObj
		AdminID = Request("AdminID")
		If AdminID <> "" Then
		   Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		   RSObj.Open "Select top 1 * From KS_Admin Where AdminID=" & AdminID, Conn, 1, 3
		   If Not RSObj.EOF Then UserName = Trim(RSObj("UserName"))
		   RSObj.Close:Set RSObj = Nothing
		Else
		  UserName=KS.C("AdminName")
		End If
		
		     If Request("Flag") = "SetOK" Then
			   If Trim(Request.Form("PassWord")) <> Trim(Request.Form("ConPassWord")) Then
				Response.Write ("<Script>alert('确认密码有误!');history.back(-1);</script>")
				Exit Sub
				Response.End
			   Else
			    Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		         RSObj.Open "Select * From KS_Admin Where UserName='" & UserName & "'", Conn, 1, 3
				 RSObj("PASSWord") = MD5(KS.R(Trim(KS.S("PassWord"))),16)
				 RSObj.Update
				 Dim PrUserName:PrUserName=RSObj("PrUserName")
				  RSObj.Close: Set RSObj = Nothing
				  If UserName=KS.C("UserName") Then  Response.Cookies(KS.SiteSn)("AdminPass")=MD5(KS.R(Trim(KS.S("PassWord"))),16)
				  
				  If KS.ChkClng(request("UpdateUserPass"))=1 Then
				    Conn.Execute("Update KS_User Set [PassWord]='" &MD5(KS.R(Trim(KS.S("PassWord"))),16) &"' Where UserName='" & PrUserName & "'")
					Response.Cookies(KS.SiteSn)("Password")=MD5(KS.R(Trim(KS.S("PassWord"))),16)
				  End If
				  
				 Response.Write ("<Script>alert('密码设置成功!!!');top.box.close();</script>")
			   End If
			 End If
			 
		Response.Write "<body style=""background: #fff;"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "  <form action=""?Action=SetPass"" name=""AdminForm"" method=""post"">"
		Response.Write "   <input name=""Flag"" type=""hidden"" id=""Flag"" value=""SetOK"">"
		Response.Write "   <input name=""AdminID"" type=""hidden"" value=""" & AdminID & """><br>"
		Response.Write "  <table width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td>"
		Response.Write "      <FIELDSET align=center>" & vbCrLf
		Response.Write "        <LEGEND align=left>设置密码</LEGEND>" & vbCrLf
		Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
		Response.Write "          <tr>"
		Response.Write "            <td width=""179"" height=""35"" align=""center""> <div align=""center"">管理员名</div></td>" & vbCrLf
		Response.Write "            <td width=""542"" align=""left"">"
		Response.Write ("<input name=""UserName"" Readonly value=""" & UserName & """ type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
		Response.Write "              用于登录后台的名称</td>" & vbCrLf
		Response.Write "          </tr>"
		
		Response.Write "          <tr>"
		Response.Write "            <td height=""35"" align=""center""> <div align=""center"">新 密 码</div></td>"
		Response.Write "           <td align=""left""> <input name=""PassWord"" type=""password"" size=""34"" class=""textbox"" onkeyup=""javascript:if(this.value!=''){document.getElementById('ps').style.display='';setPasswordLevel(this, document.getElementById('passwordLevel'))}"" onmouseout=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"" onblur=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"">" & vbCrLf
		Response.Write "              不少于6位 </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr style='display:none' id='ps'>"
		Response.Write "             <td align=""center"" height=""25"">密码强度</td>"
		Response.Write "                <td align='left'><input name=""Input"" disabled=""disabled"" class=""rank r0"" id=""passwordLevel"" /></td>"
		Response.Write "           </tr>"
		Response.Write "          <tr>"
		Response.Write "            <td height=""35"" align=""center"">确定密码</td>" & vbCrLf
		Response.Write "            <td align=""left""> <input name=""ConPassWord""  type=""password"" size=""34"" class=""textbox"">" & vbCrLf
		Response.Write "              同上</td>"
		Response.Write "          </tr>"
		
		Response.Write "          <tr>"
		Response.Write "            <td height=""35"" align=""center"">更新前台密码</td>" & vbCrLf
		Response.Write "            <td align=""left""> <label><input name=""UpdateUserPass""  type=""checkbox"" value=""1"" checked>是</label> <font color=red>如果选择是,那么前台会员中心的登录密码将和后台的一致</font></td>"
		Response.Write "          </tr>"
		
		Response.Write "        </table>"
		Response.Write "         </FIELDSET>" & vbCrLf
		Response.Write "       </td>" & vbCrLf
		Response.Write "    </tr>"
		Response.Write "    </table>" & vbCrLf
		Response.Write "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
		Response.Write "        <input type=""button"" class='button' name=""Submit"" Onclick=""CheckForm()"" value="" 确 定 "">"
		Response.Write "        <input type=""button"" class='button' name=""Submit2"" onclick=""top.box.close()"" value="" 取 消 "">" & vbCrLf
		Response.Write "      </td>" & vbCrLf
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write " &nbsp;<font color=green>PS:为了您的系统安全，您应该不定期的修改后台登录密码!</font>"
		Response.Write "  </form>" & vbCrLf
		Response.Write "</body>"
		Response.Write "</html>" & vbCrLf
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "<!--" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.AdminForm;" & vbCrLf
		Response.Write "    if (form.PassWord.value=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "     alert(""请输入新密码!"");" & vbCrLf
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    else if (form.PassWord.value.length<6)" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "      alert(""初始密码不能少于6位!"");"
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   if (form.ConPassWord.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请输入确定密码!"");" & vbCrLf
		Response.Write "     form.ConPassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   else if(form.ConPassWord.value.length<6)" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""确定密码不能少于6位!"");" & vbCrLf
		Response.Write "     form.ConPassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "   if (form.PassWord.value!=form.ConPassWord.value)" & vbCrLf
		Response.Write "    {"
		Response.Write "    alert(""两次输入的密码不一致!"");" & vbCrLf
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}"
		Response.Write "//-->"
		Response.Write "</Script>"
		End Sub
		
		'删除管理员
		Sub AdminDel()
		Dim k, AdminID,RSObj
		AdminID = Trim(KS.G("AdminID"))
		AdminID = Split(AdminID, ",")
		For k = LBound(AdminID) To UBound(AdminID)
			   Set RSObj = Conn.Execute("Select SuperTF,PrUserName From KS_Admin Where  AdminID=" & AdminID(k))
			   If Not RSObj.EOF Then
				 If RSObj("SuperTF") = 1 Then
				  Response.Write ("<script>alert('系统默认管理员不能删除!');location.href='KS.Admin.asp';</script>")
				 Else
				  '更改前台注册会员，使之变为注册会员身份
				  Conn.Execute("Update KS_User Set GroupID=3 Where UserName='" & RSObj("PrUserName") & "'")
				  Conn.Execute("Delete From KS_Admin Where AdminID =" & AdminID(k))
				 End If
			  End If
			  RSObj.Close
		Next
		Set RSObj = Nothing
		Response.Write ("<script>location.href='KS.Admin.asp';</script>")
		End Sub
		'删除角色
		Sub DeleteRole()
		 Dim ID:ID=KS.ChkCLng(KS.G("ID"))
		 If ID=0 Then KS.AlertHintScript "error!":KS.Die ""
		 Conn.Execute("Update KS_User Set GroupID=3 Where UserName in(select PrUserName from ks_admin Where GroupID=" & ID & ")")
		 Conn.Execute("Delete From ks_admin Where GroupID=" & ID)
		 Conn.Execute("Delete From ks_usergroup Where ID=" & ID)
		Response.Write ("<script>location.href='KS.Admin.asp?action=Role';</script>")
		End Sub
		
		
 Sub ClassList(ChannelID)
 %>
 <div style="border: 5px solid #E7E7E7;height:150px; overflow: auto; width:100%"> 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <% 
					  Dim Node, CheckStr,SpaceStr,TJ,k  
				      KS.LoadClassConfig
					  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks12=" & ChannelID&"]")                     
	                  if KS.ChkClng(KS.G("ID"))<>0 and KS.FoundInArr(Node.SelectSingleNode("@ks16").text,KS.ChkClng(KS.G("ID")),",")=true then CheckStr=" checked"
					  SpaceStr="&nbsp;&nbsp;&nbsp;&nbsp;"
					  TJ=Node.SelectSingleNode("@ks10").text
					  If TJ>1 Then
						 For k = 1 To TJ - 1
							SpaceStr = SpaceStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
						 Next
					  End If
					  %>
                          
                          <tr> 
                            <td><table border="0" cellspacing="0" cellpadding="0">
                                <tr align="left" class="TempletItem"> 
                                  <td><%=SpaceStr%></td>
                                  <td>
                                  <input name="AdminPurview" type="checkbox" value="<% =Node.SelectSingleNode("@ks0").text %>"<%=CheckStr%>> 
                                  <%if tj>1 then%>
                                    <% = Node.SelectSingleNode("@ks1").text %>      
                                  <%else%>
                                    <strong><% = Node.SelectSingleNode("@ks1").text %></strong>      
                                  <%end if%>
                                                                  
								 </td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	                     CheckStr = ""
	                Next
					   %>
                        </table>
                      </div>
 <%
 End Sub
 
 Sub BasePurview(PowerList,SQL,I)
 %>
     <tr> 
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10002"<%if InStr(1, PowerList,"M" & SQL(0,I)&"10002" ,1)<>0 then Response.Write( " checked") %>>
                      添加<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10003"<%if InStr(1, PowerList,"M"&SQL(0,I)&"10003" ,1)<>0 then Response.Write( " checked") %>>
                      编辑<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10004"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10004" ,1)<>0 then Response.Write( " checked") %>>
                      删除<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10010"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10010" ,1)<>0 then Response.Write(" checked") %>>
审核<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10006"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10006" ,1)<>0 then Response.Write(" checked") %>>
                      加入专题</td>
                    <td width="13%"> <input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10007"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10007" ,1)<>0 then Response.Write(" checked") %>>
                      加入JS</td>
                  </tr>
                  <tr> 
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10008"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10008" ,1)<>0 then Response.Write( " checked") %>> 栏目管理</td>
                 <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10005"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10005" ,1)<>0 then Response.Write(" checked") %>> 回 收 站</td>

                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10009"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10009" ,1)<>0 then Response.Write(" checked") %>> 上传文件</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10011"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10011" ,1)<>0 then Response.Write(" checked") %>> 复制粘贴</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20005" ,1)<>0 then Response.Write(" checked") %>> 发布<%=sql(3,i)%></td>
                  </tr>
<%
 End Sub

 
 '从数组中删除指定项
 Function DelItemInArr(byval str,byval delstr,spliter)
   if KS.IsNul(str) Then Exit Function
   Dim strA:strA=Split(str,spliter)
   Dim I,temp,newstr
   For I=0 To Ubound(Stra)
      If lcase(strA(i))<>lcase(delstr) Then
	    if newstr="" then
		 newstr=stra(i)
		else
		 newstr=newstr & spliter & stra(i)
		end if
	  End If
   Next
   DelItemInArr=newstr
 End Function
 
End Class
%> 

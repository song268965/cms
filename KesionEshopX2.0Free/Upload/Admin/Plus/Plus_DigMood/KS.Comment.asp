<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_Comment
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Comment
        Private KS,ChannelID,Page,KSCls,Action,TableXML,Table
		Private I, totalPut, CurrentPage, SqlStr,InfoID, ClassID,ProjectID
        Private RSObj,MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 18
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		   InfoID = KS.G("InfoID")
		   ClassID = KS.G("ClassID")
		   If Trim(ClassID) = "" Then ClassID = "0"
		   If ClassID <> "0" Then ClassID = "'" & Replace(ClassID, ",", "','") & "'"
		   If InfoID = "" Then InfoID = "0"
		   If InfoID <> "0" Then  InfoID = "'" & Replace(InfoID, ",", "','") & "'"
           Page = KS.G("Page")
		   ChannelID=KS.ChkClng(KS.G("ChannelID"))
		   ProjectID=KS.ChkClng(Request("ProjectID"))
		   
			If Not KS.ReturnPowerResult(0, "M010002") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End iF
			
			Table=KS.G("Table")
			If KS.IsNul(Table) Then Table="KS_Comment"
		
			 Action=KS.G("Action")
			 Select Case Action
			  Case "View"  Call CommentView()
			  Case "Verific" Call CommentVerific()
			  Case "Del" Call CommentDel()
			  Case "DelAllRecord" Call DelAllRecord()
			  Case "DelUnVerify" Call DelUnVerify()
			  Case "PostTable" Call PostTable()
			  Case "DoPostTableSave" Call DoPostTableSave()
			  Case "DoDelPostTable" Call DoDelPostTable()
			  Case "DoPostTableModifySave" Call DoPostTableModifySave()
			  Case Else	 Call CommentList()
			 End Select
		
		End Sub
		
		Sub DelUnVerify()
		  Conn.Execute("Delete From " & Table &" Where verific=0")
		  KS.Die "<script>alert('恭喜，一键清除了所有未审核的评论!');location.href='KS.Comment.asp?Table=" & Table &"&ProjectID=" & ProjectID &"';</script>"
		End Sub
		
'评论数据表管理
Sub PostTable()
		set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		TableXML.async = false
		TableXML.setProperty "ServerHTTPRequest", true 
		TableXML.load(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
%>
<!DOCTYPE html><html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<title>评论数据表管理</title>
<link href='../../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script src="../../../KS_Inc/Jquery.js"></script>
</head>
<body>

<div class="pageCont2 mt20">
	<div class="tabTitle">评论数据表管理</div>
	  <form name="myform" action="KS.Comment.asp?action=DoPostTableModifySave" method="post">
	  <table width="100%" align='center' border="0" cellpadding="0" cellspacing="0">
      <tr class="sort">
	    <td>序号</td>
	    <td>表名称</td>
	    <td>类型</td>
		<td>当前默认</td>
		<td>记录数</td>
		<td>说明</td>
		<td>管理操作</td>
	  </tr>
<%
  If TableXML.DocumentElement.SelectNodes("item").length=0 Then
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>您没有添加评论数据表!</td></tr>"
  Else
	  Dim Node,N:N=0
	  For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  %>
			  <tr  onmouseout="this.className='list'" onMouseOver="this.className='listmouseover'">               
			   <td class='splittd' height="30" align="center"><%=Node.SelectSingleNode("@id").text%></td>
			   <td class='splittd' height="30"><%=Node.SelectSingleNode("tablename").text%></td>
			   <td class='splittd' style='text-align:center' height="30"><%
			   if Node.SelectSingleNode("@issys").text="1" then
			    response.write "<span style='color:red'>系统</span>"
			   else
			    response.write "<span style='color:green'>自定义</span>"
			   end if
			   %></td>
			   <td class='splittd' align="center">
			   <%
				 if node.selectSingleNode("@isdefault").text="1" then
				  response.write "<input type='radio' name='isdefault' value='" & Node.SelectSingleNode("@id").text & "' checked>"
				 else
				  response.write "<input type='radio' name='isdefault' value='" & Node.SelectSingleNode("@id").text & "'>"
				 end if
				%>
			   </td>
			   <td class='splittd' align="center">
			   <%
			     dim num
				 num=conn.execute("select count(1) from " & Node.SelectSingleNode("tablename").text)(0)
				 response.write "<font color='#ff6600'>" & num & "</font>"
			   %>
			   </td>
			   <td class='splittd' align="center">
			   <%=Node.SelectSingleNode("descript").text%>
			   </td>
			   
			   <td class='splittd' align="center">
			    <%if node.selectSingleNode("@isdefault").text="1" or num>0 or Node.SelectSingleNode("@issys").text="1" then%>
				 <span style="color:#999999">删除</span>
				<%else%>
				 <a href="?action=DoDelPostTable&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('确定删除该任务吗?'))">删除</a>
				<%end if%>
			   </td>
			  </tr>
	  <%
		n=n+1
	  Next
  End If
  %>
		
	  </table>
       <br/>
	   <div style="text-align:center">
	    <input name="Submit" type="submit"  class="button" value="批量设置">
		
	   </div>
	 </form>
	   <br/>
       
	   
	   <script type="text/javascript">
	    function check(){
		 var tobj=$("#TableName");
		 if (tobj.val()==''){
		  alert('请输入数据表名!');
		  tobj.focus();
		  return false;
		 }
		 if (tobj.val().toLowerCase().indexOf('ks_comment_')==-1){
		  alert('数据表名必须与KS_Comment_开头!');
		  tobj.focus();
		  return false;
		 }
		 return true;
		}
	   </script>
       
	   <form name="myform" action="KS.Comment.asp?action=DoPostTableSave" method="post" id="myform">
	  <table width='99%' style="margin:4px" align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>数据表名称:</strong></td>
		   <td><input type="text" class="textbox" name="TableName" id="TableName" value="KS_Comment_"> 
		   如:KS_Comment_1,KS_Comment_2等,必须以KS_Comment_开头</td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>数据表说明:</strong></td>
		   <td><input type="text" class="textbox" name="Descript" size="50" id="Descript" > 简单描述该表的用途</td>
		  </tr>
		  <tr class='tdbg'>               
		   <td height="30" align='right' colspan="2" >
		    <input type="submit" value="确定新增" onClick="return(check())" class="button"/>
			<label><input type="checkbox" name="isdefault" value="1" checked>设置成当前数据表</label>
		   </td>
		  </tr>
	  </table>
	   </form>	  
	</div>
		  
	   <div class="attention">
<strong>特别提醒：</strong><br/>
1、数据表中选中的为当前评论系统使用来保存评论数据的表，一般情况下每个表中的数据越少评论的显示速度越快，当您上列单个评论数据表中的数据有超过几万的数据时不妨新添一个数据表来保存评论数据,您会发现评论系统的速度快很多很多。<br/>
2、您也可以将当前所使用的数据表在上列数据表中切换，当前所使用的评论数据表即当前评论用户发贴时默认的保存的评论数据表。<br/>
3、以免出错，当前正在使用的数据表、已有记录的数据表或是系统自带数据表不允许删除。
</div>
</body>
</html>
<%
End Sub
		
'保存添加数据表
Sub DoPostTableSave()
set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
TableXML.async = false
TableXML.setProperty "ServerHTTPRequest", true 
TableXML.load(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))

Dim TableName:TableName=KS.G("TableName")
 Dim Descript:Descript=Request.Form("Descript")
 Dim Node,isdefault:isdefault=KS.ChkClng(Request.Form("isdefault"))
 If Len(TableName)<12 or lcase(left(TableName,11))<>"ks_comment_" Then
  Call KS.AlertHintScript ("数据表格式不正确!")
 End If
 Dim ItemID
 '取得唯一ID号
 If TableXML.DocumentElement.SelectNodes("item").length<>0 Then
   ItemID=TableXML.DocumentElement.SelectNodes("item").length+1
 Else
   ItemID=1
 End If
 

 Dim sql:sql="CREATE TABLE ["&TableName&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&TableName&" PRIMARY KEY,"&_
                        "ChannelID int default 0,"&_
                        "InfoID int default 0,"&_
						"UserIP varchar(100),"&_
						"Content ntext,"&_
						"Score int default 0,"&_
						"OScore int default 0,"&_
						"AddDate datetime,"&_
						"Verific tinyint default 0,"&_
						"Anonymous int default 0,"&_
						"UserName nvarchar(50),"&_
						"Point int default 0,"&_
						"quotecontent ntext,"&_
						"ReplyContent ntext,"&_
						"ReplyTime datetime,"&_
						"ReplyUser nvarchar(50),"&_
						"ProjectID int default 0,"&_
						"Title nvarchar(255),"&_
						"M0 int default 0,"&_
						"M1 int default 0,"&_
						"M2 int default 0,"&_
						"M3 int default 0,"&_
						"M4 int default 0,"&_
						"M5 int default 0,"&_
						"M6 int default 0,"&_
						"M7 int default 0,"&_
						"M8 int default 0,"&_
						"M9 int default 0,"&_
						"M10 int default 0,"&_
						"M11 int default 0,"&_
						"M12 int default 0,"&_
						"M13 int default 0,"&_
						"M14 int default 0"&_
						")"
 Conn.Execute sql
 On Error Resume Next
 Conn.Execute("CREATE INDEX [ChannelId] ON " & TableName & "([ChannelId])")
 Conn.Execute("CREATE INDEX [InfoId] ON " & TableName & "([InfoId])")

 If Err Then Err.Clear
 
     Dim NodeStr,brstr
     brstr=chr(13)&chr(10)&chr(9)
     NodeStr="<item isdefault=""0"" id=""" & ItemID &""" issys=""0"">" &brstr
	 NodeStr=NodeStr & "<tablename>" & TableName & "</tablename>"&brstr
	 NodeStr=NodeStr & "<descript><![CDATA[ " & descript & "]]></descript>" & brstr
	 NodeStr=NodeStr & " </item>"&brstr
	 Dim XML2:set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
     XML2.LoadXml(NodeStr)
	 Dim NewNode:set NewNode=XML2.documentElement
	 
	 Dim TN:Set TN=TableXML.DocumentElement
	 TN.appendChild(NewNode)
	 TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
	 If isdefault=1 Then
	  For Each Node In TableXML.DocumentElement.SelectNodes("item")
		 If KS.ChkClng(ItemID)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
			Node.Attributes.getNamedItem("isdefault").text=1
		 Else
			Node.Attributes.getNamedItem("isdefault").text=0
		 End If
	  Next
	  TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
	 End If
	 
	 Response.Write "<script> alert('恭喜,评论数据表添加成功!');location.href='?action=PostTable'</script>"
End Sub
'删除评论数据表
Sub DoDelPostTable()
	set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
  Dim ItemID:ItemID=KS.ChkClng(Request("itemid"))
  If ItemID=0 Then KS.AlertHintScript "对不起,参数出错!"
  Dim DelNode,Node,ID
  Set DelNode=TableXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
  If DelNode Is Nothing  Then
   KS.AlertHintScript "对不起,参数出错!"
  End If
  '删除表
  Conn.Execute ("Drop Table "&delnode.selectsinglenode("tablename").text&"")
  
  TableXML.DocumentElement.RemoveChild(DelNode)
  
  '更新比当前任务ID大的ID号,依次减一
  For Each Node In TableXML.DocumentElement.SelectNodes("item")
     ID=KS.ChkClng(Node.SelectSingleNode("@id").text)
	 If ID>ItemID Then
	    Node.SelectSingleNode("@id").text=ID-1
	 End If
  Next
  '保存
  TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
  KS.AlertHintScript "恭喜,评论数据表已删除!"
End Sub

'批量设置
Sub DoPostTableModifySave()
	set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
  Dim Node,id,isdefault:isdefault=KS.ChkClng(Request.Form("isdefault"))

  For Each Node In TableXML.DocumentElement.SelectNodes("item")
     ID=KS.ChkClng(Node.SelectSingleNode("@id").text)
	 If ID=isdefault Then
	    Node.Attributes.getNamedItem("isdefault").text=1
     Else
	    Node.Attributes.getNamedItem("isdefault").text=0
	 End If
  Next

	 TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
	 KS.AlertHintScript "恭喜,设置成功!"

End Sub
		
Sub DelAllRecord()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_Y & ",AddDate," & SqlNowString & ")>=1"
		   Case 7 Param="datediff(" & DataPart_Y & ",AddDate," & SqlNowString & ")>=2"
		  End Select
		  If ProjectID<>0 Then Param=Param &" and ProjectID=" & ProjectID
   		  If Param<>"" Then 
		   	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select UserName,Anonymous,ID,Content,AddDate,channelid,infoid From " & Table &" Where " & Param,conn,1,1
			 Do While Not RS.Eof
			  Call ProcessUserScore(RS)
			  RS.MoveNext
			 Loop
			 RS.Close:Set RS=Nothing
		     Conn.Execute("Delete From KS_Comment Where " & Param)
		  End If
		  KS.echo "<script src=""../../../ks_inc/jquery.js""></script>"
          KS.echo "<script>$(parent.document).find('#ajaxmsg').toggle();alert('恭喜,删除指定日期评论成功!');location.href='KS.Comment.asp?Table=" & Table &"&ProjectID=" & ProjectID &"';</script>"
		 End Sub
		
        Sub CommentList
		If Request("page") <> "" Then
			  CurrentPage = KS.chkclng(Request("page"))
		Else
			  CurrentPage = 1
		End If
	With KS
	  .echo "<!DOCTYPE html><html>"
	  .echo "<head>"
	  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
	  .echo "<title>评论管理</title>"
	  .echo "<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
	  .echo "<script language=""JavaScript"">" & vbCrLf
	  .echo "var ChannelID=""" & ChannelID & """;" & vbCrLf
	  .echo "var Page='" & CurrentPage & "';" & vbCrLf
	  .echo "var InfoID=""" & InfoID & """;" & vbCrLf
	  .echo "var ClassID=""" & ClassID & """;" & vbCrLf
	  .echo "</script>" & vbCrLf
	
	  .echo "<script src=""../../../KS_Inc/common.js""></script>" & vbCrLf
	  .echo "<script src=""../../../KS_Inc/JQuery.js""></script>" & vbCrLf
%>
	<script language="javascript">
	function set(v){
	 if (v==1)
	  Verific(1,0);
	 else if (v==2)
	  Verific(0,0);
	 else if(v==3)
	  DelComment();
	}
	function Verific(OpType,CommentID)
	{
	if (CommentID==0) 
	 {
	 var ids=get_Ids(document.myform);
	if (ids!='')
	 {
	       $("#action").val("Verific");
		   $("#VerificType").val(OpType);
		   $("#myform").submit();
	 }	
	else
	 alert('请选择评论!');
	 }
	 else
	   location.href="KS.Comment.asp?Table=<%=Table%>&ProjectID=<%=ProjectID%>&Action=Verific&ChannelID="+ChannelID+"&VerificType="+OpType+"&InfoID="+InfoID+"&ClassID="+ClassID+"&Page="+Page+"&ID="+CommentID;
	}
	function DelComment()
	{
		var ids=get_Ids(document.myform);
		if (ids!=''){ 
	     if (confirm('真的要删除选中的评论吗?'))
		   $("#action").val("Del");
		   $("#myform").submit();
		}
		else{ alert('请选择要删除的评论!');}
	}
	function CommentDataBase(){
	       location.href="KS.Comment.asp?ChannelID="+ChannelID+"&Action=PostTable&InfoID="+InfoID+"&ClassID="+ClassID+"&Page="+Page;
	}
	function GetKeyDown()
	{ 
	if (event.ctrlKey)
	  switch  (event.keyCode)
	  {  case 90 : location.reload(); break;
		 case 65 : Select(0);break;
		 case 86 : event.keyCode=0;event.returnValue=false;ViewComment(0); break;
		 case 83 : event.keyCode=0;event.returnValue=false;Verific(1,0);break;
		 case 67 : event.keyCode=0;event.returnValue=false;Verific(0,0);break;
		 case 68 : DelComment();break;
	   }	
	else	
	 if (event.keyCode==46)DelComment();
	}
</script>
<%
	  .echo "</head>"
	 .echo "<body topmargin=""0"" leftmargin=""0"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
	  .echo "<ul id='menu_top' class='menu_top_fixed'>"
	  .echo "<li onclick='javascript:Verific(1,0);' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon audit'></i>审核评论</span></li>"
	  .echo "<li onclick='Verific(0,0);' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>取消审核</span></li>"
	  .echo "<li onclick='DelComment()' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon delete'></i>删除评论</span></li>"
	  .echo "<li onclick='CommentDataBase()' class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon merge'></i>评论数据表</span></li>"
	  .echo "<li></li><form action='?table=" & KS.G("Table") &"' method='post'><div class='quicktz'>条件:<select name='searchtype'><option value='1'>文档标题</option><option value='2'>评论者</option><option value='3'>评论内容</option></select>关键字:<input type='text' class='textbox' name='keyword'> <input class='button' type='submit' value=' 搜 索 '></div></form></ul>"
	  .echo "<div class=""menu_top_fixed_height""></div>"
	  .echo "<div class=""pageCont2"">"
	  
	  if ProjectID<>0 then
	   dim rst:set rst=conn.execute("select top 1 * from KS_MoodProject where id=" & ProjectID)
	   if not rst.eof then
	    .echo "<div style='font-size:14px;margin:10px;font-weight:bold'>查看点评项目[<font color=green>" & RST("ProjectName") &"</font>]的所有用户点评&nbsp;<input type='button' class='button' value='返回点评管理中心' onclick=""location.href='KS.Mood.asp?typeflag=1'""/></div>"
	   end if
	   rst.close:set rst=nothing
	 else
	  
   %>
   <div class='tabTitle'>评论管理</div>
   <div style="height:30px;line-height:30px; margin-bottom:10px;">
   <strong>选择数据表：</strong><select name="table" onChange="location.href='KS.Comment.asp?table='+this.value">
   <%
 
    Dim Node,TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
	Dim Url,SelStr,N:N=0
    For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  If KS.S("Table")=Node.SelectSingleNode("tablename").text Then
	   SelStr=" selected"
	   Table=Node.SelectSingleNode("tablename").text
	  ElseIf Node.SelectSingleNode("@isdefault").text="1" and KS.S("Table")="" Then
	   SelStr=" selected"
	    Table=Node.SelectSingleNode("tablename").text
	  Else
	   SelStr=""
	  End If
	  	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"'" &SelStr&">评论表(" & Node.SelectSingleNode("tablename").text&" 共" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &"")(0) &"条)</option>"

	Next
	
	totalPut = Conn.Execute("Select count(id) from [" & Table & "] ")(0)
 %>
   </select>
   
   当前正在管理的评论数据表：<font color=blue><%=Table%></font>,共有 <font color=red><%=totalput%></font> 条
   </div>
    <%
  end if
	  .echo ("<form name='myform' id='myform' method='Post' action='?channelid="& channelid & "'>")
	  .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
	  .echo ("<input type='hidden' name='ProjectID' value='" & ProjectID &"'>")
	  .echo ("<input type='hidden' name='Table' value='" & Table &"'>")
	  .echo ("<input type='hidden' name='VerificType' id='VerificType' value='0'>")
	  .echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
	  .echo "        <tr>"
	  .echo "          <td class=""sort"" width='35' align='center'>选择</td>"
	  .echo "          <td class=""sort"" align=""center"">评论内容</td>"
	  .echo "          <td width=""10%"" class=""sort"" align=""center"">发表人</td>"
	  .echo "          <td width=""10%"" class=""sort"" align=""center"">评论IP</td>"
	  .echo "          <td width=""15%"" align=""center"" class=""sort"">发表时间</td>"
	  .echo "          <td width=""10%"" class=""sort"" align=""center"">状态</td>"
	  .echo "          <td width=""12%"" class=""sort"" align=""center"">管理操作</td>"
	  .echo "        </tr>"

		      Set RSObj = Server.CreateObject("ADODB.RecordSet")
		 
			   Dim Param
			   If KS.G("ComeFrom")="Verify" Then
			   Param=" Where verific=0"
			   Else
			   Param=" Where 1=1"
			   End If
			   If ProjectID<>0 Then  Param=Param & " and projectid=" & KS.ChkClng(Request("ProjectID"))
			   If ChannelID<>0 Then Param=Param & " and ChannelID="& ChannelID&" "

			   If InfoID <> "0" Then
				 Param = Param & " And InfoID IN  (" & InfoID & ")"
			   End If
			   If KS.G("KeyWord")<>"" Then
			    Select Case KS.ChkClng(KS.S("SearchType"))
				 Case 1 Param=Param & " and InfoID In (select InfoID From [KS_ItemInfo] Where Title Like '%" & KS.G("KeyWord") & "%')"
				 Case 2 Param=Param & " and username='" & KS.G("KeyWord") & "'"
				 Case 3 Param=Param & " and Content Like '%" & KS.G("KeyWord") & "%'"
				End Select
			   End If
			   
			 SqlStr ="SELECT * FROM " & Table &  Param & " order by AddDate desc"
			 RSObj.Open SqlStr, conn, 1, 1
			 If RSObj.EOF And RSObj.BOF Then
			    .echo "<tr><td colspan='7' style='text-align:center' class='splittd'>暂时没有评论！</td></tr>"
			 Else
				        totalPut = conn.execute("select count(id) from " & Table & " " & param)(0)
						If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								RSObj.Move (CurrentPage - 1) * MaxPerPage
						End If
						Dim CommentXml:Set CommentXml=KS.ArrayToxml(RSObj.GetRows(MaxPerPage),RSObj,"row","xmlroot")
						Call showContent1(CommentXml)
						Set CommentXml=Nothing

		End If

      RSObj.Close:Set RSOBj=Nothing
	  CloseConn
	  .echo "</table>"
	  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	  .echo ("<tr><td width='180' class='operatingBox'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	  .echo ("</td>")
	  .echo ("<td class='operatingBox'><select onchange='set(this.value)' name='setattribute'><option value=0>快速设置...</option><option value='1'>设为已审</option><option value='2'>设为未审</option><option value='3'>执行删除</option></select></td>")
	  .echo ("</form><td align='right'>")
	  .echo ("</td></tr></table>")
	  
	  	  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		 
		 .echo ("<div style=""clear:both""></div>")
	     .echo ("<form action='KS.Comment.asp?Table=" & Table &"&action=DelAllRecord&ProjectID=" & ProjectID&"' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>特别提醒： </strong><br/>当站点运行一段时间后,网站的评论表可能存放着大量的记录,为使系统的运行性能更佳,您可以一段时间后清理一次。")
		 .echo ("<br /> <strong>删除范围：</strong><input name=""deltype"" type=""radio"" value=1>10天前 <input name=""deltype"" type=""radio"" value=""2"" /> 1个月前 <input name=""deltype"" type=""radio"" value=""3"" />2个月前 <input name=""deltype"" type=""radio"" value=""4"" />3个月前 <input name=""deltype"" type=""radio"" value=""5"" /> 6个月前 <input name=""deltype"" type=""radio"" value=""6""/> 1年前  <input name=""deltype"" type=""radio"" value=""7"" checked=""checked"" /> 2年前 <input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""执行删除""> <input type=""button"" onclick=""if (confirm('此操作不可逆，确定删除所有未审核的评论吗？')){location.href='?Table=" & Table &"&ProjectID=" & ProjectID &"&action=DelUnVerify';}""  class=""button"" value=""一键删除所有未审核的记录"">")
		 
		 .echo ("</div>")
		 .echo "</form>"
		 .echo "</div>"
		 
	  .echo "</body>"
	  .echo "</html>"
	 End With
	End Sub
	Sub ShowContent1(CommentXml)
	  With KS
	  Dim Node,ID
	  If IsObject(CommentXml) Then
		  For Each Node In CommentXml.DocumentElement.SelectNodes("row")
			  ID=Node.SelectSingleNode("@id").text
			    .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & ID & "' onclick=""chk_iddiv('" & ID & "')"">"
			    .echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &ID & "')"" type='checkbox' id='c"& ID & "' value='" &ID & "'></td>"
			    .echo "  <td height='20' class='splittd' title='双击查看详细内容'><span CommentID='" & ID & "' ondblclick=""this.submit()"" title=""" & Node.SelectSingleNode("@content").text & """><img src='../../Images/ico_friend.gif' align='absmiddle'>"
			    .echo "  <span style='cursor:default;'>" & KS.GotTopic(Node.SelectSingleNode("@content").text, 42) & " "
			  If Node.SelectSingleNode("@replycontent").text<>"" Then   .echo "<font color=red>已回复</font>"
			    .echo " </span></span> </td>"
			    .echo "  <td align='center' class='splittd'>" & Node.SelectSingleNode("@username").text & " </td>"
			    .echo "  <td align='center' class='splittd'>" &Node.SelectSingleNode("@userip").text & " </td>"
			    .echo "  <td align='center' class='splittd'><FONT Color=red>" & Node.SelectSingleNode("@adddate").text & "</font> </td>"
			  If Node.SelectSingleNode("@verific").text = 0 Then
			     .echo "  <td align='center' class='splittd'><font color=red><span style='cursor:pointer' onclick='Verific(1," & ID & ")'>未审</span></font></td>"
			  Else
			     .echo "  <td align='center' class='splittd'><span style='cursor:pointer' onclick='Verific(0," & ID & ")'>已审</span></td>"
			  End If
			    .echo "  <td align='center' class='splittd'><a href='KS.Comment.asp?Table=" & Table &"&ProjectID=" & ProjectID &"&Action=View&ChannelID=" & ChannelID & "&CommentID=" & ID & "'>查看/回复</a>  <a href='KS.Comment.asp?Table=" & KS.G("Table") &"&ChannelID=" & ChannelID & "&Action=Del&ID=" & ID & "&ProjectID=" & ProjectID &"' onclick=""return(confirm('确定删除吗?'))"">删除</a></td>"
			    .echo "</tr>"	  
		  Next
	  End If
	 End With
	End Sub
	

     '删除评论
    Sub CommentDel()
	         on error resume next
		 	 Dim K, CommentID,ChannelIDStr,InfoIDStr,ProjectIdStr
			 CommentID = KS.FilterIds(KS.G("ID"))
			 If CommentID="" Then Call KS.AlertHintScript("没有选择记录!")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select UserName,Anonymous,ID,Content,AddDate,channelid,infoid,ProjectId From " & Table&" Where ID In(" & CommentID & ")",conn,1,1
			 Do While Not RS.Eof
			  ChannelIDStr=ChannelIDStr &"," & RS("ChannelID")
			  InfoIDStr=InfoIDStr&","&RS("InfoID")
			  ProjectIdStr=ProjectIdStr&"," &RS("ProjectId")
			  Call ProcessUserScore(RS)
			  RS.MoveNext
			 Loop
			 RS.Close:Set RS=Nothing
			 Conn.Execute("Delete From " & Table&" Where id in(" & CommentID & ")")
			 
			 ChannelIDStr=KS.FilterIds(ChannelIDStr)
			 InfoIDStr=KS.FilterIds(InfoIDStr)
			 ProjectIdStr=KS.FilterIds(ProjectIdStr)
			 If InfoIDStr<>"" Then
			  ChannelIDStr=Split(ChannelIDStr,",")
			  InfoIDStr=Split(InfoIDStr&",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,",",")
			  ProjectIdStr=Split(ProjectIdStr&",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,",",")
			  For K=0 To Ubound(ChannelIDStr)
			   Call updateRatingAvg(ChannelIDStr(k),KS.ChkClng(InfoIDStr(k)),KS.ChkClng(ProjectIdStr(k)))
			   If KS.ChkClng(ChannelIDStr(k))=1000 Then
			   Conn.Execute("Update KS_GroupBuy Set CmtNum=" & LFCls.GetCmtNum(Table,ChannelIDStr(k),InfoIDStr(k)) & " Where ID=" & InfoIDStr(k)) '更新评论数
			   Else
			   Conn.Execute("Update " & KS.C_S(ChannelIDStr(k),2) &" Set CmtNum=" & LFCls.GetCmtNum(Table,ChannelIDStr(k),InfoIDStr(k)) & " Where ID=" & InfoIDStr(k)) '更新评论数
			   End If
			  Next
			 End If
			 
			 KS.AlertHintScript "恭喜，评论删除成功！"
		 End Sub
		 
		 '扣除一个月内被删除的用户积分
		 Sub ProcessUserScore(RS)
		      If Cint(RS(1))=0 And DateDiff("m",RS(4),Now)<1 Then
			     Dim RSU:Set RSU=Server.CreateObject("ADODB.RECORDSET")
				 RSU.Open "Select top 1 groupid From KS_User Where UserName='" & RS(0) & "'",conn,1,1
				 If Not RSU.Eof Then
				    If KS.ChkClng(KS.U_S(RSU(0),7))>0 and not Conn.Execute("Select top 1 id From KS_LogScore Where UserName='" & rs(0) & "' and ChannelID=1002 and InfoID=" & rs("channelid") & "" & rs("InfoID") & " And InOrOutFlag=1").Eof then
					Conn.Execute("Update KS_User Set Score=Score-" & KS.ChkClng(KS.U_S(RSU("GroupID"),7))  & " Where UserName='" & RS(0) & "'")
					
				    Dim CurrScore:CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & RS(0) & "'")(0)
					
			        Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP,Channelid,InfoID) values('" & RS(0) & "',2," & KS.ChkClng(KS.U_S(RSU("GroupID"),7)) & ","&CurrScore & ",'系统','评论[" & KS.GotTopic(KS.HTMLEncode(RS(3)),36) & "]被删除!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "',1002," & RS("ChannelID") & RS("InfoID") & ")")
					
					End If
				 End If
				 RSU.Close
			   End If
		 End Sub
		 
		 '审核评论
		 Sub CommentVerific()
		 
		    If Not KS.ReturnPowerResult(0, "M0100021") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End iF
			
		 	Dim K , CommentID,VerificType
			 VerificType = KS.ChkClng(KS.G("VerificType"))
			 CommentID = KS.FilterIds(KS.G("ID"))
			 If CommentID="" Then Call KS.AlertHintScript("没有选择记录!")
			 If VerificType=1 Then 
			    Dim IDArr:IDArr=Split(CommentID,",")
				For K=0 To Ubound(IDArr)
				  Call VerifyAddScore(IDArr(k))
				Next
			 End If
			 Conn.Execute ("Update " & Table &" set Verific=" & VerificType & " Where ID in(" & CommentID & ")")
			 Dim RS:Set RS=Conn.Execute("select channelid,infoid,ProjectID from " & Table & " Where ID in(" & CommentID&")")
			 Do While Not RS.Eof
			  If rs("ChannelID")=1000 Then
			   Conn.Execute("Update KS_GroupBuy Set CmtNum=" & LFCls.GetCmtNum(Table,RS("ChannelID"),RS("InfoID")) & " Where ID=" & RS("InfoID")) '更新评论数
			  Else
			   Conn.Execute("Update " & KS.C_S(RS("ChannelID"),2) &" Set CmtNum=" & LFCls.GetCmtNum(Table,RS("ChannelID"),RS("InfoID")) & " Where ID=" & RS("InfoID")) '更新评论数
			    If KS.C_S(rs("ChannelID"),7)=1 or KS.C_S(rs("ChannelID"),7)=2 Then
						 Dim KSRObj:Set KSRObj=New Refresh
						 Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
						 RSS.Open "select top 1 * From " & KS.C_S(rs("ChannelID"),2) & " Where ID=" & rs("InfoID"),Conn,1,1
						 Dim DocXML:Set DocXML=KS.RsToXml(RSS,"row","root")
						 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
						  KSRObj.ModelID=rs("ChannelID")
						  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
						  Call KSRObj.RefreshContent()
						  Set KSRobj=Nothing
						RSS.Close
						Set RSS=Nothing
		         End If
			    Call updateRatingAvg(rs("channelid"),rs("infoid"),rs("ProjectID"))
			   
			  End If
			 RS.MoveNext
			 Loop
			 RS.Close
			 Set RS=Nothing
			 Call KS.AlertHintScript("恭喜，更改审核状态操作成功!")
		 End Sub
		 
		 '计算点评的平均分
		 sub updateRatingAvg(channelid,infoid,ProjectID)
		       '=========================计算平均分_begin===========================================
			   if ProjectId=0 Then Exit Sub
			   Dim rs:set rs=server.CreateObject("adodb.recordset")
			   RS.Open "select top 1 * From KS_MoodProject Where ID=" & ProjectID,conn,1,1
			   If Not RS.Eof Then
			        Dim ProjectContentArr:ProjectContentArr=Split(RS("ProjectContent"),"$$$")
				   dim avgscore:avgscore=0
				   dim n,tn:tn=0
				   for n=0 to Ubound(ProjectContentArr)
					if Split(ProjectContentArr(N),"|")(0)<>"" then
						 tn=tn+1
						 dim score:score=conn.execute("select avg(m" & n&") from ks_comment where verific=1 and channelid=" & ChannelID & " and infoid=" & InfoId &" and projectid=" & ProjectID)(0)
						 if not isnumeric(score) then score=0
						 avgscore=avgscore+score
					end if
				   next
				   if tn>0 then avgscore=Round(avgscore/tn,2)
				   Conn.Execute("Update " & KS.C_S(ChannelID,2) & " set avgscore=" & avgscore &" where id=" & infoid)
			   End If
			   RS.Close
			   Set RS=Nothing
				'=========================计算平均分_end===========================================

		 end sub
		 
		sub VerifyAddScore(ID)
		          Dim RS:Set RS=Server.CreateObject("adodb.recordset")
				  rs.open "select top 1 u.userName,u.groupid,c.channelid,c.infoid from " & Table &" c inner join ks_user u on c.userName=u.username where c.anonymous=0 and c.id=" & id,conn,1,1
				  If Not RS.Eof Then
				    If rs("channelid")<>1000 and KS.ChkClng(KS.U_S(rs(1),6))>0 Then
					 Dim RSA:Set RSA=Server.CreateObject("adodb.recordset")
					 RSA.Open "Select top 1 Title,Tid,Fname,adddate From " & KS.C_S(rs("ChannelID"),2) & " Where ID=" & rs("InfoID"),conn,1,1
					 If Not RSA.Eof Then
					 
						 Call  KS.ScoreInOrOut(rs("UserName"),1,KS.ChkClng(KS.U_S(rs("GroupID"),6)),"系统","参与文档[<a href=""" & KS.GetItemUrl(rs("channelid"),rsa(1),rs("infoid"),rsa(2),rsa(3)) & """ target=""_blank"">" & RSa(0) & "</a>]的评论!",1002,""&rs("ChannelID")&""&rs("InfoID"))
					 
					 End If
					 RSA.Close:Set RSA=Nothing
					End If
				  End If
				  rs.close:set rs=nothing
		End Sub
		
		'查看评论 
		Sub CommentView()
    	Dim CommentID:CommentID = KS.G("CommentID")
		Dim RSObj:Set RSObj=Server.CreateObject("ADODB.Recordset")
		RSObj.Open "Select top 1 * From " & Table &" Where ID=" & CommentID, conn, 1, 3
		If RSObj.EOF And RSObj.BOF Then KS.Echo ("参数传递出错!"):Exit Sub
		If KS.G("Flag")="Save" Then
		 RSObj("verific")=KS.ChkClng(Request.Form("verific"))
		 RSObj("Content")=Request.Form("Content")
		 RSObj("ReplyContent")=Request.Form("ReplyContent")
		 RSObj("ReplyTime")=Request.Form("ReplyTime")
		 RSObj("ReplyUser")=Request.Form("ReplyUser")
		 ChannelID=RSObj("ChannelID")
		 Dim InfoID:InfoID=RSOBj("InfoID")
		 RSObj.Update
		 Call updateRatingAvg(rsobj("channelid"),rsobj("infoid"),rsobj("ProjectID"))
		 If KS.ChkClng(Request.Form("verific"))=1 Then
		  Call VerifyAddScore(CommentID)
		    '自动生成内容页
			 If KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2 Then
				 Dim KSRObj:Set KSRObj=New Refresh
				Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,Conn,1,1
						 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
						 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
						  KSRObj.ModelID=ChannelID
						  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
						  Call KSRObj.RefreshContent()
						  Set KSRobj=Nothing
				RS.Close
				Set RS=Nothing
			 End If
		 End If
		 If ChannelID=1000 Then
			   Conn.Execute("Update KS_GroupBuy Set CmtNum=" & LFCls.GetCmtNum(Table,ChannelID,InfoID) & " Where ID=" & InfoID) '更新评论数
		 Else
			   Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set CmtNum=" & LFCls.GetCmtNum(Table,ChannelID,InfoID) & " Where ID=" & InfoID) '更新评论数
		End If
			   
		 KS.Echo "<script>alert('恭喜,评论修改成功!');location.href='" & Request.Form("ComeUrl") & "';</script>"
		End If
        With KS
			Dim ARS, Url,SqlStr,ChannelID,ReplyTime,ReplyUser
			ChannelID=KS.ChkClng(RSObj("ChannelID"))
			if channelid=1000 then
			 sqlstr="select top 1 ID,subject as Title,classid as Tid,0,0,0,0,adddate from KS_GroupBuy Where ID=" & RSObj("InfoID")
			Else
			Select Case KS.C_S(ChannelID,6)
			 Case 1 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 2 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 3 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 4 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 5 SQLStr="select top 1 ID,Title,Tid,0,0,Fname,0,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 7 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 8 SqlStr="select top 1 ID,Title,Tid,0,0,Fname,0,adddate from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			End Select
		 End If
			
			ReplyTime=RSObj("ReplyTime")
			If ReplyTime="" Or IsNull(ReplyTime) Then
			 ReplyTime=Now
			End If
			ReplyUser=RSObj("ReplyUser")
			If ReplyUser=""  Or IsNull(ReplyUser) Then
			ReplyUser=KS.C("AdminName")
			End If
			dim itemname
			if channelid=1000 then
			itemname="团购"
			else
			itemname=KS.C_S(ChannelID,3)
			end if
			
				  .echo "<!DOCTYPE html><html>"
				  .echo "<head>"
				  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				  .echo "<link href=""../../include/Admin_Style.css"" rel=""stylesheet"">"
				  .echo "<script language=""JavaScript"" src=""../../../KS_Inc/common.js""></script>"
				  .echo "<title>查看评论</title>"
				  .echo "</head>"
				  .echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				  .echo "<div class='topdashed sort'>评论查看/回复</div>"
				  .echo "  <br>"
				  .echo "   <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"" Class=""Ctable"">"
				  .echo "    <form name='myform' action='KS.Comment.asp' method='post'>"
				  .echo "    <input type='hidden' value='" & Request.ServerVariables("HTTP_REFERER") & "' name='ComeUrl'>"
				  .echo "    <input type='hidden' value='" & ChannelID & "' name='ChannelID'>"
				  .echo "    <input type='hidden' value='" & CommentID & "' name='CommentID'>"
				  .echo "    <input type='hidden' value='" & ProjectID & "' name='ProjectID'>"
				  .echo "    <input type='hidden' value='" & Table & "' name='Table'>"
				  .echo "    <input type='hidden' value='View' name='Action'>"
				  .echo "    <input type='hidden' value='Save' name='Flag'>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td width=""200"" class='clefttitle' height=""25"">" &itemname &"标题</td>"
				  .echo "            <td> "
				  
				   Set Ars= Conn.Execute(SqlStr)
				   If Not ARS.EOF Then
					 Url = KS.GetItemUrl(ChannelID,aRS(2),ars(0),ars(5),ars(7))
					 If ChannelID=1000 Then
					    Url="../shop/groupbuyshow.asp?id=" & ars(0)
					 Else
						 If ChannelID=1 Then
						  If ARS("Changes")=1 Then Url=ARS("Fname")
						 End IF
					 End If
					   .echo "<a href=""" & Url & """ target=""_blank"">" & ARS("title") & "</a>"
				   End If
				   ARS.Close:Set ARS = Nothing
				  .echo "          </td></tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td class='clefttitle' height=""25"">发表人</td>"
				  .echo "            <td> " & RSObj("UserName") & " 发表于 " & RSObj("AddDate") & " 用户IP:" & RSObj("UserIP") &"</td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">票数</td>"
				  .echo "            <td>支持:" & RSObj("score") & "票  反对" & RSObj("oscore") & "票</td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">评论状态</td><td>"
				  .echo " <input type='radio' value='1' name='verific'"
				 If RSObj("verific")=1 Then   .echo " checked"
				  .echo ">已审核"
				  .echo " <input type='radio' value='0' name='verific'"
				 If RSObj("verific")=0 Then   .echo " checked"
				  .echo ">未审核"
				  .echo "          </td></tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">评论内容"
				If RSObj("QuoteContent")<>"" And Not IsNull(RSObj("QuoteContent")) Then
				   .echo "<div style='color:red;font-weight:bold'><br />该评论内容有引用</div>"
				End If
				  .echo "</td>"
				  .echo "            <td><textarea name='Content' style=""overflow:auto;height:100px; width:380px;"">" & ReplaceFace(RSObj("Content")) & "</textarea></td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">回复内容</td>"
				  .echo "            <td><textarea name='ReplyContent' style=""overflow:auto;height:60px; width:380px;"">" & RSObj("ReplyContent") & "</textarea></td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">回复时间</td>"
				  .echo "            <td><input type='text' name='ReplyTime' class='textbox' value='" & ReplyTime & "'></td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">回复人</td>"
				  .echo "            <td><input type='text' name='ReplyUser' class='textbox' value='" & ReplyUser & "'></td>"
				  .echo "          </tr>"
				
				  .echo "        </table>"

				  .echo "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				  .echo "    <tr>"
				  .echo "      <td height=""40"" align=""center"">"
				  .echo "        <input type='submit' class='button' value='确定修改'>"
				  .echo "        <input type=""button"" name=""Submit1"" onclick=""javascript:window.open('" & Url & "','new','');"" value=""查看" & itemname &""" class='button'>"
				  .echo "      </td>"
				  .echo "    </tr>"
				  .echo "</form>"
				  .echo "  </table>"
				  .echo "  <br>"
				  .echo "</body>"
				  .echo "</html>"
			End With
		End Sub
		
		Function ReplaceFace(c)
		 Dim str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K
		 For K=0 To 19
		  c=replace(c,"[e"&K &"]","<img title=""" & strarr(k) & """ src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">")
		 Next
		 ReplaceFace=C
		End Function

End Class
%> 

<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KS:Set KS=New PublicCls
If Not KS.ReturnPowerResult(0, "KSMB10002") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
End iF
Dim TableXML,Node,N,TaskUrl,Taskid,Action
'Set TableXML=LFCls.GetXMLFromFile("task")
set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
TableXML.async = false
TableXML.setProperty "ServerHTTPRequest", true 
TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))

Action=Request.QueryString("Action")
Select Case Action
  case "DoSave" DoSave
  case "ModifySave" ModifySave
  case "del" del
  case else
    Manage
End Select


Sub manage()
%>
<!DOCTYPE html><html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<title>论坛数据表管理</title>
<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script src="../../KS_Inc/Jquery.js"></script>
</head>
<body>
<ul id='mt'> <div id='mtl'>论坛数据表管理</div></ul>
<div class="pageCont2"><form name="myform" action="KS.GuestTable.asp?action=ModifySave" method="post">
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
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>您没有添加小论坛数据表!</td></tr>"
  Else
	  N=0
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
				 <a href="?action=del&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('确定删除该任务吗?'))">删除</a>
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
		 if (tobj.val().toLowerCase().indexOf('ks_guest_')==-1){
		  alert('数据表名必须与KS_Guest_开头!');
		  tobj.focus();
		  return false;
		 }
		 return true;
		}
	   </script>
	   <form name="myform" action="KS.GuestTable.asp?action=DoSave" method="post" id="myform">
	  <table width='99%' style="margin:4px" align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>数据表名称:</strong></td>
		   <td><input type="text" class="textbox" name="TableName" id="TableName" value="KS_Guest_"> 
		   如:KS_Guest_BBS1,KS_Guest_BBS2等,必须以KS_Guest_开头</td>
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
		  
		  
	   <div class="attention">
<strong>特别提醒：</strong><br/>
1、数据表中选中的为当前论坛所使用来保存回复帖子数据的表，一般情况下每个表中的数据越少论坛帖子显示速度越快，当您上列单个帖子数据表中的数据有超过几万的帖子时不妨新添一个数据表来保存帖子数据,您会发现论坛速度快很多很多。<br/>
2、您也可以将当前所使用的数据表在上列数据表中切换，当前所使用的帖子数据表即当前论坛用户发贴时默认的保存帖子数据表。<br/>
3、以免出错，当前正在使用的数据表、已有记录的数据表或是系统自带数据表不允许删除。
</div>
</div>
</body>
</html>
<%
End Sub


'保存
Sub DoSave()
 Dim TableName:TableName=KS.G("TableName")
 Dim Descript:Descript=Request.Form("Descript")
 Dim isdefault:isdefault=KS.ChkClng(Request.Form("isdefault"))
 If Len(TableName)<10 or lcase(left(TableName,9))<>"ks_guest_" Then
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
						"UserName nvarchar(50),"&_
						"UserID int default 0,"&_
						"UserIP varchar(100),"&_
						"TopicID int Default 0,"&_
						"ParentID int default 0,"&_
						"TxtHead varchar(100),"&_
						"Content ntext,"&_
						"ReplayTime datetime,"&_
						"Verific tinyint default 0,"&_
						"showip tinyint default 0,"&_
						"showsign tinyint Default 0,"&_
						"Opposition int Default 0,"&_
						"deltf tinyint Default 0,"&_
						"Support int Default 0"&_
						")"
 Conn.Execute sql
 On Error Resume Next
 Conn.Execute("CREATE INDEX [TopicId] ON " & TableName & "([TopicId])")
 Conn.Execute("CREATE INDEX [IX_" & TableName & "] ON [" & TableName & "] ( TopicID,Deltf) ")

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
	 TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	 If isdefault=1 Then
	  For Each Node In TableXML.DocumentElement.SelectNodes("item")
		 If ItemID=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
			Node.Attributes.getNamedItem("isdefault").text=1
		 Else
			Node.Attributes.getNamedItem("isdefault").text=0
		 End If
	  Next
	  TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	 End If
	 
	 Response.Write "<script> alert('恭喜,论坛数据表添加成功!');location.href='?'</script>"
End Sub

'保存修改
Sub ModifySave()
  Dim id,isdefault:isdefault=KS.ChkClng(Request.Form("isdefault"))

  For Each Node In TableXML.DocumentElement.SelectNodes("item")
     ID=KS.ChkClng(Node.SelectSingleNode("@id").text)
	 If ID=isdefault Then
	    Node.Attributes.getNamedItem("isdefault").text=1
     Else
	    Node.Attributes.getNamedItem("isdefault").text=0
	 End If
  Next

	 TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	 Response.Write "<script>alert('恭喜,设置成功!');location.href='?'</script>"
End Sub

Sub Del()
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
  TableXML.Save(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
  KS.AlertHintScript "恭喜,论坛数据表已删除!"
End Sub



Set KS=Nothing
CloseConn
%>
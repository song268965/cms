<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Template.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Link
KSCls.Kesion()
Set KSCls = Nothing

Class Link
        Private KS,KSUser,ModelTable,Param,XML,Node,StartTime,FormID,TableName,id
		Private AdminUserList,LoginTF
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		Public Sub Kesion()
		    LoginTF=Cbool(KSUser.UserLoginChecked)
			IF LoginTF=false Then
			   KS.Die "<script>top.location.href='../../user/Login';parent.box.close();</script>"
			End If
		   dim rs,Templ_url
		   FormID=KS.ChkClng(KS.G("FormID"))
		   ID=KS.ChkClng(KS.G("ID"))
		   if  FormID=0 then Call KS.AlertHistory("ID错误!",-1):response.end
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   RS.Open "Select TOP 1 TableName,AdminUserList From KS_Form Where ID=" & FormID,conn,1,1
		   If RS.EOF And RS.Bof Then
			 Call KS.AlertHistory("没有数据!",-1):response.end
		   else
			 TableName=RS("tablename")
			 AdminUserList=RS("AdminUserList")
		   End If
		   RS.Close
		   if checkadminpower=false Then
		     KS.Die "<script>alert('对不起，您没有权限!');parent.box.close();</script>"
		   End If
		   if request("action")="replaysave" then
		    ReplaySave
		   else
		    ShowReply
		   end if
		   
	   End Sub
	   
	   Sub ReplaySave()
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim ID:ID=KS.ChkClng(KS.G("id"))
		 Dim TableName:TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & FormID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 note,replydate,username  From " & TableName &" Where ID=" & ID,conn,1,3
          RS(0)=Request.Form("Content")
		  if Request.Form("replydate")="" then
		  RS(1)= formatdatetime(now(),2)
		 else
		 RS(1)=Request.Form("replydate")
		 end if
		 RS(2)=ksuser.username
		 RS.Update
		 RS.Close
		 Set RS=Nothing
		 ks.die "<script>alert('恭喜，留言回复成功!');top.location.reload();</script>"
	   End sub
	   
	   Sub ShowReply()
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From " & TableName &" Where ID=" & ID,conn,1,1
		 If RS.Eof Then
		  response.end
		 End If
         %>
		 <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../../user/images/css.css" type=text/css rel=stylesheet>
<script src="../../KS_Inc/jquery.js"></script>
<script src="../../KS_Inc/DatePicker/WdatePicker.js"></script>
<%=EchoUeditorHead()%>
</head>
<body>
		 <iframe src="about:blank" style="display:none" name="hiddenframe"></iframe>
		 <form action="reply.asp?action=replaysave&formid=<%=formid%>&id=<%=id%>" method="post" name="myform" target="hiddenframe">
		  <div style="margin:6px;text-align:center;font-weight:bold;color:red">查看并回复</div>
		  		  <table width='99%' align='center' border='0' cellpadding='1'  cellspacing='1' class='ctable'> 

		  <%
		   RS.Close

			   Dim S_Content,sql,k,ReturnInfo,UpFiles
			   set rs=conn.execute("select FieldName,title,MustFillTF,FieldType,ShowUnit from ks_formfield where itemid=" & Formid & " and ShowOnForm=1 order by orderid")
			   sql=rs.getrows(-1)
			   rs.close
			   rs.open "select top 1 * From " & TableName & " Where ID=" & ID,conn,1,1
			   for k=0 to ubound(sql,2)
				
				s_content=s_content &"<tr class=""tdbg"">" & vbcrlf
				s_content=s_content & "<td width=120 align=right class='clefttitle'>" & sql(1,k) & "：</td>" & vbcrlf
				s_content=s_content & "<td>" 
				
				s_content=s_content & rs(trim(sql(0,k)))
				
				s_content=s_content & "</td>" & vbcrlf
				s_content=s_content & "</tr>" & vbcrlf
			   next
			    response.write s_content
			 
		  %>
		  
		  <tr class="tdbg">
		    <td align="right" class="clefttitle">发表时间：</td>
			<td><%=rs("adddate")%></td>
		  </tr>
		  <tr class="tdbg">
		    <td align="right" class="clefttitle">回复时间：</td>
			<td><input type="text"  style="padding:2px;width:200px;" id="replydate" name="replydate" onclick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'});" value="<%=rs("replydate")%>" ></span>
			</td>
		  </tr>

		    
		  <tr class="tdbg">
		   <td align="right" class="clefttitle">回复内容：</td>
		   <td>
		    <%
				 Response.Write "<script id=""content"" name=""content"" type=""text/plain"" style=""width:730px;height:250px;"">" &KS.ClearBadChr(rs("note"))&"</script>"
	             Response.Write "<script>setTimeout(""var editor = " & GetEditorTag() &".getEditor('content',{toolbars:[" & GetEditorToolBar("newstool") &"],wordCount:false,autoHeightEnabled:false,minFrameHeight:150 });"",10);</script>"
				%>
		   </td>
		  </tr>
		  
		  <tr  class="tdbg">
		    <td colspan="2" height="35" style="text-align:center"><input type="submit" class="button" value="提交回复">&nbsp;<input type="button" class="button" value="关闭窗口" onClick="parent.box.close();"></td>
		  </tr>
		  
		  </table>
		 </form>
		 <%
		 
			
		 
		 
		  RS.Close:Set RS=Nothing
		End Sub
	   
	   
	   	'检查有没有管理表单权限
		function checkadminpower()
		     if not ks.isnul(adminuserlist) then
			    if ks.foundinarr(adminuserlist,ks.c("username"),",") and LoginTF=True then
			      checkadminpower=true
				  exit function
			    end if
			   end if
			checkadminpower=false
		end function
	   
	  
End Class
%>

 

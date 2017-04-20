<%@ Language="VBSCRIPT" codepage="65001" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
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
Response.CharSet="utf-8"
Dim KSCls
Set KSCls = New AjaxCls
KSCls.Kesion()
Set KSCls = Nothing

Class AjaxCls
      Private KS,KSUser
	  Private Action,Template,id,groupadmin
	  Private totalPut,MaxPerPage,PageNum
	  Private Sub Class_Initialize()
	   Set KS=New PublicCls
       Set KSUser=New UserCls
      End Sub
	 Private Sub Class_Terminate()
	  Set KS=Nothing
	  Set KSUser=Nothing
	  CloseConn()
	 End Sub

     Sub Kesion()
      Action=KS.S("Action")
	   Select Case Action
		Case "photocmt","logcmt"
		 Call photocmt()
	   End Select	
	 End Sub	
	 

  
  '照片及日志评论列表
  Sub photocmt()
    maxperpage=5
	dim substr,table,param
	if action="logcmt" then
	  table="KS_BlogComment"
	  param=" Where LogID=" & KS.ChkClng(request("ID"))
	else
	  table="KS_PhotoComment"
	  param=" Where photoID=" & KS.ChkClng(request("ID"))
	end if
	Dim sqlstr:SqlStr="Select * From " & table & param  & " Order By AddDate Desc,id"
    Dim rs:Set rs=Server.CreateObject("ADODB.RECORDSET")
	rs.Open SqlStr,Conn,1,1
	IF rs.Eof and rs.bof Then
	 substr="没有评论！"
	Else
			totalPut = conn.execute("select count(1) from " & table & param)(0)
			If (totalPut Mod MaxPerPage) = 0 Then
				pagenum = totalPut \ MaxPerPage
			Else
				pagenum = totalPut \ MaxPerPage + 1
			End If
			If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
				rs.Move (CurrentPage - 1) * MaxPerPage
			End If
     substr=substr & "<div style=""border-bottom:1px solid #f1f1f1;padding-bottom:2px;font-weight:bold;font-size:14px;text-align:left"">&nbsp;&nbsp;有 <font color=red>" & totalPut & " </font> 条评论，共分 <font color=red>" & pagenum & "</font> 页,第 <font color=red>" & CurrentPage & "</font> 页</div>"
    substr=substr & "<table  width='99%' border='0' align='center' cellpadding='0' cellspacing='1'>"
    Dim FaceStr,Publish,i,n
     If CurrentPage=1 Then
	  N=TotalPut
	 Else
	  N=totalPut-MaxPerPage*(CurrentPage-1)
	 End IF
  Do While Not RS.Eof 
   FaceStr=KS.Setting(3) & "images/face/boy.jpg"

    Publish=RS("AnounName")
	If not Conn.Execute("Select top 1 UserFace From KS_User Where UserName='"& Publish & "'").eof Then
      FaceStr=Conn.Execute("Select top 1 UserFace From KS_User Where UserName='"& Publish & "'")(0)
	  If lcase(left(FaceStr,4))<>"http" and left(facestr,1)<>"/" then FaceStr=KS.Setting(3) & FaceStr
   End IF
	
   substr=substr & "<tr>"
   substr=substr & "<td width='70' rowspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;' valign='top'><img width=""50"" height=""52"" src=""" & facestr & """ border=""1"" class=""faceborder"" style=""margin-top:2px;margin-bottom:5px""></td>"

   substr=substr & "<td height='25'>" & ReplaceFace(RS("Content"))
   		 If Not KS.IsNul(RS("Replay")) Then
		 substr=substr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>以下为space主人的回复:</b><br>" & RS("Replay") & "<br><div align=right>时间:" & rs("replaydate") &"</div></div>"
         End If
   substr=substr & "	 </td>"
   substr=substr & "</tr>"
   substr=substr & "<tr>"
   
   			 Dim MoreStr,KSUser,LoginTF
			 IF trim(KS.C("UserName"))=trim(RS("UserName")) Then
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> | <a href='../User/user_message.asp?Action=CommentDel&id=" & RS("ID") & "&flag=" & action &"' onclick=""return(confirm('确定删除该留言吗?'));"">删除</a> | <a href='../user/user_message.asp?id=" & RS("ID") & "&Action=reply"&action &"' target='_blank'>回复</a>"
			 Else
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> "
			 End If

   substr=substr & "<td align='right' colspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><font color='#999999'>(" & publish & " 发表于：" & RS("AddDate") &")</font>&nbsp;&nbsp;" & MoreStr & " </td>"
   substr=substr & "</tr>"
   N=N-1
   RS.MoveNext
		I = I + 1
	  If I >= MaxPerPage Then Exit Do
  loop
 substr=substr & "</table>"

	End If
	substr=substr &"<div class=""clear""></div>"
	Response.write substr
	Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|条||1"
	RS.Close:Set RS=Nothing
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
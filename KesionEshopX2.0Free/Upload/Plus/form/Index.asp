<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
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
        Private KS,KSUser,ChannelID,ModelTable,Param,XML,Node,StartTime,FormID,TableName,FormName
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,stype,OrderStr,AdminUserList,LoginTF
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		Public Sub Kesion()

		 Dim Template,KSR
		  Call FCls.SetClassInfo(1,"20124415294268","0")
		
		 LoginTF=KSUser.UserLoginChecked
		 
		 Set KSR = New Refresh
		   dim rs,Templ_url
		   FormID=KS.ChkClng(KS.G("id"))
		   if  FormID=0 then Call KS.AlertHistory("ID错误!",-1):response.end
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   RS.Open "Select top 1 FormName,PostByStep,TableName,Template,Templ_url,AdminUserList,MaxPerPage_s From KS_Form where id=" & FormID,conn,1,1
		   If RS.EOF And RS.Bof Then
			 Call KS.AlertHistory("ID错误!",-1):response.end
		   else
			 Templ_url=RS(4):TableName=RS(2):AdminUserList=RS(5)
			 MaxPerPage=KS.ChkCLng(RS("MaxPerPage_s"))
			 If MaxPerPage=0 Then MaxPerPage=10
			 FormName=RS(0)
		   End If
		   RS.Close

		select case request("action")
		   case "verify" call verify()
		   case "delete" call formdelete()
		 end select

		   
		   
		   Template = KSR.LoadTemplate(Templ_url)
		   Template =Replace(Template,"{$" & KS.S("Type") &"}"," selected")
		   Template =Replace(Template,"{$ShowFormName}",formname)
		   Template =Replace(Template,"{$ShowFormID}",formid)
		   Template = KSR.KSLabelReplaceAll(Template)
		   Set KSR = Nothing
		   StartTime = Timer()
		   InitialSearch
		   Scan Template
	   End Sub
	   Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "loop"
				      If IsObject(XML) Then
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
					   Next
					  Else
					   echo "<div class='border' style='clear:both;text-align:center'>对不起,根据您的查找条件,找不到任何相关记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName
				case "search" 
				          select case sTokenName
						    case "showpage" echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							case "leavetime" 
							   dim leavetime:leavetime=FormatNumber((timer-starttime),5)
							   if leavetime<1 then leavetime="0"&leavetime
							   echo leavetime
							case "keyword" echo KS.R(key)
							case "channelid" echo channelid
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "adddate" echo formatdatetime(GetNodeText("AddDate"),2)
			case "userip" 
			   dim userip:userip=GetNodeText("userip")
			   if not ks.isnul(userip) then
			     dim useriparr:useriparr=split(userip&"...." ,".")
				 echo useriparr(0)&"."&useriparr(1)&".***.***"
			   end if
			case "linkurl"
			   echo KS.GetDomain & "Plus/form/content.asp?FormID="&FormID& "&id=" & GetNodeText("id")
			case "intro" 
			case "cnote"
				 if not ks.isnul(GetNodeText("note")) then
					echo "<span style='color:green'>有回复</span>"		
				 else
					echo "<span style='color:red'>未回复</span>"	
				 end if
		    case "showmanage"
			  if checkadminpower() then
			     if GetNodeText("status")="0" then
				  echo "<a href='?action=verify&v=2&id=" & formid & "&rid=" & GetNodeText("id") & "' title='点此审核' style='color:red'>未审核</a>"
				 else
				  echo "<a href='?action=verify&v=0&id=" & formid &"&rid=" & GetNodeText("id") & "' title='点此取消审核' style='color:green'>已审核</a>"
				 end if
			      echo " | <a href='?action=delete&id=" & formid & "&rid=" & GetNodeText("id") & "' onclick=""return(confirm('此操作不可恢复，确定删除吗？'))"">删除</a> | <a href='javascript:showReply("&FormID& "," & GetNodeText("id") &");'>回复</a>"
			   end if
			case else
			  echo GetNodeText(sTokenName)
		  End Select
		End Sub
		Function GetNodeText(NodeName)
		 Dim N,Str
		
		 NodeName=Lcase(NodeName)
		 
		 If IsObject(Node) Then
		  	set N=node.SelectSingleNode("@" & NodeName) 
		  	If Not N is Nothing Then Str=N.text
			
		  	GetNodeText=Str
		  End If
		End Function
		
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
		
		sub verify()
		 if checkadminpower()=false then
		   ks.die "<script>alert('对不起，你没有权限!');history.back();</script>"
		 end if
		 if ks.chkclng(request("v"))=2 then
		 conn.execute("update " & TableName & " set status=2 where id=" & KS.ChkClng(request("rid")))
		 ks.alerthintscript "恭喜，审核成功!"
		 else
		 conn.execute("update " & TableName & " set status=0 where id=" & KS.ChkClng(request("rid")))
		 ks.alerthintscript "恭喜，取消审核成功!"
		 end if
		end sub
		
		sub formdelete()
		 if checkadminpower()=false then
		   ks.die "<script>alert('对不起，你没有权限!');history.back();</script>"
		 end if
		 conn.execute("delete from " & TableName & " where id=" & KS.ChkClng(request("rid")))
		 ks.alerthintscript "恭喜，删除成功!"
		end sub
		
		
		
		Sub InitialSearch()
		  Dim FieldStr,SqlStr,TopStr,TopNum
		  ChannelID=KS.ChkClng(Request("M"))
		  CurrPage=KS.ChkClng(Request("Page"))
		  If CurrPage<=0 Then CurrPage=1

	
		 	 ModelTable=TableName
			 if checkadminpower then
		        Param = " where 1=1"
		     else
			    Param = " where status=2"
			 end if
		   dim ks_type:ks_type=KS.ChkClng(KS.S("type"))
		   if ks_type<>0 then 
		    select case ks_type
			 case 1
		      param=param &"  and ks_type='在线留言'"		   
			 case 2
		      param=param &"  and ks_type='校长信箱'"		   
			 case 3
		      param=param &"  and ks_type='投诉建议'"		   
			end select
		   end if

          OrderStr=" order by id desc"
		  SqlStr="Select " & TopStr & " * From " & ModelTable & Param & OrderStr
		  'ks.echo sqlstr
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     TotalPut = Conn.Execute("select Count(1) from " & ModelTable & " " & Param)(0)
			 If TotalPut>TopNum And TopNum<>0 Then TotalPut=TopNum
			 If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrPage - 1) * MaxPerPage
			 Else
					CurrPage = 1
			 End If
			 Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
		  End If
		 RS.Close
		 Set RS=Nothing
		End Sub
	
		
End Class
%>

 

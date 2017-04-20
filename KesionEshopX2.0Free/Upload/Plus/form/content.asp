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
        Private KS,ModelTable,Param,XML,Node,StartTime,FormID,TableName,id,adminuserlist,Cipher,FormName
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		 Dim Template,KSR
		 
		  Call FCls.SetClassInfo(1,"20124415294268","0")

		 Set KSR = New Refresh
		   dim rs,Templ_url
		   FormID=KS.ChkClng(KS.G("FormID"))
		   ID=KS.ChkClng(KS.G("ID"))
		   if  FormID=0 then Call KS.AlertHistory("ID错误!",-1):response.end
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   RS.Open "Select top 1 FormName,PostByStep,TableName,Template,Tempc_url,adminuserlist,Cipher From KS_Form Where ID=" & FormID,conn,1,1
		   If RS.EOF And RS.Bof Then
			 Call KS.AlertHistory("没有数据!",-1):response.end
		   else
			 Templ_url=RS(4):TableName=RS(2):adminuserlist=rs("adminuserlist")
			 Cipher=RS("Cipher")
		   End If
		   RS.Close
		   Template = KSR.LoadTemplate(Templ_url)
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
				Case "cont"
					  If IsObject(XML) Then
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
					    if instr(GetNodeText("PassWord"),"{o|yes|o")<>0 then
						'if Cipher=1 then
						  if checkadminpower()=false then
							%>
							<script>
							function content_k(FormID,id){
							
								if (pass=$('input[name=PassWord_k]').val()=="")
								{
									alert("请输入密码!")
									return false;
								}
								return true;
							} 
							</script>
							<%
							echo "<form name=""myform"" action=""content.asp"" method=""post"" onsubmit=""return(content_k('"& FormID &"','"& GetNodeText("id") &"'))"">"
							echo "<input type=""hidden"" value="""& GetNodeText("id") &""" name=""id"">"
							echo "<input type=""hidden"" value="""& FormID &""" name=""FormID"">"
							dim PassWord_ok
							PassWord_ok=Replace(GetNodeText("PassWord"),"{o|yes|o}","")
							
							if KS.G("PassWord_k")="" then
								echo "<div  id=""PassWord_"" >输入密码可查看:<input name=""PassWord_k"" class=""PassWord_k"" type=""password""  value=""""/><input name="""" type=""submit"" value=""确定"" class=""button_p"" /></div>"
							else
								if KS.G("PassWord_k")=PassWord_ok then
									Scan sTemplate
								else
									echo "<div  id=""PassWord_"" ><font color=red>密码不正确</font>,请重新输入:<input name=""PassWord_k"" class=""PassWord_k"" type=""password""  value=""""/><input name="""" type=""submit"" value=""确定"" class=""button_p""   /></div>"	
								end if
							end if
							echo "</form>"
						  else
						    echo "<div style='padding:10px;color:brown;border:1px solid #f1f1f1;'>TIPS:以下内容需要登录密码才可以查看,由于您是管理员,所以可以查看!</div>"
						    Scan sTemplate
						  end if
					    else
							Scan sTemplate
						end if
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,根据您的查找条件,找不到任何相关记录!</div>"
					  End If
			End Select 
			
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName     
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "adddate" echo GetNodeText("AddDate")
			case "userip" 
			   dim userip:userip=GetNodeText("userip")
			   if not ks.isnul(userip) then
			     dim useriparr:useriparr=split(userip&"...." ,".")
				 echo useriparr(0)&"."&useriparr(1)&".***.***"
			   end if
			case "linkurl"
			   echo KS.GetDomain & "Plus/from/content.asp?FormID="&FormID& "&id=" & GetNodeText("id")
			case "classname" 
			  If ChannelID=102 Then
			   echo GetNodeText("pclassname") & GetNodeText("classname")
			  Else
			   echo KS.C_C(GetNodeText("tid"),1)
			  End If
			case "classurl" 
			 If ChannelID=102 Then
			  echo KS.GetDomain & "ask/showlist.asp?id=" & Node.SelectSingleNode("@classid").text
			 Else
			  echo KS.GetFolderPath(GetNodeText("tid"))
			 End If
			case "intro" 
			 Dim Intro:intro=KS.Gottopic(KS.LoseHtml(GetNodeText("KS_Content")),160)
			 Intro=Replace(Intro,"&nbsp;","")
			 echo intro

			case "cnote"
			 if GetNodeText("PassWord")<>"{o|no|o}" and GetNodeText("PassWord")<>"" then
				 if not ks.isnul(GetNodeText("note")) then
						echo GetNodeText("note")
				 else
					echo "未回复"	
				 end if
			 else
			 	echo GetNodeText("note") 
			 end if
			case else
			  echo GetNodeText(sTokenName)
		  End Select
		End Sub
		Function GetNodeText(NodeName)
		Dim N,Str
		
		 NodeName=Lcase(NodeName)
		 
		 If IsObject(Node) Then
		 	
		  if Instr(NodeName,"zd*")<>0  then
		  	dim form_str:form_str=Split(NodeName,"*")
			if UBound(form_str)=1 then
				if not ks.isnul(form_str(1)) then 
					Str=GetNodeText(form_str(1))
					GetNodeText=Str
				end if
			end if
		  else
		  	set N=node.SelectSingleNode("@" & NodeName)
		  	If Not N is Nothing Then Str=N.text		 
		  end if
		  
		  'if NodeName="title" Then Str= GetNodeText("KS_title")
		  	GetNodeText=Str
		  End If
		End Function
		
		
		Sub InitialSearch()
		  Dim SqlStr
		  ModelTable=TableName
		  Param = " where id=" & id &" "
		  SqlStr="Select top 1 *  From " & ModelTable & Param
		  
		  'Select Cipher From KS_Form
		  'ks.echo sqlstr
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     Set XML=KS.RsToxml(RS,"row","root")
		  End If
		 RS.Close
		 Set RS=Nothing
		End Sub
		
		
		'检查有没有管理表单权限
		function checkadminpower()
		     dim ksuser:set ksuser=new usercls
		     dim logintf: LoginTF=KSUser.UserLoginChecked
			 set ksuser=nothing
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

 

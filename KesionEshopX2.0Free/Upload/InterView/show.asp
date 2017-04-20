<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 9.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New CartCls
KSCls.Kesion()
Set KSCls = Nothing

Class CartCls
        Private KS, KSR,KSUser,DomainStr,Template,TotalPut,MaxPerPage,CurrPage
		Private XML,Node,LoginTF,id,Templates
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		  Set KSUser = New UserCls
		  Set KSR = New Refresh
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Sub Echo(sStr)
			Templates    = Templates & sStr 
		End Sub
		
		Public Sub Kesion()
		  ID=KS.ChkClng(KS.S("ID"))
		  If ID=0 Then KS.Die "error"
		  Fcls.RefreshFolderID = "0"        '设置当前刷新目录ID 为"0" 以取得通用标签
          CurrPage=KS.ChkClng(KS.S("Page"))
		  If CurrPage<0 Then CurrPage=1
		   Initial()
		   Scan Template
          Templates = KSR.KSLabelReplaceAll(Templates)
		   response.write Templates
	   End Sub
	   
	   Sub Initial()
	      Dim SqlStr:SqlStr="Select top 1 * From KS_InterView  Where ID=" & ID
		  'ks.echo sqlstr
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		    KS.Die "<script>alert('找不到访谈记录！');location.href='../';</script>"
		  Else
			 Set XML=KS.RsToxml(RS,"row","xmlroot")
			 Set Node=Xml.DocumentElement.SelectSingleNode("row")
			 Template = KSR.LoadTemplate(Node.SelectSingleNode("@templateid").text)
		  End If
		 RS.Close
		 Set RS=Nothing
	   
	   
	   End Sub
	  
	   
	   
	   public Sub Scan(ByVal sTemplate)
			Dim iPosLast, iPosCur
			iPosLast = 1
			Dim Tags,Key,yllen
			
          do  while (true)
                iPosCur = findTags(sTemplate, tags, key, iPosLast,yllen)
                if (iPosCur <>0) then
                    Echo mid(sTemplate,iPosLast, iPosCur - iPosLast)
                    select case (tags)
                        case "{#"
                            Parse sTemplate, key
                    end select
                    iPosLast = yllen + 1
                else
                    Echo    Mid(sTemplate, iPosLast)
                    exit do
				end if
          loop
		End Sub 
		
		Function FindTags(sTemplate, byref tags, ByRef key, iPosLast, ByRef yllen)
				dim a:a = array("{#")   '定义标签开始标记
				dim i, cur, posCur
				cur=0
				for i=0 to ubound(a)
						posCur=instr(iPosLast,sTemplate,a(i))
						if (posCur<>0 and (cur=0 or posCur<cur)) then
							cur=posCur
							tags=a(i)
							yllen=instr(posCur,sTemplate,"}")
							key=mid(sTemplate,posCur+len(a(i)),yllen-posCur-len(a(i)))
							if (cur <= 0) then exit for   '说明已经是最小了,可以退出
						end if
				next
				FindTags=cur
		End Function		
				
	   Function Parse(sTemplate, sTemp)
	     if lcase(sTemp)="nickname" then
		   echo KS.C("UserName")
		 elseif lcase(stemp)="showphotos" then
		   Call ShowPhotos()
		 else
	      echo GetNodeText(Lcase(sTemp))   
		 end if
	   End Function
	   
	   
		
		Function GetNodeText(NodeName)
		 Dim N,Str
		 NodeName=Lcase(NodeName)
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then Str=N.text
		  GetNodeText=Str
		 End If
		End Function
		
		'显示现场图片
		Sub ShowPhotos()
		  Dim RS:Set RS=Conn.Execute("select * From KS_InterViewPic Where InterViewID=" & ID & " order by id")
		  If Not RS.Eof Then
		    Do While Not RS.Eof
			 echo "<li><img style=""cursor:pointer"" onclick=""showPhoto(this.src,this.alt);"" src='" & RS("PhotoUrl") & "' alt='" & RS("Content") & "' border='0'/><br/><span>" & rs("content")&"</span></li>"
			RS.MoveNext
			Loop
		  End If
		  RS.Close
		  Set RS=Nothing
		End Sub
	  
End Class
%>

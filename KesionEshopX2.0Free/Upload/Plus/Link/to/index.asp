<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New ToLink
KSCls.Kesion()
Set KSCls = Nothing

Class ToLink
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim LinkID, ObjRS,Url,SiteName
		LinkID = KS.ChkClng(request.QueryString)
		Set ObjRS = Server.CreateObject("Adodb.RecordSet")
		ObjRS.Open "Select top 1 Url,hits,SiteName From KS_Link Where LinkID=" & LinkID, Conn, 1, 3
		If Not ObjRS.EOF Then
		  ObjRS(1) = ObjRS(1) + 1
		  ObjRS.Update
		  Url=ObjRS(0)
		  sitename=ObjRS(2)
		  ObjRS.Close:Set ObjRS=Nothing
		  
		  
		   '========点友情链接加积分==================
		 if KS.Setting(168)="1" And KS.ChkClng(KS.Setting(169))>0 Then
		   If KS.C("UserName")<>"" Then
		   
		     If KS.Setting(145)="0" Then  '积分
				  If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1001 and infoid=" & LinkID).Eof Then
					 '判断有没有到达每天增加的总限
					 Dim TodayScore:TodayScore=0
					 If KS.ChkClng(KS.Setting(165))<>0 Then
					  TodayScore=KS.ChkClng(Conn.Execute("select sum(Score) from ks_logscore where InOrOutFlag=1 and year(adddate)=year(" & SQLNowString & ") and month(adddate)=month(" & SQLNowString & ") and day(adddate)=day(" & SQLNowString & ") and username='" & ks.c("UserName") & "'")(0))
					 End If
					 If TodayScore+KS.ChkClng(KS.Setting(169))<KS.ChkClng(KS.Setting(165)) Then
	
						  Conn.Execute("Update KS_User Set Score=Score+" & KS.ChkClng(KS.Setting(169)) & " Where UserName='" & KS.C("UserName") & "'")
						  'on error resume next
						  Dim CurrScore:CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & KS.C("UserName") & "'")(0)
						  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP,Channelid,InfoID) values('" & KS.C("UserName") & "',1," & KS.ChkClng(KS.Setting(169)) & ","&CurrScore & ",'系统','点击友情链接[" & sitename & "(" & url & ")]所得!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "',1001," & LinkID & ")")
					 End If
	
				  End If
			ElseIf KS.Setting(145)="1" Then  '点券
			     If Conn.Execute("Select top 1 * From KS_LogPoint Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1001 and infoid=" & LinkID).Eof Then
					  Call KS.PointInOrOut(1001,linkid,KS.C("UserName"),1,KS.ChkClng(KS.Setting(169)),"系统","点击友情链接[" & sitename & "(" & url & ")]所得!!",0)
			      End If
			Else
			         If Conn.Execute("Select top 1 * From KS_LogMoney Where UserName='" & KS.C("UserName") & "' and year(PayTime)=year(" & SQLNowString  &") and month(PayTime)=month(" & SQLNowString &") and day(PayTime)=day(" & SQLNowString & ") and channelid=1001 and infoid=" & linkid).Eof Then
			          Call KS.MoneyInOrOut(KS.C("UserName"),KS.C("UserName"),KS.ChkClng(KS.Setting(169)),4,1,now,0,"System","点击友情链接[" & sitename & "(" & url & ")]所得!!",1001,linkid,0)
			        End If
			
			
			End If
			  
		   End If
		 End If
		'=====================================
		  
		  
		  
		  Response.Redirect url
		Else
		  Response.Write "参数传递有误!"
		End If
		  ObjRS.Close
		  Set ObjRS = Nothing
		End Sub

End Class
%>

 

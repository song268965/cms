<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KS,KSUser
Set KS=New PublicCls
Dim ID,Node,Action,BSetting,LoginTF,Confirm,Score,LimitScore,FileName,Title,inputer
ID = KS.ChkClng(KS.S("ID"))
Action=KS.G("Action")
Confirm=KS.G("Confirm")


If Action="hits" Then
   Set RS=Conn.Execute("Select top 1 hits From KS_UploadFiles Where ID=" &ID)
   If RS.Eof Then
     response.Write "document.write('0');"
   ELSE
     Response.Write "document.write('" & RS(0) & "');"
   End If
   RS.Close : Set RS=Nothing
Else
   Set KSUser=New UserCls
   LoginTF=KSUser.UserLoginChecked
   Set RS=Server.CreateObject("adodb.recordset")
   RS.Open "Select top 1 * From KS_UploadFiles Where ID=" & ID,conn,1,1
   If RS.Eof Then
     RS.Close : Set RS=Nothing
	 head
     KS.Die "<script>$.dialog.tips('附件已不存在!',2,'error.gif',function(){});</script>"
   Else
	   FileName=RS("FileName")
	   Title=RS("Title")
	   Dim ChannelID:ChannelID=KS.ChkClng(RS("ChannelID"))
	   Dim InfoID:InfoID=KS.ChkClng(RS("InfoID"))
	   Dim ClassID:ClassID=RS("ClassID")
	   Dim UserName:UserName=RS("UserName")
	   If ChannelID>0 AND ChannelID<1000 AND InfoID>0 Then
		    Dim RSS:Set RSS=Conn.Execute("Select top 1 Title,inputer From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID) 
			If Not RSS.Eof Then
			   Title=RSS(0)
			   inputer=RSS(1)
			End If
			RSS.Close
			Set RSS=NOthing
		 End If
	  
	   RS.Close : Set RS=Nothing
	   If ChannelID<2000 Then      '模型附件
	     Dim AnnexPoint:AnnexPoint=KS.ChkClng(KS.C_S(ChannelID,50))
		 If AnnexPoint<=0 Then
		   Call DownLoad()
		 Else
		   Dim ModelChargeType:ModelChargeType=KS.ChkClng(KS.C_S(ChannelID,34))
		   Call CheckConfirm(AnnexPoint,ModelChargeType)
		 End If
	   ElseIf ChannelID=9994 and ClassID<>0 Then  '论坛附件
	     KS.LoadClubBoard
		 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & ClassID &"]")
		 If Node Is Nothing Then head:KS.Die "<script>$.dialog.tips('无法下载，非法参数,附件可能不存在了！',2,'error.gif',function(){});</script>"
		 BSetting=Node.SelectSingleNode("@settings").text
		 BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
		 BSetting=Split(BSetting,"$")
		 LimitScore=KS.ChkClng(BSetting(15))
		 Score=KS.ChkClng(BSetting(16))
		 If (LimitScore>0 or Score>0) And LoginTF=false Then
		  head:KS.Die "<script>$.dialog.tips('本附件设置需要积分验证,请先登录!',1,'error.gif',function(){ShowLogin();})</script>"
		 End If
		 If LimitScore>0 and KS.ChkClng(KSUser.GetUserInfo("Score"))<LimitScore Then
		  head:KS.Die "<script>$.dialog.alert('对不起,本附件设置用户积分达到" & LimitScore & "分才可以下载,您当前积分"+KSUser.GetUserInfo("Score")+"分!',function(){});</script>"
		 End If
		 If BSetting(0)="0" Then  '不允许游客浏览时才进一步判断权限
			 Dim CheckResult:CheckResult=CheckPermissions(KSUser,BSetting) '检查访问检查
			 If CheckResult<>"true" Then
			  %>
			   <html>
			   <head>
			    <title>没权限提示</title>
			   <style type="text/css">
			   .guest_box{text-align:center}
			   .errtips{margin:0px auto;background:url(../club/images/err.gif)  no-repeat 60px 40px; background-color:#ffffff;border:1px solid #f2f2f2;min-height:121px;width:500px;margin-top:80px;margin-bottom:120px;}
				.tishixx{text-align:left;word-break : break-all;width:360px;padding-left:128px;padding-top:20px;line-height:30px;font-size:14px;color:#000;}
				.tishi{font-size:14px;font-weight:bold}
				.tishixx span{color:red}
				.closebut{height:50px;line-height:50px; margin-left:130px;}
				.closebut a{border:1px solid #fff;padding:3px;margin:3px;width:100px;line-height:30px;height:30px;}
			   </style>
			   </head>
			   <body>
			  <%
			  KS.Die CheckResult
			 End If
		  End If
		  Call CheckConfirm(Score,2)
	   End If
	   
	   DownLoad()
   End If 
 
End If


'下载论坛附件，需先检查进入版面权限
Function CheckPermissions(KSUser,BSetting)
   If KSUser.GroupID="1" Then CheckPermissions="true":Exit Function
   Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(1)) and (KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false Or LoginTF=false) Then GroupPurview=false
   Dim UserPurview:UserPurview=True : If Not KS.IsNul(BSetting(10)) and (KS.FoundInArr(BSetting(10),KSUser.UserName,",")=false or LoginTF=false) Then UserPurview=false
   If KSUser.GetUserInfo("ClubSpecialPower")="1" Then UserPurview=true:GroupPurview=True
   Dim ScorePurview:ScorePurview=KS.ChkClng(BSetting(11))
   Dim MoneyPurview:MoneyPurview=KS.ChkClng(BSetting(12))
   Dim Edays:Edays=0:If LoginTF=True Then Edays=KSUser.GetEdays
   If  BSetting(0)="0" And KS.IsNul(KS.C("UserName")) Then
		CheckPermissions=GetClubErrTips("error1",true)
   ElseIf Bsetting(54)="2" And KS.ChkClng(Edays)>0 Then
	    CheckPermissions="true"
   ElseIf Bsetting(54)="1" And KS.ChkClng(Edays)<0 Then
		CheckPermissions=GetClubErrTips("error2",true)
   Else
	   If ((GroupPurview=false and Not KS.IsNul(BSetting(10))) or (UserPurview=false)) and boardid<>0 Then
			CheckPermissions=GetClubErrTips("error1",true)
	   ElseIf KS.ChkClng(KSUser.GetUserInfo("Score"))<ScorePurView And ScorePurView>0 Then
			CheckPermissions=Replace(Replace(GetClubErrTips("error3",true),"{$Tips}","积分<span>" &ScorePurView&"</span>分"),"{$CurrTips}","积分<span>" & KSUser.GetUserInfo("Score") & "</span>分")
	   ElseIf KS.ChkClng(KSUser.GetUserInfo("Money"))<MoneyPurview And MoneyPurview>0 Then
			CheckPermissions=Replace(Replace(GetClubErrTips("error3",true),"{$Tips}","资金￥<span>" &formatnumber(MoneyPurview,2,-1,-1)&"</span>元"),"{$CurrTips}","资金￥<span>" & formatnumber(KSUser.GetUserInfo("money"),2,-1,-1) & "</span>元")
	   Else
		  CheckPermissions="true"
	   End If
  End If
End Function
Function GetClubErrTips(ErrId,ShowBack)
    Dim Str:str="<div class=""guest_box""><div class=""errtips"">" &_
	           "<div  class=""tishixx"">" & LFCls.GetConfigFromXML("GuestBook","/guestbook/template",ErrId) & "</div>"&_
			   "<div class=""clear""></div>"
	If ShowBack Then
	     str=str &"<div class=""closebut""> <a href=""javascript:window.close()"">关闭本页</a></div>"
	End If
         GetClubErrTips=str &"</div></div>"
End Function

'权限下载附件并扣费处理
Sub CheckConfirm(Point,ModelChargeType)
  If Point<=0 Then DownLoad() : Exit Sub
	Dim ChargeStr,TableName,DateField,CurrPoint
	Select Case ModelChargeType
			case 0 ChargeStr=KS.Setting(46)&KS.Setting(45) : TableName="KS_LogPoint" : DateField="AddDate" : CurrPoint=KSUser.GetUserInfo("Point")
			case 1 ChargeStr="元人民币": TableName="KS_LogMoney" : DateField="PayTime": CurrPoint=KSUser.GetUserInfo("Money")
			case 2 ChargeStr="分积分": TableName="KS_LogScore": DateField="AddDate": CurrPoint=KSUser.GetUserInfo("Score")
			case else exit sub
	End Select
If Point>0 And Cbool(KSUser.UserLoginChecked)=false Then

        If instr(request.servervariables("http_user_agent"),"Mobile")>0 Then '手机访问,自动跳到手机版
             If KS.WSetting(0)=1 Then
				    KS.DIE "<script>top.location.href='" & KS.Setting(3) & KS.WSetting(4)&"/login.asp" & "';</script>"
				    Exit Sub
			 End If
		End If
		  head:KS.Die "<script>ShowLogin();</script>"
ElseIf Point>0 and KS.ChkClng(CurrPoint)<Point and ksuser.getedays<0 Then
		  head:KS.Die "<script>$.dialog.alert('对不起,下载本附件需要消费" & Point & ChargeStr & ",您当前剩余" & CurrPoint & ChargeStr&",不足支付!',function(){});</script>"
Elseif ksuser.getedays<=0 then			
  If Conn.Execute("Select top 1 * From " & TableName & " Where UserName='" & KSUser.UserName & "' and datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<24 and ChannelID=9994 and InfoID=" & ID).Eof And KSUser.UserName<>UserName Then
		       If Confirm<>"true" and KSUser.UserName<>inputer Then
		    	head:KS.Die "<script>$.dialog.confirm('下载本附件需要消费<font color=""red"">" & Point & "</font>"& ChargeStr & ",确定下载吗?',function(){location.href='" & KS.GetDomain & "item/filedown.asp?confirm=true&id=" & id&"&ext=" & request("ext") & "&fname=" & Request("FName") & "';},function(){});</script>"
			   Else
				 Dim PayPoint :PayPoint=0
				 If ClassID>0 Then
				  PayPoint=Point* (KS.C_CID(ClassID,11))/100
				 End If
			     Dim Remark:Remark="下载附件[附件:" & Title & "]!"
				 Dim TcMsg:TcMsg="附件“" & Title & "”被下载的提成"
			     Select Case ModelChargeType
				  case 0
					  If Round(KSUser.GetUserInfo("Point"))-round(point)<0 Then
					     head:KS.Die "<script>$.dialog.alert('对不起，您的可用" & ChargeStr & "不足支付!',function(){});</script>"
					  ElseIF Cbool(KS.PointInOrOut(9994,ID,KSUser.UserName,2,Point,"系统",Remark,0))=True Then 
					   If PayPoint>0 and Not KS.IsNul(inputer) Then Call KS.PointInOrOut(9994,ID,inputer,1,PayPoint,"系统",TcMsg,0)
					   DownLoad()
					  Else
					   head:KS.Die "<script>$.dialog.alert('扣费处理出错,请联系管理人员!',function(){});</script>"
					  End If
					  
				  case 1
					  If Round(KSUser.GetUserInfo("money"))-round(point)<0 Then
					     head:KS.Die "<script>$.dialog.alert('对不起，您的可用" & ChargeStr & "不足支付!',function(){});</script>"
					  ElseIF Cbool(KS.MoneyInOrOut(KSUser.UserName,KSUser.UserName,Point,4,2,now,0,"系统",Remark,9994,ID,1))=True Then 
					    If PayPoint>0 and Not KS.IsNul(inputer) Then Call KS.MoneyInOrOut(inputer,inputer,PayPoint,4,1,now,0,"系统",TcMsg,9994,ID,1)
					   DownLoad()
					  Else
					   head:KS.Die "<script>$.dialog.alert('扣费处理出错,请联系管理人员!',function(){});</script>"
					  End If
				  case 2
				   Session("ScoreHasUse")="+" '设置只累计消费积分
					If Round(KSUser.GetUserInfo("score"))-round(point)<0 Then
					     head:KS.Die "<script>$.dialog.alert('对不起，您的可用" & ChargeStr & "不足支付!',function(){});</script>"
					ElseIf Cbool(KS.ScoreInOrOut(KSUser.UserName,1,Point,"系统",Remark,9994,id)) Then
					   If PayPoint>0 and Not KS.IsNul(inputer) Then Call KS.PointInOrOut(9994,ID,inputer,1,PayPoint,"系统",TcMsg,0)
					  DownLoad()
					Else
					  head:KS.Die "<script>$.dialog.alert('扣费处理出错,请联系管理人员!',function(){});</script>"
					End If
					Session("ScoreHasUse")=""
				 end select
			   End If
  Else
		      DownLoad()
  End If
 End If
End Sub
Function ReplaceBadChar(str)
  str=replace(str,Chr(13)&Chr(10),"")
  str=replace(str,chr(13),"")
  str=replace(str,chr(10),"")
  str=replace(str,chr(9),"")
  str=replace(str,vbcrlf,"")
End Function
Sub DownLoad()
       Call KSUser.UseLogConsum(KS.C_S(3,6),3,ID,Title)   '新增，增加根据用户组里的设置，限制每天下载次数
       Conn.Execute("Update KS_UploadFiles Set Hits=Hits+1 Where ID=" & ID)
	   Dim FileOldName:FileOldName=ReplaceBadChar(Request("FName"))
	   'ks.die FileName
	   If KS.IsNul(FileOldName) Then
	   Response.Redirect FileName
	   Else
	    FileOldName=replace(FileOldName,"&amp;","&")
		if instr(lcase(FileOldName),lcase(request("ext")))=0 then FileOldName=FileOldName & "." & ReplaceBadChar(ks.s("ext"))
		if left(lcase(FileName),4)="http" then FileName=Replace(FileName,KS.Setting(2),"")
		if left(lcase(FileName),4)="http" or right(lcase(FileName),4)=".asp" then 
		  response.Redirect(FileName)
		else
	     call downloadFile(Server.MapPath(FileName),FileOldName)
		end if
	   End If
	   KS.Die ""
End Sub

Sub downloadFile(strFile,FileOldName) 
             strFile=trim(strFile)
            If instr("00000000"&lcase(strfile),".asp")<>0 or instr("00000000"&lcase(strfile),".aspx")<>0 or instr("00000000"&lcase(strfile),".asa")<>0 or instr("00000000"&lcase(strfile),".php")<>0 or instr("00000000"&lcase(strfile),".jsp")<>0 then
			  ks.die "<script>alert('非法下载!');window.close();</script>"
			end if

			
		    Server.ScriptTimeOut=999999 
			Dim fso,f,intFilelength,strFilename,DownFileName 
			Set fso = Server.CreateObject("Scripting.FileSystemObject") 
			If Not fso.FileExists(strFile) Then 
			  head:Response.Write("<script>$.dialog.alert('系统找不到指定文件',function(){});</script>") 
			  Exit Sub 
			End If 
			Set f = fso.GetFile(strFile) 
			Set fso=Nothing 
		    If KS.IsNul(FileOldName) Then DownFileName=f.name Else DownFileName=FileOldName 
		    Dim Stream,offset,TotalSize,ChunkSize ,strChunk
			Response.Buffer=False '将Response.Buffer设为否 
			Response.Charset="gb2312"
			response.CodePage=936
			Response.ContentType = "application/octet-stream" 
			response.AddHeader "Content-Disposition","attachment;filename=" & KS.CheckXSS(replace(DownFileName," ",""))
			Set Stream = Server.CreateObject("ADODB.Stream") 
			Stream.type=1 
			Stream.Open 
			Stream.LoadFromFile strFile
			offset = 0 
			ChunkSize = 2048*1024 'ChunkSize小于IIS配制文件中的AspBufferingLimit项所设置的大小 
			TotalSize = Stream.Size 
			while offset < TotalSize 
			if (TotalSize - offset < ChunkSize) then 
			ChunkSize = TotalSize-offset 
			end if 
			strChunk = Stream.Read(ChunkSize) 
			Response.BinaryWrite strChunk 
			offset = offset + ChunkSize 
			wend 
			Stream.Close
End Sub 

Sub Head()
%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name="author" content="kesion.com"/>
<script src="../ks_inc/jquery.js"></script>
<script src="../ks_inc/common.js"></script>
<body>
<%
End Sub
Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%> 

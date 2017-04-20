<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/UploadFunction.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
Server.ScriptTimeout=9999999
Response.CharSet="utf-8"
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
If Request("From")="Common" Then
 Set KSCls = New UpFileSaveByCommon
Else
 Set KSCls = New UpFileSaveBySwfUpload
End If
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileSaveBySwfUpload
        Private KS,KSUser,FileTitles,Title
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType,BoardID,EditorID
		Dim FormName,Path,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName,LoginTF
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,U_FileSize,FormID,FieldID,MustCheckSpaceSize,AllowNoUserUpload
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		Function CheckIsLogin(UserID,UserName,Pass)
		     If UserName="" Or Pass="" Or UserID="" Then Check=false: Exit Function
		     Dim ChkRS:Set ChkRS =Conn.Execute("Select top 1 * From KS_User Where UserID=" & KS.ChkClng(UserID))
			 If ChkRS.EOF And ChkRS.BOF Then
			   CheckIsLogin=false
			 Else
			   If ChkRS("RndPassWord")=Pass Then 
			               CheckIsLogin=true 
				           If EnabledSubDomain Then
							 Response.Cookies(KS.SiteSn).domain=RootDomain					
							Else
                             Response.Cookies(KS.SiteSn).path = "/"
							End If
						    Response.Cookies(KS.SiteSn).Expires = Date + 365
							Response.Cookies(KS.SiteSn)("UserID") = ChkRS("UserID")
							Response.Cookies(KS.SiteSn)("UserName") = ChkRS("UserName")
							Response.Cookies(KS.SiteSn)("Password") = ChkRS("PassWord")
							Response.Cookies(KS.SiteSN)("RndPassword")= ChkRS("RndPassWord")
							Response.Cookies(KS.SiteSN)("GroupID")= ChkRS("GroupID")
			   Else 
			    CheckIsLogin=false
			   End If
			 End If
		     ChkRS.Close:Set ChkRS = Nothing
		End Function
		
		Sub Kesion()
		

		Set UpFileObj = New UpFileClass
		on error resume next
		UpFileObj.GetData
		If ERR.Number<>0 Then Set UpFileObj=Nothing : err.clear:KS.Die "error:" & escape("上传失败，可能您的上传的文件太大!")
		EditorID =UpFileObj.Form("EditorID")
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType")) 
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 

		UpType=UpFileObj.Form("UpType")
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		LoginTF=Cbool(KSUser.UserLoginChecked)
		If LoginTF=false Then
		  If cbool(CheckIsLogin(UpFileObj.Form("UserID"),UpFileObj.Form("UserName"),UpFileObj.Form("RndPassWord"))) =true Then  '兼容Firefox
			 LoginTF=Cbool(KSUser.UserLoginChecked)
		  End If
		End If

		FormID=KS.ChkClng(UpFileObj.Form("FormID")) 
		FieldID=KS.ChkClng(UpFileObj.Form("FieldID")) 
		BoardID=KS.ChkClng(UpFileObj.Form("BoardID"))
		
		dim CurrentDir:CurrentDir=UpFileObj.Form("CurrentDir")
        CurrentDir=Trim(Replace(Replace(CurrentDir,"../",""),".",""))
		CurrentDir=KS.CheckXSS(CurrentDir)
		If KS.ChkClng(KS.C_S(ChannelID,6))<10 and KS.ChkClng(KS.C_S(ChannelID,6))>0 then
		 If KS.C_S(ChannelID,26)=0 then KS.Die "error:" & escape("对不起，此频道不允许上传!")
		end if
		
	    MustCheckSpaceSize=true : AllowNoUserUpload=0
		Dim RS,FieldName
		If UpType="Pic" And ChannelID<>9994 and channelid<>9993 Then
			If DefaultThumb=1 Then CreateThumbsFlag=true Else CreateThumbsFlag=false
			If ChannelID=7999 or ChannelID=7998  or channelid=9992 Then '企业动态,简介　
				MaxFileSize = 200    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(7999,KSUser.GetUserInfo("UserID"))
			ElseIf ChannelID=9996  Then '圈子图片　
				MaxFileSize = 100    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9996,KSUser.GetUserInfo("UserID"))
            ElseIf ChannelID=9998  Then '相册封面
				MaxFileSize = 100    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9998,KSUser.GetUserInfo("UserID"))
			ElseIf ChannelID=9999  Then   '用户头像
			    session("urel")=""
				MaxFileSize = 1024*2    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的图片
				FormPath = KS.ReturnChannelUserUpFilesDir(9999,KSUser.GetUserInfo("UserID"))
			ElseIf ChannelID=55666  Then   '用户头像
				session("urel")=""
				MaxFileSize = 1024    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的图片
				FormPath = KS.ReturnChannelUserUpFilesDir(55666,KSUser.GetUserInfo("UserID"))
			ElseIf ChannelID=9990 Then '企业图片
				MaxFileSize = 1024    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的图片
				FormPath = KS.ReturnChannelUserUpFilesDir(9990,KSUser.GetUserInfo("UserID"))
			ElseIf ChannelID=8000  Then  '模板DIY图片
				MaxFileSize = 500    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的图片
				FormPath = KS.ReturnChannelUserUpFilesDir(8000,KSUser.GetUserInfo("UserID"))
			Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
                If KS.C("UserName")="" And KS.C("PassWord")="" Then
					 If KS.C_S(ChannelID,26)=2 Then
					  AllowNoUserUpload=1: MustCheckSpaceSize=false
					  FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"  
					 End If
				Else
					AllowNoUserUpload=0
				    FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If				
			End If
		Elseif UpType="File" Then   '上传附件
			MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
			AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
			If KS.C("UserName")="" And KS.C("PassWord")="" And KS.ChkClng(KS.C_S(ChannelID,6))>0 Then
				If KS.ChkClng(KS.C_S(ChannelID,26))=2 Then
					  AllowNoUserUpload=1: MustCheckSpaceSize=false
					  FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"  
				End If
			Else
			  AllowNoUserUpload=0
			  FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			End If
		ElseIf (ChannelID=101 Or FormID<>0) and FieldID<>0 Then  '自定义表单或注册表单自定义字段
		    If ChannelID<>101 Then
			 Set RS=Conn.Execute("select top 1 AnonymousUpload From KS_Form Where ID=" & FormID)
			 If RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
			   KS.Die "error:" & escape("出错啦!")
			 End If
			 AllowNoUserUpload=rs(0)
			 RS.Close
			 Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_FormField Where FieldID=" & FieldID)
           Else
		      If KS.Setting(60)="1" Then AllowNoUserUpload=1 Else AllowNoUserUpload=0
			 Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_Field Where ChannelID=101 and FieldID=" & FieldID)
		   End If
			 If Not RS.Eof Then
				MaxFileSize=RS(2):AllowFileExtStr=RS(1)
				RS.Close
				If ChannelID=101 Then
				    If KS.C("UserName")<>"" Then
				 	FormPath =KS.Setting(3) & KS.Setting(91)& "user/" & KSUser.GetUserInfo("userid") & "/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"        
					Else
				 	FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"        
					End If
				Else
					Set RS=Conn.Execute("Select top 1 UploadDir From KS_Form Where ID=" &FormID)
					If Not RS.Eof Then 
					 FormPath =KS.Setting(3) & KS.Setting(91)&RS(0) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"        
					Else
					 FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"        
					End If
					End If
				If AllowNoUserUpload=1 Then MustCheckSpaceSize=false
			 Else
				KS.Die "error:" & escape("参数有误!")
			 End IF
			 RS.Close:Set RS=Nothing
       ElseIf ChannelID<>0 And BasicType<>0 and FieldID<>0 Then '模型自定义字段
	        Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & FieldID)
		   If Not RS.Eof Then
		    FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
			FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			If KS.C_S(ChannelID,26)=2 Then  
			 AllowNoUserUpload=1
			 If LoginTF=false Then 
			  MustCheckSpaceSize=false
			  FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"  
			 End If
			End If
		   Else
		    KS.Die "error:" & escape("参数有误!")
		   End IF
		   RS.Close:Set RS=Nothing
	   Else
			Select Case BasicType
			  Case 1,3,4,7,9    '下载,影片等
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					If BasicType=4 Then
					 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,2)
					ElseIf BasicType=7 Then  '影片
				     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,2) &"|" & KS.ReturnChannelAllowUpFilesType(ChannelID,3) & "|"& KS.ReturnChannelAllowUpFilesType(ChannelID,4)  
					Else
					 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
					End If
					If KS.C("UserName")="" And KS.C("PassWord")="" Then
					 If KS.C_S(ChannelID,26)=2 Then
					  AllowNoUserUpload=1: MustCheckSpaceSize=false
					  FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"  
					 End If
					Else
					 AllowNoUserUpload=0
			         FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					End If
			  Case 2     '图片中心
					CreateThumbsFlag=true
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
					If KS.C("UserName")="" And KS.C("PassWord")="" Then
					 If KS.C_S(ChannelID,26)=2 Then
					  AllowNoUserUpload=1: MustCheckSpaceSize=false
					  FormPath =KS.Setting(3) & KS.Setting(91)&"Temp/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"  
					 End If
					Else
					  AllowNoUserUpload=0
					  FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					End If
					
			case 8666		'会员中心短消息
			  If KS.ChkClng(KS.U_S(KSUser.GroupID,22))=1 Then
				MaxFileSize = KS.ChkClng(KS.U_S(KSUser.GroupID,24))    '设定文件上传最大字节数
				AllowFileExtStr = KS.U_S(KSUser.GroupID,23) '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(8666,KSUser.GetUserInfo("UserID"))
			  Else
			    KS.Die "error:" & escape("对不起，此频道不允许上传附件!")
			  End If
			Case 9995  '音乐
				MaxFileSize = 102400    '设定文件上传最大字节数
				AllowFileExtStr = "mp3"  '取允许上传的动漫类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9995,KSUser.GetUserInfo("UserID"))
			 Case 9997  '相片
				MaxFileSize = KS.ChkClng(KS.SSetting(32))    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
				FormPath = KS.ReturnChannelUserUpFilesDir(9997,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			Case 9992 '问答附件	
			 	 If KS.ASetting(42)<>"1" Then
				   KS.Die "error:" & escape("对不起，此频道不允许上传附件!")
				ElseIf LoginTF=false or (not KS.IsNul(KS.ASetting(46)) and KS.FoundInArr(KS.ASetting(46),KSUser.GroupID,",")=false) Then
				 KS.Die "error:" & escape("对不起,您没有在此频道上传的权限!")
                 End If
				MaxFileSize =KS.ChkClng(KS.ASetting(44))    '设定文件上传最大字节数
				AllowFileExtStr = KS.ASetting(43)  '取允许上传的类型
				FormPath = KS.ReturnChannelUserUpFilesDir(9997,KSUser.GetUserInfo("UserID")) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			 Case 9994  '论坛
			    If BoardID=0 Then
				  KS.Die "error:" & escape("非法参数!")
				End If
				KS.LoadClubBoard
				Dim BNode,BSetting
				Set BNode=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & BoardID &"]") 
				If BNode Is Nothing Then KS.Die "error:" & escape("非法参数!")
				BSetting=BNode.SelectSingleNode("@settings").text
				BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				BSetting=Split(BSetting,"$")
				If KS.ChkClng(BSetting(36))<>1 Then
				  KS.Die "error:" & escape("对不起，系统不允许此频道上传文件,请与网站管理员联系!")
				End If
				If  LoginTF=true  and (KS.IsNul(BSetting(17)) Or KS.FoundInArr(BSetting(17),KSUser.GroupID,",")) Then
				    If KS.ChkClng(BSetting(39))<>0 Then
					 If Conn.Execute("select count(1) From KS_UploadFiles Where ClassID=" & BoardID & " and datediff(" & DataPart_D & ",AddDate," & SQLNowString & ")<1 and username='" & KSUser.UserName &"'")(0)>=KS.ChkClng(BSetting(39)) Then
					  KS.Die "error:" & escape("对不起，本版面每天每人限制只能上传" & KS.ChkClng(BSetting(39))&"个文件!")
					 End If
					End If
					MaxFileSize = KS.ChkClng(BSetting(38))    '设定文件上传最大字节数
					AllowFileExtStr = BSetting(37)  '取允许上传的类型
					FormPath =KS.ReturnChannelUserUpFilesDir(9994,KS.Setting(67))
				Else
				  KS.Die "error:" & escape("对不起，您没有在本论坛上传附件的权限!")
				End If
			Case 9993  '写日志附件
			    If KS.ChkClng(KS.SSetting(26))=0 Then
				  KS.Die "error:" & escape("对不起，系统不允许此频道上传文件,请与网站管理员联系!")
			   ElseIf LoginTF=false or (not KS.IsNul(KS.SSetting(30)) and KS.FoundInArr(KS.SSetting(30),KSUser.GroupID,",")=false) Then 
			    KS.Die "error:" & escape("对不起,您没有在此频道上传的权限!")
			   End If
				MaxFileSize = KS.ChkClng(KS.SSetting(28))    '设定文件上传最大字节数
				AllowFileExtStr = KS.SSetting(27)  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9993,KSUser.GetUserInfo("UserID"))
			Case 9991  '微博
			    If KS.ChkClng(KS.SSetting(50))=0 Then
				  KS.Die "error:" & escape("对不起，系统不允许此频道上传文件,请与网站管理员联系!")
			   ElseIf LoginTF=false or (not KS.IsNul(KS.SSetting(53)) and KS.FoundInArr(KS.SSetting(53),KSUser.GroupID,",")=false) Then 
			    KS.Die "error:" & escape("对不起,您没有在此频道上传的权限!")
			   End If
				MaxFileSize = KS.ChkClng(KS.SSetting(51))    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9994,KS.SSetting(54))
			Case 99999
				MaxFileSize = KS.U_S(KSUser.GroupID,24)    '设定文件上传最大字节数
				AllowFileExtStr = "gif|jpg|png|swf|flv|mp3|doc"  '取允许上传的类型
				if CurrentDir<>"" then 
				FormPath =KS.ReturnChannelUserUpFilesDir(99999,KSUser.GetUserInfo("UserID") &"/" & CurrentDir)
				else
				FormPath =KS.ReturnChannelUserUpFilesDir(99999,KSUser.GetUserInfo("UserID"))
				end if
				Formpath=replace(FormPath,"//","/")
			End Select
		End If
		If AllowNoUserUpload="0" And LoginTF=false Then 
		   KS.Die "error:" & escape("对不起，只有登录后才可以使用上传!")
		End If

		FormPath=Replace(FormPath,".","")
		IF Instr(FormPath,KS.Setting(3))=0 Then FormPath=KS.Setting(3) & FormPath
		FilePath=Server.MapPath(FormPath) & "\"
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
        If KS.ChkClng(KS.Setting(97))=1 Then FormPath=KS.Setting(2) & FormPath
		ReturnValue = CheckUpFile(KSUser,MustCheckSpaceSize,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)

		If Not KS.IsNul(UpFileObj.Form("fileNames")) Then FileTitles=unescape(UpFileObj.Form("fileNames")) '防止中文乱码
		If UpFileObj.Form("NoReName")="1" Then  '不更名
		        Dim PhysicalPath,FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
		        PhysicalPath = Server.MapPath(replace(TempFileStr,"|",""))
				TempFileStr= mid(TempFileStr,1, InStrRev(TempFileStr, "/")) &  FileTitles
				
				 Dim NoAllowExtArr:NoAllowExtArr=split(NoAllowExt,"|")
				 Dim KK
				 for kk=0 to ubound(NoAllowExtArr)
						   if instr(replace(lcase(TempFileStr),lcase(KS.Setting(2)),""),"." & NoAllowExtArr(kk))<>0 then
							 If FsoObj.FileExists(PhysicalPath)=true Then call KS.DeleteFile(PhysicalPath)
							 KS.Die "error:" & escape("文件上传失败,文件名不合法!")
						   end if
				 Next
				
				If FsoObj.FileExists(PhysicalPath)=true Then
				 FsoObj.MoveFile PhysicalPath,server.MapPath(TempFileStr)
			    End If
		End If

		if ReturnValue <> "" then
		     ReturnValue=replace(ReturnValue,"\n","。")
		     If Instr(ReturnValue,"上传失败")<>0 Then
		     KS.Die "error:" & escape("上传失败" & Replace(Split(ReturnValue,"上传失败")(1),"'","\'"))
			 Else
		     KS.Die "error:" & escape(Replace(ReturnValue,"'","\'"))
			 End If
		else 
			 TempFileStr=replace(TempFileStr,"'","\'")
			 If UpType="Field" Then
			 	KS.Die replace(TempFileStr,"|","")
			 Elseif UpType="File" Or UpType="BBSFile" Then   '上传附件
				  Call AddAnnexToDB(ChannelID,KS.C("UserName"),TempFileStr,FileTitles,KS.ChkClng(BoardID),false,EditorID)
			 ElseIf UpType="Pic" Then

				   if basictype=9999 then
			 		Call KSUser.AddToWeibo(KSUser.UserName,"更换了自己的形象照片，[url={$GetSiteUrl}user/weibo.asp?userid=" & KSUser.GetUserInfo("userid") &"]TA的微博[/url] [url={$GetSiteUrl}space/?" & KSUser.GetUserInfo("userid") &"]TA的空间[/url][br][url=" & replace(TempFileStr,"|","") & "][img]" & replace(TempFileStr,"|","") & "[/img][/url]",6)
				  end if
			    
			      If BasicType=1 Or BasicType=5 Or BasicType=3  Or BasicType=8 or channelid=7999 or channelid=7998 Then
				   if ThumbPathFileName="" then ThumbPathFileName=replace(TempFileStr,"|","")
			       KS.Die ThumbPathFileName  &"@"& replace(TempFileStr,"|","") 
				  Else
				   if DefaultThumb=1 then
				     KS.Echo ThumbPathFileName
 					 if replace(ThumbPathFileName,"|","")<>replace(TempFileStr,"|","") then
				      Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
					 end if
				   Else
				     KS.Echo escape(replace(TempFileStr,"|",""))
				   End If
                   KS.Die ""
				  End If
			 Else
				 Select Case BasicType
				      Case 3 KS.Die escape(replace(TempFileStr,"|","")) & "|" & U_FileSize
					  Case 2         '图片
					       if ThumbPathFileName="" then ThumbPathFileName=replace(TempFileStr,"|","")
						   KS.Die replace(TempFileStr,"|","") &  "@" & ThumbPathFileName & "@" & escape(FileTitles)
					  Case 9997    '相片
						   KS.Die replace(TempFileStr,"|","") &  "@" & replace(TempFileStr,"|","") & "@" & escape(FileTitles)
					  Case Else KS.Die escape(replace(TempFileStr,"|",""))
				 End Select
			 End If
		  End iF
		Set UpFileObj=Nothing
 End Sub
End Class




'普通上传处理类 
Class UpFileSaveByCommon
        Private KS,KSUser,FileTitles,BoardID,LoginTF
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType
		Dim FormName,Path,TempFileStr,FormPath,ThumbPathFileName
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		Function IsSelfRefer() 
			Dim sHttp_Referer, sServer_Name 
			sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER")) 
			sServer_Name = CStr(Request.ServerVariables("SERVER_NAME")) 
			If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then 
			IsSelfRefer = True 
			Else 
			IsSelfRefer = False 
			End If 
		End Function 
		Sub Kesion()
		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {background:#f0f0f0;" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)
		
		If Cbool(KSUser.UserLoginChecked)=false Then
			Response.Write "<script>alert('没有登录!');history.back();</script>"
			Response.end
		End If
		
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('非法上传1！');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"user_upfile.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"post.asp")<=0 then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 end if
		 if IsSelfRefer=false Then
			Response.Write "<script>alert('请不要非法上传！');history.back();</script>"
			Response.end
		 End If
		 
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		BoardID=KS.ChkClng(UpFileObj.Form("BoardID"))
		LoginTF=Cbool(KSUser.UserLoginChecked)
		
		Select Case ChannelID
		  Case 9999
		    FormPath=KS.ReturnChannelUserUpFilesDir(9999,KSUser.GetUserInfo("UserID")) '上传头像路径
			MaxFileSize = 1024*5    '设定文件上传最大字节数
			AllowFileExtStr ="jpg|gif|png"
		  Case 9994  '论坛
			    If BoardID=0 Then
				  KS.Die "<script>alert('非法参数!');</script>"
				End If
				KS.LoadClubBoard
				Dim BNode,BSetting
				Set BNode=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & BoardID &"]") 
				If BNode Is Nothing Then KS.Die "<script>alert('非法参数!');</script>"
				BSetting=BNode.SelectSingleNode("@settings").text
				BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				BSetting=Split(BSetting,"$")
				If KS.ChkClng(BSetting(36))<>1 Then
				   KS.Die "<script>alert('对不起，系统不允许此频道上传文件,请与网站管理员联系!');</script>"
				End If
				If  LoginTF=true  and (KS.IsNul(BSetting(17)) Or KS.FoundInArr(BSetting(17),KSUser.GroupID,",")) Then
				    If KS.ChkClng(BSetting(39))<>0 Then
					 If Conn.Execute("select count(1) From KS_UploadFiles Where ClassID=" & BoardID & " and datediff(" & DataPart_D & ",AddDate," & SQLNowString & ")<1 and username='" & KSUser.UserName &"'")(0)>=KS.ChkClng(BSetting(39)) Then
					   KS.Die "<script>alert('对不起，本版面每天每人限制只能上传" & KS.ChkClng(BSetting(39))&"个文件!');</script>"
					 End If
					End If
					MaxFileSize = KS.ChkClng(BSetting(38))    '设定文件上传最大字节数
					AllowFileExtStr = BSetting(37)  '取允许上传的类型
					FormPath =KS.ReturnChannelUserUpFilesDir(9994,KS.Setting(67))
				Else
				  KS.Die "<script>alert('对不起，您没有在本论坛上传附件的权限!');</script>"
				End If
		  case else
		    ks.die "error!"
        End Select	
		
		
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		FilePath=Server.MapPath(FormPath) & "\"
		If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
			
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		
		If UpType="Pic" Then
			If DefaultThumb=1 Then CreateThumbsFlag=true Else CreateThumbsFlag=false
		End If	
		

		ReturnValue = CheckUpFile(KSUser,true,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
		
		if ReturnValue <> "" then
		     ReturnValue = Replace(ReturnValue,"'","\'")
		     Response.Write "<script>alert('" & ReturnValue &"');history.back();</script>"
			 Response.End()
		else 
			    TempFileStr=replace(TempFileStr,"'","\'")
				   If BasicType=9999 Then  '头像
				    Dim UserFace :UserFace=replace(TempFileStr,"|","")
				    if DefaultThumb=0 then
					else
					        Dim ThumbFileName:ThumbFileName=split(UserFace,".")(0)&"_S.jpg"
					        Dim FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
							If FsoObj.FileExists(server.MapPath(ThumbFileName))=true Then
							 Call KS.DeleteFile(UserFace)  '删除原图片
							 FsoObj.MoveFile Server.MapPath(ThumbFileName),server.MapPath(UserFace)
							End If
					end if
				     Conn.Execute("Update KS_User Set UserFace='"& UserFace &"' Where UserName='" & KSUser.UserName&"'")
					 Session(KS.SiteSN&"UserInfo")=""
					 KS.Die "<script>alert('恭喜，头像上传成功!');top.location.reload();</script>"
				   ElseIf BasicType=9994 Then
				     KS.Die "<script>parent.uploadOk('" & replace(TempFileStr,"|","") & "');</script>"
				   Else
				          Response.Write("<script language=""JavaScript"">")
						  if DefaultThumb=0 then
						   Response.Write("parent.document.getElementById('PhotoUrl').value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
						   response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
						  else
						   Response.Write("parent.document.getElementById('PhotoUrl').value='" & ThumbPathFileName & "';")
						   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
						  end if
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../System/KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
						  Response.Write("</script>")
				   End IF
		  End iF
		Set UpFileObj=Nothing
		End Sub
		
		'上传默认图成功
		Sub SuccessDefaultPhoto()
	      Response.Write("<script language=""JavaScript"">")
		    if DefaultThumb=0 then
				 Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
		    else
				 Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
				 Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
			end if
		   Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../System/KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
		  Response.Write "</script>"
		End Sub
			
End Class
%> 

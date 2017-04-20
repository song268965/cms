<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../ks_cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../plus/md5.asp"-->
<!--#include file="Include/3GCls.asp"-->
<!--#include file="../api/cls_api.asp"-->
<%
'****************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: htF_C://www.kesion.com htF_C://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New RegCls
KSCls.Kesion()
Set KSCls = Nothing

Class RegCls
        Private KS,KSR,F_C
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
		Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#include file="include/function.asp"-->
		<%
		Public Sub Kesion()
		   if request.ServerVariables("HTF_C_REFERER")<>"" then
		    if instr(request.ServerVariables("HTF_C_REFERER"),"reg.asp")=0 then
		      GCls.ComeUrl=request.ServerVariables("HTF_C_REFERER")
		    end if
		  end if
		  IF KS.Setting(21)=0 Then 
		   KS.Die "<script>alert('对不起，本站暂停新会员注册!');history.back();</script>"
		  End If
		   F_C = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & TemplatePath & "/reg.html")
		   FCls.RefreshType="UserRegStep1"
		   If Trim(F_C) = "" Then F_C = "模板不存在!"
          Dim UserRegMustFill:UserRegMustFill=KS.Setting(33)
		  Dim ShowCheckEmailTF:ShowCheckEmailTF=true
		  Dim ShowVerifyCodeTF:ShowVerifyCodeTF=false
		 
		 IF KS.Setting(28)="1" Then ShowCheckEmailTF=false
		 IF KS.Setting(27)="1" then ShowVerifyCodeTF=true
		 
		 If KS.Setting(33)="0" Then
		 F_C = Replace(F_C, "{$ShowUserType}", "<input type='hidden' id='GroupID' value='2'/>")
		 F_C = Replace(F_C, "{$DisplayUserType}", " style='display:none'")
		 Else
		 F_C = Replace(F_C, "{$ShowUserType}", UserGroupList())
		 F_C = Replace(F_C, "{$DisplayUserType}", "")
		 End If
		 
		 If KS.Setting(32)="1" Then 
		 F_C = Replace(F_C, "{$Show_Detail}", " style='display:none'")
		 F_C = Replace(F_C, "{$Show_DetailTF}", 1)
		 Else
		 F_C = Replace(F_C, "{$Show_Detail}", "")
		 F_C = Replace(F_C, "{$Show_DetailTF}", 2)
		 End If
		 
		 If KS.Setting(148)="1" Then
		 F_C = Replace(F_C, "{$DisplayQestion}", "")
		 Else
		 F_C = Replace(F_C, "{$DisplayQestion}", " style=""display:none""")
		 End If

		 If KS.Setting(149)="1" Then
		 F_C = Replace(F_C, "{$DisplayMobile}", "")
		 Else
		 F_C = Replace(F_C, "{$DisplayMobile}", " style=""display:none""")
		 End If
		 If KS.Setting(143)="1" Then
		 F_C = Replace(F_C, "{$DisplayAlliance}", "")
		 Else
		 F_C = Replace(F_C, "{$DisplayAlliance}", " style=""display:none""")
		 End If
		 
		 F_C = Replace(F_C, "{$Show_OutTimes}", KS.ChkClng(split(KS.Setting(156)&"∮","∮")(1)))
		 If KS.IsNul(Split(KS.Setting(155)&"∮","∮")(0)) or KS.Setting(157)="0" Then
		 F_C = Replace(F_C, "{$DisplayMobileCode}"," style=""display:none""")
		 F_C = Replace(F_C,"{$Show_MobileCodeTF}",0)
		 F_C = Replace(F_C, "{$Show_Mobile}", KS.Setting(149))
		 Else
		 F_C = Replace(F_C,"{$Show_MobileCodeTF}",1)
		 F_C = Replace(F_C, "{$Show_Mobile}", "1")
		 End If
		 
		 If Mid(KS.Setting(161),1,1)="1" Then
		 Dim RndReg:rndReg=GetRegRnd()
		 F_C = Replace(F_C, "{$DisplayRegQuestion}", "")
		 F_C = Replace(F_C, "{$RegQuestion}", GetRegQuestion(RndReg))
		 F_C = Replace(F_C, "{$AnswerRnd}", GetRegAnswerRnd(RndReg))
		 Else
		 F_C = Replace(F_C, "{$DisplayRegQuestion}", " style=""display:none""")
		 F_C = Replace(F_C, "{$RegQuestion}", "")
		 F_C = Replace(F_C, "{$AnswerRnd}", "")
		 End If
		 
		 F_C = Replace(F_C, "{$Show_Question}", KS.Setting(148))
		 If Request("u")<>"" Then
		 F_C = Replace(F_C, "{$UserName}", " value=""" & split(Request("u"),"@")(0) & """")
		 Else
		 F_C = Replace(F_C, "{$UserName}", "")
		 End If
		 If KS.S("Uid")<>"" Then
		  if not Conn.Execute("Select top 1 UserName From KS_User Where UserID=" & KS.ChkClng(KS.S("Uid"))).eof then
		  F_C = Replace(F_C, "{$AllianceUser}", " value=""" & Conn.Execute("Select top 1 UserName From KS_User Where UserID=" & KS.ChkClng(KS.S("Uid")))(0) & """ readonly")
		  End If
		  F_C = Replace(F_C, "{$Friend}", " value=""" & KS.S("F") & """")
		 Else
		  F_C = Replace(F_C, "{$AllianceUser}", "")
		  F_C = Replace(F_C, "{$Friend}", "")
		 End If

		 F_C = Replace(F_C, "{$GetUserRegLicense}", Replace(KS.Setting(23),chr(10),"<br/>"))
		 F_C = Replace(F_C,"{$Show_UserNameLimitChar}",KS.Setting(29))
		 F_C = Replace(F_C,"{$Show_UserNameMaxChar}",KS.Setting(30))
		 F_C = Replace(F_C, "{$Show_CheckEmail}", IsShow(ShowCheckEmailTF))
		 F_C = Replace(F_C, "{$Show_VerifyCodeTF}", IsShow(ShowVerifyCodeTF))
		 
		 Dim LoginStr
		 If cbool(API_QQEnable) Then
		   LoginStr="<a title=""使用qq账号登录"" href=""" & KS.GetDomain & "api/qq/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_qq.png"" align=""absmiddle""/> QQ登录</a>&nbsp;&nbsp;"
		 End If
		 If cbool(API_SinaEnable) Then
		   LoginStr=LoginStr & " <a title=""使用新浪微博账号登录"" href=""" & KS.GetDomain & "api/sina/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_sina.png"" align=""absmiddle""/> 新浪微博</a>&nbsp;&nbsp;"
		 End If
		 If cbool(API_AlipayEnable) Then
		   LoginStr=LoginStr & " <a title=""使用支付宝登录"" href=""" & KS.GetDomain & "api/alipay/alipay_auth_authorize.asp""><img src=""" &KS.GetDomain & "images/default/icon_alipay.png"" align=""absmiddle""/> 支付宝</a>"
		 End If
		 If LoginStr<>"" Then LoginStr=LoginStr & "<br/><div style='color:#999;margin:5px'>与合作网站内容互通，快速登录</div>"

		 F_C=Replace(F_C,"{$ShowQuickLogin}",LoginStr)		 
		 
	     InitialCommon
		  
         F_C = KSR.KSLabelReplaceAll(F_C) '替换函数标签
		 Response.Write F_C  
		end sub
		
		Function GetRegRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.GetCurrQuestion(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetRegRnd=RandNum
		End Function
		Function GetRegQuestion(ByVal RndReg)
		  Dim QuestionArr:QuestionArr=Split(KS.GetCurrQuestion(162),vbcrlf)
		  GetRegQuestion=QuestionArr(rndReg)
		End Function
		Function GetRegAnswerRnd(ByVal RndReg)
		  GetRegAnswerRnd=md5(rndReg,16)
		End Function        '会员类型
		Function UserGroupList()
			If  KS.Setting(33)="0" Then UserGroupList="":Exit Function
			 Dim Node,Tips
			 Call KS.LoadUserGroup()
			 If KS.ChkClng(KS.S("GroupID"))<>0 Then
				Set Node=Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectSingleNode("row[@id=" & KS.S("GroupID") & "]")
				If Not Node Is Nothing Then
				If KS.ChkClng(Node.SelectSingleNode("@showonreg").text)=0 Then KS.Die "<script>alert('对不起，该用户组不允许注册!');</script>"
				UserGroupList="<span style='font-weight:bold;color:#ff6600'>" & Node.SelectSingleNode("@groupname").text &"</span><input type='hidden' value='" & KS.S("GroupID") & "' id='GroupID' name='GroupID'><span style='display:none' id='tips_" &Node.SelectSingleNode("@id").text&"'>" &Node.SelectSingleNode("@descript").text &"</span>"
			    End If 
				Set Node=Nothing
			Else
			  For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
			  If UserGroupList="" Then
			  Tips=Node.SelectSingleNode("@descript").text
			  UserGroupList="<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"" checked>" & Node.SelectSingleNode("@groupname").text  & "</label><span style='display:none' id='tips_" &Node.SelectSingleNode("@id").text&"'>" &Node.SelectSingleNode("@descript").text &"</span>"
			  Else
			  UserGroupList=UserGroupList & " <label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"">" & Node.SelectSingleNode("@groupname").text & "</label><span style='display:none' id='tips_" &Node.SelectSingleNode("@id").text&"'>" &Node.SelectSingleNode("@descript").text &"</span>"
			  End If
			 Next
			End If
		End Function
		
		Function IsShow(Show)
			If Show =true Then
				IsShow = ""
			Else
				IsShow = " Style='display:none'"
			End If
		End Function		
		
		'===================会员注册结束=====================
End Class
%>
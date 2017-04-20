<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<%
const uploaddir="smrz"    '定义证件上传的目录，建议更改

'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
If Request("action")="checkemail" Then
   checkemail
   response.end
End If

Sub checkemail()
    Dim KS:Set KS=New PublicCls
	Dim UserID:UserID=KS.ChkClng(KS.G("UserID"))
	Dim RndPassWord:RndPassWord=KS.G("Rnd")
	Dim RS:Set RS=Server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from ks_user where userid=" & userid & " and rndpassword='" & RndPassWord & "'",conn,1,1
	If RS.Eof And rs.bof Then
	  RS.CLose:Set RS=Nothing
	  KS.Die "<script>alert('邮箱认证失败，请重新发送认证确认信!');location.href='User_rz.asp';</script>"
	End If
	conn.execute("update ks_user set email='" & KS.G("Email") & "',isemailrz=1,rndpassword='" & KS.GetRndPassword(10) & "' where userid=" & userid)
	set ks=nothing
	response.write "<script>alert('恭喜，邮箱认证成功!');location.href='user_rz.asp';</script>"
End Sub

Dim KSCls
Set KSCls = New User_RZ
KSCls.Kesion()
Set KSCls = Nothing

Class User_RZ
        Private KS,KSUser
		Private CurrentPage,totalPut,TotalPages,SQL
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
		Public Sub loadMain()	
			IF Cbool(KSUser.UserLoginChecked)=false Then
			  Response.Write "<script>top.location.href='Login';</script>"
			  Exit Sub
			End If
			KSUser.CheckPowerAndDie("s08")
			
			Call KSUser.Head()
			Call KSUser.InnerLocation("实名认证")
			Select Case KS.S("Action")
			  Case "yyrz" yyrz
			  Case "yyrzSave" yyrzSave
			  Case "sfz" sfz
			  Case "sfzSave" sfzSave
			  Case "mobilerz" mobilerz
			  Case "mobilerzSave" mobilerzSave
			  Case "emailrz" emailrz
			  Case "emailrzSave" emailrzSave
			  Case Else Main
			End Select
	   End Sub	
	   

	   
	   Sub yyrz()
	    KSUser.InnerLocation("营业执照认证")
		Dim CompanyName,BusinessLicense,LegalPeople,Address,RegisteredCapital,Foundation,Business,photourl,isrz
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "select top 1 * From KS_EnterPrise Where UserName='" & KSUser.UserName & "'",conn,1,1
		If Not RS.Eof Then
		  ComPanyName=RS("CompanyName")
		  LegalPeople=RS("LegalPeople")
		  BusinessLicense=RS("BusinessLicense")
		  Address=RS("Address")
		  RegisteredCapital=RS("RegisteredCapital")
		  Foundation=RS("Foundation")
		  Business=RS("Business")
		  photourl=RS("photourl")
		  isrz=rs("isrz")
		Else
		  isrz=0
		End If
		RS.Close:Set RS=Nothing
		%>
		 <%if isrz="1" then%>
		 <table  cellspacing="1" cellpadding="3" class="border" align="center" border="0">
                          <tr class="tdbg">
                            <td class="clefttitle" align="right">公司名称：</td>
                            <td><%=CompanyName%></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle" align="right">注 册 号：</td>
                            <td><%=BusinessLicense%></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle" align="right">企业法人：</td>
                            <td><%=LegalPeople%></td>
                          </tr>
						   <tr class="tdbg">
                            <td class="clefttitle" align="right">公司地址：</td>
                            <td><%=Address%></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">注册资金：</td>
                            <td><%=RegisteredCapital%></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">经营范围：</td>
                            <td><%=Business%></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">成立日期：</td>
                            <td><%=Foundation%></td>
                          </tr>
						
						<%if photourl<>"" then%>
                       <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照预览：</td>
						  <td><a href="<%=photourl%>" target="_blank"><img src="<%=photourl%>" width="200" border="0" style="border:1px solid #cccccc;padding:1px"/></a>
						  </td>
						</tr>
					<%end if%>
						<tr class="tdbg">
                            <td class="clefttitle">&nbsp;</td>
                            <td><button class="pn" name="Submit" onclick="history.back(-1);" type="button"><strong> 返 回 </strong></button></td>
                          </tr>
		</table>
		 <%else%>
		  <script type="text/javascript">
		  function CheckForm(){
		    if ($("#CompanyName").val()==''){
			  alert('请输入公司名称!');
			  $("#CompanyName").focus();
			  return false;
			}
			if ($("#BusinessLicense").val()==''){
			  alert('请输入注册号!');
			  $("#BusinessLicense").focus();
			  return false;
			}
			if ($("#LegalPeople").val()==''){
			  alert('请输入企业法人!');
			  $("#LegalPeople").focus();
			  return false;
			}
			if ($("#Address").val()==''){
			  alert('请输入公司地址!');
			  $("#Address").focus();
			  return false;
			}
			if ($("#RegisteredCapital").val()==''){
			  alert('请输入注册资金!');
			  $("#RegisteredCapital").focus();
			  return false;
			}
			if ($("#Business").val()==''){
			  alert('请输入经营范围!');
			  $("#Business").focus();
			  return false;
			}
			return true;
		  }
		 </script>
					  <form action="?Action=yyrzSave" method="post" name="myform"  enctype="multipart/form-data" id="myform" onSubmit="return CheckForm();">
					      <input type="hidden" value="<%=KS.S("ComeUrl")%>" name="ComeUrl">
		 <table  cellspacing="1" cellpadding="3" class="border" align="center" border="0">
                          <tr class="tdbg">
                            <td class="clefttitle" align="right">公司名称：</td>
                            <td><input name="CompanyName" type="text" class="textbox" id="CompanyName" value="<%=CompanyName%>" size="30" maxlength="200" />
                                <span style="color: red">* </span> <span class="msgtips">请填写你在工商局注册登记的名称。</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle" align="right">注 册 号：</td>
                            <td><input name="BusinessLicense" class="textbox" type="text" id="BusinessLicense" value="<%=BusinessLicense%>" size="30" maxlength="50" /> <span class="msgtips">填写你的营业执照上的注册号码。</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle" align="right">企业法人：</td>
                            <td><input name="LegalPeople" class="textbox" type="text" id="LegalPeople" value="<%=LegalPeople%>" size="30" maxlength="50" />
                            <span style="color: red">* </span> <span class="msgtips">填写公司的法人或是主要负责人。</span></td>
                          </tr>
						   <tr class="tdbg">
                            <td class="clefttitle" align="right">公司地址：</td>
                            <td><input name="Address" class="textbox" type="text" id="Address" value="<%=Address%>" size="30" maxlength="50" /> <span class="msgtips">填写公司的注册地址</span></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">注册资金：</td>
                            <td>
							<input type="text" name="RegisteredCapital" value="<%=RegisteredCapital%>" class="textbox"/>
							<span class="msgtips">如：100万元</span></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">经营范围：</td>
                            <td>
							<textarea name="Business" id="Business" class="textbox" style="width:300px;height:60px"><%=Business%></textarea><span class="msgtips">请填写营业执照上的经营范围</span></td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle" align="right">成立日期：</td>
                            <td>
							<input type="text" name="Foundation" id="Foundation" value="<%=Foundation%>" class="textbox"/>
							<span class="msgtips">请填写营业执照上的成立日期</span></td>
                          </tr>
						 <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照复印件：</td>
						  <td><input type="file" class="textbox" name="photourl1" size="40">
						  <input type="hidden" name="photourl" value="<%=photourl%>"/>
						  </td>
						</tr>
						<%if photourl<>"" then%>
                       <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照预览：</td>
						  <td><a href="<%=photourl%>" target="_blank"><img src="<%=photourl%>" width="200" border="0" style="border:1px solid #cccccc;padding:1px"/></a>
						  </td>
						</tr>
					<%end if%>
						<tr class="tdbg">
                            <td class="clefttitle">&nbsp;</td>
                            <td><button class="pn" name="Submit" type="submit"><strong>提交审核</strong></button></td>
                          </tr>
		</table>
		</form>
		
		<%end if
	   End Sub
	   
	   Sub yyrzSave()
	        on error resume next
		     Dim fobj:Set FObj = New UpFileClass
			 FObj.GetData
			 if err.number<>0 then
			  call KS.AlertHistory("对不起,文件超出允许上传的大小!",-1)
			  response.end
			 end if
            Dim MaxFileSize:MaxFileSize = 1024   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.Setting(3)&KS.Setting(91)& uploaddir & "/" & KSUser.GetUserInfo("userid") & "/"
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"yyrz")
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):response.End()
			End Select
			If ReturnValue="" And Fobj.Form("photourl")="" Then
			 KS.Die "<script>alert('营业执照复印件必须上传!');history.back();</script>"
			End If
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select top 1 * From KS_EnterPrise Where UserName='" & KSUser.UserName & "'",conn,1,3
			If RS.Eof And RS.Bof Then
			 RS.AddNew
			End If
			 RS("CompanyName")=Fobj.Form("CompanyName")
			 RS("BusinessLicense")=Fobj.Form("BusinessLicense")
			 RS("LegalPeople")=Fobj.Form("LegalPeople")
			 RS("Address")=Fobj.Form("Address")
			 RS("RegisteredCapital")=Fobj.Form("RegisteredCapital")
			 RS("Business")=Fobj.Form("Business")
			 RS("Foundation")=Fobj.Form("Foundation")
	         If ReturnValue<>"" Then
			  RS("PhotoUrl")=ReturnValue
			 Else
			  RS("photourl")=Fobj.Form("photourl")
			 End If
			 RS("LastRzdate")=now
			 RS("IsRz")=2
			RS.Update
			RS.Close
			Set Fobj=Nothing
			Conn.Execute("Update KS_User Set IsRz=2 Where UserName='" & KSUser.UserName & "'")
	     KS.Die "<script>alert('恭喜，您的营业执行认证已提交，请等待管理员的审核!');location.href='user_rz.asp';</script>"
	   End Sub
	   
	 Sub Sfz()
	  KSUser.InnerLocation("身份证认证") 
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "select top 1 * from ks_user where username='" & KSUser.UserName & "'",conn,1,1
	  If RS.Eof And RS.Bof Then
	   RS.Close:Set RS=Nothing
	   KS.Die "error!"
	  End If
	 
	  If RS("Issfzrz")="1" Then
	  %>
	   <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0" class="border">
          <tr class="tdbg">
             <td class="clefttitle" align="right">真实姓名：</td>
             <td><%=RS("RealName")%></td>
         </tr>
          <tr class="tdbg">
             <td class="clefttitle" align="right">身份证号码：</td>
             <td><%=RS("IDCard")%></td>
         </tr>
		
		<%if RS("sfzphotourl")<>"" then%>
                       <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">身份证预览：</td>
						  <td><a href="<%=RS("sfzphotourl")%>" target="_blank"><img src="<%=RS("sfzphotourl")%>" width="200" border="0" style="border:1px solid #cccccc;padding:1px"/></a>
						  </td>
						</tr>
		<%end if%>
		<tr class="tdbg">
            <td class="clefttitle">&nbsp;</td>
            <td><button class="pn" name="Submit" onclick="history.back(-1);" type="button"><strong> 返 回 </strong></button></td>
        </tr>
	</table>
	  <%Else%>
	 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#RealName").val()==''){
			  $.dialog.alert('请输入您的姓名!',function(){
			  $("#RealName").focus();});
			  return false;
			}
			if ($("#IDCard").val()==''){
			  $.dialog.alert('请输入身份证号码!',function(){
			  $("#IDCard").focus();});
			  return false;
			}
			
			return true;
		  }
		 </script>
	 <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0" class="border">
		 <form action="?Action=sfzSave" method="post" name="myform"  enctype="multipart/form-data" id="myform" onSubmit="return CheckForm();">
			<input type="hidden" value="<%=KS.S("ComeUrl")%>" name="ComeUrl">
          <tr class="tdbg">
             <td class="clefttitle" align="right">真实姓名：</td>
             <td><input name="RealName" type="text" class="textbox" id="RealName" value="<%=RS("RealName")%>" size="30" maxlength="200" />
                  <span style="color: red">* </span> <span class="msgtips">必须与身份证在的姓名一致。</span></td>
         </tr>
          <tr class="tdbg">
             <td class="clefttitle" align="right">身份证号码：</td>
             <td><input name="IDCard" type="text" class="textbox" id="IDCard" value="<%=RS("IDCard")%>" size="30" maxlength="200" />
                  <span style="color: red">* </span> <span class="msgtips">填写公安机关发放的身份证号码。</span></td>
         </tr>
		 <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">身份证复印件：</td>
						  <td><input type="file" class="textbox" name="photourl1" size="40">
						  <input type="hidden" name="photourl" value="<%=RS("sfzphotourl")%>"/>
			  </td>
		</tr>
		<%if RS("sfzphotourl")<>"" then%>
                       <tr class="tdbg">
						  <td  height="25"class="clefttitle" align="right">营业执照预览：</td>
						  <td><a href="<%=RS("sfzphotourl")%>" target="_blank"><img src="<%=RS("sfzphotourl")%>" width="200" border="0" style="border:1px solid #cccccc;padding:1px"/></a>
						  </td>
						</tr>
		<%end if%>
						<tr class="tdbg">
                            <td class="clefttitle">&nbsp;</td>
                            <td><button class="pn" name="Submit" type="submit"><strong>提交审核</strong></button></td>
                          </tr>
	</form>
	</table>
	 <%
	 End If
	 RS.Close
	 Set RS=Nothing
	 End Sub
	 
	Sub sfzSave()
	        on error resume next
		     Dim fobj:Set FObj = New UpFileClass
			 FObj.GetData
			 if err.number<>0 then
			  call KS.AlertHistory("对不起,文件超出允许上传的大小!",-1)
			  response.end
			 end if
            Dim MaxFileSize:MaxFileSize = 1024   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.Setting(3)&KS.Setting(91)& uploaddir & "/" & KSUser.GetUserInfo("userid") & "/"
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"sfz")
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):response.End()
			End Select
			If ReturnValue="" And Fobj.Form("photourl")="" Then
			 KS.Die "<script>alert('身份证复印件必须上传!');history.back();</script>"
			End If
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,3
			RS("IDCard")=Fobj.Form("IDCard")
			RS("RealName")=Fobj.Form("RealName")
			If ReturnValue<>"" Then
			  RS("sfzPhotoUrl")=ReturnValue
			 Else
			  RS("sfzphotourl")=Fobj.Form("photourl")
			 End If
			RS("IsSfzRz")=2
			RS("LastRzdate")=now
		   RS.Update
			RS.Close
			Set Fobj=Nothing
			Session(KS.SiteSN&"UserInfo")=""
	     KS.Die "<script>alert('恭喜，您的身份证认证已提交，请等待管理员的审核!');location.href='user_rz.asp';</script>"
	End Sub
	
Sub mobilerz()
      KSUser.InnerLocation("手机认证")
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "select top 1 * from ks_user where username='" & KSUser.UserName & "'",conn,1,1
	  If RS.Eof And RS.Bof Then
	   RS.Close:Set RS=Nothing
	   KS.Die "error!"
	  End If
	 
	  If RS("Ismobilerz")="1" and request("flag")<>"rerz" Then
	  %>
	   <table  cellspacing="1" cellpadding="3" class="border" align="center" border="0">
          <tr class="tdbg">
             <td class="clefttitle" align="right">已通过认证的手机：</td>
             <td><%=RS("mobile")%></td>
         </tr>
		<tr class="tdbg">
            <td class="clefttitle">&nbsp;</td>
            <td><button class="pn" name="Submit" onclick="location.href='?action=mobilerz&flag=rerz'" type="button"><strong>更换手机号码</strong></button>
			<button class="pn" name="Submit" onclick="history.back(-1);" type="button"><strong> 返 回 </strong></button>
			</td>
        </tr>
	</table>
	  <%Else%>
	 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#Mobile").val()==''){
			  $.dialog.alert('请输入手机号码!',function(){
			  $("#Mobile").focus();
			  });
			  return false;
			}
			<%If Not KS.IsNul(Split(KS.Setting(155)&"∮∮","∮")(3)) And KS.Setting(157)="1" Then%>
		    if ($("#MobileCode").val()==''){
			  $.dialog.alert('请输入手机短信验证码!',function(){
			  $("#MobileCode").focus();
			  });
			  return false;
			}
			<%end if%>
			return true;
		  }
		 </script>
		 <form action="?Action=mobilerzSave" method="post" name="myform"  id="myform" onSubmit="return CheckForm();">
			<input type="hidden" value="<%=KS.S("ComeUrl")%>" name="ComeUrl">
	 <table  cellspacing="1" cellpadding="3" class="border" align="center" border="0">
			<%If Not KS.IsNul(Split(KS.Setting(155)&"∮∮","∮")(3)) And KS.Setting(157)="1" Then%>
			 	<tr class="tdbg">
				 <td class="clefttitle" align="right">手机号码：</td>
				 <td><input name="Mobile" type="text" class="textbox" id="Mobile" value="<%=RS("Mobile")%>" size="30" maxlength="200" /> <span style="color: red">* </span></span>
				 </td>
				</tr>
			 	<tr class="tdbg">
				 <td class="clefttitle" align="right">短信验证码：</td>
				 <td><input name="MobileCode" type="text" class="textbox" id="MobileCode" value="" size="10" maxlength="6" />
				 	<input type="button" value="免费获取手机验证码" id="MobileCodeBtn" onClick="getMobileCode(<%=KS.ChkClng(split(KS.Setting(156)&"∮","∮")(1))%>,'103','Mobile','MobileCodeBtn')" class="button"/>
					<input type="hidden" name="sendsms" value="1"/>
				 </td>
				</tr>
				<tr class="tdbg">
					   <td class="clefttitle">&nbsp;</td>
					   <td><button class="pn" name="Submit" type="submit"><strong>提交认证</strong></button></td>
				</tr>
			<%Else%>
				<tr class="tdbg">
				 <td class="clefttitle" align="right">手机号码：</td>
				 <td><input name="Mobile" type="text" class="textbox" id="Mobile" value="<%=RS("Mobile")%>" size="30" maxlength="200" />
				 <br/>  <span style="color: red">* </span> <span class="msgtips">管理员会人工发一条认证信息进行确定，请在收到信息后及时回复。</span>
				 </td>
				</tr>
				<tr class="tdbg">
					   <td class="clefttitle">&nbsp;</td>
					   <td><button class="pn" name="Submit" type="submit"><strong>提交审核</strong></button></td>
				</tr>
		 <%end if%>
	</table>
		</form>

	
	 <%
	 End If
	 RS.Close
	 Set RS=Nothing
	 End Sub	
	 
	 Sub mobilerzSave()
	       Dim RS
	       Dim Mobile:Mobile=KS.DelSQL(KS.S("Mobile"))
	       Dim MobileCode:MobileCode=KS.DelSQL(KS.S("MobileCode"))
	       If KS.IsNul(Mobile) Then
		    KS.Die "<script>alert('请输入手机号码!');history.back();</script>"
		   End If

	        Dim sendsms:sendsms=KS.ChkClng(KS.S("sendsms"))
			If sendsms=1 Then
			   	  If KS.IsNul(MobileCode) Then
					KS.Die "<script>alert('请输入手机短信验证码!');history.back();</script>"
				   End If
				   Dim RSM:Set RSM=Conn.Execute("Select top 1 * From KS_UserRecord Where flag=103 And UserName='" & Mobile &"' Order By ID Desc")
				 If RSM.Eof And RSM.Bof Then
				   RSM.Close
				   Set RSM=Nothing
				   Response.Write("<script>alert('对不起，您输入的手机号码与接收短消息的手机号码不一致！');history.back(-1);</script>")
				   Exit Sub
				 End If
				 Dim RightMobileCode:RightMobileCode=RSM("Note")
				 Dim RightSendDate:RightSendDate=RSM("AddDate")
				 Dim RightMobile:RightMobile=RSM("UserName")
				 RSM.Close
				 Set RSM=Nothing
				 Dim TimeAllow:TimeAllow=KS.ChkClng(split(KS.Setting(156)&"∮∮","∮")(4))
				 If RightMobile<>Mobile Then
				   Response.Write("<script>alert('对不起，您输入的手机号码与接收短消息的手机号码不一致！');history.back(-1);</script>")
				   Exit Sub
				 ElseIf MobileCode<>RightMobileCode Then
				   Response.Write("<script>alert('对不起，您输入的手机短信验证码不正确！');history.back(-1);</script>")
				   Exit Sub
				 ElseIf TimeAllow>0 and DateDiff("n",RightSendDate,Now)>TimeAllow Then
				   Response.Write("<script>alert('对不起，您输入的手机短信验证码已失效！');history.back(-1);</script>")
				   Exit Sub
				 End If
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,3
				RS("Mobile")=Mobile
				RS("IsMobileRz")=1
				RS("LastRzdate")=now
			    RS.Update
				RS.Close
				Set RS=Nothing
				Session(KS.SiteSN&"UserInfo")=""
			    KS.Die "<script>alert('恭喜，您的手机号码通过实名认证成功!');location.href='user_rz.asp';</script>"

			Else
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,3
				RS("Mobile")=Mobile
				RS("IsMobileRz")=2
				RS("LastRzdate")=now
			   RS.Update
				RS.Close
				Set RS=Nothing
				Session(KS.SiteSN&"UserInfo")=""
			    KS.Die "<script>alert('恭喜，您的手机认证已提交，请等待管理员的审核!');location.href='user_rz.asp';</script>"
		   End If
	 End Sub
	 
	 Sub emailrz()
      KSUser.InnerLocation("邮箱认证")
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "select top 1 * from ks_user where username='" & KSUser.UserName & "'",conn,1,1
	  If RS.Eof And RS.Bof Then
	   RS.Close:Set RS=Nothing
	   KS.Die "error!"
	  End If
	 
	  If RS("Isemailrz")="1" Then
	  %>
	   <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0">
          <tr class="tdbg">
             <td class="clefttitle" align="right">已通过认证的邮箱：</td>
             <td><%=RS("email")%></td>
         </tr>
		<tr class="tdbg">
            <td class="clefttitle">&nbsp;</td>
            <td><button class="pn" name="Submit" onclick="history.back(-1);" type="button"><strong> 返 回 </strong></button></td>
        </tr>
	</table>
	  <%Else%>
	 <script type="text/javascript">
		  function CheckForm(){
		    if ($("#Email").val()==''){
			  $.dialog.alert('请输入电子邮箱!',function(){
			  $("#Email").focus();});
			  return false;
			}
			return true;
		  }
		 </script>
	 <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0" class="border">
		 <form action="?Action=emailrzSave" method="post" name="myform"  id="myform" onSubmit="return CheckForm();">
			<input type="hidden" value="<%=KS.S("ComeUrl")%>" name="ComeUrl">
          <tr class="tdbg">
             <td class="clefttitle" align="right">电子邮箱：</td>
             <td><input name="Email" type="text" class="textbox" id="Email" value="<%=RS("Email")%>" size="30" maxlength="200" />
			 <br/>
                <span class="msgtips"><span style="color: red">* </span> 系统会自动发一封认证邮箱，提交后请登录您的邮箱确认。</span></td>
         </tr>
			<tr class="tdbg">
                            <td class="clefttitle">&nbsp;</td>
                            <td><button class="pn" name="Submit" type="submit"><strong>提交认证</strong></button></td>
            </tr>
	</form>
	</table>
	 <%
	 End If
	 RS.Close
	 Set RS=Nothing
	 End Sub
	 
	 Sub emailrzSave()
	  KSUser.InnerLocation("邮箱认证")
	  Dim Email:Email=KS.G("Email")
	  If Email="" Or Not KS.IsValidEmail(Email) Then
	    KS.Die "<script>alert('请输入的邮箱不正确，请重输!');history.back();</script>"
	  End If
	  Dim MailBodyStr:MailBodyStr=KSUser.UserName & ",您好!<br/>&nbsp;&nbsp;&nbsp;&nbsp;这封邮件是您在【" & KS.Setting(0) & "】网站申请电子邮箱实名认证的确认信，如果本操作是您本人提交的请点以下链接确认，否则请删除本邮件。<br/><a href='" & KS.GetDomain & "user/User_rz.asp?action=checkemail&email="&email&"&userid=" & KSUser.GetUserInfo("userid") & "&rnd=" & KS.C("RndPassWord") & "' target='_blank'>" & KS.GetDomain & "user/User_rz.asp?action=checkemail&email="&email&"&userid=" & KSUser.GetUserInfo("userid") & "&rnd=" & KS.C("RndPassWord") & "</a>。"
	  
	  
	  Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "电子邮箱实名认证确认信", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
	  IF ReturnInfo="OK" Then
		 Conn.Execute("Update KS_User Set IsEmailRz=2 where username='" & KSUser.UserName & "'") 
		 %>
		 <br/>
		 <br/>
		 <br/>
		 <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0">
			  <tr class="tdbg">
				 <td>您的邮箱实名认证邮箱确认信已发送到<font color=red><%=Email%></font>，请及时登录邮箱确认。<br/>如果长时间没有收到或是确认性失效请<input type='button' value='点此按钮' onclick="location.href='user_rz.asp?action=emailrzSave&email=<%=Email%>';" class="button"/>重发确认信!</td>
			 </tr>
				<tr class="tdbg">
					<td style="text-align:center;height:50px"><button class="pn" name="Submit" onclick="history.back(-1);" type="button"><strong> 返 回 </strong></button></td>
				</tr>
		</table>
		 <%
	  Else
	  %>		 <br/>
		 <br/>
		 <br/>

	   <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0">
			  <tr class="tdbg">
				 <td>信件发送失败!失败原因:<%= ReturnInfo%></td>
			 </tr>
				<tr class="tdbg">
					<td  style="text-align:center;height:50px"><button class="pn" name="Submit" onclick="history.back(-1);" type="button"><strong> 返 回 </strong></button>
								<input type='button' value='重发确认信' onclick="location.href='user_rz.asp?action=emailrzSave&email=<%=Email%>';" class="button"/>
								</td>
				</tr>
		</table>
		<%
	  End if
	 End Sub
	    
	Sub Main()
	 Dim IsYYRZRz:IsYYRZRz=0
	 Dim IsSFZRZ:IsSFzRz=KS.ChkClng(conn.execute("select top 1 issfzrz from ks_user where username='" & KSUser.UserName & "'")(0))
	 Dim IsMobileRZ:IsMobileRZ=KS.ChkClng(conn.execute("select top 1 ismobilerz from ks_user where username='" & KSUser.UserName & "'")(0))
	 Dim IsEmailRZ:IsEmailRZ=KS.ChkClng(conn.execute("select top 1 isemailrz from ks_user where username='" & KSUser.UserName & "'")(0))
	 Dim IsQY:IsQY=False
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "select top 1 * From KS_EnterPrise Where UserName='" & KSUser.UserName & "'",conn,1,1
	 If Not RS.Eof Then
	   IsQY=true
	   IsYYRZRz=KS.ChkClng(RS("IsRz"))
	 End If
	 RS.Close:Set RS=Nothing
	  %>
    
	 <style>
	 .rztips{color:#999999}
	 </style>
	  <div class="tabs">	
			<ul>
				<li class='puton'>实名认证</li>
			</ul>
	  </div>

	 <div style="margin:20px;font-size:14px">
	 请不要提交虚假信息，我们对用户上传的证件严格保密。<font color=green>为保证实名认证的权威性，认证通过的信息将不可以更改</font>。 <a href="../space/company/rz.asp?userid=<%=ksuser.getuserinfo("userid")%>" target="_blank" style="color:#FF0000">查看认证资料</a>
	 </div>
	<table border="0" class="border" align="center">
	 <%if IsQY then%>
				<tr>
				  <td height="50" class="splittd">
				    <%if IsYYRZRZ=1 Then%>
				     <img src="images/ok.gif" align="absmiddle"/>
					 <%elseif isyyrzrz=2 then%>
				     <img src="images/ico.gif" align="absmiddle"/>
					<%else%>
				     <img src="images/wrong.gif" align="absmiddle"/>
					<%end if%> <strong>营业执照认证</strong>
				   </td>
				  <td  class="splittd rztips"><%If IsYYRZRz=0 then%>您还没有提交企业营业执照认证，提高您在本站的信誉度，建议您及时上传认证
				  <%elseif IsYYRZRz=2 then%>
				   您的营业执照认证已提交审核，正在审核中，请耐心等待审核结果...
				  <%elseif IsYYRZRz=3 then%>
				   您的营业执照认证已提交,但没有通过审核，请重新提交认证信息...
				  <%else%>
				   恭喜，您的营业执照已通过实名认证审核。
				  <%end if%>
				  </td>
				  <td  class="splittd"><a href="?action=yyrz"><%If IsYYRZRz=0 then%>
				  立即认证
				  <%elseif isYYRZRz=1 then%>
				  已认证，查看认证资料
				  <%else%>
				   修改认证资料
				  <%end if%></a></td>
				</tr>
	 <%end if%>
				<tr>
				  <td height="50" class="splittd">
				     <%if IsSFZRZ=1 Then%>
				     <img src="images/ok.gif" align="absmiddle"/>
					 <%elseif issfzrz=2 then%>
				     <img src="images/ico.gif" align="absmiddle"/>
					 <%else%>
				     <img src="images/wrong.gif" align="absmiddle"/>
					 <%end if%> <strong>身份证认证</strong>
				   </td>
				  <td  class="splittd rztips"> <%if IsSFZRZ=0 Then%>您还没有提交身份认证信息，建议及时上传认证
				  <%elseif IsSFZRZ=2 then%>
				    您的身份证实名认证已提交，正在审核中，请耐心等待审核结果...
				  <%elseif IsSFZRZ=3 then%>
				    您的身份证实名认证已得交，但没有通过审核，请重新提交认证信息...
				  <%else%>
				    恭喜，您的身份证已通过实名认证审核。
				  <%end if%>
				    </td>
				  <td  class="splittd">
				  <a href="?action=sfz">
				  <%If IsSFZRZ=0 then%>
				  立即认证
				  <%elseif IsSFZRZ=1 then%>
				   已认证，查看认证资料
				  <%else%>
				   修改认证资料
				  <%end if%>
				  </a>
				  </td>
				</tr>
				<tr>
				  <td height="50" class="splittd">
				    <%if IsMobileRZ=1 then%>
					<img src="images/ok.gif" align="absmiddle"/>
					 <%elseif ismobilerz=2 then%>
				     <img src="images/ico.gif" align="absmiddle"/>
					<%else%>
				     <img src="images/wrong.gif" align="absmiddle"/>
					<%end if%> <strong>手机认证</strong>
				   </td>
				  <td  class="splittd rztips"><%if IsMobileRZ=0 Then%>输入的您手机号，系统会发一条短信验证码通知您确认认证
				  <%elseif IsMobileRZ=2 then%>
				    您的手机实名认证已提交，正在审核中，请耐心等待审核结果...
				  <%elseif IsMobileRZ=3 then%>
				    您的手机实名认证已得交，但没有通过审核，请重新提交认证信息...
				  <%else%>
				    恭喜，您的手机已通过实名认证审核。
				  <%end if%></td>
				  <td  class="splittd">
				  <a href="?action=mobilerz">
				  <%If IsMobileRZ=0 then%>
				  立即认证
				  <%elseif IsMobileRZ=1 then%>
				   已认证，查看认证资料
				  <%else%>
				   修改认证资料
				  <%end if%>
				  </td>
				</tr>
				<tr>
				  <td height="50" class="splittd">
				  <%if IsEmailRZ=1 then%>
				     <img src="images/ok.gif" align="absmiddle"/>
					 <%elseif isemailrz=2 then%>
				     <img src="images/ico.gif" align="absmiddle"/>
				  <%else%>
				     <img src="images/wrong.gif" align="absmiddle"/>
				  <%end if%> <strong>邮箱认证</strong>
				   </td>
				  <td  class="splittd rztips"><%if IsEmailRZ=0 Then%>输入您的常用邮箱，根据提示完成认证
				  <%elseif IsEmailRZ=2 then%>
				    您的邮箱实名认证已提交，请登录您的邮箱完成最后认证工作。
				  <%else%>
				    恭喜，您的邮箱已通过实名认证。
				  <%end if%></td>
				  <td  class="splittd">
				  <a href="?action=emailrz">
				  <%If IsEmailRZ=0 then%>
				   立即认证
				  <%elseif IsEmailRZ=1 then%>
				   已认证，查看认证资料
				  <%else%>
				   修改认证资料
				  <%end if%>
				  </a>
				  </td>
				</tr>
				<tr>
				  <td height="50" class="splittd">
				  <%if not conn.execute("select top 1 * From KS_EnterpriseZS where username='" & KSUser.UserName & "'").eof then%>
				     <img src="images/ok.gif" align="absmiddle"/>
				  <%else%>
				     <img src="images/wrong.gif" align="absmiddle"/>
				  <%end if%> <strong>其它认证证件</strong>
				   </td>
				  <td  class="splittd rztips">可以上传其它证件,如专利证书,版权证书等等。</td>
				  <td  class="splittd">
				  <a href="user_Enterprisezs.asp?Action=Add">
				   我要上传
				  </a>
				  </td>
				</tr>
				
				
	</table> 
				
		  <%
  End Sub
  
  
    
  
End Class
%> 

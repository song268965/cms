<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Down_Param
KSCls.Kesion()
Set KSCls = Nothing

Class Down_Param
        Private KS,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 With KS
		  	.echo "<!DOCTYPE html><html>"
			.echo "<title>下载基本参数设置</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "</head>"
			
			Dim RS, Action, SQLStr
			Dim DownLb, DownYY, DownSQ, DownPT, JyDownUrl, JyDownWin
			Action = KS.G("Action")
			ChannelID= KS.ChkClng(KS.G("ChannelID"))
			If ChannelID=0 Then ChannelID=3
			If Not KS.ReturnPowerResult(0, "KMST20001") Then Call KS.ReturnErr(1, "")   '下载基本参数设置权限检查
			
			SQLStr = "Select * From KS_DownParam Where ChannelID=" & ChannelID
			Set RS = Server.CreateObject("Adodb.RecordSet")
			If Action = "save" Then
			  RS.Open SQLStr, conn, 1, 3
			  If RS.Eof Then
			   RS.AddNew
			   RS("ChannelID")=ChannelID
			  End If
			  RS("DownLb") = KS.G("DownLB")
			  RS("DownYY") = KS.G("DownYY")
			  RS("DownSQ") = KS.G("DownSQ")
			  RS("DownPT") = KS.G("DownPT")
			  RS.Update
			  .echo ("<script>top.$.dialog.alert('下载参数修改成功!');</script>")
			  RS.Close
			End If
			 RS.Open SQLStr, conn, 1, 1
			  If Not RS.EOF Then
			   DownLb = RS("DownLb")
			   DownYY = RS("DownYy")
			   DownSQ = RS("DownSQ")
			   DownPT = RS("DownPT")
			  End If
			RS.Close
			
			Set RS = Nothing
			.echo "<body topmargin=""0"" leftmargin=""0"">"
			.echo "      <div class='tabTitle mt20'>"
			.echo "      [" & KS.C_S(ChannelID,1) &"]参数设置"
			.echo "      </div>"
			
			.echo "<div class='pageCont2'><dl class=""dtable""><dd style='display:none'><div>模型:</div><select"
			if request("channelid")<>"" then .echo " disabled"
			.echo " id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.echo " <option value='0'>---请选择模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks6=3]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			    .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.echo "</select></dd>"

			.echo "<form action=""?ChannelID=" & ChannelID &"&Action=save"" method=""post"" name=""DownParamForm"">"
			.echo "  <dd><div>软件类别</div>"
			.echo "        <textarea name=""DownLb"" cols=""70"" rows=""10"">" & DownLb & "</textarea>"
			.echo "        <span class=""block"">说明：每一个类别为一行</span></dd>"
			.echo "      <dd><div>设定语言：</div>"
			.echo "      <textarea name=""DownYy"" cols=""70"" rows=""5"">" & DownYY & "</textarea>"
			.echo "      <span class=""block"">说明：每一种语言为一行</span></dd>"
			.echo "      <dd><div>授权形式：</div>"
			.echo "      <textarea name=""DownSq"" cols=""70"" rows=""5"">" & DownSQ & "</textarea>"
			.echo "        <br>"
			.echo "        <span class=""block"">说明：每一种授权方式为一行</span></dd>"
			.echo "      <dd><div>运行平台：</div>"
			.echo "      <textarea name=""DownPt"" cols=""70"" rows=""5"">" & DownPT & "</textarea>"
			.echo "      <span class=""block"">说明：每一种运行平台为一行</span></dd>"
			.echo "  </dl>"
			.echo "</form></div>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language=""javascript"">"
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.DownParamForm;" & vbCrLf
			  .echo "    form.submit();" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->"
			.echo "</Script>"
			End With
		End Sub

End Class
%> 
